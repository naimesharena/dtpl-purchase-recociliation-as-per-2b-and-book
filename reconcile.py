"""
Purchase reconciliation : Books (Tally purchase registers) vs GSTR-2B (B2B sheet).

Handles the client's data quirks:
  * Non-GST purchases in books (no CGST/SGST/IGST) -> will never appear in 2B.
  * Ineligible-credit bills booked at TOTAL amount (tax merged in value, tax cols blank).
  * Same supplier registered in multiple states -> invoice reported under a
    different GSTIN of the SAME PAN. Matched on PAN + invoice no + amount.

Output: PURCHASE-RECO-BOOK-VS-2B.xlsx  (one reconciliation sheet + summary)
"""

import re
import pandas as pd
import numpy as np
from datetime import datetime

BOOK_FILES = {
    "HR": "PURCHASE-HR.xls",
    "PL": "PURCHASE-PL.xls",
    "VL": "PURCHASE-VL.xls",
}
GSTR2B_FILE = "gstr-2B.xls"
OUT_FILE = "PURCHASE-RECO-BOOK-VS-2B.xlsx"

AMT_TOL = 2.0        # rupee tolerance for amount matching
DATE_TOL_DAYS = 30   # tolerance for probable (amount based) matching


# ----------------------------------------------------------------- helpers
def norm_inv(v):
    """Invoice no key: uppercase alnum only, leading zeros of each numeric block stripped."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    s = str(v).strip().upper()
    if s.endswith(".0"):
        s = s[:-2]
    s = re.sub(r"[^A-Z0-9]", "", s)
    s = s.lstrip("0")                     # drop leading zeros
    return s


def inv_digits(v):
    s = re.sub(r"\D", "", str(v) if v is not None else "")
    s = s.lstrip("0")
    return s


def pan_of(g):
    g = str(g).strip().upper()
    return g[2:12] if len(g) >= 12 else ""


def to_num(s):
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def to_date(s):
    return pd.to_datetime(s, errors="coerce", dayfirst=True)


# ----------------------------------------------------------------- loaders
def load_books():
    frames = []
    for branch, f in BOOK_FILES.items():
        raw = pd.read_excel(f, header=None, engine="openpyxl", skiprows=9)
        raw.columns = [str(c).strip() for c in raw.iloc[0]]
        df = raw.iloc[1:].copy()
        df = df[df["Particulars"].astype(str).str.strip().str.lower() != "grand total"]
        df = df[df["Date"].notna()]

        out = pd.DataFrame({
            "Branch": branch,
            "Book Date": to_date(df["Date"]),
            "Supplier (Books)": df["Particulars"].astype(str).str.strip(),
            "Voucher Type": df["Voucher Type"],
            "Voucher No.": df["Voucher No."].astype(str).str.strip(),
            "Book Invoice No.": df["Supplier Invoice No."].astype(str).str.strip(),
            "Book Invoice Date": to_date(df["Supplier Invoice Date"]),
            "Book GSTIN": df["GSTIN/UIN"].astype(str).str.strip().str.upper()
                            .replace({"NAN": "", "NONE": ""}),
            "Book Total": to_num(df["Gross Total"]),
            "Book CGST": to_num(df.get("CGST TAX")),
            "Book SGST": to_num(df.get("SGST TAX")),
            "Book IGST": to_num(df.get("IGST TAX")),
            "Book Round Off": to_num(df.get("Round Off")),
        })
        out["Book Tax"] = out["Book CGST"] + out["Book SGST"] + out["Book IGST"]

        # Taxable value = sum of the expense / asset ledger columns of the voucher.
        # (Do NOT back-calculate from Gross Total: the client nets off TDS in the
        #  voucher total, so Total = Taxable + Tax + RoundOff - TDS.)
        ledger_cols = [c for c in df.columns
                       if str(c).strip() in ("Purchase Accounts", "Indirect Expenses",
                                             "Direct Expenses", "Fixed Assets")]
        ledger = sum(to_num(df[c]) for c in ledger_cols) if ledger_cols else 0.0
        ledger = pd.Series(ledger, index=out.index).round(2)
        fallback = (out["Book Total"] - out["Book Tax"] - out["Book Round Off"]).round(2)
        out["Book Taxable"] = np.where(ledger.abs() > 0.009, ledger, fallback)

        # Gross invoice value as per the bill, before any TDS / other deduction
        out["Book Invoice Value"] = (out["Book Taxable"] + out["Book Tax"]
                                     + out["Book Round Off"]).round(2)
        out["Book TDS / Deduction"] = (out["Book Invoice Value"]
                                       - out["Book Total"]).round(2)
        out.loc[out["Book TDS / Deduction"].abs() < 0.01, "Book TDS / Deduction"] = 0.0
        frames.append(out)

    b = pd.concat(frames, ignore_index=True)
    b["No Tax in Books"] = b["Book Tax"].abs() < 0.01
    b["Book Nature"] = np.where(
        ~b["No Tax in Books"], "GST Purchase",
        np.where(b["Book GSTIN"].str.len() == 15,
                 "Booked at Total (Ineligible ITC / Tax not split)",
                 "Non-GST Purchase"))
    b["PAN"] = b["Book GSTIN"].map(pan_of)
    b["ikey"] = b["Book Invoice No."].map(norm_inv)
    b["vkey"] = b["Voucher No."].map(norm_inv)
    b["idig"] = b["Book Invoice No."].map(inv_digits)
    b["vdig"] = b["Voucher No."].map(inv_digits)
    b["_bi"] = range(len(b))
    return b


def load_2b():
    df = pd.read_excel(GSTR2B_FILE, sheet_name="B2B", header=None,
                       engine="xlrd", skiprows=6)
    df.columns = ["GSTIN", "Name", "InvNo", "InvType", "InvDate", "InvValue",
                  "POS", "RCM", "Taxable", "IGST", "CGST", "SGST", "Cess",
                  "Period", "FileDate", "ITCAvail", "Reason", "ApplPct",
                  "Source", "IRN", "IRNDate"][:df.shape[1]]
    df = df[df["GSTIN"].astype(str).str.len() == 15].copy()

    t = pd.DataFrame({
        "2B GSTIN": df["GSTIN"].str.strip().str.upper(),
        "2B Supplier": df["Name"].astype(str).str.strip(),
        "2B Invoice No.": df["InvNo"].astype(str).str.strip(),
        "2B Invoice Date": to_date(df["InvDate"]),
        "2B Invoice Value": to_num(df["InvValue"]),
        "2B Taxable": to_num(df["Taxable"]),
        "2B IGST": to_num(df["IGST"]),
        "2B CGST": to_num(df["CGST"]),
        "2B SGST": to_num(df["SGST"]),
        "2B Cess": to_num(df["Cess"]),
        "2B ITC Available": df["ITCAvail"],
        "2B Period": df["Period"],
    })
    t["2B Tax"] = t["2B IGST"] + t["2B CGST"] + t["2B SGST"] + t["2B Cess"]
    t["PAN"] = t["2B GSTIN"].map(pan_of)
    t["ikey"] = t["2B Invoice No."].map(norm_inv)
    t["idig"] = t["2B Invoice No."].map(inv_digits)
    t["_ti"] = range(len(t))
    return t


# ----------------------------------------------------------------- matching
def close(a, b, tol=AMT_TOL):
    return abs(float(a) - float(b)) <= tol


def amount_agrees(brow, trow):
    """True if the book value lines up with the 2B figures in any valid way.

    - normal GST bill  : taxable ~ taxable  and  total ~ invoice value
    - booked at total  : book total ~ invoice value  (tax merged, not split)
    - ineligible ITC   : book total ~ taxable + tax  (same as invoice value)
    """
    bt, btax = brow["Book Taxable"], brow["Book Tax"]
    bgross, btot = brow["Book Invoice Value"], brow["Book Total"]
    return (close(bt, trow["2B Taxable"])
            or close(bgross, trow["2B Invoice Value"])
            or close(bgross, trow["2B Taxable"] + trow["2B Tax"])
            or close(btot, trow["2B Invoice Value"])
            or close(btot, trow["2B Taxable"] + trow["2B Tax"])
            or close(btot, trow["2B Taxable"])
            or close(bgross, trow["2B Taxable"])
            or (close(btax, trow["2B Tax"]) and close(bt, trow["2B Taxable"], 5)))


def reconcile(books, t2b):
    used_t, pairs = set(), []
    t_by_gstin_ikey, t_by_pan_ikey, t_by_pan_idig, t_by_pan = {}, {}, {}, {}
    for r in t2b.to_dict("records"):
        t_by_gstin_ikey.setdefault((r["2B GSTIN"], r["ikey"]), []).append(r)
        t_by_pan_ikey.setdefault((r["PAN"], r["ikey"]), []).append(r)
        t_by_pan_idig.setdefault((r["PAN"], r["idig"]), []).append(r)
        t_by_pan.setdefault(r["PAN"], []).append(r)

    brecs = books.to_dict("records")

    def take(cands, brow, need_amount):
        best = None
        for c in cands:
            if c["_ti"] in used_t:
                continue
            ok = amount_agrees(brow, c)
            if need_amount and not ok:
                continue
            score = (0 if ok else 1,
                     abs((brow["Book Invoice Date"] - c["2B Invoice Date"]).days)
                     if pd.notna(brow["Book Invoice Date"]) and pd.notna(c["2B Invoice Date"]) else 999)
            if best is None or score < best[0]:
                best = (score, c, ok)
        return best

    # ---- pass 1 : same GSTIN + invoice no + amount
    # ---- pass 2 : same GSTIN + invoice no (amount differs)
    # ---- pass 3 : same PAN (other state GSTIN) + invoice no + amount
    # ---- pass 4 : same PAN + invoice no
    # ---- pass 5 : same PAN + numeric part of invoice + amount
    # ---- pass 6 : same PAN + amount + date within tolerance
    def g_ikey(b):   # same GSTIN, invoice no. OR voucher no.
        return (t_by_gstin_ikey.get((b["Book GSTIN"], b["ikey"]), [])
                + t_by_gstin_ikey.get((b["Book GSTIN"], b["vkey"]), []))

    def p_ikey(b):   # same PAN (other state GSTIN)
        return (t_by_pan_ikey.get((b["PAN"], b["ikey"]), [])
                + t_by_pan_ikey.get((b["PAN"], b["vkey"]), []))

    def p_idig(b):
        return (t_by_pan_idig.get((b["PAN"], b["idig"]), [])
                + t_by_pan_idig.get((b["PAN"], b["vdig"]), []))

    passes = [
        ("Matched", g_ikey, True,
         "Exact match - GSTIN, invoice no. & amount"),
        ("Matched with Difference", g_ikey, False,
         "Same GSTIN & invoice no., value difference"),
        ("Matched - GSTIN Mismatch (Same PAN)", p_ikey, True,
         "Supplier billed from another state GSTIN of same PAN - invoice no. & amount match"),
        ("Matched - GSTIN Mismatch (Same PAN)", p_ikey, False,
         "Same PAN & invoice no., value difference"),
        ("Probable Match - Invoice No. Differs", p_idig, True,
         "Invoice no. format differs (numeric part & amount match)"),
        ("Probable Match - Amount Based", lambda b: t_by_pan.get(b["PAN"], []), True,
         "Matched on supplier PAN + value within tolerance"),
    ]

    pending = brecs
    for status, getc, need_amt, remark in passes:
        nxt = []
        for b in pending:
            if not b["PAN"]:
                nxt.append(b)
                continue
            if status.startswith("Probable Match - Amount") :
                if not b["ikey"]:
                    nxt.append(b); continue
            if need_amt is False and not b["ikey"]:
                nxt.append(b); continue
            res = take(getc(b), b, need_amt)
            if res is None:
                nxt.append(b); continue
            score, c, ok = res
            if status.startswith("Probable Match - Amount"):
                d = score[1]
                if d > DATE_TOL_DAYS:
                    nxt.append(b); continue
            used_t.add(c["_ti"])
            pairs.append((b, c, status, remark))
        pending = nxt

    for b in pending:
        if b["No Tax in Books"] and b["Book Nature"] == "Non-GST Purchase":
            st, rm = "In Books Only - Non-GST", "Non-GST purchase, not reportable in GSTR-2B"
        elif b["No Tax in Books"]:
            st, rm = ("In Books Only - Booked at Total",
                      "Tax not split in books (ineligible credit / total value booked)")
        else:
            st, rm = "In Books Only", "Not found in GSTR-2B (B2B)"
        pairs.append((b, None, st, rm))

    for c in t2b.to_dict("records"):
        if c["_ti"] not in used_t:
            pairs.append((None, c, "In GSTR-2B Only", "Not recorded in books"))

    return pairs


# ----------------------------------------------------------------- output
BOOK_COLS = ["Branch", "Book Date", "Supplier (Books)", "Book GSTIN", "Voucher Type",
             "Voucher No.", "Book Invoice No.", "Book Invoice Date", "Book Taxable",
             "Book IGST", "Book CGST", "Book SGST", "Book Tax", "Book Invoice Value",
             "Book TDS / Deduction", "Book Total", "Book Nature"]
T2B_COLS = ["2B GSTIN", "2B Supplier", "2B Invoice No.", "2B Invoice Date",
            "2B Taxable", "2B IGST", "2B CGST", "2B SGST", "2B Tax",
            "2B Invoice Value", "2B ITC Available", "2B Period"]

STATUS_ORDER = ["Matched", "Matched - GSTIN Mismatch (Same PAN)",
                "Matched with Difference", "Probable Match - Invoice No. Differs",
                "Probable Match - Amount Based", "In Books Only",
                "In Books Only - Booked at Total", "In Books Only - Non-GST",
                "In GSTR-2B Only"]


def build_output(pairs):
    rows = []
    for b, c, status, remark in pairs:
        r = {"Status": status, "Match Remark": remark}
        for k in BOOK_COLS:
            r[k] = b[k] if b else None
        for k in T2B_COLS:
            r[k] = c[k] if c else None
        if b and c:
            r["Diff - Taxable"] = round(b["Book Taxable"] - c["2B Taxable"], 2)
            r["Diff - Tax"] = round(b["Book Tax"] - c["2B Tax"], 2)
            r["Diff - Invoice Value"] = round(b["Book Invoice Value"] - c["2B Invoice Value"], 2)
            r["GSTIN Same?"] = "Yes" if b["Book GSTIN"] == c["2B GSTIN"] else "No - same PAN, other state"
        rows.append(r)

    df = pd.DataFrame(rows, columns=["Status", "Match Remark"] + BOOK_COLS + T2B_COLS +
                      ["Diff - Taxable", "Diff - Tax", "Diff - Invoice Value", "GSTIN Same?"])
    df["_o"] = df["Status"].map({s: i for i, s in enumerate(STATUS_ORDER)}).fillna(99)
    df = df.sort_values(["_o", "Book Date", "2B Invoice Date"], na_position="last").drop(columns="_o")
    df.insert(0, "Sr.", range(1, len(df) + 1))
    return df


def summary(df):
    g = df.groupby("Status").agg(
        **{"No. of Records": ("Sr.", "count"),
           "Book Taxable": ("Book Taxable", "sum"),
           "Book Tax": ("Book Tax", "sum"),
           "Book Invoice Value": ("Book Invoice Value", "sum"),
           "Book TDS / Deduction": ("Book TDS / Deduction", "sum"),
           "Book Total": ("Book Total", "sum"),
           "2B Taxable": ("2B Taxable", "sum"),
           "2B Tax": ("2B Tax", "sum"),
           "2B Invoice Value": ("2B Invoice Value", "sum")}).reset_index()
    g["_o"] = g["Status"].map({s: i for i, s in enumerate(STATUS_ORDER)}).fillna(99)
    g = g.sort_values("_o").drop(columns="_o")
    tot = {"Status": "TOTAL"}
    for c in g.columns[1:]:
        tot[c] = g[c].sum()
    return pd.concat([g, pd.DataFrame([tot])], ignore_index=True)


def write_excel(df, summ):
    with pd.ExcelWriter(OUT_FILE, engine="openpyxl") as xl:
        summ.to_excel(xl, sheet_name="Summary", index=False, startrow=2)
        df.to_excel(xl, sheet_name="Book vs 2B Reco", index=False, startrow=2)

        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        fills = {
            "Matched": "C6EFCE",
            "Matched - GSTIN Mismatch (Same PAN)": "FFF2CC",
            "Matched with Difference": "FCE4D6",
            "Probable Match - Invoice No. Differs": "FFF2CC",
            "Probable Match - Amount Based": "FFF2CC",
            "In Books Only": "FFC7CE",
            "In Books Only - Booked at Total": "DDEBF7",
            "In Books Only - Non-GST": "E2EFDA",
            "In GSTR-2B Only": "D9D2E9",
        }
        thin = Side(style="thin", color="BFBFBF")

        for name, data in (("Summary", summ), ("Book vs 2B Reco", df)):
            ws = xl.sheets[name]
            ws["A1"] = ("DRIVE TRUCKING PVT LTD - Purchase Reconciliation : "
                        "As per Books vs As per GSTR-2B (B2B)  |  FY 2025-26")
            ws["A1"].font = Font(bold=True, size=13)
            ws["A2"] = ("Books: PURCHASE-HR / PL / VL   |   2B: B2B sheet   |   "
                        f"Generated {datetime.now():%d-%m-%Y %H:%M}")
            ws["A2"].font = Font(italic=True, size=9, color="555555")

            hdr = 3
            for cell in ws[hdr]:
                if cell.value is None:
                    continue
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="305496")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = Border(thin, thin, thin, thin)
            ws.freeze_panes = ws.cell(row=hdr + 1, column=4 if name != "Summary" else 2)
            ws.auto_filter.ref = ws.dimensions

            cols = list(data.columns)
            si = cols.index("Status") + 1
            for row in ws.iter_rows(min_row=hdr + 1, max_row=hdr + len(data)):
                st = ws.cell(row=row[0].row, column=si).value
                f = fills.get(st)
                for cell in row:
                    cell.border = Border(thin, thin, thin, thin)
                    if f:
                        cell.fill = PatternFill("solid", fgColor=f)
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = "#,##0.00"
                    if isinstance(cell.value, datetime):
                        cell.number_format = "dd-mm-yyyy"

            widths = {"Book TDS / Deduction": 18, "Book Invoice Value": 16, "Status": 34, "Match Remark": 48, "Supplier (Books)": 30,
                      "2B Supplier": 30, "Book GSTIN": 17, "2B GSTIN": 17,
                      "Book Invoice No.": 20, "2B Invoice No.": 20,
                      "Book Nature": 38, "GSTIN Same?": 22}
            for i, c in enumerate(cols, start=1):
                ws.column_dimensions[get_column_letter(i)].width = widths.get(c, 15)
            if name == "Summary":
                last = hdr + len(data)
                for cell in ws[last]:
                    cell.font = Font(bold=True)


if __name__ == "__main__":
    books = load_books()
    t2b = load_2b()
    print(f"Books rows: {len(books)}   GSTR-2B B2B rows: {len(t2b)}")
    pairs = reconcile(books, t2b)
    df = build_output(pairs)
    summ = summary(df)
    write_excel(df, summ)
    print(summ.to_string(index=False))
    print("Written ->", OUT_FILE)
