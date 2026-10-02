"""
Purchase reconciliation : Books (Tally purchase registers) vs GSTR-2B (B2B sheet).

Handles the client's data quirks:
  * Non-GST purchases in books (no CGST/SGST/IGST) -> will never appear in 2B.
  * Ineligible-credit bills booked at TOTAL amount (tax merged in value, tax cols blank).
  * Same supplier registered in multiple states -> invoice reported under a
    different GSTIN of the SAME PAN. Matched on PAN + invoice no + amount.

Sheets produced in PURCHASE-RECO-BOOK-VS-2B.xlsx
  * Summary                 - status wise counts & values of both reconciliations
  * Book vs 2B Reco         - purchase registers  vs  2B B2B (B2BA amendments applied)
  * Debit Note vs 2B CDNR   - debit note registers vs  2B B2B-CDNR
  * Other 2B Sheets         - B2B(Rejected) / B2B-CDNR(Rejected) / B2B-CDNRA(Rejected)
                              / ECO, each checked against the books
"""

import glob
import os
import re
import pandas as pd
import numpy as np
from datetime import datetime

BOOK_FILES = {
    "HR": "PURCHASE-HR.xls",
    "PL": "PURCHASE-PL.xls",
    "VL": "PURCHASE-VL.xls",
}
GSTR2B_DIR = "GSTR-2B"      # monthly GSTR-2B downloads (xlsx), consolidated
GSTR2B_FILE = "gstr-2B.xls"  # legacy single consolidated file (fallback only)
DN_FILES = {
    "HR": "DEBIT NOTE-HR.xls",
    "PL": "DEBIT NOTE-PL.xls",
    "VL": "DEBIT NOTE-VL.xls",
}
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


def _read_b2b_sheet(path, engine=None, skiprows=6):
    df = pd.read_excel(path, sheet_name="B2B", header=None,
                       engine=engine, skiprows=skiprows)
    cols = ["GSTIN", "Name", "InvNo", "InvType", "InvDate", "InvValue",
            "POS", "RCM", "Taxable", "IGST", "CGST", "SGST", "Cess",
            "Period", "FileDate", "ITCAvail", "Reason", "ApplPct",
            "Source", "IRN", "IRNDate"]
    df = df.iloc[:, :len(cols)]
    df.columns = cols[:df.shape[1]]
    return df[df["GSTIN"].astype(str).str.strip().str.len() == 15].copy()


def _period_label(fname):
    """012026_... -> Jan-2026"""
    m = re.match(r"(\d{2})(\d{4})", os.path.basename(fname))
    if not m:
        return ""
    mm, yy = int(m.group(1)), m.group(2)
    return f"{MONTHS[mm - 1]}-{yy}"


MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def _sheet_rows(path, sheet, skiprows):
    """Rows of a 2B sheet that actually carry a supplier GSTIN."""
    x = pd.ExcelFile(path)
    if sheet not in x.sheet_names:
        return None
    d = pd.read_excel(x, sheet, header=None, skiprows=skiprows)
    keep = d.apply(lambda r: any(isinstance(v, str) and len(str(v).strip()) == 15
                                 and str(v).strip()[:2].isdigit() for v in r.iloc[:4]),
                   axis=1)
    d = d[keep]
    return d if len(d) else None


def load_b2ba():
    """B2BA - amendments. Revised figures replace the original B2B invoice."""
    rows = []
    for f in sorted(glob.glob(os.path.join(GSTR2B_DIR, "*.xls*"))):
        d = _sheet_rows(f, "B2BA", 6)
        if d is None:
            continue
        for _, r in d.iterrows():
            rows.append({
                "orig_inv": r.iloc[0], "orig_date": r.iloc[1],
                "2B GSTIN": str(r.iloc[2]).strip().upper(),
                "2B Supplier": str(r.iloc[3]).strip(),
                "2B Invoice No.": str(r.iloc[4]).strip(),
                "2B Invoice Date": r.iloc[6],
                "2B Invoice Value": r.iloc[7],
                "2B Taxable": r.iloc[10],
                "2B IGST": r.iloc[11], "2B CGST": r.iloc[12],
                "2B SGST": r.iloc[13], "2B Cess": r.iloc[14],
                "2B ITC Available": r.iloc[23],
                "2B Period": r.iloc[21],
                "2B Source File": os.path.basename(f),
            })
    return pd.DataFrame(rows)


def load_cdnr():
    """B2B-CDNR - supplier credit / debit notes (books: debit note register)."""
    rows = []
    for f in sorted(glob.glob(os.path.join(GSTR2B_DIR, "*.xls*"))):
        d = _sheet_rows(f, "B2B-CDNR", 6)
        if d is None:
            continue
        for _, r in d.iterrows():
            rows.append({
                "2B GSTIN": str(r.iloc[0]).strip().upper(),
                "2B Supplier": str(r.iloc[1]).strip(),
                "2B Invoice No.": str(r.iloc[2]).strip(),
                "2B Note Type": r.iloc[3],
                "2B Invoice Date": r.iloc[5],
                "2B Invoice Value": r.iloc[6],
                "2B Taxable": r.iloc[9],
                "2B IGST": r.iloc[10], "2B CGST": r.iloc[11],
                "2B SGST": r.iloc[12], "2B Cess": r.iloc[13],
                "2B ITC Available": r.iloc[22],
                "2B Period": r.iloc[20],
                "2B Source File": os.path.basename(f),
            })
    t = pd.DataFrame(rows)
    return _finalise_2b(t)


OTHER_SHEETS = {
    # sheet : (col positions - gstin, name, docno, doctype, docdate, docvalue,
    #                          taxable, igst, cgst, sgst, cess, period, skiprows)
    "B2B(Rejected)":        dict(c=(0, 1, 2, 3, 4, 5, 7, 8, 9, 10, 11, 13), skip=6),
    "B2B-CDNR(Rejected)":   dict(c=(0, 1, 2, 3, 5, 6, 8, 9, 10, 11, 12, 14), skip=6),
    "B2B-CDNRA(Rejected)":  dict(c=(3, 4, 5, 6, 8, 9, 11, 12, 13, 14, 15, 17), skip=6),
    "ECO":                  dict(c=(0, 1, 2, 3, 4, 5, 7, 8, 9, 10, 11, 12), skip=6),
}


def load_other_sheets():
    out = []
    for sheet, cfg in OTHER_SHEETS.items():
        for f in sorted(glob.glob(os.path.join(GSTR2B_DIR, "*.xls*"))):
            d = _sheet_rows(f, sheet, cfg["skip"])
            if d is None:
                continue
            c = cfg["c"]
            for _, r in d.iterrows():
                out.append({
                    "2B Sheet": sheet,
                    "2B GSTIN": str(r.iloc[c[0]]).strip().upper(),
                    "2B Supplier": str(r.iloc[c[1]]).strip(),
                    "2B Invoice No.": str(r.iloc[c[2]]).strip(),
                    "2B Doc Type": r.iloc[c[3]],
                    "2B Invoice Date": r.iloc[c[4]],
                    "2B Invoice Value": r.iloc[c[5]],
                    "2B Taxable": r.iloc[c[6]],
                    "2B IGST": r.iloc[c[7]], "2B CGST": r.iloc[c[8]],
                    "2B SGST": r.iloc[c[9]], "2B Cess": r.iloc[c[10]],
                    "2B Period": r.iloc[c[11]],
                    "2B Source File": os.path.basename(f),
                })
    return _finalise_2b(pd.DataFrame(out))


def _finalise_2b(t):
    """Common numeric / key preparation for any 2B style frame."""
    if t.empty:
        t["_ti"] = []
        return t
    t["2B Invoice Date"] = to_date(t["2B Invoice Date"])
    for c in ["2B Invoice Value", "2B Taxable", "2B IGST", "2B CGST", "2B SGST", "2B Cess"]:
        t[c] = to_num(t[c])
    t["2B Tax"] = t["2B IGST"] + t["2B CGST"] + t["2B SGST"] + t["2B Cess"]
    t["2B Period"] = t["2B Period"].astype(str).str.strip().replace({"nan": ""})
    if "2B ITC Available" in t:
        t["2B ITC Available"] = t["2B ITC Available"].astype(str).str.strip().replace({"nan": ""})
    t["PAN"] = t["2B GSTIN"].map(pan_of)
    t["ikey"] = t["2B Invoice No."].map(norm_inv)
    t["idig"] = t["2B Invoice No."].map(inv_digits)
    t = t.reset_index(drop=True)
    t["_ti"] = range(len(t))
    return t


def apply_amendments(t2b, amd):
    """Replace amended invoices in the B2B pool with their revised figures."""
    if amd.empty:
        return t2b, 0
    amd = amd.copy()
    amd["okey"] = amd["orig_inv"].map(norm_inv)
    amd["PAN"] = amd["2B GSTIN"].map(pan_of)
    t2b = t2b.copy()
    t2b["_drop"] = False
    added = []
    for _, a in amd.iterrows():
        hit = (t2b["PAN"] == a["PAN"]) & (t2b["ikey"] == a["okey"])
        t2b.loc[hit, "_drop"] = True
        row = {k: a[k] for k in a.index if k.startswith("2B")}
        row["2B Period"] = f"{a['2B Period']} (B2BA amended)"
        added.append(row)
    n = int(t2b["_drop"].sum())
    t2b = t2b[~t2b["_drop"]].drop(columns="_drop")
    new = _finalise_2b(pd.DataFrame(added))
    out = pd.concat([t2b, new], ignore_index=True)
    out = out.reset_index(drop=True)
    out["_ti"] = range(len(out))
    return out, n


def load_2b():
    files = sorted(glob.glob(os.path.join(GSTR2B_DIR, "*.xls*")))
    parts = []
    if files:
        for f in files:
            d = _read_b2b_sheet(f)
            d["SrcFile"] = os.path.basename(f)
            d["Period2"] = _period_label(f)
            parts.append(d)
            print(f"  2B {os.path.basename(f)} : {len(d)} B2B rows")
    else:
        d = _read_b2b_sheet(GSTR2B_FILE, engine="xlrd")
        d["SrcFile"] = GSTR2B_FILE
        d["Period2"] = ""
        parts.append(d)
    df = pd.concat(parts, ignore_index=True)

    t = pd.DataFrame({
        "2B GSTIN": df["GSTIN"].astype(str).str.strip().str.upper(),
        "2B Supplier": df["Name"].astype(str).str.strip(),
        "2B Invoice No.": df["InvNo"].astype(str).str.strip(),
        "2B Invoice Date": to_date(df["InvDate"]),
        "2B Invoice Value": to_num(df["InvValue"]),
        "2B Taxable": to_num(df["Taxable"]),
        "2B IGST": to_num(df["IGST"]),
        "2B CGST": to_num(df["CGST"]),
        "2B SGST": to_num(df["SGST"]),
        "2B Cess": to_num(df["Cess"]),
        "2B ITC Available": df["ITCAvail"].astype(str).str.strip()
                              .replace({"nan": ""}),
        "2B Period": df["Period2"].astype(str).str.strip(),
        "2B Source File": df["SrcFile"],
    })
    # period from the 2B file itself if the file name could not be parsed
    raw_period = df["Period"].astype(str).str.strip()
    t.loc[t["2B Period"] == "", "2B Period"] = raw_period
    t["2B Tax"] = t["2B IGST"] + t["2B CGST"] + t["2B SGST"] + t["2B Cess"]
    t["PAN"] = t["2B GSTIN"].map(pan_of)
    t["ikey"] = t["2B Invoice No."].map(norm_inv)
    t["idig"] = t["2B Invoice No."].map(inv_digits)

    # a monthly download can repeat an invoice (e.g. re-filed) - drop duplicates
    before = len(t)
    t = t.drop_duplicates(subset=["2B GSTIN", "ikey", "2B Invoice Date",
                                  "2B Taxable", "2B Tax"], keep="first")
    if before != len(t):
        print(f"  removed {before - len(t)} duplicate 2B rows across months")
    t = t.reset_index(drop=True)
    t["_ti"] = range(len(t))
    return t


def load_debit_notes():
    """Debit note registers (books). Supplier credit notes appear here as DN."""
    TAXC = {"IGST TAX", "CGST TAX", "SGST TAX", "Round Off", "Gross Total"}
    SKIP = {"Date", "Particulars", "Voucher Type", "Voucher No.", "Voucher Ref. No.",
            "Voucher Ref. Date", "GSTIN/UIN", "Shipping / Bill of Entry No.",
            "Shipping / Bill of Entry Date", "Bill of Entry No.", "Bill of Entry Date"}
    frames = []
    for branch, f in DN_FILES.items():
        raw = pd.read_excel(f, header=None, engine="openpyxl", skiprows=9)
        raw.columns = [str(c).strip() for c in raw.iloc[0]]
        df = raw.iloc[1:].copy()
        df = df[df["Particulars"].astype(str).str.strip().str.lower() != "grand total"]
        df = df[df["Date"].notna()]
        if df.empty:
            continue
        out = pd.DataFrame({
            "Branch": branch,
            "Book Date": to_date(df["Date"]),
            "Supplier (Books)": df["Particulars"].astype(str).str.strip(),
            "Voucher Type": df["Voucher Type"],
            "Voucher No.": df["Voucher No."].astype(str).str.strip(),
            "Book Invoice No.": df["Voucher Ref. No."].astype(str).str.strip()
                                  .replace({"nan": ""}),
            "Book Invoice Date": to_date(df["Voucher Ref. Date"]),
            "Book GSTIN": df["GSTIN/UIN"].astype(str).str.strip().str.upper()
                            .replace({"NAN": "", "NONE": ""}),
            "Book Total": to_num(df["Gross Total"]),
            "Book CGST": to_num(df.get("CGST TAX")),
            "Book SGST": to_num(df.get("SGST TAX")),
            "Book IGST": to_num(df.get("IGST TAX")),
            "Book Round Off": to_num(df.get("Round Off")),
        })
        out["Book Tax"] = out["Book CGST"] + out["Book SGST"] + out["Book IGST"]
        ledger_cols = [c for c in df.columns if c not in TAXC and c not in SKIP]
        ledger = sum(to_num(df[c]) for c in ledger_cols) if ledger_cols else 0.0
        ledger = pd.Series(ledger, index=out.index).round(2)
        fallback = (out["Book Total"] - out["Book Tax"] - out["Book Round Off"]).round(2)
        out["Book Taxable"] = np.where(ledger.abs() > 0.009, ledger, fallback)
        out["Book Invoice Value"] = (out["Book Taxable"] + out["Book Tax"]
                                     + out["Book Round Off"]).round(2)
        out["Book TDS / Deduction"] = (out["Book Invoice Value"] - out["Book Total"]).round(2)
        out.loc[out["Book TDS / Deduction"].abs() < 0.01, "Book TDS / Deduction"] = 0.0
        frames.append(out)

    b = pd.concat(frames, ignore_index=True)
    b["No Tax in Books"] = b["Book Tax"].abs() < 0.01
    b["Book Nature"] = np.where(~b["No Tax in Books"], "GST Debit Note",
                                np.where(b["Book GSTIN"].str.len() == 15,
                                         "Debit Note booked at Total (no tax split)",
                                         "Non-GST Debit Note"))
    b["PAN"] = b["Book GSTIN"].map(pan_of)
    b["ikey"] = b["Book Invoice No."].map(norm_inv)
    b["vkey"] = b["Voucher No."].map(norm_inv)
    b["idig"] = b["Book Invoice No."].map(inv_digits)
    b["vdig"] = b["Voucher No."].map(inv_digits)
    b["_bi"] = range(len(b))
    return b


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




def check_other_sheets(other, books, dnotes):
    """Flag whether each rejected / ECO document is present in the books."""
    if other.empty:
        return other
    pool = []
    for src, df in (("Purchase Register", books), ("Debit Note Register", dnotes)):
        for r in df.to_dict("records"):
            pool.append((r["PAN"], r["ikey"], r["vkey"], r["Book Invoice Value"],
                         src, r["Supplier (Books)"], r["Branch"],
                         r["Book Invoice No."]))
    rows = []
    for _, o in other.iterrows():
        found, where, bval, bsup, bbr = "Not in Books", "", None, "", ""
        for pan, ik, vk, val, src, sup, br, inv in pool:
            if pan and pan == o["PAN"] and o["ikey"] and o["ikey"] in (ik, vk):
                found, where, bval, bsup, bbr = "Present in Books", src, val, sup, br
                break
        if found == "Not in Books":
            for pan, ik, vk, val, src, sup, br, inv in pool:
                if (pan and pan == o["PAN"]
                        and abs(float(val) - float(o["2B Invoice Value"])) <= AMT_TOL):
                    found, where, bval, bsup, bbr = ("Present in Books (value match)",
                                                     src, val, sup, br)
                    break
        rows.append({"Status in Books": found, "Found In": where,
                     "Book Branch": bbr, "Book Supplier": bsup,
                     "Book Invoice Value": bval})
    return pd.concat([other.reset_index(drop=True), pd.DataFrame(rows)], axis=1)


# ----------------------------------------------------------------- output
BOOK_COLS = ["Branch", "Book Date", "Supplier (Books)", "Book GSTIN", "Voucher Type",
             "Voucher No.", "Book Invoice No.", "Book Invoice Date", "Book Taxable",
             "Book IGST", "Book CGST", "Book SGST", "Book Tax", "Book Invoice Value",
             "Book TDS / Deduction", "Book Total", "Book Nature"]
T2B_COLS = ["2B Source File", "2B GSTIN", "2B Supplier", "2B Invoice No.", "2B Invoice Date",
            "2B Taxable", "2B IGST", "2B CGST", "2B SGST", "2B Tax",
            "2B Invoice Value", "2B ITC Available", "2B Period"]

TEXT_COLS = {"2B Period", "2B Source File", "Book Invoice No.", "2B Invoice No.",
             "Voucher No.", "Book GSTIN", "2B GSTIN", "Status", "Match Remark",
             "Branch", "Supplier (Books)", "2B Supplier", "Voucher Type",
             "Book Nature", "GSTIN Same?", "2B ITC Available", "2B Sheet",
             "Status in Books", "Found In", "Book Supplier", "Book Branch",
             "2B Doc Type", "2B Note Type", "Reconciliation"}

STATUS_ORDER = ["Matched", "Matched - GSTIN Mismatch (Same PAN)",
                "Matched with Difference", "Probable Match - Invoice No. Differs",
                "Probable Match - Amount Based", "In Books Only",
                "In Books Only - Booked at Total", "In Books Only - Non-GST",
                "In GSTR-2B Only",
                "Other 2B Sheet - Present in Books",
                "Other 2B Sheet - Not in Books"]


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
    df = df.copy()
    if "Sr." not in df:
        df.insert(0, "Sr.", range(1, len(df) + 1))
    for c in ["Book Taxable", "Book Tax", "Book Invoice Value",
              "Book TDS / Deduction", "Book Total", "2B Taxable", "2B Tax",
              "2B Invoice Value"]:
        if c not in df:
            df[c] = 0.0
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
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


def other_to_reco_rows(other):
    """Map rejected / ECO documents into the common reconciliation layout."""
    if other.empty:
        return pd.DataFrame()
    rows = []
    for _, o in other.iterrows():
        present = str(o["Status in Books"]).startswith("Present")
        rows.append({
            "Reconciliation": f"Other 2B Sheet - {o['2B Sheet']}",
            "Status": ("Other 2B Sheet - Present in Books" if present
                       else "Other 2B Sheet - Not in Books"),
            "Match Remark": (f"{o['2B Sheet']} document - {o['Status in Books']}"
                             + (f" ({o['Found In']})" if o["Found In"] else "")),
            "Branch": o["Book Branch"], "Supplier (Books)": o["Book Supplier"],
            "Book Invoice Value": o["Book Invoice Value"],
            "2B GSTIN": o["2B GSTIN"], "2B Supplier": o["2B Supplier"],
            "2B Invoice No.": o["2B Invoice No."],
            "2B Invoice Date": o["2B Invoice Date"],
            "2B Taxable": o["2B Taxable"], "2B IGST": o["2B IGST"],
            "2B CGST": o["2B CGST"], "2B SGST": o["2B SGST"],
            "2B Tax": o["2B Tax"], "2B Invoice Value": o["2B Invoice Value"],
            "2B Period": o["2B Period"], "2B Source File": o["2B Source File"],
        })
    return pd.DataFrame(rows)


def write_excel(sheets, summ, other):
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
        "Other 2B Sheet - Present in Books": "FFF2CC",
        "Other 2B Sheet - Not in Books": "F2F2F2",
        "Present in Books": "C6EFCE",
        "Present in Books (value match)": "FFF2CC",
        "Not in Books": "FFC7CE",
    }
    thin = Side(style="thin", color="BFBFBF")

    with pd.ExcelWriter(OUT_FILE, engine="openpyxl") as xl:
        summ.to_excel(xl, sheet_name="Summary", index=False, startrow=2)
        for name, data in sheets:
            data.to_excel(xl, sheet_name=name, index=False, startrow=2)

        allsheets = [("Summary", summ)] + list(sheets)

        for name, data in allsheets:
            ws = xl.sheets[name]
            ws["A1"] = ("DRIVE TRUCKING PVT LTD - Purchase Reconciliation : "
                        "As per Books vs As per GSTR-2B  |  FY 2025-26")
            ws["A1"].font = Font(bold=True, size=13)
            ws["A2"] = ("Books: PURCHASE / DEBIT NOTE - HR, PL, VL   |   "
                        "2B: 12 monthly downloads consolidated   |   "
                        f"Generated {datetime.now():%d-%m-%Y %H:%M}")
            ws["A2"].font = Font(italic=True, size=9, color="555555")

            hdr = 3
            for cell in ws[hdr]:
                if cell.value is None:
                    continue
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="305496")
                cell.alignment = Alignment(horizontal="center", vertical="center",
                                           wrap_text=True)
                cell.border = Border(thin, thin, thin, thin)
            ws.freeze_panes = ws.cell(row=hdr + 1, column=4 if name != "Summary" else 2)
            ws.auto_filter.ref = ws.dimensions

            cols = list(data.columns)
            scol = "Status" if "Status" in cols else (
                "Status in Books" if "Status in Books" in cols else None)
            si = cols.index(scol) + 1 if scol else None
            for row in ws.iter_rows(min_row=hdr + 1, max_row=hdr + len(data)):
                f = fills.get(ws.cell(row=row[0].row, column=si).value) if si else None
                for cell in row:
                    cell.border = Border(thin, thin, thin, thin)
                    if f:
                        cell.fill = PatternFill("solid", fgColor=f)
                    col_name = cols[cell.column - 1] if cell.column <= len(cols) else ""
                    if col_name in TEXT_COLS:
                        cell.number_format = "@"
                        cell.alignment = Alignment(horizontal="left")
                    elif col_name in ("Sr.", "No. of Records"):
                        cell.number_format = "0"
                    elif isinstance(cell.value, datetime):
                        cell.number_format = "dd-mm-yyyy"
                    elif isinstance(cell.value, (int, float)):
                        cell.number_format = "#,##0.00"

            widths = {"Reconciliation": 26, "2B Source File": 34, "2B Period": 12, "Sr.": 7,
                      "Book TDS / Deduction": 18, "Book Invoice Value": 16,
                      "Status": 34, "Match Remark": 48, "Supplier (Books)": 30,
                      "2B Supplier": 30, "Book GSTIN": 17, "2B GSTIN": 17,
                      "Book Invoice No.": 20, "2B Invoice No.": 20,
                      "Book Nature": 38, "GSTIN Same?": 22, "2B Sheet": 22,
                      "Status in Books": 26, "Found In": 20, "Book Supplier": 28}
            for i, c in enumerate(cols, start=1):
                ws.column_dimensions[get_column_letter(i)].width = widths.get(c, 15)
            if name == "Summary":
                for cell in ws[hdr + len(data)]:
                    cell.font = Font(bold=True)


def run_pair(books, t2b, label):
    pairs = reconcile(books, t2b)
    df = build_output(pairs)
    df.insert(1, "Reconciliation", label)
    s = summary(df)
    s.insert(0, "Reconciliation", label)
    return df, s


if __name__ == "__main__":
    books = load_books()
    t2b = load_2b()

    amd = load_b2ba()
    t2b, n_amended = apply_amendments(t2b, amd)
    print(f"B2BA amendments applied: {len(amd)} (replaced {n_amended} original B2B rows)")

    print(f"Books rows: {len(books)}   GSTR-2B B2B rows: {len(t2b)}")
    df_b2b, s_b2b = run_pair(books, t2b, "Purchase vs B2B")

    dnotes = load_debit_notes()
    cdnr = load_cdnr()
    print(f"Debit notes in books: {len(dnotes)}   2B B2B-CDNR rows: {len(cdnr)}")
    df_dn, s_dn = run_pair(dnotes, cdnr, "Debit Note vs B2B-CDNR")

    other = check_other_sheets(load_other_sheets(), books, dnotes)
    df_ot = other_to_reco_rows(other)
    print(f"Other 2B sheets (rejected / ECO): {len(df_ot)} documents")

    # ---- everything in ONE sheet
    full = pd.concat([df_b2b, df_dn, df_ot], ignore_index=True)
    order = {"Purchase vs B2B": 0, "Debit Note vs B2B-CDNR": 1}
    full["_r"] = full["Reconciliation"].map(order).fillna(2)
    full["_s"] = full["Status"].map({s: i for i, s in enumerate(STATUS_ORDER)}).fillna(99)
    full = full.sort_values(["_r", "_s", "Book Date", "2B Invoice Date"],
                            na_position="last").drop(columns=["_r", "_s"])
    full["Sr."] = range(1, len(full) + 1)
    cols = ["Sr.", "Reconciliation"] + [c for c in full.columns
                                        if c not in ("Sr.", "Reconciliation")]
    full = full[cols]

    s_ot = summary(df_ot) if not df_ot.empty else pd.DataFrame()
    if not s_ot.empty:
        s_ot.insert(0, "Reconciliation", "Other 2B Sheets (Rejected / ECO)")
    summ = pd.concat([s_b2b, s_dn, s_ot], ignore_index=True)

    write_excel([("Book vs 2B Reco", full)], summ, df_ot)

    print(summ.to_string(index=False))
    print("Written ->", OUT_FILE)
