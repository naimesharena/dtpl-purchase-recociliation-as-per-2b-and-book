"""Parser for GSTR-2B Excel exports (quarterly file from GST portal).

The file is in legacy .xls (BIFF) format read with xlrd.

Sheets of interest:
  * B2B            — Regular invoices (ITC available)
  * B2B-CDNR       — Original credit / debit notes
  * B2BA           — Amendments to invoices
  * B2B-CDNRA      — Amendments to credit/debit notes
  * B2B(Rejected)  — Invoices where ITC is rejected
  * B2B-CDNR(Rejected) — Credit/debit notes where ITC is rejected

We also skip the "Read me" and "ITC *" summary sheets.

Each sheet has a 4/5/6 row header that is hard-coded per GSTR-2B layout.
"""
from __future__ import annotations

from datetime import date, datetime
from typing import List, Optional

import xlrd

from .models import NormalizedTransaction, DocumentType


def _to_float(v) -> float:
    if v is None or v == "":
        return 0.0
    try:
        f = float(v)
        return 0.0 if str(f) == "nan" else f
    except (TypeError, ValueError):
        return 0.0


def _parse_gstr_date(v, datemode: int = 0) -> Optional[date]:
    if v is None or v == "":
        return None
    if isinstance(v, float):
        # xlrd date as Excel serial
        try:
            t = xlrd.xldate_as_tuple(v, datemode)
            return date(t[0], t[1], t[2])
        except Exception:
            return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        s = v.strip()
        for fmt in ("%d-%b-%Y", "%d-%m-%Y", "%d/%m/%Y", "%d-%b-%y"):
            try:
                return datetime.strptime(s, fmt).date()
            except ValueError:
                continue
    return None


def _month_str_from_period(period: str) -> str:
    """GSTR-1 period like '042025' → '042025' (already MMYYYY)."""
    return str(period).strip()


def _month_str(d: Optional[date]) -> str:
    return f"{d.month:02d}{d.year}" if d else ""


class GSTR2BParser:
    """Parse GSTR-2B .xls into normalized transactions."""

    def __init__(self, path: str):
        self.path = path

    def parse_all(self) -> List[NormalizedTransaction]:
        wb = xlrd.open_workbook(self.path)
        datemode = wb.datemode
        txs: List[NormalizedTransaction] = []

        # B2B
        if "B2B" in wb.sheet_names():
            txs.extend(self._parse_b2b(wb.sheet_by_name("B2B"), datemode))
        # B2B-CDNR
        if "B2B-CDNR" in wb.sheet_names():
            txs.extend(self._parse_cdnr(wb.sheet_by_name("B2B-CDNR"), datemode))
        # B2BA (amendments)
        if "B2BA" in wb.sheet_names():
            txs.extend(self._parse_b2ba(wb.sheet_by_name("B2BA"), datemode))
        # B2B-CDNRA
        if "B2B-CDNRA" in wb.sheet_names():
            txs.extend(self._parse_cdnra(wb.sheet_by_name("B2B-CDNRA"), datemode))
        # Rejected sheets (ITC not eligible)
        for rej_name in ["B2B(Rejected)", "B2B-CDNR(Rejected)"]:
            if rej_name in wb.sheet_names():
                is_cdnr = "CDNR" in rej_name
                txs.extend(self._parse_rejected(wb.sheet_by_name(rej_name), datemode, is_cdnr=is_cdnr))

        return txs

    # ------------------------------------------------------------------
    # B2B (regular invoices)
    # ------------------------------------------------------------------
    def _parse_b2b(self, sh, datemode: int) -> List[NormalizedTransaction]:
        # Header ends at row 5, data starts row 6.
        # Cols:
        #  0=GSTIN, 1=name, 2=inv_no, 3=inv_type, 4=inv_date, 5=inv_val,
        #  6=pos, 7=rcm, 8=taxable, 9=igst, 10=cgst, 11=sgst, 12=cess,
        # 13=period, 14=filing_date, 15=itc_avail, ...
        out = []
        for i in range(6, sh.nrows):
            gstin = str(sh.cell_value(i, 0)).strip()
            if not gstin or not gstin[0].isalnum():
                continue
            inv_no = str(sh.cell_value(i, 2)).strip()
            if not inv_no:
                continue
            inv_date = _parse_gstr_date(sh.cell_value(i, 4), datemode)
            period = _month_str_from_period(sh.cell_value(i, 13))
            itc = str(sh.cell_value(i, 15)).strip()
            out.append(NormalizedTransaction(
                source="GSTR2B",
                location="ALL",
                supplier_gstin=gstin.upper(),
                supplier_name=str(sh.cell_value(i, 1)).strip(),
                document_number=inv_no,
                document_type=DocumentType.INVOICE,
                document_date=inv_date,
                document_month=_month_str(inv_date),
                gstr2b_month=period,
                invoice_value=_to_float(sh.cell_value(i, 5)),
                taxable_value=_to_float(sh.cell_value(i, 8)),
                igst=_to_float(sh.cell_value(i, 9)),
                cgst=_to_float(sh.cell_value(i, 10)),
                sgst=_to_float(sh.cell_value(i, 11)),
                cess=_to_float(sh.cell_value(i, 12)),
                itc_available=itc,
                place_of_supply=str(sh.cell_value(i, 6)).strip(),
                reverse_charge=str(sh.cell_value(i, 7)).strip(),
            ))
        return out

    # ------------------------------------------------------------------
    # B2B-CDNR (credit / debit notes)
    # ------------------------------------------------------------------
    def _parse_cdnr(self, sh, datemode: int) -> List[NormalizedTransaction]:
        # Cols:
        #  0=gstin, 1=name, 2=note_no, 3=note_type (C/D), 4=supply_type,
        #  5=note_date, 6=note_val, 7=pos, 8=rcm, 9=taxable,
        # 10=igst, 11=cgst, 12=sgst, 13=cess, 14=period, 15=filing_date,
        # 16=itc_avail
        out = []
        for i in range(6, sh.nrows):
            gstin = str(sh.cell_value(i, 0)).strip()
            if not gstin or not gstin[0].isalnum():
                continue
            note_no = str(sh.cell_value(i, 2)).strip()
            if not note_no:
                continue
            note_type = str(sh.cell_value(i, 3)).strip().upper()
            if note_type == "C":
                dtype = DocumentType.CREDIT_NOTE
            elif note_type == "D":
                dtype = DocumentType.DEBIT_NOTE
            else:
                dtype = DocumentType.CREDIT_NOTE  # default to credit note if unknown
            note_date = _parse_gstr_date(sh.cell_value(i, 5), datemode)
            period = _month_str_from_period(sh.cell_value(i, 14))
            itc = str(sh.cell_value(i, 16)).strip()
            out.append(NormalizedTransaction(
                source="GSTR2B",
                location="ALL",
                supplier_gstin=gstin.upper(),
                supplier_name=str(sh.cell_value(i, 1)).strip(),
                document_number=note_no,
                document_type=dtype,
                document_date=note_date,
                document_month=_month_str(note_date),
                gstr2b_month=period,
                invoice_value=_to_float(sh.cell_value(i, 6)),
                taxable_value=_to_float(sh.cell_value(i, 9)),
                igst=_to_float(sh.cell_value(i, 10)),
                cgst=_to_float(sh.cell_value(i, 11)),
                sgst=_to_float(sh.cell_value(i, 12)),
                cess=_to_float(sh.cell_value(i, 13)),
                itc_available=itc,
                place_of_supply=str(sh.cell_value(i, 7)).strip(),
                reverse_charge=str(sh.cell_value(i, 8)).strip(),
            ))
        return out

    # ------------------------------------------------------------------
    # B2BA (amendments to invoices)
    # ------------------------------------------------------------------
    def _parse_b2ba(self, sh, datemode: int) -> List[NormalizedTransaction]:
        # Original in cols 0..1; Revised starts at col 4 (invoice number).
        # Cols from row 6:
        #  0=orig_inv_no, 1=orig_inv_date, 2=gstin, 3=name,
        #  4=rev_inv_no, 5=rev_type, 6=rev_date, 7=rev_value, 8=pos, 9=rcm,
        # 10=taxable, 11=igst, 12=cgst, 13=sgst, 14=cess,
        # 15=period, 16=filing, 17=itc, 18=reason, 19=rate
        out = []
        for i in range(7, sh.nrows):
            gstin = str(sh.cell_value(i, 2)).strip()
            if not gstin or not gstin[0].isalnum():
                continue
            rev_no = str(sh.cell_value(i, 4)).strip()
            if not rev_no:
                continue
            rev_date = _parse_gstr_date(sh.cell_value(i, 6), datemode)
            period = _month_str_from_period(sh.cell_value(i, 15))
            itc = str(sh.cell_value(i, 17)).strip()
            out.append(NormalizedTransaction(
                source="GSTR2B",
                location="ALL",
                supplier_gstin=gstin.upper(),
                supplier_name=str(sh.cell_value(i, 3)).strip(),
                document_number=rev_no,
                document_type=DocumentType.INVOICE,
                document_date=rev_date,
                document_month=_month_str(rev_date),
                gstr2b_month=period,
                invoice_value=_to_float(sh.cell_value(i, 7)),
                taxable_value=_to_float(sh.cell_value(i, 10)),
                igst=_to_float(sh.cell_value(i, 11)),
                cgst=_to_float(sh.cell_value(i, 12)),
                sgst=_to_float(sh.cell_value(i, 13)),
                cess=_to_float(sh.cell_value(i, 14)),
                itc_available=itc,
                place_of_supply=str(sh.cell_value(i, 8)).strip(),
                reverse_charge=str(sh.cell_value(i, 9)).strip(),
                particulars="AMENDMENT",
            ))
        return out

    # ------------------------------------------------------------------
    # B2B-CDNRA (amendments to notes)
    # ------------------------------------------------------------------
    def _parse_cdnra(self, sh, datemode: int) -> List[NormalizedTransaction]:
        # Row 5/6 headers:
        # 0=orig_note_type, 1=orig_note_no, 2=orig_note_date, 3=gstin, 4=name,
        # 5=rev_note_no, 6=rev_note_type, 7=rev_supply_type, 8=rev_note_date,
        # 9=rev_note_val, 10=pos, 11=rcm, 12=taxable, 13=igst, 14=cgst, 15=sgst,
        # 16=cess, 17=period, 18=filing, 19=itc, 20=reason, 21=rate
        out = []
        for i in range(7, sh.nrows):
            gstin = str(sh.cell_value(i, 3)).strip()
            if not gstin or not gstin[0].isalnum():
                continue
            rev_no = str(sh.cell_value(i, 5)).strip()
            if not rev_no:
                continue
            note_type = str(sh.cell_value(i, 6)).strip().upper()
            dtype = DocumentType.CREDIT_NOTE if note_type == "C" else (
                DocumentType.DEBIT_NOTE if note_type == "D" else DocumentType.CREDIT_NOTE
            )
            rev_date = _parse_gstr_date(sh.cell_value(i, 8), datemode)
            period = _month_str_from_period(sh.cell_value(i, 17))
            itc = str(sh.cell_value(i, 19)).strip()
            out.append(NormalizedTransaction(
                source="GSTR2B",
                location="ALL",
                supplier_gstin=gstin.upper(),
                supplier_name=str(sh.cell_value(i, 4)).strip(),
                document_number=rev_no,
                document_type=dtype,
                document_date=rev_date,
                document_month=_month_str(rev_date),
                gstr2b_month=period,
                invoice_value=_to_float(sh.cell_value(i, 9)),
                taxable_value=_to_float(sh.cell_value(i, 12)),
                igst=_to_float(sh.cell_value(i, 13)),
                cgst=_to_float(sh.cell_value(i, 14)),
                sgst=_to_float(sh.cell_value(i, 15)),
                cess=_to_float(sh.cell_value(i, 16)),
                itc_available=itc,
                place_of_supply=str(sh.cell_value(i, 10)).strip(),
                reverse_charge=str(sh.cell_value(i, 11)).strip(),
                particulars="AMENDMENT",
            ))
        return out

    # ------------------------------------------------------------------
    # Rejected sheets
    # ------------------------------------------------------------------
    def _parse_rejected(self, sh, datemode: int, is_cdnr: bool) -> List[NormalizedTransaction]:
        # Use same column mapping as B2B / CDNR; header ends at row 5.
        # We mark itc_available as "No" so downstream reports can separate.
        if is_cdnr:
            return self._parse_cdnr(sh, datemode)  # columns same layout as CDNR
        return self._parse_b2b(sh, datemode)
