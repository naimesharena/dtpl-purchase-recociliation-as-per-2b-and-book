"""Parser for Tally-exported Books data (Purchase Register and Debit Note Register).

The repository contains 3 branch/location exports for purchases (HR, PL, VL) and 3
for debit notes. All are `.xls` that are actually xlsx files (openpyxl).

The header row is at row index 9 (0-based) in every sheet. The columns vary slightly
between locations in the order of Purchase Accounts / Indirect Expenses / Direct
Expenses / Fixed Assets, but we detect them by header name.

For every purchase ledger row we compute:
  * taxable_value = Purchase Accounts + Indirect Expenses + Direct Expenses + Fixed Assets
                   (the ledger value on which GST was calculated)
  * cgst, sgst, igst = SGST TAX, CGST TAX, IGST TAX columns respectively.
    NOTE: Tally labels the central-tax column "CGST TAX" and state-tax "SGST TAX";
    we preserve those labels.

For Debit Note Register rows we compute taxable_value from whichever expense/purchase
columns are populated (the register has many ad-hoc columns). These are classified as
CREDIT_NOTE documents (they correspond to supplier credit notes and reduce ITC).
"""
from __future__ import annotations

import os
import glob
from datetime import date, datetime
from typing import List, Optional

import pandas as pd
import numpy as np

from .models import NormalizedTransaction, DocumentType


_TAX_COLS = ["SGST TAX", "CGST TAX", "IGST TAX"]
_PURCHASE_VALUE_COLS = [
    "Purchase Accounts",
    "Indirect Expenses",
    "Direct Expenses",
    "Fixed Assets",
]


def _to_float(x) -> float:
    if x is None:
        return 0.0
    try:
        if isinstance(x, str) and not x.strip():
            return 0.0
        v = float(x)
        if pd.isna(v):
            return 0.0
        return v
    except (TypeError, ValueError):
        return 0.0


def _to_date(x) -> Optional[date]:
    if x is None:
        return None
    if isinstance(x, float):
        if pd.isna(x):
            return None
        # Excel serial date
        try:
            return pd.Timestamp.from_pydatetime(datetime(1899, 12, 30) + pd.Timedelta(days=int(x))).date()
        except Exception:
            return None
    if isinstance(x, (pd.Timestamp,)):
        try:
            return x.to_pydatetime().date()
        except Exception:
            return None
    # numpy datetime64 / anything pandas converts
    try:
        ts = pd.to_datetime(x, errors="coerce")
        if pd.isna(ts):
            return None
        return ts.date()
    except Exception:
        pass
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, date):
        return x
    if isinstance(x, str):
        s = x.strip()
        if not s:
            return None
        for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d-%b-%Y"):
            try:
                return datetime.strptime(s, fmt).date()
            except ValueError:
                pass
    return None


def _month_str(d) -> str:
    if d is None:
        return ""
    try:
        return f"{d.month:02d}{d.year}"
    except AttributeError:
        try:
            ts = pd.to_datetime(d, errors="coerce")
            if pd.isna(ts):
                return ""
            return f"{ts.month:02d}{ts.year}"
        except Exception:
            return ""


class BooksParser:
    """Parse Tally Purchase & Debit Note Excel exports into NormalizedTransactions."""

    def __init__(self, purchase_files: List[str] | None = None, debit_note_files: List[str] | None = None):
        self.purchase_files = purchase_files or []
        self.debit_note_files = debit_note_files or []

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def parse_all(self) -> List[NormalizedTransaction]:
        txs: List[NormalizedTransaction] = []
        for f in self.purchase_files:
            txs.extend(self._parse_purchase_file(f))
        for f in self.debit_note_files:
            txs.extend(self._parse_debit_note_file(f))
        return txs

    # ------------------------------------------------------------------
    # Purchase Register
    # ------------------------------------------------------------------
    def _parse_purchase_file(self, path: str) -> List[NormalizedTransaction]:
        loc = self._location_from_filename(path, "PURCHASE")
        df = pd.read_excel(path, sheet_name="Purchase Register", engine="openpyxl", header=9)
        df = df.dropna(subset=["Voucher Type", "Date"], how="any")
        txs: List[NormalizedTransaction] = []
        for _, row in df.iterrows():
            tx = self._row_to_purchase_tx(row, loc)
            if tx is not None:
                txs.append(tx)
        return txs

    def _row_to_purchase_tx(self, row, location: str) -> Optional[NormalizedTransaction]:
        doc_date = _to_date(row.get("Date"))
        gstin = str(row.get("GSTIN/UIN") or "").strip()
        if gstin.lower() in ("nan", "none"):
            gstin = ""

        # Taxable value = sum of purchase/expense/fixed-asset ledger amounts
        taxable = 0.0
        for c in _PURCHASE_VALUE_COLS:
            if c in row.index:
                taxable += _to_float(row.get(c))
        cgst = _to_float(row.get("CGST TAX"))
        sgst = _to_float(row.get("SGST TAX"))
        igst = _to_float(row.get("IGST TAX"))
        gross = _to_float(row.get("Gross Total"))

        # If taxable is 0 (rare: rounding-only entry), fall back to gross minus GST
        if taxable == 0.0 and gross:
            taxable = max(0.0, gross - (cgst + sgst + igst))

        # Document number selection:
        # For most suppliers the supplier's tax invoice number is entered as Voucher No.
        # in Tally (it's the number that appears on the physical tax invoice and hence
        # on GSTR-2B). For some suppliers "Supplier Invoice No." is entered as a
        # different internal order reference (e.g. Daimler order number), so Voucher No.
        # is the primary match key. We record Supplier Invoice No. as particulars for
        # reference.
        sup_inv = str(row.get("Supplier Invoice No.") or "").strip()
        vou_no = str(row.get("Voucher No.") or "").strip()
        if sup_inv.lower() in ("nan", "none"):
            sup_inv = ""
        if vou_no.lower() in ("nan", "none"):
            vou_no = ""
        # Primary = Voucher No. (tax invoice number reported in GSTR-2B).
        doc_no = vou_no or sup_inv
        if not doc_no:
            return None

        voucher_type = str(row.get("Voucher Type") or "").strip()
        particulars = str(row.get("Particulars") or "").strip()
        sup_inv_date = _to_date(row.get("Supplier Invoice Date"))
        doc_month = _month_str(sup_inv_date or doc_date)
        books_month = _month_str(doc_date)

        # Classification flags
        has_gstin = bool(gstin)
        has_gst = (cgst + sgst + igst) > 0.005  # any GST claimed
        is_non_gst = False
        is_ineligible = False
        if not has_gst:
            if not has_gstin:
                is_non_gst = True
            else:
                # Has GSTIN but no GST claimed → ineligible ITC
                # (e.g. purchases where ITC is blocked and not booked separately)
                is_ineligible = True

        return NormalizedTransaction(
            source="BOOKS",
            location=location,
            supplier_gstin=gstin.upper(),
            supplier_name=particulars,
            document_number=doc_no,
            document_type=DocumentType.INVOICE,
            document_date=sup_inv_date or doc_date,
            document_month=doc_month,
            books_month=books_month,
            invoice_value=gross,
            taxable_value=taxable,
            cgst=cgst,
            sgst=sgst,
            igst=igst,
            is_non_gst=is_non_gst,
            is_ineligible_itc=is_ineligible,
            voucher_type=voucher_type,
            voucher_no=vou_no,
            particulars=particulars,
        )

    # ------------------------------------------------------------------
    # Debit Note Register
    # ------------------------------------------------------------------
    def _parse_debit_note_file(self, path: str) -> List[NormalizedTransaction]:
        loc = self._location_from_filename(path, "DEBIT NOTE")
        df = pd.read_excel(path, sheet_name="Debit Note Register", engine="openpyxl", header=9)
        df = df.dropna(subset=["Voucher Type", "Date"], how="any")
        txs: List[NormalizedTransaction] = []
        for _, row in df.iterrows():
            tx = self._row_to_debit_note_tx(row, loc)
            if tx is not None:
                txs.append(tx)
        return txs

    def _row_to_debit_note_tx(self, row, location: str) -> Optional[NormalizedTransaction]:
        doc_date = _to_date(row.get("Date"))
        gstin = str(row.get("GSTIN/UIN") or "").strip()
        if gstin.lower() in ("nan", "none"):
            gstin = ""
        cgst = _to_float(row.get("CGST TAX"))
        sgst = _to_float(row.get("SGST TAX"))
        igst = _to_float(row.get("IGST TAX"))
        gross = _to_float(row.get("Gross Total"))
        voucher_no = str(row.get("Voucher No.") or "").strip()
        voucher_ref = str(row.get("Voucher Ref. No.") or "").strip()
        if voucher_ref.lower() in ("nan", "none"):
            voucher_ref = ""

        # Debit note register document number = Voucher Ref. No. when present (matches
        # supplier's credit note number), else Voucher No.
        doc_no = voucher_ref or voucher_no
        if not doc_no:
            return None

        particulars = str(row.get("Particulars") or "").strip()
        voucher_type = str(row.get("Voucher Type") or "").strip()
        ref_date = _to_date(row.get("Voucher Ref. Date"))
        doc_month = _month_str(ref_date or doc_date)
        books_month = _month_str(doc_date)

        # Taxable value for debit notes = gross - GST (since ledger columns may vary).
        # Sum any purchase/expense columns found to cross-check.
        taxable = max(0.0, gross - (cgst + sgst + igst))
        # Also sum recognizable columns (schema varies by location)
        extra_taxable = 0.0
        for col in row.index:
            cl = str(col).lower()
            if any(k in cl for k in ("dicv", "discount on purchase - gst", "internet service",
                                     "tools machinary", "electrical fitting",
                                     "electrical expenses", "spartpart", "rent expenses")):
                extra_taxable += _to_float(row.get(col))
        if extra_taxable and abs(extra_taxable - taxable) > 5:
            taxable = extra_taxable

        has_gst = (cgst + sgst + igst) > 0.005
        has_gstin = bool(gstin)
        is_non_gst = False
        is_ineligible = False
        if not has_gst:
            if not has_gstin:
                is_non_gst = True
            else:
                is_ineligible = True

        return NormalizedTransaction(
            source="BOOKS",
            location=location,
            supplier_gstin=gstin.upper(),
            supplier_name=particulars,
            document_number=doc_no,
            document_type=DocumentType.CREDIT_NOTE,  # buyer's debit note = supplier credit note
            document_date=ref_date or doc_date,
            document_month=doc_month,
            books_month=books_month,
            invoice_value=gross,
            taxable_value=taxable,
            cgst=cgst,
            sgst=sgst,
            igst=igst,
            is_non_gst=is_non_gst,
            is_ineligible_itc=is_ineligible,
            voucher_type=voucher_type,
            voucher_no=voucher_no,
            particulars=particulars,
        )

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    @staticmethod
    def _location_from_filename(path: str, kind: str) -> str:
        base = os.path.basename(path).upper()
        for loc in ("HR", "PL", "VL"):
            if f"-{loc}" in base:
                return loc
        return os.path.basename(path)

    @staticmethod
    def default_files(base_dir: str) -> "BooksParser":
        purchase = sorted(glob.glob(os.path.join(base_dir, "PURCHASE-*.xls")))
        debit = sorted(glob.glob(os.path.join(base_dir, "DEBIT NOTE-*.xls")))
        return BooksParser(purchase_files=purchase, debit_note_files=debit)
