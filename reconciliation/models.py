"""Data models for GSTR-2B vs Books reconciliation."""
from __future__ import annotations

import uuid
from dataclasses import dataclass, field, asdict
from enum import Enum
from typing import Optional, List, Dict, Any
from datetime import date, datetime


class DocumentType(str, Enum):
    """Normalized document type.

    INVOICE = regular purchase invoice (adds ITC).
    CREDIT_NOTE = supplier credit note / buyer's debit note (reduces ITC).
    DEBIT_NOTE = supplier debit note (adds ITC; rare in 2B).
    """
    INVOICE = "INVOICE"
    CREDIT_NOTE = "CREDIT_NOTE"
    DEBIT_NOTE = "DEBIT_NOTE"


class ReconciliationStatus(str, Enum):
    """Status of a reconciled item (one primary status per item)."""
    MATCHED = "MATCHED"                                    # exact match (all values within tolerance)
    MATCHED_WITH_DIFFERENCE = "MATCHED_WITH_DIFFERENCE"    # doc found but GST/taxable differ
    TIMING_DIFFERENCE = "TIMING_DIFFERENCE"                # present in books, expected in future 2B period
    BOOKS_ONLY = "BOOKS_ONLY"                              # in books, not in 2B and not expected (mismatch)
    GSTR2B_ONLY = "GSTR2B_ONLY"                            # in 2B, not found in books
    NON_GST = "NON_GST"                                    # non-GST purchase (not expected in 2B)
    INELIGIBLE_ITC = "INELIGIBLE_ITC"                      # ineligible ITC (GST not separately claimed)
    CREDIT_NOTE_TIMING_DIFFERENCE = "CREDIT_NOTE_TIMING_DIFFERENCE"
    DUPLICATE = "DUPLICATE"
    POSSIBLE_MATCH = "POSSIBLE_MATCH"                      # fuzzy match — needs human review


@dataclass
class NormalizedTransaction:
    """A transaction from either Books or GSTR-2B, normalized to a common schema.

    All monetary fields use the internal sign convention after normalization:
      * Regular purchase invoice → taxable, cgst, sgst, igst are POSITIVE (adds ITC).
      * Credit note              → taxable, cgst, sgst, igst are NEGATIVE (reduces ITC).
      * Debit note (supplier)    → taxable, cgst, sgst, igst are POSITIVE (adds ITC).
    """
    id: str = field(default_factory=lambda: uuid.uuid4().hex[:12])
    source: str = ""                                   # "BOOKS" or "GSTR2B"
    location: str = ""                                 # HR / PL / VL / etc.
    supplier_gstin: str = ""
    supplier_name: str = ""
    document_number: str = ""
    document_type: DocumentType = DocumentType.INVOICE
    document_date: Optional[date] = None
    document_month: str = ""                           # MMYYYY derived from document_date
    books_month: str = ""                              # MMYYYY of accounting month (books only)
    gstr2b_month: str = ""                             # MMYYYY of GSTR-1/2B filing period (2B only)
    invoice_value: float = 0.0
    taxable_value: float = 0.0                         # raw value as reported
    cgst: float = 0.0
    sgst: float = 0.0
    igst: float = 0.0
    cess: float = 0.0
    itc_available: str = ""                            # "Yes" / "No" from 2B
    place_of_supply: str = ""
    reverse_charge: str = ""
    is_non_gst: bool = False
    is_ineligible_itc: bool = False
    voucher_type: str = ""                             # original books voucher type
    voucher_no: str = ""
    particulars: str = ""
    raw_row: Dict[str, Any] = field(default_factory=dict)

    # --- normalized values (sign-corrected) ---
    @property
    def normalized_taxable(self) -> float:
        return self.taxable_value if self._sign_factor() > 0 else -self.taxable_value

    @property
    def normalized_cgst(self) -> float:
        return self.cgst if self._sign_factor() > 0 else -self.cgst

    @property
    def normalized_sgst(self) -> float:
        return self.sgst if self._sign_factor() > 0 else -self.sgst

    @property
    def normalized_igst(self) -> float:
        return self.igst if self._sign_factor() > 0 else -self.igst

    @property
    def normalized_total_itc(self) -> float:
        return self.normalized_cgst + self.normalized_sgst + self.normalized_igst

    def _sign_factor(self) -> int:
        """Return +1 if values are already in our convention, -1 if they need flipping.

        The raw data in this repository is always positive; the sign is determined by
        document type and source (see ``normalizer.DefaultSignConvention``).
        """
        # Sign decision is handled externally via SignNormalizer; this property is a fallback.
        # The normalizer sets the raw values so that normalized_* match the convention, so
        # we simply return +1 here (values written are already normalized).
        return 1

    def key(self) -> str:
        """Stable (gstin, doc_type, doc_no) key used for exact matching."""
        return f"{(self.supplier_gstin or '').strip().upper()}|{self.document_type.value}|{self._doc_no_norm()}"

    def _doc_no_norm(self) -> str:
        return (self.document_number or "").strip().upper().replace(" ", "").replace("/", "").replace("-", "")

    def to_row(self) -> Dict[str, Any]:
        return {
            "id": self.id,
            "source": self.source,
            "location": self.location,
            "supplier_gstin": self.supplier_gstin,
            "supplier_name": self.supplier_name,
            "document_number": self.document_number,
            "document_type": self.document_type.value,
            "document_date": self.document_date.isoformat() if self.document_date else "",
            "document_month": self.document_month,
            "books_month": self.books_month,
            "gstr2b_month": self.gstr2b_month,
            "invoice_value": round(self.invoice_value, 2),
            "taxable_value": round(self.taxable_value, 2),
            "cgst": round(self.cgst, 2),
            "sgst": round(self.sgst, 2),
            "igst": round(self.igst, 2),
            "cess": round(self.cess, 2),
            "normalized_taxable": round(self.normalized_taxable, 2),
            "normalized_cgst": round(self.normalized_cgst, 2),
            "normalized_sgst": round(self.normalized_sgst, 2),
            "normalized_igst": round(self.normalized_igst, 2),
            "normalized_total_itc": round(self.normalized_total_itc, 2),
            "itc_available": self.itc_available,
            "place_of_supply": self.place_of_supply,
            "is_non_gst": self.is_non_gst,
            "is_ineligible_itc": self.is_ineligible_itc,
            "voucher_type": self.voucher_type,
            "voucher_no": self.voucher_no,
            "particulars": self.particulars,
        }


@dataclass
class ReconciliationMatch:
    """Reconciliation result for a single books-side transaction (or 2B-only item)."""
    status: ReconciliationStatus
    books_tx: Optional[NormalizedTransaction] = None
    gstr2b_tx: Optional[NormalizedTransaction] = None
    reconciliation_month: str = ""           # MMYYYY for which reconciliation was run
    expected_gstr2b_month: str = ""          # for TIMING_DIFFERENCE
    difference_taxable: float = 0.0
    difference_cgst: float = 0.0
    difference_sgst: float = 0.0
    difference_igst: float = 0.0
    notes: str = ""

    def to_row(self) -> Dict[str, Any]:
        b = self.books_tx
        g = self.gstr2b_tx
        return {
            "status": self.status.value,
            "reconciliation_month": self.reconciliation_month,
            "expected_gstr2b_month": self.expected_gstr2b_month,
            "books_id": b.id if b else "",
            "books_location": b.location if b else "",
            "books_supplier_gstin": b.supplier_gstin if b else "",
            "books_supplier_name": b.supplier_name if b else (g.supplier_name if g else ""),
            "books_document_number": b.document_number if b else "",
            "books_document_type": b.document_type.value if b else "",
            "books_document_date": b.document_date.isoformat() if b and b.document_date else "",
            "books_voucher_no": b.voucher_no if b else "",
            "books_voucher_type": b.voucher_type if b else "",
            "books_invoice_value": round(b.invoice_value, 2) if b else 0.0,
            "books_taxable": round(b.taxable_value, 2) if b else 0.0,
            "books_cgst": round(b.cgst, 2) if b else 0.0,
            "books_sgst": round(b.sgst, 2) if b else 0.0,
            "books_igst": round(b.igst, 2) if b else 0.0,
            "gstr2b_id": g.id if g else "",
            "gstr2b_supplier_gstin": g.supplier_gstin if g else "",
            "gstr2b_document_number": g.document_number if g else "",
            "gstr2b_document_type": g.document_type.value if g else "",
            "gstr2b_document_date": g.document_date.isoformat() if g and g.document_date else "",
            "gstr2b_period": g.gstr2b_month if g else "",
            "gstr2b_taxable": round(g.taxable_value, 2) if g else 0.0,
            "gstr2b_cgst": round(g.cgst, 2) if g else 0.0,
            "gstr2b_sgst": round(g.sgst, 2) if g else 0.0,
            "gstr2b_igst": round(g.igst, 2) if g else 0.0,
            "gstr2b_itc_available": g.itc_available if g else "",
            "diff_taxable": round(self.difference_taxable, 2),
            "diff_cgst": round(self.difference_cgst, 2),
            "diff_sgst": round(self.difference_sgst, 2),
            "diff_igst": round(self.difference_igst, 2),
            "notes": self.notes,
        }


@dataclass
class MonthlySummary:
    """Aggregated summary for one reconciliation month."""
    month: str = ""
    # Books side
    books_taxable: float = 0.0
    books_cgst: float = 0.0
    books_sgst: float = 0.0
    books_igst: float = 0.0
    books_total_itc: float = 0.0
    books_count: int = 0
    # GSTR-2B side
    gstr2b_taxable: float = 0.0
    gstr2b_cgst: float = 0.0
    gstr2b_sgst: float = 0.0
    gstr2b_igst: float = 0.0
    gstr2b_total_itc: float = 0.0
    gstr2b_count: int = 0
    # Reconciled categories
    matched_count: int = 0
    matched_taxable: float = 0.0
    matched_itc: float = 0.0
    matched_with_diff_count: int = 0
    matched_with_diff_taxable: float = 0.0
    matched_with_diff_itc_diff: float = 0.0
    timing_diff_count: int = 0
    timing_diff_taxable: float = 0.0
    timing_diff_itc: float = 0.0
    non_gst_count: int = 0
    non_gst_amount: float = 0.0
    ineligible_itc_count: int = 0
    ineligible_itc_amount: float = 0.0
    books_only_count: int = 0
    books_only_taxable: float = 0.0
    books_only_itc: float = 0.0
    gstr2b_only_count: int = 0
    gstr2b_only_taxable: float = 0.0
    gstr2b_only_itc: float = 0.0
    duplicate_count: int = 0
    possible_match_count: int = 0

    def to_row(self) -> Dict[str, Any]:
        def r(x):
            return round(x, 2) if isinstance(x, float) else x
        return {k: r(v) for k, v in asdict(self).items()}


@dataclass
class ReconciliationConfig:
    """Tunable parameters controlling reconciliation behaviour."""
    # Monetary tolerance for matching (absolute, in rupees)
    tolerance_taxable: float = 1.00
    tolerance_cgst: float = 1.00
    tolerance_sgst: float = 1.00
    tolerance_igst: float = 1.00
    tolerance_total_itc: float = 5.00
    # How many future GSTR-2B months to search for timing-difference resolution
    future_month_window: int = 3
    # Treat entries with zero GST and no GSTIN as non-GST
    classify_no_gstin_non_gst: bool = True
    # Treat entries with GSTIN but zero GST (and taxable == gross) as ineligible ITC
    classify_zero_gst_as_ineligible_when_gstin_present: bool = True
    # Round-off differences are treated as matched
    round_off_tolerance: float = 5.00
