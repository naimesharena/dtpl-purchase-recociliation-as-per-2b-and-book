"""Core reconciliation records.  The service is deliberately independent of Excel."""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from enum import Enum
from typing import Any

ZERO = Decimal(0)


class DocumentType(str, Enum):
    INVOICE = "invoice"
    CREDIT_NOTE = "credit_note"
    DEBIT_NOTE = "debit_note"
    UNKNOWN = "unknown"


class ReconciliationStatus(str, Enum):
    MATCHED = "MATCHED"
    MATCHED_WITH_DIFFERENCE = "MATCHED_WITH_DIFFERENCE"
    TIMING_DIFFERENCE = "TIMING_DIFFERENCE"
    CREDIT_NOTE_TIMING_DIFFERENCE = "CREDIT_NOTE_TIMING_DIFFERENCE"
    BOOKS_ONLY = "BOOKS_ONLY"
    GSTR2B_ONLY = "GSTR2B_ONLY"
    NON_GST = "NON_GST"
    INELIGIBLE_ITC = "INELIGIBLE_ITC"
    DUPLICATE = "DUPLICATE"
    POSSIBLE_MATCH = "POSSIBLE_MATCH"


@dataclass(slots=True)
class Transaction:
    """One source document before reconciliation.

    ``books_month``, ``document_month`` and ``gstr2b_month`` are intentionally
    separate.  Amounts are raw source values; sign normalization happens in
    the service and never mutates the imported row.
    """

    transaction_id: str
    source: str
    supplier_gstin: str = ""
    supplier_name: str = ""
    document_number: str = ""
    document_type: str = DocumentType.INVOICE.value
    document_date: str | None = None
    books_month: str | None = None
    document_month: str | None = None
    gstr2b_month: str | None = None
    supplier_report_month: str | None = None
    taxable_value: Decimal = ZERO
    cgst: Decimal = ZERO
    sgst: Decimal = ZERO
    igst: Decimal = ZERO
    cess: Decimal = ZERO
    total_value: Decimal = ZERO
    classification: str = "auto"
    itc_eligible: bool | None = None
    itc_availability: str | None = None
    is_amendment: bool = False
    amends_document_number: str = ""
    original_invoice_number: str = ""
    source_file: str = ""
    source_sheet: str = ""
    source_row: int | None = None
    raw_document_type: str = ""
    classification_reason: str = ""
    metadata: dict[str, Any] = field(default_factory=dict)

    def __post_init__(self) -> None:
        self.source = str(self.source).strip().lower()
        self.supplier_gstin = str(self.supplier_gstin or "").strip().upper()
        self.document_number = str(self.document_number or "").strip()
        self.document_type = str(getattr(self.document_type, "value", self.document_type)).strip().lower()
        self.amends_document_number = str(self.amends_document_number or "").strip()
        for field_name in ("taxable_value", "cgst", "sgst", "igst", "cess", "total_value"):
            value = getattr(self, field_name)
            if not isinstance(value, Decimal):
                setattr(self, field_name, Decimal(str(value or 0)))

    @property
    def raw_total_itc(self) -> Decimal:
        return self.cgst + self.sgst + self.igst + self.cess

    @property
    def has_gst(self) -> bool:
        return any((self.cgst, self.sgst, self.igst, self.cess))

    @property
    def is_book_entry(self) -> bool:
        return self.source == "books"

    @property
    def is_gstr2b_entry(self) -> bool:
        return self.source == "gstr2b"


@dataclass(slots=True)
class OutstandingItem:
    """A book entry carried into a future 2B period.

    The explicit normalized fields make the carry-forward JSON auditable and
    preserve the fields requested for a persistent open-item register.
    """

    transaction: Transaction
    expected_gstr2b_month: str | None
    normalized_taxable_value: Decimal
    normalized_cgst: Decimal
    normalized_sgst: Decimal
    normalized_igst: Decimal
    normalized_cess: Decimal = ZERO
    status: str = ReconciliationStatus.TIMING_DIFFERENCE.value
    matched_transaction_id: str | None = None

    @property
    def id(self) -> str:
        return self.transaction.transaction_id

    @property
    def supplier_gstin(self) -> str:
        return self.transaction.supplier_gstin

    @property
    def document_number(self) -> str:
        return self.transaction.document_number

    @property
    def document_type(self) -> str:
        return self.transaction.document_type

    @property
    def document_date(self) -> str | None:
        return self.transaction.document_date

    @property
    def books_month(self) -> str | None:
        return self.transaction.books_month

    @property
    def gstr2b_month(self) -> str | None:
        return self.transaction.gstr2b_month

    @property
    def taxable_value(self) -> Decimal:
        return self.transaction.taxable_value

    @property
    def cgst(self) -> Decimal:
        return self.transaction.cgst

    @property
    def sgst(self) -> Decimal:
        return self.transaction.sgst

    @property
    def igst(self) -> Decimal:
        return self.transaction.igst

    @property
    def cess(self) -> Decimal:
        return self.transaction.cess


@dataclass(slots=True)
class ReconciliationRecord:
    record_id: str
    status: str
    books: Transaction | None = None
    gstr2b: Transaction | None = None
    match_level: str | None = None
    reason: str = ""
    differences: dict[str, Decimal] = field(default_factory=dict)
    expected_gstr2b_month: str | None = None
    resolved: bool = False
    possible_match_ids: list[str] = field(default_factory=list)

    @property
    def books_month(self) -> str | None:
        return self.books.books_month if self.books else None

    @property
    def gstr2b_month(self) -> str | None:
        return self.gstr2b.gstr2b_month if self.gstr2b else None


@dataclass(slots=True)
class MonthlySummary:
    month: str
    books_taxable_value: Decimal = ZERO
    books_cgst: Decimal = ZERO
    books_sgst: Decimal = ZERO
    books_igst: Decimal = ZERO
    books_cess: Decimal = ZERO
    books_total_itc: Decimal = ZERO
    books_eligible_itc: Decimal = ZERO
    gstr2b_taxable_value: Decimal = ZERO
    gstr2b_cgst: Decimal = ZERO
    gstr2b_sgst: Decimal = ZERO
    gstr2b_igst: Decimal = ZERO
    gstr2b_cess: Decimal = ZERO
    gstr2b_total_itc: Decimal = ZERO
    gstr2b_eligible_itc: Decimal = ZERO
    matched_amount: Decimal = ZERO
    matched_with_timing_amount: Decimal = ZERO
    timing_difference: Decimal = ZERO
    timing_difference_cleared: Decimal = ZERO
    non_gst_amount: Decimal = ZERO
    ineligible_taxable_value: Decimal = ZERO
    ineligible_itc_amount: Decimal = ZERO
    books_only_amount: Decimal = ZERO
    books_only_itc: Decimal = ZERO
    gstr2b_only_amount: Decimal = ZERO
    gstr2b_only_itc: Decimal = ZERO
    difference_amount: Decimal = ZERO
    difference_taxable_value: Decimal = ZERO


@dataclass(slots=True)
class ReconciliationResult:
    records: list[ReconciliationRecord]
    outstanding_items: list[OutstandingItem]
    monthly_summaries: list[MonthlySummary]
    as_of_month: str | None = None

    @property
    def section_a_books_not_in_2b(self) -> list[ReconciliationRecord]:
        statuses = {
            ReconciliationStatus.BOOKS_ONLY.value,
            ReconciliationStatus.TIMING_DIFFERENCE.value,
            ReconciliationStatus.CREDIT_NOTE_TIMING_DIFFERENCE.value,
            ReconciliationStatus.NON_GST.value,
            ReconciliationStatus.INELIGIBLE_ITC.value,
            ReconciliationStatus.POSSIBLE_MATCH.value,
            ReconciliationStatus.DUPLICATE.value,
        }
        return [record for record in self.records if record.books is not None and record.status in statuses]

    @property
    def section_b_2b_not_in_books(self) -> list[ReconciliationRecord]:
        return [
            record
            for record in self.records
            if record.gstr2b is not None
            and record.books is None
            and record.status in {
                ReconciliationStatus.GSTR2B_ONLY.value,
                ReconciliationStatus.DUPLICATE.value,
                ReconciliationStatus.POSSIBLE_MATCH.value,
            }
        ]
