"""Shared pytest fixtures for reconciliation tests."""
import sys
import os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from datetime import date
import pytest

from reconciliation import (
    NormalizedTransaction,
    DocumentType,
    ReconciliationConfig,
    ReconciliationEngine,
)


def make_invoice(
    *,
    source: str = "BOOKS",
    gstin: str = "24AABCE1234F1Z1",
    doc_no: str = "INV-001",
    doc_date: date = date(2025, 4, 15),
    books_month: str = "042025",
    gstr2b_month: str = "042025",
    taxable: float = 100000.0,
    cgst: float = 9000.0,
    sgst: float = 9000.0,
    igst: float = 0.0,
    inv_val: float | None = None,
    doc_type: DocumentType = DocumentType.INVOICE,
    location: str = "TEST",
    supplier_name: str = "Test Supplier",
    is_non_gst: bool = False,
    is_ineligible: bool = False,
    itc_avail: str = "Yes",
) -> NormalizedTransaction:
    """Helper to build test transactions quickly."""
    if inv_val is None:
        inv_val = taxable + cgst + sgst + igst
    return NormalizedTransaction(
        source=source, location=location,
        supplier_gstin=gstin, supplier_name=supplier_name,
        document_number=doc_no, document_type=doc_type,
        document_date=doc_date, document_month=books_month if source == "BOOKS" else gstr2b_month,
        books_month=books_month if source == "BOOKS" else "",
        gstr2b_month=gstr2b_month if source == "GSTR2B" else "",
        invoice_value=inv_val, taxable_value=taxable,
        cgst=cgst, sgst=sgst, igst=igst,
        itc_available=itc_avail,
        is_non_gst=is_non_gst, is_ineligible_itc=is_ineligible,
    )


@pytest.fixture
def default_config():
    return ReconciliationConfig(
        tolerance_taxable=1.0, tolerance_cgst=1.0,
        tolerance_sgst=1.0, tolerance_igst=1.0,
        tolerance_total_itc=5.0, future_month_window=3,
    )


@pytest.fixture
def engine(default_config):
    return ReconciliationEngine(config=default_config)
