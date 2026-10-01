"""Tests for basic matching cases — exact matches, mismatches, GST/value errors."""
from datetime import date
from .conftest import make_invoice
from reconciliation import (
    ReconciliationEngine, ReconciliationConfig, ReconciliationStatus, DocumentType,
)


def _run(books, g2b, cfg=None):
    eng = ReconciliationEngine(config=cfg or ReconciliationConfig(future_month_window=0))
    eng.load_from_lists(books, g2b)
    matches, summaries = eng.run()
    return matches, summaries


def test_exact_invoice_match():
    """Exact match on GSTIN + doc number + all values → MATCHED."""
    b = make_invoice(taxable=100000, cgst=9000, sgst=9000, igst=0)
    g = make_invoice(source="GSTR2B", taxable=100000, cgst=9000, sgst=9000, igst=0)
    matches, _ = _run([b], [g])
    statuses = [m.status for m in matches]
    assert ReconciliationStatus.MATCHED in statuses
    # No books-only / g2b-only entries
    assert ReconciliationStatus.BOOKS_ONLY not in statuses
    assert ReconciliationStatus.GSTR2B_ONLY not in statuses


def test_taxable_value_mismatch():
    """Same doc but taxable differs → MATCHED_WITH_DIFFERENCE when within wider tolerance."""
    b = make_invoice(taxable=100000, cgst=9000, sgst=9000, igst=0)
    g = make_invoice(source="GSTR2B", taxable=100050, cgst=9000, sgst=9000, igst=0)
    cfg = ReconciliationConfig(tolerance_taxable=1, future_month_window=0)
    matches, _ = _run([b], [g], cfg=cfg)
    statuses = [m.status for m in matches]
    assert ReconciliationStatus.MATCHED_WITH_DIFFERENCE in statuses


def test_cgst_mismatch():
    b = make_invoice(taxable=100000, cgst=9000, sgst=9000, igst=0)
    g = make_invoice(source="GSTR2B", taxable=100000, cgst=9001.5, sgst=9000, igst=0)
    cfg = ReconciliationConfig(tolerance_cgst=1, future_month_window=0)
    matches, _ = _run([b], [g], cfg=cfg)
    assert any(m.status == ReconciliationStatus.MATCHED_WITH_DIFFERENCE for m in matches)


def test_sgst_mismatch():
    b = make_invoice(taxable=100000, cgst=9000, sgst=9000, igst=0)
    g = make_invoice(source="GSTR2B", taxable=100000, cgst=9000, sgst=9002, igst=0)
    cfg = ReconciliationConfig(tolerance_sgst=1, future_month_window=0)
    matches, _ = _run([b], [g], cfg=cfg)
    assert any(m.status == ReconciliationStatus.MATCHED_WITH_DIFFERENCE for m in matches)


def test_igst_mismatch():
    b = make_invoice(taxable=100000, cgst=0, sgst=0, igst=18000)
    g = make_invoice(source="GSTR2B", taxable=100000, cgst=0, sgst=0, igst=18005)
    cfg = ReconciliationConfig(tolerance_igst=1, future_month_window=0)
    matches, _ = _run([b], [g], cfg=cfg)
    assert any(m.status == ReconciliationStatus.MATCHED_WITH_DIFFERENCE for m in matches)


def test_gstin_mismatch_results_in_books_only():
    """Different GSTIN → no match, books only + gstr2b only."""
    b = make_invoice(gstin="24AABCE1234F1Z1")
    g = make_invoice(source="GSTR2B", gstin="24AABCE9999F1Z9")
    matches, _ = _run([b], [g])
    statuses = [m.status for m in matches]
    assert ReconciliationStatus.BOOKS_ONLY in statuses
    assert ReconciliationStatus.GSTR2B_ONLY in statuses


def test_invoice_number_mismatch_results_in_unmatched():
    b = make_invoice(doc_no="INV-001")
    g = make_invoice(source="GSTR2B", doc_no="INV-999")
    matches, _ = _run([b], [g])
    statuses = [m.status for m in matches]
    assert ReconciliationStatus.BOOKS_ONLY in statuses
    assert ReconciliationStatus.GSTR2B_ONLY in statuses


def test_within_tolerance_counts_as_matched():
    b = make_invoice(taxable=100000, cgst=9000, sgst=9000, igst=0)
    g = make_invoice(source="GSTR2B", taxable=100000.5, cgst=9000.3, sgst=9000.3, igst=0)
    cfg = ReconciliationConfig(tolerance_taxable=1, tolerance_cgst=1,
                               tolerance_sgst=1, tolerance_igst=1, future_month_window=0)
    matches, _ = _run([b], [g], cfg=cfg)
    assert any(m.status == ReconciliationStatus.MATCHED for m in matches)
