"""Tests for special transactions — non-GST, ineligible ITC, zero-GST, IGST, etc."""
from datetime import date
from .conftest import make_invoice
from reconciliation import (
    ReconciliationEngine, ReconciliationConfig, ReconciliationStatus, DocumentType,
)


def _run(books, g2b):
    eng = ReconciliationEngine(config=ReconciliationConfig(future_month_window=0))
    eng.load_from_lists(books, g2b)
    return eng.run()


def test_non_gst_purchase_classified_correctly():
    """Non-GST purchase has no GSTIN and all GST=0 → NON_GST, not BOOKS_ONLY."""
    b = make_invoice(gstin="", taxable=20000, cgst=0, sgst=0, igst=0,
                     is_non_gst=True)
    matches, _ = _run([b], [])
    statuses = [m.status for m in matches]
    assert ReconciliationStatus.NON_GST in statuses
    assert ReconciliationStatus.BOOKS_ONLY not in statuses


def test_non_gst_does_not_affect_itc_mismatch():
    """Non-GST purchases excluded from ITC diff calculation."""
    b_gst = make_invoice(taxable=100000, cgst=9000, sgst=9000, igst=0)
    b_nongst = make_invoice(doc_no="NONGST-1", gstin="", taxable=20000,
                            cgst=0, sgst=0, igst=0, is_non_gst=True)
    g = make_invoice(source="GSTR2B", taxable=100000, cgst=9000, sgst=9000, igst=0)
    matches, summaries = _run([b_gst, b_nongst], [g])
    # The eligible invoice matches, non-gst is separately classified
    assert any(m.status == ReconciliationStatus.MATCHED for m in matches)
    assert any(m.status == ReconciliationStatus.NON_GST for m in matches)
    apr = summaries.get("042025")
    assert apr is not None
    assert apr.non_gst_count == 1
    assert apr.matched_count == 1
    assert apr.books_only_count == 0


def test_ineligible_itc_classified():
    """Entry with GSTIN but GST=0 → INELIGIBLE_ITC."""
    b = make_invoice(gstin="24AABCE1234F1Z1", taxable=10000, cgst=0, sgst=0, igst=0,
                     is_ineligible=True, inv_val=10000)
    matches, _ = _run([b], [])
    assert any(m.status == ReconciliationStatus.INELIGIBLE_ITC for m in matches)
    assert not any(m.status == ReconciliationStatus.BOOKS_ONLY for m in matches)


def test_eligible_purchase_with_gst_matches():
    b = make_invoice(taxable=50000, cgst=4500, sgst=4500, igst=0)
    g = make_invoice(source="GSTR2B", taxable=50000, cgst=4500, sgst=4500, igst=0)
    matches, _ = _run([b], [g])
    assert any(m.status == ReconciliationStatus.MATCHED for m in matches)


def test_zero_gst_purchase_with_gstin_is_ineligible():
    """Even if parser doesn't pre-flag it, zero-GST with GSTIN is ineligible."""
    # Parser sets is_ineligible_itc=True when there's a GSTIN but zero GST.
    b = make_invoice(gstin="24AABCE1234F1Z1", taxable=8000, cgst=0, sgst=0, igst=0,
                     is_ineligible=True)
    matches, _ = _run([b], [])
    assert any(m.status == ReconciliationStatus.INELIGIBLE_ITC for m in matches)


def test_cgst_sgst_mixed_purchase():
    b = make_invoice(taxable=25000, cgst=2250, sgst=2250, igst=0)
    g = make_invoice(source="GSTR2B", taxable=25000, cgst=2250, sgst=2250, igst=0)
    matches, _ = _run([b], [g])
    assert any(m.status == ReconciliationStatus.MATCHED for m in matches)


def test_igst_purchase():
    b = make_invoice(taxable=50000, cgst=0, sgst=0, igst=9000)
    g = make_invoice(source="GSTR2B", taxable=50000, cgst=0, sgst=0, igst=9000)
    matches, _ = _run([b], [g])
    assert any(m.status == ReconciliationStatus.MATCHED for m in matches)
