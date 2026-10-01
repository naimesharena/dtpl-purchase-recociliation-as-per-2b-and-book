"""Tests for month-shifted reconciliation and carry-forward."""
from datetime import date
from .conftest import make_invoice
from reconciliation import (
    ReconciliationEngine, ReconciliationConfig, ReconciliationStatus,
)


def _engine(window=3):
    return ReconciliationEngine(config=ReconciliationConfig(future_month_window=window))


def test_books_september_gstr2b_october_timing_then_matched():
    """Books September, GSTR-2B October → TIMING_DIFFERENCE in Sep, MATCHED in Oct."""
    b_sep = make_invoice(doc_no="INV-S1", taxable=100000, cgst=9000, sgst=9000,
                         doc_date=date(2025, 9, 10), books_month="092025")
    g_oct = make_invoice(source="GSTR2B", doc_no="INV-S1", taxable=100000, cgst=9000, sgst=9000,
                         doc_date=date(2025, 9, 10), gstr2b_month="102025")
    eng = _engine(window=3)
    eng.load_from_lists([b_sep], [g_oct])
    matches, _ = eng.run()
    assert any(m.status == ReconciliationStatus.TIMING_DIFFERENCE
               and m.reconciliation_month == "092025"
               and m.expected_gstr2b_month == "102025"
               for m in matches)
    assert any(m.status == ReconciliationStatus.MATCHED
               and m.reconciliation_month == "102025"
               for m in matches)


def test_books_august_gstr2b_september_previous_credit():
    """Books August, GSTR-2B September → forward timing difference (like Scenario A)."""
    b_aug = make_invoice(doc_no="INV-A1", taxable=100000, cgst=9000, sgst=9000,
                         doc_date=date(2025, 8, 15), books_month="082025")
    g_sep = make_invoice(source="GSTR2B", doc_no="INV-A1", taxable=100000, cgst=9000, sgst=9000,
                         doc_date=date(2025, 8, 15), gstr2b_month="092025")
    eng = _engine(window=3)
    eng.load_from_lists([b_aug], [g_sep])
    matches, _ = eng.run()
    assert any(m.status == ReconciliationStatus.TIMING_DIFFERENCE
               and m.reconciliation_month == "082025"
               and m.expected_gstr2b_month == "092025"
               for m in matches)
    assert any(m.status == ReconciliationStatus.MATCHED
               and m.reconciliation_month == "092025"
               for m in matches)


def test_item_unmatched_after_window_becomes_books_only():
    """Item with no matching 2B even after future_month_window → BOOKS_ONLY."""
    b_apr = make_invoice(doc_no="INV-X", taxable=50000, cgst=4500, sgst=4500,
                         doc_date=date(2025, 4, 1), books_month="042025")
    # Only provide 2B data for April (no future months)
    g_apr = make_invoice(source="GSTR2B", doc_no="INV-OTHER", taxable=1000, cgst=90, sgst=90,
                         doc_date=date(2025, 4, 1), gstr2b_month="042025")
    eng = ReconciliationEngine(config=ReconciliationConfig(future_month_window=0))
    eng.load_from_lists([b_apr], [g_apr])
    matches, _ = eng.run()
    assert any(m.status == ReconciliationStatus.BOOKS_ONLY
               and m.books_tx.document_number == "INV-X"
               for m in matches)


def test_timing_diff_matched_after_multiple_months():
    """Books April, 2B July (3 months later) → carried forward then matched."""
    b_apr = make_invoice(doc_no="INV-LATE", taxable=10000, cgst=900, sgst=900,
                         doc_date=date(2025, 4, 10), books_month="042025")
    g_jul = make_invoice(source="GSTR2B", doc_no="INV-LATE", taxable=10000, cgst=900, sgst=900,
                         doc_date=date(2025, 4, 10), gstr2b_month="072025")
    # add may/june empty pools so engine has those months
    eng = _engine(window=4)
    eng.load_from_lists([b_apr], [g_jul])
    matches, _ = eng.run()
    # In April: TIMING_DIFFERENCE
    assert any(m.status == ReconciliationStatus.TIMING_DIFFERENCE
               and m.reconciliation_month == "042025"
               and m.expected_gstr2b_month == "072025"
               for m in matches)
    # In July: MATCHED
    assert any(m.status == ReconciliationStatus.MATCHED
               and m.reconciliation_month == "072025"
               for m in matches)


def test_late_booked_gstr2b_from_previous_month_matches_in_current():
    """2B April, Books May → reverse timing (2B arrives before booking)."""
    g_apr = make_invoice(source="GSTR2B", doc_no="INV-R", taxable=10000, cgst=900, sgst=900,
                         doc_date=date(2025, 4, 20), gstr2b_month="042025")
    b_may = make_invoice(doc_no="INV-R", taxable=10000, cgst=900, sgst=900,
                         doc_date=date(2025, 4, 20), books_month="052025")
    eng = _engine(window=3)
    eng.load_from_lists([b_may], [g_apr])
    matches, _ = eng.run()
    assert any(m.status == ReconciliationStatus.GSTR2B_ONLY
               and m.reconciliation_month == "042025"
               for m in matches)
    assert any(m.status == ReconciliationStatus.MATCHED
               and m.reconciliation_month == "052025"
               and "Late-booked" in m.notes
               for m in matches)
