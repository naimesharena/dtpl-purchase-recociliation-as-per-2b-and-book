"""Tests for monthly summary calculations."""
from datetime import date
from .conftest import make_invoice
from reconciliation import ReconciliationEngine, ReconciliationConfig, ReconciliationStatus


def test_summary_totals_for_full_example():
    """Test scenario from the spec's 'Example' section.

    September Books:
      Eligible Purchase: taxable=100000, CGST=9000, SGST=9000, IGST=0
      Non-GST Purchase:  taxable=20000,  CGST=0, SGST=0, IGST=0
      Ineligible ITC:    taxable=10000,  CGST=0, SGST=0, IGST=0 (gross=taxable)
      Credit Note:       taxable=-10000, CGST=-900, SGST=-900, IGST=0
    September GSTR-2B: only the eligible invoice
    October GSTR-2B:   the credit note
    """
    from reconciliation import DocumentType
    # Books for September
    b1 = make_invoice(doc_no="INV-ELIG", taxable=100000, cgst=9000, sgst=9000, igst=0,
                     doc_date=date(2025, 9, 5), books_month="092025")
    b2 = make_invoice(doc_no="NONGST-1", gstin="", taxable=20000, cgst=0, sgst=0, igst=0,
                     doc_date=date(2025, 9, 7), books_month="092025", is_non_gst=True)
    b3 = make_invoice(doc_no="INELIG-1", gstin="24OTHER9999K1Z1", taxable=10000, cgst=0, sgst=0,
                     igst=0, inv_val=10000, doc_date=date(2025, 9, 9), books_month="092025",
                     is_ineligible=True)
    b4 = make_invoice(doc_type=DocumentType.CREDIT_NOTE, doc_no="CN-1", taxable=10000, cgst=900,
                     sgst=900, igst=0, doc_date=date(2025, 9, 30), books_month="092025")

    # GSTR-2B September
    g_sep = make_invoice(source="GSTR2B", doc_no="INV-ELIG", taxable=100000, cgst=9000, sgst=9000,
                        igst=0, doc_date=date(2025, 9, 5), gstr2b_month="092025")
    # GSTR-2B October (credit note)
    g_oct_cn = make_invoice(source="GSTR2B", doc_type=DocumentType.CREDIT_NOTE, doc_no="CN-1",
                           taxable=10000, cgst=900, sgst=900, igst=0,
                           doc_date=date(2025, 9, 30), gstr2b_month="102025")

    eng = ReconciliationEngine(config=ReconciliationConfig(future_month_window=2))
    eng.load_from_lists([b1, b2, b3, b4], [g_sep, g_oct_cn])
    matches, summaries = eng.run()

    sep = summaries["092025"]
    # Books totals
    assert sep.books_count == 4
    assert abs(sep.books_taxable - (100000 + 20000 + 10000 - 10000)) < 1
    # ITC totals should include eligible invoice and credit note (normalized to negative)
    # but NOT non-gst/ineligible (they have zero GST)
    assert abs(sep.books_total_itc - (9000 + 9000 - 900 - 900)) < 1

    # Classification
    assert sep.matched_count == 1           # eligible invoice
    assert sep.non_gst_count == 1
    assert sep.ineligible_itc_count == 1
    assert sep.timing_diff_count == 1       # credit note expected in October
    assert sep.books_only_count == 0

    # October summary: credit note should now be matched as carry-forward
    oct_ = summaries["102025"]
    assert oct_.matched_count == 1  # CN-1 resolved

    # Verify statuses for September
    sep_matches = [m for m in matches if m.reconciliation_month == "092025"]
    statuses = {m.status for m in sep_matches}
    assert ReconciliationStatus.MATCHED in statuses
    assert ReconciliationStatus.NON_GST in statuses
    assert ReconciliationStatus.INELIGIBLE_ITC in statuses
    assert ReconciliationStatus.CREDIT_NOTE_TIMING_DIFFERENCE in statuses


def test_mutually_exclusive_status_per_transaction():
    """Each transaction must be in exactly one primary status per month."""
    from reconciliation import DocumentType
    b1 = make_invoice(doc_no="A", taxable=10000, cgst=900, sgst=900)
    b2 = make_invoice(doc_no="B", taxable=20000, cgst=1800, sgst=1800, books_month="042025")
    g1 = make_invoice(source="GSTR2B", doc_no="A", taxable=10000, cgst=900, sgst=900)
    matches, _ = ReconciliationEngine(config=ReconciliationConfig(future_month_window=0)
                                      ).load_from_lists([b1, b2], [g1]) or (None, None)
    # Re-run clean
    eng = ReconciliationEngine(config=ReconciliationConfig(future_month_window=0))
    eng.load_from_lists([b1, b2], [g1])
    matches, _ = eng.run()
    # b1 (doc A) should appear once
    b1_entries = [m for m in matches if m.books_tx and m.books_tx.document_number == "A"]
    assert len(b1_entries) == 1
    # b2 appears once (BOOKS_ONLY)
    b2_entries = [m for m in matches if m.books_tx and m.books_tx.document_number == "B"]
    assert len(b2_entries) == 1
    assert b2_entries[0].status == ReconciliationStatus.BOOKS_ONLY
