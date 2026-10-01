"""Tests for credit note sign normalization and matching."""
from datetime import date
from .conftest import make_invoice
from reconciliation import (
    ReconciliationEngine, ReconciliationConfig, ReconciliationStatus, DocumentType,
    SignNormalizer,
)


def _run(books, g2b, cfg=None, window=0):
    cfg = cfg or ReconciliationConfig(future_month_window=window)
    eng = ReconciliationEngine(config=cfg)
    eng.load_from_lists(books, g2b)
    return eng.run()


def test_credit_note_opposite_signs_normalize_and_match():
    """Books Debit Note (positive values) vs 2B Credit Note (positive raw) —
    after normalization both should be negative and match.

    Note: the SignNormalizer flips signs on CREDIT_NOTE transactions for both
    sources because raw books-debit-note and raw 2B-credit-note values are
    positive but reduce ITC. The engine.load_from_lists applies normalization,
    so we inspect engine.books_txs / engine.gstr2b_txs after loading.
    """
    b = make_invoice(
        doc_type=DocumentType.CREDIT_NOTE,
        doc_no="CN-001", taxable=10000, cgst=900, sgst=900, igst=0,
        doc_date=date(2025, 4, 20), books_month="042025",
    )
    g = make_invoice(
        source="GSTR2B", doc_type=DocumentType.CREDIT_NOTE,
        doc_no="CN-001", taxable=10000, cgst=900, sgst=900, igst=0,
        doc_date=date(2025, 4, 20), gstr2b_month="042025",
    )
    cfg = ReconciliationConfig(future_month_window=0)
    eng = ReconciliationEngine(config=cfg)
    eng.load_from_lists([b], [g])
    # After normalization, both should be negative
    assert eng.books_txs[0].normalized_cgst < 0
    assert eng.gstr2b_txs[0].normalized_cgst < 0
    assert eng.books_txs[0].normalized_taxable < 0
    matches, _ = eng.run()
    assert any(m.status == ReconciliationStatus.MATCHED for m in matches)


def test_credit_note_same_month_match():
    b = make_invoice(doc_type=DocumentType.CREDIT_NOTE, doc_no="CN-10",
                     taxable=5000, cgst=450, sgst=450, igst=0,
                     doc_date=date(2025, 5, 31), books_month="052025")
    g = make_invoice(source="GSTR2B", doc_type=DocumentType.CREDIT_NOTE, doc_no="CN-10",
                     taxable=5000, cgst=450, sgst=450, igst=0,
                     doc_date=date(2025, 5, 31), gstr2b_month="052025")
    matches, _ = _run([b], [g])
    assert any(m.status == ReconciliationStatus.MATCHED for m in matches)


def test_previous_month_credit_received_current_month():
    """Credit note is from April (doc date & 2B period), but booked in May Books →
    April 2B is GSTR2B_ONLY (carried forward), May shows late-booked MATCHED."""
    g_april = make_invoice(source="GSTR2B", doc_type=DocumentType.CREDIT_NOTE,
                           doc_no="CN-1", taxable=5000, cgst=450, sgst=450, igst=0,
                           doc_date=date(2025, 4, 30), gstr2b_month="042025")
    b_may = make_invoice(doc_type=DocumentType.CREDIT_NOTE, doc_no="CN-1",
                         taxable=5000, cgst=450, sgst=450, igst=0,
                         doc_date=date(2025, 4, 30), books_month="052025")
    cfg = ReconciliationConfig(future_month_window=3)
    matches, summaries = _run([b_may], [g_april], cfg=cfg, window=3)
    # In April: GSTR2B_ONLY (carrying forward)
    assert any(m.status == ReconciliationStatus.GSTR2B_ONLY and m.reconciliation_month == "042025"
               for m in matches)
    # In May: MATCHED (late-booked)
    assert any(m.status == ReconciliationStatus.MATCHED and m.reconciliation_month == "052025"
               for m in matches)


def test_current_month_credit_appears_next_month():
    """Books May, 2B June for a credit note → TIMING_DIFFERENCE in May, MATCHED in June."""
    b_may = make_invoice(doc_type=DocumentType.CREDIT_NOTE, doc_no="CN-2",
                         taxable=5000, cgst=450, sgst=450, igst=0,
                         doc_date=date(2025, 5, 31), books_month="052025")
    g_june = make_invoice(source="GSTR2B", doc_type=DocumentType.CREDIT_NOTE, doc_no="CN-2",
                          taxable=5000, cgst=450, sgst=450, igst=0,
                          doc_date=date(2025, 5, 31), gstr2b_month="062025")
    cfg = ReconciliationConfig(future_month_window=3)
    matches, _ = _run([b_may], [g_june], cfg=cfg, window=3)
    assert any(m.status == ReconciliationStatus.CREDIT_NOTE_TIMING_DIFFERENCE
               and m.reconciliation_month == "052025"
               and m.expected_gstr2b_month == "062025"
               for m in matches)
    assert any(m.status == ReconciliationStatus.MATCHED
               and m.reconciliation_month == "062025"
               for m in matches)
