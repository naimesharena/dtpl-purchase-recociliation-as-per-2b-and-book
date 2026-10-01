"""Tests for duplicates / amendments / credit note + original invoice relationship."""
from datetime import date
from .conftest import make_invoice
from reconciliation import (
    ReconciliationEngine, ReconciliationConfig, ReconciliationStatus, DocumentType,
)


def _run(books, g2b, window=0):
    eng = ReconciliationEngine(config=ReconciliationConfig(future_month_window=window))
    eng.load_from_lists(books, g2b)
    return eng.run()


def test_duplicate_books_invoice_one_consumes_g2b_other_books_only():
    """If Books has two entries with the same key, only one should match 2B, the other
    remains BOOKS_ONLY."""
    b1 = make_invoice(doc_no="INV-DUP", taxable=10000, cgst=900, sgst=900, igst=0)
    b2 = make_invoice(doc_no="INV-DUP", taxable=10000, cgst=900, sgst=900, igst=0)
    g = make_invoice(source="GSTR2B", doc_no="INV-DUP", taxable=10000, cgst=900, sgst=900, igst=0)
    matches, _ = _run([b1, b2], [g])
    assert sum(1 for m in matches if m.status == ReconciliationStatus.MATCHED) == 1
    assert sum(1 for m in matches if m.status == ReconciliationStatus.BOOKS_ONLY) == 1


def test_duplicate_gstr2b_invoice_leaves_one_g2b_only():
    b = make_invoice(doc_no="INV-D2", taxable=10000, cgst=900, sgst=900)
    g1 = make_invoice(source="GSTR2B", doc_no="INV-D2", taxable=10000, cgst=900, sgst=900)
    g2 = make_invoice(source="GSTR2B", doc_no="INV-D2", taxable=10000, cgst=900, sgst=900)
    matches, _ = _run([b], [g1, g2])
    assert sum(1 for m in matches if m.status == ReconciliationStatus.MATCHED) == 1
    assert sum(1 for m in matches if m.status == ReconciliationStatus.GSTR2B_ONLY) == 1


def test_amendment_in_gstr2b_matches_books_invoice():
    """B2BA (amendment) entries carry same doc number and should still match books."""
    b = make_invoice(doc_no="INV-AMEND", taxable=19613, cgst=1765.17, sgst=1765.17, igst=0,
                     doc_date=date(2025, 6, 1))
    g = make_invoice(source="GSTR2B", doc_no="INV-AMEND", taxable=19613, cgst=1765.17,
                     sgst=1765.17, igst=0, doc_date=date(2025, 6, 1),
                     gstr2b_month="072025", supplier_name="(Amended)")
    cfg = ReconciliationConfig(future_month_window=3)
    matches, _ = _run([b], [g], window=3)
    assert any(m.status in (ReconciliationStatus.MATCHED, ReconciliationStatus.MATCHED_WITH_DIFFERENCE,
                            ReconciliationStatus.TIMING_DIFFERENCE)
               for m in matches)


def test_credit_note_and_original_invoice_distinct():
    """Original invoice and credit note should be matched independently."""
    b_inv = make_invoice(doc_no="INV-1", taxable=100000, cgst=9000, sgst=9000,
                         doc_date=date(2025, 4, 10), books_month="042025")
    b_cn = make_invoice(doc_type=DocumentType.CREDIT_NOTE, doc_no="CN-1", taxable=10000,
                        cgst=900, sgst=900, doc_date=date(2025, 4, 30), books_month="042025")
    g_inv = make_invoice(source="GSTR2B", doc_no="INV-1", taxable=100000, cgst=9000, sgst=9000,
                         doc_date=date(2025, 4, 10), gstr2b_month="042025")
    g_cn = make_invoice(source="GSTR2B", doc_type=DocumentType.CREDIT_NOTE, doc_no="CN-1",
                        taxable=10000, cgst=900, sgst=900, doc_date=date(2025, 4, 30), gstr2b_month="042025")
    matches, _ = _run([b_inv, b_cn], [g_inv, g_cn])
    matched = [m for m in matches if m.status == ReconciliationStatus.MATCHED]
    assert len(matched) == 2
    types_matched = {(m.books_tx.document_type if m.books_tx else m.gstr2b_tx.document_type).value
                     for m in matched}
    assert "INVOICE" in types_matched
    assert "CREDIT_NOTE" in types_matched


def test_non_gst_not_counted_as_mismatch():
    """Ensure non-GST purchase isn't counted as missing from 2B."""
    b_nongst = make_invoice(doc_no="NG-1", gstin="", taxable=15000, cgst=0, sgst=0, igst=0,
                            is_non_gst=True)
    b_reg = make_invoice(doc_no="INV-R", taxable=50000, cgst=4500, sgst=4500, igst=0)
    g_reg = make_invoice(source="GSTR2B", doc_no="INV-R", taxable=50000, cgst=4500, sgst=4500, igst=0)
    matches, summaries = _run([b_nongst, b_reg], [g_reg])
    apr = summaries["042025"]
    assert apr.non_gst_count == 1
    assert apr.matched_count == 1
    assert apr.books_only_count == 0
    assert apr.gstr2b_only_count == 0
