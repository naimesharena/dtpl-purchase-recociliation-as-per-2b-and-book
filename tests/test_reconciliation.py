from __future__ import annotations

import unittest
from decimal import Decimal

from gstr2b_recon import (
    ReconciliationConfig,
    ReconciliationService,
    ReconciliationStatus,
    Transaction,
)

D = Decimal


def transaction(
    transaction_id: str,
    source: str,
    *,
    gstin: str = "24ABCDE1234F1Z5",
    number: str = "INV-1",
    doc_type: str = "invoice",
    books_month: str | None = "2025-09",
    gstr_month: str | None = None,
    document_month: str | None = "2025-09",
    taxable: str = "1000",
    cgst: str = "90",
    sgst: str = "90",
    igst: str = "0",
    cess: str = "0",
    total: str = "1180",
    classification: str = "eligible",
    itc_eligible: bool | None = True,
    **kwargs,
) -> Transaction:
    return Transaction(
        transaction_id=transaction_id,
        source=source,
        supplier_gstin=gstin,
        document_number=number,
        document_type=doc_type,
        document_date=(document_month + "-15") if document_month else None,
        books_month=books_month if source == "books" else None,
        document_month=document_month,
        gstr2b_month=gstr_month if source == "gstr2b" else None,
        taxable_value=D(taxable),
        cgst=D(cgst),
        sgst=D(sgst),
        igst=D(igst),
        cess=D(cess),
        total_value=D(total),
        classification=classification,
        itc_eligible=itc_eligible,
        **kwargs,
    )


class ReconciliationServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.config = ReconciliationConfig()
        self.service = ReconciliationService(self.config)

    def reconcile_pair(self, books: Transaction, gstr: Transaction):
        return self.service.reconcile([books], [gstr], as_of_month="2025-09").records[0]

    def test_exact_invoice_match(self) -> None:
        record = self.reconcile_pair(
            transaction("b1", "books"),
            transaction("g1", "gstr2b", gstr_month="2025-09"),
        )
        self.assertEqual(record.status, ReconciliationStatus.MATCHED.value)
        self.assertEqual(record.match_level, "GSTIN_DOCUMENT_NUMBER")

    def test_gstr2b_ineligible_flag_prevents_a_false_itc_match(self) -> None:
        book = transaction("b1", "books")
        gstr = transaction("g1", "gstr2b", gstr_month="2025-09", itc_eligible=False, itc_availability="No")
        result = self.service.reconcile([book], [gstr], as_of_month="2025-09")
        self.assertEqual(result.records[0].status, ReconciliationStatus.MATCHED_WITH_DIFFERENCE.value)
        self.assertIn("ITC unavailable", result.records[0].reason)
        self.assertEqual(result.monthly_summaries[0].gstr2b_eligible_itc, D(0))
        self.assertEqual(result.monthly_summaries[0].difference_amount, D("180"))

    def test_gstr2b_only_document_is_reported_in_section_b(self) -> None:
        result = self.service.reconcile(
            [],
            [transaction("g-only", "gstr2b", number="NOT-BOOKED", gstr_month="2025-09")],
            as_of_month="2025-09",
        )
        self.assertEqual(len(result.section_b_2b_not_in_books), 1)
        self.assertEqual(result.section_b_2b_not_in_books[0].status, ReconciliationStatus.GSTR2B_ONLY.value)

    def test_taxable_cgst_sgst_and_igst_mismatches_are_reported_by_component(self) -> None:
        variants = (
            ("taxable", "1010", "1000", "taxable_value"),
            ("cgst", "91", "90", "cgst"),
            ("sgst", "91", "90", "sgst"),
            ("igst", "1", "0", "igst"),
            ("cess", "1", "0", "cess"),
        )
        for field, book_value, gstr_value, difference_field in variants:
            with self.subTest(field=field):
                book_values = {"taxable": "1000", "cgst": "90", "sgst": "90", "igst": "0", "cess": "0"}
                gstr_values = dict(book_values)
                book_values[field] = book_value
                gstr_values[field] = gstr_value
                record = self.reconcile_pair(
                    transaction("b1", "books", **book_values),
                    transaction("g1", "gstr2b", gstr_month="2025-09", **gstr_values),
                )
                self.assertEqual(record.status, ReconciliationStatus.MATCHED_WITH_DIFFERENCE.value)
                self.assertNotEqual(record.differences.get(difference_field, D(0)), D(0))

    def test_document_type_difference_is_not_reported_as_an_exact_match(self) -> None:
        record = self.reconcile_pair(
            transaction("b1", "books", doc_type="invoice"),
            transaction("g1", "gstr2b", doc_type="debit_note", gstr_month="2025-09"),
        )
        self.assertEqual(record.status, ReconciliationStatus.MATCHED_WITH_DIFFERENCE.value)
        self.assertIn("Document type differs", record.reason)

    def test_gstin_mismatch_is_a_probable_match_not_an_exact_match(self) -> None:
        record = self.reconcile_pair(
            transaction("b1", "books", gstin="24AAAAA1111A1Z1"),
            transaction("g1", "gstr2b", gstin="27BBBBB2222B1Z2", gstr_month="2025-09"),
        )
        self.assertEqual(record.status, ReconciliationStatus.POSSIBLE_MATCH.value)
        self.assertEqual(record.match_level, "DOCUMENT_NUMBER_AND_VALUES")
        self.assertIn("GSTIN differs", record.reason)

    def test_invoice_number_mismatch_with_close_values_is_probable_match(self) -> None:
        record = self.reconcile_pair(
            transaction("b1", "books", number="INV-1001"),
            transaction("g1", "gstr2b", number="INV-1OO1", gstr_month="2025-09"),
        )
        self.assertEqual(record.status, ReconciliationStatus.POSSIBLE_MATCH.value)
        self.assertEqual(record.match_level, "APPROXIMATE_VALUES")

    def test_credit_note_opposite_signs_normalize_and_match(self) -> None:
        book = transaction(
            "b-cn", "books", number="CN-1", doc_type="credit_note",
            taxable="10000", cgst="900", sgst="900", total="11800",
        )
        gstr = transaction(
            "g-cn", "gstr2b", number="CN-1", doc_type="C", gstr_month="2025-09",
            taxable="-10000", cgst="-900", sgst="-900", total="-11800",
        )
        record = self.reconcile_pair(book, gstr)
        self.assertEqual(record.status, ReconciliationStatus.MATCHED.value)
        self.assertEqual(record.books.document_type, "credit_note")
        self.assertEqual(record.gstr2b.document_type, "credit_note")

    def test_source_sign_and_document_alias_rules_can_be_overridden(self) -> None:
        config = ReconciliationConfig.from_dict({
            "document_type_aliases": {"books": {"credit note": "credit_note"}},
            "sign_conventions": {"books": {"credit_note": -1}},
        })
        result = ReconciliationService(config).reconcile(
            [transaction("b-cn", "books", number="CN-CFG", doc_type="Credit Note", taxable="10", cgst="1", sgst="1")],
            [transaction("g-cn", "gstr2b", number="CN-CFG", doc_type="C", gstr_month="2025-09", taxable="-10", cgst="-1", sgst="-1")],
            as_of_month="2025-09",
        )
        self.assertEqual(result.records[0].status, ReconciliationStatus.MATCHED.value)
        self.assertEqual(result.records[0].books.document_type, "credit_note")

    def test_repository_debit_note_register_maps_buyer_debit_note_to_supplier_credit_note(self) -> None:
        book = transaction("b-dn", "books", number="CN-2", doc_type="Debit Note", taxable="500", cgst="45", sgst="45")
        gstr = transaction("g-cn", "gstr2b", number="CN-2", doc_type="C", gstr_month="2025-09", taxable="-500", cgst="-45", sgst="-45", total="-590")
        record = self.reconcile_pair(book, gstr)
        self.assertEqual(record.status, ReconciliationStatus.MATCHED.value)
        self.assertEqual(record.books.document_type, "credit_note")

    def test_same_month_credit_note_is_matched(self) -> None:
        record = self.reconcile_pair(
            transaction("b-cn", "books", number="CN-3", doc_type="credit_note", taxable="100", cgst="9", sgst="9"),
            transaction("g-cn", "gstr2b", number="CN-3", doc_type="C", gstr_month="2025-09", taxable="-100", cgst="-9", sgst="-9"),
        )
        self.assertEqual(record.status, ReconciliationStatus.MATCHED.value)

    def test_previous_month_credit_note_in_current_2b_is_timing_difference(self) -> None:
        book = transaction("b-cn", "books", number="CN-4", doc_type="credit_note", books_month="2025-08", document_month="2025-08", taxable="100", cgst="9", sgst="9")
        gstr = transaction("g-cn", "gstr2b", number="CN-4", doc_type="C", document_month="2025-08", gstr_month="2025-09", taxable="-100", cgst="-9", sgst="-9")
        result = self.service.reconcile([book], [gstr], as_of_month="2025-09")
        record = result.records[0]
        self.assertEqual(record.status, ReconciliationStatus.CREDIT_NOTE_TIMING_DIFFERENCE.value)
        self.assertTrue(record.resolved)
        self.assertEqual(record.books_month, "2025-08")
        self.assertEqual(record.gstr2b_month, "2025-09")
        self.assertEqual(result.outstanding_items, [])

    def test_current_month_book_credit_note_carries_forward_and_clears_next_month(self) -> None:
        book = transaction("b-cn", "books", number="CN-5", doc_type="credit_note", taxable="100", cgst="9", sgst="9")
        first = self.service.reconcile([book], [], as_of_month="2025-09")
        self.assertEqual(first.records[0].status, ReconciliationStatus.CREDIT_NOTE_TIMING_DIFFERENCE.value)
        self.assertEqual(first.records[0].expected_gstr2b_month, "2025-10")
        self.assertEqual(len(first.outstanding_items), 1)

        gstr = transaction("g-cn", "gstr2b", number="CN-5", doc_type="C", gstr_month="2025-10", taxable="-100", cgst="-9", sgst="-9")
        next_period = self.service.reconcile([], [gstr], outstanding=first.outstanding_items, as_of_month="2025-10")
        self.assertEqual(len(next_period.records), 1)
        self.assertEqual(next_period.records[0].status, ReconciliationStatus.CREDIT_NOTE_TIMING_DIFFERENCE.value)
        self.assertTrue(next_period.records[0].resolved)
        self.assertEqual(next_period.records[0].gstr2b_month, "2025-10")
        self.assertEqual(next_period.outstanding_items, [])

    def test_late_match_after_multiple_months_clears_open_item(self) -> None:
        book = transaction("b1", "books", number="INV-2", books_month="2025-08", document_month="2025-08")
        pending = self.service.reconcile([book], [], as_of_month="2025-08")
        gstr = transaction("g1", "gstr2b", number="INV-2", gstr_month="2025-11", document_month="2025-08")
        cleared = self.service.reconcile([], [gstr], outstanding=pending.outstanding_items, as_of_month="2025-11")
        self.assertEqual(cleared.records[0].status, ReconciliationStatus.TIMING_DIFFERENCE.value)
        self.assertTrue(cleared.records[0].resolved)
        self.assertFalse(cleared.outstanding_items)

    def test_as_of_horizon_uses_latest_2b_period_not_future_books_month(self) -> None:
        books = [
            transaction("b-pending", "books", number="PENDING", books_month="2025-09", document_month="2025-09"),
            transaction("b-future", "books", number="FUTURE", books_month="2026-03", document_month="2026-03"),
        ]
        gstr = [transaction("g-unrelated", "gstr2b", gstin="27ZZZZZ9999Z1Z9", number="OTHER", gstr_month="2025-10")]
        result = self.service.reconcile(books, gstr)
        self.assertEqual(result.as_of_month, "2025-10")
        by_number = {row.books.document_number: row for row in result.records if row.books}
        self.assertEqual(by_number["PENDING"].status, ReconciliationStatus.TIMING_DIFFERENCE.value)

    def test_item_older_than_configured_window_is_books_only(self) -> None:
        config = ReconciliationConfig(max_future_months=2)
        result = ReconciliationService(config).reconcile(
            [transaction("b1", "books", books_month="2025-08", document_month="2025-08")],
            [],
            as_of_month="2025-11",
        )
        self.assertEqual(result.records[0].status, ReconciliationStatus.BOOKS_ONLY.value)
        self.assertEqual(result.outstanding_items, [])

    def test_non_gst_and_ineligible_zero_tax_rows_are_excluded_from_mismatch(self) -> None:
        non_gst = transaction("b-non", "books", gstin="", number="NON-1", cgst="0", sgst="0", igst="0", classification="auto", itc_eligible=None)
        ineligible = transaction("b-inel", "books", number="INEL-1", taxable="1000", total="1000", cgst="0", sgst="0", igst="0", classification="auto", itc_eligible=None)
        result = self.service.reconcile([non_gst, ineligible], [], as_of_month="2025-09")
        by_number = {record.books.document_number: record for record in result.records}
        self.assertEqual(by_number["NON-1"].status, ReconciliationStatus.NON_GST.value)
        self.assertEqual(by_number["INEL-1"].status, ReconciliationStatus.INELIGIBLE_ITC.value)
        summary = result.monthly_summaries[0]
        self.assertEqual(summary.non_gst_amount, D("1000"))
        self.assertEqual(summary.ineligible_taxable_value, D("1000"))
        self.assertEqual(summary.difference_amount, D("0"))

    def test_ineligible_books_entry_matched_to_a_2b_row_stays_out_of_itc_difference(self) -> None:
        book = transaction("b-inel", "books", number="INEL-2", taxable="1000", total="1000", cgst="0", sgst="0", classification="ineligible_itc", itc_eligible=False)
        gstr = transaction("g-inel", "gstr2b", number="INEL-2", gstr_month="2025-09", taxable="1000", cgst="90", sgst="90")
        result = self.service.reconcile([book], [gstr], as_of_month="2025-09")
        self.assertEqual(result.records[0].status, ReconciliationStatus.INELIGIBLE_ITC.value)
        self.assertEqual(result.monthly_summaries[0].difference_amount, D(0))

    def test_eligible_zero_tax_purchase_can_be_marked_explicitly(self) -> None:
        book = transaction("b-zero", "books", number="ZERO-1", cgst="0", sgst="0", igst="0", classification="eligible", itc_eligible=True)
        result = self.service.reconcile([book], [], as_of_month="2025-09")
        self.assertEqual(result.records[0].status, ReconciliationStatus.TIMING_DIFFERENCE.value)

    def test_mixed_cgst_sgst_and_igst_purchase_match(self) -> None:
        books = transaction("b-mix", "books", number="MIX-1", cgst="90", sgst="90", igst="0")
        gstr = transaction("g-mix", "gstr2b", number="MIX-1", gstr_month="2025-09", cgst="90", sgst="90", igst="0")
        self.assertEqual(self.reconcile_pair(books, gstr).status, ReconciliationStatus.MATCHED.value)
        books_igst = transaction("b-igst", "books", number="IGST-1", cgst="0", sgst="0", igst="180")
        gstr_igst = transaction("g-igst", "gstr2b", number="IGST-1", gstr_month="2025-09", cgst="0", sgst="0", igst="180")
        self.assertEqual(self.reconcile_pair(books_igst, gstr_igst).status, ReconciliationStatus.MATCHED.value)

    def test_books_duplicates_are_not_automatically_consumed(self) -> None:
        books = [transaction("b1", "books"), transaction("b2", "books")]
        gstr = [transaction("g1", "gstr2b", gstr_month="2025-09")]
        result = self.service.reconcile(books, gstr, as_of_month="2025-09")
        duplicate_books = [row for row in result.records if row.books and row.status == ReconciliationStatus.DUPLICATE.value]
        self.assertEqual(len(duplicate_books), 2)
        self.assertTrue(all(row.gstr2b is None for row in duplicate_books))

    def test_gstr2b_duplicates_are_flagged(self) -> None:
        books = [transaction("b1", "books")]
        gstr = [transaction("g1", "gstr2b", gstr_month="2025-09"), transaction("g2", "gstr2b", gstr_month="2025-09")]
        result = self.service.reconcile(books, gstr, as_of_month="2025-09")
        duplicates = [row for row in result.records if row.status == ReconciliationStatus.DUPLICATE.value]
        self.assertEqual(len(duplicates), 2)
        self.assertTrue(all(row.gstr2b is not None for row in duplicates))

    def test_amendment_supersedes_base_invoice_without_double_counting(self) -> None:
        book = transaction("b1", "books", number="INV-OLD", taxable="1100")
        base = transaction("g-base", "gstr2b", number="INV-OLD", gstr_month="2025-09", taxable="1000")
        amendment = transaction(
            "g-amend", "gstr2b", number="INV-NEW", gstr_month="2025-10", taxable="1100",
            is_amendment=True, amends_document_number="INV-OLD",
        )
        result = self.service.reconcile([book], [base, amendment], as_of_month="2025-10")
        self.assertEqual(len(result.records), 1)
        self.assertEqual(result.records[0].gstr2b.transaction_id, "g-amend")
        self.assertIn("g-base", result.records[0].gstr2b.metadata["superseded_transaction_ids"])
        self.assertEqual(result.records[0].status, ReconciliationStatus.TIMING_DIFFERENCE.value)
        self.assertEqual(result.monthly_summaries[0].gstr2b_taxable_value, D(0))
        self.assertEqual(result.monthly_summaries[1].gstr2b_taxable_value, D("1100"))

    def test_credit_note_original_invoice_relationship_is_preserved_and_not_double_counted(self) -> None:
        book_invoice = transaction("b-inv", "books", number="INV-10", taxable="1000", cgst="90", sgst="90")
        gstr_invoice = transaction("g-inv", "gstr2b", number="INV-10", gstr_month="2025-09", taxable="1000", cgst="90", sgst="90")
        book_note = transaction("b-cn", "books", number="CN-10", doc_type="credit_note", taxable="100", cgst="9", sgst="9", original_invoice_number="INV-10")
        gstr_note = transaction("g-cn", "gstr2b", number="CN-10", doc_type="C", gstr_month="2025-09", taxable="-100", cgst="-9", sgst="-9", original_invoice_number="INV-10")
        result = self.service.reconcile([book_invoice, book_note], [gstr_invoice, gstr_note], as_of_month="2025-09")
        self.assertEqual(len(result.records), 2)
        self.assertTrue(all(row.status == ReconciliationStatus.MATCHED.value for row in result.records))
        note_record = next(row for row in result.records if row.books.document_type == "credit_note")
        self.assertEqual(note_record.books.original_invoice_number, "INV-10")

    def test_monthly_summary_excludes_non_gst_ineligible_and_carries_timing_separately(self) -> None:
        book = transaction("b1", "books", number="CN-11", doc_type="credit_note", taxable="100", cgst="9", sgst="9")
        result = self.service.reconcile([book], [], as_of_month="2025-09")
        row = result.monthly_summaries[0]
        self.assertEqual(row.timing_difference, D("-18"))
        self.assertEqual(row.difference_amount, D("0"))
        self.assertEqual(row.books_eligible_itc, D("-18"))


if __name__ == "__main__":
    unittest.main()
