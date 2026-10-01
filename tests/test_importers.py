from __future__ import annotations

import importlib.util
import unittest
from pathlib import Path

from gstr2b_recon import (
    ReconciliationConfig,
    ReconciliationService,
    ReconciliationStatus,
)
from gstr2b_recon.importers import read_books_workbooks, read_gstr2b_workbooks

ROOT = Path(__file__).resolve().parents[1]
HAS_XLRD = importlib.util.find_spec("xlrd") is not None


@unittest.skipUnless(HAS_XLRD, "xlrd is required to open the repository's legacy GSTR-2B .xls")
class RepositoryWorkbookImportTests(unittest.TestCase):
    def test_tally_xlsx_register_imports_alternate_voucher_key_and_amount_columns(self) -> None:
        path = ROOT / "PURCHASE-PL.xls"
        rows = read_books_workbooks([path], ReconciliationConfig())
        row = next(item for item in rows if item.document_number == "1633190092")
        self.assertEqual(row.supplier_gstin, "33AABCF1590N1ZJ")
        self.assertEqual(row.books_month, "2025-04")
        self.assertEqual(row.document_month, "2025-04")
        self.assertEqual(str(row.taxable_value), "117467.48")
        self.assertIn("3325052779", row.metadata["document_number_aliases"])

    def test_legacy_gstr_workbook_maps_note_sign_type_and_filing_month(self) -> None:
        path = ROOT / "gstr-2B.xls"
        rows = read_gstr2b_workbooks([path], config=ReconciliationConfig(), period_source="filing-month")
        note = next(item for item in rows if item.source_sheet == "B2B-CDNR" and item.document_number == "25-26/2")
        self.assertEqual(note.document_type, "credit_note")
        self.assertEqual(note.gstr2b_month, "2025-05")
        self.assertEqual(note.supplier_report_month, "2025-04")
        self.assertEqual(note.cgst, note.sgst)

        books = read_books_workbooks([ROOT / "PURCHASE-PL.xls"], ReconciliationConfig())
        book_invoice = next(row for row in books if row.document_number == "1633190092")
        gstr_invoice = next(
            row for row in rows
            if row.source_sheet == "B2B" and row.supplier_gstin == book_invoice.supplier_gstin
            and row.document_number == "3325052779"
        )
        reconciled = ReconciliationService(ReconciliationConfig()).reconcile(
            [book_invoice], [gstr_invoice], as_of_month="2025-05"
        )
        self.assertEqual(reconciled.records[0].status, ReconciliationStatus.TIMING_DIFFERENCE.value)
        self.assertEqual(reconciled.records[0].match_level, "GSTIN_DOCUMENT_NUMBER")

    def test_explicit_month_overrides_source_period_for_a_monthly_export(self) -> None:
        rows = read_gstr2b_workbooks(
            [ROOT / "gstr-2B.xls"],
            config=ReconciliationConfig(),
            gstr2b_period="2025-09",
            period_source="supplier-period",
        )
        self.assertTrue(rows)
        self.assertEqual({row.gstr2b_month for row in rows}, {"2025-09"})
        self.assertEqual(rows[0].metadata["gstr2b_month_basis"], "explicit-report-period")


if __name__ == "__main__":
    unittest.main()
