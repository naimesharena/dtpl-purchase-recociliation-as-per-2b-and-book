"""Build Excel reconciliation report with multiple sheets."""
from __future__ import annotations

from typing import Dict, List

import pandas as pd

from .models import ReconciliationMatch, MonthlySummary, ReconciliationStatus


STATUS_ORDER = [
    ReconciliationStatus.MATCHED,
    ReconciliationStatus.MATCHED_WITH_DIFFERENCE,
    ReconciliationStatus.TIMING_DIFFERENCE,
    ReconciliationStatus.CREDIT_NOTE_TIMING_DIFFERENCE,
    ReconciliationStatus.NON_GST,
    ReconciliationStatus.INELIGIBLE_ITC,
    ReconciliationStatus.BOOKS_ONLY,
    ReconciliationStatus.GSTR2B_ONLY,
    ReconciliationStatus.DUPLICATE,
    ReconciliationStatus.POSSIBLE_MATCH,
]


class ExcelReportBuilder:
    """Produce a multi-sheet Excel workbook with reconciliation output."""

    def __init__(self, output_path: str):
        self.output_path = output_path

    def build(self,
              matches: List[ReconciliationMatch],
              summaries: Dict[str, MonthlySummary],
              books_txs=None,
              gstr2b_txs=None) -> str:
        with pd.ExcelWriter(self.output_path, engine="xlsxwriter") as writer:
            wb = writer.book

            # Formats
            header_fmt = wb.add_format({
                "bold": True, "bg_color": "#1F4E78", "font_color": "white",
                "border": 1, "text_wrap": True, "valign": "top",
            })
            money_fmt = wb.add_format({"num_format": "#,##0.00"})
            pct_fmt = wb.add_format({"num_format": "0.00%"})

            # ---- Sheet 1: README / Summary ----
            self._write_summary_sheet(writer, summaries, header_fmt, money_fmt)

            # ---- One sheet per month (detailed reconciliation) ----
            months = sorted(summaries.keys())
            for month in months:
                self._write_month_sheet(writer, month, matches, summaries[month], header_fmt, money_fmt)

            # ---- Section A: Books not in GSTR-2B (all months combined) ----
            self._write_books_not_in_2b(writer, matches, header_fmt, money_fmt)

            # ---- Section B: GSTR-2B not in Books (all months combined) ----
            self._write_2b_not_in_books(writer, matches, header_fmt, money_fmt)

            # ---- Timing differences (open items) ----
            self._write_timing_open(writer, matches, header_fmt, money_fmt)

            # ---- Credit notes (with sign-normalization view) ----
            self._write_credit_notes(writer, matches, header_fmt, money_fmt)

            # ---- All matched / diff ----
            self._write_all_status(writer, "All Matched",
                                   [ReconciliationStatus.MATCHED, ReconciliationStatus.MATCHED_WITH_DIFFERENCE],
                                   matches, header_fmt, money_fmt)

            # ---- Full dump ----
            self._write_all_matches(writer, matches, header_fmt, money_fmt)

            # ---- Source data sheets ----
            if books_txs is not None:
                bdf = pd.DataFrame([t.to_row() for t in books_txs])
                bdf.to_excel(writer, sheet_name="Books - Normalized", index=False)
                self._format_df_sheet(writer, "Books - Normalized", bdf, header_fmt, money_fmt)
            if gstr2b_txs is not None:
                gdf = pd.DataFrame([t.to_row() for t in gstr2b_txs])
                gdf.to_excel(writer, sheet_name="GSTR2B - Normalized", index=False)
                self._format_df_sheet(writer, "GSTR2B - Normalized", gdf, header_fmt, money_fmt)

        return self.output_path

    # ------------------------------------------------------------------
    # Sheet writers
    # ------------------------------------------------------------------
    def _write_summary_sheet(self, writer, summaries, header_fmt, money_fmt):
        rows = []
        for month in sorted(summaries.keys()):
            s = summaries[month]
            rows.append({
                "Month": self._pretty_month(s.month),
                "Books Count": s.books_count,
                "Books Taxable Value": s.books_taxable,
                "Books CGST": s.books_cgst,
                "Books SGST": s.books_sgst,
                "Books IGST": s.books_igst,
                "Books Total ITC": s.books_total_itc,
                "GSTR-2B Count": s.gstr2b_count,
                "GSTR-2B Taxable Value": s.gstr2b_taxable,
                "GSTR-2B CGST": s.gstr2b_cgst,
                "GSTR-2B SGST": s.gstr2b_sgst,
                "GSTR-2B IGST": s.gstr2b_igst,
                "GSTR-2B Total ITC": s.gstr2b_total_itc,
                "Matched Count": s.matched_count,
                "Matched Taxable": s.matched_taxable,
                "Matched ITC": s.matched_itc,
                "Matched w/ Diff Count": s.matched_with_diff_count,
                "Timing Diff Count": s.timing_diff_count,
                "Timing Diff ITC": s.timing_diff_itc,
                "Non-GST Count": s.non_gst_count,
                "Non-GST Amount": s.non_gst_amount,
                "Ineligible ITC Count": s.ineligible_itc_count,
                "Ineligible ITC Amount": s.ineligible_itc_amount,
                "Books Only Count": s.books_only_count,
                "Books Only ITC": s.books_only_itc,
                "GSTR2B Only Count": s.gstr2b_only_count,
                "GSTR2B Only ITC": s.gstr2b_only_itc,
                "Difference ITC (Books - 2B)": s.books_total_itc - s.gstr2b_total_itc,
            })
        df = pd.DataFrame(rows)
        df.to_excel(writer, sheet_name="Monthly Summary", index=False)
        self._format_df_sheet(writer, "Monthly Summary", df, header_fmt, money_fmt)

        # Add guide / how-to-read
        ws = writer.sheets["Monthly Summary"]
        guide_row = len(df) + 3
        ws.write(guide_row, 0, "Reconciliation Status Guide", header_fmt)
        guide = [
            ("MATCHED", "Exact match on GSTIN + Doc Number + Doc Type; values within tolerance."),
            ("MATCHED_WITH_DIFFERENCE", "Document found but taxable/GST values differ; needs review."),
            ("TIMING_DIFFERENCE", "Books entry this month but GSTR-2B appearing in future month (carried forward)."),
            ("CREDIT_NOTE_TIMING_DIFFERENCE", "Same as above for credit/debit notes."),
            ("NON_GST", "Non-GST purchase (no GSTIN, no GST); not expected in GSTR-2B."),
            ("INELIGIBLE_ITC", "GSTIN present but GST not separately claimed — ineligible ITC."),
            ("BOOKS_ONLY", "Present in Books, not in GSTR-2B even after future-month window."),
            ("GSTR2B_ONLY", "Present in GSTR-2B, missing from Books."),
            ("POSSIBLE_MATCH", "Fuzzy match requiring manual review."),
        ]
        ws.write(guide_row + 1, 0, "Status", header_fmt)
        ws.write(guide_row + 1, 1, "Meaning", header_fmt)
        for i, (st, desc) in enumerate(guide, start=guide_row + 2):
            ws.write(i, 0, st)
            ws.write(i, 1, desc)

    def _write_month_sheet(self, writer, month, matches, summary, header_fmt, money_fmt):
        name = f"M-{self._pretty_month(month).replace(' ', '-')}"
        mm = [m for m in matches if m.reconciliation_month == month]
        # Sort by status order, then supplier
        order_idx = {s: i for i, s in enumerate(STATUS_ORDER)}
        mm.sort(key=lambda x: (order_idx.get(x.status, 99),
                               (x.books_tx.supplier_name if x.books_tx
                                else (x.gstr2b_tx.supplier_name if x.gstr2b_tx else ""))))
        df = pd.DataFrame([m.to_row() for m in mm])
        if df.empty:
            df = pd.DataFrame([{"status": "(no data)"}])
        df.to_excel(writer, sheet_name=name[:31], index=False)
        self._format_df_sheet(writer, name[:31], df, header_fmt, money_fmt)

    def _write_books_not_in_2b(self, writer, matches, header_fmt, money_fmt):
        stati = [ReconciliationStatus.BOOKS_ONLY,
                 ReconciliationStatus.TIMING_DIFFERENCE,
                 ReconciliationStatus.CREDIT_NOTE_TIMING_DIFFERENCE,
                 ReconciliationStatus.NON_GST,
                 ReconciliationStatus.INELIGIBLE_ITC]
        rows = [m for m in matches if m.status in stati]
        df = pd.DataFrame([m.to_row() for m in rows])
        if df.empty:
            df = pd.DataFrame([{"status": "(no items)"}])
        df.to_excel(writer, sheet_name="Sec-A - Books Not In 2B", index=False)
        self._format_df_sheet(writer, "Sec-A - Books Not In 2B", df, header_fmt, money_fmt)

    def _write_2b_not_in_books(self, writer, matches, header_fmt, money_fmt):
        rows = [m for m in matches if m.status == ReconciliationStatus.GSTR2B_ONLY]
        df = pd.DataFrame([m.to_row() for m in rows])
        if df.empty:
            df = pd.DataFrame([{"status": "(no items)"}])
        df.to_excel(writer, sheet_name="Sec-B - 2B Not In Books", index=False)
        self._format_df_sheet(writer, "Sec-B - 2B Not In Books", df, header_fmt, money_fmt)

    def _write_timing_open(self, writer, matches, header_fmt, money_fmt):
        rows = [m for m in matches if m.status in (
            ReconciliationStatus.TIMING_DIFFERENCE,
            ReconciliationStatus.CREDIT_NOTE_TIMING_DIFFERENCE,
        )]
        df = pd.DataFrame([m.to_row() for m in rows])
        if df.empty:
            df = pd.DataFrame([{"status": "(no open timing differences)"}])
        df.to_excel(writer, sheet_name="Open Timing Differences", index=False)
        self._format_df_sheet(writer, "Open Timing Differences", df, header_fmt, money_fmt)

    def _write_credit_notes(self, writer, matches, header_fmt, money_fmt):
        rows = []
        for m in matches:
            b = m.books_tx
            g = m.gstr2b_tx
            if (b and b.document_type.value == "CREDIT_NOTE") or (g and g.document_type.value == "CREDIT_NOTE"):
                rows.append(m)
        df = pd.DataFrame([m.to_row() for m in rows])
        if df.empty:
            df = pd.DataFrame([{"status": "(no credit notes)"}])
        df.to_excel(writer, sheet_name="Credit Notes Reconciliation", index=False)
        self._format_df_sheet(writer, "Credit Notes Reconciliation", df, header_fmt, money_fmt)

    def _write_all_status(self, writer, sheet_name, statuses, matches, header_fmt, money_fmt):
        rows = [m for m in matches if m.status in statuses]
        df = pd.DataFrame([m.to_row() for m in rows])
        if df.empty:
            df = pd.DataFrame([{"status": "(none)"}])
        df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
        self._format_df_sheet(writer, sheet_name[:31], df, header_fmt, money_fmt)

    def _write_all_matches(self, writer, matches, header_fmt, money_fmt):
        df = pd.DataFrame([m.to_row() for m in matches])
        df.to_excel(writer, sheet_name="All Reconciliation Items", index=False)
        self._format_df_sheet(writer, "All Reconciliation Items", df, header_fmt, money_fmt)

    # ------------------------------------------------------------------
    # Formatting helpers
    # ------------------------------------------------------------------
    def _format_df_sheet(self, writer, sheet_name, df, header_fmt, money_fmt):
        ws = writer.sheets[sheet_name]
        # Header
        for col_idx, col in enumerate(df.columns):
            ws.write(0, col_idx, col, header_fmt)
        # Column widths + money format
        for col_idx, col in enumerate(df.columns):
            max_len = max(
                [len(str(col))] + [len(str(v)) for v in df[col].astype(str).head(200).tolist()]
            ) if len(df) else len(str(col))
            width = min(max(max_len + 2, 12), 40)
            fmt = None
            if any(k in col.lower() for k in ("taxable", "cgst", "sgst", "igst", "itc", "value",
                                              "amount", "diff_", "invoice_value", "gross")):
                fmt = money_fmt
            ws.set_column(col_idx, col_idx, width, fmt)
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, max(len(df), 1), len(df.columns) - 1 if len(df.columns) else 0)

    @staticmethod
    def _pretty_month(mmyyyy: str) -> str:
        if not mmyyyy or len(mmyyyy) != 6:
            return mmyyyy
        m = int(mmyyyy[:2])
        y = int(mmyyyy[2:])
        names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                 "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        return f"{names[m-1]}-{y}"
