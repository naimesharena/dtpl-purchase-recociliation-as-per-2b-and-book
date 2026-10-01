"""Top-level reconciliation engine that orchestrates month-by-month reconciliation."""
from __future__ import annotations

import os
from collections import defaultdict
from typing import Dict, List, Tuple

from .models import (
    NormalizedTransaction,
    ReconciliationMatch,
    MonthlySummary,
    ReconciliationConfig,
    ReconciliationStatus,
)
from .normalizer import SignNormalizer
from .parser_books import BooksParser
from .parser_gstr2b import GSTR2BParser
from .matcher import ReconciliationMatcher, _month_key


class ReconciliationEngine:
    """Run full GSTR-2B vs Books reconciliation for all months present in data."""

    def __init__(self, config: ReconciliationConfig | None = None,
                 normalizer: SignNormalizer | None = None):
        self.config = config or ReconciliationConfig()
        self.normalizer = normalizer or SignNormalizer()
        self.matcher = ReconciliationMatcher(self.config)

        self.books_txs: List[NormalizedTransaction] = []
        self.gstr2b_txs: List[NormalizedTransaction] = []
        self.matches: List[ReconciliationMatch] = []
        self.summaries: Dict[str, MonthlySummary] = {}

    def load_from_directory(self, base_dir: str) -> None:
        bp = BooksParser.default_files(base_dir)
        books_raw = bp.parse_all()
        gp = GSTR2BParser(os.path.join(base_dir, "gstr-2B.xls"))
        g2b_raw = gp.parse_all()

        self.books_txs = self.normalizer.normalize(books_raw)
        self.gstr2b_txs = self.normalizer.normalize(g2b_raw)

    def load_from_lists(self, books: List[NormalizedTransaction],
                        gstr2b: List[NormalizedTransaction]) -> None:
        self.books_txs = self.normalizer.normalize(books)
        self.gstr2b_txs = self.normalizer.normalize(gstr2b)

    def run(self) -> Tuple[List[ReconciliationMatch], Dict[str, MonthlySummary]]:
        books_by_month: Dict[str, List[NormalizedTransaction]] = defaultdict(list)
        g2b_by_month: Dict[str, List[NormalizedTransaction]] = defaultdict(list)

        for b in self.books_txs:
            if b.books_month:
                books_by_month[b.books_month].append(b)
        for g in self.gstr2b_txs:
            if g.gstr2b_month:
                g2b_by_month[g.gstr2b_month].append(g)

        all_months = sorted(set(books_by_month.keys()) | set(g2b_by_month.keys()),
                            key=_month_key)

        outstanding_books: List[NormalizedTransaction] = []
        outstanding_2b: List[NormalizedTransaction] = []
        all_matches: List[ReconciliationMatch] = []
        summaries: Dict[str, MonthlySummary] = {}

        for month in all_months:
            month_matches, outstanding_books, outstanding_2b = self.matcher.match_month(
                month=month,
                books_for_month=books_by_month.get(month, []),
                gstr2b_by_month=dict(g2b_by_month),
                outstanding_books=outstanding_books,
                outstanding_2b=outstanding_2b,
            )
            all_matches.extend(month_matches)
            summaries[month] = self._build_summary(month, month_matches,
                                                   books_by_month.get(month, []),
                                                   g2b_by_month.get(month, []))

        self.matches = all_matches
        self.summaries = summaries
        return all_matches, summaries

    def _build_summary(self, month, matches, books_in_month, g2b_in_month):
        s = MonthlySummary(month=month)
        for b in books_in_month:
            s.books_taxable += b.normalized_taxable
            s.books_cgst += b.normalized_cgst
            s.books_sgst += b.normalized_sgst
            s.books_igst += b.normalized_igst
            s.books_total_itc += b.normalized_total_itc
            s.books_count += 1
        for g in g2b_in_month:
            s.gstr2b_taxable += g.normalized_taxable
            s.gstr2b_cgst += g.normalized_cgst
            s.gstr2b_sgst += g.normalized_sgst
            s.gstr2b_igst += g.normalized_igst
            s.gstr2b_total_itc += g.normalized_total_itc
            s.gstr2b_count += 1
        for m in matches:
            b = m.books_tx; g = m.gstr2b_tx
            if m.status == ReconciliationStatus.MATCHED:
                s.matched_count += 1
                if b:
                    s.matched_taxable += b.normalized_taxable
                    s.matched_itc += b.normalized_total_itc
            elif m.status == ReconciliationStatus.MATCHED_WITH_DIFFERENCE:
                s.matched_with_diff_count += 1
                if b:
                    s.matched_with_diff_taxable += b.normalized_taxable
                s.matched_with_diff_itc_diff += abs(m.difference_cgst + m.difference_sgst + m.difference_igst)
            elif m.status in (ReconciliationStatus.TIMING_DIFFERENCE,
                              ReconciliationStatus.CREDIT_NOTE_TIMING_DIFFERENCE):
                s.timing_diff_count += 1
                if b:
                    s.timing_diff_taxable += b.normalized_taxable
                    s.timing_diff_itc += b.normalized_total_itc
            elif m.status == ReconciliationStatus.NON_GST:
                s.non_gst_count += 1
                if b:
                    s.non_gst_amount += b.invoice_value or b.taxable_value
            elif m.status == ReconciliationStatus.INELIGIBLE_ITC:
                s.ineligible_itc_count += 1
                if b:
                    s.ineligible_itc_amount += b.invoice_value or b.taxable_value
            elif m.status == ReconciliationStatus.BOOKS_ONLY:
                s.books_only_count += 1
                if b:
                    s.books_only_taxable += b.normalized_taxable
                    s.books_only_itc += b.normalized_total_itc
            elif m.status == ReconciliationStatus.GSTR2B_ONLY:
                s.gstr2b_only_count += 1
                if g:
                    s.gstr2b_only_taxable += g.normalized_taxable
                    s.gstr2b_only_itc += g.normalized_total_itc
            elif m.status == ReconciliationStatus.DUPLICATE:
                s.duplicate_count += 1
            elif m.status == ReconciliationStatus.POSSIBLE_MATCH:
                s.possible_match_count += 1
        return s
