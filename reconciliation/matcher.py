"""Hierarchical matcher between Books and GSTR-2B transactions.

Matching levels (applied per month, with bi-directional carry-forward):

Level 1 — Exact match on (gstin, document_type, normalized_doc_number) AND values
          match within tolerance.
Level 2 — Same key, values differ but within tolerance for total ITC
          → MATCHED_WITH_DIFFERENCE.
Level 3 — Timing difference:
            * Books entry in month M but GSTR-2B in a future month (forward carry)
            * Books entry in month M relates to GSTR-2B in a recent past month
              (late booking / previous-month credit received now — backward match)
Level 4 — NON_GST / INELIGIBLE_ITC (pre-classified by parser).
Level 5 — Fall-through → BOOKS_ONLY / GSTR2B_ONLY.
"""
from __future__ import annotations

from collections import defaultdict
from typing import Dict, List, Tuple, Optional

from .models import (
    NormalizedTransaction,
    ReconciliationMatch,
    ReconciliationStatus,
    ReconciliationConfig,
    DocumentType,
)


def _doc_key(tx: NormalizedTransaction) -> Tuple[str, str, str]:
    g = (tx.supplier_gstin or "").strip().upper()
    d = tx.document_type.value
    n = (tx.document_number or "").strip().upper().replace(" ", "").replace("/", "").replace("-", "").replace(".", "")
    return (g, d, n)


def _values_match(b: NormalizedTransaction, g: NormalizedTransaction, cfg: ReconciliationConfig) -> Tuple[bool, Dict[str, float]]:
    d_t = abs(b.normalized_taxable - g.normalized_taxable)
    d_c = abs(b.normalized_cgst - g.normalized_cgst)
    d_s = abs(b.normalized_sgst - g.normalized_sgst)
    d_i = abs(b.normalized_igst - g.normalized_igst)
    exact = (d_t <= cfg.tolerance_taxable and d_c <= cfg.tolerance_cgst
             and d_s <= cfg.tolerance_sgst and d_i <= cfg.tolerance_igst)
    return exact, {"taxable": d_t, "cgst": d_c, "sgst": d_s, "igst": d_i}


def _total_itc_diff(b: NormalizedTransaction, g: NormalizedTransaction) -> float:
    return abs(b.normalized_total_itc - g.normalized_total_itc)


def _month_key(m: str) -> Tuple[int, int]:
    return (int(m[2:]), int(m[:2]))


def _add_months(month: str, delta: int) -> str:
    if not month or len(month) != 6:
        return ""
    mm = int(month[:2])
    yy = int(month[2:])
    total = yy * 12 + (mm - 1) + delta
    ny, nm = divmod(total, 12)
    return f"{nm + 1:02d}{ny}"


def _months_between(a: str, b: str) -> int:
    ka, kb = _month_key(a), _month_key(b)
    return (kb[0] - ka[0]) * 12 + (kb[1] - ka[1])


class ReconciliationMatcher:
    """Perform hierarchical matching for a single month's reconciliation."""

    def __init__(self, config: ReconciliationConfig | None = None):
        self.config = config or ReconciliationConfig()

    def match_month(
        self,
        month: str,
        books_for_month: List[NormalizedTransaction],
        gstr2b_by_month: Dict[str, List[NormalizedTransaction]],
        outstanding_books: List[NormalizedTransaction] | None = None,
        outstanding_2b: List[NormalizedTransaction] | None = None,
    ) -> Tuple[List[ReconciliationMatch], List[NormalizedTransaction], List[NormalizedTransaction]]:
        """
        Returns (matches, outstanding_books_for_future, outstanding_2b_for_future).

        outstanding_books: book items from earlier months that we expected to
                           appear in 2B later — resolve against current 2B pool.
        outstanding_2b:    2B items from earlier months that we expected to be
                           booked later — resolve against current books pool.
        """
        cfg = self.config
        matches: List[ReconciliationMatch] = []

        # Build per-month 2B pools (immutable view)
        all_months_sorted = sorted(gstr2b_by_month.keys(), key=_month_key)

        # Active 2B pool for THIS reconciliation cycle = current month only.
        # Previous months' 2B entries that were GSTR2B_ONLY are now in outstanding_2b.
        current_2b = list(gstr2b_by_month.get(month, []))
        current_2b_by_key: Dict[Tuple[str, str, str], List[NormalizedTransaction]] = defaultdict(list)
        for g in current_2b:
            current_2b_by_key[_doc_key(g)].append(g)
        consumed_2b_ids: set = set()

        # Future 2B pools (peek only, no consumption)
        future_window = []
        for i in range(1, cfg.future_month_window + 1):
            fm = _add_months(month, i)
            if fm in gstr2b_by_month:
                future_window.append(fm)
        future_pools_by_key: Dict[str, Dict[Tuple[str, str, str], List[NormalizedTransaction]]] = {}
        for fm in future_window:
            pool: Dict[Tuple[str, str, str], List[NormalizedTransaction]] = defaultdict(list)
            for g in gstr2b_by_month.get(fm, []):
                pool[_doc_key(g)].append(g)
            future_pools_by_key[fm] = pool

        new_outstanding_books: List[NormalizedTransaction] = []
        new_outstanding_2b: List[NormalizedTransaction] = []

        # --------------------------------------------------------------
        # Step 1 — Try to resolve outstanding BOOKS carry-forwards
        #          (books we expected to see in current-month 2B).
        # --------------------------------------------------------------
        still_outstanding_books: List[NormalizedTransaction] = []
        for b in outstanding_books or []:
            expected = getattr(b, "_expected_month", "") or ""
            if expected and _months_between(month, expected) > 0:
                # Not yet due
                still_outstanding_books.append(b)
                continue
            m = self._match_in_pool(b, current_2b_by_key, consumed_2b_ids, cfg)
            if m is not None:
                m.reconciliation_month = month
                m.notes = (m.notes + " | Books carried forward from earlier month").strip(" |")
                if m.gstr2b_tx.gstr2b_month != b.books_month:
                    m.status = ReconciliationStatus.MATCHED_WITH_DIFFERENCE if (
                        m.status == ReconciliationStatus.MATCHED_WITH_DIFFERENCE
                    ) else ReconciliationStatus.MATCHED
                    m.notes += f" (books {b.books_month} vs 2B {m.gstr2b_tx.gstr2b_month})"
                matches.append(m)
                continue
            # Not found in current month — peek future
            found = self._peek(b, future_pools_by_key, cfg)
            if found:
                b._expected_month = found
                still_outstanding_books.append(b)
                status = (ReconciliationStatus.CREDIT_NOTE_TIMING_DIFFERENCE
                          if b.document_type == DocumentType.CREDIT_NOTE
                          else ReconciliationStatus.TIMING_DIFFERENCE)
                matches.append(ReconciliationMatch(
                    status=status, books_tx=b, reconciliation_month=month,
                    expected_gstr2b_month=found,
                    notes=f"Still outstanding; expected in {found}",
                ))
            else:
                matches.append(ReconciliationMatch(
                    status=ReconciliationStatus.BOOKS_ONLY, books_tx=b,
                    reconciliation_month=month,
                    notes="Timing-difference window expired without match",
                ))
        new_outstanding_books.extend(still_outstanding_books)

        # --------------------------------------------------------------
        # Step 2 — Resolve outstanding GSTR2B_ONLY carry-forwards against
        #          current books pool (late-booked items).
        # --------------------------------------------------------------
        consumed_books_keys: set = set()  # we track matched books by (id) via matches,
        # but easier: build current-books keyed list and track ids.
        books_by_id = {id(b): b for b in books_for_month}
        # Flatten current books for matching
        book_pool_by_key: Dict[Tuple[str, str, str], List[NormalizedTransaction]] = defaultdict(list)
        for b in books_for_month:
            if b.is_non_gst or b.is_ineligible_itc:
                continue  # these are classified separately
            book_pool_by_key[_doc_key(b)].append(b)
        matched_book_ids: set = set()

        still_outstanding_2b: List[NormalizedTransaction] = []
        for g in outstanding_2b or []:
            key = _doc_key(g)
            found_b = None
            for b in book_pool_by_key.get(key, []):
                if b.id in matched_book_ids:
                    continue
                exact, diffs = _values_match(b, g, cfg)
                close = (not exact and diffs["taxable"] <= max(cfg.tolerance_taxable*100, 100)
                         and _total_itc_diff(b, g) <= max(cfg.tolerance_total_itc*20, 100))
                if exact or close:
                    found_b = b
                    status = ReconciliationStatus.MATCHED if exact else ReconciliationStatus.MATCHED_WITH_DIFFERENCE
                    matched_book_ids.add(b.id)
                    matches.append(ReconciliationMatch(
                        status=status, books_tx=b, gstr2b_tx=g,
                        reconciliation_month=month,
                        difference_taxable=b.normalized_taxable - g.normalized_taxable,
                        difference_cgst=b.normalized_cgst - g.normalized_cgst,
                        difference_sgst=b.normalized_sgst - g.normalized_sgst,
                        difference_igst=b.normalized_igst - g.normalized_igst,
                        notes=f"Late-booked: 2B in {g.gstr2b_month}, booked in {b.books_month}",
                    ))
                    break
            if found_b:
                continue
            # Compute age (months since 2B was filed). Carry forward up to window.
            age = _months_between(g.gstr2b_month, month)
            if age <= cfg.future_month_window:
                still_outstanding_2b.append(g)
            # If expired: drop without adding another GSTR2B_ONLY (already reported
            # in the original month; summary counts still reflect it).
        new_outstanding_2b.extend(still_outstanding_2b)

        # --------------------------------------------------------------
        # Step 3 — Classify NON_GST / INELIGIBLE_ITC (never expected in 2B)
        # --------------------------------------------------------------
        eligible_books: List[NormalizedTransaction] = []
        for b in books_for_month:
            if b.id in matched_book_ids:
                continue
            if b.is_non_gst:
                matches.append(ReconciliationMatch(
                    status=ReconciliationStatus.NON_GST, books_tx=b,
                    reconciliation_month=month,
                    notes="Non-GST purchase (no GSTIN, no GST); not expected in GSTR-2B",
                ))
                continue
            if b.is_ineligible_itc:
                matches.append(ReconciliationMatch(
                    status=ReconciliationStatus.INELIGIBLE_ITC, books_tx=b,
                    reconciliation_month=month,
                    notes="Ineligible ITC (GSTIN present but GST not claimed)",
                ))
                continue
            eligible_books.append(b)

        # --------------------------------------------------------------
        # Step 4 — Match current eligible books against current 2B pool.
        # --------------------------------------------------------------
        # Also consider past N months' *unmatched* 2B (for late-booked items).
        # We already received those as outstanding_2b; here we also handle the
        # case where past GSTR2B items were reported GSTR2B_ONLY but the current
        # month books them — they arrived via outstanding_2b above.

        for b in eligible_books:
            if b.id in matched_book_ids:
                continue
            m = self._match_in_pool(b, current_2b_by_key, consumed_2b_ids, cfg)
            if m is not None:
                m.reconciliation_month = month
                matched_book_ids.add(b.id)
                matches.append(m)
                continue

            # Not found in current 2B — peek future (forward carry)
            found = self._peek(b, future_pools_by_key, cfg)
            if found:
                b._expected_month = found
                new_outstanding_books.append(b)
                matched_book_ids.add(b.id)
                status = (ReconciliationStatus.CREDIT_NOTE_TIMING_DIFFERENCE
                          if b.document_type == DocumentType.CREDIT_NOTE
                          else ReconciliationStatus.TIMING_DIFFERENCE)
                matches.append(ReconciliationMatch(
                    status=status, books_tx=b, reconciliation_month=month,
                    expected_gstr2b_month=found,
                    notes=f"Expected in GSTR-2B of {found}",
                ))
            else:
                # BOOKS_ONLY (no matching 2B found within window)
                matches.append(ReconciliationMatch(
                    status=ReconciliationStatus.BOOKS_ONLY, books_tx=b,
                    reconciliation_month=month,
                    notes="Not found in GSTR-2B within search window",
                ))

        # --------------------------------------------------------------
        # Step 5 — Remaining (unmatched) current-month 2B.
        #
        # Report as GSTR2B_ONLY in this month AND carry forward for up to
        # ``future_month_window`` months so that, if the books entry arrives
        # next month, it can be matched as a "late-booked" resolution.
        # When we carry it forward we do NOT re-report GSTR2B_ONLY each
        # subsequent month (to avoid double counting in the summary); the
        # late resolution shows up as a MATCHED entry in the later month.
        # --------------------------------------------------------------
        for g in current_2b:
            if g.id in consumed_2b_ids:
                continue
            note = ""
            if g.itc_available and g.itc_available.lower() == "no":
                note = "ITC not available/rejected per GSTR-2B"
            else:
                note = "Present in GSTR-2B but not found in Books; carrying forward"
                g._expected_book_month = _add_months(month, 1)
                new_outstanding_2b.append(g)
            matches.append(ReconciliationMatch(
                status=ReconciliationStatus.GSTR2B_ONLY, gstr2b_tx=g,
                reconciliation_month=month, notes=note,
            ))

        # Deduplicate outstanding lists
        new_outstanding_books = self._dedup(new_outstanding_books)
        new_outstanding_2b = self._dedup(new_outstanding_2b)
        return matches, new_outstanding_books, new_outstanding_2b

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _match_in_pool(self, b, pool_by_key, consumed, cfg):
        key = _doc_key(b)
        for g in pool_by_key.get(key, []):
            if g.id in consumed:
                continue
            exact, diffs = _values_match(b, g, cfg)
            if exact:
                consumed.add(g.id)
                return ReconciliationMatch(
                    status=ReconciliationStatus.MATCHED, books_tx=b, gstr2b_tx=g,
                    difference_taxable=diffs["taxable"], difference_cgst=diffs["cgst"],
                    difference_sgst=diffs["sgst"], difference_igst=diffs["igst"],
                    notes=f"Matched in {g.gstr2b_month}",
                )
            if (diffs["taxable"] <= max(cfg.tolerance_taxable*100, 100)
                    and _total_itc_diff(b, g) <= max(cfg.tolerance_total_itc*20, 100)):
                consumed.add(g.id)
                return ReconciliationMatch(
                    status=ReconciliationStatus.MATCHED_WITH_DIFFERENCE,
                    books_tx=b, gstr2b_tx=g,
                    difference_taxable=b.normalized_taxable - g.normalized_taxable,
                    difference_cgst=b.normalized_cgst - g.normalized_cgst,
                    difference_sgst=b.normalized_sgst - g.normalized_sgst,
                    difference_igst=b.normalized_igst - g.normalized_igst,
                    notes=(f"Matched with value difference in {g.gstr2b_month}; "
                           f"Δtaxable={diffs['taxable']:.2f} ΔITC={_total_itc_diff(b,g):.2f}"),
                )
        return None

    def _peek(self, b, future_pools, cfg) -> str:
        key = _doc_key(b)
        for m in sorted(future_pools.keys(), key=_month_key):
            for g in future_pools[m].get(key, []):
                d_t = abs(b.normalized_taxable - g.normalized_taxable)
                d_itc = _total_itc_diff(b, g)
                if d_t <= max(cfg.tolerance_taxable*500, 500) and d_itc <= max(cfg.tolerance_total_itc*100, 200):
                    return m
        return ""

    @staticmethod
    def _dedup(items):
        seen = set()
        out = []
        for it in items:
            if it.id in seen:
                continue
            seen.add(it.id)
            out.append(it)
        return out

    @staticmethod
    def _future_months_from(month, available, window):
        sorted_avail = sorted(available, key=_month_key)
        mk = _month_key(month)
        out = []
        for m in sorted_avail:
            if _month_key(m) > mk and _months_between(month, m) <= window:
                out.append(m)
        return out
