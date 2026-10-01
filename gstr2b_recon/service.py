"""Reusable month-aware reconciliation engine with one-to-one matching."""
from __future__ import annotations

from collections import defaultdict
from collections.abc import Iterable, Sequence
from dataclasses import replace
from decimal import ROUND_FLOOR, Decimal
from hashlib import sha1
from itertools import product

from .config import ReconciliationConfig
from .models import (
    MonthlySummary,
    OutstandingItem,
    ReconciliationRecord,
    ReconciliationResult,
    ReconciliationStatus,
    Transaction,
)
from .normalization import (
    add_month,
    month_distance,
    normalize_document_number,
    normalize_document_type,
    normalize_gstin,
    normalized_values,
)

_AMOUNT_FIELDS = ("taxable_value", "cgst", "sgst", "igst", "cess")
_ITC_FIELDS = ("cgst", "sgst", "igst", "cess")


class ReconciliationService:
    """Reconcile Books and GSTR-2B transactions without tying the engine to Excel.

    The engine performs strict document-key matching first, then unique
    probable matching.  A matched document is represented by exactly one
    record, even when its Books and 2B months differ.
    """

    def __init__(self, config: ReconciliationConfig | None = None) -> None:
        self.config = config or ReconciliationConfig()

    def reconcile(
        self,
        books: Iterable[Transaction],
        gstr2b: Iterable[Transaction],
        *,
        outstanding: Iterable[OutstandingItem | Transaction] = (),
        as_of_month: str | None = None,
    ) -> ReconciliationResult:
        current_books = [self._prepare(t) for t in books]
        current_ids = {t.transaction_id for t in current_books}
        current_keys = {
            (t.supplier_gstin, t.document_type, number)
            for t in current_books
            for number in self._document_numbers(t)
        }
        carried_books: list[Transaction] = []
        for item in outstanding:
            tx = self._prepare(item.transaction if isinstance(item, OutstandingItem) else item)
            tx_keys = {
                (tx.supplier_gstin, tx.document_type, number)
                for number in self._document_numbers(tx)
            }
            if tx.transaction_id not in current_ids and not current_keys.intersection(tx_keys):
                carried_books.append(tx)
        all_books = carried_books + current_books
        gstr_transactions = [self._prepare(t) for t in gstr2b]
        gstr_transactions = self._apply_amendments(gstr_transactions)

        if as_of_month is None:
            gstr_months = [t.gstr2b_month for t in gstr_transactions if t.gstr2b_month]
            book_months = [t.books_month for t in all_books if t.books_month]
            # The latest imported 2B period defines the reconciliation horizon;
            # future-dated Books rows must not age out an older missing 2B item.
            as_of_month = max(gstr_months) if gstr_months else max(book_months) if book_months else None

        book_duplicates = self._find_duplicates(all_books)
        gstr_duplicates = self._find_duplicates(gstr_transactions)
        records: list[ReconciliationRecord] = []
        used_books: set[str] = set()
        used_gstr: set[str] = set()

        for tx in all_books:
            if tx.transaction_id in book_duplicates:
                group = book_duplicates[tx.transaction_id]
                records.append(self._record(
                    ReconciliationStatus.DUPLICATE,
                    books=tx,
                    reason="Duplicate Books document key; review the source entries before matching.",
                    possible_match_ids=group,
                ))
                used_books.add(tx.transaction_id)
        for tx in gstr_transactions:
            if tx.transaction_id in gstr_duplicates:
                group = gstr_duplicates[tx.transaction_id]
                records.append(self._record(
                    ReconciliationStatus.DUPLICATE,
                    gstr2b=tx,
                    reason="Duplicate GSTR-2B document key; duplicate excluded from matching totals.",
                    possible_match_ids=group,
                ))
                used_gstr.add(tx.transaction_id)

        active_books = [t for t in all_books if t.transaction_id not in used_books]
        active_gstr = [t for t in gstr_transactions if t.transaction_id not in used_gstr]
        candidate_edges = self._candidate_edges(active_books, active_gstr)
        selected, ambiguous_books, ambiguous_gstr = self._select_unambiguous(candidate_edges)

        for book, gstr, level, reason in selected:
            record = self._matched_record(book, gstr, level, reason)
            records.append(record)
            used_books.add(book.transaction_id)
            used_gstr.add(gstr.transaction_id)

        # A probable match is advisory only. Keep each source transaction in a
        # single primary category; ambiguous suggestions are attached as IDs.
        for book in active_books:
            if book.transaction_id in used_books:
                continue
            suggestions = ambiguous_books.get(book.transaction_id, [])
            classification, classification_reason = self._book_classification(book)
            if suggestions:
                record = self._record(
                    ReconciliationStatus.POSSIBLE_MATCH,
                    books=book,
                    reason="More than one plausible GSTR-2B candidate; no automatic match made.",
                    possible_match_ids=suggestions,
                )
            elif classification == "non_gst":
                record = self._record(
                    ReconciliationStatus.NON_GST,
                    books=book,
                    reason=classification_reason,
                )
            elif classification == "ineligible_itc":
                record = self._record(
                    ReconciliationStatus.INELIGIBLE_ITC,
                    books=book,
                    reason=classification_reason,
                )
            elif self._timing_is_open(book, as_of_month):
                status = self._timing_status(book)
                record = self._record(
                    status,
                    books=book,
                    reason="Eligible Books entry has not appeared in the imported GSTR-2B periods yet.",
                    expected_gstr2b_month=add_month(book.books_month, self.config.expected_lag_months),
                    resolved=False,
                )
            else:
                record = self._record(
                    ReconciliationStatus.BOOKS_ONLY,
                    books=book,
                    reason="No GSTR-2B document found within the configured month search window.",
                )
            records.append(record)
            used_books.add(book.transaction_id)

        for gstr in active_gstr:
            if gstr.transaction_id in used_gstr:
                continue
            suggestions = ambiguous_gstr.get(gstr.transaction_id, [])
            records.append(self._record(
                ReconciliationStatus.POSSIBLE_MATCH if suggestions else ReconciliationStatus.GSTR2B_ONLY,
                gstr2b=gstr,
                reason=(
                    "More than one plausible Books candidate; no automatic match made."
                    if suggestions else
                    "GSTR-2B document has no Books match in the configured month search window."
                ),
                possible_match_ids=suggestions,
            ))
            used_gstr.add(gstr.transaction_id)

        records.sort(key=self._record_sort_key)
        outstanding_items = self._make_outstanding(records)
        summaries = self._summarize(records, all_books, gstr_transactions)
        return ReconciliationResult(
            records=records,
            outstanding_items=outstanding_items,
            monthly_summaries=summaries,
            as_of_month=as_of_month,
        )

    def _prepare(self, transaction: Transaction) -> Transaction:
        """Copy and canonicalize a transaction; raw source values remain intact."""
        tx = replace(transaction, metadata=dict(transaction.metadata or {}))
        tx.source = (tx.source or "").lower()
        tx.supplier_gstin = normalize_gstin(tx.supplier_gstin)
        tx.document_type = normalize_document_type(tx.document_type, tx.source, self.config)
        tx.document_number = str(tx.document_number or "").strip()
        tx.amends_document_number = str(tx.amends_document_number or "").strip()
        tx.original_invoice_number = str(tx.original_invoice_number or "").strip()
        return tx

    def _apply_amendments(self, transactions: list[Transaction]) -> list[Transaction]:
        """Treat each revised 2B document as the effective row, not a duplicate.

        B2BA/CDNRA rows carry the original number as an alias for matching a
        Books document that still uses the pre-amendment number. If multiple
        amendments target the same original document, only the latest period
        remains active in the current snapshot.
        """
        amendment_groups: dict[tuple[str, str], list[Transaction]] = defaultdict(list)
        for tx in transactions:
            if tx.source == "gstr2b" and tx.is_amendment and tx.amends_document_number:
                amendment_groups[(tx.supplier_gstin, normalize_document_number(tx.amends_document_number))].append(tx)

        superseded_ids: set[str] = set()
        for (gstin, original_number), amendments in amendment_groups.items():
            amendments.sort(key=lambda t: (t.gstr2b_month or "", t.source_row or 0, t.transaction_id))
            latest = amendments[-1]
            superseded_document_numbers: list[str] = []
            superseded_transaction_ids: list[str] = []
            for earlier in amendments[:-1]:
                superseded_ids.add(earlier.transaction_id)
                superseded_document_numbers.append(earlier.document_number)
                superseded_transaction_ids.append(earlier.transaction_id)
            for candidate in transactions:
                if candidate.transaction_id == latest.transaction_id or candidate.is_amendment:
                    continue
                if (
                    candidate.supplier_gstin == gstin
                    and normalize_document_number(candidate.document_number) == original_number
                ):
                    superseded_ids.add(candidate.transaction_id)
                    superseded_document_numbers.append(candidate.document_number)
                    superseded_transaction_ids.append(candidate.transaction_id)
            latest.metadata["superseded_document_numbers"] = sorted(set(superseded_document_numbers))
            latest.metadata["superseded_transaction_ids"] = sorted(set(superseded_transaction_ids))

        output: list[Transaction] = []
        for tx in transactions:
            if tx.transaction_id in superseded_ids:
                continue
            output.append(tx)
        return output

    @staticmethod
    def _business_key(tx: Transaction) -> tuple[str, str, str] | None:
        number = normalize_document_number(tx.document_number)
        if not number:
            return None
        return (normalize_gstin(tx.supplier_gstin), number, tx.document_type)

    def _find_duplicates(self, transactions: Sequence[Transaction]) -> dict[str, list[str]]:
        grouped: dict[tuple[str, str, str], list[Transaction]] = defaultdict(list)
        for tx in transactions:
            key = self._business_key(tx)
            if key:
                grouped[key].append(tx)
        result: dict[str, list[str]] = {}
        for group in grouped.values():
            if len(group) > 1:
                ids = sorted(t.transaction_id for t in group)
                for tx in group:
                    result[tx.transaction_id] = [candidate for candidate in ids if candidate != tx.transaction_id]
        return result

    def _candidate_edges(
        self,
        books: Sequence[Transaction],
        gstr2b: Sequence[Transaction],
    ) -> list[tuple[Transaction, Transaction, tuple[int, int, Decimal], str, str]]:
        """Build candidate pairs from indexes rather than a books × 2B scan."""
        edges: list[tuple[Transaction, Transaction, tuple[int, int, Decimal], str, str]] = []
        by_gstin_document: dict[tuple[str, str], list[Transaction]] = defaultdict(list)
        by_number_type: dict[tuple[str, str], list[Transaction]] = defaultdict(list)
        by_value_bucket: dict[tuple[str, str, tuple[int, ...]], list[Transaction]] = defaultdict(list)
        gstr_values = {tx.transaction_id: normalized_values(tx, self.config) for tx in gstr2b}
        tolerance = max(self.config.value_tolerance, Decimal("0.01"))

        for gstr in gstr2b:
            numbers = self._document_numbers(gstr)
            for number in numbers:
                by_gstin_document[(gstr.supplier_gstin, number)].append(gstr)
                by_number_type[(number, gstr.document_type)].append(gstr)
            if gstr.supplier_gstin:
                bucket = self._amount_bucket(gstr_values[gstr.transaction_id], tolerance)
                by_value_bucket[(gstr.supplier_gstin, gstr.document_type, bucket)].append(gstr)

        for book in books:
            book_numbers = self._document_numbers(book)
            book_amounts = normalized_values(book, self.config)
            possible: dict[str, Transaction] = {}

            # Level 1/2: GSTIN + document number, including an amendment's
            # original number. Document type is then tested as a difference.
            if book.supplier_gstin:
                for number in book_numbers:
                    for gstr in by_gstin_document.get((book.supplier_gstin, number), ()):
                        possible[gstr.transaction_id] = gstr

            # Level 3: same GSTIN/type and nearby taxable/tax amount buckets,
            # even when invoice numbers differ. Neighbour buckets guarantee
            # values within the configured tolerance are considered.
            if book.supplier_gstin and book_numbers:
                bucket = self._amount_bucket(book_amounts, tolerance)
                for offsets in product((-1, 0, 1), repeat=len(_AMOUNT_FIELDS)):
                    nearby = tuple(value + offset for value, offset in zip(bucket, offsets))
                    for gstr in by_value_bucket.get((book.supplier_gstin, book.document_type, nearby), ()):
                        possible[gstr.transaction_id] = gstr

            # GSTIN mismatch candidate: only suggest when document number and
            # type are exact and all normalized value components are close.
            for number in book_numbers:
                for gstr in by_number_type.get((number, book.document_type), ()):
                    possible[gstr.transaction_id] = gstr

            for gstr in possible.values():
                if not self._within_search_window(book, gstr):
                    continue
                gstr_numbers = self._document_numbers(gstr)
                same_number = bool(book_numbers.intersection(gstr_numbers))
                same_gstin = bool(book.supplier_gstin and book.supplier_gstin == gstr.supplier_gstin)
                same_type = book.document_type == gstr.document_type
                gstr_amount = gstr_values[gstr.transaction_id]
                close = self._amounts_close(book_amounts, gstr_amount, self.config.value_tolerance)
                month_gap = self._month_gap(book, gstr)
                amount_delta = sum((abs(book_amounts[k] - gstr_amount[k]) for k in _AMOUNT_FIELDS), Decimal(0))

                if same_number and same_gstin:
                    level = "GSTIN_DOCUMENT_NUMBER" if same_type else "GSTIN_DOCUMENT_NUMBER_TYPE_DIFFERENCE"
                    reason = "Same supplier and document number; values and document type were compared."
                    rank = 0 if same_type else 1
                    edges.append((book, gstr, (rank, month_gap, amount_delta), level, reason))
                elif same_type and close and same_gstin and book_numbers and gstr_numbers:
                    edges.append((
                        book, gstr, (2, month_gap, amount_delta), "APPROXIMATE_VALUES",
                        "Probable match by supplier, document type and values; document number differs.",
                    ))
                elif same_number and same_type and close and (not same_gstin):
                    edges.append((
                        book, gstr, (3, month_gap, amount_delta), "DOCUMENT_NUMBER_AND_VALUES",
                        "Probable match by document number and values; GSTIN differs or is missing.",
                    ))
        return edges

    @staticmethod
    def _amount_bucket(values: dict[str, Decimal], tolerance: Decimal) -> tuple[int, ...]:
        return tuple(
            int((values[field] / tolerance).to_integral_value(rounding=ROUND_FLOOR))
            for field in _AMOUNT_FIELDS
        )

    @staticmethod
    def _document_numbers(transaction: Transaction) -> set[str]:
        values = {normalize_document_number(transaction.document_number)}
        aliases = transaction.metadata.get("document_number_aliases", []) if transaction.metadata else []
        if isinstance(aliases, str):
            aliases = [aliases]
        values.update(normalize_document_number(value) for value in aliases)
        if transaction.amends_document_number:
            values.add(normalize_document_number(transaction.amends_document_number))
        values.discard("")
        return values

    def _within_search_window(self, book: Transaction, gstr: Transaction) -> bool:
        if not book.books_month or not gstr.gstr2b_month:
            return True
        delta = month_distance(gstr.gstr2b_month, book.books_month)
        return -self.config.max_past_months <= delta <= self.config.max_future_months

    @staticmethod
    def _month_gap(book: Transaction, gstr: Transaction) -> int:
        if not book.books_month or not gstr.gstr2b_month:
            return 0
        return abs(month_distance(gstr.gstr2b_month, book.books_month))

    @staticmethod
    def _amounts_close(left: dict[str, Decimal], right: dict[str, Decimal], tolerance: Decimal) -> bool:
        return all(abs(left[field] - right[field]) <= tolerance for field in _AMOUNT_FIELDS)

    @staticmethod
    def _edge_rank(edge: tuple[Transaction, Transaction, tuple[int, int, Decimal], str, str]) -> tuple[int, int, Decimal]:
        return edge[2]

    def _select_unambiguous(
        self,
        edges: list[tuple[Transaction, Transaction, tuple[int, int, Decimal], str, str]],
    ) -> tuple[list[tuple[Transaction, Transaction, str, str]], dict[str, list[str]], dict[str, list[str]]]:
        by_book: dict[str, list[tuple[Transaction, Transaction, tuple[int, int, Decimal], str, str]]] = defaultdict(list)
        by_gstr: dict[str, list[tuple[Transaction, Transaction, tuple[int, int, Decimal], str, str]]] = defaultdict(list)
        for edge in edges:
            by_book[edge[0].transaction_id].append(edge)
            by_gstr[edge[1].transaction_id].append(edge)

        best_book: dict[str, tuple[Transaction, Transaction, tuple[int, int, Decimal], str, str]] = {}
        best_gstr: dict[str, tuple[Transaction, Transaction, tuple[int, int, Decimal], str, str]] = {}
        tied_book: dict[str, list[tuple[Transaction, Transaction, tuple[int, int, Decimal], str, str]]] = {}
        tied_gstr: dict[str, list[tuple[Transaction, Transaction, tuple[int, int, Decimal], str, str]]] = {}
        for tx_id, choices in by_book.items():
            ordered = sorted(choices, key=self._edge_rank)
            best_rank = ordered[0][2]
            ties = [edge for edge in ordered if edge[2] == best_rank]
            if len(ties) == 1:
                best_book[tx_id] = ties[0]
            else:
                tied_book[tx_id] = ties
        for tx_id, choices in by_gstr.items():
            ordered = sorted(choices, key=self._edge_rank)
            best_rank = ordered[0][2]
            ties = [edge for edge in ordered if edge[2] == best_rank]
            if len(ties) == 1:
                best_gstr[tx_id] = ties[0]
            else:
                tied_gstr[tx_id] = ties

        selected: list[tuple[Transaction, Transaction, str, str]] = []
        used_books: set[str] = set()
        used_gstr: set[str] = set()
        for book_id, edge in best_book.items():
            book, gstr, _rank, level, reason = edge
            other_best = best_gstr.get(gstr.transaction_id)
            if other_best is edge and book_id not in used_books and gstr.transaction_id not in used_gstr:
                selected.append((book, gstr, level, reason))
                used_books.add(book_id)
                used_gstr.add(gstr.transaction_id)

        ambiguous_books: dict[str, list[str]] = {}
        ambiguous_gstr: dict[str, list[str]] = {}
        for tx_id, choices in by_book.items():
            if tx_id not in used_books:
                ids = sorted({edge[1].transaction_id for edge in choices if edge[1].transaction_id not in used_gstr})
                if ids:
                    ambiguous_books[tx_id] = ids
        for tx_id, choices in by_gstr.items():
            if tx_id not in used_gstr:
                ids = sorted({edge[0].transaction_id for edge in choices if edge[0].transaction_id not in used_books})
                if ids:
                    ambiguous_gstr[tx_id] = ids
        return selected, ambiguous_books, ambiguous_gstr

    def _matched_record(
        self,
        book: Transaction,
        gstr: Transaction,
        level: str,
        reason: str,
    ) -> ReconciliationRecord:
        book_values = normalized_values(book, self.config)
        gstr_values = normalized_values(gstr, self.config)
        if level in {"APPROXIMATE_VALUES", "DOCUMENT_NUMBER_AND_VALUES"}:
            return self._record(
                ReconciliationStatus.POSSIBLE_MATCH,
                books=book,
                gstr2b=gstr,
                match_level=level,
                reason=reason + " Confirm the supplier/document identity before posting an adjustment.",
                differences={
                    field: book_values[field] - gstr_values[field]
                    for field in _AMOUNT_FIELDS
                    if abs(book_values[field] - gstr_values[field]) > self.config.exact_tolerance
                },
                resolved=False,
            )
        differences = {
            field: book_values[field] - gstr_values[field]
            for field in _AMOUNT_FIELDS
            if abs(book_values[field] - gstr_values[field]) > self.config.exact_tolerance
        }
        classification, classification_reason = self._book_classification(book)
        if classification in {"non_gst", "ineligible_itc"}:
            status = ReconciliationStatus.NON_GST if classification == "non_gst" else ReconciliationStatus.INELIGIBLE_ITC
            return self._record(
                status,
                books=book,
                gstr2b=gstr,
                match_level=level,
                reason=classification_reason + " The document was identified in 2B, but this Books classification is excluded from the ITC mismatch.",
                differences=differences,
                resolved=True,
            )
        if gstr.itc_eligible is False and classification == "eligible":
            return self._record(
                ReconciliationStatus.MATCHED_WITH_DIFFERENCE,
                books=book,
                gstr2b=gstr,
                match_level=level,
                reason=reason + " The 2B row is marked ITC unavailable/rejected; document presence does not make its tax eligible.",
                differences=differences,
                resolved=True,
            )
        type_difference = level == "GSTIN_DOCUMENT_NUMBER_TYPE_DIFFERENCE"
        if type_difference:
            reason += " Document type differs between Books and GSTR-2B."
        same_period = bool(book.books_month and gstr.gstr2b_month and book.books_month == gstr.gstr2b_month)
        periods_unavailable = not book.books_month or not gstr.gstr2b_month
        if not same_period and not periods_unavailable:
            status = self._timing_status(book)
            reason = "Matched to a different GSTR-2B month; " + reason
            expected = add_month(book.books_month, self.config.expected_lag_months)
            return self._record(
                status, books=book, gstr2b=gstr, match_level=level, reason=reason,
                differences=differences, expected_gstr2b_month=expected, resolved=True,
            )
        if differences or type_difference:
            return self._record(
                ReconciliationStatus.MATCHED_WITH_DIFFERENCE,
                books=book, gstr2b=gstr, match_level=level,
                reason=reason + (" Amount components exceed the exact-match tolerance." if differences else ""),
                differences=differences, resolved=True,
            )
        if periods_unavailable:
            reason += " Period comparison unavailable; document identity and values matched."
        return self._record(
            ReconciliationStatus.MATCHED,
            books=book, gstr2b=gstr, match_level=level, reason=reason, resolved=True,
        )

    @staticmethod
    def _timing_status(book: Transaction) -> ReconciliationStatus:
        return (
            ReconciliationStatus.CREDIT_NOTE_TIMING_DIFFERENCE
            if book.document_type == "credit_note"
            else ReconciliationStatus.TIMING_DIFFERENCE
        )

    def _timing_is_open(self, book: Transaction, as_of_month: str | None) -> bool:
        if not book.books_month or not as_of_month:
            return True
        age = month_distance(as_of_month, book.books_month)
        # A period ahead of the latest imported 2B is expected to appear later.
        return age <= self.config.max_future_months

    def _book_classification(self, book: Transaction) -> tuple[str, str]:
        classification = str(book.classification or "auto").strip().lower()
        aliases = {
            "non-gst": "non_gst", "non gst": "non_gst", "non_gst": "non_gst",
            "ineligible": "ineligible_itc", "ineligible_itc": "ineligible_itc",
            "eligible": "eligible", "auto": "auto", "": "auto",
        }
        classification = aliases.get(classification, classification)
        if classification == "non_gst":
            return "non_gst", book.classification_reason or "Books account is explicitly classified as non-GST."
        if classification == "ineligible_itc" or book.itc_eligible is False:
            return "ineligible_itc", book.classification_reason or "Books entry is explicitly classified as ineligible ITC."
        if classification == "eligible" or book.itc_eligible is True:
            return "eligible", book.classification_reason
        if book.document_type in {"credit_note", "debit_note"}:
            return "eligible", "Credit/debit note retained for document matching."
        if book.has_gst:
            return "eligible", "GST is recorded separately in Books."
        policy = self.config.zero_tax_with_gstin if book.supplier_gstin else self.config.zero_tax_without_gstin
        if policy == "non_gst":
            return "non_gst", "No GST components are recorded and the configured zero-tax policy classifies this as non-GST."
        if policy in {"ineligible", "ineligible_itc"}:
            return "ineligible_itc", "No GST components are recorded; configured zero-tax policy classifies this as ineligible ITC."
        return "eligible", "No GST components are recorded; the configured policy treats this entry as eligible for reconciliation."

    def _make_outstanding(self, records: Sequence[ReconciliationRecord]) -> list[OutstandingItem]:
        result: list[OutstandingItem] = []
        for record in records:
            if (
                record.books is None
                or record.resolved
                or record.status not in {
                    ReconciliationStatus.TIMING_DIFFERENCE.value,
                    ReconciliationStatus.CREDIT_NOTE_TIMING_DIFFERENCE.value,
                }
            ):
                continue
            values = normalized_values(record.books, self.config)
            result.append(OutstandingItem(
                transaction=record.books,
                expected_gstr2b_month=record.expected_gstr2b_month,
                normalized_taxable_value=values["taxable_value"],
                normalized_cgst=values["cgst"],
                normalized_sgst=values["sgst"],
                normalized_igst=values["igst"],
                normalized_cess=values["cess"],
                status=record.status,
            ))
        return result

    def _summarize(
        self,
        records: Sequence[ReconciliationRecord],
        books: Sequence[Transaction],
        gstr2b: Sequence[Transaction],
    ) -> list[MonthlySummary]:
        months: dict[str, MonthlySummary] = {}

        def summary(month: str | None) -> MonthlySummary | None:
            if not month:
                return None
            if month not in months:
                months[month] = MonthlySummary(month=month)
            return months[month]

        book_classes: dict[str, str] = {}
        for tx in books:
            row = summary(tx.books_month)
            if row is None:
                continue
            values = normalized_values(tx, self.config)
            classification, _ = self._book_classification(tx)
            book_classes[tx.transaction_id] = classification
            row.books_taxable_value += values["taxable_value"]
            row.books_cgst += values["cgst"]
            row.books_sgst += values["sgst"]
            row.books_igst += values["igst"]
            row.books_cess += values["cess"]
            book_itc = sum(values[field] for field in _ITC_FIELDS)
            row.books_total_itc += book_itc
            if classification not in {"non_gst", "ineligible_itc"}:
                row.books_eligible_itc += book_itc
            elif classification == "non_gst":
                row.non_gst_amount += values["taxable_value"]
            else:
                row.ineligible_taxable_value += values["taxable_value"]
                row.ineligible_itc_amount += book_itc

        for tx in gstr2b:
            row = summary(tx.gstr2b_month)
            if row is None:
                continue
            values = normalized_values(tx, self.config)
            row.gstr2b_taxable_value += values["taxable_value"]
            row.gstr2b_cgst += values["cgst"]
            row.gstr2b_sgst += values["sgst"]
            row.gstr2b_igst += values["igst"]
            row.gstr2b_cess += values["cess"]
            itc = sum(values[field] for field in _ITC_FIELDS)
            row.gstr2b_total_itc += itc
            if tx.itc_eligible is not False:
                row.gstr2b_eligible_itc += itc

        for record in records:
            status = record.status
            book = record.books
            gstr = record.gstr2b
            if status in {ReconciliationStatus.NON_GST.value, ReconciliationStatus.INELIGIBLE_ITC.value}:
                continue
            if status == ReconciliationStatus.DUPLICATE.value:
                # Duplicates remain visible but are not counted again in the
                # reconciliation delta (source totals above remain auditable).
                continue
            if status in {ReconciliationStatus.MATCHED.value, ReconciliationStatus.MATCHED_WITH_DIFFERENCE.value} and book and gstr:
                book_month_row = summary(book.books_month)
                if book_month_row:
                    book_itc = sum(normalized_values(book, self.config)[field] for field in _ITC_FIELDS)
                    gstr_itc = sum(normalized_values(gstr, self.config)[field] for field in _ITC_FIELDS)
                    eligible_book_itc = book_itc if book_classes.get(book.transaction_id, "eligible") == "eligible" else Decimal(0)
                    eligible_gstr_itc = gstr_itc if gstr.itc_eligible is not False else Decimal(0)
                    book_month_row.matched_amount += eligible_book_itc
                    if status == ReconciliationStatus.MATCHED_WITH_DIFFERENCE.value:
                        book_month_row.difference_amount += eligible_book_itc - eligible_gstr_itc
                        book_month_row.difference_taxable_value += (
                            normalized_values(book, self.config)["taxable_value"]
                            - normalized_values(gstr, self.config)["taxable_value"]
                        )
            elif status in {
                ReconciliationStatus.TIMING_DIFFERENCE.value,
                ReconciliationStatus.CREDIT_NOTE_TIMING_DIFFERENCE.value,
            } and book:
                book_row = summary(book.books_month)
                book_values = normalized_values(book, self.config)
                book_itc = sum(book_values[field] for field in _ITC_FIELDS)
                if book_row:
                    book_row.timing_difference += book_itc
                if gstr:
                    gstr_row = summary(gstr.gstr2b_month)
                    gstr_itc = sum(normalized_values(gstr, self.config)[field] for field in _ITC_FIELDS)
                    if gstr_row:
                        gstr_row.timing_difference_cleared += gstr_itc
                        gstr_row.matched_with_timing_amount += gstr_itc
                    if book_row:
                        book_row.difference_amount += book_itc - gstr_itc
                        book_row.difference_taxable_value += (
                            book_values["taxable_value"]
                            - normalized_values(gstr, self.config)["taxable_value"]
                        )
            elif status in {ReconciliationStatus.BOOKS_ONLY.value, ReconciliationStatus.POSSIBLE_MATCH.value} and book and not gstr:
                book_row = summary(book.books_month)
                if book_row and status == ReconciliationStatus.BOOKS_ONLY.value:
                    book_values = normalized_values(book, self.config)
                    book_itc = sum(book_values[field] for field in _ITC_FIELDS)
                    book_row.books_only_amount += book_values["taxable_value"]
                    book_row.books_only_itc += book_itc
                    book_row.difference_taxable_value += book_values["taxable_value"]
                    if book_classes.get(book.transaction_id, "eligible") == "eligible":
                        book_row.difference_amount += book_itc
            elif status in {ReconciliationStatus.GSTR2B_ONLY.value, ReconciliationStatus.POSSIBLE_MATCH.value} and gstr and not book:
                gstr_row = summary(gstr.gstr2b_month)
                if gstr_row and status == ReconciliationStatus.GSTR2B_ONLY.value:
                    gstr_values = normalized_values(gstr, self.config)
                    gstr_itc = sum(gstr_values[field] for field in _ITC_FIELDS)
                    gstr_row.gstr2b_only_amount += gstr_values["taxable_value"]
                    gstr_row.difference_taxable_value -= gstr_values["taxable_value"]
                    if gstr.itc_eligible is not False:
                        gstr_row.gstr2b_only_itc += gstr_itc
                        gstr_row.difference_amount -= gstr_itc

        return [months[key] for key in sorted(months)]

    @staticmethod
    def _record_sort_key(record: ReconciliationRecord) -> tuple[str, str, str]:
        month = record.books_month or record.gstr2b_month or "9999-99"
        number = record.books.document_number if record.books else record.gstr2b.document_number if record.gstr2b else ""
        return (month, record.status, number)

    @staticmethod
    def _record(
        status: ReconciliationStatus,
        *,
        books: Transaction | None = None,
        gstr2b: Transaction | None = None,
        match_level: str | None = None,
        reason: str = "",
        differences: dict[str, Decimal] | None = None,
        expected_gstr2b_month: str | None = None,
        resolved: bool = False,
        possible_match_ids: list[str] | None = None,
    ) -> ReconciliationRecord:
        keys = ":".join(filter(None, [books.transaction_id if books else "", gstr2b.transaction_id if gstr2b else "", status.value]))
        record_id = sha1(keys.encode("utf-8")).hexdigest()[:16]
        return ReconciliationRecord(
            record_id=record_id,
            status=status.value,
            books=books,
            gstr2b=gstr2b,
            match_level=match_level,
            reason=reason,
            differences=differences or {},
            expected_gstr2b_month=expected_gstr2b_month,
            resolved=resolved,
            possible_match_ids=possible_match_ids or [],
        )
