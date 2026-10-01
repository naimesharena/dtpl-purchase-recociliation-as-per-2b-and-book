# GSTR-2B vs Books Reconciliation (DTPL)

A Python reconciliation engine for matching Tally-exported Books data (Purchase Register + Debit Note Register) against GSTR-2B returns.

## Key Features

- **Sign normalization**: Books Debit Notes (= supplier Credit Notes) and GSTR-2B Credit Notes are both normalized to negative values before matching, so they can be compared directly.
- **Month-wise reconciliation** over the full financial year (Apr 2025 – Mar 2026).
- **Bi-directional timing-difference carry-forward**:
  - Books April, 2B July → reported as TIMING_DIFFERENCE in April, automatically MATCHED in July.
  - 2B April, Books May (late-booked invoice / previous-month credit) → GSTR2B_ONLY in April, MATCHED in May as "Late-booked".
- **Non-GST purchases**: entries without a GSTIN and zero GST are classified `NON_GST` and excluded from ITC mismatch.
- **Ineligible ITC**: entries with a GSTIN but zero GST (blocked credits not booked separately) are classified `INELIGIBLE_ITC`.
- **Hierarchical matching**: Exact → value-difference → fuzzy → timing → fall-through.
- **Configurable tolerance** and future-month search window.
- **Excel output** with per-month sheets, two required sections (Books not in 2B, 2B not in Books), open timing differences, credit-note view, monthly summary, and normalized data dumps.

## Quick Start

```bash
python reconcile.py
```

The script looks in the current directory for:

- `PURCHASE-HR.xls`, `PURCHASE-PL.xls`, `PURCHASE-VL.xls` (Tally Purchase Register exports — xlsx content)
- `DEBIT NOTE-HR.xls`, `DEBIT NOTE-PL.xls`, `DEBIT NOTE-VL.xls` (Tally Debit Note Register exports)
- `gstr-2B.xls` (GSTR-2B quarterly Excel from GST Portal)

And produces `Reconciliation_Report.xlsx`.

### CLI options

```
python reconcile.py --base <DIR> --output <FILE> --future-window N --tolerance R
```

- `--future-window N` — how many future GSTR-2B months to search for timing differences (default: 3).
- `--tolerance R` — per-amount tolerance in rupees for exact match (default: 1).

## Reconciliation Statuses

| Status | Meaning |
|--------|---------|
| MATCHED | Exact match (values within tolerance) |
| MATCHED_WITH_DIFFERENCE | Document found but taxable/GST values differ; review |
| TIMING_DIFFERENCE | Books entry this month; expected in future GSTR-2B |
| CREDIT_NOTE_TIMING_DIFFERENCE | Same as above for credit notes |
| NON_GST | Non-GST purchase (no GSTIN, no GST); not in 2B |
| INELIGIBLE_ITC | GSTIN present but GST not claimed |
| BOOKS_ONLY | In books, not in GSTR-2B within search window |
| GSTR2B_ONLY | In GSTR-2B, missing from books (or late-booked) |

## Project Layout

```
reconcile.py                      # CLI entry point
reconciliation/
├── __init__.py
├── models.py                     # Dataclasses: NormalizedTransaction, ReconciliationMatch, MonthlySummary, ReconciliationConfig, enums
├── normalizer.py                 # SignNormalizer + DefaultSignConvention
├── parser_books.py               # Tally Purchase/Debit-Note register parser
├── parser_gstr2b.py              # GSTR-2B .xls parser (B2B, B2B-CDNR, B2BA, B2B-CDNRA, rejected)
├── matcher.py                    # Hierarchical matcher with bi-directional carry-forward
├── engine.py                     # Orchestrates month-by-month reconciliation
└── excel_report.py               # Generates multi-sheet Excel report
tests/                            # Pytest test suite (31 tests)
```

## Running Tests

```bash
pip install pytest --break-system-packages
python -m pytest tests/ -v
```
