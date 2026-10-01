# GSTR-2B vs Books reconciliation

This repository arrived with seven spreadsheets and no application/database code: three Tally-style purchase registers, three debit-note registers and one GST-portal GSTR-2B export. This implementation reuses those files directly; it does not add a parallel bookkeeping model or alter the source workbooks.

## Data flow and source mapping

```text
PURCHASE-*.xls ─┐
DEBIT NOTE-*.xls ├─> workbook adapters ─> normalized Transactions ─> ReconciliationService
GSTR-2B.xls ─────┘                                                ├─> detail + two sections
                                                                    ├─> month summaries
                                                                    └─> outstanding open items
```

- `PURCHASE-*.xls` and `DEBIT NOTE-*.xls` are XLSX/OOXML workbooks with an `.xls` filename. The adapter reads their register headers and rows with Python's standard library.
- `gstr-2B.xls` is a legacy binary XLS export. The adapter reads its B2B invoices, B2BA amendments, B2B-CDNR notes and relevant rejected-record sheets. `xlrd` is the only runtime dependency.
- Books purchase rows use **Supplier Invoice No.** as the primary document number and keep **Voucher No.** as an alternate match key. Debit-note rows prefer **Voucher Ref. No.** and retain the voucher number as an alternate key. Book date, supplier-document date, and 2B month remain separate.
- Taxable value is derived from the non-tax ledger columns; CGST, SGST and IGST are read separately. The `All Items`, `Gross Total`, tax and round-off columns are not added to taxable value. If a row has no usable ledger detail, the importer falls back to gross less GST and flags that limitation in its ledger metadata.
- The source registers do not contain a dedicated ITC-eligibility field. Explicit ledger names (for example, `Non GST` or `Blocked ITC`) take precedence. Otherwise, zero-tax rows use the configurable fallback: no GSTIN → `NON_GST`; GSTIN present → `INELIGIBLE_ITC`. This is a reviewable heuristic, not a tax/legal determination; change it in the JSON config or mark the transaction explicitly in a future source adapter when your chart of accounts has a better rule.

## Sign convention and note mapping

All values are normalized to ITC impact before comparison: invoices are positive, supplier credit notes reduce ITC (negative), and supplier debit notes increase ITC (positive). Thus a positive Books value and a negative GSTR-2B value for the same credit note normalize to the same negative amount.

In the provided purchase-register convention, a buyer-side **Debit Note** is mapped to a supplier **Credit Note** by default. GSTR-2B note type `C` maps to `credit_note`; `D` maps to `debit_note`. Both the aliases and per-source sign multipliers are configurable in `config.example.json`. The model also preserves an optional link from a credit/debit note to its original invoice without using that link as the note's own matching key.

## Month logic and carry-forward

Each transaction has separate `books_month`, `document_month`, and `gstr2b_month` values. A match can span periods inside the configured past/future search window. A cross-period match has one primary status (`TIMING_DIFFERENCE` or `CREDIT_NOTE_TIMING_DIFFERENCE`) and one record containing both source rows; it is not also emitted as `BOOKS_ONLY` or `GSTR2B_ONLY`. An unmatched eligible Books row remains an open timing item until the search window expires. State is written to JSON and loaded on the next run so that a newly imported 2B period can clear a previous open item.

**Important for the supplied 2B workbook:** it aggregates many periods and does not retain a definitive per-document GSTR-2B availability month. The default `filing-month` mapping uses the supplier's GSTR-1 filing-date month as an availability proxy, falling back to the supplier-reported period when filing date is missing. This often represents the next-month appearance, but it is not guaranteed to equal the 2B tax period (late filing/cut-off effects can differ). For a single monthly 2B download, pass `--gstr2b-period YYYY-MM`; use `--period-source supplier-period` if you intentionally want the supplier's reported period instead. Do not interpret the default proxy as an official return-period assignment.

## Run

```bash
python -m pip install -r requirements.txt
python -m gstr2b_recon
```

With the supplied filenames, the command auto-discovers the purchase/debit-note registers and `gstr-2B.xls`. Example for a single monthly import:

```bash
python -m gstr2b_recon \
  --books PURCHASE-HR.xls PURCHASE-PL.xls PURCHASE-VL.xls \
         'DEBIT NOTE-HR.xls' 'DEBIT NOTE-PL.xls' 'DEBIT NOTE-VL.xls' \
  --gstr2b gstr-2B.xls \
  --gstr2b-period 2025-09 \
  --as-of 2025-09 \
  --output-dir reconciliation_output
```

If the 2B workbook contains multiple months, do not pass one `--gstr2b-period` for all rows. Use the supplied period column/date mapping (or prepare separate monthly files) and review the printed mapping note. `--state-file PATH` chooses a persistent open-item file; `--no-state` disables loading/saving it. `--future-months N` overrides the configured future search window.

## Reports

The output directory contains:

- `GSTR2B_vs_Books_Reconciliation.xlsx` — formatted workbook with Read Me, monthly summary, full detail, Section A, Section B and open timing items tabs.
- `reconciliation_detail.csv` — one primary classification per record, source details, raw/normalized amounts, difference components and timing information.
- `books_not_in_gstr2b.csv` — Section A: `BOOKS_ONLY`, pending/resolved timing, `NON_GST`, `INELIGIBLE_ITC`, and possible-match review items.
- `gstr2b_not_in_books.csv` — Section B: GSTR-2B-only and duplicate GSTR-2B items.
- `monthly_summary.csv` — source totals, eligible ITC, matched/timing amounts, excluded non-GST/ineligible values, ITC `difference_amount`, and a separate `difference_taxable_value` so value-only mismatches remain visible.
- `reconciliation.json` — full structured result.
- `outstanding_items.json` and `open_items.json` — open timing items (the latter is the persistent carry-forward state).

`MATCHED_WITH_DIFFERENCE` captures a document-key match with component differences. `POSSIBLE_MATCH` is advisory and requires review; it is never silently treated as a confirmed match. Duplicate rows are flagged and excluded from matching. Rejected/ITC-unavailable 2B rows retain their availability flag.

## Configuration and tests

Copy `config.example.json` and edit matching tolerance, timing windows, note aliases/sign rules or zero-tax classification policies. Core callers can use `Transaction` and `ReconciliationService` directly without Excel.

```bash
python -m unittest discover -s tests -v
```
