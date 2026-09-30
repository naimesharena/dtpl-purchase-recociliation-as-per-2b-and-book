# Software Format Reconciliation (separate from the main project)

Reconciliation reports built from the accounting-software export
**`software format.xls`** (repo root) — "GSTR-2B Reconciliation Details,
Month: All" for **DRIVE TRUCKING PRIVATE LIMITED (24AAJCD4457C1ZV), FY 2025-26**.

Everything in this folder is self-contained: it reads only `software format.xls`
and writes only into this folder. The main project files (`reconcile.py`,
`make_deliverables.py`, the PURCHASE / DEBIT NOTE / gstr-2B registers and the
`GST_ITC_Reconciliation_FY2025-26.*` deliverables) are **not touched**.

## Input

`software format.xls` (legacy .xls, xlrd-readable) contains the software's own
matching result, sheet per document type:

| Sheet     | Rows  | Contents |
|-----------|-------|----------|
| `invoice` | 9,121 | invoice reconciliation: As-per-Records side, As-per-GSTR-2B side, per-row difference block (2B − Books) |
| `note`    |   58  | credit note reconciliation (same layout + Note Type) |
| `invoice_ims`, `note_ims` | 0 data | IMS blocks, empty for this FY |

Status values (as reported by the software): `Matched`, `Partly Mat`
(matched on invoice no/GSTIN but with differences), `Not in Rec` (in GSTR-2B,
not in Books), `Not in 2B` (in Books, not in GSTR-2B).

## Run

```bash
python read_software.py     # parse the .xls -> software_data.pkl
python make_reports.py      # -> Software_Format_Reconciliation_FY2025-26.xlsx + .md
```

(Any venv with `pandas`, `xlrd`, `openpyxl` works; `read_software.py` has built-in
sanity checks: status/side consistency and difference-block identity.)

## Output — `Software_Format_Reconciliation_FY2025-26.xlsx`

| Sheet | What it is |
|-------|------------|
| `0. Executive Summary` | status mix, ITC totals (Books vs GSTR-2B), where the difference comes from, top unmatched suppliers |
| `A. Monthly Bridge` | month-wise ITC bridge: Books → month differences → value differences → (Not in 2B) → (Not in Rec) → **= GSTR-2B**, plus Final Difference; FY total block |
| `A2. Bridge Line Details` | every document behind every bridge line (filter by Month + Line); sums tie out exactly to sheet A |
| `B. Unmatched Details` | all "Not in Rec" + "Not in 2B" rows, Books \| GSTR-2B \| Diff column blocks |
| `C. Value Differences` | matched pairs where amounts differ (Books \| 2B \| Diff + mismatched fields) |
| `D. Field Mismatches` | "Partly Matched" rows with no amount difference (date/POS/other field differences) |
| `E. Matched (Clean)` | fully equal matched pairs |

A companion Markdown summary is written to `Software_Format_Reconciliation_FY2025-26.md`.

## Verified

* Bridge identity holds for **all 12 months**:
  `Books + month-diff + value-diff − (Not in 2B) + (Not in Rec) = GSTR-2B` (max deviation < 0.01).
* A2 detail rows sum to the sheet-A line values exactly, every month and line.
* Totals agree with the source file's own totals row (Books 93,120,792.92 taxable /
  17,588,016.25 GST; GSTR-2B 159,646,479.24 / 30,416,937.02; net diff +12,828,920.77 GST).

## Key observations

* The software reports a large **"Not in Rec"** pool: 4,603 invoices + 58 credit
  notes present in GSTR-2B but not found in its records (GST ≈ 13.25 M). This is
  the main driver of the net GSTR-2B > Books difference.
* **Month-difference lines are zero** — the software's matching always pairs the
  books and 2B sides in the same period, so no cross-month timing adjustment is
  needed (the rows are kept in the bridge for design consistency).
* All 58 GSTR-2B credit notes are "Not in Rec" because the software records hold
  no debit note entries (books side of the note sheet is empty). The book-side
  debit note entries live in the separate DEBIT NOTE register used by the main
  project reconciliation, so part of this pool is that registration/timing gap.
