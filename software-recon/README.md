# Software Format Reconciliation (separate from the main project)

All work in this folder is self-contained: it only **reads** the files at the
repo root and writes only inside this folder. The main project files
(`reconcile.py`, `make_deliverables.py`, the registers and the
`GST_ITC_Reconciliation_FY2025-26.*` deliverables) are not modified.

## 1. PRIMARY deliverable - our data in the software format

**`GSTR-2B_Reconciliation_Software_Format_FY2025-26.xlsx`**

The reconciliation of **our uploaded files** — `gstr-2B.xls` +
`PURCHASE-HR/PL/VL.xls` + `DEBIT NOTE-HR/PL/VL.xls` (FY 2025-26,
DRIVE TRUCKING PRIVATE LIMITED, 24AAJCD4457C1ZV) — presented in the exact
layout of the sample software export **`software format.xls`**
("GSTR-2B Reconciliation Details, Month: All").

Sheets (headers copied verbatim from the sample file):

| Sheet | Rows | Contents |
|-------|------|----------|
| `invoice` | 9,767 | every book invoice (PURCHASE registers) + every GSTR-2B invoice not in books, in the As-per-Records \| As-per-GSTR-2B \| Difference layout with Status, per-row difference block (2B − Books) and a totals row |
| `note` | 67 | same for DEBIT NOTE register (book side, Note Type "Debit") vs GSTR-2B credit notes (Note Type "Credit") |
| `invoice_ims`, `note_ims` | 0 | IMS blocks, kept empty as in the sample |
| `0. Executive Summary` | - | status mix + ITC totals (Books net vs GSTR-2B net, net difference) + where the difference comes from |
| `A. Monthly Bridge` | - | month-wise ITC bridge: Books -> month differences -> value differences -> (Not in 2B) -> (Not in Rec) = GSTR-2B, plus Final Difference and FY total block |
| `A2. Bridge Line Details` | 12,454 | every document behind every bridge line (filter by Month + Line); sums tie out exactly to sheet A |

Statuses (same vocabulary as the software):
`Matched` (amounts equal), `Partly Mat` (matched with differences),
`Not in Rec` (in GSTR-2B, not in Books), `Not in 2B` (in Books, not in GSTR-2B).

Run:

```bash
python reconcile.py                    # repo root - main matching engine (writes /tmp pickles)
python software-recon/make_software_format.py
```

Verified:
* header rows identical to the sample, column counts 45/47/45/47
* totals row equals the register / GSTR-2B sums
* NET DIFFERENCE (GSTR-2B minus Books) = **+224,726.72 GST** - identical to the
  main project's headline figure, i.e. the same matching, different presentation
* Monthly Bridge ties out exactly for all 12 months
  (`Books + month-diff + value-diff - (Not in 2B) + (Not in Rec) = GSTR-2B`),
  and Books / GSTR-2B / 2B-minus-Books are month-by-month identical to the main
  project's Monthly Bridge
* every A2 detail row sums to its sheet-A line, every month (auto-checked)

Result at a glance (GST): Books invoices 29,310,878.53, less debit notes
116,178.18 → Books net 29,194,700.35; GSTR-2B invoices 29,938,553.72, less
credit notes 519,126.65 → 2B net 29,419,427.07. Invoices: 8,175 Matched,
707 Partly Mat, 747 Not in 2B, 138 Not in Rec. Notes: 52 Matched, 2 Partly Mat,
7 Not in 2B, 6 Not in Rec.

## 2. Analysis of the sample software export itself

**`Software_Format_Reconciliation_FY2025-26.xlsx` / `.md`** (see
`read_software.py`, `make_reports.py`) - the same bridge/drill-down treatment
applied to the uploaded `software format.xls` on its own (company-wide
software export, 9,179 rows). Kept for reference; not part of our file set.

Run:

```bash
python software-recon/read_software.py
python software-recon/make_reports.py
```
