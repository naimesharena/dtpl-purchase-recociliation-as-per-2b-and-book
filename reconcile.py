#!/usr/bin/env python3
"""Entry-point script: run GSTR-2B vs Books reconciliation and emit Excel report.

Usage:
    python reconcile.py [--base DIR] [--output FILE] [--future-window N] [--tolerance R]
"""
from __future__ import annotations

import argparse
import os
import sys

from reconciliation import (
    ReconciliationConfig,
    ReconciliationEngine,
    ExcelReportBuilder,
)


def main(argv=None):
    parser = argparse.ArgumentParser(description="GSTR-2B vs Books Reconciliation")
    parser.add_argument("--base", default=os.path.dirname(os.path.abspath(__file__)),
                        help="Directory containing Books .xls files and gstr-2B.xls")
    parser.add_argument("--output", default=None,
                        help="Output Excel file path (default: Reconciliation_Report.xlsx in base dir)")
    parser.add_argument("--future-window", type=int, default=3,
                        help="How many future GSTR-2B months to search for timing differences (default: 3)")
    parser.add_argument("--tolerance", type=float, default=1.0,
                        help="Per-amount tolerance in rupees for exact match (default: 1)")
    args = parser.parse_args(argv)

    base = os.path.abspath(args.base)
    if not os.path.isfile(os.path.join(base, "gstr-2B.xls")):
        print(f"ERROR: gstr-2B.xls not found in {base}", file=sys.stderr)
        return 1

    output = args.output or os.path.join(base, "Reconciliation_Report.xlsx")

    cfg = ReconciliationConfig(
        tolerance_taxable=args.tolerance,
        tolerance_cgst=args.tolerance,
        tolerance_sgst=args.tolerance,
        tolerance_igst=args.tolerance,
        future_month_window=args.future_window,
    )

    engine = ReconciliationEngine(config=cfg)
    print(f"[1/4] Loading Books (Purchase + Debit Note registers) from {base} ...")
    print(f"[2/4] Loading GSTR-2B from {os.path.join(base, 'gstr-2B.xls')} ...")
    engine.load_from_directory(base)

    print(f"      Books transactions loaded: {len(engine.books_txs)}")
    print(f"      GSTR-2B transactions loaded: {len(engine.gstr2b_txs)}")

    print("[3/4] Running month-by-month reconciliation ...")
    matches, summaries = engine.run()

    matched = sum(1 for m in matches if m.status.value == "MATCHED")
    matched_wd = sum(1 for m in matches if m.status.value == "MATCHED_WITH_DIFFERENCE")
    timing = sum(1 for m in matches if m.status.value in ("TIMING_DIFFERENCE", "CREDIT_NOTE_TIMING_DIFFERENCE"))
    nongst = sum(1 for m in matches if m.status.value == "NON_GST")
    inelig = sum(1 for m in matches if m.status.value == "INELIGIBLE_ITC")
    bonly = sum(1 for m in matches if m.status.value == "BOOKS_ONLY")
    g2bonly = sum(1 for m in matches if m.status.value == "GSTR2B_ONLY")

    print("      ------- Reconciliation Totals -------")
    print(f"      MATCHED:                    {matched}")
    print(f"      MATCHED_WITH_DIFFERENCE:    {matched_wd}")
    print(f"      TIMING_DIFFERENCE (open):   {timing}")
    print(f"      NON_GST:                    {nongst}")
    print(f"      INELIGIBLE_ITC:             {inelig}")
    print(f"      BOOKS_ONLY:                 {bonly}")
    print(f"      GSTR2B_ONLY:                {g2bonly}")
    print(f"      Months processed:           {len(summaries)}")

    print(f"[4/4] Writing Excel report to {output} ...")
    builder = ExcelReportBuilder(output)
    builder.build(matches, summaries, books_txs=engine.books_txs, gstr2b_txs=engine.gstr2b_txs)
    print("Done.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
