"""Command-line entry point: python -m gstr2b_recon."""
from __future__ import annotations

import argparse
import glob
from collections import Counter
from pathlib import Path

from .config import ReconciliationConfig
from .importers import read_books_workbooks, read_gstr2b_workbooks
from .normalization import parse_month
from .service import ReconciliationService
from .storage import load_outstanding_state, save_outstanding_state, write_reports


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="gstr2b-recon",
        description="Reconcile Tally Books purchase/debit-note registers with GSTR-2B exports.",
    )
    parser.add_argument("--books", nargs="+", help="Books purchase/debit-note workbooks (default: PURCHASE-*.xls and DEBIT NOTE-*.xls).")
    parser.add_argument("--gstr2b", nargs="+", help="One or more GSTR-2B workbooks (default: gstr-2B.xls).")
    parser.add_argument("--gstr2b-period", help="Authoritative report month (YYYY-MM) for a monthly 2B file; overrides per-row inference.")
    parser.add_argument(
        "--period-source", choices=("filing-month", "supplier-period", "none"), default="filing-month",
        help="When report month is not supplied, map 2B availability to the supplier filing month (default), supplier GSTR-1 period, or leave it unset.",
    )
    parser.add_argument("--as-of", help="Latest reconciliation month (YYYY-MM); defaults to the latest imported month.")
    parser.add_argument("--config", help="Optional JSON config for sign rules, matching tolerances and zero-tax classifications.")
    parser.add_argument("--future-months", type=int, help="Override the configured future search/carry-forward window.")
    parser.add_argument("--output-dir", default="reconciliation_output", help="Directory for CSV/JSON reports (default: reconciliation_output).")
    parser.add_argument("--state-file", help="Persistent open-item JSON file (default: OUTPUT_DIR/open_items.json).")
    parser.add_argument("--no-state", action="store_true", help="Do not load/save persistent carry-forward state.")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        config = ReconciliationConfig.load(args.config)
        if args.future_months is not None:
            config.max_future_months = max(0, args.future_months)
        as_of_month = parse_month(args.as_of) if args.as_of else None
        if args.as_of and not as_of_month:
            raise ValueError(f"Invalid --as-of {args.as_of!r}; use YYYY-MM")

        books_paths = [Path(value) for value in args.books] if args.books else _default_books()
        gstr_paths = [Path(value) for value in args.gstr2b] if args.gstr2b else _default_gstr2b()
        if not books_paths:
            raise FileNotFoundError("No Books workbooks found; pass --books with purchase and debit-note registers.")
        if not gstr_paths:
            raise FileNotFoundError("No GSTR-2B workbook found; pass --gstr2b.")

        books = read_books_workbooks(books_paths, config)
        gstr = read_gstr2b_workbooks(
            gstr_paths,
            config=config,
            gstr2b_period=args.gstr2b_period,
            period_source=args.period_source,
        )
        state_path = Path(args.state_file) if args.state_file else Path(args.output_dir) / "open_items.json"
        prior_items = [] if args.no_state else load_outstanding_state(state_path)
        service = ReconciliationService(config)
        result = service.reconcile(books, gstr, outstanding=prior_items, as_of_month=as_of_month)
        report_paths = write_reports(result, args.output_dir, config)
        if not args.no_state:
            save_outstanding_state(state_path, result.outstanding_items)

        counts = Counter(record.status for record in result.records)
        print(f"Books documents: {len(books)} | GSTR-2B documents: {len(gstr)} | reconciliation records: {len(result.records)}")
        print(f"As of: {result.as_of_month or 'unknown'} | open timing items: {len(result.outstanding_items)}")
        for status, count in sorted(counts.items()):
            print(f"  {status}: {count}")
        if not args.gstr2b_period:
            basis_counts = Counter(
                tx.metadata.get("gstr2b_month_basis", "unknown") for tx in gstr
            )
            print("2B period mapping: " + ", ".join(f"{key}={value}" for key, value in sorted(basis_counts.items())))
            if "filing-month" == args.period_source:
                print("Note: filing-month is a proxy for availability, not a guaranteed 2B tax-period month; use --gstr2b-period for a monthly file or an explicit month mapping when the export contains multiple report periods.")
        print("Reports:")
        for name, path in report_paths.items():
            print(f"  {name}: {path}")
        if not args.no_state:
            print(f"  carry-forward state: {state_path}")
        return 0
    except Exception as error:  # noqa: BLE001 - top-level CLI boundary reports any input/runtime failure cleanly.
        print(f"Error: {error}")
        return 2


def _default_books() -> list[Path]:
    candidates = sorted(set(glob.glob("PURCHASE-*.xls") + glob.glob("DEBIT NOTE-*.xls")))
    return [Path(value) for value in candidates]


def _default_gstr2b() -> list[Path]:
    path = Path("gstr-2B.xls")
    return [path] if path.exists() else []
