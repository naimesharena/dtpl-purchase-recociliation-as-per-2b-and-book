"""Document identity, date and sign normalization helpers."""
from __future__ import annotations

import re
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from typing import Any

from .config import ReconciliationConfig
from .models import Transaction


def normalize_gstin(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "")).upper()


def normalize_document_number(value: Any) -> str:
    # Remove whitespace only. Slashes, dashes and other punctuation can be
    # meaningful parts of Indian supplier document numbers.
    return re.sub(r"\s+", "", str(value or "")).upper()


def normalize_document_type(raw_type: Any, source: str, config: ReconciliationConfig) -> str:
    raw = str(raw_type or "").strip().lower()
    key = re.sub(r"\s+", " ", raw)
    aliases = config.document_type_aliases.get(source, {})
    if key in aliases:
        return aliases[key]
    if raw in {"c", "cn", "credit note", "credit-note"}:
        return "credit_note"
    if raw in {"d", "dn", "debit note", "debit-note"}:
        # A purchaser's debit note commonly corresponds to a supplier credit
        # note; the repository default is configured above, not fixed here.
        return "debit_note"
    if "credit" in key:
        return "credit_note"
    if "debit" in key:
        return "debit_note"
    if "purchase" in key or "invoice" in key or key in {"regular", "sezwp", "sezwop", "de"}:
        return "invoice"
    return "unknown"


def normalize_amount(value: Decimal, transaction: Transaction, config: ReconciliationConfig) -> Decimal:
    source_rules = config.sign_conventions.get(transaction.source, {})
    multiplier = int(source_rules.get(transaction.document_type, 1))
    if transaction.document_type in {"credit_note", "debit_note"}:
        return abs(value) * Decimal(multiplier)
    return value * Decimal(multiplier)


def normalized_values(transaction: Transaction, config: ReconciliationConfig) -> dict[str, Decimal]:
    return {
        "taxable_value": normalize_amount(transaction.taxable_value, transaction, config),
        "cgst": normalize_amount(transaction.cgst, transaction, config),
        "sgst": normalize_amount(transaction.sgst, transaction, config),
        "igst": normalize_amount(transaction.igst, transaction, config),
        "cess": normalize_amount(transaction.cess, transaction, config),
    }


def parse_month(value: Any) -> str | None:
    """Parse month identifiers as YYYY-MM, date, MMYYYY or YYYYMM."""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.strftime("%Y-%m")
    if isinstance(value, date):
        return value.strftime("%Y-%m")
    text = str(value).strip()
    if not text:
        return None
    match = re.fullmatch(r"(\d{4})[-/]?(\d{1,2})", text)
    if match:
        year, month = int(match.group(1)), int(match.group(2))
        if 1 <= month <= 12:
            return f"{year:04d}-{month:02d}"
    match = re.fullmatch(r"(\d{1,2})[-/]?(\d{4})", text)
    if match:
        month, year = int(match.group(1)), int(match.group(2))
        if 1 <= month <= 12:
            return f"{year:04d}-{month:02d}"
    for fmt in ("%b %Y", "%B %Y", "%b,%Y", "%B,%Y", "%m/%Y"):
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m")  # noqa: DTZ007 - calendar month only; no time zone is involved.
        except ValueError:
            pass
    # Date-like values such as DD-MM-YYYY or DD/MM/YYYY.
    parsed = parse_date(value)
    return parsed.strftime("%Y-%m") if parsed else None


def parse_date(value: Any) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    if not text:
        return None
    for fmt in (
        "%d-%m-%Y", "%d/%m/%Y", "%d-%b-%Y", "%d-%B-%Y",
        "%Y-%m-%d", "%m/%d/%Y", "%d.%m.%Y", "%d-%m-%y",
        "%d/%m/%y", "%d-%b-%y", "%d-%b-%Y %H:%M:%S",
    ):
        try:
            return datetime.strptime(text, fmt).date()  # noqa: DTZ007 - source value is a local calendar date.
        except ValueError:
            pass
    # Excel date serials reach this helper only for known date columns.
    try:
        serial = Decimal(text)
        if serial >= 1 and serial < 2958466:
            return (datetime(1899, 12, 30) + timedelta(days=float(serial))).date()  # noqa: DTZ001 - Excel serials encode local date-only values.
    except (InvalidOperation, ValueError, OverflowError, TypeError):
        pass
    return None


def month_distance(later_month: str, earlier_month: str) -> int:
    """Signed calendar-month distance: positive means later_month is later."""
    ly, lm = map(int, later_month.split("-"))
    ey, em = map(int, earlier_month.split("-"))
    return (ly * 12 + lm) - (ey * 12 + em)


def add_month(month: str | None, months: int) -> str | None:
    if not month:
        return None
    year, number = map(int, month.split("-"))
    absolute = year * 12 + (number - 1) + months
    return f"{absolute // 12:04d}-{absolute % 12 + 1:02d}"


def excel_serial_to_date(value: Any) -> date | None:
    """Excel's 1900 date system (used by the uploaded register exports)."""
    try:
        serial = Decimal(str(value))
    except (InvalidOperation, ValueError, OverflowError, TypeError):
        return parse_date(value)
    if serial < 1 or serial >= 2958466:
        return parse_date(value)
    return (datetime(1899, 12, 30) + timedelta(days=float(serial))).date()  # noqa: DTZ001 - Excel serials encode local date-only values.
