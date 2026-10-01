"""Adapters for the uploaded Tally-style XLSX registers and GST portal XLS."""
from __future__ import annotations

import posixpath
import re
import zipfile
from collections.abc import Iterable
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from hashlib import sha1
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .config import ReconciliationConfig
from .models import Transaction
from .normalization import excel_serial_to_date, parse_date, parse_month

_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_NS = {"m": _MAIN_NS, "r": _REL_NS, "p": _PKG_REL_NS}
_ZERO = Decimal(0)


def read_books_workbooks(
    paths: Iterable[str | Path],
    config: ReconciliationConfig | None = None,
) -> list[Transaction]:
    """Read purchase and debit-note registers into source-neutral records."""
    config = config or ReconciliationConfig()
    transactions: list[Transaction] = []
    for path_value in paths:
        path = Path(path_value)
        sheets = _read_workbook(path)
        for sheet_name, rows in sheets.items():
            header_index = _find_books_header(rows)
            if header_index is None:
                continue
            headers = [_text(value) for value in rows[header_index][1]]
            kind = _books_register_kind(sheet_name, headers)
            if kind is None:
                continue
            for excel_row, values in rows[header_index + 1:]:
                record = _row_dict(headers, values)
                if _is_empty_books_row(record):
                    continue
                tx = _book_transaction(path, sheet_name, excel_row, record, kind, config)
                if tx:
                    transactions.append(tx)
    return transactions


def read_gstr2b_workbooks(
    paths: Iterable[str | Path],
    *,
    config: ReconciliationConfig | None = None,
    gstr2b_period: str | None = None,
    period_source: str = "filing-month",
) -> list[Transaction]:
    """Read B2B, amendments and debit/credit notes from GST portal exports.

    A monthly report period supplied via ``gstr2b_period`` is authoritative.
    Otherwise, ``filing-month`` uses the supplier's GSTR-1 filing-date month as
    a practical availability proxy; ``supplier-period`` uses the period shown
    in the supplier's GSTR-1 row.  The distinction is important for quarterly
    or year-aggregated workbooks that do not retain the actual 2B availability
    month per document.
    """
    config = config or ReconciliationConfig()
    explicit_month = parse_month(gstr2b_period) if gstr2b_period else None
    if gstr2b_period and explicit_month is None:
        raise ValueError(f"Invalid --gstr2b-period {gstr2b_period!r}; use YYYY-MM")
    if period_source not in {"filing-month", "supplier-period", "none"}:
        raise ValueError("period_source must be filing-month, supplier-period or none")
    transactions: list[Transaction] = []
    for path_value in paths:
        path = Path(path_value)
        sheets = _read_workbook(path)
        transactions.extend(_parse_gstr_sheets(path, sheets, config, explicit_month, period_source))
    return transactions


def _read_workbook(path: Path) -> dict[str, list[tuple[int, list[Any]]]]:
    if not path.exists():
        raise FileNotFoundError(path)
    try:
        if zipfile.is_zipfile(path):
            return _read_xlsx_zip(path)
    except (OSError, zipfile.BadZipFile):
        pass
    try:
        import xlrd  # type: ignore[import-not-found]
    except ImportError as error:
        raise RuntimeError(
            f"{path} is a legacy .xls workbook. Install the project dependency with: python -m pip install -r requirements.txt"
        ) from error
    try:
        workbook = xlrd.open_workbook(str(path), on_demand=True)
    except Exception as error:
        raise ValueError(f"Could not read Excel workbook {path}: {error}") from error
    output: dict[str, list[tuple[int, list[Any]]]] = {}
    for sheet in workbook.sheets():
        output[sheet.name] = [
            (row_index + 1, [sheet.cell_value(row_index, col) for col in range(sheet.ncols)])
            for row_index in range(sheet.nrows)
        ]
    workbook.release_resources()
    return output


def _read_xlsx_zip(path: Path) -> dict[str, list[tuple[int, list[Any]]]]:
    """Small stdlib XLSX reader; handles workbooks uploaded with an .xls suffix."""
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {rel.attrib["Id"]: rel.attrib["Target"] for rel in relationships.findall("p:Relationship", _NS)}
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in names:
            shared_root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in shared_root.findall("m:si", _NS):
                shared_strings.append("".join(text.text or "" for text in item.iter(f"{{{_MAIN_NS}}}t")))
        output: dict[str, list[tuple[int, list[Any]]]] = {}
        sheets = workbook.find("m:sheets", _NS)
        if sheets is None:
            return output
        for sheet in sheets:
            rel_id = sheet.attrib.get(f"{{{_REL_NS}}}id")
            target = targets.get(rel_id or "")
            if not target:
                continue
            member = target.lstrip("/")
            if not member.startswith("xl/"):
                member = posixpath.normpath(posixpath.join("xl", member))
            if member not in names:
                continue
            root = ET.fromstring(archive.read(member))
            sheet_data = root.find("m:sheetData", _NS)
            parsed_rows: list[tuple[int, list[Any]]] = []
            if sheet_data is not None:
                for row in sheet_data.findall("m:row", _NS):
                    cells: dict[int, Any] = {}
                    max_col = 0
                    for cell in row.findall("m:c", _NS):
                        reference = cell.attrib.get("r", "")
                        col_index = _column_index(reference)
                        max_col = max(max_col, col_index + 1)
                        cell_type = cell.attrib.get("t", "")
                        value_node = cell.find("m:v", _NS)
                        raw = value_node.text if value_node is not None else ""
                        if cell_type == "s" and raw:
                            try:
                                value: Any = shared_strings[int(raw)]
                            except (ValueError, IndexError):
                                value = raw
                        elif cell_type == "inlineStr":
                            value = "".join(node.text or "" for node in cell.iter(f"{{{_MAIN_NS}}}t"))
                        elif cell_type == "b":
                            value = raw == "1"
                        else:
                            value = raw
                        cells[col_index] = value
                    parsed_rows.append((int(row.attrib.get("r", len(parsed_rows) + 1)), [cells.get(i, "") for i in range(max_col)]))
            output[sheet.attrib.get("name", member)] = parsed_rows
        return output


def _file_token(path: Path) -> str:
    """Stable path token; two monthly files with the same basename stay distinct."""
    return sha1(str(path.resolve()).encode("utf-8")).hexdigest()[:12]


def _column_index(reference: str) -> int:
    letters = "".join(character for character in reference if character.isalpha())
    value = 0
    for character in letters.upper():
        value = value * 26 + (ord(character) - ord("A") + 1)
    return max(0, value - 1)


def _books_register_kind(sheet_name: str, headers: list[str]) -> str | None:
    name = sheet_name.lower()
    if "debit note" in name or "credit note" in name:
        return "note"
    if "purchase" in name:
        return "purchase"
    lowered = [header.lower() for header in headers]
    if "voucher type" in lowered and any("supplier invoice" in header for header in lowered):
        return "purchase"
    if "voucher type" in lowered and any("voucher ref" in header for header in lowered):
        return "note"
    return None


def _find_books_header(rows: list[tuple[int, list[Any]]]) -> int | None:
    for index, (_row_number, values) in enumerate(rows):
        normalized = {_header_key(value) for value in values if _text(value)}
        if "voucher type" in normalized and ("gstin uin" in normalized or "gstin" in normalized):
            return index
    return None


def _header_key(value: Any) -> str:
    text = _text(value).lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _row_dict(headers: list[str], values: list[Any]) -> dict[str, Any]:
    output: dict[str, Any] = {}
    for index, header in enumerate(headers):
        if header:
            output[header] = values[index] if index < len(values) else ""
    return output


def _find_header(record: dict[str, Any], *candidates: str) -> Any:
    by_key = {_header_key(key): value for key, value in record.items()}
    for candidate in candidates:
        key = _header_key(candidate)
        if key in by_key:
            return by_key[key]
    return ""


def _is_empty_books_row(record: dict[str, Any]) -> bool:
    particulars = _text(_find_header(record, "Particulars"))
    if particulars.lower() == "grand total":
        return True
    voucher_type = _text(_find_header(record, "Voucher Type"))
    voucher_no = _text(_find_header(record, "Voucher No."))
    supplier_no = _text(_find_header(record, "Supplier Invoice No."))
    reference_no = _text(_find_header(record, "Voucher Ref. No."))
    return not any((voucher_type, voucher_no, supplier_no, reference_no))


def _book_transaction(
    path: Path,
    sheet_name: str,
    excel_row: int,
    record: dict[str, Any],
    kind: str,
    config: ReconciliationConfig,
) -> Transaction | None:
    book_date_raw = _find_header(record, "Date", "Voucher Date")
    book_date = _date_string(book_date_raw)
    book_month = parse_month(book_date) if book_date else None
    if kind == "purchase":
        raw_type = _text(_find_header(record, "Voucher Type")) or "Purchase"
        document_number = _text(_find_header(record, "Supplier Invoice No.", "Supplier Invoice Number"))
        if not document_number:
            document_number = _text(_find_header(record, "Voucher No.", "Voucher Number"))
        document_date_raw = _find_header(record, "Supplier Invoice Date", "Invoice Date")
    else:
        raw_type = _text(_find_header(record, "Voucher Type")) or "Debit Note"
        document_number = _text(_find_header(record, "Voucher Ref. No.", "Voucher Reference Number"))
        if not document_number:
            document_number = _text(_find_header(record, "Voucher No.", "Voucher Number"))
        document_date_raw = _find_header(record, "Voucher Ref. Date", "Note Date")
    document_date = _date_string(document_date_raw) or book_date
    document_month = parse_month(document_date) if document_date else book_month
    gstin = _text(_find_header(record, "GSTIN/UIN", "GSTIN"))
    supplier = _text(_find_header(record, "Particulars", "Supplier Name"))
    gross_total = _amount(_find_header(record, "Gross Total", "Total"))
    cgst = _amount(_find_header(record, "CGST TAX", "CGST"))
    sgst = _amount(_find_header(record, "SGST TAX", "SGST"))
    igst = _amount(_find_header(record, "IGST TAX", "IGST"))
    cess = _amount(_find_header(record, "CESS TAX", "CESS"))

    excluded_headers = {
        "date", "particulars", "voucher type", "voucher no", "voucher ref no",
        "voucher ref date", "supplier invoice no", "supplier invoice date",
        "gstin uin", "gstin", "shipping bill of entry no", "shipping bill of entry date",
        "bill of entry no", "bill of entry date", "gross total", "round off",
    }
    account_values: list[tuple[str, Decimal]] = []
    for header, value in record.items():
        key = _header_key(header)
        if key in excluded_headers or "tax" in key or "round off" in key or "all items" in key:
            continue
        parsed = _amount(value)
        if parsed:
            account_values.append((header, parsed))
    taxable_value = sum((value for _header, value in account_values), _ZERO)
    # Some Tally registers have no account column in the exported row. The
    # fallback is explicitly less reliable when TDS/other adjustments exist.
    if not taxable_value and gross_total:
        taxable_value = gross_total - cgst - sgst - igst - cess
    account_names = [header for header, _value in account_values]
    has_tax = any((cgst, sgst, igst, cess))
    explicit_ineligible = any(
        keyword in _header_key(name)
        for name in account_names
        for keyword in config.ineligible_account_keywords
    )
    non_gst_value = sum(
        (value for name, value in account_values if any(keyword in name.lower() for keyword in config.non_gst_account_keywords)),
        _ZERO,
    )
    if explicit_ineligible:
        classification = "ineligible_itc"
        classification_reason = "A Books ledger name identifies the entry as blocked/ineligible ITC."
        itc_eligible: bool | None = False
    elif not has_tax and account_values and abs(non_gst_value - taxable_value) <= Decimal("0.02"):
        classification = "non_gst"
        classification_reason = "All taxable ledger value is assigned to an explicitly named non-GST account."
        itc_eligible = False
    elif has_tax:
        classification = "eligible"
        classification_reason = "GST is recorded separately in the Books register."
        itc_eligible = True
    else:
        policy = config.zero_tax_with_gstin if gstin else config.zero_tax_without_gstin
        if policy in {"ineligible", "ineligible_itc"}:
            classification = "ineligible_itc"
            classification_reason = "No GST is separately recorded; zero-tax/GSTIN policy classifies this as ineligible ITC."
            itc_eligible = False
        elif policy == "non_gst":
            classification = "non_gst"
            classification_reason = "No GST is separately recorded; zero-tax policy classifies this as non-GST."
            itc_eligible = False
        else:
            classification = "eligible"
            classification_reason = "No GST is separately recorded; the configured zero-tax policy treats this entry as eligible for reconciliation."
            itc_eligible = None

    identifier = f"books:{path.name}:{_file_token(path)}:{sheet_name}:{excel_row}"
    raw_reference = _text(_find_header(record, "Voucher Ref. No."))
    voucher_number = _text(_find_header(record, "Voucher No.", "Voucher Number"))
    supplier_invoice_number = _text(_find_header(record, "Supplier Invoice No.", "Supplier Invoice Number"))
    aliases = sorted({value for value in (voucher_number, supplier_invoice_number, raw_reference) if value and value != document_number})
    return Transaction(
        transaction_id=identifier,
        source="books",
        supplier_gstin=gstin,
        supplier_name=supplier,
        document_number=document_number,
        document_type=raw_type,
        raw_document_type=raw_type,
        document_date=document_date,
        books_month=book_month,
        document_month=document_month,
        taxable_value=taxable_value,
        cgst=cgst,
        sgst=sgst,
        igst=igst,
        cess=cess,
        total_value=gross_total,
        classification=classification,
        itc_eligible=itc_eligible,
        source_file=path.name,
        source_sheet=sheet_name,
        source_row=excel_row,
        classification_reason=classification_reason,
        metadata={"raw_voucher_type": raw_type, "voucher_reference_number": raw_reference,
                  "document_number_aliases": aliases,
                  "supplier_invoice_number": supplier_invoice_number,
                  "voucher_number": voucher_number,
                  "ledger_components": {name: str(value) for name, value in account_values}},
    )


def _parse_gstr_sheets(
    path: Path,
    sheets: dict[str, list[tuple[int, list[Any]]]],
    config: ReconciliationConfig,
    explicit_month: str | None,
    period_source: str,
) -> list[Transaction]:
    output: list[Transaction] = []
    for sheet_name, rows in sheets.items():
        normalized_sheet = re.sub(r"\s+", " ", sheet_name.strip().lower())
        if normalized_sheet == "b2b":
            for row_number, row in rows[6:]:
                if not _text(_at(row, 0)) or not _text(_at(row, 2)):
                    continue
                output.append(_gstr_transaction(
                    path, sheet_name, row_number, config, explicit_month, period_source,
                    gstin=_at(row, 0), name=_at(row, 1), number=_at(row, 2), raw_type="invoice",
                    document_date=_at(row, 4), total=_at(row, 5), taxable=_at(row, 8),
                    igst=_at(row, 9), cgst=_at(row, 10), sgst=_at(row, 11), cess=_at(row, 12),
                    supplier_period=_at(row, 13), filing_date=_at(row, 14),
                    itc_availability=_at(row, 15), reason=_at(row, 16),
                ))
        elif normalized_sheet == "b2ba":
            for row_number, row in rows[7:]:
                if not _text(_at(row, 2)) or not _text(_at(row, 4)):
                    continue
                output.append(_gstr_transaction(
                    path, sheet_name, row_number, config, explicit_month, period_source,
                    gstin=_at(row, 2), name=_at(row, 3), number=_at(row, 4), raw_type="invoice",
                    document_date=_at(row, 6), total=_at(row, 7), taxable=_at(row, 10),
                    igst=_at(row, 11), cgst=_at(row, 12), sgst=_at(row, 13), cess=_at(row, 14),
                    supplier_period=_at(row, 15), filing_date=_at(row, 16),
                    itc_availability=_at(row, 17), reason=_at(row, 18),
                    amendment=True, amends_document_number=_text(_at(row, 0)),
                ))
        elif normalized_sheet == "b2b-cdnr":
            for row_number, row in rows[6:]:
                if not _text(_at(row, 0)) or not _text(_at(row, 2)):
                    continue
                output.append(_gstr_transaction(
                    path, sheet_name, row_number, config, explicit_month, period_source,
                    gstin=_at(row, 0), name=_at(row, 1), number=_at(row, 2), raw_type=_at(row, 3),
                    document_date=_at(row, 5), total=_at(row, 6), taxable=_at(row, 9),
                    igst=_at(row, 10), cgst=_at(row, 11), sgst=_at(row, 12), cess=_at(row, 13),
                    supplier_period=_at(row, 14), filing_date=_at(row, 15),
                    itc_availability=_at(row, 16), reason=_at(row, 17),
                ))
        elif normalized_sheet == "b2b-cdnra":
            for row_number, row in rows[7:]:
                if not _text(_at(row, 3)) or not _text(_at(row, 5)):
                    continue
                output.append(_gstr_transaction(
                    path, sheet_name, row_number, config, explicit_month, period_source,
                    gstin=_at(row, 3), name=_at(row, 4), number=_at(row, 5), raw_type=_at(row, 6),
                    document_date=_at(row, 8), total=_at(row, 9), taxable=_at(row, 12),
                    igst=_at(row, 13), cgst=_at(row, 14), sgst=_at(row, 15), cess=_at(row, 16),
                    supplier_period=_at(row, 17), filing_date=_at(row, 18),
                    itc_availability=_at(row, 19), reason=_at(row, 20),
                    amendment=True, amends_document_number=_text(_at(row, 1)),
                ))
        elif normalized_sheet == "b2ba(rejected)":
            for row_number, row in rows[7:]:
                if not _text(_at(row, 2)) or not _text(_at(row, 4)):
                    continue
                output.append(_gstr_transaction(
                    path, sheet_name, row_number, config, explicit_month, period_source,
                    gstin=_at(row, 2), name=_at(row, 3), number=_at(row, 4), raw_type="invoice",
                    document_date=_at(row, 6), total=_at(row, 7), taxable=_at(row, 10),
                    igst=_at(row, 11), cgst=_at(row, 12), sgst=_at(row, 13), cess=_at(row, 14),
                    supplier_period=_at(row, 15), filing_date=_at(row, 16),
                    itc_availability="No", reason="Amendment rejected in GSTR-2B/IMS section.",
                    amendment=True, amends_document_number=_text(_at(row, 0)), force_ineligible=True,
                ))
        elif normalized_sheet == "b2b-dnra":
            for row_number, row in rows[7:]:
                if not _text(_at(row, 3)) or not _text(_at(row, 5)):
                    continue
                output.append(_gstr_transaction(
                    path, sheet_name, row_number, config, explicit_month, period_source,
                    gstin=_at(row, 3), name=_at(row, 4), number=_at(row, 5), raw_type=_at(row, 6) or "D",
                    document_date=_at(row, 8), total=_at(row, 9), taxable=_at(row, 12),
                    igst=_at(row, 13), cgst=_at(row, 14), sgst=_at(row, 15), cess=_at(row, 16),
                    supplier_period=_at(row, 17), filing_date=_at(row, 18),
                    itc_availability=_at(row, 19), reason=_at(row, 20),
                    amendment=True, amends_document_number=_text(_at(row, 1)),
                ))
        elif normalized_sheet == "b2b(rejected)":
            for row_number, row in rows[6:]:
                if not _text(_at(row, 0)) or not _text(_at(row, 2)):
                    continue
                output.append(_gstr_transaction(
                    path, sheet_name, row_number, config, explicit_month, period_source,
                    gstin=_at(row, 0), name=_at(row, 1), number=_at(row, 2), raw_type="invoice",
                    document_date=_at(row, 4), total=_at(row, 5), taxable=_at(row, 8),
                    igst=_at(row, 9), cgst=_at(row, 10), sgst=_at(row, 11), cess=_at(row, 12),
                    supplier_period=_at(row, 13), filing_date=_at(row, 14),
                    itc_availability="No", reason="Rejected in GSTR-2B/IMS section.",
                    force_ineligible=True,
                ))
        elif normalized_sheet == "b2b-cdnr(rejected)":
            for row_number, row in rows[5:]:
                if not _text(_at(row, 0)) or not _text(_at(row, 2)):
                    continue
                output.append(_gstr_transaction(
                    path, sheet_name, row_number, config, explicit_month, period_source,
                    gstin=_at(row, 0), name=_at(row, 1), number=_at(row, 2), raw_type=_at(row, 3),
                    document_date=_at(row, 5), total=_at(row, 6), taxable=_at(row, 8),
                    igst=_at(row, 9), cgst=_at(row, 10), sgst=_at(row, 11), cess=_at(row, 12),
                    supplier_period=_at(row, 13), filing_date=_at(row, 14),
                    itc_availability="No", reason="Rejected in GSTR-2B/IMS section.",
                    force_ineligible=True,
                ))
        elif normalized_sheet == "b2b-cdnra(rejected)":
            for row_number, row in rows[6:]:
                if not _text(_at(row, 3)) or not _text(_at(row, 5)):
                    continue
                output.append(_gstr_transaction(
                    path, sheet_name, row_number, config, explicit_month, period_source,
                    gstin=_at(row, 3), name=_at(row, 4), number=_at(row, 5), raw_type=_at(row, 6),
                    document_date=_at(row, 8), total=_at(row, 9), taxable=_at(row, 11),
                    igst=_at(row, 12), cgst=_at(row, 13), sgst=_at(row, 14), cess=_at(row, 15),
                    supplier_period=_at(row, 16), filing_date=_at(row, 17),
                    itc_availability="No", reason="Rejected in GSTR-2B/IMS section.",
                    amendment=True, amends_document_number=_text(_at(row, 1)), force_ineligible=True,
                ))
    return output


def _gstr_transaction(
    path: Path,
    sheet_name: str,
    row_number: int,
    config: ReconciliationConfig,
    explicit_month: str | None,
    period_source: str,
    *,
    gstin: Any,
    name: Any,
    number: Any,
    raw_type: Any,
    document_date: Any,
    total: Any,
    taxable: Any,
    igst: Any,
    cgst: Any,
    sgst: Any,
    cess: Any,
    supplier_period: Any,
    filing_date: Any,
    itc_availability: Any,
    reason: Any = "",
    amendment: bool = False,
    amends_document_number: str = "",
    force_ineligible: bool = False,
) -> Transaction:
    document_date_str = _date_string(document_date)
    report_month = parse_month(supplier_period)
    filing_month = parse_month(_date_string(filing_date) or filing_date)
    if explicit_month:
        available_month = explicit_month
        availability_basis = "explicit-report-period"
    elif period_source == "filing-month":
        available_month = filing_month or report_month
        availability_basis = "supplier-filing-date-month" if filing_month else "supplier-period-fallback"
    elif period_source == "supplier-period":
        available_month = report_month
        availability_basis = "supplier-report-period"
    else:
        available_month = None
        availability_basis = "not-assigned"
    availability = _text(itc_availability)
    eligibility: bool | None
    if force_ineligible or availability.strip().lower() in {"no", "not available", "rejected", "r"}:
        eligibility = False
    elif availability.strip().lower() in {"yes", "available", "eligible"}:
        eligibility = True
    else:
        eligibility = None
    raw_doc_type = _text(raw_type) or "invoice"
    canonical_type = normalize_gstr_doc_type(raw_doc_type, config)
    return Transaction(
        transaction_id=f"gstr2b:{path.name}:{_file_token(path)}:{sheet_name}:{row_number}",
        source="gstr2b",
        supplier_gstin=_text(gstin),
        supplier_name=_text(name),
        document_number=_text(number),
        document_type=canonical_type,
        raw_document_type=raw_doc_type,
        document_date=document_date_str,
        document_month=parse_month(document_date_str) if document_date_str else None,
        gstr2b_month=available_month,
        supplier_report_month=report_month,
        taxable_value=_amount(taxable),
        cgst=_amount(cgst),
        sgst=_amount(sgst),
        igst=_amount(igst),
        cess=_amount(cess),
        total_value=_amount(total),
        itc_eligible=eligibility,
        itc_availability=availability or None,
        is_amendment=amendment,
        amends_document_number=amends_document_number,
        source_file=path.name,
        source_sheet=sheet_name,
        source_row=row_number,
        metadata={
            "supplier_report_month": report_month,
            "filing_date": _date_string(filing_date) or _text(filing_date),
            "gstr2b_month_basis": availability_basis,
            "gstr2b_reason": _text(reason),
        },
    )


def normalize_gstr_doc_type(raw_type: Any, config: ReconciliationConfig) -> str:
    key = _text(raw_type).strip().lower()
    aliases = config.document_type_aliases.get("gstr2b", {})
    if key in aliases:
        return aliases[key]
    if key.startswith("c"):
        return "credit_note"
    if key.startswith("d"):
        return "debit_note"
    return "invoice"


def _at(row: list[Any], index: int) -> Any:
    return row[index] if 0 <= index < len(row) else ""


def _text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, Decimal):
        return format(value, "f")
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return str(value)
    return str(value).strip()


def _amount(value: Any) -> Decimal:
    if value is None or value == "":
        return _ZERO
    if isinstance(value, Decimal):
        return value
    text = _text(value).replace(",", "").replace("₹", "").replace("Rs.", "").strip()
    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1]
    if text in {"-", "—", "–"}:
        return _ZERO
    try:
        amount = Decimal(text)
        if not amount.is_finite():
            return _ZERO
        return amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except InvalidOperation:
        return _ZERO


def _date_string(value: Any) -> str | None:
    parsed: date | None
    if isinstance(value, datetime):
        parsed = value.date()
    elif isinstance(value, date):
        parsed = value
    else:
        parsed = parse_date(value)
        if parsed is None and isinstance(value, (float, int, Decimal)):
            parsed = excel_serial_to_date(value)
    return parsed.isoformat() if parsed else None
