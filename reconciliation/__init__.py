"""GSTR-2B vs Books Reconciliation Module.

Purpose: reconcile purchase / ITC data recorded in the Books against
the corresponding data appearing in GSTR-2B, with proper handling of
credit notes, timing differences, non-GST purchases, ineligible ITC,
and month-wise carry-forward of open items.
"""

from .models import (
    DocumentType,
    ReconciliationStatus,
    NormalizedTransaction,
    ReconciliationMatch,
    MonthlySummary,
    ReconciliationConfig,
)
from .normalizer import SignNormalizer, DefaultSignConvention
from .parser_books import BooksParser
from .parser_gstr2b import GSTR2BParser
from .matcher import ReconciliationMatcher
from .engine import ReconciliationEngine
from .excel_report import ExcelReportBuilder

__all__ = [
    "DocumentType",
    "ReconciliationStatus",
    "NormalizedTransaction",
    "ReconciliationMatch",
    "MonthlySummary",
    "ReconciliationConfig",
    "SignNormalizer",
    "DefaultSignConvention",
    "BooksParser",
    "GSTR2BParser",
    "ReconciliationMatcher",
    "ReconciliationEngine",
    "ExcelReportBuilder",
]
