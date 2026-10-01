"""Sign normalization rules for Books vs GSTR-2B.

The accounting system in this repository uses the following conventions for raw values:

Books (Tally Purchase Register / Debit Note Register):
  * Purchase invoice — CGST/SGST/IGST are recorded positive. Taxable is the sum of
    Purchase Accounts + Indirect/Direct Expenses + Fixed Assets and is positive.
  * Debit note (which is the buyer's record of a supplier credit note) —
    CGST/SGST/IGST are recorded positive in the Debit Note Register, but the
    transaction REDUCES eligible ITC. In normalized form, these values must be NEGATIVE.

GSTR-2B:
  * B2B Regular invoice — taxable_value, CGST/SGST/IGST are positive and INCREASE ITC.
  * B2B-CDNR with Note type 'C' (Credit Note) — taxable, CGST/SGST/IGST are positive
    in the Excel file but the GST portal says these "should be net off against ITC",
    so they REDUCE eligible ITC. In normalized form they must be NEGATIVE.
  * B2B-CDNR with Note type 'D' (Debit Note) — values positive and INCREASE ITC.
  * Amendments (B2BA / B2B-CDNRA) supersede original entries; normalized using the same
    rule according to the REVISED note type.
  * ITC not available / Rejected sections carry the same signs but are flagged so the
    matcher reports them separately.

The normalizer returns *new* transactions whose taxable/cgst/sgst/igst fields are
already in normalized form (positive = ITC increase, negative = ITC reduction).
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import List

from .models import NormalizedTransaction, DocumentType


@dataclass
class SignConvention:
    """Encapsulates the sign decision per source/document type."""

    def sign_for(self, source: str, doc_type: DocumentType) -> int:
        """Return +1 if raw values already match our convention, -1 if they need to be negated."""
        raise NotImplementedError


class DefaultSignConvention(SignConvention):
    """Default convention for the DTPL Tally + GSTR-2B dataset."""

    def sign_for(self, source: str, doc_type: DocumentType) -> int:
        if source == "BOOKS":
            if doc_type == DocumentType.CREDIT_NOTE:
                # Books "Debit Note Register" entries record the supplier credit note
                # using positive numbers but reducing ITC → flip sign.
                return -1
            return +1
        if source == "GSTR2B":
            if doc_type == DocumentType.CREDIT_NOTE:
                # 2B Credit Note ('C') is reported with positive numbers but reduces ITC → flip.
                return -1
            return +1
        return +1


class SignNormalizer:
    """Applies a :class:`SignConvention` to a list of transactions.

    After normalization the taxable/cgst/sgst/igst fields directly represent ITC
    impact (positive = add, negative = reduce).
    """

    def __init__(self, convention: SignConvention | None = None):
        self.convention = convention or DefaultSignConvention()

    def normalize(self, txs: List[NormalizedTransaction]) -> List[NormalizedTransaction]:
        out: List[NormalizedTransaction] = []
        for tx in txs:
            sign = self.convention.sign_for(tx.source, tx.document_type)
            # We mutate a copy so that the original parser output is preserved
            nt = NormalizedTransaction(**{**tx.__dict__})
            if sign < 0:
                nt.taxable_value = -abs(nt.taxable_value)
                nt.cgst = -abs(nt.cgst)
                nt.sgst = -abs(nt.sgst)
                nt.igst = -abs(nt.igst)
                nt.cess = -abs(nt.cess)
                nt.invoice_value = -abs(nt.invoice_value) if nt.invoice_value else 0.0
            else:
                nt.taxable_value = abs(nt.taxable_value)
                nt.cgst = abs(nt.cgst)
                nt.sgst = abs(nt.sgst)
                nt.igst = abs(nt.igst)
                nt.cess = abs(nt.cess)
                nt.invoice_value = abs(nt.invoice_value) if nt.invoice_value else 0.0
            out.append(nt)
        return out
