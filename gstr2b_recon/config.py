"""Configuration for matching, sign conventions and source-specific aliases."""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from typing import Any


@dataclass(slots=True)
class ReconciliationConfig:
    exact_tolerance: Decimal = Decimal("0.01")
    value_tolerance: Decimal = Decimal("1.00")
    max_future_months: int = 3
    max_past_months: int = 3
    expected_lag_months: int = 1
    # This is a transparent fallback for the supplied register, which has no
    # explicit ITC-eligibility column. Change either policy to "eligible",
    # "non_gst" or "ineligible_itc" to match the company's accounting rules.
    zero_tax_with_gstin: str = "ineligible_itc"
    zero_tax_without_gstin: str = "non_gst"
    document_type_aliases: dict[str, dict[str, str]] = field(default_factory=lambda: {
        "books": {
            "invoice": "invoice",
            "purchase": "invoice",
            "debit note": "credit_note",
            "credit note": "debit_note",
            "credit_note": "credit_note",
            "debit_note": "debit_note",
        },
        "gstr2b": {
            "invoice": "invoice",
            "regular": "invoice",
            "c": "credit_note",
            "credit note": "credit_note",
            "credit_note": "credit_note",
            "d": "debit_note",
            "debit note": "debit_note",
            "debit_note": "debit_note",
        },
    })
    # The values are the sign applied to absolute credit/debit note values.
    # Invoice signs are retained as imported unless explicitly configured.
    sign_conventions: dict[str, dict[str, int]] = field(default_factory=lambda: {
        "books": {"invoice": 1, "credit_note": -1, "debit_note": 1},
        "gstr2b": {"invoice": 1, "credit_note": -1, "debit_note": 1},
    })
    non_gst_account_keywords: list[str] = field(default_factory=lambda: ["non gst", "non-gst", "without gst"])
    ineligible_account_keywords: list[str] = field(default_factory=lambda: [
        "ineligible itc", "blocked itc", "ineligible credit", "blocked credit",
    ])

    def __post_init__(self) -> None:
        self.exact_tolerance = _decimal(self.exact_tolerance)
        self.value_tolerance = _decimal(self.value_tolerance)
        if self.exact_tolerance < 0 or self.value_tolerance < 0:
            raise ValueError("Amount tolerances must be non-negative")
        valid_policies = {"eligible", "non_gst", "ineligible", "ineligible_itc"}
        if self.zero_tax_with_gstin not in valid_policies or self.zero_tax_without_gstin not in valid_policies:
            raise ValueError(f"Zero-tax policy must be one of {sorted(valid_policies)}")
        self.max_future_months = max(0, int(self.max_future_months))
        self.max_past_months = max(0, int(self.max_past_months))
        self.expected_lag_months = max(0, int(self.expected_lag_months))
        self.zero_tax_with_gstin = str(self.zero_tax_with_gstin).lower()
        self.zero_tax_without_gstin = str(self.zero_tax_without_gstin).lower()

    @classmethod
    def from_dict(cls, data: dict[str, Any] | None) -> ReconciliationConfig:
        if not data:
            return cls()
        allowed = {
            "exact_tolerance", "value_tolerance", "max_future_months", "max_past_months",
            "expected_lag_months", "zero_tax_with_gstin", "zero_tax_without_gstin",
            "document_type_aliases", "sign_conventions", "non_gst_account_keywords",
            "ineligible_account_keywords",
        }
        values = {key: value for key, value in data.items() if key in allowed}
        defaults = cls()
        for nested_key in ("document_type_aliases", "sign_conventions"):
            if nested_key in values:
                merged = {source: dict(rules) for source, rules in getattr(defaults, nested_key).items()}
                for source, rules in values[nested_key].items():
                    merged.setdefault(source, {}).update(rules)
                values[nested_key] = merged
        return cls(**values)

    @classmethod
    def load(cls, path: str | Path | None) -> ReconciliationConfig:
        if path is None:
            return cls()
        with Path(path).open(encoding="utf-8") as stream:
            data = json.load(stream)
        if not isinstance(data, dict):
            raise TypeError("Configuration JSON must contain an object")
        return cls.from_dict(data)

    def to_dict(self) -> dict[str, Any]:
        return {
            "exact_tolerance": str(self.exact_tolerance),
            "value_tolerance": str(self.value_tolerance),
            "max_future_months": self.max_future_months,
            "max_past_months": self.max_past_months,
            "expected_lag_months": self.expected_lag_months,
            "zero_tax_with_gstin": self.zero_tax_with_gstin,
            "zero_tax_without_gstin": self.zero_tax_without_gstin,
            "document_type_aliases": self.document_type_aliases,
            "sign_conventions": self.sign_conventions,
            "non_gst_account_keywords": self.non_gst_account_keywords,
            "ineligible_account_keywords": self.ineligible_account_keywords,
        }


def _decimal(value: Any) -> Decimal:
    if isinstance(value, Decimal):
        return value
    return Decimal(str(value))
