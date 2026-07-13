from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import pandas as pd


@dataclass(frozen=True)
class UnitConversionRule:
    from_unit: str
    to_unit: str
    factor: float


UNIT_CONVERSION_RULES = [
    UnitConversionRule(from_unit="MWh", to_unit="kWh", factor=1000),
]


def apply_unit_conversions(
    df: pd.DataFrame,
    amount_column: str = "Amount of energy consumed",
    unit_column: str = "Energy Unit",
    original_unit_column: str = "Original Energy Unit",
    conversion_rules: list[UnitConversionRule] | None = None,
) -> pd.DataFrame:
    result = df.copy()
    if amount_column not in result.columns or unit_column not in result.columns:
        return result

    result[original_unit_column] = result[unit_column]

    rules_by_unit = {
        _normalize_unit(rule.from_unit): rule
        for rule in (conversion_rules or UNIT_CONVERSION_RULES)
    }

    for index, row in result.iterrows():
        rule = rules_by_unit.get(_normalize_unit(row.get(unit_column)))
        if rule is None:
            continue

        amount = _coerce_number(row.get(amount_column))
        if amount is None:
            continue

        converted_amount = amount * rule.factor
        result.at[index, amount_column] = _clean_number(converted_amount)
        result.at[index, unit_column] = rule.to_unit

    return _move_column_after(result, original_unit_column, "Total amount of energy consumed")


def _normalize_unit(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip().casefold()


def _coerce_number(value: Any) -> float | None:
    if value is None or isinstance(value, bool) or pd.isna(value):
        return None
    if isinstance(value, int | float):
        return float(value)

    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _clean_number(value: float) -> float | int:
    if float(value).is_integer():
        return int(value)
    return value


def _move_column_after(df: pd.DataFrame, column: str, after_column: str) -> pd.DataFrame:
    if column not in df.columns or after_column not in df.columns:
        return df
    columns = [item for item in df.columns if item != column]
    insert_at = columns.index(after_column) + 1
    columns[insert_at:insert_at] = [column]
    return df.loc[:, columns]
