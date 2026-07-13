from __future__ import annotations

from pathlib import Path
from typing import Any
import re

import pandas as pd

from src.lookup.mapping_lookup import (
    build_lookup_key,
    load_lookup_config,
    load_lookup_table,
    normalize_lookup_value,
)


ACCOUNT_MAPPING_COLUMNS = [
    "division",
    "legal_entity",
    "unit_name",
    "supplier_name",
    "division_shorthand",
    "facility_type",
    "scope",
    "facility_identifier",
    "activity_group",
    "invoice_count",
    "invoice_frequency",
    "consumption_unit",
    "decimal_separator",
    "currency",
]
OCR_CONTEXT_COLUMNS = {"legal_entity", "unit_name"}


def add_account_mapping_columns(
    df: pd.DataFrame,
    config_dir: str | Path = "config",
    lookup_name: str = "mapping",
    return_mapped_fields: bool = False,
) -> pd.DataFrame | tuple[pd.DataFrame, dict[str, pd.Series]]:
    config = load_lookup_config(lookup_name, config_dir=config_dir)
    table = load_lookup_table(config)
    _require_mapping_columns(table)

    lookup_by_account = _lookup_by_account_number(table)
    lookup_by_unit_alias = _lookup_by_alias(table, "ocr_unit_name_list")
    lookup_by_legal_entity_alias = _lookup_by_alias(table, "ocr_legal_entity_name_list")
    result = df.copy()
    mapped_fields = {
        column: pd.Series(False, index=result.index)
        for column in ACCOUNT_MAPPING_COLUMNS
    }

    for column in ACCOUNT_MAPPING_COLUMNS:
        if column in OCR_CONTEXT_COLUMNS and column in result.columns:
            result[column] = result[column].astype(object)
        else:
            result[column] = pd.Series([""] * len(result), index=result.index, dtype=object)

    for index, row in result.iterrows():
        mapping_row = _find_mapping_row(
            row.to_dict(),
            lookup_by_account=lookup_by_account,
            lookup_by_unit_alias=lookup_by_unit_alias,
            lookup_by_legal_entity_alias=lookup_by_legal_entity_alias,
        )
        if mapping_row is None:
            continue
        for column in ACCOUNT_MAPPING_COLUMNS:
            value = mapping_row.get(column)
            if not _is_blank(value):
                result.at[index, column] = value
                mapped_fields[column].at[index] = True

    ordered_result = _order_mapping_columns_after_account_number(result)
    if return_mapped_fields:
        return ordered_result, mapped_fields
    return ordered_result


def _find_mapping_row(
    row: dict[str, Any],
    lookup_by_account: dict[str, dict[str, Any]],
    lookup_by_unit_alias: dict[str, dict[str, Any]],
    lookup_by_legal_entity_alias: dict[str, dict[str, Any]],
) -> dict[str, Any] | None:
    account_match = lookup_by_account.get(normalize_lookup_value(row.get("account_number")))
    if account_match is not None:
        return account_match

    unit_match = lookup_by_unit_alias.get(normalize_lookup_value(row.get("unit_name")))
    if unit_match is not None:
        return unit_match

    return lookup_by_legal_entity_alias.get(normalize_lookup_value(row.get("legal_entity")))


def _lookup_by_account_number(table: pd.DataFrame) -> dict[str, dict[str, Any]]:
    lookup: dict[str, dict[str, Any]] = {}
    for _, row in table.iterrows():
        record = row.drop(labels=["_python_lookup_key"], errors="ignore").to_dict()
        key = build_lookup_key(record, ("account_number",))
        if key and key not in lookup:
            lookup[key] = record
    return lookup


def _lookup_by_alias(table: pd.DataFrame, alias_column: str) -> dict[str, dict[str, Any]]:
    if alias_column not in table.columns:
        return {}

    lookup: dict[str, dict[str, Any]] = {}
    for _, row in table.iterrows():
        record = row.drop(labels=["_python_lookup_key"], errors="ignore").to_dict()
        for alias in _split_aliases(record.get(alias_column)):
            key = normalize_lookup_value(alias)
            if key and key not in lookup:
                lookup[key] = record
    return lookup


def _split_aliases(value: Any) -> list[str]:
    if _is_blank(value):
        return []
    return [
        part.strip()
        for part in re.split(r"[;,\n\r|]+", str(value))
        if part.strip()
    ]


def _with_empty_mapping_columns(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    for column in ACCOUNT_MAPPING_COLUMNS:
        if column not in result.columns:
            result[column] = pd.Series([""] * len(result), index=result.index, dtype=object)
        else:
            result[column] = result[column].astype(object)
    return result


def _require_mapping_columns(table: pd.DataFrame) -> None:
    missing_columns = [
        column
        for column in ["account_number", *ACCOUNT_MAPPING_COLUMNS]
        if column not in table.columns
    ]
    if missing_columns:
        raise KeyError(f"Missing account mapping columns: {missing_columns}")


def _order_mapping_columns_after_account_number(df: pd.DataFrame) -> pd.DataFrame:
    columns = [
        column
        for column in df.columns
        if column not in ACCOUNT_MAPPING_COLUMNS and column != "account_number_confidence"
    ]
    insert_at = columns.index("account_number") + 1 if "account_number" in columns else len(columns)
    account_confidence_columns = ["account_number_confidence"] if "account_number_confidence" in df.columns else []
    ordered_columns = [
        *columns[:insert_at],
        *account_confidence_columns,
        *columns[insert_at:],
        *ACCOUNT_MAPPING_COLUMNS,
    ]
    return df.loc[:, ordered_columns]


def _is_blank(value: Any) -> bool:
    return value is None or pd.isna(value) or str(value).strip() == ""
