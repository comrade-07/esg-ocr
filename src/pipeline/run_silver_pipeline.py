from calendar import monthrange
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import pandas as pd

from src.core.config_loader import load_yaml
from src.core.logger import get_logger
from src.core.path_settings import silver_excel_output_dir
from src.lookup.mapping_lookup import (
    build_lookup_key,
    load_lookup_config,
    load_lookup_table,
    normalize_lookup_value,
)
from src.normalize.text_normalizer import clean_energy_unit
from src.output.xlsx_writer import write_xlsx
from src.pipeline.run_review_pipeline import checkpoint_dir_for_invoice_type
from src.review.approved_silver import APPROVED_SILVER_FILENAME
from src.review.manual_data_entry import MANUAL_DATA_ENTRY_DECISIONS_FILENAME
from src.transform.account_mapping import add_account_mapping_columns
from src.transform.date_fields import add_normalized_date_columns
from src.transform.quantity_fields import clean_quantity_columns
from src.transform.text_fields import clean_text_columns

logger = get_logger(__name__)


DEFAULT_INVOICE_TYPE = "scope2"
DATA_QUALITY_COLUMN = "data_quality"
OCR_DATA_QUALITY_VALUE = "Actual"
MANUAL_DATA_ENTRY_COMPLETE_STATUS = "COMPLETED"
MANUAL_DATA_ENTRY_HELPER_COLUMN = "manual_data_entry_portion"
MANUAL_DATA_ENTRY_HELPER_VALUE = "Yes"

SILVER_OUTPUT_STEPS = {
    "reviewed": "01",
    "normalized": "02",
    "curated": "03",
    "aggregated": "04",
    "proration_calculation": "05",
    "proration_split": "06",
    "prorated": "07",
    "manual_proration_calculation": "07A",
    "manual_proration_split": "07B",
    "manual_prorated": "07C",
    "business_mapping": "08",
    "template_preparation": "09",
    "template_output": "10",
}


def silver_output_filename(invoice_type: str, layer: str) -> str:
    step = SILVER_OUTPUT_STEPS[layer]
    return f"{step}_{invoice_type}_silver_{layer}.xlsx"


def legacy_silver_output_filename(invoice_type: str, layer: str) -> str:
    return f"{invoice_type}_silver_{layer}.xlsx"


SILVER_CURATED_COLUMNS = [
    DATA_QUALITY_COLUMN,
    "division",
    "legal_entity",
    "unit_name",
    "account_number",
    "supplier_name",
    "consumption_start_date_1_normalized",
    "consumption_end_date_1_normalized",
    "consumption_start_date_2_normalized",
    "consumption_end_date_2_normalized",
    "quantity_1",
    "quantity_2",
    "quantity_3",
    "quantity_4",
    "quantity_5",
    "quantity_6",
    "quantity_unit_1",
    "consumption_unit",
    "consumption_quantity_unit",
    "quantity_unit_2",
    "total_amount",
    "solar_export",
    "solar_banking_charge",
    "solar_total_generation",
    "division_shorthand",
    "facility_type",
    "scope",
    "facility_identifier",
    "activity_group",
    "invoice_count",
    "source_file",
    "sharepoint_link",
    "status",
    "createdDateTme",
    "lastUpdateDateTime",
    "apiVersion",
    "modelID",
    "missing_fields",
    "low_confidence_fields",
    "approval_status",
    "manual_review_tag",
]


SILVER_CURATED_SOURCE_ALIASES = {
    "createdDateTme": ["createdDateTme", "createdDateTime"],
    "lastUpdateDateTime": ["lastUpdateDateTime", "lastUpdatedDateTime"],
    "modelID": ["modelID", "modelId"],
}


SILVER_AGGREGATED_COLUMNS = [
    column
    for column in SILVER_CURATED_COLUMNS
    if column not in {f"quantity_{index}" for index in range(1, 7)}
]
SILVER_AGGREGATED_COLUMNS.insert(
    SILVER_AGGREGATED_COLUMNS.index("quantity_unit_1"),
    "aggregated_quantity_2",
)
SILVER_AGGREGATED_COLUMNS.insert(
    SILVER_AGGREGATED_COLUMNS.index("aggregated_quantity_2"),
    "aggregated_quantity_1",
)
MANUAL_DATA_ENTRY_FLOW_COLUMNS = [
    "manual_entry_invoice_id",
    "manual_entry_line_id",
    "manual_entry_reviewer",
    "manual_entry_comment",
    "manual_entry_data_quality",
    "manual_entry_transaction_date",
]
for column in [*MANUAL_DATA_ENTRY_FLOW_COLUMNS, MANUAL_DATA_ENTRY_HELPER_COLUMN]:
    if column not in SILVER_AGGREGATED_COLUMNS:
        SILVER_AGGREGATED_COLUMNS.append(column)


SILVER_PRORATED_HELPER_COLUMNS = [
    "proration_source_row_id",
    "average_consumption_1",
    "date_1_proration",
    "date_1_total_days",
    "prorated_quantity_1",
    "date_1_proration_month",
    "date_1_proration_start",
    "date_1_proration_end",
    "average_consumption_date_2",
    "date_2_proration",
    "date_2_total_days",
    "prorated_quantity_2",
    "date_2_proration_month",
    "date_2_proration_start",
    "date_2_proration_end",
]


SILVER_PRORATION_CALCULATION_COLUMNS = SILVER_AGGREGATED_COLUMNS.copy()
_prorated_helper_insert_at = SILVER_PRORATION_CALCULATION_COLUMNS.index("aggregated_quantity_2") + 1
SILVER_PRORATION_CALCULATION_COLUMNS[
    _prorated_helper_insert_at:_prorated_helper_insert_at
] = SILVER_PRORATED_HELPER_COLUMNS


SILVER_PRORATION_SPLIT_SOURCE_COLUMNS = [
    "consumption_start_date_1_normalized",
    "consumption_end_date_1_normalized",
    "aggregated_quantity_1",
    "average_consumption_1",
    "date_1_proration",
    "date_1_total_days",
    "prorated_quantity_1",
    "date_1_proration_month",
    "date_1_proration_start",
    "date_1_proration_end",
    "consumption_start_date_2_normalized",
    "consumption_end_date_2_normalized",
    "aggregated_quantity_2",
    "average_consumption_date_2",
    "date_2_proration",
    "date_2_total_days",
    "prorated_quantity_2",
    "date_2_proration_month",
    "date_2_proration_start",
    "date_2_proration_end",
]


SILVER_PRORATION_SPLIT_VALUE_COLUMNS = [
    "consumption_start_date_proration",
    "consumption_end_date_proration",
    "aggregate_quantity1_to_6",
    "average_consumption",
    "date_split_days_proration",
    "date_total_days_proration",
    "quantity_proration",
    "date_month_proration",
    "date_split_start_proration",
    "date_split_end_proration",
]


SILVER_PRORATION_SPLIT_COLUMNS = []
for column in SILVER_PRORATION_CALCULATION_COLUMNS:
    if column == "consumption_start_date_1_normalized":
        SILVER_PRORATION_SPLIT_COLUMNS.extend(SILVER_PRORATION_SPLIT_VALUE_COLUMNS)
    if column not in SILVER_PRORATION_SPLIT_SOURCE_COLUMNS:
        SILVER_PRORATION_SPLIT_COLUMNS.append(column)


SILVER_PRORATED_SUMMARY_COLUMNS = [
    "consumption_month",
    "month_start_date",
    "month_end_date",
    "monthly_consumption",
    "complete_month_data_captured",
    "covered_days_in_month",
    "calendar_days_in_month",
    "proration_component_count",
    "proration_source_row_ids",
    "proration_split_keys",
    "source_file_count",
    "source_files",
]


SILVER_PRORATED_COLUMNS = (
    SILVER_PRORATION_SPLIT_COLUMNS.copy()
    + SILVER_PRORATED_SUMMARY_COLUMNS
)


SILVER_BUSINESS_MAPPING_FIXED_COLUMNS = [
    "Data Quality",
    "KPI Component",
    "Division",
    "Legal Entity Name",
    "Unit",
    "Consumption start date",
    "Consumption end date",
    "Transaction Date",
    "Energy Type",
    "Purchased or acquired",
    "Amount of energy consumed",
    "Energy Unit",
    "lookup_key",
    "fossil_fuel_%",
    "renewable_energy_%",
    "nuclear_%",
    "contractual_instruments",
    "Account number",
    "Supplier Name",
    "solar_export",
    "solar_banking_charge",
    "solar_total_generation",
    "total_solar_consumed",
    "proration_component_count",
    "proration_source_row_ids",
    "proration_split_keys",
    "source_file_count",
    "source_files",
    "covered_days_in_month",
    "calendar_days_in_month",
    "complete_month_data_captured",
    "status",
    "createdDateTme",
    "lastUpdateDateTime",
    "apiVersion",
    "modelID",
    "missing_fields",
    "low_confidence_fields",
    "approval_status",
    "manual_review_tag",
    MANUAL_DATA_ENTRY_HELPER_COLUMN,
]


SILVER_BUSINESS_MAPPING_MAX_SHAREPOINT_LINKS = 4
SILVER_BUSINESS_MAPPING_MIN_SHAREPOINT_LINKS = 2


SILVER_TEMPLATE_PREPARATION_COLUMNS = [
    "Data Quality",
    "KPI Component",
    "Division",
    "Legal Entity Name",
    "Unit",
    "Consumption start date",
    "Consumption end date",
    "Transaction Date",
    "Energy KPI",
    "Energy Type",
    "Purchased or acquired",
    "Amount of energy consumed",
    "Energy Unit",
    "contractual_instruments",
    "Total amount of energy consumed",
    "fossil_fuel_%",
    "renewable_energy_%",
    "nuclear_%",
    "Account number",
    "Supplier Name",
    "lookup_key",
    "solar_export",
    "solar_banking_charge",
    "solar_total_generation",
    "total_solar_consumed",
    "proration_component_count",
    "proration_source_row_ids",
    "proration_split_keys",
    "source_file_count",
    "source_files",
    "covered_days_in_month",
    "calendar_days_in_month",
    "complete_month_data_captured",
    "status",
    "lastUpdateDateTime",
    "apiVersion",
    "modelID",
    "missing_fields",
    "low_confidence_fields",
    "approval_status",
    "manual_review_tag",
    "sharepoint_link_1",
    "sharepoint_link_2",
    "sharepoint_link_3",
    "sharepoint_link_4",
    MANUAL_DATA_ENTRY_HELPER_COLUMN,
]


SILVER_TEMPLATE_PREPARATION_KPI_COLUMNS = [
    ("fossil_fuel_%", "Fossil Fuels", "percentage"),
    ("renewable_energy_%", "Renewable sources", "percentage"),
    ("nuclear_%", "Nuclear sources", "percentage"),
    ("solar_total_generation", "Renewable energy production", "value"),
    (
        "total_solar_consumed",
        "Consumption of self-generated non-fuel renewable energy",
        "value",
    ),
]


def run_silver_pipeline(
    config_dir: str | Path = "config",
    bronze_file: str | Path | None = None,
    approved_file: str | Path | None = None,
    invoice_type: str | None = None,
) -> Path:
    config_dir = Path(config_dir)
    settings = load_yaml(config_dir / "settings.yaml")
    invoice_type = invoice_type or settings.get("pipeline", {}).get("invoice_type", DEFAULT_INVOICE_TYPE)

    silver_dir = silver_excel_output_dir(settings)
    checkpoint_dir = checkpoint_dir_for_invoice_type(settings, invoice_type)
    input_file = _silver_input_file(
        checkpoint_dir=checkpoint_dir,
        bronze_file=bronze_file,
        approved_file=approved_file,
    )
    reviewed_file = silver_output_filename(invoice_type, "reviewed")
    normalized_file = silver_output_filename(invoice_type, "normalized")
    curated_file = silver_output_filename(invoice_type, "curated")
    aggregated_file = silver_output_filename(invoice_type, "aggregated")
    proration_calculation_file = silver_output_filename(invoice_type, "proration_calculation")
    proration_split_file = silver_output_filename(invoice_type, "proration_split")
    prorated_file = silver_output_filename(invoice_type, "prorated")
    manual_proration_calculation_file = silver_output_filename(invoice_type, "manual_proration_calculation")
    manual_proration_split_file = silver_output_filename(invoice_type, "manual_proration_split")
    manual_prorated_file = silver_output_filename(invoice_type, "manual_prorated")
    business_mapping_file = silver_output_filename(invoice_type, "business_mapping")
    template_preparation_file = silver_output_filename(invoice_type, "template_preparation")
    template_output_file = silver_output_filename(invoice_type, "template_output")

    if not input_file.exists():
        raise FileNotFoundError(f"Approved review CSV does not exist: {input_file}")

    _remove_legacy_silver_outputs(silver_dir, invoice_type)

    logger.info("Reading approved review CSV for silver layers from %s", input_file)
    reviewed_df = pd.read_csv(input_file, dtype=object, keep_default_na=False)
    reviewed_df = _with_data_quality_column(reviewed_df)
    reviewed_path = write_xlsx(reviewed_df, silver_dir, reviewed_file, sheet_name="SilverReviewed")
    logger.info("Wrote silver-reviewed Excel workbook to %s", reviewed_path)

    normalized_df, mapped_fields = add_account_mapping_columns(
        reviewed_df,
        config_dir=config_dir,
        return_mapped_fields=True,
    )
    normalized_df = clean_quantity_columns(normalized_df)
    normalized_df = add_normalized_date_columns(normalized_df)
    normalized_df = clean_text_columns(normalized_df, skip_field_masks=mapped_fields)
    normalized_df = add_consumption_quantity_unit(normalized_df)
    normalized_df = _with_data_quality_column(normalized_df)

    silver_path = write_xlsx(normalized_df, silver_dir, normalized_file, sheet_name="SilverNormalized")
    logger.info("Wrote silver-normalized Excel workbook to %s", silver_path)

    curated_df = build_silver_curated(normalized_df)
    curated_path = write_xlsx(curated_df, silver_dir, curated_file, sheet_name="SilverCurated")
    logger.info("Wrote silver-curated Excel workbook to %s", curated_path)

    aggregated_df = build_silver_aggregated(curated_df)
    aggregated_path = write_xlsx(
        aggregated_df,
        silver_dir,
        aggregated_file,
        sheet_name="SilverAggregated",
    )
    logger.info("Wrote silver-aggregated Excel workbook to %s", aggregated_path)

    proration_calculation_df = build_silver_proration_calculation(aggregated_df)
    proration_calculation_path = write_xlsx(
        proration_calculation_df,
        silver_dir,
        proration_calculation_file,
        sheet_name="SilverProrationCalculation",
    )
    logger.info(
        "Wrote silver-proration-calculation Excel workbook to %s",
        proration_calculation_path,
    )

    proration_split_df = build_silver_proration_split(proration_calculation_df)
    proration_split_path = write_xlsx(
        proration_split_df,
        silver_dir,
        proration_split_file,
        sheet_name="SilverProrationSplit",
    )
    logger.info(
        "Wrote silver-proration-split Excel workbook to %s",
        proration_split_path,
    )

    prorated_df = build_silver_prorated(proration_split_df)
    prorated_path = write_xlsx(
        prorated_df,
        silver_dir,
        prorated_file,
        sheet_name="SilverProrated",
    )
    logger.info("Wrote silver-prorated Excel workbook to %s", prorated_path)

    # Manual entries skip OCR review, but they still need the same monthly split
    # before joining OCR rows for downstream business mapping.
    manual_decisions_df = load_completed_manual_data_entry_decisions(
        checkpoint_dir / MANUAL_DATA_ENTRY_DECISIONS_FILENAME
    )
    manual_proration_input_df = build_manual_data_entry_proration_input(
        manual_decisions_df,
        config_dir=config_dir,
    )
    manual_proration_calculation_df = build_silver_proration_calculation(manual_proration_input_df)
    manual_proration_calculation_path = write_xlsx(
        manual_proration_calculation_df,
        silver_dir,
        manual_proration_calculation_file,
        sheet_name="ManualProrationCalculation",
    )
    logger.info(
        "Wrote manual-data-entry-proration-calculation Excel workbook to %s",
        manual_proration_calculation_path,
    )

    manual_proration_split_df = build_silver_proration_split(manual_proration_calculation_df)
    manual_proration_split_path = write_xlsx(
        manual_proration_split_df,
        silver_dir,
        manual_proration_split_file,
        sheet_name="ManualProrationSplit",
    )
    logger.info(
        "Wrote manual-data-entry-proration-split Excel workbook to %s",
        manual_proration_split_path,
    )

    manual_prorated_df = build_silver_prorated(manual_proration_split_df)
    manual_prorated_path = write_xlsx(
        manual_prorated_df,
        silver_dir,
        manual_prorated_file,
        sheet_name="ManualProrated",
    )
    logger.info("Wrote manual-data-entry-prorated Excel workbook to %s", manual_prorated_path)

    business_mapping_input_df = pd.concat(
        [prorated_df, manual_prorated_df],
        ignore_index=True,
        sort=False,
    )
    business_mapping_input_df = _reindex_silver_prorated_columns(business_mapping_input_df)
    business_mapping_df = build_silver_business_mapping(business_mapping_input_df, config_dir=config_dir)
    business_mapping_path = write_xlsx(
        business_mapping_df,
        silver_dir,
        business_mapping_file,
        sheet_name="SilverBusinessMapping",
    )
    logger.info("Wrote silver-business-mapping Excel workbook to %s", business_mapping_path)

    template_preparation_df = build_silver_template_preparation(
        business_mapping_df,
    )
    template_preparation_path = write_xlsx(
        template_preparation_df,
        silver_dir,
        template_preparation_file,
        sheet_name="SilverTemplatePreparation",
    )
    logger.info("Wrote silver-template-preparation Excel workbook to %s", template_preparation_path)

    template_output_df = build_silver_template_output(template_preparation_df)
    template_output_path = write_xlsx(
        template_output_df,
        silver_dir,
        template_output_file,
        sheet_name="SilverTemplateOutput",
    )
    logger.info("Wrote silver-template-output Excel workbook to %s", template_output_path)

    return silver_path


def build_silver_curated(normalized_df: pd.DataFrame) -> pd.DataFrame:
    curated_columns = {}
    for column in SILVER_CURATED_COLUMNS:
        source_column = _first_existing_column(
            normalized_df,
            SILVER_CURATED_SOURCE_ALIASES.get(column, [column]),
        )
        if source_column is None and column == DATA_QUALITY_COLUMN:
            curated_columns[column] = OCR_DATA_QUALITY_VALUE
        elif source_column is None:
            curated_columns[column] = ""
        else:
            curated_columns[column] = normalized_df[source_column]
    return pd.DataFrame(curated_columns, index=normalized_df.index)


def _with_data_quality_column(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    if DATA_QUALITY_COLUMN in result.columns:
        data_quality = result[DATA_QUALITY_COLUMN].apply(
            lambda value: OCR_DATA_QUALITY_VALUE if _is_blank(value) else value
        )
        result = result.drop(columns=[DATA_QUALITY_COLUMN])
    else:
        data_quality = pd.Series([OCR_DATA_QUALITY_VALUE] * len(result), index=result.index)

    insert_at = (
        result.columns.get_loc("document_confidence") + 1
        if "document_confidence" in result.columns
        else len(result.columns)
    )
    result.insert(insert_at, DATA_QUALITY_COLUMN, data_quality)
    return result


def add_consumption_quantity_unit(normalized_df: pd.DataFrame) -> pd.DataFrame:
    result = normalized_df.copy()
    if "consumption_unit" not in result.columns:
        result["consumption_unit"] = ""
    if "quantity_unit_1" not in result.columns:
        result["quantity_unit_1"] = ""

    consumption_quantity_unit = result.apply(
        lambda row: _effective_consumption_quantity_unit(
            quantity_unit=row.get("quantity_unit_1"),
            consumption_unit=row.get("consumption_unit"),
        ),
        axis=1,
    )
    result = result.drop(columns=["consumption_quantity_unit"], errors="ignore")
    insert_at = result.columns.get_loc("consumption_unit") + 1
    result.insert(insert_at, "consumption_quantity_unit", consumption_quantity_unit)
    return result


def build_silver_aggregated(curated_df: pd.DataFrame) -> pd.DataFrame:
    aggregated_df = curated_df.drop(
        columns=[f"quantity_{index}" for index in range(1, 7)],
        errors="ignore",
    ).copy()
    aggregated_quantities = curated_df.apply(_aggregate_quantities, axis=1, result_type="expand")
    aggregated_quantities.columns = ["aggregated_quantity_1", "aggregated_quantity_2"]
    insert_at = aggregated_df.columns.get_loc("quantity_unit_1")
    aggregated_df.insert(insert_at, "aggregated_quantity_1", aggregated_quantities["aggregated_quantity_1"])
    aggregated_df.insert(insert_at + 1, "aggregated_quantity_2", aggregated_quantities["aggregated_quantity_2"])
    return aggregated_df.reindex(columns=SILVER_AGGREGATED_COLUMNS, fill_value="")


def build_silver_proration_calculation(aggregated_df: pd.DataFrame) -> pd.DataFrame:
    prorated_rows = []

    for source_row_id, (_, row) in enumerate(aggregated_df.iterrows(), start=1):
        row = row.copy()
        row["proration_source_row_id"] = source_row_id
        expanded_rows = [row.copy()]
        for period_index in (1, 2):
            next_rows = []
            for expanded_row in expanded_rows:
                next_rows.extend(_prorate_row_period(expanded_row, period_index))
            expanded_rows = next_rows
        prorated_rows.extend(expanded_rows)

    return (
        pd.DataFrame(prorated_rows)
        .reindex(columns=SILVER_PRORATION_CALCULATION_COLUMNS, fill_value="")
        .reset_index(drop=True)
    )


def build_silver_proration_split(proration_calculation_df: pd.DataFrame) -> pd.DataFrame:
    split_rows = []
    seen_segments = set()

    for _, row in proration_calculation_df.iterrows():
        for period_index in (1, 2):
            split_row = _split_proration_period_row(row, period_index)
            if split_row is None:
                continue

            segment_key = (
                split_row.get("proration_source_row_id"),
                period_index,
                split_row.get("date_month_proration"),
                split_row.get("date_split_start_proration"),
                split_row.get("date_split_end_proration"),
                split_row.get("quantity_proration"),
            )
            if segment_key in seen_segments:
                continue
            seen_segments.add(segment_key)
            split_rows.append(split_row)

    split_df = pd.DataFrame(split_rows).reindex(columns=SILVER_PRORATION_SPLIT_COLUMNS, fill_value="")
    if not split_df.empty:
        split_df = split_df.sort_values(
            by=["proration_source_row_id", "date_split_start_proration", "date_split_end_proration"],
            kind="stable",
        )
    return split_df.reset_index(drop=True)


def build_silver_prorated(proration_split_df: pd.DataFrame) -> pd.DataFrame:
    monthly_records = []

    for _, row in proration_split_df.iterrows():
        quantity = _coerce_number(row.get("quantity_proration"))
        segment_month = row.get("date_month_proration")
        segment_start = _coerce_date(row.get("date_split_start_proration"))
        segment_end = _coerce_date(row.get("date_split_end_proration"))
        if quantity is None or _is_blank(segment_month) or segment_start is None or segment_end is None:
            continue

        record = row.to_dict()
        record.update({
            "consumption_month": str(segment_month),
            "month_start_date": date(segment_start.year, segment_start.month, 1),
            "month_end_date": date(
                segment_start.year,
                segment_start.month,
                monthrange(segment_start.year, segment_start.month)[1],
            ),
            "monthly_consumption": quantity,
            "segment_start": segment_start,
            "segment_end": segment_end,
        })
        monthly_records.append(record)

    if not monthly_records:
        return pd.DataFrame(columns=SILVER_PRORATED_COLUMNS)

    records_df = pd.DataFrame(monthly_records)
    group_columns = [
        "account_number",
        "unit_name",
        "consumption_month",
    ]
    grouped_rows = []
    for group_values, group_df in records_df.groupby(group_columns, dropna=False, sort=False):
        account_number, unit_name, consumption_month = group_values
        month_start = group_df["month_start_date"].iloc[0]
        month_end = group_df["month_end_date"].iloc[0]
        covered_days = _covered_days(group_df[["segment_start", "segment_end"]].itertuples(index=False, name=None))
        calendar_days = month_end.day
        source_files = sorted({
            str(value)
            for value in group_df["source_file"].tolist()
            if not _is_blank(value)
        })
        sharepoint_links = _unique_text_values(group_df["sharepoint_link"])
        proration_source_row_ids = _unique_text_values(group_df["proration_source_row_id"])
        proration_split_keys = [
            "|".join([
                str(row.proration_source_row_id),
                str(row.date_month_proration),
                str(row.date_split_start_proration),
                str(row.date_split_end_proration),
            ])
            for row in group_df[
                [
                    "proration_source_row_id",
                    "date_month_proration",
                    "date_split_start_proration",
                    "date_split_end_proration",
                ]
            ].itertuples(index=False)
        ]
        grouped_row = {
            "account_number": account_number,
            "unit_name": unit_name,
            "consumption_month": consumption_month,
            "month_start_date": month_start,
            "month_end_date": month_end,
            "monthly_consumption": group_df["monthly_consumption"].sum(),
            "complete_month_data_captured": covered_days >= calendar_days,
            "covered_days_in_month": covered_days,
            "calendar_days_in_month": calendar_days,
            "proration_component_count": len(group_df),
            "proration_source_row_ids": "; ".join(proration_source_row_ids),
            "proration_split_keys": "; ".join(proration_split_keys),
            "source_file_count": len(source_files),
            "source_files": "; ".join(source_files),
        }
        for index, sharepoint_link in enumerate(sharepoint_links, start=1):
            grouped_row[f"sharepoint_link_{index}"] = sharepoint_link
        for column in SILVER_PRORATION_SPLIT_COLUMNS:
            if column in grouped_row:
                continue
            if column in {"aggregate_quantity1_to_6", "quantity_proration"}:
                grouped_row[column] = _sum_present(group_df[column])
            elif column in {
                "consumption_start_date_proration",
                "consumption_end_date_proration",
                "date_month_proration",
                "date_split_start_proration",
                "date_split_end_proration",
                "source_file",
                "sharepoint_link",
            }:
                grouped_row[column] = _unique_join(group_df[column])
            elif column in {"date_split_days_proration", "date_total_days_proration"}:
                grouped_row[column] = _sum_present(group_df[column])
            elif column in {
                "unit_name",
                "legal_entity",
                "supplier_name",
                "division",
                "division_shorthand",
                "facility_type",
                "scope",
                "facility_identifier",
                "activity_group",
            }:
                grouped_row[column] = _unique_join(group_df[column])
            else:
                grouped_row[column] = _first_present(group_df[column])
        grouped_rows.append(grouped_row)

    grouped_df = pd.DataFrame(grouped_rows)
    return _reindex_silver_prorated_columns(grouped_df)


def build_silver_business_mapping(
    prorated_df: pd.DataFrame,
    config_dir: str | Path = "config",
) -> pd.DataFrame:
    energy_source_allocation_lookup = _business_lookup_table(
        "energy_source_allocation",
        config_dir=config_dir,
    )
    contracts_lookup = _business_lookup_table("contracts", config_dir=config_dir)
    business_rows = []
    for _, row in prorated_df.iterrows():
        activity_group = row.get("activity_group", "")
        activity_group_text = "" if _is_blank(activity_group) else str(activity_group).strip()
        month_start_date = _format_yyyy_mm_dd(row.get("month_start_date"))
        lookup_key = _business_lookup_key(
            row.get("unit_name", ""),
            row.get("supplier_name", ""),
            month_start_date,
        )
        lookup_key_candidates = _business_lookup_key_candidates(row, month_start_date, lookup_key)
        energy_source_allocation = _find_business_lookup_row(
            energy_source_allocation_lookup,
            lookup_key_candidates,
        )
        contract = _find_business_lookup_row(contracts_lookup, lookup_key_candidates)
        business_row = {
            "Data Quality": row.get(DATA_QUALITY_COLUMN, ""),
            "KPI Component": _purchased_kpi_component(activity_group_text),
            "Division": row.get("division", ""),
            "Legal Entity Name": row.get("legal_entity", ""),
            "Unit": row.get("unit_name", ""),
            "Consumption start date": month_start_date,
            "Consumption end date": _format_yyyy_mm_dd(row.get("month_end_date")),
            "Transaction Date": _format_yyyy_mm_dd(row.get("createdDateTme")),
            "Energy Type": activity_group_text,
            "Purchased or acquired": "Purchased",
            "Amount of energy consumed": row.get("quantity_proration", ""),
            "Energy Unit": row.get("consumption_quantity_unit", ""),
            "fossil_fuel_%": energy_source_allocation.get("fossil_fuel_%", ""),
            "renewable_energy_%": energy_source_allocation.get("renewable_energy_%", ""),
            "nuclear_%": energy_source_allocation.get("nuclear_%", ""),
            "contractual_instruments": contract.get("contractual_instruments", ""),
            "Account number": row.get("account_number", ""),
            "Supplier Name": row.get("supplier_name", ""),
            "lookup_key": lookup_key,
            "solar_export": row.get("solar_export", ""),
            "solar_banking_charge": row.get("solar_banking_charge", ""),
            "solar_total_generation": row.get("solar_total_generation", ""),
            "total_solar_consumed": _subtract_numbers(
                row.get("solar_total_generation"),
                row.get("solar_export"),
            ),
            "proration_component_count": row.get("proration_component_count", ""),
            "proration_source_row_ids": row.get("proration_source_row_ids", ""),
            "proration_split_keys": row.get("proration_split_keys", ""),
            "source_file_count": row.get("source_file_count", ""),
            "source_files": row.get("source_files", ""),
            "covered_days_in_month": row.get("covered_days_in_month", ""),
            "calendar_days_in_month": row.get("calendar_days_in_month", ""),
            "complete_month_data_captured": row.get("complete_month_data_captured", ""),
            "status": row.get("status", ""),
            "createdDateTme": row.get("createdDateTme", ""),
            "lastUpdateDateTime": row.get("lastUpdateDateTime", ""),
            "apiVersion": row.get("apiVersion", ""),
            "modelID": row.get("modelID", ""),
            "missing_fields": row.get("missing_fields", ""),
            "low_confidence_fields": row.get("low_confidence_fields", ""),
            "approval_status": row.get("approval_status", ""),
            "manual_review_tag": row.get("manual_review_tag", ""),
            MANUAL_DATA_ENTRY_HELPER_COLUMN: row.get(MANUAL_DATA_ENTRY_HELPER_COLUMN, ""),
        }
        for index in range(1, SILVER_BUSINESS_MAPPING_MAX_SHAREPOINT_LINKS + 1):
            sharepoint_link_column = f"sharepoint_link_{index}"
            if sharepoint_link_column in prorated_df.columns:
                business_row[sharepoint_link_column] = row.get(sharepoint_link_column, "")
        business_rows.append(business_row)

    business_df = pd.DataFrame(business_rows)
    return _reindex_silver_business_mapping_columns(business_df, prorated_df)


def build_silver_template_preparation(
    business_mapping_df: pd.DataFrame,
) -> pd.DataFrame:
    template_rows = []

    for _, row in business_mapping_df.iterrows():
        total_energy_consumed = row.get("Amount of energy consumed", "")

        base_row = {
            "Data Quality": row.get("Data Quality", ""),
            "KPI Component": row.get("KPI Component", ""),
            "Division": row.get("Division", ""),
            "Legal Entity Name": row.get("Legal Entity Name", ""),
            "Unit": row.get("Unit", ""),
            "Consumption start date": row.get("Consumption start date", ""),
            "Consumption end date": row.get("Consumption end date", ""),
            "Transaction Date": row.get("Transaction Date", ""),
            "Energy Type": row.get("Energy Type", ""),
            "Purchased or acquired": row.get("Purchased or acquired", "Purchased"),
            "Energy Unit": row.get("Energy Unit", ""),
            "contractual_instruments": row.get("contractual_instruments", ""),
            "Total amount of energy consumed": total_energy_consumed,
            "fossil_fuel_%": row.get("fossil_fuel_%", ""),
            "renewable_energy_%": row.get("renewable_energy_%", ""),
            "nuclear_%": row.get("nuclear_%", ""),
            "Account number": row.get("Account number", ""),
            "Supplier Name": row.get("Supplier Name", ""),
            "lookup_key": row.get("lookup_key", ""),
            "solar_export": row.get("solar_export", ""),
            "solar_banking_charge": row.get("solar_banking_charge", ""),
            "solar_total_generation": row.get("solar_total_generation", ""),
            "total_solar_consumed": row.get("total_solar_consumed", ""),
            "proration_component_count": row.get("proration_component_count", ""),
            "proration_source_row_ids": row.get("proration_source_row_ids", ""),
            "proration_split_keys": row.get("proration_split_keys", ""),
            "source_file_count": row.get("source_file_count", ""),
            "source_files": row.get("source_files", ""),
            "covered_days_in_month": row.get("covered_days_in_month", ""),
            "calendar_days_in_month": row.get("calendar_days_in_month", ""),
            "complete_month_data_captured": row.get("complete_month_data_captured", ""),
            "status": row.get("status", ""),
            "lastUpdateDateTime": row.get("lastUpdateDateTime", ""),
            "apiVersion": row.get("apiVersion", ""),
            "modelID": row.get("modelID", ""),
            "missing_fields": row.get("missing_fields", ""),
            "low_confidence_fields": row.get("low_confidence_fields", ""),
            "approval_status": row.get("approval_status", ""),
            "manual_review_tag": row.get("manual_review_tag", ""),
            MANUAL_DATA_ENTRY_HELPER_COLUMN: row.get(MANUAL_DATA_ENTRY_HELPER_COLUMN, ""),
        }
        for index in range(1, SILVER_BUSINESS_MAPPING_MAX_SHAREPOINT_LINKS + 1):
            sharepoint_link_column = f"sharepoint_link_{index}"
            base_row[sharepoint_link_column] = row.get(sharepoint_link_column, "")

        for source_column, energy_kpi, calculation_method in SILVER_TEMPLATE_PREPARATION_KPI_COLUMNS:
            template_row = base_row.copy()
            template_row["Energy KPI"] = energy_kpi
            if calculation_method == "percentage":
                template_row["Amount of energy consumed"] = _multiply_numbers(
                    total_energy_consumed,
                    template_row.get(source_column),
                )
            else:
                template_row["Amount of energy consumed"] = _number_or_blank(
                    template_row.get(source_column),
                )
            template_rows.append(template_row)

    template_df = pd.DataFrame(template_rows)
    return template_df.reindex(columns=SILVER_TEMPLATE_PREPARATION_COLUMNS, fill_value="")


def build_silver_template_output(template_preparation_df: pd.DataFrame) -> pd.DataFrame:
    if template_preparation_df.empty:
        return template_preparation_df.reindex(columns=SILVER_TEMPLATE_PREPARATION_COLUMNS, fill_value="")

    keep_rows = template_preparation_df["Amount of energy consumed"].map(_is_nonzero_number)
    return (
        template_preparation_df.loc[keep_rows]
        .reindex(columns=SILVER_TEMPLATE_PREPARATION_COLUMNS, fill_value="")
        .reset_index(drop=True)
    )


def load_completed_manual_data_entry_decisions(path: str | Path) -> pd.DataFrame:
    decisions_path = Path(path)
    if not decisions_path.exists():
        return pd.DataFrame()
    try:
        decisions_df = pd.read_csv(decisions_path, dtype=object, keep_default_na=False)
    except pd.errors.EmptyDataError:
        return pd.DataFrame()
    if decisions_df.empty or "manual_entry_status" not in decisions_df.columns:
        return pd.DataFrame()
    complete_mask = (
        decisions_df["manual_entry_status"].astype(str).str.upper()
        == MANUAL_DATA_ENTRY_COMPLETE_STATUS
    )
    return decisions_df.loc[complete_mask].copy().reset_index(drop=True)


def build_manual_data_entry_proration_input(
    manual_decisions_df: pd.DataFrame,
    config_dir: str | Path | None = None,
) -> pd.DataFrame:
    rows = []
    for index, row in enumerate(manual_decisions_df.to_dict("records"), start=1):
        rows.append({
            DATA_QUALITY_COLUMN: row.get("data_quality", ""),
            "division": row.get("division", ""),
            "legal_entity": row.get("legal_entity_name", ""),
            "unit_name": row.get("unit", ""),
            "account_number": row.get("account_number", ""),
            "supplier_name": row.get("supplier_name", ""),
            "consumption_start_date_1_normalized": row.get("consumption_start_date", ""),
            "consumption_end_date_1_normalized": row.get("consumption_end_date", ""),
            "consumption_start_date_2_normalized": "",
            "consumption_end_date_2_normalized": "",
            "aggregated_quantity_1": row.get("amount_of_energy_consumed", ""),
            "aggregated_quantity_2": "",
            "quantity_unit_1": row.get("energy_unit", ""),
            "consumption_unit": row.get("energy_unit", ""),
            "consumption_quantity_unit": clean_energy_unit(row.get("energy_unit", "")),
            "quantity_unit_2": "",
            "source_file": row.get("source_file", ""),
            "sharepoint_link": row.get("sharepoint_link", ""),
            "status": row.get("manual_entry_status", ""),
            "createdDateTme": row.get("transaction_date", ""),
            "lastUpdateDateTime": row.get("reviewed_at", ""),
            "approval_status": row.get("manual_entry_status", ""),
            "manual_review_tag": "MANUAL_DATA_ENTRY",
            "manual_entry_invoice_id": row.get("invoice_id", ""),
            "manual_entry_line_id": row.get("line_id", ""),
            "manual_entry_reviewer": row.get("reviewed_by", ""),
            "manual_entry_comment": row.get("review_comment", ""),
            "manual_entry_data_quality": row.get("data_quality", ""),
            "manual_entry_transaction_date": row.get("transaction_date", ""),
            MANUAL_DATA_ENTRY_HELPER_COLUMN: MANUAL_DATA_ENTRY_HELPER_VALUE,
        })

    manual_input_df = pd.DataFrame(rows)
    if config_dir is not None and not manual_input_df.empty:
        manual_input_df = add_account_mapping_columns(manual_input_df, config_dir=config_dir)
    return manual_input_df.reindex(columns=_manual_proration_input_columns(), fill_value="")


def _split_proration_period_row(row: pd.Series, period_index: int) -> dict[str, Any] | None:
    source_columns = {
        "consumption_start_date_proration": f"consumption_start_date_{period_index}_normalized",
        "consumption_end_date_proration": f"consumption_end_date_{period_index}_normalized",
        "aggregate_quantity1_to_6": f"aggregated_quantity_{period_index}",
        "average_consumption": (
            "average_consumption_1"
            if period_index == 1
            else "average_consumption_date_2"
        ),
        "date_split_days_proration": f"date_{period_index}_proration",
        "date_total_days_proration": f"date_{period_index}_total_days",
        "quantity_proration": f"prorated_quantity_{period_index}",
        "date_month_proration": f"date_{period_index}_proration_month",
        "date_split_start_proration": f"date_{period_index}_proration_start",
        "date_split_end_proration": f"date_{period_index}_proration_end",
    }
    if all(_is_blank(row.get(source_column)) for source_column in source_columns.values()):
        return None

    split_row = row.to_dict()
    for source_column in SILVER_PRORATION_SPLIT_SOURCE_COLUMNS:
        split_row.pop(source_column, None)
    for target_column, source_column in source_columns.items():
        split_row[target_column] = row.get(source_column, "")
    if period_index == 2 and not _is_blank(row.get("quantity_unit_2")):
        split_row["consumption_quantity_unit"] = row.get("quantity_unit_2", "")
    return split_row


def _aggregate_quantities(row: pd.Series) -> tuple[Any, Any]:
    has_second_consumption_period = not (
        _is_blank(row.get("consumption_start_date_2_normalized"))
        and _is_blank(row.get("consumption_end_date_2_normalized"))
    )
    if has_second_consumption_period:
        return row.get("quantity_1", ""), row.get("quantity_2", "")

    quantity_values = [
        _coerce_number(row.get(f"quantity_{index}"))
        for index in range(1, 7)
    ]
    present_values = [value for value in quantity_values if value is not None]
    if not present_values:
        return "", ""
    total = sum(present_values)
    if float(total).is_integer():
        total = int(total)
    return total, ""


def _effective_consumption_quantity_unit(quantity_unit: Any, consumption_unit: Any) -> Any:
    cleaned_quantity_unit = clean_energy_unit(quantity_unit)
    if not _is_blank(cleaned_quantity_unit):
        return cleaned_quantity_unit
    return clean_energy_unit(consumption_unit)


def _prorate_row_period(row: pd.Series, period_index: int) -> list[pd.Series]:
    start_column = f"consumption_start_date_{period_index}_normalized"
    end_column = f"consumption_end_date_{period_index}_normalized"
    quantity_column = f"aggregated_quantity_{period_index}"
    total_days_column = f"date_{period_index}_total_days"
    proration_column = f"date_{period_index}_proration"
    prorated_quantity_column = (
        "prorated_quantity_1"
        if period_index == 1
        else "prorated_quantity_2"
    )
    average_column = (
        "average_consumption_1"
        if period_index == 1
        else "average_consumption_date_2"
    )
    start_date = _coerce_date(row.get(start_column))
    end_date = _coerce_date(row.get(end_column))
    quantity = _coerce_number(row.get(quantity_column))

    if start_date is None or end_date is None:
        return [row]

    segments = _monthly_segments(start_date, end_date)
    total_days = sum(segment_days for _, _, segment_days in segments)
    if total_days <= 0:
        return [row]

    prorated_rows = []
    average_consumption = ""
    prorated_quantity = ""
    if quantity is not None:
        average_consumption = quantity / total_days

    for segment_start, segment_end, segment_days in segments:
        prorated_row = row.copy()
        if average_consumption != "":
            prorated_quantity = average_consumption * segment_days
        prorated_row[total_days_column] = total_days
        prorated_row[proration_column] = segment_days
        prorated_row[average_column] = average_consumption
        prorated_row[prorated_quantity_column] = prorated_quantity
        prorated_row[f"date_{period_index}_proration_month"] = segment_start.strftime("%Y-%m")
        prorated_row[f"date_{period_index}_proration_start"] = segment_start.isoformat()
        prorated_row[f"date_{period_index}_proration_end"] = segment_end.isoformat()
        prorated_rows.append(prorated_row)
    return prorated_rows


def _monthly_segments(start_date: date, end_date: date) -> list[tuple[date, date, int]]:
    if end_date < start_date:
        return [(start_date, end_date, 0)]
    if start_date.year == end_date.year and start_date.month == end_date.month:
        return [(start_date, end_date, _inclusive_days(start_date, end_date))]

    segments = []
    cursor = start_date
    while cursor <= end_date:
        next_month_start = _first_day_of_next_month(cursor)
        segment_end = min(next_month_start - timedelta(days=1), end_date)
        segment_days = _inclusive_days(cursor, segment_end)
        segments.append((cursor, segment_end, segment_days))
        cursor = segment_end + timedelta(days=1)
    return segments


def _inclusive_days(start_date: date, end_date: date) -> int:
    return (end_date - start_date).days + 1


def _first_day_of_next_month(value: date) -> date:
    _, last_day = monthrange(value.year, value.month)
    return date(value.year, value.month, last_day) + timedelta(days=1)


def _coerce_date(value: Any) -> date | None:
    if _is_blank(value):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.date()


def _coerce_number(value: Any) -> float | None:
    if _is_blank(value) or isinstance(value, bool):
        return None
    if isinstance(value, int | float):
        return float(value)

    text = str(value).strip().replace(",", "")
    try:
        return float(text)
    except ValueError:
        return None


def _is_blank(value: Any) -> bool:
    return value is None or pd.isna(value) or str(value).strip() == ""


def _covered_days(intervals: Any) -> int:
    normalized_intervals = sorted(
        (start, end)
        for start, end in intervals
        if start is not None and end is not None and end >= start
    )
    if not normalized_intervals:
        return 0

    merged = []
    for start, end in normalized_intervals:
        if not merged or start > merged[-1][1] + timedelta(days=1):
            merged.append([start, end])
        elif end > merged[-1][1]:
            merged[-1][1] = end

    return sum((end - start).days + 1 for start, end in merged)


def _first_present(values: pd.Series) -> Any:
    for value in values:
        if not _is_blank(value):
            return value
    return ""


def _clear_inactive_period_values(record: dict[str, Any], active_period_index: int) -> None:
    inactive_period_index = 2 if active_period_index == 1 else 1
    inactive_columns = [
        f"aggregated_quantity_{inactive_period_index}",
        f"date_{inactive_period_index}_proration",
        f"date_{inactive_period_index}_total_days",
        f"prorated_quantity_{inactive_period_index}",
        f"date_{inactive_period_index}_proration_month",
        f"date_{inactive_period_index}_proration_start",
        f"date_{inactive_period_index}_proration_end",
    ]
    if inactive_period_index == 1:
        inactive_columns.append("average_consumption_1")
    else:
        inactive_columns.append("average_consumption_date_2")
    for column in inactive_columns:
        record[column] = ""


def _sum_present(values: pd.Series) -> Any:
    numbers = [
        number
        for number in (_coerce_number(value) for value in values)
        if number is not None
    ]
    if not numbers:
        return ""
    total = sum(numbers)
    if float(total).is_integer():
        return int(total)
    return total


def _unique_join(values: pd.Series) -> str:
    return "; ".join(_unique_text_values(values))


def _unique_text_values(values: pd.Series) -> list[str]:
    unique_values = []
    seen_values = set()
    for value in values:
        if _is_blank(value):
            continue
        text = value.isoformat() if isinstance(value, date) else str(value)
        if text in seen_values:
            continue
        seen_values.add(text)
        unique_values.append(text)
    return unique_values


def _reindex_silver_prorated_columns(df: pd.DataFrame) -> pd.DataFrame:
    sharepoint_link_columns = sorted(
        (
            column
            for column in df.columns
            if _is_numbered_sharepoint_link_column(column)
        ),
        key=lambda column: int(column.rsplit("_", 1)[1]),
    )
    columns = SILVER_PRORATED_COLUMNS.copy()
    if sharepoint_link_columns:
        insert_at = columns.index("sharepoint_link") + 1
        columns[insert_at:insert_at] = sharepoint_link_columns
    result = df.reindex(columns=columns, fill_value="")
    if sharepoint_link_columns:
        result.loc[:, sharepoint_link_columns] = result.loc[:, sharepoint_link_columns].fillna("")
    return result


def _is_numbered_sharepoint_link_column(column: Any) -> bool:
    if not isinstance(column, str) or not column.startswith("sharepoint_link_"):
        return False
    suffix = column.rsplit("_", 1)[1]
    return suffix.isdigit()


def _reindex_silver_business_mapping_columns(
    business_df: pd.DataFrame,
    prorated_df: pd.DataFrame,
) -> pd.DataFrame:
    sharepoint_link_columns = []
    for index in range(1, SILVER_BUSINESS_MAPPING_MAX_SHAREPOINT_LINKS + 1):
        column = f"sharepoint_link_{index}"
        if index <= SILVER_BUSINESS_MAPPING_MIN_SHAREPOINT_LINKS or column in prorated_df.columns:
            sharepoint_link_columns.append(column)
    return business_df.reindex(
        columns=SILVER_BUSINESS_MAPPING_FIXED_COLUMNS + sharepoint_link_columns,
        fill_value="",
    )


def _business_lookup_table(
    lookup_name: str,
    config_dir: str | Path = "config",
) -> dict[str, dict[str, Any]]:
    config = load_lookup_config(lookup_name, config_dir=config_dir)
    table = load_lookup_table(config)
    lookup: dict[str, dict[str, Any]] = {}
    for _, row in table.iterrows():
        record = row.drop(labels=["_python_lookup_key"], errors="ignore").to_dict()
        keys = []
        if "lookup_key" in record:
            keys.append(record.get("lookup_key"))
        keys.append(row.get("_python_lookup_key"))
        for key in keys:
            normalized_key = normalize_lookup_value(key)
            if normalized_key and normalized_key not in lookup:
                lookup[normalized_key] = record
    return lookup


def _find_business_lookup_row(
    lookup: dict[str, dict[str, Any]],
    lookup_key_candidates: list[str],
) -> dict[str, Any]:
    for lookup_key in lookup_key_candidates:
        match = lookup.get(normalize_lookup_value(lookup_key))
        if match is not None:
            return match
    return {}


def _business_lookup_key_candidates(
    row: pd.Series,
    month_start_date: str,
    lookup_key: str,
) -> list[str]:
    candidates = [lookup_key]
    serial_lookup_key = build_lookup_key(
        {
            "unit_name": row.get("unit_name", ""),
            "supplier_name": row.get("supplier_name", ""),
            "start_date": month_start_date,
            "contract_start_date": month_start_date,
        },
        ("unit_name", "supplier_name", "start_date"),
        delimiter="_",
        date_key_fields=("start_date",),
    )
    candidates.append(serial_lookup_key)
    return candidates


def _purchased_kpi_component(activity_group: str) -> str:
    if _is_blank(activity_group):
        return ""
    return f"Purchased {str(activity_group).strip().lower()}"


def _business_lookup_key(unit_name: Any, supplier_name: Any, month_start_date: str) -> str:
    values = [unit_name, supplier_name, month_start_date]
    if all(_is_blank(value) for value in values):
        return ""
    return "_".join("" if _is_blank(value) else str(value).strip() for value in values).lower()


def _format_yyyy_mm_dd(value: Any) -> str:
    parsed_date = _coerce_date(value)
    if parsed_date is None:
        return ""
    return parsed_date.isoformat()


def _subtract_numbers(left: Any, right: Any) -> Any:
    left_number = _coerce_number(left)
    if left_number is None:
        return ""
    right_number = _coerce_number(right)
    if right_number is None:
        right_number = 0
    result = left_number - right_number
    if float(result).is_integer():
        return int(result)
    return result


def _multiply_numbers(left: Any, right: Any) -> Any:
    left_number = _coerce_number(left)
    right_number = _coerce_number(right)
    if left_number is None or right_number is None:
        return ""
    result = left_number * right_number
    if float(result).is_integer():
        return int(result)
    return result


def _number_or_blank(value: Any) -> Any:
    number = _coerce_number(value)
    if number is None:
        return ""
    if float(number).is_integer():
        return int(number)
    return number


def _is_nonzero_number(value: Any) -> bool:
    number = _coerce_number(value)
    return number is not None and number != 0


def _first_existing_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for column in candidates:
        if column in df.columns:
            return column
    return None


def _manual_proration_input_columns() -> list[str]:
    columns = SILVER_AGGREGATED_COLUMNS.copy()
    for column in [*MANUAL_DATA_ENTRY_FLOW_COLUMNS, MANUAL_DATA_ENTRY_HELPER_COLUMN]:
        if column not in columns:
            columns.append(column)
    return columns


def _silver_input_file(
    checkpoint_dir: Path,
    bronze_file: str | Path | None,
    approved_file: str | Path | None,
) -> Path:
    if approved_file is not None:
        return Path(approved_file)
    if bronze_file is not None:
        return Path(bronze_file)
    return checkpoint_dir / APPROVED_SILVER_FILENAME


def _remove_legacy_silver_outputs(silver_dir: Path, invoice_type: str) -> None:
    for layer in SILVER_OUTPUT_STEPS:
        legacy_path = silver_dir / legacy_silver_output_filename(invoice_type, layer)
        _remove_silver_output(legacy_path)

    expected_numbered_outputs = {
        silver_output_filename(invoice_type, layer)
        for layer in SILVER_OUTPUT_STEPS
    }
    for numbered_path in silver_dir.glob(f"*_{invoice_type}_silver_*.xlsx"):
        if numbered_path.name not in expected_numbered_outputs:
            _remove_silver_output(numbered_path)


def _remove_silver_output(path: Path) -> None:
    if not path.exists():
        return
    try:
        path.unlink()
    except PermissionError as exc:
        raise PermissionError(
            f"Cannot remove stale silver output {path}. "
            "Close the workbook if it is open in Excel or another app, then run the pipeline again."
        ) from exc
