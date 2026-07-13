from __future__ import annotations

from pathlib import Path

import pandas as pd

from src.core.config_loader import load_yaml
from src.core.logger import get_logger
from src.core.path_settings import gold_output_dir, silver_excel_output_dir
from src.output.xlsx_writer import write_xlsx
from src.pipeline.run_silver_pipeline import DEFAULT_INVOICE_TYPE, silver_output_filename
from src.transform.unit_conversion import apply_unit_conversions


logger = get_logger(__name__)


def gold_output_filename(invoice_type: str, layer: str) -> str:
    return f"{invoice_type}_gold_{layer}.xlsx"


def run_gold_pipeline(
    config_dir: str | Path = "config",
    invoice_type: str | None = None,
    template_output_file: str | Path | None = None,
) -> Path:
    config_dir = Path(config_dir)
    settings = load_yaml(config_dir / "settings.yaml")
    invoice_type = invoice_type or settings.get("pipeline", {}).get("invoice_type", DEFAULT_INVOICE_TYPE)

    silver_dir = silver_excel_output_dir(settings)
    gold_dir = gold_output_dir(settings)
    template_output_path = (
        Path(template_output_file)
        if template_output_file is not None
        else silver_dir / silver_output_filename(invoice_type, "template_output")
    )

    if not template_output_path.exists():
        raise FileNotFoundError(f"Silver template output workbook does not exist: {template_output_path}")

    logger.info("Reading silver-template-output workbook for gold template from %s", template_output_path)
    template_output_df = pd.read_excel(
        template_output_path,
        sheet_name="SilverTemplateOutput",
        dtype=object,
        keep_default_na=False,
        engine="openpyxl",
    )
    gold_template_df = build_gold_template(template_output_df)
    gold_template_path = write_xlsx(
        gold_template_df,
        gold_dir,
        gold_output_filename(invoice_type, "template"),
        sheet_name="GoldTemplate",
    )
    logger.info("Wrote gold-template Excel workbook to %s", gold_template_path)
    return gold_template_path


def build_gold_template(template_output_df: pd.DataFrame) -> pd.DataFrame:
    return apply_unit_conversions(template_output_df)
