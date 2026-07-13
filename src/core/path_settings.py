from pathlib import Path


def bronze_output_dir(settings: dict) -> Path:
    paths = settings.get("paths", {})
    return Path(paths.get("bronze_output", paths.get("csv_output", "data/bronze")))


def silver_excel_output_dir(settings: dict) -> Path:
    paths = settings.get("paths", {})
    return Path(paths.get("silver_excel_output", paths.get("silver_output", "data/silver")))


def checkpoint_output_dir(settings: dict) -> Path:
    paths = settings.get("paths", {})
    return Path(
        paths.get(
            "review_checkpoint_output",
            paths.get("checkpoint_output", "data/output/checkpoints"),
        )
    )


def manual_data_entry_upload_dir(settings: dict) -> Path:
    paths = settings.get("paths", {})
    return Path(paths.get("manual_data_entry_uploads", "data/manual_uploads"))


def gold_output_dir(settings: dict) -> Path:
    paths = settings.get("paths", {})
    return Path(paths.get("gold_output", "data/gold"))
