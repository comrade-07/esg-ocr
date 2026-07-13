from pathlib import Path
import json


def list_json_files(input_dir: str | Path) -> list[Path]:
    input_path = Path(input_dir)
    if not input_path.exists():
        raise FileNotFoundError(f"Input folder does not exist: {input_path}")
    return sorted(input_path.glob("*.json"))


def read_json_file(path: str | Path) -> dict:
    path = Path(path)
    with path.open("r", encoding="utf-8") as file:
        return json.load(file)
