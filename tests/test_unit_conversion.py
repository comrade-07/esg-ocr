import pandas as pd

from src.transform.unit_conversion import (
    UnitConversionRule,
    apply_unit_conversions,
)


def test_apply_unit_conversions_converts_mwh_to_kwh():
    df = pd.DataFrame([
        {
            "Amount of energy consumed": 1.5,
            "Energy Unit": "MWh",
        },
        {
            "Amount of energy consumed": "2",
            "Energy Unit": "MWH",
        },
        {
            "Amount of energy consumed": 300,
            "Energy Unit": "kWh",
        },
    ])

    result = apply_unit_conversions(df)

    assert list(result["Amount of energy consumed"]) == [1500, 2000, 300]
    assert list(result["Energy Unit"]) == ["kWh", "kWh", "kWh"]
    assert list(result["Original Energy Unit"]) == ["MWh", "MWH", "kWh"]


def test_apply_unit_conversions_supports_custom_rules():
    df = pd.DataFrame([
        {
            "Amount of energy consumed": 3,
            "Energy Unit": "GWh",
        },
    ])

    result = apply_unit_conversions(
        df,
        conversion_rules=[
            UnitConversionRule(from_unit="GWh", to_unit="kWh", factor=1_000_000),
        ],
    )

    assert result.loc[0, "Amount of energy consumed"] == 3000000
    assert result.loc[0, "Energy Unit"] == "kWh"
    assert result.loc[0, "Original Energy Unit"] == "GWh"


def test_apply_unit_conversions_places_original_unit_after_total_amount():
    df = pd.DataFrame([
        {
            "Energy KPI": "Fossil Fuels",
            "Amount of energy consumed": 1,
            "Energy Unit": "MWh",
            "Total amount of energy consumed": 2,
            "source_files": "one.json",
        },
    ])

    result = apply_unit_conversions(df)

    assert list(result.columns) == [
        "Energy KPI",
        "Amount of energy consumed",
        "Energy Unit",
        "Total amount of energy consumed",
        "Original Energy Unit",
        "source_files",
    ]
