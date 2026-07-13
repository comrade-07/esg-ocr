import pandas as pd

from src.normalize.text_normalizer import clean_energy_unit, clean_proper_text, clean_upper_text
from src.transform.text_fields import clean_text_columns


def test_clean_proper_text_strips_edges_and_replaces_separators():
    assert clean_proper_text("  _STOLT-NIELSEN MIDDLE EAST DMCC, ") == "Stolt Nielsen Middle East DMCC"
    assert clean_proper_text("(Durban)") == "Durban"
    assert clean_proper_text("STOLT TANK CONTAINERS COLOMBIA SAS") == "Stolt Tank Containers Colombia SAS"


def test_clean_upper_text_strips_edges_replaces_separators_and_uppercases():
    assert clean_upper_text("  Dubai-Electricity & Water_Authority, ") == "DUBAI ELECTRICITY & WATER AUTHORITY"


def test_clean_energy_unit_keeps_energy_units_as_uppercase():
    assert clean_energy_unit(" kWh ") == "KWH"
    assert clean_energy_unit("MWH") == "MWH"
    assert clean_energy_unit("kw-h") == "KWH"
    assert clean_energy_unit("megawatt hours") == "MWH"


def test_clean_energy_unit_removes_non_energy_units():
    assert clean_energy_unit("litres") == ""
    assert clean_energy_unit("m3") == ""
    assert clean_energy_unit("gallons") == ""


def test_clean_text_columns_only_targets_configured_text_fields():
    df = pd.DataFrame([{
        "legal_entity": " STOLT-NIELSEN MIDDLE EAST DMCC ",
        "unit_name": "(Durban)",
        "supplier": " Dubai Electricity & Water Authority, ",
        "quantity_unit_1": " kWh ",
        "quantity_unit_2": "m3",
        "invoice_date": "25 Jul 2025",
    }])

    result = clean_text_columns(df)

    assert result.loc[0, "legal_entity"] == "Stolt Nielsen Middle East DMCC"
    assert result.loc[0, "unit_name"] == "Durban"
    assert result.loc[0, "supplier"] == "DUBAI ELECTRICITY & WATER AUTHORITY"
    assert result.loc[0, "quantity_unit_1"] == "KWH"
    assert result.loc[0, "quantity_unit_2"] == ""
    assert result.loc[0, "invoice_date"] == "25 Jul 2025"
