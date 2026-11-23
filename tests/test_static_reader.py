"""Tests for static_reader.py — parsing static condominium data XLSX."""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from siemens_converter.static_reader import StaticDataError, parse_static_data

FIXTURE = Path(__file__).parent / "fixtures" / "static_data_TEST.xlsx"
REPORT = Path(__file__).parent / "fixtures" / "FC_report_TEST_9999_2026-01-15.xls"

TEMPLATE_ADVICE = (
    "Scegliere il file dei dati statici, compilato a partire da "
    "static_data_template.xlsx (si scarica dalla pagina dell'ultima versione: "
    "https://github.com/gjed/siemens-converter/releases/latest), "
    "e ripetere la conversione."
)
ACCEPTED_LABELS = (
    "«Energia elettrica», «Gas metano», «Acqua condominio», "
    "«Conduzione e manutenzione», «Contabilizzazione», "
    "«Acqua sanitaria manutenzione»"
)


def refusal(path: Path) -> str:
    with pytest.raises(StaticDataError) as excinfo:
        parse_static_data(path)
    return str(excinfo.value)


def edited(tmp_path: Path, sheet: str, **cells) -> Path:
    """Copy the fixture into tmp_path with the given cells of *sheet* changed."""
    wb = openpyxl.load_workbook(FIXTURE)
    for cell, value in cells.items():
        wb[sheet][cell] = value
    path = tmp_path / "dati_2026.xlsx"
    wb.save(path)
    return path


def test_parse_valid_file():
    sd = parse_static_data(FIXTURE)
    assert len(sd.apartments) == 10
    assert len(sd.millesimali) == 10
    assert len(sd.costs) == 6
    assert len(sd.meters) == 3
    assert len(sd.previous_readings) == 10


def test_apartments_populated():
    sd = parse_static_data(FIXTURE)
    apt1 = sd.apartments[0]
    assert apt1.apartment_number == 1
    assert apt1.proprietario == "Rossi Mario"
    assert apt1.inquilino == "Bianchi Anna"


def test_apartment_empty_inquilino():
    sd = parse_static_data(FIXTURE)
    apt10 = sd.apartments[9]
    assert apt10.apartment_number == 10
    assert apt10.inquilino == ""


def test_millesimali_populated():
    sd = parse_static_data(FIXTURE)
    m1 = sd.millesimali[0]
    assert m1.apartment_number == 1
    assert m1.subalterno == 20
    assert m1.heat_energy_kwh == 1344.03
    assert m1.water_energy_kwh == 1101.36


def test_costs_populated():
    sd = parse_static_data(FIXTURE)
    cost_dict = {c.label: c.amount for c in sd.costs}
    assert cost_dict["Energia elettrica"] == 263.44
    assert cost_dict["Gas metano"] == 2141.0
    assert cost_dict["Acqua condominio"] == 0.0


def test_meters_populated():
    sd = parse_static_data(FIXTURE)
    # Fixture now stores uppercase names matching the template display style
    m = {mr.name.upper(): mr for mr in sd.meters}
    assert m["ENERGIA ELETTRICA CT"].initial == 0
    assert m["ENERGIA ELETTRICA CT"].final == 1086
    assert m["GAS METANO CT"].final == 1718


def test_previous_readings_populated():
    sd = parse_static_data(FIXTURE)
    pr1 = sd.previous_readings[0]
    assert pr1.apartment_number == 1
    assert pr1.heat_kwh == 4048
    assert pr1.water_m3 == 25.5
    assert pr1.cold_water_m3 == 10.2


def test_file_that_is_not_a_workbook(tmp_path: Path):
    path = tmp_path / "dati_2026.xlsx"
    path.write_bytes(REPORT.read_bytes())

    assert refusal(path) == (
        "«dati_2026.xlsx» non è un file dei dati statici: non si apre come file "
        "Excel (.xlsx). " + TEMPLATE_ADVICE
    )


def test_file_with_another_extension(tmp_path: Path):
    path = tmp_path / "FC_report.xls"
    path.write_bytes(REPORT.read_bytes())

    assert refusal(path).startswith(
        "«FC_report.xls» non è un file dei dati statici: non si apre come file Excel"
    )


def test_missing_sheet(tmp_path: Path):
    wb = openpyxl.load_workbook(FIXTURE)
    del wb["Contatori"]
    path = tmp_path / "missing_sheet.xlsx"
    wb.save(path)

    assert refusal(path) == (
        "«missing_sheet.xlsx» non è un file dei dati statici: manca il foglio "
        "«Contatori». Se il file è quello giusto, controllare che nessun foglio "
        "sia stato rinominato o eliminato. " + TEMPLATE_ADVICE
    )


def test_missing_sheets(tmp_path: Path):
    wb = openpyxl.load_workbook(FIXTURE)
    wb["Costi"].title = "Spese"
    del wb["Inquilini"]
    path = tmp_path / "missing_sheets.xlsx"
    wb.save(path)

    assert refusal(path) == (
        "«missing_sheets.xlsx» non è un file dei dati statici: mancano i fogli "
        "«Costi», «Inquilini». Se il file è quello giusto, controllare che nessun "
        "foglio sia stato rinominato o eliminato. " + TEMPLATE_ADVICE
    )


def test_missing_cost_label_defaults_to_zero(tmp_path: Path):
    """Costi sheet missing 'Acqua condominio' row should default to 0.0."""
    wb = openpyxl.load_workbook(FIXTURE)
    ws = wb["Costi"]
    # New layout: row1=title, row2=header, row3=Energia el., row4=Gas metano,
    #             row5=Acqua condominio, ...
    ws.delete_rows(5)  # remove "Acqua condominio"
    patched = tmp_path / "missing_cost.xlsx"
    wb.save(patched)

    sd = parse_static_data(patched)
    cost_dict = {c.label: c.amount for c in sd.costs}
    assert cost_dict["Acqua condominio"] == 0.0
    # Other costs should still be present
    assert cost_dict["Gas metano"] == 2141.0


def test_fewer_apartment_rows(tmp_path: Path):
    """Inquilini with only 3 data rows should produce 3 apartments (not error)."""
    wb = openpyxl.load_workbook(FIXTURE)
    ws = wb["Inquilini"]
    # New layout: row1=title, row2=header, rows3-12=apts 1-10.
    # Keep only rows 3,4,5 (apts 1,2,3) — delete rows 6-12 (7 rows from row 6).
    ws.delete_rows(6, 7)
    patched = tmp_path / "fewer_apts.xlsx"
    wb.save(patched)

    sd = parse_static_data(patched)
    assert len(sd.apartments) == 3
    assert sd.apartments[0].apartment_number == 1
    assert sd.apartments[2].apartment_number == 3


def test_cost_label_case_and_whitespace_insensitive(tmp_path: Path):
    """'gas METANO ' must still map to Gas metano, not silently become 0."""
    wb = openpyxl.load_workbook(FIXTURE)
    ws = wb["Costi"]
    ws["A4"] = "gas METANO "  # row 4 = Gas metano
    patched = tmp_path / "case_cost.xlsx"
    wb.save(patched)

    sd = parse_static_data(patched)
    cost_dict = {c.label: c.amount for c in sd.costs}
    assert cost_dict["Gas metano"] == 2141.0


def test_unknown_cost_label(tmp_path: Path):
    path = edited(tmp_path, "Costi", A4="  Gas metanno ")

    assert refusal(path) == (
        "«dati_2026.xlsx», foglio «Costi»: queste voci di spesa non sono tra quelle "
        "accettate:\n"
        "- «Gas metanno» (cella A4)\n\n"
        f"Le voci accettate sono: {ACCEPTED_LABELS}. Scrivere ogni voce come nel "
        "modello, oppure eliminare la riga, e ripetere la conversione."
    )


def test_every_unknown_cost_label_is_listed(tmp_path: Path):
    path = edited(tmp_path, "Costi", A3="Luce", A7="Pulizie")

    message = refusal(path)

    assert "\n- «Luce» (cella A3)\n- «Pulizie» (cella A7)\n" in message


@pytest.mark.parametrize(
    ("sheet", "cell", "value", "kind"),
    [
        ("Costi", "B4", "abc", "un numero"),
        ("Millesimali", "C4", "1.344,03 kWh", "un numero"),
        ("Millesimali", "E5", "n/d", "un numero"),
        ("Millesimali", "B6", "terzo", "un numero intero"),
        ("Contatori", "C3", "zero", "un numero"),
        ("Contatori", "D5", "680 mc", "un numero"),
        ("Letture_precedenti", "B3", "quattro", "un numero intero"),
        ("Letture_precedenti", "C4", "30,1 m3", "un numero"),
        ("Letture_precedenti", "D5", "?", "un numero"),
    ],
)
def test_value_that_is_not_a_number(tmp_path: Path, sheet, cell, value, kind):
    path = edited(tmp_path, sheet, **{cell: value})

    assert refusal(path) == (
        f"«dati_2026.xlsx», foglio «{sheet}», cella {cell}: il valore «{value}» "
        f"non è {kind}. Scrivere solo il numero, senza unità di misura e senza "
        "simboli, e ripetere la conversione."
    )


def test_blank_number_cells_still_read_as_zero(tmp_path: Path):
    path = edited(
        tmp_path,
        "Letture_precedenti",
        B3=None,
        C3=None,
        D3=None,
    )
    wb = openpyxl.load_workbook(path)
    wb["Costi"]["B4"] = None
    wb["Contatori"]["D3"] = None
    wb["Millesimali"]["C4"] = None
    wb.save(path)

    sd = parse_static_data(path)

    assert (sd.previous_readings[0].heat_kwh, sd.previous_readings[0].water_m3) == (
        0,
        0.0,
    )
    assert sd.previous_readings[0].cold_water_m3 == 0.0
    assert {c.label: c.amount for c in sd.costs}["Gas metano"] == 0.0
    assert sd.meters[0].final == 0.0
    assert sd.millesimali[0].heat_energy_kwh == 0.0
