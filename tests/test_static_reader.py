"""Tests for static_reader.py — parsing static condominium data XLSX."""

from __future__ import annotations

import dataclasses
from datetime import datetime
from pathlib import Path

import openpyxl
import pytest

from siemens_converter.models import CentralMeterStart, PreviousReading
from siemens_converter.parser import parse_fc_report
from siemens_converter.static_reader import (
    StaticDataError,
    check_against_report,
    parse_static_data,
)

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
FIRST_ROW = {
    "Inquilini": 3,
    "Millesimali": 5,
    "Letture_precedenti": 3,
}
APARTMENT_SHEETS = list(FIRST_ROW)


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


def patched(tmp_path: Path, change) -> Path:
    """Copy the fixture into tmp_path after change(workbook) has altered it."""
    wb = openpyxl.load_workbook(FIXTURE)
    change(wb)
    path = tmp_path / "dati_2026.xlsx"
    wb.save(path)
    return path


# -- A valid file --


def test_parse_valid_file():
    sd = parse_static_data(FIXTURE)
    assert len(sd.apartments) == 10
    assert len(sd.millesimali) == 10
    assert len(sd.costs) == 6
    assert len(sd.meters) == 3
    assert len(sd.previous_readings) == 10
    assert len(sd.central_meters) == 2


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
    assert m1.subalterno == 101
    assert m1.heat_energy_kwh == 1250.5
    assert m1.water_energy_kwh == 820.0


def test_costs_populated():
    sd = parse_static_data(FIXTURE)
    cost_dict = {c.label: c.amount for c in sd.costs}
    assert cost_dict["Energia elettrica"] == 310.5
    assert cost_dict["Gas metano"] == 2450.0
    assert cost_dict["Acqua condominio"] == 0.0  # a cost that is really zero


def test_meters_populated():
    sd = parse_static_data(FIXTURE)
    m = {mr.name.upper(): mr for mr in sd.meters}
    assert m["ENERGIA ELETTRICA CT"].initial == 5000
    assert m["ENERGIA ELETTRICA CT"].final == 6150
    assert m["GAS METANO CT"].final == 13800


def test_previous_readings_populated():
    sd = parse_static_data(FIXTURE)
    pr1 = sd.previous_readings[0]
    assert pr1.apartment_number == 1
    assert pr1.heat_kwh == 4200
    assert pr1.water_m3 == 12.5
    assert pr1.cold_water_m3 == 20.4


def test_exercise_populated():
    sd = parse_static_data(FIXTURE)
    assert sd.exercise.year == 2025
    assert sd.exercise.start_date == datetime(2024, 12, 31)


def test_central_meters_populated():
    sd = parse_static_data(FIXTURE)
    heating, acs = sd.central_meters
    assert (heating.description, heating.initial_kwh) == ("Riscaldamento", 24000)
    assert (acs.description, acs.initial_kwh) == ("Sanitario", 22000)
    assert heating.final_kwh is None and heating.final_date is None


def test_central_meter_manual_final_reading(tmp_path: Path):
    path = edited(tmp_path, "Esercizio", D5=27500, E5=datetime(2026, 1, 10))

    heating, acs = parse_static_data(path).central_meters

    assert (heating.final_kwh, heating.final_date) == (27500, datetime(2026, 1, 10))
    assert (acs.final_kwh, acs.final_date) == (None, None)


def test_text_in_the_total_row_is_the_only_row_that_is_not_data(tmp_path: Path):
    def add_totals(wb):
        wb["Inquilini"]["A13"] = "Totale"
        wb["Letture_precedenti"]["A13"] = "TOTALE"
        wb["Costi"]["A10"] = "totale"
        wb["Contatori"]["A6"] = "TOTALE"

    sd = parse_static_data(patched(tmp_path, add_totals))

    assert len(sd.apartments) == len(sd.previous_readings) == 10


def test_row_below_a_blank_row_is_read(tmp_path: Path):
    def blank_row_between(wb):
        wb["Costi"].insert_rows(5)  # a blank row between two costs

    sd = parse_static_data(patched(tmp_path, blank_row_between))

    assert {c.label: c.amount for c in sd.costs}["Conduzione e manutenzione"] == 420


# -- Files that are not static data files --


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


def test_file_made_with_the_previous_template_has_no_exercise_sheet(tmp_path: Path):
    wb = openpyxl.load_workbook(FIXTURE)
    del wb["Esercizio"]
    path = tmp_path / "dati_2025.xlsx"
    wb.save(path)

    assert refusal(path) == (
        "«dati_2025.xlsx» non è un file dei dati statici: manca il foglio "
        "«Esercizio». Se il file è quello giusto, controllare che nessun foglio "
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


# -- Cost rows --


def test_cost_label_case_and_whitespace_insensitive(tmp_path: Path):
    """'gas METANO ' must still map to Gas metano, not be refused."""
    sd = parse_static_data(edited(tmp_path, "Costi", A4="gas METANO "))

    assert {c.label: c.amount for c in sd.costs}["Gas metano"] == 2450.0


def test_missing_cost_row_is_refused(tmp_path: Path):
    path = patched(tmp_path, lambda wb: wb["Costi"].delete_rows(5))  # Acqua condominio

    assert refusal(path) == (
        "«dati_2026.xlsx», foglio «Costi»: manca la voce «Acqua condominio». "
        "Aggiungere la riga della voce, scritta come nel modello e con i suoi "
        "valori (se un valore è davvero zero, scrivere 0), e ripetere la conversione."
    )


def test_every_missing_cost_row_is_listed(tmp_path: Path):
    path = patched(tmp_path, lambda wb: wb["Costi"].delete_rows(7, 2))

    assert refusal(path) == (
        "«dati_2026.xlsx», foglio «Costi»: mancano le voci «Contabilizzazione», "
        "«Acqua sanitaria manutenzione». Aggiungere la riga delle voci, scritta "
        "come nel modello e con i suoi valori (se un valore è davvero zero, "
        "scrivere 0), e ripetere la conversione."
    )


def test_duplicate_cost_row(tmp_path: Path):
    path = edited(tmp_path, "Costi", A8="Gas metano")

    assert refusal(path) == (
        "«dati_2026.xlsx», foglio «Costi»: la voce «Gas metano» compare due volte "
        "(celle A4 e A8). Lasciare una sola riga e ripetere la conversione."
    )


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


def test_stray_text_row_in_costi_is_refused_not_skipped(tmp_path: Path):
    path = edited(tmp_path, "Costi", A10="Note", B10=1)

    assert refusal(path).startswith(
        "«dati_2026.xlsx», foglio «Costi»: queste voci di spesa non sono tra "
        "quelle accettate:\n- «Note» (cella A10)\n"
    )


def test_cost_amount_without_its_label(tmp_path: Path):
    path = edited(tmp_path, "Costi", A5=None)

    assert refusal(path) == (
        "«dati_2026.xlsx», foglio «Costi», cella A5: manca il nome (la voce), ma "
        "la riga contiene altri valori. Scriverlo come nel modello o cancellare "
        "la riga, e ripetere la conversione."
    )


# -- Meter rows --


def test_missing_meter_row(tmp_path: Path):
    path = patched(tmp_path, lambda wb: wb["Contatori"].delete_rows(4))

    assert refusal(path) == (
        "«dati_2026.xlsx», foglio «Contatori»: manca il contatore «GAS METANO CT». "
        "Aggiungere la riga del contatore, scritta come nel modello e con i suoi "
        "valori (se un valore è davvero zero, scrivere 0), e ripetere la conversione."
    )


def test_duplicate_meter_row(tmp_path: Path):
    path = edited(tmp_path, "Contatori", A6="gas metano ct", C6=1, D6=2)

    assert refusal(path) == (
        "«dati_2026.xlsx», foglio «Contatori»: il contatore «GAS METANO CT» "
        "compare due volte (celle A4 e A6). Lasciare una sola riga e ripetere la "
        "conversione."
    )


def test_unknown_meter(tmp_path: Path):
    path = edited(tmp_path, "Contatori", A5="Acqua di pozzo")

    assert refusal(path) == (
        "«dati_2026.xlsx», foglio «Contatori»: questi contatori non sono tra quelli "
        "accettati:\n- «Acqua di pozzo» (cella A5)\n\n"
        "I contatori accettati sono: «ENERGIA ELETTRICA CT», «GAS METANO CT», "
        "«ACQUA GENERALE». Scrivere ogni contatore come nel modello, oppure "
        "eliminare la riga, e ripetere la conversione."
    )


# -- Numbers --


@pytest.mark.parametrize(
    ("sheet", "cell", "value", "kind"),
    [
        ("Costi", "B4", "abc", "un numero"),
        ("Millesimali", "C5", "1.250,5 kWh", "un numero"),
        ("Millesimali", "E6", "n/d", "un numero"),
        ("Millesimali", "B7", "terzo", "un numero intero"),
        ("Contatori", "C3", "zero", "un numero"),
        ("Contatori", "D5", "700 mc", "un numero"),
        ("Letture_precedenti", "B3", "quattro", "un numero intero"),
        ("Letture_precedenti", "C4", "20,1 m3", "un numero"),
        ("Letture_precedenti", "D5", "?", "un numero"),
        ("Esercizio", "C5", "24000 kWh", "un numero"),
        ("Esercizio", "C6", "ventimila", "un numero"),
    ],
)
def test_value_that_is_not_a_number(tmp_path: Path, sheet, cell, value, kind):
    path = edited(tmp_path, sheet, **{cell: value})

    assert refusal(path) == (
        f"«dati_2026.xlsx», foglio «{sheet}», cella {cell}: il valore «{value}» "
        f"non è {kind}. Scrivere solo il numero, senza unità di misura e senza "
        "simboli, e ripetere la conversione."
    )


@pytest.mark.parametrize(
    ("sheet", "cell"),
    [
        ("Costi", "B4"),
        ("Millesimali", "C5"),
        ("Millesimali", "E5"),
        ("Millesimali", "B5"),
        ("Contatori", "C3"),
        ("Contatori", "D3"),
        ("Letture_precedenti", "B3"),
        ("Letture_precedenti", "C3"),
        ("Letture_precedenti", "D3"),
        ("Esercizio", "C5"),
        ("Esercizio", "C6"),
    ],
)
def test_blank_number_cell_is_refused(tmp_path: Path, sheet, cell):
    path = edited(tmp_path, sheet, **{cell: None})

    assert refusal(path) == (
        f"«dati_2026.xlsx», foglio «{sheet}», cella {cell}: la cella è vuota ma "
        "serve un numero. Scrivere il numero (se è davvero zero, scrivere 0) e "
        "ripetere la conversione."
    )


def test_a_real_zero_is_accepted(tmp_path: Path):
    path = edited(tmp_path, "Costi", B4=0)

    assert {c.label: c.amount for c in parse_static_data(path).costs}["Gas metano"] == 0


# -- Apartment sheets --


@pytest.mark.parametrize("sheet", APARTMENT_SHEETS)
def test_duplicate_apartment_row(tmp_path: Path, sheet):
    first = FIRST_ROW[sheet]
    # Apartment 10 appears a second time, on the row after the last one
    path = edited(tmp_path, sheet, **{f"A{first + 10}": 10})

    assert refusal(path) == (
        f"«dati_2026.xlsx», foglio «{sheet}»: l'appartamento 10 compare due volte "
        f"(celle A{first + 9} e A{first + 10}). Lasciare una sola riga per ogni "
        "appartamento e ripetere la conversione."
    )


@pytest.mark.parametrize("sheet", APARTMENT_SHEETS)
def test_duplicate_apartment_replacing_another_row(tmp_path: Path, sheet):
    first = FIRST_ROW[sheet]
    path = edited(tmp_path, sheet, **{f"A{first + 4}": 2})  # apartment 5 becomes 2

    assert refusal(path) == (
        f"«dati_2026.xlsx», foglio «{sheet}»: l'appartamento 2 compare due volte "
        f"(celle A{first + 1} e A{first + 4}). Lasciare una sola riga per ogni "
        "appartamento e ripetere la conversione."
    )


@pytest.mark.parametrize("sheet", APARTMENT_SHEETS)
def test_missing_apartment_row(tmp_path: Path, sheet):
    first = FIRST_ROW[sheet]
    path = patched(tmp_path, lambda wb: wb[sheet].delete_rows(first + 6))

    assert refusal(path) == (
        f"«dati_2026.xlsx», foglio «{sheet}»: manca l'appartamento 7. Aggiungere "
        "la riga dell'appartamento, con il suo numero nella colonna A, e "
        "ripetere la conversione."
    )


def test_fewer_apartment_rows_are_refused(tmp_path: Path):
    # Keep only the rows of apartments 1, 2 and 3
    path = patched(tmp_path, lambda wb: wb["Inquilini"].delete_rows(6, 7))

    assert refusal(path) == (
        "«dati_2026.xlsx», foglio «Inquilini»: mancano gli appartamenti 4, 5, 6, "
        "7, 8, 9 e 10. Aggiungere la riga di ciascuno, con il suo numero nella "
        "colonna A, e ripetere la conversione."
    )


@pytest.mark.parametrize("sheet", APARTMENT_SHEETS)
def test_stray_text_row_is_refused_not_skipped(tmp_path: Path, sheet):
    first = FIRST_ROW[sheet]
    # A note typed in the middle of the table
    path = edited(tmp_path, sheet, **{f"A{first + 4}": "Verificare"})

    assert refusal(path) == (
        f"«dati_2026.xlsx», foglio «{sheet}», cella A{first + 4}: il valore "
        "«Verificare» non è un numero di appartamento. In questo foglio sono "
        "ammesse solo le righe degli appartamenti da 1 a 10 e la riga TOTALE: "
        "correggere il valore o cancellare la riga, e ripetere la conversione."
    )


@pytest.mark.parametrize("sheet", APARTMENT_SHEETS)
def test_stray_text_row_below_the_table_is_refused(tmp_path: Path, sheet):
    first = FIRST_ROW[sheet]
    path = edited(tmp_path, sheet, **{f"A{first + 11}": "Note"})

    assert refusal(path).startswith(
        f"«dati_2026.xlsx», foglio «{sheet}», cella A{first + 11}: il valore «Note»"
    )


@pytest.mark.parametrize("sheet", APARTMENT_SHEETS)
def test_apartment_number_out_of_range(tmp_path: Path, sheet):
    first = FIRST_ROW[sheet]
    path = edited(tmp_path, sheet, **{f"A{first + 9}": 11})

    assert refusal(path) == (
        f"«dati_2026.xlsx», foglio «{sheet}», cella A{first + 9}: l'appartamento "
        "«11» non esiste: gli appartamenti vanno da 1 a 10. Correggere il numero "
        "o cancellare la riga, e ripetere la conversione."
    )


@pytest.mark.parametrize("sheet", APARTMENT_SHEETS)
def test_row_without_an_apartment_number(tmp_path: Path, sheet):
    first = FIRST_ROW[sheet]
    path = edited(tmp_path, sheet, **{f"A{first + 2}": None})

    assert refusal(path) == (
        f"«dati_2026.xlsx», foglio «{sheet}», cella A{first + 2}: manca il numero "
        "dell'appartamento, ma la riga contiene altri valori. Scrivere il numero "
        "dell'appartamento (da 1 a 10) o cancellare la riga, e ripetere la "
        "conversione."
    )


# -- Exercise sheet --


def test_year_is_blank(tmp_path: Path):
    assert refusal(edited(tmp_path, "Esercizio", C3=None)) == (
        "«dati_2026.xlsx», foglio «Esercizio», cella C3: l'anno dell'esercizio è "
        "vuoto. Scrivere l'anno con quattro cifre (per esempio 2026) e ripetere "
        "la conversione."
    )


@pytest.mark.parametrize("value", [25, "duemilaventisei", 20260, 2025.5])
def test_year_is_not_a_four_digit_number(tmp_path: Path, value):
    assert refusal(edited(tmp_path, "Esercizio", C3=value)) == (
        f"«dati_2026.xlsx», foglio «Esercizio», cella C3: il valore «{value}» non "
        "è un anno. Scrivere l'anno con quattro cifre (per esempio 2026) e "
        "ripetere la conversione."
    )


def test_start_date_is_blank(tmp_path: Path):
    assert refusal(edited(tmp_path, "Esercizio", C4=None)) == (
        "«dati_2026.xlsx», foglio «Esercizio», cella C4: la data di inizio "
        "dell'esercizio è vuota. Scrivere la data nel formato gg/mm/aaaa (per "
        "esempio 31/12/2025) e ripetere la conversione."
    )


def test_start_date_is_text(tmp_path: Path):
    assert refusal(edited(tmp_path, "Esercizio", C4="fine dicembre")) == (
        "«dati_2026.xlsx», foglio «Esercizio», cella C4: il valore «fine "
        "dicembre» non è una data. Scrivere la data nel formato gg/mm/aaaa (per "
        "esempio 31/12/2025) e ripetere la conversione."
    )


def test_manual_final_reading_without_a_date(tmp_path: Path):
    assert refusal(edited(tmp_path, "Esercizio", D6=26000)) == (
        "«dati_2026.xlsx», foglio «Esercizio», cella E6: manca la data della "
        "lettura finale scritta in D6. Scrivere la data nel formato gg/mm/aaaa "
        "(per esempio 31/12/2025), oppure cancellare la lettura finale, e "
        "ripetere la conversione."
    )


def test_manual_final_date_without_a_reading(tmp_path: Path):
    path = edited(tmp_path, "Esercizio", E5=datetime(2026, 1, 10))

    assert refusal(path) == (
        "«dati_2026.xlsx», foglio «Esercizio», cella D5: manca la lettura finale "
        "della data scritta in E5. Scrivere la lettura (solo il numero), "
        "oppure cancellare la data, e ripetere la conversione."
    )


def test_manual_final_reading_with_a_text_date(tmp_path: Path):
    path = edited(tmp_path, "Esercizio", D5=27500, E5="10 gennaio")

    assert refusal(path).startswith(
        "«dati_2026.xlsx», foglio «Esercizio», cella E5: il valore «10 gennaio» "
        "non è una data."
    )


# -- Static data against the report --
# The fixture report reads 28000 kWh (heating) and 25000 kWh (ACS) on the central
# meters; apartment 1 reads 5243 kWh, 31.613 m³ of hot water and 45.72 m³ of cold water.


def checked(static):
    return check_against_report(parse_fc_report(REPORT), static)


def with_previous(static, apartment: int, **changes):
    return dataclasses.replace(
        static,
        previous_readings=[
            dataclasses.replace(p, **changes) if p.apartment_number == apartment else p
            for p in static.previous_readings
        ],
    )


def test_static_data_of_the_same_exercise_is_accepted():
    checked(parse_static_data(FIXTURE))


def test_equal_readings_are_accepted():
    static = with_previous(
        parse_static_data(FIXTURE),
        1,
        heat_kwh=5243,
        water_m3=31.613,
        cold_water_m3=45.72,
    )
    static = dataclasses.replace(
        static,
        central_meters=[
            CentralMeterStart("Riscaldamento", 28000),
            CentralMeterStart("Sanitario", 25000),
        ],
    )

    checked(static)


def refusal_against_report(static) -> str:
    with pytest.raises(StaticDataError) as excinfo:
        checked(static)
    return str(excinfo.value)


ADVICE_AGAINST_REPORT = (
    "Controllare di aver scelto i dati statici dello stesso esercizio del report "
    "FC (le letture iniziali sono le letture finali dell'esercizio precedente) e "
    "ripetere la conversione."
)


def test_central_meter_below_its_start_reading():
    static = dataclasses.replace(
        parse_static_data(FIXTURE),
        central_meters=[
            CentralMeterStart("Riscaldamento", 28001),
            CentralMeterStart("Sanitario", 22000),
        ],
    )

    assert refusal_against_report(static) == (
        "Alcune letture finali sono più basse delle letture iniziali dei dati "
        "statici, e non può essere:\n"
        "- contatore centrale «Riscaldamento»: lettura finale 28000 kWh, lettura "
        "iniziale 28001 kWh\n\n" + ADVICE_AGAINST_REPORT
    )


def test_manual_final_reading_is_the_one_compared():
    low = dataclasses.replace(
        parse_static_data(FIXTURE),
        central_meters=[
            CentralMeterStart("Riscaldamento", 24000, 23000, datetime(2026, 1, 9)),
            CentralMeterStart("Sanitario", 22000),
        ],
    )
    rescued = dataclasses.replace(
        parse_static_data(FIXTURE),
        # The report's 28000 would be below 29000, the manual 29500 is not
        central_meters=[
            CentralMeterStart("Riscaldamento", 29000, 29500, datetime(2026, 1, 9)),
            CentralMeterStart("Sanitario", 22000),
        ],
    )

    assert refusal_against_report(low).splitlines()[1] == (
        "- contatore centrale «Riscaldamento»: lettura finale 23000 kWh (dai dati "
        "statici), lettura iniziale 24000 kWh"
    )
    checked(rescued)


def test_every_apartment_reading_that_went_backwards_is_named():
    static = with_previous(
        parse_static_data(FIXTURE),
        1,
        heat_kwh=5244,
        water_m3=31.7,
        cold_water_m3=45.72,
    )
    static = with_previous(static, 8, cold_water_m3=9.5)

    assert refusal_against_report(static) == (
        "Alcune letture finali sono più basse delle letture iniziali dei dati "
        "statici, e non può essere:\n"
        "- contacalorie «App, 01 Rossi Mario»: lettura finale 5243 kWh, lettura "
        "iniziale 5244 kWh\n"
        "- contatore dell'acqua calda «App, 01 Rossi Mario»: lettura finale "
        "31,613 m³, lettura iniziale 31,7 m³\n"
        "- contatore dell'acqua fredda «App, 08 Rosa Paolo»: lettura finale 8,27 "
        "m³, lettura iniziale 9,5 m³\n\n" + ADVICE_AGAINST_REPORT
    )


def test_static_data_of_the_next_exercise_is_refused():
    """Next year's start readings with this year's report: every consumption negative."""
    static = parse_static_data(FIXTURE)
    next_year = dataclasses.replace(
        static,
        previous_readings=[
            PreviousReading(p.apartment_number, p.heat_kwh + 10000, 500.0, 500.0)
            for p in static.previous_readings
        ],
    )

    message = refusal_against_report(next_year)

    assert message.count("\n- ") == 30
