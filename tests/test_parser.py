import re
import shutil
from pathlib import Path

import pytest

from siemens_converter.parser import ReportError, parse_fc_report

FIXTURE = Path(__file__).parent / "fixtures" / "FC_report_TEST_9999_2026-01-15.xls"


def test_header_parsed():
    report = parse_fc_report(FIXTURE)
    assert report.header.serial == "TE99999999"
    assert report.header.total_wired == 22


def test_central_meters_count():
    report = parse_fc_report(FIXTURE)
    assert len(report.central_meters) == 2
    descs = {m.description for m in report.central_meters}
    assert "Riscaldamento" in descs
    assert "Sanitario" in descs


def test_central_meter_values():
    report = parse_fc_report(FIXTURE)
    risc = next(m for m in report.central_meters if m.description == "Riscaldamento")
    assert isinstance(risc.heat_energy_kwh, int)
    assert risc.heat_energy_kwh > 0


def test_water_meters_count():
    report = parse_fc_report(FIXTURE)
    assert len(report.water_meters) == 10


def test_water_meter_fields():
    report = parse_fc_report(FIXTURE)
    wm = report.water_meters[0]
    assert wm.apartment_number >= 1
    assert wm.water_volume_m3 > 0


def test_heat_allocators_count():
    report = parse_fc_report(FIXTURE)
    assert len(report.heat_allocators) == 10


def test_heat_allocator_fields():
    report = parse_fc_report(FIXTURE)
    ha = report.heat_allocators[0]
    assert ha.apartment_number >= 1
    assert ha.heat_energy_mwh > 0
    assert ha.heat_energy_kwh == round(ha.heat_energy_mwh * 1000)
    assert ha.aux1_volume_m3 > 0


def test_comma_decimal_parsing():
    report = parse_fc_report(FIXTURE)
    ha = report.heat_allocators[0]
    assert isinstance(ha.heat_energy_mwh, float)


def test_apartment_numbers_sorted():
    report = parse_fc_report(FIXTURE)
    water_nums = [wm.apartment_number for wm in report.water_meters]
    heat_nums = [ha.apartment_number for ha in report.heat_allocators]
    assert water_nums == sorted(water_nums)
    assert heat_nums == sorted(heat_nums)


# -- Files that are not an FC report --

STATIC_DATA = Path(__file__).parent / "fixtures" / "static_data_TEST.xlsx"


def refusal(path: Path) -> str:
    with pytest.raises(ReportError) as excinfo:
        parse_fc_report(path)
    return str(excinfo.value)


def not_a_report(name: str, why: str) -> str:
    return (
        f"il file «{name}» non è un report FC: {why}. Scegliere il file FC_report "
        "esportato dal sistema Siemens (il suo nome comincia con FC_report) e "
        "ripetere la conversione."
    )


def edited(tmp_path: Path, edit) -> Path:
    """Copy the fixture into tmp_path after *edit* has changed its text."""
    path = tmp_path / "FC_report_modificato.xls"
    path.write_text(edit(FIXTURE.read_text(encoding="utf-8")), encoding="utf-8")
    return path


def first_rows(count: int, cells: int | None = None):
    """An edit that keeps the first *count* rows, each cut to *cells* cells."""

    def edit(text: str) -> str:
        start = text.index("<table")
        rows = re.findall(r"<tr>.*?</tr>", text[start:])[:count]
        if cells is not None:
            rows = [
                "<tr>"
                + "".join(re.findall(r"<t[dh]>.*?</t[dh]>", row)[:cells])
                + "</tr>"
                for row in rows
            ]
        return text[:start] + "<table>" + "".join(rows) + "</table>"

    return edit


def test_static_data_is_not_a_report(tmp_path):
    path = tmp_path / "dati_statici.xlsx"
    shutil.copy(STATIC_DATA, path)

    assert refusal(path) == not_a_report(
        "dati_statici.xlsx", "non contiene testo leggibile"
    )


def test_text_without_a_table(tmp_path):
    path = edited(tmp_path, lambda text: "<html><body>Ciao</body></html>")

    assert refusal(path) == not_a_report(path.name, "non contiene nessuna tabella")


def test_empty_file(tmp_path):
    path = edited(tmp_path, lambda text: "")

    assert refusal(path) == not_a_report(path.name, "non contiene nessuna tabella")


def test_too_few_rows(tmp_path):
    path = edited(tmp_path, first_rows(3))

    assert refusal(path) == not_a_report(
        path.name, "la tabella è troppo corta, mancano le righe dei dispositivi"
    )


def test_too_few_columns(tmp_path):
    path = edited(tmp_path, first_rows(30, cells=10))

    assert refusal(path) == not_a_report(
        path.name, "la tabella ha meno colonne del previsto"
    )


def test_too_few_columns_in_the_metadata_row(tmp_path):
    def edit(text: str) -> str:
        meta = re.findall(r"<tr><td>FC_report.*?</tr>", text)[0]
        return text.replace(meta, "<tr><td>FC_report</td><td>15/01/2026</td></tr>")

    path = edited(tmp_path, edit)

    assert refusal(path) == not_a_report(
        path.name, "la tabella ha meno colonne del previsto"
    )


def test_wired_total_is_not_a_number(tmp_path):
    path = edited(
        tmp_path, lambda text: text.replace("<td>22</td>", "<td>venti</td>", 1)
    )

    assert refusal(path) == not_a_report(
        path.name, "il numero dei dispositivi collegati non è un numero"
    )


def test_device_serial_number_is_not_a_number(tmp_path):
    path = edited(
        tmp_path, lambda text: text.replace("<td>22050103</td>", "<td>n/d</td>", 1)
    )

    assert refusal(path) == not_a_report(
        path.name, "il numero di serie «n/d» di un dispositivo non è un numero"
    )


def test_device_without_an_apartment_number(tmp_path):
    path = edited(
        tmp_path,
        lambda text: text.replace(
            "App, 03 Verdi Giuseppe", "Appartamento tre Verdi Giuseppe", 1
        ),
    )

    assert refusal(path) == (
        "contatore dell'acqua calda «Appartamento tre Verdi Giuseppe»: nella "
        "descrizione manca il numero dell'appartamento, che deve comparire come "
        "«App, 01». Far correggere la descrizione nel sistema Siemens, scaricare "
        "di nuovo il report FC e ripetere la conversione."
    )
