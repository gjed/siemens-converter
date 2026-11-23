"""Tests for the checks run on an FC report: refusals, warnings and how they are shown.

Broken reports are built by editing the text of the committed fixture into tmp_path.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import openpyxl
import pytest

from siemens_converter.models import StaticData
from siemens_converter.parser import ReportError, parse_fc_report

FIXTURE = Path(__file__).parent / "fixtures" / "FC_report_TEST_9999_2026-01-15.xls"
STATIC_DATA = Path(__file__).parent / "fixtures" / "static_data_TEST.xlsx"

# Column indexes of the device table
DATE, STATUS, HEAT_ENERGY, WATER_VOLUME, AUX1_VOLUME = 9, 11, 15, 24, 26

DATE_CELLS = ("D28", "D31", "C76", "C126", "C178")

ADVICE = (
    "Scaricare di nuovo il report FC e ripetere la conversione. "
    "Se il messaggio compare ancora, far controllare sull'impianto "
    "i dispositivi indicati."
)


def edited_report(tmp_path, edit=lambda rows: None) -> Path:
    """Copy the fixture into tmp_path after *edit* has changed its device rows."""
    text = FIXTURE.read_text(encoding="utf-8")
    start = text.index("<tr><td>0</td>")  # first device row
    end = text.index("</table>")
    rows = [
        re.findall(r"<td>(.*?)</td>", tr) for tr in text[start:end].split("</tr>")[:-1]
    ]
    edit(rows)
    body = "".join(
        "<tr>" + "".join(f"<td>{cell}</td>" for cell in row) + "</tr>" for row in rows
    )
    path = tmp_path / FIXTURE.name
    path.write_text(text[:start] + body + text[end:], encoding="utf-8")
    return path


def device(rows, description: str, detail: str = "") -> list[str]:
    """The row of a device: central meters have no detail, apartment meters do."""
    return next(r for r in rows if r[4].startswith(description) and r[5] == detail)


def set_cell(description: str, detail: str, col: int, value: str):
    def edit(rows):
        device(rows, description, detail)[col] = value

    return edit


def refusal(path: Path) -> str:
    with pytest.raises(ReportError) as excinfo:
        parse_fc_report(path)
    return str(excinfo.value)


def test_unedited_copy_is_the_fixture(tmp_path):
    assert edited_report(tmp_path).read_text(encoding="utf-8") == FIXTURE.read_text(
        encoding="utf-8"
    )


# -- Completeness --


def test_complete_report_has_no_warnings():
    assert parse_fc_report(FIXTURE).warnings == []


def test_missing_central_meter(tmp_path):
    path = edited_report(tmp_path, lambda rows: rows.remove(device(rows, "Sanitario")))

    assert refusal(path) == (
        "il report non è completo e il riparto non è stato creato.\n\n"
        "Dispositivi mancanti:\n"
        "- contatore centrale «Sanitario»\n\n" + ADVICE
    )


def test_missing_apartment_meters(tmp_path):
    def edit(rows):
        rows.remove(device(rows, "App, 03", "Contacalorie"))
        rows.remove(device(rows, "App, 07", "Acqua calda"))

    assert refusal(edited_report(tmp_path, edit)) == (
        "il report non è completo e il riparto non è stato creato.\n\n"
        "Dispositivi mancanti:\n"
        "- contacalorie dell'appartamento 3\n"
        "- contatore dell'acqua calda dell'appartamento 7\n\n" + ADVICE
    )


def test_every_missing_device_is_listed(tmp_path):
    def edit(rows):
        del rows[:]

    message = refusal(edited_report(tmp_path, edit))

    assert message.count("\n- ") == 22
    assert "- contatore centrale «Riscaldamento»" in message
    assert "- contacalorie dell'appartamento 10" in message
    assert "- contatore dell'acqua calda dell'appartamento 1\n" in message


# -- Required readings --


@pytest.mark.parametrize(
    ("description", "detail", "col", "line"),
    [
        (
            "Riscaldamento",
            "",
            HEAT_ENERGY,
            "- contatore centrale «Riscaldamento»: la lettura dell'energia "
            "(colonna «heat_energy») è vuota",
        ),
        (
            "App, 02",
            "Contacalorie",
            HEAT_ENERGY,
            "- contacalorie «App, 02 Bianchi Luigi»: la lettura dell'energia "
            "(colonna «heat_energy») è vuota",
        ),
        (
            "App, 02",
            "Contacalorie",
            AUX1_VOLUME,
            "- contacalorie «App, 02 Bianchi Luigi»: la lettura dell'acqua fredda "
            "(colonna «aux1_volume») è vuota",
        ),
        (
            "App, 02",
            "Acqua calda",
            WATER_VOLUME,
            "- contatore dell'acqua calda «App, 02 Bianchi Luigi»: la lettura "
            "dell'acqua calda (colonna «water_volume») è vuota",
        ),
    ],
)
def test_blank_required_reading(tmp_path, description, detail, col, line):
    path = edited_report(tmp_path, set_cell(description, detail, col, " "))

    assert refusal(path) == (
        "il report non è completo e il riparto non è stato creato.\n\n"
        "Letture mancanti o non valide:\n" + line + "\n\n" + ADVICE
    )


def test_reading_that_is_not_a_number(tmp_path):
    path = edited_report(
        tmp_path, set_cell("App, 05", "Acqua calda", WATER_VOLUME, "n/d")
    )

    assert (
        "- contatore dell'acqua calda «App, 05 Gialli Marco»: la lettura "
        "dell'acqua calda (colonna «water_volume») non è un numero valido («n/d»)"
    ) in refusal(path)


def test_missing_device_and_blank_reading_in_one_message(tmp_path):
    def edit(rows):
        rows.remove(device(rows, "Riscaldamento"))
        device(rows, "App, 04", "Contacalorie")[HEAT_ENERGY] = ""

    assert refusal(edited_report(tmp_path, edit)) == (
        "il report non è completo e il riparto non è stato creato.\n\n"
        "Dispositivi mancanti:\n"
        "- contatore centrale «Riscaldamento»\n\n"
        "Letture mancanti o non valide:\n"
        "- contacalorie «App, 04 Neri Francesco»: la lettura dell'energia "
        "(colonna «heat_energy») è vuota\n\n" + ADVICE
    )


# -- Readout dates --


def test_central_meters_read_before_the_apartments(tmp_path):
    def edit(rows):
        device(rows, "Riscaldamento")[DATE] = "2026/01/03"
        device(rows, "Sanitario")[DATE] = "2026/01/03"

    assert parse_fc_report(edited_report(tmp_path, edit)).warnings == [
        "Il contatore centrale «Riscaldamento» è stato letto il 03/01/2026, "
        "12 giorni prima dei contacalorie degli appartamenti (15/01/2026). "
        "I due consumi non coprono lo stesso periodo: controllare il riparto "
        "prima di usarlo.",
        "Il contatore centrale «Sanitario» è stato letto il 03/01/2026, "
        "12 giorni prima dei contatori dell'acqua calda degli appartamenti "
        "(15/01/2026). I due consumi non coprono lo stesso periodo: controllare "
        "il riparto prima di usarlo.",
    ]


def test_central_meter_read_one_day_after(tmp_path):
    path = edited_report(tmp_path, set_cell("Sanitario", "", DATE, "2026/01/16"))

    (warning,) = parse_fc_report(path).warnings
    assert warning.startswith(
        "Il contatore centrale «Sanitario» è stato letto il 16/01/2026, "
        "1 giorno dopo i contatori dell'acqua calda degli appartamenti "
        "(15/01/2026)."
    )


def test_apartment_meter_left_behind(tmp_path):
    path = edited_report(
        tmp_path, set_cell("App, 06", "Contacalorie", DATE, "2026/01/12")
    )

    assert parse_fc_report(path).warnings == [
        "I contacalorie degli appartamenti non sono stati letti tutti lo stesso "
        "giorno. Nel riparto è indicata la data più recente (15/01/2026); letti "
        "prima: «App, 06 Blu Anna» (12/01/2026). Controllare queste letture "
        "prima di usare il riparto."
    ]


def test_each_date_cell_is_dated_from_its_own_meters(tmp_path):
    from siemens_converter.writer import write_xlsx

    def edit(rows):
        device(rows, "Riscaldamento")[DATE] = "2026/01/10"
        device(rows, "Sanitario")[DATE] = "2026/01/09"
        for row in rows:
            if row[5] == "Acqua calda":
                row[DATE] = "2026/01/14"

    out = tmp_path / "out.xlsx"
    write_xlsx(parse_fc_report(edited_report(tmp_path, edit)), out, StaticData())

    ws = openpyxl.load_workbook(out).worksheets[0]
    assert {ref: f"{ws[ref].value:%Y-%m-%d}" for ref in DATE_CELLS} == {
        "D28": "2026-01-10",
        "D31": "2026-01-09",
        "C76": "2026-01-15",
        "C126": "2026-01-14",
        "C178": "2026-01-15",
    }


# -- Communication status --


def test_central_meter_status_not_ok(tmp_path):
    path = edited_report(
        tmp_path,
        set_cell("Riscaldamento", "", STATUS, "Portata negativa | Potenza negativa"),
    )

    assert parse_fc_report(path).warnings == [
        "Lo stato di questi dispositivi nel report non è «Ok»: contatore centrale "
        "«Riscaldamento» (Portata negativa | Potenza negativa). Controllare le "
        "loro letture prima di usare il riparto."
    ]


def test_both_meters_of_an_apartment_not_ok(tmp_path):
    def edit(rows):
        device(rows, "App, 01", "Contacalorie")[STATUS] = "Errore"
        device(rows, "App, 01", "Acqua calda")[STATUS] = ""

    assert parse_fc_report(edited_report(tmp_path, edit)).warnings == [
        "Lo stato di questi dispositivi nel report non è «Ok»: contacalorie "
        "«App, 01 Rossi Mario» (Errore); contatore dell'acqua calda "
        "«App, 01 Rossi Mario» (vuoto). Controllare le loro letture prima di "
        "usare il riparto."
    ]


# -- How errors and warnings reach the user --


def run_main(
    monkeypatch, path: Path, console: bool = True, boxes: list | None = None
) -> list[tuple[str, str]]:
    """Run the CLI on *path*; return the message boxes shown as (title, text)."""
    import siemens_converter.__main__ as cli

    boxes = [] if boxes is None else boxes
    monkeypatch.setattr(sys, "argv", ["siemens-converter", str(path), str(STATIC_DATA)])
    monkeypatch.setattr(cli, "_has_console", lambda: console)
    monkeypatch.setattr(cli, "_msgbox", lambda title, text: boxes.append((title, text)))
    cli.main()
    return boxes


def test_refused_report_on_console(monkeypatch, tmp_path, capsys):
    path = edited_report(tmp_path, lambda rows: rows.remove(device(rows, "Sanitario")))

    with pytest.raises(SystemExit) as excinfo:
        run_main(monkeypatch, path)

    assert excinfo.value.code == 1
    out = capsys.readouterr().out
    assert "Errore nel report FC: il report non è completo" in out
    assert "- contatore centrale «Sanitario»" in out
    assert "Traceback" not in out
    assert list(tmp_path.glob("Riparto_*.xlsx")) == []


def test_refused_report_without_console(monkeypatch, tmp_path):
    path = edited_report(tmp_path, set_cell("Riscaldamento", "", HEAT_ENERGY, ""))
    boxes: list[tuple[str, str]] = []

    with pytest.raises(SystemExit) as excinfo:
        run_main(monkeypatch, path, console=False, boxes=boxes)

    assert excinfo.value.code == 1
    ((title, text),) = boxes
    assert title == "Errore"
    assert text.startswith("Errore nel report FC: il report non è completo")
    assert "contatore centrale «Riscaldamento»: la lettura dell'energia" in text
    assert list(tmp_path.glob("Riparto_*.xlsx")) == []


def test_warnings_on_console_before_done(monkeypatch, tmp_path, capsys):
    path = edited_report(tmp_path, set_cell("Sanitario", "", STATUS, "Errore"))

    run_main(monkeypatch, path)

    out = capsys.readouterr().out
    assert "Attenzione:\n- Lo stato di questi dispositivi" in out
    assert out.index("Attenzione:") < out.index("Fatto:")
    assert len(list(tmp_path.glob("Riparto_*.xlsx"))) == 1


def test_warnings_in_final_message_box(monkeypatch, tmp_path):
    def edit(rows):
        device(rows, "Riscaldamento")[DATE] = "2026/01/03"
        device(rows, "Sanitario")[STATUS] = "Errore"

    path = edited_report(tmp_path, edit)

    ((title, text),) = run_main(monkeypatch, path, console=False)

    assert title == "siemens-converter"
    assert text.startswith(
        "Fatto!\n\nRiparto_TE99999999_15-01-2026.xlsx\n\n"
        "Attenzione:\n"
        "- Il contatore centrale «Riscaldamento» è stato letto il 03/01/2026, "
        "12 giorni prima"
    )
    assert "\n- Lo stato di questi dispositivi nel report non è «Ok»" in text


def test_no_warnings_leaves_the_output_as_it_was(monkeypatch, tmp_path, capsys):
    path = edited_report(tmp_path)

    assert run_main(monkeypatch, path, console=False) == [
        ("siemens-converter", "Fatto!\n\nRiparto_TE99999999_15-01-2026.xlsx")
    ]

    run_main(monkeypatch, path)
    assert "Attenzione" not in capsys.readouterr().out
