"""Tests for CLI entry point: routing, success output and every message a user can see."""

from __future__ import annotations

import importlib.util
import shutil
import sys
from pathlib import Path

import openpyxl
import pytest

import siemens_converter.__main__ as cli

FIXTURE_DIR = Path(__file__).parent / "fixtures"
FC_REPORT = FIXTURE_DIR / "FC_report_TEST_9999_2026-01-15.xls"
STATIC_DATA = FIXTURE_DIR / "static_data_TEST.xlsx"

FIND_REPORT = (
    "il file FC_report esportato dal sistema Siemens (il file .xls, il cui nome "
    "comincia con FC_report)"
)
TEMPLATE_ADVICE = (
    "Scegliere il file dei dati statici, compilato a partire da "
    "static_data_template.xlsx (si scarica dalla pagina dell'ultima versione: "
    "https://github.com/gjed/siemens-converter/releases/latest), "
    "e ripetere la conversione."
)
READING = f"Lettura del report {FC_REPORT.name} ...\n"


def run(monkeypatch, capsys, *args: Path, console: bool = True):
    """Run the CLI on *args*; return (exit code, console output, message boxes)."""
    boxes: list[tuple[str, str]] = []
    monkeypatch.setattr(sys, "argv", ["siemens-converter", *map(str, args)])
    monkeypatch.setattr(cli, "_has_console", lambda: console)
    monkeypatch.setattr(cli, "_msgbox", lambda title, text: boxes.append((title, text)))
    code = None
    try:
        cli.main()
    except SystemExit as exc:
        code = exc.code
    return code, capsys.readouterr().out, boxes


def refusal(monkeypatch, capsys, *args: Path, console: bool, progress: str = "") -> str:
    """Run the CLI, expect exit code 1 and one message; return that message."""
    code, out, boxes = run(monkeypatch, capsys, *args, console=console)
    assert code == 1
    assert "Traceback" not in out
    if console:
        assert boxes == []
        assert out.startswith(progress)
        return out.removeprefix(progress).removesuffix("\n")
    assert out == ""
    ((title, text),) = boxes
    assert title == "Errore"
    return text


def copy(src: Path, tmp_path: Path, name: str | None = None) -> Path:
    dest = tmp_path / (name or src.name)
    shutil.copy(src, dest)
    return dest


both_ways = pytest.mark.parametrize("console", [True, False], ids=["console", "box"])


# -- Usage and files that do not exist --


def test_usage_on_console(monkeypatch, capsys):
    code, out, boxes = run(monkeypatch, capsys)

    assert code == 1
    assert boxes == []
    assert out == (
        "Uso: siemens-converter <FC_report.xls> [dati_statici.xlsx]\n"
        "\n"
        "Argomenti:\n"
        "  FC_report.xls      il file esportato dal sistema Siemens (obbligatorio)\n"
        "  dati_statici.xlsx  il file dei dati statici (facoltativo)\n"
    )


@both_ways
def test_report_not_found(monkeypatch, capsys, tmp_path, console):
    missing = tmp_path / "FC_report_nuovo.xls"

    assert refusal(monkeypatch, capsys, missing, console=console) == (
        f"Il file FC_report «{missing}» non esiste. Controllare che non sia stato "
        "spostato o rinominato e ripetere la conversione."
    )


@both_ways
def test_static_file_not_found(monkeypatch, capsys, tmp_path, console):
    report = copy(FC_REPORT, tmp_path)
    missing = tmp_path / "dati_2026.xlsx"

    assert refusal(monkeypatch, capsys, report, missing, console=console) == (
        f"Il file dei dati statici «{missing}» non esiste. Controllare che non sia "
        "stato spostato o rinominato e ripetere la conversione."
    )
    assert list(tmp_path.glob("Riparto_*.xlsx")) == []


# -- Which file is which --


@both_ways
def test_only_static_data_given(monkeypatch, capsys, tmp_path, console):
    static = copy(STATIC_DATA, tmp_path)

    assert refusal(monkeypatch, capsys, static, console=console) == (
        f"Manca il file FC_report: è stato scelto solo «{static.name}», che è un "
        "file .xlsx e quindi viene usato come file dei dati statici. "
        f"Scegliere anche {FIND_REPORT} e ripetere la conversione."
    )
    assert list(tmp_path.glob("Riparto_*.xlsx")) == []


@both_ways
def test_two_xlsx_given(monkeypatch, capsys, tmp_path, console):
    first = copy(STATIC_DATA, tmp_path)
    second = copy(STATIC_DATA, tmp_path, "dati_2025.xlsx")

    assert refusal(monkeypatch, capsys, first, second, console=console) == (
        f"Sono stati scelti due file .xlsx, «{first.name}» e «{second.name}», ma "
        "ne serve uno solo: il file dei dati statici. Manca il file FC_report. "
        f"Scegliere {FIND_REPORT}, insieme a un solo file dei dati statici, e "
        "ripetere la conversione."
    )


@both_ways
def test_two_files_that_are_not_xlsx_given(monkeypatch, capsys, tmp_path, console):
    first = copy(FC_REPORT, tmp_path)
    second = copy(FC_REPORT, tmp_path, "FC_report_vecchio.xls")

    assert refusal(monkeypatch, capsys, first, second, console=console) == (
        f"Sono stati scelti due file che non sono .xlsx, «{first.name}» e "
        f"«{second.name}», ma ne serve uno solo: il file FC_report. Il file dei "
        "dati statici deve essere un file .xlsx. Scegliere il solo file "
        "FC_report, oppure il file FC_report insieme al file dei dati statici, e "
        "ripetere la conversione."
    )
    assert list(tmp_path.glob("Riparto_*.xlsx")) == []


def test_dialogs_route_files_like_drag_and_drop(monkeypatch, capsys, tmp_path):
    # run() starts without file arguments, so the no-console build opens the dialogs
    """An .xlsx picked in the first dialog is the static data, not the report."""
    static = copy(STATIC_DATA, tmp_path)
    picks = iter([str(static), None])
    monkeypatch.setattr(cli, "_win_open_file_dialog", lambda title, types: next(picks))

    code, out, boxes = run(monkeypatch, capsys, console=False)

    assert code == 1
    ((title, text),) = boxes
    assert title == "Errore"
    assert text.startswith(
        f"Manca il file FC_report: è stato scelto solo «{static.name}»"
    )


def test_dialog_cannot_open(monkeypatch):
    boxes: list[tuple[str, str]] = []
    monkeypatch.setattr(cli, "_msgbox", lambda title, text: boxes.append((title, text)))
    monkeypatch.setattr(sys, "platform", "linux")  # there is no Windows API here

    assert cli._win_open_file_dialog("titolo", "filtro") is None

    assert boxes == [
        (
            "Errore",
            "Impossibile aprire la finestra di selezione dei file. "
            "Trascinare i file da convertire sopra l'icona del programma.",
        )
    ]


# -- Files that are not what they should be --


@both_ways
def test_report_is_not_text(monkeypatch, capsys, tmp_path, console):
    not_a_report = copy(STATIC_DATA, tmp_path, "FC_report_sbagliato.xls")

    assert refusal(
        monkeypatch,
        capsys,
        not_a_report,
        console=console,
        progress=f"Lettura del report {not_a_report.name} ...\n",
    ) == (
        "Errore nel report FC: il file «FC_report_sbagliato.xls» non è un report "
        "FC: non contiene testo leggibile. Scegliere il file FC_report esportato "
        "dal sistema Siemens (il suo nome comincia con FC_report) e ripetere la "
        "conversione."
    )
    assert list(tmp_path.glob("Riparto_*.xlsx")) == []


@both_ways
def test_static_data_is_not_a_workbook(monkeypatch, capsys, tmp_path, console):
    report = copy(FC_REPORT, tmp_path)
    not_static = copy(FC_REPORT, tmp_path, "dati_2026.xlsx")

    assert refusal(
        monkeypatch,
        capsys,
        report,
        not_static,
        console=console,
        progress=READING + "Lettura dei dati statici da dati_2026.xlsx ...\n",
    ) == (
        "Errore nel file dati statici: «dati_2026.xlsx» non è un file dei dati "
        "statici: non si apre come file Excel (.xlsx). " + TEMPLATE_ADVICE
    )
    assert list(tmp_path.glob("Riparto_*.xlsx")) == []


@both_ways
def test_static_data_value_is_not_a_number(monkeypatch, capsys, tmp_path, console):
    report = copy(FC_REPORT, tmp_path)
    wb = openpyxl.load_workbook(STATIC_DATA)
    wb["Costi"]["B4"] = "2.141 euro"
    static = tmp_path / "dati_2026.xlsx"
    wb.save(static)

    assert refusal(
        monkeypatch,
        capsys,
        report,
        static,
        console=console,
        progress=READING + "Lettura dei dati statici da dati_2026.xlsx ...\n",
    ) == (
        "Errore nel file dati statici: «dati_2026.xlsx», foglio «Costi», cella "
        "B4: il valore «2.141 euro» non è un numero. Scrivere solo il numero, "
        "senza unità di misura e senza simboli, e ripetere la conversione."
    )


@both_ways
def test_report_is_a_folder(monkeypatch, capsys, tmp_path, console):
    folder = tmp_path / "FC_report_cartella.xls"
    folder.mkdir()

    assert refusal(
        monkeypatch,
        capsys,
        folder,
        console=console,
        progress=f"Lettura del report {folder.name} ...\n",
    ) == (
        "Errore nel report FC: il file «FC_report_cartella.xls» non si può aprire "
        "per leggerlo. Controllare che sia un file e non una cartella, che non "
        "sia aperto in un altro programma e che si possa leggere, e ripetere la "
        "conversione."
    )


@both_ways
def test_output_cannot_be_written(monkeypatch, capsys, tmp_path, console):
    report = copy(FC_REPORT, tmp_path)
    output = tmp_path / "Riparto_TE99999999_15-01-2026.xlsx"
    output.mkdir()  # a file of that name cannot replace a folder

    assert refusal(
        monkeypatch,
        capsys,
        report,
        console=console,
        progress=READING + f"Scrittura di {output.name} ...\n",
    ) == (
        f"Impossibile creare il file «{output.name}» nella cartella {tmp_path}. "
        "Se è aperto in Excel, chiuderlo; altrimenti controllare di poter "
        "scrivere in quella cartella. Poi ripetere la conversione."
    )


# -- Successful runs --


def test_console_progress_and_final_line(monkeypatch, capsys, tmp_path):
    report = copy(FC_REPORT, tmp_path)
    static = copy(STATIC_DATA, tmp_path)
    output = tmp_path / "Riparto_TE99999999_15-01-2026.xlsx"

    code, out, boxes = run(monkeypatch, capsys, report, static)

    assert code is None
    assert boxes == []
    assert out == (
        f"Lettura del report {report.name} ...\n"
        f"Lettura dei dati statici da {static.name} ...\n"
        f"Scrittura di {output.name} ...\n"
        f"Fatto: {output}\n"
    )


def test_windows_console_waits_for_enter(monkeypatch, capsys, tmp_path):
    prompts: list[str] = []
    monkeypatch.setattr(sys, "platform", "win32")
    monkeypatch.setattr("builtins.input", prompts.append)

    run(monkeypatch, capsys, copy(FC_REPORT, tmp_path))

    assert prompts == ["\nPremere Invio per chiudere..."]


def test_final_message_box(monkeypatch, capsys, tmp_path):
    code, out, boxes = run(
        monkeypatch, capsys, copy(FC_REPORT, tmp_path), console=False
    )

    assert code is None
    assert out == ""
    assert boxes == [
        ("siemens-converter", "Fatto!\n\nRiparto_TE99999999_15-01-2026.xlsx")
    ]


# -- Last resort --


def load_entry_point():
    path = Path(__file__).parent.parent / "scripts" / "pyinstaller_entry.py"
    spec = importlib.util.spec_from_file_location("pyinstaller_entry", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


SENTENCE = (
    "La conversione non è riuscita per un errore imprevisto del programma. "
    "Inviare i dettagli tecnici qui sotto a chi ha fornito il programma."
)


def test_last_resort_on_console(monkeypatch, capsys):
    monkeypatch.setattr(sys, "platform", "linux")

    load_entry_point().report_failure("Traceback: KeyError")

    assert capsys.readouterr().err == (
        f"{SENTENCE}\n\nDettagli tecnici:\nTraceback: KeyError\n"
    )


def test_last_resort_in_message_box(monkeypatch):
    import ctypes

    shown: list[tuple] = []

    class User32:
        @staticmethod
        def MessageBoxW(*args):
            shown.append(args)

    monkeypatch.setattr(sys, "platform", "win32")
    monkeypatch.setattr(sys, "stdout", None)
    monkeypatch.setattr(
        ctypes, "windll", type("W", (), {"user32": User32}), raising=False
    )

    load_entry_point().report_failure("Traceback: KeyError")

    assert shown == [
        (
            0,
            f"{SENTENCE}\n\nDettagli tecnici:\nTraceback: KeyError",
            "siemens-converter - Errore",
            0,
        )
    ]


def test_main_success_single_file(monkeypatch, tmp_path):
    """Single FC_report argument — backward compatible."""
    input_file = tmp_path / FC_REPORT.name
    shutil.copy(FC_REPORT, input_file)

    monkeypatch.setattr(sys, "argv", ["siemens-converter", str(input_file)])
    from siemens_converter.__main__ import main

    main()

    outputs = list(tmp_path.glob("Riparto_*.xlsx"))
    assert len(outputs) == 1

    # Inquilini tenant column should be empty (no static data)
    wb = openpyxl.load_workbook(outputs[0])
    ws = wb["Inquilini"]
    assert ws["B2"].value is None


def test_main_success_two_files(monkeypatch, tmp_path):
    """Both FC_report and static data arguments."""
    input_file = tmp_path / FC_REPORT.name
    shutil.copy(FC_REPORT, input_file)
    static_file = tmp_path / STATIC_DATA.name
    shutil.copy(STATIC_DATA, static_file)

    monkeypatch.setattr(
        sys,
        "argv",
        ["siemens-converter", str(input_file), str(static_file)],
    )
    from siemens_converter.__main__ import main

    main()

    outputs = list(tmp_path.glob("Riparto_*.xlsx"))
    assert len(outputs) == 1

    # Inquilini tenant column should be populated
    wb = openpyxl.load_workbook(outputs[0])
    ws = wb["Inquilini"]
    assert ws["B2"].value == "Bianchi Anna"  # Apt 1 tenant from static data

    # Millesimali should be populated
    ws_mill = wb["Tabelle millesimali"]
    assert ws_mill.cell(row=4, column=2).value == 20  # subalterno apt 1

    # Costs should be populated
    ws_tab = wb["Tabella_2026"]
    assert ws_tab.cell(row=3, column=5).value == 263.44  # Energia elettrica


def test_main_two_files_reversed_order(monkeypatch, tmp_path):
    """Drag-and-drop may pass the .xlsx first; files are routed by suffix."""
    input_file = tmp_path / FC_REPORT.name
    shutil.copy(FC_REPORT, input_file)
    static_file = tmp_path / STATIC_DATA.name
    shutil.copy(STATIC_DATA, static_file)

    monkeypatch.setattr(
        sys,
        "argv",
        ["siemens-converter", str(static_file), str(input_file)],
    )
    from siemens_converter.__main__ import main

    main()

    outputs = list(tmp_path.glob("Riparto_*.xlsx"))
    assert len(outputs) == 1
    wb = openpyxl.load_workbook(outputs[0])
    assert wb["Inquilini"]["B2"].value == "Bianchi Anna"
