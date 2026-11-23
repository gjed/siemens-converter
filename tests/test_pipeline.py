"""Tests for the conversion sequence the CLI and the web page share."""

from __future__ import annotations

import shutil
from pathlib import Path

import openpyxl
import pytest

from siemens_converter.pipeline import ConversionError, convert, route_files

FIXTURE_DIR = Path(__file__).parent / "fixtures"
FC_REPORT = FIXTURE_DIR / "FC_report_TEST_9999_2026-01-15.xls"
STATIC_DATA = FIXTURE_DIR / "static_data_TEST.xlsx"


def test_route_files_in_either_order():
    assert route_files([FC_REPORT, STATIC_DATA]) == (FC_REPORT, STATIC_DATA)
    assert route_files([STATIC_DATA, FC_REPORT]) == (FC_REPORT, STATIC_DATA)


@pytest.mark.parametrize(
    ("names", "start"),
    [
        (["dati.xlsx"], "Manca il file FC_report: è stato scelto solo «dati.xlsx»"),
        (["a.xlsx", "b.xlsx"], "Sono stati scelti due file .xlsx, «a.xlsx» e «b.xlsx»"),
        (["a.xls", "b.xls"], "Sono stati scelti due file che non sono .xlsx, «a.xls»"),
        (["FC_report.xls"], "Manca il file dei dati statici: è stato scelto solo"),
    ],
)
def test_route_files_refuses_anything_but_one_of_each(names, start):
    with pytest.raises(ConversionError) as caught:
        route_files([Path(n) for n in names])
    assert str(caught.value).startswith(start)


def test_convert_writes_the_riparto_and_reports_progress(tmp_path):
    lines: list[str] = []

    result = convert(FC_REPORT, STATIC_DATA, tmp_path, lines.append)

    assert result.output_path == tmp_path / "Riparto_TE99999999_15-01-2026.xlsx"
    assert openpyxl.load_workbook(result.output_path).sheetnames
    assert lines == [
        f"Lettura del report {FC_REPORT.name} ...",
        f"Lettura dei dati statici da {STATIC_DATA.name} ...",
        f"Scrittura di {result.output_path.name} ...",
    ]


def test_convert_prefixes_a_report_error(tmp_path):
    bad = tmp_path / "FC_report_finto.xls"
    bad.write_text("ciao")

    with pytest.raises(ConversionError, match="^Errore nel report FC: "):
        convert(bad, STATIC_DATA, tmp_path)


def test_convert_prefixes_a_static_data_error(tmp_path):
    bad = tmp_path / "dati.xlsx"
    bad.write_text("ciao")

    with pytest.raises(ConversionError, match="^Errore nel file dati statici: "):
        convert(FC_REPORT, bad, tmp_path)


def test_convert_refuses_an_output_that_cannot_be_written(tmp_path):
    (tmp_path / "Riparto_TE99999999_15-01-2026.xlsx").mkdir()

    with pytest.raises(ConversionError, match="^Impossibile creare il file «Riparto_"):
        convert(FC_REPORT, STATIC_DATA, tmp_path)


def test_convert_does_not_modify_its_inputs(tmp_path):
    report = Path(shutil.copy(FC_REPORT, tmp_path))
    static = Path(shutil.copy(STATIC_DATA, tmp_path))
    before = (report.read_bytes(), static.read_bytes())

    convert(report, static, tmp_path)

    assert (report.read_bytes(), static.read_bytes()) == before
