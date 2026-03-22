"""Tests for XLSX writer."""

from __future__ import annotations

import openpyxl
import pytest
from pathlib import Path

from siemens_converter.writer import write_xlsx
from siemens_converter.models import (
    ReportHeader,
    CentralMeter,
    WaterMeter,
    HeatAllocator,
    ParsedReport,
)


def _make_report():
    """Build a ParsedReport with 2 apartments for testing."""
    header = ReportHeader(
        "FC_report.xls",
        "2026-03-14",
        "16:45:03",
        "-",
        "EV123",
        "3.93",
        12,
    )
    central = [
        CentralMeter("Riscaldamento", 29213, 71714666, "2026/03/02"),
        CentralMeter("Sanitario", 31310, 71731475, "2026/03/02"),
    ]
    water = [
        WaterMeter("App, 01 Rossi", 1, 31.613, 23005747, "2026/03/02"),
        WaterMeter("App, 02 Bianchi", 2, 39.18, 22079950, "2026/03/02"),
    ]
    heat = [
        HeatAllocator("App, 01 Rossi", 1, 5.243, 45.72, 71714345, "2026/03/02"),
        HeatAllocator("App, 02 Bianchi", 2, 6.26, 59.69, 71714369, "2026/03/02"),
    ]
    return ParsedReport(header, central, water, heat)


def test_output_file_created(tmp_path):
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out)
    assert out.exists()


def test_central_meter_riscaldamento(tmp_path):
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out)
    wb = openpyxl.load_workbook(out)
    ws = wb.worksheets[0]
    assert ws["C28"].value == 29213


def test_central_meter_sanitario(tmp_path):
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out)
    wb = openpyxl.load_workbook(out)
    ws = wb.worksheets[0]
    assert ws["C31"].value == 31310


def test_reading_date(tmp_path):
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out)
    wb = openpyxl.load_workbook(out)
    ws = wb.worksheets[0]
    assert ws["D28"].value is not None


def test_heat_allocator_values(tmp_path):
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out)
    wb = openpyxl.load_workbook(out)
    ws = wb.worksheets[0]
    assert ws["C78"].value == 5243  # 5.243 MWh * 1000
    assert ws["C80"].value == 6260  # 6.26 MWh * 1000


def test_water_meter_values(tmp_path):
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out)
    wb = openpyxl.load_workbook(out)
    ws = wb.worksheets[0]
    assert ws["C130"].value == 31.613
    assert ws["C132"].value == 39.18


def test_afs_values(tmp_path):
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out)
    wb = openpyxl.load_workbook(out)
    ws = wb.worksheets[0]
    assert ws["C182"].value == 45.72
    assert ws["C184"].value == 59.69


def test_four_sheets(tmp_path):
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out)
    wb = openpyxl.load_workbook(out)
    assert len(wb.sheetnames) == 4


def test_dati_report_sheet_exists(tmp_path):
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out)
    wb = openpyxl.load_workbook(out)
    assert "Dati Report" in wb.sheetnames


def test_dati_report_header_row(tmp_path):
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out)
    wb = openpyxl.load_workbook(out)
    ws = wb["Dati Report"]
    # Row 1: metadata
    assert ws["A1"].value == "File"
    assert ws["B1"].value == "FC_report.xls"
    assert ws["C1"].value == "Data"
    assert ws["D1"].value == "2026-03-14"


def test_dati_report_column_headers(tmp_path):
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out)
    wb = openpyxl.load_workbook(out)
    ws = wb["Dati Report"]
    # Row 3: column headers
    assert ws["A3"].value == "Tipo"
    assert ws["B3"].value == "Appartamento"
    assert ws["C3"].value == "Energia termica"
    assert ws["D3"].value == "Unita"
    assert ws["E3"].value == "Volume acqua"
    assert ws["F3"].value == "Unita"
    assert ws["G3"].value == "Volume AFS"
    assert ws["H3"].value == "Unita"


def test_dati_report_water_meters(tmp_path):
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out)
    wb = openpyxl.load_workbook(out)
    ws = wb["Dati Report"]
    # Water meters start at row 4
    assert ws["A4"].value == "Acqua calda"
    assert ws["B4"].value == "App, 01 Rossi"
    assert ws["E4"].value == 31.613
    assert ws["F4"].value == "m3"


def test_dati_report_heat_allocators(tmp_path):
    out = tmp_path / "output.xlsx"
    report = _make_report()
    out = tmp_path / "output.xlsx"
    write_xlsx(report, out)
    wb = openpyxl.load_workbook(out)
    ws = wb["Dati Report"]
    # Heat allocators after water meters: row 4 + 2 water = row 6
    assert ws["A6"].value == "Contacalorie"
    assert ws["B6"].value == "App, 01 Rossi"
    assert ws["C6"].value == 5.243
    assert ws["D6"].value == "MWh"
    assert ws["G6"].value == 45.72
    assert ws["H6"].value == "m3"


def test_dati_report_central_meters(tmp_path):
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out)
    wb = openpyxl.load_workbook(out)
    ws = wb["Dati Report"]
    # Central meters after water (2) + heat (2) = row 8
    assert ws["A8"].value == "Centrale"
    assert ws["B8"].value == "Riscaldamento"
    assert ws["C8"].value == 29213
    assert ws["D8"].value == "kWh"


def test_full_pipeline(tmp_path):
    """Parse fixture -> write XLSX -> verify key values."""
    from siemens_converter.parser import parse_fc_report

    fixture = Path(__file__).parent / "fixtures" / "FC_report_TEST_9999_2026-01-15.xls"
    report = parse_fc_report(fixture)
    out = tmp_path / "output.xlsx"
    write_xlsx(report, out)

    wb = openpyxl.load_workbook(out)
    ws = wb.worksheets[0]

    # Central meters populated
    assert ws["C28"].value is not None  # Riscaldamento
    assert ws["C31"].value is not None  # Sanitario
    assert isinstance(ws["C28"].value, int)
    assert isinstance(ws["C31"].value, int)

    # All 10 apartments should have heat readings
    heat_rows = [78, 80, 82, 84, 86, 88, 90, 92, 94, 95]
    for r in heat_rows:
        assert ws.cell(row=r, column=3).value is not None, (
            f"C{r} should have heat value"
        )

    # All 10 apartments should have water readings
    water_rows = [130, 132, 134, 136, 138, 140, 142, 144, 146, 147]
    for r in water_rows:
        assert ws.cell(row=r, column=3).value is not None, (
            f"C{r} should have water value"
        )

    # All 10 apartments should have AFS readings
    afs_rows = [182, 184, 186, 188, 190, 192, 194, 196, 198, 199]
    for r in afs_rows:
        assert ws.cell(row=r, column=3).value is not None, f"C{r} should have AFS value"

    # Dates should be set
    assert ws["D28"].value is not None
    assert ws["C76"].value is not None
