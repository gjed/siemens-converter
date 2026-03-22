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


def test_dati_report_metadata(tmp_path):
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out)
    wb = openpyxl.load_workbook(out)
    ws = wb["Dati Report"]
    # Row 1: metadata labels
    assert ws["A1"].value == "File"
    assert ws["B1"].value == "Data"
    assert ws["E1"].value == "Seriale"
    # Row 2: metadata values
    assert ws["A2"].value == "FC_report.xls"
    assert ws["B2"].value == "2026-03-14"
    assert ws["E2"].value == "EV123"


def test_dati_report_column_headers(tmp_path):
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out)
    wb = openpyxl.load_workbook(out)
    ws = wb["Dati Report"]
    # Row 4: column headers (row 3 is annotations)
    assert ws["A4"].value == "Tipo"
    assert ws["B4"].value == "Appartamento"
    assert ws["C4"].value == "Energia termica"
    assert ws["E4"].value == "Volume acqua"
    assert ws["G4"].value == "Volume AFS"


def test_dati_report_water_meters(tmp_path):
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out)
    wb = openpyxl.load_workbook(out)
    ws = wb["Dati Report"]
    # Water meters start at row 5
    assert ws["A5"].value == "Acqua calda"
    assert ws["B5"].value == "App, 01 Rossi"
    assert ws["E5"].value == 31.613


def test_dati_report_heat_allocators(tmp_path):
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out)
    wb = openpyxl.load_workbook(out)
    ws = wb["Dati Report"]
    # Heat allocators after 2 water meters: row 7
    assert ws["A7"].value == "Contacalorie"
    assert ws["B7"].value == "App, 01 Rossi"
    assert ws["C7"].value == 5.243
    assert ws["D7"].value == "MWh"
    assert ws["G7"].value == 45.72


def test_dati_report_central_meters(tmp_path):
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out)
    wb = openpyxl.load_workbook(out)
    ws = wb["Dati Report"]
    # Central meters after 2 water + 2 heat = row 9
    assert ws["A9"].value == "Centrale"
    assert ws["B9"].value == "Riscaldamento"
    assert ws["C9"].value == 29213
    assert ws["D9"].value == "kWh"


def test_millesimali_names_from_report(tmp_path):
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out)
    wb = openpyxl.load_workbook(out)
    ws = wb["Tabelle millesimali"]
    # App 01 and 02 names should come from heat allocators
    assert ws["A4"].value == "App, 01 Rossi"
    assert ws["A5"].value == "App, 02 Bianchi"
    # Subalterno should remain empty (not populated by tool)
    assert ws["B4"].value is None
    # Energy values should remain empty (not populated by tool)
    assert ws["C4"].value is None
    assert ws["E4"].value is None


def test_dati_report_formatting(tmp_path):
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out)
    wb = openpyxl.load_workbook(out)
    ws = wb["Dati Report"]
    # Header cells should be bold with header fill
    assert ws["A1"].font.bold is True
    assert ws["A4"].font.bold is True
    # Green highlights on key data values
    assert ws["E5"].fill.fgColor.rgb == "FF99FFCC"  # water volume green
    assert ws["C7"].fill.fgColor.rgb == "FF99FFCC"  # heat energy green
    assert ws["G7"].fill.fgColor.rgb == "FF99FFCC"  # AFS volume green
    # Freeze panes at row 5
    assert ws.freeze_panes == "A5"


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
