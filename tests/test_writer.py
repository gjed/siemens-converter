"""Tests for XLSX writer."""

from __future__ import annotations

import openpyxl
import pytest
from pathlib import Path

from siemens_converter.writer import write_xlsx
from siemens_converter.models import (
    ApartmentInfo,
    CentralMeter,
    CostItem,
    HeatAllocator,
    MeterReading,
    Millesimali,
    ParsedReport,
    PreviousReading,
    ReportHeader,
    StaticData,
    WaterMeter,
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


def test_five_sheets(tmp_path):
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out)
    wb = openpyxl.load_workbook(out)
    assert len(wb.sheetnames) == 5
    assert "Inquilini" in wb.sheetnames
    assert "Dati Report" in wb.sheetnames


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
    # Row 1: metadata labels in FC_report column positions
    assert ws["A1"].value == "Nome File"
    assert ws["B1"].value == "Data Report"
    assert ws["E1"].value == "Versione firmware"
    assert ws["H1"].value == "Numero di serie"
    # Row 1: group headers over data columns
    assert ws.cell(row=1, column=16).value == "Riparto"  # col P
    assert ws.cell(row=1, column=25).value == "Riparto"  # col Y
    assert ws.cell(row=1, column=27).value == "Conti separati a parte"  # col AA
    # Row 2: metadata values
    assert ws["A2"].value == "FC_report.xls"
    assert ws["H2"].value == "EV123"
    # Row 3: annotation sub-labels
    assert ws.cell(row=3, column=16).value == "energia termica"
    assert ws.cell(row=3, column=25).value == "volume acqua sanitaria"
    assert ws.cell(row=3, column=27).value == "volume acqua fredda"


def test_millesimali_names_reference_inquilini(tmp_path):
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out)
    wb = openpyxl.load_workbook(out)
    ws = wb["Tabelle millesimali"]
    # Names come from Inquilini sheet via formulas
    assert ws["A4"].value == "=Inquilini!A2"  # apt 1
    assert ws["A5"].value == "=Inquilini!A3"  # apt 2
    assert ws["B4"].value is None  # subalterno empty
    assert ws["C4"].value is None  # energy empty


def test_inquilini_sheet(tmp_path):
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out)
    wb = openpyxl.load_workbook(out)
    ws = wb["Inquilini"]
    # Header
    assert ws["A1"].value == "Appartamento"
    assert ws["B1"].value == "Inquilino"
    # Apartment rows from report
    assert ws["A2"].value == "App, 01 Rossi"
    assert ws["A3"].value == "App, 02 Bianchi"
    # Tenant column empty (openpyxl saves "" as None)
    assert ws["B2"].value is None
    assert ws["B3"].value is None


def test_ripartizione_references_inquilini(tmp_path):
    """Apartment cells in Ripartizione should reference Inquilini sheet."""
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out)
    wb = openpyxl.load_workbook(out)
    ws = wb.worksheets[0]
    # Summary row 3 (apt 1) = owner -> Inquilini!A2
    assert ws["A3"].value == "=Inquilini!A2"
    # Detail row 102 (apt 1 owner) -> Inquilini!A2
    assert ws["A102"].value == "=Inquilini!A2"
    # Detail row 103 (apt 1 tenant) -> Inquilini!B2
    assert ws["A103"].value == "=Inquilini!B2"


def test_inquilini_has_fc_report_names(tmp_path):
    """Inquilini column A should have the FC_report apartment descriptions."""
    from siemens_converter.parser import parse_fc_report

    fixture = Path(__file__).parent / "fixtures" / "FC_report_TEST_9999_2026-01-15.xls"
    report = parse_fc_report(fixture)
    out = tmp_path / "output.xlsx"
    write_xlsx(report, out)

    wb = openpyxl.load_workbook(out)
    ws = wb["Inquilini"]
    assert "Rossi" in ws["A2"].value  # apt 1
    assert "Bianchi" in ws["A3"].value  # apt 2
    # Tenant column B still empty
    assert ws["B2"].value is None


def test_dati_report_full_fc_data(tmp_path):
    """Parse fixture -> verify Dati Report has all 38 columns from FC_report."""
    from siemens_converter.parser import parse_fc_report

    fixture = Path(__file__).parent / "fixtures" / "FC_report_TEST_9999_2026-01-15.xls"
    report = parse_fc_report(fixture)
    out = tmp_path / "output.xlsx"
    write_xlsx(report, out)

    wb = openpyxl.load_workbook(out)
    ws = wb["Dati Report"]

    # Row 4 should have all FC_report column headers
    assert ws.cell(row=4, column=1).value == "count"
    assert ws.cell(row=4, column=5).value == "device_description"  # col E
    assert ws.cell(row=4, column=16).value == "heat_energy"  # col P
    assert ws.cell(row=4, column=25).value == "water_volume"  # col Y
    assert ws.cell(row=4, column=27).value == "aux1_volume"  # col AA

    # Data starts at row 5 -- water meters first (10), then heat (10), then central (2)
    # First water meter
    assert ws.cell(row=5, column=6).value == "Acqua calda"  # device_detail = col F
    # First heat allocator (after 10 water meters)
    assert ws.cell(row=15, column=6).value == "Contacalorie"
    # Central meters at end (after 10 water + 10 heat = row 25)
    desc_25 = ws.cell(row=25, column=5).value
    assert desc_25 in ("Riscaldamento", "Sanitario")


def test_dati_report_hidden_columns(tmp_path):
    """Irrelevant columns should be hidden."""
    from siemens_converter.parser import parse_fc_report

    fixture = Path(__file__).parent / "fixtures" / "FC_report_TEST_9999_2026-01-15.xls"
    report = parse_fc_report(fixture)
    out = tmp_path / "output.xlsx"
    write_xlsx(report, out)

    wb = openpyxl.load_workbook(out)
    ws = wb["Dati Report"]

    # Visible columns: A(1), B(2), E(5), G(7), M(13), P(16), Q(17), R(18), Y(25), Z(26), AA(27), AB(28)
    assert ws.column_dimensions["A"].hidden is False
    assert ws.column_dimensions["E"].hidden is False
    assert ws.column_dimensions["P"].hidden is False
    # Hidden columns
    assert ws.column_dimensions["C"].hidden is True  # device_serial_number
    assert ws.column_dimensions["D"].hidden is True  # name_device
    assert ws.column_dimensions["H"].hidden is True  # wired/wireless


def test_dati_report_formatting(tmp_path):
    """Verify formatting: green highlights, alternating rows, freeze panes."""
    from siemens_converter.parser import parse_fc_report

    fixture = Path(__file__).parent / "fixtures" / "FC_report_TEST_9999_2026-01-15.xls"
    report = parse_fc_report(fixture)
    out = tmp_path / "output.xlsx"
    write_xlsx(report, out)

    wb = openpyxl.load_workbook(out)
    ws = wb["Dati Report"]

    # Header row bold
    assert ws["A1"].font.bold is True
    assert ws.cell(row=4, column=1).font.bold is True

    # Green highlight on water_volume (col Y=25) for first water meter (row 5)
    water_cell = ws.cell(row=5, column=25)
    assert water_cell.fill.fgColor.rgb == "FF99FFCC"

    # Green highlight on heat_energy (col P=16) for first heat allocator (row 15)
    heat_cell = ws.cell(row=15, column=16)
    assert heat_cell.fill.fgColor.rgb == "FF99FFCC"

    # Alternating row fill on even data rows
    alt_cell = ws.cell(row=6, column=1)  # second data row
    assert alt_cell.fill.fgColor.rgb == "FFE8E8E8"

    # Freeze panes
    assert ws.freeze_panes == "A5"

    # Row heights
    assert ws.row_dimensions[5].height == 27


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


# -- Static data integration tests --


def _make_static_data():
    """Build a StaticData matching the 2-apartment test report."""
    return StaticData(
        apartments=[
            ApartmentInfo(1, "Rossi Mario", "Bianchi Anna"),
            ApartmentInfo(2, "Verdi Giuseppe", "Neri Luigi"),
        ],
        millesimali=[
            Millesimali(1, 20, 1344.03, 1101.36),
            Millesimali(2, 21, 2595.68, 1297.82),
        ],
        costs=[
            CostItem("Energia elettrica", 263.44),
            CostItem("Gas metano", 2141.0),
            CostItem("Acqua condominio", 0.0),
            CostItem("Conduzione e manutenzione", 280.0),
            CostItem("Contabilizzazione", 68.2),
            CostItem("Acqua sanitaria manutenzione", 0.0),
        ],
        meters=[
            MeterReading("Energia elettrica CT", "kWh", 0, 1086),
            MeterReading("Gas metano CT", "mc", 0, 1718),
            MeterReading("Acqua generale", "mc", 0, 680),
        ],
        previous_readings=[
            PreviousReading(1, 4048, 25.5, 10.2),
            PreviousReading(2, 4727, 30.1, 12.5),
        ],
    )


def test_static_data_tenant_names(tmp_path):
    """Inquilini column B should have tenant names when static data provided."""
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out, static_data=_make_static_data())
    wb = openpyxl.load_workbook(out)
    ws = wb["Inquilini"]
    assert ws["B2"].value == "Bianchi Anna"
    assert ws["B3"].value == "Neri Luigi"


def test_static_data_millesimali(tmp_path):
    """Tabelle millesimali should have energy values and subalterno."""
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out, static_data=_make_static_data())
    wb = openpyxl.load_workbook(out)
    ws = wb["Tabelle millesimali"]
    # Apartment 1 -> row 4
    assert ws.cell(row=4, column=2).value == 20  # subalterno
    assert ws.cell(row=4, column=3).value == 1344.03  # heat energy kWh
    assert ws.cell(row=4, column=5).value == 1101.36  # water energy kWh
    # Apartment 2 -> row 5
    assert ws.cell(row=5, column=2).value == 21
    assert ws.cell(row=5, column=3).value == 2595.68


def test_static_data_costs(tmp_path):
    """Tabella_2026 should have cost amounts in E3-E8."""
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out, static_data=_make_static_data())
    wb = openpyxl.load_workbook(out)
    ws = wb["Tabella_2026"]
    assert ws.cell(row=3, column=5).value == 263.44  # Energia elettrica
    assert ws.cell(row=4, column=5).value == 2141.0  # Gas metano
    assert ws.cell(row=5, column=5).value == 0.0  # Acqua
    assert ws.cell(row=6, column=5).value == 280.0  # Conduzione
    assert ws.cell(row=7, column=5).value == 68.2  # Contabilizzazione
    assert ws.cell(row=8, column=5).value == 0.0  # Acqua sanitaria


def test_static_data_meter_readings(tmp_path):
    """Tabella_2026 should have meter readings in F20-G22."""
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out, static_data=_make_static_data())
    wb = openpyxl.load_workbook(out)
    ws = wb["Tabella_2026"]
    assert ws.cell(row=20, column=6).value == 0  # elettrica initial
    assert ws.cell(row=20, column=7).value == 1086  # elettrica final
    assert ws.cell(row=21, column=6).value == 0  # gas initial
    assert ws.cell(row=21, column=7).value == 1718  # gas final
    assert ws.cell(row=22, column=6).value == 0  # acqua initial
    assert ws.cell(row=22, column=7).value == 680  # acqua final


def test_static_data_previous_readings(tmp_path):
    """Ripartizione previous readings in column B."""
    out = tmp_path / "output.xlsx"
    write_xlsx(_make_report(), out, static_data=_make_static_data())
    wb = openpyxl.load_workbook(out)
    ws = wb.worksheets[0]
    # Apt 1: heat row 78, water row 130, AFS row 182
    assert ws.cell(row=78, column=2).value == 4048  # heat prev
    assert ws.cell(row=130, column=2).value == 25.5  # water prev
    assert ws.cell(row=182, column=2).value == 10.2  # AFS prev
    # Apt 2: heat row 80, water row 132, AFS row 184
    assert ws.cell(row=80, column=2).value == 4727
    assert ws.cell(row=132, column=2).value == 30.1
    assert ws.cell(row=184, column=2).value == 12.5


def test_without_static_data_unchanged(tmp_path):
    """Without static data, output is identical to current behavior."""
    out_no_static = tmp_path / "no_static.xlsx"
    out_with_none = tmp_path / "with_none.xlsx"
    report = _make_report()

    write_xlsx(report, out_no_static)
    write_xlsx(report, out_with_none, static_data=None)

    wb1 = openpyxl.load_workbook(out_no_static)
    wb2 = openpyxl.load_workbook(out_with_none)

    # Same sheets
    assert wb1.sheetnames == wb2.sheetnames

    # Inquilini tenant column still empty
    ws = wb1["Inquilini"]
    assert ws["B2"].value is None

    # Tabelle millesimali energy values still empty
    ws = wb1["Tabelle millesimali"]
    assert ws.cell(row=4, column=2).value is None  # subalterno
    assert ws.cell(row=4, column=3).value is None  # heat energy
