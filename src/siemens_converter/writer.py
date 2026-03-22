"""XLSX workbook generation from parsed FC_report data."""

from __future__ import annotations

import shutil
from datetime import datetime
from importlib.resources import files
from pathlib import Path

import openpyxl

from siemens_converter.models import ParsedReport

# Row mapping: apartment number -> {heat_row, water_row, afs_row} in Ripartizione sheet
APT_ROWS: dict[int, dict[str, int]] = {
    1: {"heat": 78, "water": 130, "afs": 182},
    2: {"heat": 80, "water": 132, "afs": 184},
    3: {"heat": 82, "water": 134, "afs": 186},
    4: {"heat": 84, "water": 136, "afs": 188},
    5: {"heat": 86, "water": 138, "afs": 190},
    6: {"heat": 88, "water": 140, "afs": 192},
    7: {"heat": 90, "water": 142, "afs": 194},
    8: {"heat": 92, "water": 144, "afs": 196},
    9: {"heat": 94, "water": 146, "afs": 198},
    10: {"heat": 95, "water": 147, "afs": 199},
}


def write_xlsx(report: ParsedReport, output_path: Path) -> None:
    """Write parsed report data into a copy of the template XLSX."""
    template = files("siemens_converter").joinpath("template.xlsx")
    shutil.copy(str(template), str(output_path))

    wb = openpyxl.load_workbook(output_path)
    ws = wb.worksheets[0]  # Ripartizione sheet

    reading_date = _parse_reading_date(report)

    # Central meter readings
    for meter in report.central_meters:
        if meter.description == "Riscaldamento":
            ws["C28"] = meter.heat_energy_kwh
            ws["D28"] = reading_date
        elif meter.description == "Sanitario":
            ws["C31"] = meter.heat_energy_kwh
            ws["D31"] = reading_date

    # Section date headers
    ws["C76"] = reading_date
    ws["C126"] = reading_date
    ws["C178"] = reading_date

    # Per-apartment heat allocator readings
    for ha in report.heat_allocators:
        rows = APT_ROWS.get(ha.apartment_number)
        if rows is None:
            continue
        ws.cell(row=rows["heat"], column=3, value=ha.heat_energy_kwh)
        ws.cell(row=rows["afs"], column=3, value=ha.aux1_volume_m3)

    # Per-apartment water meter readings
    for wm in report.water_meters:
        rows = APT_ROWS.get(wm.apartment_number)
        if rows is None:
            continue
        ws.cell(row=rows["water"], column=3, value=wm.water_volume_m3)

    # Add "Dati Report" sheet with raw FC_report data
    _write_dati_report_sheet(wb, report)

    wb.save(output_path)


def _write_dati_report_sheet(wb: openpyxl.Workbook, report: ParsedReport) -> None:
    """Add a 'Dati Report' sheet with the parsed FC_report data."""
    ws = wb.create_sheet("Dati Report")

    # Row 1: metadata
    ws["A1"] = "File"
    ws["B1"] = report.header.filename
    ws["C1"] = "Data"
    ws["D1"] = report.header.report_date
    ws["E1"] = "Seriale"
    ws["F1"] = report.header.serial

    # Row 3: column headers
    headers = [
        "Tipo",
        "Appartamento",
        "Energia termica",
        "Unita",
        "Volume acqua",
        "Unita",
        "Volume AFS",
        "Unita",
    ]
    for col, h in enumerate(headers, 1):
        ws.cell(row=3, column=col, value=h)

    # Data rows: water meters, then heat allocators, then central meters
    row = 4

    for wm in report.water_meters:
        ws.cell(row=row, column=1, value="Acqua calda")
        ws.cell(row=row, column=2, value=wm.description)
        ws.cell(row=row, column=5, value=wm.water_volume_m3)
        ws.cell(row=row, column=6, value="m3")
        row += 1

    for ha in report.heat_allocators:
        ws.cell(row=row, column=1, value="Contacalorie")
        ws.cell(row=row, column=2, value=ha.description)
        ws.cell(row=row, column=3, value=ha.heat_energy_mwh)
        ws.cell(row=row, column=4, value="MWh")
        ws.cell(row=row, column=7, value=ha.aux1_volume_m3)
        ws.cell(row=row, column=8, value="m3")
        row += 1

    for cm in report.central_meters:
        ws.cell(row=row, column=1, value="Centrale")
        ws.cell(row=row, column=2, value=cm.description)
        ws.cell(row=row, column=3, value=cm.heat_energy_kwh)
        ws.cell(row=row, column=4, value="kWh")
        row += 1


def _parse_reading_date(report: ParsedReport) -> datetime | None:
    """Extract a datetime from the first available readout_date string."""
    for meter in report.central_meters:
        if meter.readout_date:
            try:
                return datetime.strptime(meter.readout_date, "%Y/%m/%d")
            except ValueError:
                pass
    return None
