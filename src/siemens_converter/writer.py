"""XLSX workbook generation from parsed FC_report data."""

from __future__ import annotations

import shutil
from datetime import datetime
from importlib.resources import files
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from siemens_converter.models import ParsedReport

# -- Styles matching the original "File con dati necessari" formatting --

_FONT = Font(name="Aptos Narrow", size=10)
_FONT_BOLD = Font(name="Aptos Narrow", size=10, bold=True)
_FONT_HEADER_ANNOTATION = Font(
    name="Aptos Narrow", size=10, bold=True, color="FFFF0000"
)

_FILL_HEADER = PatternFill(
    start_color="FFDBE8F0", end_color="FFDBE8F0", fill_type="solid"
)
_FILL_GREEN = PatternFill(
    start_color="FF99FFCC", end_color="FF99FFCC", fill_type="solid"
)
_FILL_YELLOW = PatternFill(
    start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid"
)

_ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
_ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")

_THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

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

    # Populate Tabelle millesimali apartment names from FC_report
    _write_millesimali_names(wb, report)

    # Add "Dati Report" sheet with raw FC_report data
    _write_dati_report_sheet(wb, report)

    wb.save(output_path)


# Millesimali sheet: apartment number -> row (rows 4-13)
_MILL_ROWS: dict[int, int] = {i: i + 3 for i in range(1, 11)}


def _write_millesimali_names(wb: openpyxl.Workbook, report: ParsedReport) -> None:
    """Populate apartment names in the Tabelle millesimali sheet."""
    ws = wb["Tabelle millesimali"]
    for ha in report.heat_allocators:
        row = _MILL_ROWS.get(ha.apartment_number)
        if row is None:
            continue
        ws.cell(row=row, column=1, value=ha.description)


def _write_dati_report_sheet(wb: openpyxl.Workbook, report: ParsedReport) -> None:
    """Add a 'Dati Report' sheet with the parsed FC_report data."""
    ws = wb.create_sheet("Dati Report")
    num_cols = 8

    # -- Row 1: metadata labels --
    meta_labels = ["File", "Data", "Ora", "Riferimento", "Seriale", "Firmware"]
    meta_values = [
        report.header.filename,
        report.header.report_date,
        report.header.report_time,
        report.header.reference,
        report.header.serial,
        report.header.firmware,
    ]
    for col, label in enumerate(meta_labels, 1):
        c = ws.cell(row=1, column=col, value=label)
        c.font = _FONT_BOLD
        c.fill = _FILL_HEADER
        c.alignment = _ALIGN_CENTER
        c.border = _THIN_BORDER
    ws.row_dimensions[1].height = 35

    # -- Row 2: metadata values --
    for col, val in enumerate(meta_values, 1):
        c = ws.cell(row=2, column=col, value=val)
        c.font = _FONT
        c.fill = _FILL_HEADER
        c.alignment = _ALIGN_CENTER
        c.border = _THIN_BORDER

    # -- Row 3: annotation row (highlight key data columns) --
    annotations = {3: "energia termica", 5: "volume acqua", 7: "volume AFS"}
    for col, label in annotations.items():
        c = ws.cell(row=3, column=col, value=label)
        c.font = _FONT_HEADER_ANNOTATION
        c.fill = _FILL_GREEN
        c.alignment = _ALIGN_CENTER
        c.border = _THIN_BORDER
    ws.row_dimensions[3].height = 37

    # -- Row 4: column headers --
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
        c = ws.cell(row=4, column=col, value=h)
        c.font = _FONT_BOLD
        c.fill = _FILL_HEADER
        c.alignment = _ALIGN_LEFT
        c.border = _THIN_BORDER
    ws.row_dimensions[4].height = 35

    # -- Data rows --
    row = 5

    for wm in report.water_meters:
        _set_data_cell(ws, row, 1, "Acqua calda")
        _set_data_cell(ws, row, 2, wm.description, bold=True)
        _set_data_cell(ws, row, 5, wm.water_volume_m3, fmt="0.000", green=True)
        _set_data_cell(ws, row, 6, "m\u00b3")
        # Fill empty cols with border
        for empty_col in [3, 4, 7, 8]:
            _set_data_cell(ws, row, empty_col, None)
        row += 1

    for ha in report.heat_allocators:
        _set_data_cell(ws, row, 1, "Contacalorie")
        _set_data_cell(ws, row, 2, ha.description, bold=True)
        _set_data_cell(ws, row, 3, ha.heat_energy_mwh, fmt="0.000", green=True)
        _set_data_cell(ws, row, 4, "MWh")
        _set_data_cell(ws, row, 5, None)
        _set_data_cell(ws, row, 6, None)
        _set_data_cell(ws, row, 7, ha.aux1_volume_m3, fmt="0.00", green=True)
        _set_data_cell(ws, row, 8, "m\u00b3")
        row += 1

    for cm in report.central_meters:
        _set_data_cell(ws, row, 1, "Centrale")
        _set_data_cell(ws, row, 2, cm.description, bold=True)
        _set_data_cell(ws, row, 3, cm.heat_energy_kwh, green=True)
        _set_data_cell(ws, row, 4, "kWh")
        for empty_col in [5, 6, 7, 8]:
            _set_data_cell(ws, row, empty_col, None)
        row += 1

    # -- Column widths --
    col_widths = {1: 14, 2: 35, 3: 16, 4: 8, 5: 14, 6: 8, 7: 14, 8: 8}
    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # -- Freeze panes: freeze header rows --
    ws.freeze_panes = "A5"

    # -- Auto-filter on data table --
    last_row = row - 1
    ws.auto_filter.ref = f"A4:H{last_row}"


def _set_data_cell(
    ws: openpyxl.worksheet.worksheet.Worksheet,  # type: ignore[name-defined]
    row: int,
    col: int,
    value: object,
    *,
    bold: bool = False,
    fmt: str | None = None,
    green: bool = False,
) -> None:
    """Write a value to a cell with consistent data-row styling."""
    c = ws.cell(row=row, column=col, value=value)
    c.font = _FONT_BOLD if bold else _FONT
    c.border = _THIN_BORDER
    if green:
        c.fill = _FILL_GREEN
        c.font = _FONT_BOLD
    if fmt:
        c.number_format = fmt
    if isinstance(value, (int, float)):
        c.alignment = _ALIGN_RIGHT
    else:
        c.alignment = _ALIGN_LEFT


def _parse_reading_date(report: ParsedReport) -> datetime | None:
    """Extract a datetime from the first available readout_date string."""
    for meter in report.central_meters:
        if meter.readout_date:
            try:
                return datetime.strptime(meter.readout_date, "%Y/%m/%d")
            except ValueError:
                pass
    return None
