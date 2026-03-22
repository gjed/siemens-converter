"""Parse Siemens FC report .xls files (HTML-encoded tables)."""

from __future__ import annotations

from html.parser import HTMLParser
from pathlib import Path

from siemens_converter.models import (
    CentralMeter,
    HeatAllocator,
    ParsedReport,
    ReportHeader,
    WaterMeter,
    extract_apartment_number,
)


def parse_decimal(s: str) -> float:
    """Parse an Italian-format decimal string (comma separator) to float.

    Returns 0.0 for empty or whitespace-only strings.
    """
    s = s.strip()
    if not s:
        return 0.0
    return float(s.replace(",", "."))


class _TableParser(HTMLParser):
    """Extract all rows/cells from the first <table> in an HTML document."""

    def __init__(self) -> None:
        super().__init__()
        self.rows: list[list[str]] = []
        self._current_row: list[str] | None = None
        self._current_cell: list[str] | None = None
        self._in_table = False

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag == "table":
            self._in_table = True
        elif tag == "tr" and self._in_table:
            self._current_row = []
        elif tag in ("td", "th") and self._current_row is not None:
            self._current_cell = []

    def handle_endtag(self, tag: str) -> None:
        if tag in ("td", "th") and self._current_cell is not None:
            assert self._current_row is not None
            self._current_row.append("".join(self._current_cell))
            self._current_cell = None
        elif tag == "tr" and self._current_row is not None:
            self.rows.append(self._current_row)
            self._current_row = None
        elif tag == "table":
            self._in_table = False

    def handle_data(self, data: str) -> None:
        if self._current_cell is not None:
            self._current_cell.append(data)


def parse_fc_report(path: str | Path) -> ParsedReport:
    """Parse a Siemens FC report file and return a structured ParsedReport."""
    path = Path(path)
    html = path.read_text(encoding="utf-8")

    parser = _TableParser()
    parser.feed(html)
    rows = parser.rows

    # Row 0: metadata headers (ignored)
    # Row 1: metadata values
    meta = rows[1]
    header = ReportHeader(
        filename=meta[0],
        report_date=meta[1],
        report_time=meta[2],
        reference=meta[3],
        serial=meta[7],
        firmware=meta[4],
        total_wired=int(meta[5]),
    )

    # Row 2: blank separator
    # Row 3: column headers
    # Rows 4+: device data
    central_meters: list[CentralMeter] = []
    water_meters: list[WaterMeter] = []
    heat_allocators: list[HeatAllocator] = []

    for row in rows[4:]:
        name_device = row[3].strip()
        device_description = row[4].strip()
        device_detail = row[5].strip()
        serial_number = int(row[2])
        readout_date = row[9]

        if name_device == "contacalorie CT" and device_detail == "":
            # Central meter
            central_meters.append(
                CentralMeter(
                    description=device_description,
                    heat_energy_kwh=int(row[15]),
                    serial_number=serial_number,
                    readout_date=readout_date,
                )
            )
        elif device_detail == "Acqua calda":
            # Water meter
            water_meters.append(
                WaterMeter(
                    description=device_description,
                    apartment_number=extract_apartment_number(device_description),
                    water_volume_m3=parse_decimal(row[24]),
                    serial_number=serial_number,
                    readout_date=readout_date,
                )
            )
        elif device_detail == "Contacalorie":
            # Heat allocator
            heat_allocators.append(
                HeatAllocator(
                    description=device_description,
                    apartment_number=extract_apartment_number(device_description),
                    heat_energy_mwh=parse_decimal(row[15]),
                    aux1_volume_m3=parse_decimal(row[26]),
                    serial_number=serial_number,
                    readout_date=readout_date,
                )
            )

    water_meters.sort(key=lambda m: m.apartment_number)
    heat_allocators.sort(key=lambda m: m.apartment_number)

    return ParsedReport(
        header=header,
        central_meters=central_meters,
        water_meters=water_meters,
        heat_allocators=heat_allocators,
    )
