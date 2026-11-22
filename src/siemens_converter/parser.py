"""Parse Siemens FC report .xls files (HTML-encoded tables)."""

from __future__ import annotations

from collections.abc import Callable
from html.parser import HTMLParser
from pathlib import Path

from siemens_converter.models import (
    CentralMeter,
    HeatAllocator,
    ParsedReport,
    ReportHeader,
    WaterMeter,
    extract_apartment_number,
    latest_readout_date,
    parse_readout_date,
)
from siemens_converter.writer import APT_ROWS

# Device kinds as the user reads them: (singular, plural)
_CENTRAL = ("contatore centrale", "contatori centrali")
_HEAT = ("contacalorie", "contacalorie")
_WATER = ("contatore dell'acqua calda", "contatori dell'acqua calda")

# Readings the workbook needs: column index -> (name for the user, column header)
_READINGS = {
    15: ("dell'energia", "heat_energy"),
    24: ("dell'acqua calda", "water_volume"),
    26: ("dell'acqua fredda", "aux1_volume"),
}


class ReportError(ValueError):
    """An FC report that must not become a workbook; the message is for the user."""


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
    bad_readings: list[str] = []

    for row in rows[4:]:
        name_device = row[3].strip()
        device_description = row[4].strip()
        device_detail = row[5].strip()
        serial_number = int(row[2])
        readout_date = row[9]
        status = row[11].strip()

        if name_device == "contacalorie CT" and device_detail == "":
            # Central meter
            central_meters.append(
                CentralMeter(
                    description=device_description,
                    heat_energy_kwh=_reading(row, 15, int, _CENTRAL, bad_readings),
                    serial_number=serial_number,
                    readout_date=readout_date,
                    communication_status=status,
                )
            )
        elif device_detail == "Acqua calda":
            # Water meter
            water_meters.append(
                WaterMeter(
                    description=device_description,
                    apartment_number=extract_apartment_number(device_description),
                    water_volume_m3=_reading(
                        row, 24, parse_decimal, _WATER, bad_readings
                    ),
                    serial_number=serial_number,
                    readout_date=readout_date,
                    communication_status=status,
                )
            )
        elif device_detail == "Contacalorie":
            # Heat allocator
            heat_allocators.append(
                HeatAllocator(
                    description=device_description,
                    apartment_number=extract_apartment_number(device_description),
                    heat_energy_mwh=_reading(
                        row, 15, parse_decimal, _HEAT, bad_readings
                    ),
                    aux1_volume_m3=_reading(
                        row, 26, parse_decimal, _HEAT, bad_readings
                    ),
                    serial_number=serial_number,
                    readout_date=readout_date,
                    communication_status=status,
                )
            )

    water_meters.sort(key=lambda m: m.apartment_number)
    heat_allocators.sort(key=lambda m: m.apartment_number)

    central_found = {m.description for m in central_meters}
    missing = [
        f"{_CENTRAL[0]} «{description}»"
        for description in ("Riscaldamento", "Sanitario")
        if description not in central_found
    ]
    for kind, meters in ((_HEAT, heat_allocators), (_WATER, water_meters)):
        found = {m.apartment_number for m in meters}
        missing += [
            f"{kind[0]} dell'appartamento {apt}" for apt in APT_ROWS if apt not in found
        ]
    if missing or bad_readings:
        raise ReportError(_refusal(missing, bad_readings))

    # Preserve raw data for the Dati Report sheet
    column_headers = rows[3] if len(rows) > 3 else []
    # Sort raw rows: water meters first, then heat allocators, then central
    water_rows = [r for r in rows[4:] if r[5].strip() == "Acqua calda"]
    heat_rows = [r for r in rows[4:] if r[5].strip() == "Contacalorie"]
    central_rows = [r for r in rows[4:] if r[3].strip() == "contacalorie CT"]

    # Sort apartment rows by apartment number
    def _apt_sort_key(r: list[str]) -> int:
        try:
            return extract_apartment_number(r[4])
        except ValueError:
            return 9999

    water_rows.sort(key=_apt_sort_key)
    heat_rows.sort(key=_apt_sort_key)
    raw_device_rows = water_rows + heat_rows + central_rows

    report = ParsedReport(
        header=header,
        central_meters=central_meters,
        water_meters=water_meters,
        heat_allocators=heat_allocators,
        column_headers=column_headers,
        raw_device_rows=raw_device_rows,
    )
    report.warnings = _warnings(report)
    return report


def _reading(
    row: list[str],
    col: int,
    convert: Callable[[str], float],
    kind: tuple[str, str],
    problems: list[str],
) -> float:
    """Convert a required reading; a blank or non-numeric one is added to *problems*."""
    text = row[col].strip()
    name, column = _READINGS[col]
    try:
        if text:
            return convert(text)
        found = "è vuota"
    except ValueError:
        found = f"non è un numero valido («{text}»)"
    problems.append(
        f"{kind[0]} «{row[4].strip()}»: la lettura {name} (colonna «{column}») {found}"
    )
    return 0


def _refusal(missing: list[str], bad_readings: list[str]) -> str:
    """Build the message for a report that cannot be converted."""
    parts = ["il report non è completo e il riparto non è stato creato."]
    for title, items in (
        ("Dispositivi mancanti:", missing),
        ("Letture mancanti o non valide:", bad_readings),
    ):
        if items:
            parts.append("\n".join([title, *(f"- {item}" for item in items)]))
    parts.append(
        "Scaricare di nuovo il report FC e ripetere la conversione. "
        "Se il messaggio compare ancora, far controllare sull'impianto "
        "i dispositivi indicati."
    )
    return "\n\n".join(parts)


def _warnings(report: ParsedReport) -> list[str]:
    """Problems the user must know about but that do not stop the conversion."""
    warnings: list[str] = []
    day = "%d/%m/%Y"

    # Each central meter against the apartment meters it is billed against
    for central, kind, meters in (
        ("Riscaldamento", _HEAT, report.heat_allocators),
        ("Sanitario", _WATER, report.water_meters),
    ):
        section = latest_readout_date(meters)
        if section is None:
            continue
        for meter in report.central_meters:
            read = parse_readout_date(meter.readout_date)
            if meter.description != central or read is None or read == section:
                continue
            days = abs((read - section).days)
            warnings.append(
                f"Il {_CENTRAL[0]} «{central}» è stato letto il {read:{day}}, "
                f"{days} {'giorno' if days == 1 else 'giorni'} "
                f"{'prima dei' if read < section else 'dopo i'} {kind[1]} degli "
                f"appartamenti ({section:{day}}). I due consumi non coprono lo "
                "stesso periodo: controllare il riparto prima di usarlo."
            )
        earlier = [
            f"«{m.description}» ({read:{day}})"
            for m in meters
            if (read := parse_readout_date(m.readout_date)) and read != section
        ]
        if earlier:
            warnings.append(
                f"I {kind[1]} degli appartamenti non sono stati letti tutti lo "
                f"stesso giorno. Nel riparto è indicata la data più recente "
                f"({section:{day}}); letti prima: {', '.join(earlier)}. "
                "Controllare queste letture prima di usare il riparto."
            )

    not_ok = [
        f"{kind[0]} «{m.description}» ({m.communication_status or 'vuoto'})"
        for kind, meters in (
            (_CENTRAL, report.central_meters),
            (_HEAT, report.heat_allocators),
            (_WATER, report.water_meters),
        )
        for m in meters
        if m.communication_status != "Ok"
    ]
    if not_ok:
        warnings.append(
            f"Lo stato di questi dispositivi nel report non è «Ok»: "
            f"{'; '.join(not_ok)}. Controllare le loro letture prima di usare "
            "il riparto."
        )
    return warnings
