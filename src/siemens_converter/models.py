"""Data classes for Siemens report structures and static condominium data (no I/O)."""

from __future__ import annotations

import re
from collections.abc import Iterable
from dataclasses import dataclass, field
from datetime import datetime


@dataclass
class ReportHeader:
    filename: str
    report_date: str
    report_time: str
    reference: str
    serial: str
    firmware: str
    total_wired: int


@dataclass
class CentralMeter:
    description: str
    heat_energy_kwh: int
    serial_number: int
    readout_date: str
    communication_status: str = "Ok"


@dataclass
class WaterMeter:
    description: str
    apartment_number: int
    water_volume_m3: float
    serial_number: int
    readout_date: str
    communication_status: str = "Ok"


@dataclass
class HeatAllocator:
    description: str
    apartment_number: int
    heat_energy_mwh: float
    aux1_volume_m3: float
    serial_number: int
    readout_date: str
    communication_status: str = "Ok"

    @property
    def heat_energy_kwh(self) -> int:
        return round(self.heat_energy_mwh * 1000)


@dataclass
class ParsedReport:
    header: ReportHeader
    central_meters: list[CentralMeter]
    water_meters: list[WaterMeter]
    heat_allocators: list[HeatAllocator]
    column_headers: list[str] | None = None
    raw_device_rows: list[list[str]] | None = None
    warnings: list[str] = field(default_factory=list)


# -- Static condominium data (non-Siemens, from administrator's records) --


@dataclass
class ApartmentInfo:
    apartment_number: int
    proprietario: str
    inquilino: str


@dataclass
class Millesimali:
    apartment_number: int
    subalterno: int
    heat_energy_kwh: float
    water_energy_kwh: float


@dataclass
class CostItem:
    label: str
    amount: float


@dataclass
class MeterReading:
    name: str
    unit: str
    initial: float
    final: float


@dataclass
class PreviousReading:
    apartment_number: int
    heat_kwh: int
    water_m3: float
    cold_water_m3: float


@dataclass
class ExerciseInfo:
    year: int
    start_date: datetime


@dataclass
class CentralMeterStart:
    """Start reading of a central heat meter and its optional manual final reading.

    *description* is the one the FC_report gives the meter: Riscaldamento or Sanitario.
    A manual final reading comes with its date, and replaces the report's.
    """

    description: str
    initial_kwh: float
    final_kwh: float | None = None
    final_date: datetime | None = None


@dataclass
class StaticData:
    apartments: list[ApartmentInfo] = field(default_factory=list)
    millesimali: list[Millesimali] = field(default_factory=list)
    costs: list[CostItem] = field(default_factory=list)
    meters: list[MeterReading] = field(default_factory=list)
    previous_readings: list[PreviousReading] = field(default_factory=list)
    exercise: ExerciseInfo | None = None
    central_meters: list[CentralMeterStart] = field(default_factory=list)


def parse_readout_date(text: str) -> datetime | None:
    """Parse a 'YYYY/MM/DD' readout date; None when blank or malformed."""
    try:
        return datetime.strptime(text.strip(), "%Y/%m/%d")
    except ValueError:
        return None


def latest_readout_date(
    meters: Iterable[CentralMeter | WaterMeter | HeatAllocator],
) -> datetime | None:
    """Most recent readout date among *meters*; None when none of them has one."""
    dates = [d for m in meters if (d := parse_readout_date(m.readout_date))]
    return max(dates, default=None)


def extract_apartment_number(description: str) -> int:
    """Extract apartment number from strings like 'App, 01 Rossi Mario' or 'App, 10 Verdi'."""
    match = re.search(r"App[.,]\s*(\d{1,2})", description)
    if match is None:
        raise ValueError(
            f"nella descrizione «{description}» manca il numero dell'appartamento, "
            "che deve comparire come «App, 01»"
        )
    return int(match.group(1))
