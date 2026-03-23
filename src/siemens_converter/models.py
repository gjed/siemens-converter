"""Data classes for Siemens report structures (no I/O)."""

from __future__ import annotations

import re
from dataclasses import dataclass


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


@dataclass
class WaterMeter:
    description: str
    apartment_number: int
    water_volume_m3: float
    serial_number: int
    readout_date: str


@dataclass
class HeatAllocator:
    description: str
    apartment_number: int
    heat_energy_mwh: float
    aux1_volume_m3: float
    serial_number: int
    readout_date: str

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


def extract_apartment_number(description: str) -> int:
    """Extract apartment number from strings like 'App, 01 Rossi Mario' or 'App, 10 Verdi'."""
    match = re.search(r"App[.,]\s*(\d{1,2})", description)
    if match is None:
        raise ValueError(f"Cannot extract apartment number from: {description!r}")
    return int(match.group(1))
