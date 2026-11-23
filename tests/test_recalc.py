"""Recalculation tests: the generated workbook's formulas, evaluated by LibreOffice.

openpyxl does not calculate, so every other test only sees formula text.
These tests run the real pipeline, let LibreOffice recalculate the result
and assert on the computed values.
"""

from __future__ import annotations

import dataclasses
import os
import shutil
import subprocess
from datetime import datetime
from pathlib import Path

import openpyxl
import pytest

from siemens_converter.models import (
    CentralMeterStart,
    CostItem,
    MeterReading,
    StaticData,
)
from siemens_converter.parser import parse_fc_report
from siemens_converter.static_reader import parse_static_data
from siemens_converter.writer import write_xlsx

FIXTURES = Path(__file__).parent / "fixtures"

# Per-apartment totals in the "Ripartizione totale" table, apartments 1-10.
APARTMENT_TOTALS = [
    f"I{row}" for row in (229, 231, 233, 235, 237, 239, 241, 243, 245, 246)
]


def _with_meter_start(static: StaticData, start: float) -> StaticData:
    """Same consumption on every utility meter, initial reading moved to *start*."""
    return dataclasses.replace(
        static,
        meters=[
            MeterReading(m.name, m.unit, start, start + (m.final - m.initial))
            for m in static.meters
        ],
    )


def _with_cost(static: StaticData, label: str, amount: float) -> StaticData:
    return dataclasses.replace(
        static,
        costs=[
            CostItem(c.label, amount if c.label == label else c.amount)
            for c in static.costs
        ],
    )


def _with_manual_final_reading(static: StaticData) -> StaticData:
    """Heating central meter: final reading and date typed in the static data."""
    heating, acs = static.central_meters
    return dataclasses.replace(
        static,
        central_meters=[
            CentralMeterStart(
                heating.description,
                heating.initial_kwh,
                heating.initial_kwh + 3500,
                datetime(2026, 1, 9),
            ),
            acs,
        ],
    )


def _static_variants() -> dict[str, StaticData]:
    base = parse_static_data(FIXTURES / "static_data_TEST.xlsx")
    return {
        "zero": _with_meter_start(base, 0),
        "offset": _with_meter_start(base, 1000),
        "millesimali": dataclasses.replace(
            _with_meter_start(base, 0),
            millesimali=[
                dataclasses.replace(m, heat_energy_kwh=m.heat_energy_kwh * 3)
                for m in base.millesimali
            ],
        ),
        # A cost that needs the sheet-scoped name costoafs
        "water": _with_cost(_with_meter_start(base, 0), "Acqua condominio", 600.0),
        "manual": _with_manual_final_reading(_with_meter_start(base, 0)),
    }


VARIANTS = tuple(_static_variants())


def _find_soffice() -> str:
    soffice = shutil.which("soffice")
    if soffice is None:
        if os.environ.get("CI"):
            pytest.fail(
                "LibreOffice (soffice) is not on PATH: "
                "the recalculation tests must not be skipped in CI"
            )
        pytest.skip("LibreOffice (soffice) is not on PATH")
    return soffice


@pytest.fixture(scope="module")
def recalculated(tmp_path_factory) -> dict[str, openpyxl.Workbook]:
    """One workbook per static data variant, recalculated by LibreOffice."""
    soffice = _find_soffice()
    work = tmp_path_factory.mktemp("recalc")
    generated = work / "generated"
    generated.mkdir()

    report = parse_fc_report(FIXTURES / "FC_report_TEST_9999_2026-01-15.xls")
    sources = []
    for name, static in _static_variants().items():
        source = generated / f"{name}.xlsx"
        write_xlsx(report, source, static_data=static)
        sources.append(str(source))

    subprocess.run(
        [
            soffice,
            # Own profile: no clash with a LibreOffice already running.
            f"-env:UserInstallation={(work / 'profile').as_uri()}",
            "--headless",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(work / "recalculated"),
            *sources,
        ],
        check=True,
        capture_output=True,
        timeout=300,
    )
    return {
        name: openpyxl.load_workbook(
            work / "recalculated" / f"{name}.xlsx", data_only=True
        )
        for name in VARIANTS
    }


@pytest.mark.parametrize("variant", VARIANTS)
def test_no_error_cells(recalculated, variant):
    errors = [
        f"{ws.title}!{cell.coordinate}={cell.value}"
        for ws in recalculated[variant].worksheets[:3]
        for row in ws.iter_rows()
        for cell in row
        if isinstance(cell.value, str) and cell.value.startswith("#")
    ]
    assert errors == []


@pytest.mark.parametrize("variant", VARIANTS)
def test_distributed_total_equals_costs(recalculated, variant):
    wb = recalculated[variant]
    ws = wb.worksheets[0]
    costs = sum(cost.amount for cost in _static_variants()[variant].costs)

    assert wb["Tabella"]["E9"].value == pytest.approx(costs)
    assert ws["I247"].value == pytest.approx(costs)
    assert ws["L250"].value == "esatto"


def test_apartment_totals_ignore_initial_meter_reading(recalculated):
    zero = [recalculated["zero"].worksheets[0][c].value for c in APARTMENT_TOTALS]
    offset = [recalculated["offset"].worksheets[0][c].value for c in APARTMENT_TOTALS]

    assert None not in zero
    assert offset == pytest.approx(zero, rel=1e-9)


@pytest.mark.parametrize("variant", ["zero", "offset"])
def test_utility_readings_and_consumptions(recalculated, variant):
    ws = recalculated[variant].worksheets[0]
    meters = {m.name.upper(): m for m in _static_variants()[variant].meters}
    electricity = meters["ENERGIA ELETTRICA CT"]
    gas = meters["GAS METANO CT"]

    assert ws["C25"].value == electricity.final - electricity.initial
    assert ws["C23"].value == gas.final
    assert ws["C36"].value == gas.final - gas.initial


@pytest.mark.parametrize("variant", ["zero", "millesimali"])
def test_fixed_cost_split_follows_millesimali(recalculated, variant):
    wb = recalculated[variant]
    ws = wb.worksheets[0]
    shares = wb["Tabelle millesimali"]
    millesimali = _static_variants()[variant].millesimali
    heat = sum(m.heat_energy_kwh for m in millesimali)
    water = sum(m.water_energy_kwh for m in millesimali)

    assert shares["C18"].value == pytest.approx(heat / (heat + water))
    assert shares["E18"].value == pytest.approx(water / (heat + water))
    assert ws["C63"].value == pytest.approx(shares["C18"].value)
    assert ws["C64"].value == pytest.approx(shares["E18"].value)


def test_central_meter_consumption_uses_the_start_readings_of_the_static_data(
    recalculated,
):
    ws = recalculated["zero"].worksheets[0]
    static = _static_variants()["zero"]
    heating, acs = static.central_meters

    assert ws["C27"].value == heating.initial_kwh
    assert ws["C30"].value == acs.initial_kwh
    # The fixture report reads 28000 kWh (heating) and 25000 kWh (ACS)
    assert ws["C29"].value == 28000 - heating.initial_kwh
    assert ws["C32"].value == 25000 - acs.initial_kwh


def test_central_meter_consumption_uses_the_manual_final_reading(recalculated):
    ws = recalculated["manual"].worksheets[0]
    heating, acs = _static_variants()["manual"].central_meters

    assert ws["C28"].value == heating.final_kwh
    assert ws["D28"].value == heating.final_date
    assert ws["C29"].value == 3500
    # The ACS meter has no manual reading: the report's 25000 kWh stays
    assert ws["C31"].value == 25000
    assert ws["C32"].value == 25000 - acs.initial_kwh


def test_apartment_start_readings_and_exercise_labels_come_from_the_static_data(
    recalculated,
):
    ws = recalculated["zero"].worksheets[0]
    first = _static_variants()["zero"].previous_readings[0]

    assert ws["B78"].value == first.heat_kwh
    assert ws["B130"].value == first.water_m3
    assert ws["B182"].value == first.cold_water_m3
    assert ws["D78"].value == ws["C78"].value - first.heat_kwh
    assert ws["A28"].value.endswith("FINE GESTIONE 25")
    assert ws["B178"].value == datetime(2024, 12, 31)


def test_sheet_scoped_names_follow_the_renamed_sheets(recalculated):
    """costoenegiaelettrica and costoafs point at the sheets under their new names."""
    wb = recalculated["water"]
    ws = wb.worksheets[0]

    # costoenegiaelettrica (scoped to Ripartizione) is C26, the average cost in Tabella
    assert ws["C26"].value == wb["Tabella"]["E12"].value
    assert ws["C44"].value == ws["C26"].value
    # costoafs (scoped to Tabella) spreads the 600 EUR of cold water over the volumes
    assert wb["Tabella"]["E13"].value == pytest.approx(
        600 / (ws["D200"].value + ws["E200"].value)
    )
    assert ws["F200"].value == pytest.approx(600)
