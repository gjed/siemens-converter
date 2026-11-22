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
from pathlib import Path

import openpyxl
import pytest

from siemens_converter.models import MeterReading, StaticData
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

    assert wb["Tabella_2026"]["E9"].value == pytest.approx(costs)
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
