from pathlib import Path
from siemens_converter.parser import parse_fc_report

FIXTURE = Path(__file__).parent / "fixtures" / "FC_report_TEST_9999_2026-01-15.xls"


def test_header_parsed():
    report = parse_fc_report(FIXTURE)
    assert report.header.serial == "TE99999999"
    assert report.header.total_wired == 22


def test_central_meters_count():
    report = parse_fc_report(FIXTURE)
    assert len(report.central_meters) == 2
    descs = {m.description for m in report.central_meters}
    assert "Riscaldamento" in descs
    assert "Sanitario" in descs


def test_central_meter_values():
    report = parse_fc_report(FIXTURE)
    risc = next(m for m in report.central_meters if m.description == "Riscaldamento")
    assert isinstance(risc.heat_energy_kwh, int)
    assert risc.heat_energy_kwh > 0


def test_water_meters_count():
    report = parse_fc_report(FIXTURE)
    assert len(report.water_meters) == 10


def test_water_meter_fields():
    report = parse_fc_report(FIXTURE)
    wm = report.water_meters[0]
    assert wm.apartment_number >= 1
    assert wm.water_volume_m3 > 0


def test_heat_allocators_count():
    report = parse_fc_report(FIXTURE)
    assert len(report.heat_allocators) == 10


def test_heat_allocator_fields():
    report = parse_fc_report(FIXTURE)
    ha = report.heat_allocators[0]
    assert ha.apartment_number >= 1
    assert ha.heat_energy_mwh > 0
    assert ha.heat_energy_kwh == round(ha.heat_energy_mwh * 1000)
    assert ha.aux1_volume_m3 > 0


def test_comma_decimal_parsing():
    report = parse_fc_report(FIXTURE)
    ha = report.heat_allocators[0]
    assert isinstance(ha.heat_energy_mwh, float)


def test_apartment_numbers_sorted():
    report = parse_fc_report(FIXTURE)
    water_nums = [wm.apartment_number for wm in report.water_meters]
    heat_nums = [ha.apartment_number for ha in report.heat_allocators]
    assert water_nums == sorted(water_nums)
    assert heat_nums == sorted(heat_nums)
