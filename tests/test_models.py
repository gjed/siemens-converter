from siemens_converter.models import (
    ReportHeader,
    CentralMeter,
    WaterMeter,
    HeatAllocator,
    ParsedReport,
    extract_apartment_number,
)


def test_report_header():
    h = ReportHeader(
        filename="FC_report.xls",
        report_date="2026-03-14",
        report_time="16:45:03",
        reference="-",
        serial="EV23234456",
        firmware="3.93",
        total_wired=22,
    )
    assert h.serial == "EV23234456"
    assert h.total_wired == 22


def test_central_meter():
    m = CentralMeter(
        description="Riscaldamento",
        heat_energy_kwh=29213,
        serial_number=71714666,
        readout_date="2026/03/02",
    )
    assert m.heat_energy_kwh == 29213


def test_water_meter():
    w = WaterMeter(
        description="App, 01 Rossi Mario",
        apartment_number=1,
        water_volume_m3=31.613,
        serial_number=23005747,
        readout_date="2026/03/02",
    )
    assert w.apartment_number == 1
    assert w.water_volume_m3 == 31.613


def test_heat_allocator():
    h = HeatAllocator(
        description="App, 01 Rossi Mario",
        apartment_number=1,
        heat_energy_mwh=5.243,
        aux1_volume_m3=45.72,
        serial_number=71714345,
        readout_date="2026/03/02",
    )
    assert h.heat_energy_kwh == 5243  # property: round(MWh * 1000)


def test_parsed_report():
    header = ReportHeader("f.xls", "2026-01-01", "12:00:00", "-", "EV1", "1.0", 10)
    report = ParsedReport(
        header=header, central_meters=[], water_meters=[], heat_allocators=[]
    )
    assert report.header.serial == "EV1"


def test_extract_apartment_number_normal():
    assert extract_apartment_number("App, 01 Rossi Mario") == 1


def test_extract_apartment_number_two_digit():
    assert extract_apartment_number("App, 10 Verdi Giuseppe") == 10


def test_extract_apartment_number_no_space_dash():
    assert extract_apartment_number("App, 04 Neri Francesco- Bianchi") == 4


def test_extract_apartment_number_double_space():
    assert extract_apartment_number("App, 01 Rossi Mario  - Gialli") == 1
