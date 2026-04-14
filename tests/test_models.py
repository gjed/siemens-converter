from siemens_converter.models import (
    ApartmentInfo,
    CentralMeter,
    CostItem,
    HeatAllocator,
    MeterReading,
    Millesimali,
    ParsedReport,
    PreviousReading,
    ReportHeader,
    StaticData,
    WaterMeter,
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


# -- Static data models --


def test_apartment_info():
    a = ApartmentInfo(apartment_number=1, proprietario="Rossi", inquilino="Bianchi")
    assert a.apartment_number == 1
    assert a.proprietario == "Rossi"
    assert a.inquilino == "Bianchi"


def test_millesimali():
    m = Millesimali(
        apartment_number=1,
        subalterno=20,
        heat_energy_kwh=1344.03,
        water_energy_kwh=1101.36,
    )
    assert m.subalterno == 20
    assert m.heat_energy_kwh == 1344.03


def test_cost_item():
    c = CostItem(label="Gas metano", amount=2141.0)
    assert c.label == "Gas metano"
    assert c.amount == 2141.0


def test_meter_reading():
    m = MeterReading(name="Energia elettrica CT", unit="kWh", initial=0.0, final=1086.0)
    assert m.final - m.initial == 1086.0


def test_previous_reading():
    p = PreviousReading(
        apartment_number=1, heat_kwh=4048, water_m3=25.5, cold_water_m3=10.2
    )
    assert p.heat_kwh == 4048


def test_static_data_defaults():
    sd = StaticData()
    assert sd.apartments == []
    assert sd.millesimali == []
    assert sd.costs == []
    assert sd.meters == []
    assert sd.previous_readings == []


def test_static_data_with_values():
    sd = StaticData(
        apartments=[ApartmentInfo(1, "Rossi", "Bianchi")],
        costs=[CostItem("Gas", 100.0)],
    )
    assert len(sd.apartments) == 1
    assert len(sd.costs) == 1
