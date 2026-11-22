# Tasks

## 1. Models

- [x] 1.1 Add `communication_status` (default `Ok`) to `CentralMeter`, `WaterMeter` and
  `HeatAllocator`, and `warnings` to `ParsedReport`; verify the existing `tests/test_models.py`
  still passes
- [x] 1.2 Add the pure helpers `parse_readout_date` and `latest_readout_date` to `models.py`;
  verify with unit tests for a valid date, a blank date, a malformed date and a mixed group

## 2. Completeness and required readings

- [x] 2.1 Add `ReportError` to `parser.py` and raise it, listing every missing device, when the
  report lacks a central meter or an apartment meter; verify with tests that drop a central-meter
  row, a heat-meter row and a hot-water-meter row from the fixture and assert the wording
- [x] 2.2 Report blank and non-numeric `heat_energy`, `aux1_volume` and `water_volume` readings in
  the same error, naming the device; verify with tests that blank each of the four required cells
  and one that writes text in a reading
- [x] 2.3 Verify with a test that a missing device and a blank reading are reported in one error

## 3. Dates

- [x] 3.1 Write `D28` and `D31` from their own central meter and `C76`/`C178`/`C126` from the heat
  and hot-water meters in `writer.py`, removing `_parse_reading_date`; verify by replacing
  `tests/test_writer.py::test_reading_date` with a test where all four groups carry different
  dates, and a test for apartment meters of the same kind with different dates

## 4. Warnings

- [x] 4.1 Read `communication_status` in the parser and fill `ParsedReport.warnings` with the
  status warning; verify with tests that change the status of a central meter and of both meters
  of one apartment, and one that asserts no warning on the unmodified fixture
- [x] 4.2 Add the readout-date warnings (central against apartment meters, apartment meters of the
  same kind); verify with tests for 12 days before, 1 day after and one heat meter left behind

## 5. Entry point

- [x] 5.1 Wrap the `parse_fc_report` call in `__main__.py`; verify with tests that a refused
  report exits with code 1 and prints the message on the console, and shows it in a message box
  when there is no console
- [x] 5.2 Print warnings before `Done` and append them to the final message box; verify with tests
  for the console and the no-console path, and one asserting the output is unchanged without
  warnings

## 6. Integration

- [x] 6.1 Convert the test fixture with the code before and after this change and verify every
  cell of every sheet is equal
- [x] 6.2 Convert the two real reports from scratch copies and verify the expected dates and
  warnings
- [x] 6.3 Run the full test suite, `openspec validate report-validation` and the markdown lint

## Workflow follow-up

- Archive the change with `openspec archive report-validation -y` before the final commit.
