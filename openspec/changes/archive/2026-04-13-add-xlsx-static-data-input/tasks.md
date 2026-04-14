## 1. Data Models

- [x] 1.1 Add static data dataclasses to `models.py`: `ApartmentInfo`, `Millesimali`, `CostItem`, `MeterReading`, `PreviousReading`, `StaticData`
- [x] 1.2 Add unit tests for the new dataclasses in `tests/test_models.py`

## 2. Static Data Template

- [x] 2.1 Create `src/siemens_converter/static_data_template.xlsx` with five sheets (Inquilini, Millesimali, Costi, Contatori, Letture_precedenti), proper column headers, and formatted cells
- [x] 2.2 Add test in `tests/test_template.py` to verify template structure (sheet names, column headers)

## 3. Static Data Reader

- [x] 3.1 Create `src/siemens_converter/static_reader.py` with `parse_static_data(path) -> StaticData` function
- [x] 3.2 Add validation: raise `ValueError` for missing required sheets
- [x] 3.3 Handle graceful defaults: missing cost rows → 0.0, fewer apartment rows → empty strings
- [x] 3.4 Create anonymized test fixture `tests/fixtures/static_data_TEST.xlsx` with fake but structurally valid data
- [x] 3.5 Add unit tests in `tests/test_static_reader.py`: valid parse, missing sheet, missing costs, empty apartments

## 4. Writer Integration

- [x] 4.1 Add `static_data: StaticData | None = None` parameter to `write_xlsx` in `writer.py`
- [x] 4.2 Inject tenant names (inquilino) into Inquilini sheet column B when static data provided
- [x] 4.3 Inject millesimali (subalterno, energy values) into Tabelle millesimali sheet columns B, C, E
- [x] 4.4 Inject costs into Tabella_2026 sheet cells E3-E8
- [x] 4.5 Inject meter readings into Tabella_2026 sheet cells F20-G22
- [x] 4.6 Inject previous readings into Ripartizione heat/water/AFS column B rows
- [x] 4.7 Add integration tests in `tests/test_writer.py`: output with static data has correct cell values, output without static data is unchanged

## 5. CLI / Entry Point

- [x] 5.1 Update `__main__.py` to accept optional second argument (static data xlsx path)
- [x] 5.2 Add Windows file dialog for static data selection (sequential after FC_report dialog) in no-console mode
- [x] 5.3 Validate static data file exists before processing, show clear error on missing file
- [x] 5.4 Update usage message to reflect two-file workflow
- [x] 5.5 Add CLI tests in `tests/test_cli.py` for single-file and two-file invocations

## 6. Compiled Test Data

- [x] 6.1 Create `assets/static_data_2026.xlsx` from real 2026 Riparto data (extract tenant names, millesimali, costs, meter readings, previous readings into the template structure)
- [x] 6.2 Verify the compiled file is gitignored (already covered by `assets/` pattern)
- [x] 6.3 Manual end-to-end test: run converter with both FC_report and compiled static data, verify output matches expected 2026 Riparto structure

## 7. Build & Distribution

- [x] 7.1 Update PyInstaller `--add-data` in `README.md` and build scripts to include `static_data_template.xlsx`
- [x] 7.2 Update `_get_template_path` pattern or add analogous `_get_static_template_path` for PyInstaller bundling
- [x] 7.3 Run full test suite, verify CI green
