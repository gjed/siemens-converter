## Why

The converter currently processes Siemens FC_report data (meter readings) but produces output with missing information — tenant names, millesimali tables, cost figures, utility readings, and previous-period meter values were stripped because the repo is public. Without this static data, the output workbook requires extensive manual entry to be usable for actual condominium billing. Adding a second xlsx input file with this "static" (non-Siemens) data completes the pipeline end-to-end.

## What Changes

- **New xlsx input**: A second file (alongside the FC_report) provides static/administrative data: tenant names (inquilini), prepared millesimali tables, annual costs (gas, electricity, water, maintenance), utility meter readings, and previous-period allocator readings.
- **Template file**: A blank `.xlsx` template committed to the repo that users populate with their condominium's data. This template has the correct structure and column headers but no confidential content.
- **Compiled data file**: A filled-in version using real 2026 Riparto data, placed in `assets/` (gitignored), for manual integration testing.
- **Two-input Windows dialog**: The Windows entry point gains a second input field — users provide both the FC_report `.xls` and the static data `.xlsx`.
- **Data merging in writer**: `write_xlsx` receives static data and injects tenant names into the Inquilini sheet column B, millesimali into Tabelle millesimali, costs into Tabella, and previous readings into Ripartizione.

## Capabilities

### New Capabilities

- `static-data-input`: Parsing, validation, and data model for the static xlsx input file (template structure, reader, models).
- `two-file-workflow`: Updated CLI/GUI entry point accepting two input files and merging both data sources into the output.

### Modified Capabilities

_None — no existing specs to modify._

## Impact

- **Code**: `__main__.py` (two-input CLI/dialog), new `static_reader.py` module, new models in `models.py`, `writer.py` (merge static data into output).
- **Files**: New `src/siemens_converter/static_data_template.xlsx` (committed), new `assets/static_data_2026.xlsx` (gitignored).
- **Dependencies**: No new runtime dependencies (openpyxl already handles xlsx reading).
- **PyInstaller**: `--add-data` must include the static data template alongside `template.xlsx`.
- **Tests**: New fixture for static data template, unit tests for static reader, integration tests for merged output.
- **User workflow**: Changes from single drag-and-drop to two-file input on Windows.
