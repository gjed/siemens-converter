# Proposal

## Why

The output workbook distributes more than the real costs as soon as the electricity meter's
initial reading is not 0, and the converter still reports success
([issue #22](https://github.com/gjed/siemens-converter/issues/22)). Two more cells in the same
template show or use values that only hold for the 2026 workbook the template was copied from,
and no test evaluates a formula, so none of this is caught today.

## What Changes

All cells are on the first sheet of `template.xlsx` (`Ripartizione_Termo 02.03.2026`).

- `C25` (electricity consumption) reads the consumption `Tabella_2026!H20` instead of the final
  reading `Tabella_2026!G20`. The electricity cost distributed to the apartments now equals the
  electricity cost entered, whatever the initial reading.
- `C23` (gas final reading) reads the final reading `Tabella_2026!G21` instead of the consumption
  `Tabella_2026!H21`. The printed gas consumption and unit cost are right when the initial
  reading is not 0. Apartment totals do not depend on these cells.
- `C63` and `C64` (`kcli` and `kacs`, the split of the management costs between heating and hot
  water) reference `'Tabelle millesimali'!C18` and `E18` instead of holding the 2026 values as
  literals. The split follows the millesimali in the static data.
- A recalculation test generates workbooks through the real pipeline, recalculates them with
  LibreOffice and asserts on the computed values. CI installs LibreOffice and fails if the test
  cannot run.

With an initial reading of 0 and the 2026 millesimali the results are the same as before. No other
template cell changes: the remaining pasted values belong to
[issue #24](https://github.com/gjed/siemens-converter/issues/24).

## Capabilities

### New Capabilities

- `riparto-calculation`: what the formulas of the generated workbook compute from the report and
  the static data: the costs they distribute, the shares they use and the check that the
  distributed total equals the costs.

### Modified Capabilities

None.

## Impact

- `src/siemens_converter/template.xlsx`: four cells.
- `tests/test_recalc.py` (new) and `tests/test_template.py`.
- `.github/workflows/ci.yml` and `.github/workflows/release.yml`: the test jobs install
  LibreOffice Calc.
- No Python source, dependency or CLI change. Running the full suite locally needs LibreOffice;
  without it the recalculation tests are skipped.
