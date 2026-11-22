# Proposal

## Why

The converter turns any FC report it can parse into a workbook and prints `Done`, even when the
report is missing a meter, has a blank reading, or mixes readings taken weeks apart
([issue #23](https://github.com/gjed/siemens-converter/issues/23)). A missing device leaves its
"Lettura finale" cell empty, so that apartment gets a negative consumption and a negative bill;
a blank reading becomes 0 or a traceback; one readout date is stamped on five cells that belong to
four different groups of meters. The people who run the tool are non-technical and have no way to
notice any of this.

## What Changes

- The conversion stops, with exit code 1 and no workbook, when the FC report does not contain both
  central meters (`Riscaldamento`, `Sanitario`) and one heat meter and one hot-water meter for each
  apartment 1–10. The message lists what is missing.
- The conversion stops in the same way when a reading the workbook needs is blank or is not a
  number, naming the device. A blank reading no longer becomes 0, and a blank central-meter reading
  no longer ends in a traceback.
- Each date cell of the "Ripartizione" sheet is dated from its own meters: `D28` from the
  `Riscaldamento` central meter, `D31` from the `Sanitario` central meter, `C76` and `C178` from
  the apartment heat meters, `C126` from the apartment hot-water meters. Until now all five cells
  carried the readout date of the first central meter in the file.
- The conversion warns, without failing, when a central meter and the apartment meters were read on
  different days (with the number of days), when apartment meters of the same kind were read on
  different days, and when a device's communication status is not `Ok` (naming the devices).
- Warnings are printed on the console before `Done` and, in the no-console build, appended to the
  final message box. Errors in the FC report reach the user as a message instead of a traceback.
- Every new message is in plain Italian, names devices by the description shown in the report and
  says what to do next.
- A report that passes every check produces the same workbook as before, apart from the corrected
  date cells.

## Capabilities

### New Capabilities

- `report-validation`: the checks the converter runs on an FC report before and while writing the
  workbook — completeness, required readings, per-meter readout dates, and the warnings shown to
  the user.

### Modified Capabilities

None.

## Impact

- `src/siemens_converter/parser.py`: completeness and blank-reading checks, communication status,
  warnings.
- `src/siemens_converter/models.py`: communication status on each meter, warnings on the parsed
  report, a pure helper for the readout date of a group of meters.
- `src/siemens_converter/writer.py`: per-meter date cells.
- `src/siemens_converter/__main__.py`: FC report errors shown as a message, warnings shown before
  `Done` and in the final message box.
- `tests/`: new tests for every check; `tests/test_writer.py::test_reading_date` is rewritten for
  the per-meter dates.
- No new dependency, no change to the static data file, the templates or the CLI arguments.
