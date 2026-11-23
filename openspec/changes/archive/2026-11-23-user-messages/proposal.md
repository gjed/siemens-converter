# Proposal

## Why

The people who run the converter are non-technical Italian speakers who drag files onto a Windows
`.exe` and read whatever box appears. Several messages are in English, use internal terms, or are a
raw Python traceback ([issue #26](https://github.com/gjed/siemens-converter/issues/26)): dragging
only the static data `.xlsx` ends in `UnicodeDecodeError`, the static data errors are half Italian
and half English, the usage text, `File not found`, `Parsing`, `Writing` and `Done` are English
only, and the last-resort handler opens with a stack trace. The user guide explains none of the
report errors and warnings added by the `report-validation` capability
([issue #30](https://github.com/gjed/siemens-converter/issues/30)), and its "Problemi e soluzioni"
section has to be rewritten against the new wording anyway.

## What Changes

- Every message a user can see, on the console and in the message boxes, is plain Italian and says
  what is wrong, in which file (plus sheet and cell where that applies) and what to do next. No
  exception names, no English, no column indexes.
- When the files given are only an `.xlsx`, two `.xlsx` or two files that are not `.xlsx`, the
  message says which file is missing or doubled instead of parsing the static data as an
  FC_report.
- A file that is not an FC report (not text, no table, too few rows or columns) is refused with a
  message naming the file, before any other parsing error can surface. The apartment-number error
  of a device description is refused the same way, naming the device.
- A file that is not a static data file (it does not open as `.xlsx`, or lacks the required sheets)
  is refused with a message that points to `static_data_template.xlsx` on the release page. The
  unknown-cost-label error quotes the label as written and lists the accepted ones. A non-numeric
  value where a number is read is refused with a message naming file, sheet and cell. Which inputs
  are accepted does not change.
- The last-resort handler of the executable leads with one Italian sentence saying that the
  conversion failed and whom to send the details to; the technical details follow below it.
- The user guide rewrites every entry of "Problemi e soluzioni" for the new wording and adds an
  entry for the report refusal and one for each warning. It says that warnings do not stop the
  conversion and appear in the final message.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

- `two-file-workflow`: file routing errors, and the language of the messages shown while running.
- `static-data-input`: the errors of the static data reader are Italian and name file, sheet and
  cell.
- `report-validation`: files that are not an FC report are refused; the wording shown on the
  console after the report checks.

## Impact

- `src/siemens_converter/__main__.py`: routing, Italian messages, only the two handled error types
  are shown as messages.
- `src/siemens_converter/parser.py`: `ReportError` for files that are not an FC report.
- `src/siemens_converter/static_reader.py`: new `StaticDataError`, Italian messages, cell-level
  number errors.
- `src/siemens_converter/models.py`: Italian wording of the apartment-number error.
- `scripts/pyinstaller_entry.py`: last-resort handler.
- `tests/`: wording of each message is asserted for its trigger.
- `docs/uso.md`, `README.md`, `docs/development.md`: quote the new wording.
- No new dependency, no change to the templates, the CLI arguments or the output workbook.
