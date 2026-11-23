# Proposal

## Why

The converter only produces correct bills for the exercise whose readings are baked into
`template.xlsx` ([issue #24](https://github.com/gjed/siemens-converter/issues/24)). The start
readings of the two central heat meters, the exercise start date, the year in the labels, the start
readings of the 30 apartment meters and the tenancy days of a real building are literals in the
template, and the static data file has no place for them. In 2027 nothing fails and the bills are
wrong. Without a static data file the output is a sheet of `#DIV/0!` and the converter prints
`Fatto`. The public repository also carries real readings and millesimali in the template and in
the test fixture.

The static data reader also accepts malformed files and turns the gaps into zeros
([issue #25](https://github.com/gjed/siemens-converter/issues/25)): a blank millesimi cell, a
duplicated or missing apartment row, a missing cost row and a stray text row all end in a bill
computed from zeros, and nothing compares the static data with the report, so a file from the wrong
year gives negative consumptions. The output is a set of bills, so a plausible-looking wrong number
is worse than a refusal.

## What Changes

- **BREAKING** The static data file is required. A run with only the FC_report stops with an
  Italian message that points to `static_data_template.xlsx`; the second file dialog no longer
  offers to skip it and the usage text no longer says "facoltativo".
- **BREAKING** `template.xlsx` holds no exercise data: every year-specific literal is cleared
  (central meter start readings and dates, the year in two labels, the 30 apartment start readings,
  the cold water section date, the tenancy leftovers and the `??/??/????` placeholder), keeping every
  formula. A test fails when a numeric or date literal appears in the first sheet outside the
  physical constants `C40`, `C48`, `C123` and `C124`.
- **BREAKING** The sheets of the output are named `Ripartizione` and `Tabella`, fixed, instead of
  `Ripartizione_Termo 02.03.2026` and `Tabella_2026`.
- **BREAKING** `static_data_template.xlsx` gains a sixth sheet, `Esercizio`, with the exercise year,
  the exercise start date, the start readings of the heating and hot water (ACS) central heat
  meters and, optionally, a manual final reading and date for each. Static data files made with the
  previous template are refused with a message that names the missing sheet.
- The writer puts those values into the cells cleared in the template, writes the year in the two
  labels, and writes `Proprietario` into column A of the `Inquilini` sheet when it is given. A manual
  final reading of a central meter overrides the one of the FC_report.
- The reader refuses malformed files: each per-apartment sheet must list apartments 1 to 10 exactly
  once, a blank or non-numeric number cell is refused, a missing cost or meter row is refused (a real
  zero is entered as `0`), and any row that is neither data nor `TOTALE` is refused on every sheet.
- After both inputs are read, the conversion is refused when a final reading of the FC_report is
  lower than the start reading in the static data, naming the meter.
- `tests/fixtures/static_data_TEST.xlsx` is regenerated with invented values consistent with the
  test report.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

- `static-data-input`: the template gains the `Esercizio` sheet, the reader refuses blank cells,
  duplicate, missing and stray rows and missing cost rows, the cross-check against the report, and
  the writer injects the exercise data and the owner names.
- `two-file-workflow`: the static data file is required, in the command line, in drag and drop and
  in the second file dialog.
- `riparto-calculation`: the output sheets have neutral fixed names, and the template holds no
  exercise data.

## Impact

- `src/siemens_converter/template.xlsx`, `src/siemens_converter/static_data_template.xlsx`:
  cleared literals, renamed sheets, new `Esercizio` sheet.
- `src/siemens_converter/models.py`, `static_reader.py`, `writer.py`, `__main__.py`: new fields,
  stricter reader, cross-check, required static data.
- `tests/`: new tests for the template literals and for each refusal; assertions on old fixture
  values and sheet names updated; `tests/fixtures/static_data_TEST.xlsx` regenerated.
- `docs/uso.md`, `README.md`, `docs/development.md`: the second file is required, the new fields and
  the new errors.
- Users must download the new `static_data_template.xlsx` from the next release.
