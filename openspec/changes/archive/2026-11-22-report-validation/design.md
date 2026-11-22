# Design

## Context

See `proposal.md` for the motivation. The pipeline is `parse_fc_report` → `write_xlsx`, wired by
`__main__.py`. Today the parser classifies each device row and skips what it does not recognise,
the writer stamps one date on five cells, and `__main__.py` lets a parser exception escape as a
traceback.

The windowed Windows build has no console: the only thing the user sees is a message box. The
users are non-technical Italian speakers.

## Goals / Non-Goals

**Goals:**

- Refuse a report the template cannot turn into a correct bill, saying what is missing.
- Date every cell from the meters it describes.
- Tell the user about readings of different dates and devices in an abnormal state.
- Keep the workbook of a valid report unchanged apart from the date cells.

**Non-Goals:**

- Recognising a file that is not an FC report at all, and rewording the messages that already
  exist in English: both belong to
  [issue #26](https://github.com/gjed/siemens-converter/issues/26).
- Duplicate devices (two heat meters for one apartment) and apartments outside 1–10: the writer
  keeps its current behaviour.
- The user guide: the warnings are documented in a later change.

## Decisions

### The checks live in the parser, behind one entry point

`parse_fc_report` raises `ReportError` (a `ValueError`) for a refused report and returns a
`ParsedReport` whose new `warnings` list holds the warnings. Anything that calls the parser gets
the checks, and `__main__.py` only has to print what it receives.

A separate validation module called by `__main__.py` was rejected: a blank reading has to be
caught where the text is still available, before it becomes a number, so the parser would have to
carry the problem across anyway.

### One error lists every problem

The parser collects missing devices and unusable readings and raises once, at the end. A user who
gets one problem at a time has to download and drag the report again for each of them.

A device whose reading is blank is present, so it is reported under the blank reading and not
again as a missing device.

### The apartment set comes from the writer

Completeness is checked against the keys of `APT_ROWS` in `writer.py`: those are the rows the
template has, so the check cannot drift from what the writer fills.

### Readings that are not numbers are refused like blank ones

`heat_energy`, `aux1_volume` and `water_volume` go through one helper that reports an empty cell
and a cell that is not a number in the same list. Handling only the empty case would leave the
second one as an English `could not convert string to float` in the message box. `parse_decimal`
is left as it is: the required readings no longer reach it empty.

### Section dates: the most recent readout of the kind, plus a warning

`C76`/`C178` and `C126` each describe ten meters. When they do not share one readout date the cell
carries the most recent date and a warning names every meter read earlier, with its date.

- Failing instead was rejected: a meter that missed the last readout still has a usable, if older,
  reading, and the issue asks for a warning when dates differ.
- The most frequent date was rejected: it needs a tie-break and gives the same answer as the most
  recent date in the realistic case, a few meters left behind by the last readout.

The helper that computes it is a pure function in `models.py`, shared by the writer (cell values)
and the parser (warnings), so the two cannot disagree. It still returns a `datetime`, the type the
date cells already hold.

A readout date that is blank or not in `YYYY/MM/DD` form is ignored, as before: the cell stays
empty and no date warning is produced for that device.

### Central meters are compared with the meters they bill against

`Riscaldamento` is compared with the heat meters and `Sanitario` with the hot-water meters, because
those are the pairs the template subtracts to obtain the "energia involontaria". Each central meter
gets its own warning, since the two can be read on different days.

### Devices are named by kind and description

A heat meter and a hot-water meter of the same apartment have the same description in the report,
so the description alone is ambiguous. Messages use `contatore centrale «Riscaldamento»`,
`contacalorie «App, 01 …»` and `contatore dell'acqua calda «App, 01 …»`. A missing device has no
description to show and is named by apartment number: `contacalorie dell'appartamento 3`.

Readings are named in Italian with the report's column header in brackets
(`colonna «heat_energy»`), so the user can find the cell when opening the report in Excel.

### `communication_status` is stored on each meter

The three meter dataclasses get a `communication_status` field defaulting to `Ok`, so existing code
and tests that build meters by hand keep working. All devices that are not `Ok` go in one warning.

### `__main__.py` treats the FC report like the static data file

The `parse_fc_report` call is wrapped the way `parse_static_data` already is: any exception becomes
`Errore nel report FC: …` on the console or in a message box, exit code 1. `ReportError` messages
are written to read well after that prefix. Warnings are formatted once and used for both the
console and the final message box.

## Risks / Trade-offs

- [A report that used to convert is now refused] → Intended: such a report produced a workbook
  with negative bills. The message lists exactly what is missing.
- [A long list of warnings makes a tall message box] → Each kind of problem is one entry and
  devices are listed inside it; the worst realistic case is five entries.
- \[An exception that is not a `ReportError` still shows an English text after the Italian prefix\]
  → Out of scope here; covered by issue #26.
- \[The parser imports the writer for `APT_ROWS`\] → One constant; the alternative is a second copy
  of the apartment list.
