# Design

## Context

Everything that changes from one exercise to the next is either a literal in `template.xlsx` or
missing from the static data file. The reader (`static_reader.py`) reads blanks as 0, stops or skips
at the first non-numeric row depending on the sheet, and the writer overwrites duplicated rows. See
proposal.md for the motivation. The maintainer settled the three open questions of issue #24:

1. The static data file is required.
1. The static data may carry an optional manual final reading and date for each central heat meter,
   which override the FC_report.
1. Sheet names are neutral and fixed.

## Goals / Non-Goals

**Goals:**

- The template carries no exercise data and the static data file carries all of it.
- A malformed static data file, or one of another exercise, is refused before any output exists.

**Non-Goals:**

- Splitting an apartment's bill between two tenants who change during the year, and importing the
  tenancy days: the cells stay empty and are filled by hand in the workbook.
- Generating next year's static data file from a year-end run.
- Importing the reading and cost dates of `Tabella`: those cells stay empty.

## Decisions

- **The exercise data lives in a new `Esercizio` sheet, not in `Contatori`.** `Contatori` describes
  the utility meters of the plant (electricity, gas, water) with a start and a final reading each.
  The year, the start date and the two heat meters have different rows and an optional final
  reading and date. One new sheet with one row per value keeps the existing sheets unchanged. The
  cost is that a file made with the old template is refused with "manca il foglio «Esercizio»",
  which is wanted: its start readings and millesimali would be of the wrong exercise anyway.
- **Fixed rows, read by position**, like the other sheets: year in `C3`, start date in `C4`, heating
  meter in row 5, ACS meter in row 6. The label in column A is for the human only.
- **The year is a four-digit integer and only the last two digits appear in the labels**
  (`FINE GESTIONE 26`), as in the template today.
- **A manual final reading needs its date and a date needs its reading.** A reading without its date
  would mix two dates in the sheet and a date without a reading has nothing to apply to.
- **One helper walks the rows of every sheet.** It skips entirely blank rows and the `TOTALE` row
  and hands everything else to the sheet's reader, which refuses what it does not recognise. The
  per-apartment sheets share one function that checks 1–10 exactly once and returns the rows by
  apartment, so Inquilini, Millesimali and Letture_precedenti cannot disagree again. Costi and
  Contatori share a second one for the rows with a fixed label (missing, unknown, duplicate).
- **Millesimali is read from row 5.** The template has the units in row 3 and `TOTALE` in row 4;
  the old reader skipped them because it skipped every non-numeric row. With the strict rule they
  would be refused, so the data rows are read from where the spec says they start.
- **`_number` refuses a blank.** It reads a cell that must hold a number and no longer defaults to
  0; a real zero is typed `0`. Every refusal raises `StaticDataError` with the finished message and
  reuses `_TEMPLATE_ADVICE` and `_quoted`.
- **The cross-check is a function of `static_reader`, called by `write_xlsx` before it touches the
  output file.** It needs the parsed report and the parsed static data, and it is a property of the
  static data ("these are not the start readings of this report"). The writer is the first place
  that has both; `__main__` only prints the message. Running it first leaves no half-written output
  file. It compares each central meter (after the manual override) and each apartment's heat, hot
  water and cold water reading, lists all offenders in one message and treats equality as valid.
- **Required static data is enforced in `_route_files`**, which already decides what the files are,
  so drag and drop, the command line and the dialogs behave the same. Cancelling the second dialog
  passes only the FC_report to `_route_files`, which prints the same "manca il file dei dati statici"
  message.
- **Sheet names are changed once, in the template.** A one-off script renamed the sheets and
  rewrote every cross-sheet formula and the two sheet-scoped defined names, because openpyxl does
  not. The recalculation test (`tests/test_recalc.py`) proves nothing broke: every total is
  unchanged and no cell holds `#REF!`. The writer has no rewriting code.
- **`template.xlsx` is cleared by the same kind of one-off script** that sets the listed cells to
  empty and leaves formulas, and a test pins the result: no numeric or date literal in the first
  sheet outside `C40`, `C48`, `C123` and `C124`. `B77`, a stray `0` in a heading row that the
  literal scan finds, is cleared too.
- **The Contatori and Costi rows are required** like the apartment rows. The issue names the cost
  rows; the three meter rows have the same failure (a missing row leaves the cell blank and the gas
  cost divides by zero) and the same fix.

## Risks / Trade-offs

- [Users keep the old static data file] → it is refused with a message that names the missing
  sheet and points to the new template, and the release notes and `docs/uso.md` say so.
- [The static data file is now required, which breaks the single-file drag and drop] → decided by
  the maintainer; the refusal says what to choose.
- [Strict rows refuse a file that used to work, for example a note typed under the table] → the
  message names the cell and says what the sheet accepts.
- [Static data that is wrong but never above the report is accepted] → the check only catches readings that went backwards; it cannot know the year of the report.

## Migration Plan

Ship the new `static_data_template.xlsx` with the release. Users copy their values into it.
Rollback is reverting the release; the old executable and the old template work together.
