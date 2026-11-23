## MODIFIED Requirements

### Requirement: Errors and warnings reach the user as readable messages

The system SHALL show an FC report error and a static data error as a message and exit with code 1, never as a traceback: on the console when there is one, in a message box otherwise. Warnings SHALL be printed on the console before the final `Fatto:` line and, when there is no console, appended to the final message box. Messages SHALL be in Italian, name devices by their description in the report, and say what to do next.

#### Scenario: Refused report on the console

- **WHEN** the converter is run from a console on a report that fails a check
- **THEN** the error message is printed, the exit code is 1 and no traceback is shown

#### Scenario: Refused report without a console

- **WHEN** the no-console build is run on a report that fails a check
- **THEN** the error message is shown in a message box and the exit code is 1

#### Scenario: Warnings on the console

- **WHEN** the converter is run from a console on a report that produces warnings
- **THEN** the warnings are printed before the `Fatto:` line and the exit code is 0

#### Scenario: Warnings without a console

- **WHEN** the no-console build converts a report that produces warnings
- **THEN** the final message box contains the output file name followed by the warnings

#### Scenario: No warnings

- **WHEN** the report produces no warning
- **THEN** the console output and the final message box contain no warning section

## ADDED Requirements

### Requirement: Files that are not an FC report are refused

The system SHALL refuse a file that is not an FC report with a `ReportError` that names the file and asks to choose the file exported by the Siemens system, before any other parsing error can surface. A device whose description has no apartment number SHALL be refused with a message that names the device.

#### Scenario: Binary file

- **WHEN** the file is not UTF-8 text, for example the static data `.xlsx`
- **THEN** the conversion fails with an Italian message that names the file and says it is not an FC report
- **AND** no workbook is written

#### Scenario: Text without a table

- **WHEN** the file is UTF-8 text without a `<table>`
- **THEN** the conversion fails with an Italian message that names the file and says it contains no table

#### Scenario: Too few rows

- **WHEN** the table has fewer rows than the metadata and header rows need
- **THEN** the conversion fails with an Italian message that names the file and says the table is too short

#### Scenario: Too few columns

- **WHEN** a row of the table has fewer columns than the report layout needs
- **THEN** the conversion fails with an Italian message that names the file and says the table has too few columns

#### Scenario: Device without an apartment number

- **WHEN** the description of an apartment device does not contain an apartment number
- **THEN** the conversion fails with an Italian message that names the device and shows how the description should begin
