# two-file-workflow Specification

## Purpose

Define how the converter receives its two inputs, the required FC_report `.xls` and the optional static data `.xlsx`, in each of the ways it is started: from the command line, by dragging files onto the Windows executable, and through file dialogs when the executable is opened without files.

## Requirements

### Requirement: CLI accepts two input files

The CLI entry point SHALL accept two positional arguments: the FC_report `.xls` path and the static data `.xlsx` path. A path that does not exist SHALL be reported in Italian, naming the file and what to do, with exit code 1.

#### Scenario: Both files provided via CLI

- **WHEN** the user runs `python -m siemens_converter <FC_report.xls> <static_data.xlsx>`
- **THEN** the system SHALL parse both files and produce the merged output

#### Scenario: Only FC_report provided (backward compatible)

- **WHEN** the user runs `python -m siemens_converter <FC_report.xls>` without a second argument
- **THEN** the system SHALL produce the output from the FC_report alone, with no static data merged

#### Scenario: FC_report not found

- **WHEN** the user provides an FC_report path that does not exist
- **THEN** the system SHALL display an Italian message that names the file, says it does not exist and asks to check that it was not moved or renamed, and exit with code 1

#### Scenario: Static data file not found

- **WHEN** the user provides a static data path that does not exist
- **THEN** the system SHALL display an Italian message that names the file, says it does not exist and asks to check that it was not moved or renamed, and exit with code 1

#### Scenario: No arguments on a console

- **WHEN** the user runs the CLI from a console without arguments
- **THEN** the system SHALL print the usage text in Italian and exit with code 1

### Requirement: Windows dialog with two sequential file selections

When running in no-console mode (PyInstaller `--noconsole`) and started without file arguments, the system SHALL show two sequential Windows file-open dialogs:

1. First dialog: FC_report `.xls` file selection (required — cancelling exits the application)
1. Second dialog: static data `.xlsx` file selection (optional — cancelling skips static data and proceeds with FC_report only)

#### Scenario: User selects both files via dialog

- **WHEN** the user is running the `.exe` without console and selects both files
- **THEN** the system SHALL process both files and produce merged output, showing a success message box

#### Scenario: User selects only FC_report via dialog

- **WHEN** the user selects an FC_report but cancels the static data dialog
- **THEN** the system SHALL process only the FC_report (backward compatible behavior)

#### Scenario: User cancels FC_report dialog

- **WHEN** the user cancels the first file dialog (FC_report)
- **THEN** the system SHALL exit without processing

### Requirement: Drag-and-drop supports two files

When two files are dragged onto the `.exe`, the system SHALL route them by extension regardless of argument order: the `.xlsx` file is static data, the other is the FC_report. When the files given are not exactly one `.xlsx` and one other file, or are only one `.xlsx`, the system SHALL say in Italian which file is missing or doubled, and SHALL NOT parse any file.

#### Scenario: Two files dragged onto exe

- **WHEN** the user drags both an FC_report `.xls` and a static data `.xlsx` onto the executable
- **THEN** the system SHALL process both and produce merged output

#### Scenario: One file dragged onto exe

- **WHEN** the user drags only an FC_report `.xls` onto the executable
- **THEN** the system SHALL process only the FC_report (backward compatible)

#### Scenario: Only the static data file given

- **WHEN** the only file given is an `.xlsx`
- **THEN** the system SHALL display an Italian message that names the file, says the FC_report is missing and asks to choose it as well, and exit with code 1

#### Scenario: Two `.xlsx` files given

- **WHEN** both files given are `.xlsx`
- **THEN** the system SHALL display an Italian message that names both files, says that only one static data file is accepted and that the FC_report is missing, and exit with code 1

#### Scenario: Two files that are not `.xlsx` given

- **WHEN** neither file given is an `.xlsx`
- **THEN** the system SHALL display an Italian message that names both files, says that only one FC_report is accepted and that the static data file must be an `.xlsx`, and exit with code 1

### Requirement: Every message shown to the user is plain Italian

Every message the system shows, on the console and in message boxes, SHALL be in plain Italian and SHALL say what is wrong, in which file (plus sheet and cell where that applies) and what to do next. It SHALL NOT contain an exception name, a column index or English text. This covers the usage text, the progress lines, the final line, the prompt to close the console and the message shown when a file dialog cannot be opened.

#### Scenario: Progress and final lines on a console

- **WHEN** a conversion succeeds on a console
- **THEN** the progress lines and the final line are in Italian and the final line starts with `Fatto:`

#### Scenario: File dialog cannot be opened

- **WHEN** the Windows file dialog cannot be opened
- **THEN** the message box says in Italian that the dialog could not be opened and asks to drag the files onto the program instead, without the text of the underlying error

#### Scenario: Unexpected error

- **WHEN** an unexpected error ends the conversion of the executable
- **THEN** the console or the message box opens with one Italian sentence saying that the conversion failed and to send the details to whoever provided the program, followed by the technical details
