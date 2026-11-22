# two-file-workflow Specification

## Purpose

Define how the converter receives its two inputs, the required FC_report `.xls` and the optional static data `.xlsx`, in each of the ways it is started: from the command line, by dragging files onto the Windows executable, and through file dialogs when the executable is opened without files.

## Requirements

### Requirement: CLI accepts two input files

The CLI entry point SHALL accept two positional arguments: the FC_report `.xls` path and the static data `.xlsx` path.

#### Scenario: Both files provided via CLI

- **WHEN** the user runs `python -m siemens_converter <FC_report.xls> <static_data.xlsx>`
- **THEN** the system SHALL parse both files and produce the merged output

#### Scenario: Only FC_report provided (backward compatible)

- **WHEN** the user runs `python -m siemens_converter <FC_report.xls>` without a second argument
- **THEN** the system SHALL produce the output from the FC_report alone, with no static data merged

#### Scenario: Static data file not found

- **WHEN** the user provides a static data path that does not exist
- **THEN** the system SHALL display an error message and exit with code 1

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

When two files are dragged onto the `.exe`, the system SHALL route them by extension regardless of argument order: the `.xlsx` file is static data, the other is the FC_report.

#### Scenario: Two files dragged onto exe

- **WHEN** the user drags both an FC_report `.xls` and a static data `.xlsx` onto the executable
- **THEN** the system SHALL process both and produce merged output

#### Scenario: One file dragged onto exe

- **WHEN** the user drags only an FC_report `.xls` onto the executable
- **THEN** the system SHALL process only the FC_report (backward compatible)
