## MODIFIED Requirements

### Requirement: CLI accepts two input files

The CLI entry point SHALL accept two positional arguments: the FC_report `.xls` path and the static data `.xlsx` path, both required. A path that does not exist SHALL be reported in Italian, naming the file and what to do, with exit code 1.

#### Scenario: Both files provided via CLI

- **WHEN** the user runs `python -m siemens_converter <FC_report.xls> <static_data.xlsx>`
- **THEN** the system SHALL parse both files and produce the merged output

#### Scenario: Only FC_report provided (backward compatible)

- **WHEN** the user runs `python -m siemens_converter <FC_report.xls>` without a second argument
- **THEN** the system SHALL display an Italian message that names the FC_report, says the static data file is missing and points to `static_data_template.xlsx` on the release page, exit with code 1 and write no output

#### Scenario: FC_report not found

- **WHEN** the user provides an FC_report path that does not exist
- **THEN** the system SHALL display an Italian message that names the file, says it does not exist and asks to check that it was not moved or renamed, and exit with code 1

#### Scenario: Static data file not found

- **WHEN** the user provides a static data path that does not exist
- **THEN** the system SHALL display an Italian message that names the file, says it does not exist and asks to check that it was not moved or renamed, and exit with code 1

#### Scenario: No arguments on a console

- **WHEN** the user runs the CLI from a console without arguments
- **THEN** the system SHALL print the usage text in Italian, which marks both files as required, and exit with code 1

### Requirement: Windows dialog with two sequential file selections

When running in no-console mode (PyInstaller `--noconsole`) and started without file arguments, the system SHALL show two sequential Windows file-open dialogs:

1. First dialog: FC_report `.xls` file selection (required — cancelling exits the application)
1. Second dialog: static data `.xlsx` file selection (required — cancelling is refused with the message that the static data file is missing)

#### Scenario: User selects both files via dialog

- **WHEN** the user is running the `.exe` without console and selects both files
- **THEN** the system SHALL process both files and produce merged output, showing a success message box

#### Scenario: User selects only FC_report via dialog

- **WHEN** the user selects an FC_report but cancels the static data dialog
- **THEN** the system SHALL display an Italian message box that says the static data file is missing and points to `static_data_template.xlsx` on the release page, and write no output

#### Scenario: User cancels FC_report dialog

- **WHEN** the user cancels the first file dialog (FC_report)
- **THEN** the system SHALL exit without processing

### Requirement: Drag-and-drop supports two files

When two files are dragged onto the `.exe`, the system SHALL route them by extension regardless of argument order: the `.xlsx` file is static data, the other is the FC_report. When the files given are not exactly one `.xlsx` and one other file, the system SHALL say in Italian which file is missing or doubled, and SHALL NOT parse any file.

#### Scenario: Two files dragged onto exe

- **WHEN** the user drags both an FC_report `.xls` and a static data `.xlsx` onto the executable
- **THEN** the system SHALL process both and produce merged output

#### Scenario: One file dragged onto exe

- **WHEN** the user drags only an FC_report `.xls` onto the executable
- **THEN** the system SHALL display an Italian message that names the file, says the static data file is missing and points to `static_data_template.xlsx` on the release page, and exit with code 1

#### Scenario: Only the static data file given

- **WHEN** the only file given is an `.xlsx`
- **THEN** the system SHALL display an Italian message that names the file, says the FC_report is missing and asks to choose it as well, and exit with code 1

#### Scenario: Two `.xlsx` files given

- **WHEN** both files given are `.xlsx`
- **THEN** the system SHALL display an Italian message that names both files, says that only one static data file is accepted and that the FC_report is missing, and exit with code 1

#### Scenario: Two files that are not `.xlsx` given

- **WHEN** neither file given is an `.xlsx`
- **THEN** the system SHALL display an Italian message that names both files, says that only one FC_report is accepted and that the static data file must be an `.xlsx`, and exit with code 1
