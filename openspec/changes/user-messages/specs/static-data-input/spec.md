## MODIFIED Requirements

### Requirement: Static data reader

The system SHALL provide a `parse_static_data` function that reads a user-populated static data xlsx file and returns a structured `StaticData` object. Every error it raises SHALL be a `StaticDataError` whose message is plain Italian, names the file (and the sheet and cell where that applies) and says what to do next.

#### Scenario: Parse valid static data file

- **WHEN** a valid xlsx file matching the template structure is provided
- **THEN** the function SHALL return a `StaticData` object with all apartment data, costs, meter readings, and previous readings populated

#### Scenario: File that is not an xlsx

- **WHEN** the provided file cannot be opened as an xlsx workbook
- **THEN** the function SHALL raise a `StaticDataError` that names the file, says it is not a static data file and points to `static_data_template.xlsx` on the release page

#### Scenario: Missing required sheet

- **WHEN** the provided xlsx file is missing one of the five required sheets
- **THEN** the function SHALL raise a `StaticDataError` that names the file, says it is not a static data file, lists the missing sheets and points to `static_data_template.xlsx` on the release page

#### Scenario: Missing required cost row

- **WHEN** the Costi sheet is missing a required cost item
- **THEN** the function SHALL treat the missing cost as 0.0 (graceful default)

#### Scenario: Unknown cost label

- **WHEN** the Costi sheet has a label that is not one of the six cost items
- **THEN** the function SHALL raise a `StaticDataError` that names the file, the sheet and the cell, quotes the label as written and lists the accepted labels

#### Scenario: Value that is not a number

- **WHEN** a cell the reader converts to a number holds text that is not a number
- **THEN** the function SHALL raise a `StaticDataError` that names the file, the sheet and the cell and quotes the value found

#### Scenario: Blank number cell

- **WHEN** a cell the reader converts to a number is blank
- **THEN** the function SHALL read it as 0

#### Scenario: Empty apartment rows

- **WHEN** the Inquilini sheet has fewer rows than the number of apartments in the Siemens report
- **THEN** the function SHALL return only the apartments that are present, and the writer SHALL leave the tenant name of the remaining apartments empty
