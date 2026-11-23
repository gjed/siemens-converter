# static-data-input Specification

## Purpose

Define the optional static data workbook through which a building administrator supplies what the Siemens FC_report does not contain: tenant names, millesimali, yearly costs, main utility meter readings and each apartment's previous readings. The capability covers the blank template that is filled in, the reader that turns a filled workbook into a `StaticData` object, the data model, and the cells of the output workbook that receive the data.

## Requirements

### Requirement: Static data template structure

The system SHALL provide a blank xlsx template file (`static_data_template.xlsx`) with the following named sheets and column structure. Columns are read by position; the header text of the template is given for each column.

- **Inquilini** sheet: column A `App.` (int, apartment number 1-10), column B `Proprietario / Condomino` (str, owner name), column C `Inquilino / Conduttore` (str, tenant name/description for Ripartizione rows).
- **Millesimali** sheet: column A `App.` (int), column B `Subal-terno` (int), column C `Energia utile riscaldamento` (float, kWh), column D `Millesimi risc.` (formula, not read), column E `Energia utile ACS` (float, kWh), column F `Millesimi ACS` (formula, not read). Row 3 holds the units and row 4 the `TOTALE` formulas.
- **Costi** sheet: column A `Voce di spesa` (str, cost item label), column B `€ / anno` (float, amount in EUR). Required rows: `Energia elettrica`, `Gas metano`, `Acqua condominio`, `Conduzione e manutenzione`, `Contabilizzazione`, `Acqua sanitaria manutenzione`, followed by a `TOTALE` formula row.
- **Contatori** sheet: column A `Contatore` (str), column B `Unità` (str), column C `Lettura iniziale` (float), column D `Lettura finale` (float), column E `Consumo` (formula, not read). Required rows: `ENERGIA ELETTRICA CT`, `GAS METANO CT`, `ACQUA GENERALE`.
- **Letture_precedenti** sheet: column A `App.` (int), column B `Riscaldamento lettura prec.` (int, previous heat allocator reading in kWh), column C `ACS lettura prec.` (float, previous water meter reading in m³), column D `AFS lettura prec.` (float, previous cold water reading in m³).

The template SHALL be committed to the repository, declared as package data, and published as a separate asset of every GitHub release. PyInstaller builds bundle only `template.xlsx`; the static data template SHALL NOT be required inside the executable.

#### Scenario: Template file exists and is valid xlsx

- **WHEN** the application is installed as a Python package
- **THEN** the file `static_data_template.xlsx` SHALL be available as a package resource alongside `template.xlsx`

#### Scenario: Template is published with the release

- **WHEN** a tagged release is built
- **THEN** the file `static_data_template.xlsx` SHALL be attached to the release as its own asset, next to the executables and not inside them

#### Scenario: Template has correct sheet structure

- **WHEN** the template file is opened
- **THEN** it SHALL contain exactly the five sheets listed above, with a section title in row 1, column headers in row 2, and data rows starting at row 3, except for the Millesimali sheet, where data rows start at row 5

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

### Requirement: Static data model

The system SHALL define dataclasses for static data following the project convention of pure models with no I/O.

- `ApartmentInfo`: apartment_number (int), proprietario (str), inquilino (str)
- `Millesimali`: apartment_number (int), subalterno (int), heat_energy_kwh (float), water_energy_kwh (float)
- `CostItem`: label (str), amount (float)
- `MeterReading`: name (str), unit (str), initial (float), final (float)
- `PreviousReading`: apartment_number (int), heat_kwh (int), water_m3 (float), cold_water_m3 (float)
- `StaticData`: apartments (list[ApartmentInfo]), millesimali (list[Millesimali]), costs (list[CostItem]), meters (list[MeterReading]), previous_readings (list[PreviousReading])

#### Scenario: StaticData is a pure dataclass

- **WHEN** `StaticData` is instantiated
- **THEN** it SHALL contain no I/O logic and no side effects

### Requirement: Writer merges static data into output

The `write_xlsx` function SHALL accept an optional `StaticData` parameter. When provided, it SHALL inject:

- **Inquilini sheet column B**: tenant names (inquilino field) from StaticData, matching by apartment number.
- **Tabelle millesimali sheet**: `Subalterno` in column B, heating energy (kWh) in column C, ACS energy (kWh) in column E — for rows 4-13 (apartments 1-10).
- **Tabella_2026 sheet**: cost amounts into cells E3-E8 (matching by cost item label), and meter readings into cells F20-F22 (initial) and G20-G22 (final).
- **Ripartizione sheet**: previous heat readings into column B of heat rows (B78, B80, B82, ...), previous water readings into column B of water rows, previous cold water readings into column B of AFS rows.

#### Scenario: Output with static data has tenant names

- **WHEN** `write_xlsx` is called with a valid `StaticData` containing tenant names
- **THEN** the Inquilini sheet column B SHALL contain the tenant names for each apartment

#### Scenario: Output with static data has millesimali

- **WHEN** `write_xlsx` is called with `StaticData` containing millesimali
- **THEN** the Tabelle millesimali sheet SHALL have energy values in columns C and E, and computed millesimi formulas in D and F SHALL produce correct results

#### Scenario: Output with static data has costs

- **WHEN** `write_xlsx` is called with `StaticData` containing costs
- **THEN** the Tabella_2026 sheet SHALL have cost amounts in the correct cells

#### Scenario: Output without static data is unchanged

- **WHEN** `write_xlsx` is called without `StaticData` (None)
- **THEN** the output SHALL hold the FC_report readings only: no static data cell SHALL be written and column B of the Inquilini sheet SHALL be empty
