## ADDED Requirements

### Requirement: Static data template structure

The system SHALL provide a blank xlsx template file (`static_data_template.xlsx`) with the following named sheets and column structure:

- **Inquilini** sheet: columns `Appartamento` (int, apartment number 1-10), `Proprietario` (str, owner name), `Inquilino` (str, tenant name/description for Ripartizione rows).
- **Millesimali** sheet: columns `Appartamento` (int), `Subalterno` (int), `Energia_riscaldamento_kWh` (float), `Energia_ACS_kWh` (float).
- **Costi** sheet: columns `Voce` (str, cost item label), `Importo` (float, amount in EUR). Required rows: `Energia elettrica`, `Gas metano`, `Acqua condominio`, `Conduzione e manutenzione`, `Contabilizzazione`, `Acqua sanitaria manutenzione`.
- **Contatori** sheet: columns `Contatore` (str), `Unita` (str), `Lettura_iniziale` (float), `Lettura_finale` (float). Required rows: `Energia elettrica CT`, `Gas metano CT`, `Acqua generale`.
- **Letture_precedenti** sheet: columns `Appartamento` (int), `Riscaldamento_kWh` (int, previous heat allocator reading), `ACS_m3` (float, previous water meter reading), `AFS_m3` (float, previous cold water reading).

The template SHALL be committed to the repository and bundled in PyInstaller builds.

#### Scenario: Template file exists and is valid xlsx
- **WHEN** the application is installed or built
- **THEN** the file `static_data_template.xlsx` SHALL be available as a package resource alongside `template.xlsx`

#### Scenario: Template has correct sheet structure
- **WHEN** the template file is opened
- **THEN** it SHALL contain exactly the five sheets listed above with correct column headers in row 1

### Requirement: Static data reader

The system SHALL provide a `parse_static_data` function that reads a user-populated static data xlsx file and returns a structured `StaticData` object.

#### Scenario: Parse valid static data file
- **WHEN** a valid xlsx file matching the template structure is provided
- **THEN** the function SHALL return a `StaticData` object with all apartment data, costs, meter readings, and previous readings populated

#### Scenario: Missing required sheet
- **WHEN** the provided xlsx file is missing one of the five required sheets
- **THEN** the function SHALL raise a `ValueError` with a message identifying the missing sheet

#### Scenario: Missing required cost row
- **WHEN** the Costi sheet is missing a required cost item
- **THEN** the function SHALL treat the missing cost as 0.0 (graceful default)

#### Scenario: Empty apartment rows
- **WHEN** the Inquilini sheet has fewer rows than the number of apartments in the Siemens report
- **THEN** the function SHALL populate available apartments and leave remaining apartments with empty strings

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
- **Tabelle millesimali sheet**: `Subalterno` in column B, `Energia_riscaldamento_kWh` in column C, `Energia_ACS_kWh` in column E — for rows 4-13 (apartments 1-10).
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
- **THEN** the output SHALL be identical to the current behavior
