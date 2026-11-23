## MODIFIED Requirements

### Requirement: Static data template structure

The system SHALL provide a blank xlsx template file (`static_data_template.xlsx`) with the following named sheets and column structure. Columns are read by position; the header text of the template is given for each column.

- **Inquilini** sheet: column A `App.` (int, apartment number 1-10), column B `Proprietario / Condomino` (str, owner name), column C `Inquilino / Conduttore` (str, tenant name/description for Ripartizione rows).
- **Millesimali** sheet: column A `App.` (int), column B `Subal-terno` (int), column C `Energia utile riscaldamento` (float, kWh), column D `Millesimi risc.` (formula, not read), column E `Energia utile ACS` (float, kWh), column F `Millesimi ACS` (formula, not read). Row 3 holds the units and row 4 the `TOTALE` formulas.
- **Costi** sheet: column A `Voce di spesa` (str, cost item label), column B `€ / anno` (float, amount in EUR). Required rows: `Energia elettrica`, `Gas metano`, `Acqua condominio`, `Conduzione e manutenzione`, `Contabilizzazione`, `Acqua sanitaria manutenzione`, followed by a `TOTALE` formula row.
- **Contatori** sheet: column A `Contatore` (str), column B `Unità` (str), column C `Lettura iniziale` (float), column D `Lettura finale` (float), column E `Consumo` (formula, not read). Required rows: `ENERGIA ELETTRICA CT`, `GAS METANO CT`, `ACQUA GENERALE`.
- **Letture_precedenti** sheet: column A `App.` (int), column B `Riscaldamento lettura prec.` (int, previous heat allocator reading in kWh), column C `ACS lettura prec.` (float, previous water meter reading in m³), column D `AFS lettura prec.` (float, previous cold water reading in m³).
- **Esercizio** sheet: column A `Dato`, column B `Unità`, column C `Valore (lettura iniziale per i contatori)`, column D `Lettura finale (facoltativa)`, column E `Data lettura finale (facoltativa)`. Rows 3 to 6, in this order: `Anno dell'esercizio` (int, four digits, in column C), `Data di inizio dell'esercizio` (date, in column C), `Contacalorie centrale RISCALDAMENTO` and `Contacalorie centrale ACS` (kWh: start reading in column C, optional manual final reading in column D and its date in column E).

The template SHALL be committed to the repository, declared as package data, and published as a separate asset of every GitHub release. PyInstaller builds bundle only `template.xlsx`; the static data template SHALL NOT be required inside the executable.

#### Scenario: Template file exists and is valid xlsx

- **WHEN** the application is installed as a Python package
- **THEN** the file `static_data_template.xlsx` SHALL be available as a package resource alongside `template.xlsx`

#### Scenario: Template is published with the release

- **WHEN** a tagged release is built
- **THEN** the file `static_data_template.xlsx` SHALL be attached to the release as its own asset, next to the executables and not inside them

#### Scenario: Template has correct sheet structure

- **WHEN** the template file is opened
- **THEN** it SHALL contain exactly the six sheets listed above, with a section title in row 1, column headers in row 2, and data rows starting at row 3, except for the Millesimali sheet, where data rows start at row 5

#### Scenario: Template holds no exercise data

- **WHEN** the template file is opened
- **THEN** every cell the user fills in SHALL be empty: no name, millesimi, cost, reading, year or date

### Requirement: Static data reader

The system SHALL provide a `parse_static_data` function that reads a user-populated static data xlsx file and returns a structured `StaticData` object. Every error it raises SHALL be a `StaticDataError` whose message is plain Italian, names the file (and the sheet and cell, or the apartment or meter, where that applies) and says what to do next.

#### Scenario: Parse valid static data file

- **WHEN** a valid xlsx file matching the template structure is provided
- **THEN** the function SHALL return a `StaticData` object with all apartment data, costs, meter readings, previous readings, the exercise year and start date and the central meter readings populated

#### Scenario: File that is not an xlsx

- **WHEN** the provided file cannot be opened as an xlsx workbook
- **THEN** the function SHALL raise a `StaticDataError` that names the file, says it is not a static data file and points to `static_data_template.xlsx` on the release page

#### Scenario: Missing required sheet

- **WHEN** the provided xlsx file is missing one of the six required sheets, as a file made with an older template is missing `Esercizio`
- **THEN** the function SHALL raise a `StaticDataError` that names the file, says it is not a static data file, lists the missing sheets and points to `static_data_template.xlsx` on the release page

#### Scenario: Missing required cost row

- **WHEN** the Costi sheet is missing a required cost item
- **THEN** the function SHALL raise a `StaticDataError` that names the file, the sheet and the missing item and says that a cost that is really zero is entered as `0`

#### Scenario: Unknown cost label

- **WHEN** the Costi sheet has a label that is not one of the six cost items
- **THEN** the function SHALL raise a `StaticDataError` that names the file, the sheet and the cell, quotes the label as written and lists the accepted labels

#### Scenario: Value that is not a number

- **WHEN** a cell the reader converts to a number holds text that is not a number
- **THEN** the function SHALL raise a `StaticDataError` that names the file, the sheet and the cell and quotes the value found

#### Scenario: Blank number cell

- **WHEN** a cell the reader converts to a number is blank
- **THEN** the function SHALL raise a `StaticDataError` that names the file, the sheet and the cell and says to write the number, `0` if it really is zero

#### Scenario: Empty apartment rows

- **WHEN** the Inquilini, Millesimali or Letture_precedenti sheet has fewer rows than the apartments 1 to 10
- **THEN** the function SHALL raise a `StaticDataError` that names the file, the sheet and the missing apartment

### Requirement: Static data model

The system SHALL define dataclasses for static data following the project convention of pure models with no I/O.

- `ApartmentInfo`: apartment_number (int), proprietario (str), inquilino (str)
- `Millesimali`: apartment_number (int), subalterno (int), heat_energy_kwh (float), water_energy_kwh (float)
- `CostItem`: label (str), amount (float)
- `MeterReading`: name (str), unit (str), initial (float), final (float)
- `PreviousReading`: apartment_number (int), heat_kwh (int), water_m3 (float), cold_water_m3 (float)
- `ExerciseInfo`: year (int), start_date (date)
- `CentralMeterStart`: description (str, `Riscaldamento` or `Sanitario` as in the FC_report), initial_kwh (float), final_kwh (float or none), final_date (date or none)
- `StaticData`: apartments (list[ApartmentInfo]), millesimali (list[Millesimali]), costs (list[CostItem]), meters (list[MeterReading]), previous_readings (list[PreviousReading]), exercise (ExerciseInfo or none), central_meters (list[CentralMeterStart])

#### Scenario: StaticData is a pure dataclass

- **WHEN** `StaticData` is instantiated
- **THEN** it SHALL contain no I/O logic and no side effects

### Requirement: Writer merges static data into output

The `write_xlsx` function SHALL require a `StaticData` parameter and SHALL inject:

- **Inquilini sheet column A**: the owner name (proprietario field) when it is given, matching by apartment number; otherwise the device description of the FC_report.
- **Inquilini sheet column B**: tenant names (inquilino field) from StaticData, matching by apartment number.
- **Tabelle millesimali sheet**: `Subalterno` in column B, heating energy (kWh) in column C, ACS energy (kWh) in column E — for rows 4-13 (apartments 1-10).
- **Tabella sheet**: cost amounts into cells E3-E8 (matching by cost item label), and meter readings into cells F20-F22 (initial) and G20-G22 (final).
- **Ripartizione sheet**: previous heat readings into column B of heat rows (B78, B80, B82, ...), previous water readings into column B of water rows, previous cold water readings into column B of AFS rows.
- **Ripartizione sheet, exercise data**: the start readings of the central meters into `C27` (heating) and `C30` (ACS), the exercise start date into `D27`, `D30` and `B178`, and the exercise year, as its last two digits, in the labels `A28` and `A31` that end with `FINE GESTIONE`.
- **Ripartizione sheet, central meter final readings**: the manual final reading and date of a central meter when the static data gives them, otherwise the reading and date of the FC_report.

#### Scenario: Output with static data has tenant names

- **WHEN** `write_xlsx` is called with a valid `StaticData` containing tenant names
- **THEN** the Inquilini sheet column B SHALL contain the tenant names for each apartment

#### Scenario: Output with static data has owner names

- **WHEN** the static data gives an owner name for an apartment
- **THEN** the Inquilini sheet column A of that apartment SHALL hold the owner name, and an apartment without an owner name SHALL keep the device description

#### Scenario: Output with static data has millesimali

- **WHEN** `write_xlsx` is called with `StaticData` containing millesimali
- **THEN** the Tabelle millesimali sheet SHALL have energy values in columns C and E, and computed millesimi formulas in D and F SHALL produce correct results

#### Scenario: Output with static data has costs

- **WHEN** `write_xlsx` is called with `StaticData` containing costs
- **THEN** the Tabella sheet SHALL have cost amounts in the correct cells

#### Scenario: Output has the exercise data of the static file

- **WHEN** `write_xlsx` is called with static data for exercise 2026 starting on 31/12/2025 with central meter start readings 1000 and 2000
- **THEN** `C27` SHALL be 1000, `C30` SHALL be 2000, `D27`, `D30` and `B178` SHALL be 31/12/2025 and the labels `A28` and `A31` SHALL end with `FINE GESTIONE 26`

#### Scenario: Manual final reading overrides the report

- **WHEN** the static data gives a manual final reading and date for the heating central meter
- **THEN** `C28` and `D28` SHALL hold them instead of the values of the FC_report, and the ACS meter, which has none, SHALL keep the values of the FC_report in `C31` and `D31`

#### Scenario: Output without static data is unchanged

- **WHEN** the converter is run without static data
- **THEN** no output workbook SHALL be written

## ADDED Requirements

### Requirement: Apartment sheets list each apartment exactly once

On the Inquilini, Millesimali and Letture_precedenti sheets the reader SHALL accept one row for each apartment from 1 to 10, no more and no fewer.

#### Scenario: Duplicate apartment row

- **WHEN** an apartment appears in two rows of the same sheet
- **THEN** the function SHALL raise a `StaticDataError` that names the file, the sheet, the apartment and both cells

#### Scenario: Apartment number out of range

- **WHEN** the apartment cell of a row holds a number that is not from 1 to 10
- **THEN** the function SHALL raise a `StaticDataError` that names the file, the sheet and the cell and says the apartments go from 1 to 10

### Requirement: Only data rows and the TOTALE row are accepted

On every sheet of the static data file the reader SHALL accept data rows, entirely blank rows and the `TOTALE` row, and SHALL refuse any other row instead of skipping it or ignoring what follows it.

#### Scenario: Stray text row

- **WHEN** a row of Inquilini, Millesimali, Costi, Contatori or Letture_precedenti holds text that is neither an apartment number, an expected label nor `TOTALE`
- **THEN** the function SHALL raise a `StaticDataError` that names the file, the sheet and the cell, quotes the text and says what the sheet accepts

#### Scenario: Row below a blank row

- **WHEN** a data row follows a blank row
- **THEN** the function SHALL read it like any other data row

#### Scenario: Row without its label

- **WHEN** the apartment or label cell of a row is blank but other cells of the row are filled
- **THEN** the function SHALL raise a `StaticDataError` that names the file, the sheet and the cell

### Requirement: Utility meter rows are required

The Contatori sheet SHALL contain each of the three required meter rows exactly once.

#### Scenario: Missing meter row

- **WHEN** the Contatori sheet has no row for one of the three meters
- **THEN** the function SHALL raise a `StaticDataError` that names the file, the sheet and the missing meter

#### Scenario: Duplicate cost or meter row

- **WHEN** a cost item or a meter appears in two rows
- **THEN** the function SHALL raise a `StaticDataError` that names the file, the sheet, the item and both cells

### Requirement: Exercise data

The reader SHALL read from the Esercizio sheet the exercise year, the exercise start date, the start reading of the heating and of the ACS central heat meters and, for each, an optional manual final reading together with its date.

#### Scenario: Year or start date missing

- **WHEN** the year or the start date is blank, or the year is not a four-digit number, or the start date is not a date
- **THEN** the function SHALL raise a `StaticDataError` that names the file, the sheet and the cell and says what to write

#### Scenario: Central meter start reading missing

- **WHEN** the start reading of a central meter is blank or not a number
- **THEN** the function SHALL raise a `StaticDataError` that names the file, the sheet and the cell

#### Scenario: Manual final reading with its date

- **WHEN** a central meter has both a manual final reading and its date
- **THEN** the function SHALL return them with the meter

#### Scenario: Manual final reading without a date, or a date without a reading

- **WHEN** a central meter has only one of the manual final reading and its date
- **THEN** the function SHALL raise a `StaticDataError` that names the file, the sheet and the cell that is missing

### Requirement: Final readings are not lower than the start readings

After the FC_report and the static data are read and before the output is written, the system SHALL compare each final reading with the matching start reading, and SHALL refuse to write the output when a final reading is lower.

#### Scenario: Static data of another exercise

- **WHEN** the final reading of an apartment heat allocator, hot water meter or cold water volume in the FC_report is lower than the previous reading of that apartment in the static data
- **THEN** the system SHALL refuse the conversion with one Italian message that names every such meter, with the apartment and the kind of reading, and says to check that the static data is of the same exercise as the report

#### Scenario: Central meter below its start reading

- **WHEN** the final reading of a central heat meter, from the static data when given, otherwise from the FC_report, is lower than its start reading
- **THEN** the system SHALL refuse the conversion with one Italian message that names the meter

#### Scenario: Equal readings

- **WHEN** a final reading equals the start reading
- **THEN** the system SHALL NOT refuse the conversion
