# Spec Delta

## Purpose

Defines what the formulas of the generated workbook compute from the Siemens report and the static
data, so that the cost apportionment (riparto) stays correct when the input values change.

## ADDED Requirements

### Requirement: Electricity cost is independent of the meter's initial reading

The workbook SHALL compute the electricity consumption of the heating plant as the difference
between the final and the initial reading of the electricity meter, so that the electricity cost
it distributes equals the electricity cost entered in the static data.

#### Scenario: Initial reading is not zero

- **WHEN** the electricity meter reads 1000 at the start and 1000 + N at the end of the period
- **THEN** the electricity consumption in the first sheet (`C25`) SHALL be N
- **AND** every per-apartment total SHALL equal the one obtained with readings 0 and N

#### Scenario: Initial reading is zero

- **WHEN** the electricity meter reads 0 at the start and N at the end of the period
- **THEN** the electricity consumption in the first sheet (`C25`) SHALL be N

### Requirement: Gas readings are shown as entered

The first sheet SHALL show the gas meter's final reading as entered in the static data, and the
gas consumption SHALL be the difference between the final and the initial reading.

#### Scenario: Gas initial reading is not zero

- **WHEN** the gas meter reads 1000 at the start and 2718 at the end of the period
- **THEN** the gas final reading in the first sheet (`C23`) SHALL be 2718
- **AND** the gas consumption in the first sheet (`C36`) SHALL be 1718

### Requirement: Management cost split follows the millesimali

The shares that split the management costs between heating (`kcli`) and domestic hot water
(`kacs`) SHALL be computed from the heating and hot water energy totals of the millesimali table
in the same workbook. They SHALL NOT be constants.

#### Scenario: Shares match the millesimali table

- **WHEN** the workbook is generated from static data and recalculated
- **THEN** `C63` of the first sheet SHALL equal `'Tabelle millesimali'!C18`
- **AND** `C64` of the first sheet SHALL equal `'Tabelle millesimali'!E18`

#### Scenario: Millesimali change

- **WHEN** the heating energy values of the millesimali are tripled and the hot water values are
  unchanged
- **THEN** `C63` SHALL be the new heating share, heating total divided by the sum of the heating
  and hot water totals
- **AND** `C64` SHALL be the new hot water share

### Requirement: Distributed total equals the costs

The sum of the amounts distributed to the apartments SHALL equal the sum of the costs entered in
the static data, and the workbook's own check cell SHALL confirm it.

#### Scenario: Complete static data

- **WHEN** the workbook is generated from a report and static data that give every apartment its
  millesimali, every cost its amount and every utility meter its readings, and is recalculated
- **THEN** the grand total `I247` of the first sheet SHALL equal `Tabella_2026!E9`
- **AND** the check cell `L250` of the first sheet SHALL read `esatto`

#### Scenario: Utility meters do not start at zero

- **WHEN** the three utility meters read 1000 at the start of the period
- **THEN** the grand total `I247` SHALL still equal `Tabella_2026!E9`
- **AND** the check cell `L250` SHALL read `esatto`

### Requirement: Recalculated workbook has no error cells

A workbook generated from a report and complete static data SHALL contain no formula error after
a spreadsheet application recalculates it.

#### Scenario: Recalculation by a spreadsheet application

- **WHEN** the generated workbook is recalculated
- **THEN** no cell of the sheets `Ripartizione_Termo 02.03.2026`, `Tabella_2026` and
  `Tabelle millesimali` SHALL hold an error value such as `#DIV/0!`, `#REF!` or `#NAME?`
