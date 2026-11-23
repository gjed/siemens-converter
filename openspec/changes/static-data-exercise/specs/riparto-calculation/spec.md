## MODIFIED Requirements

### Requirement: Distributed total equals the costs

The sum of the amounts distributed to the apartments SHALL equal the sum of the costs entered in
the static data, and the workbook's own check cell SHALL confirm it.

#### Scenario: Complete static data

- **WHEN** the workbook is generated from a report and static data that give every apartment its
  millesimali, every cost its amount and every utility meter its readings, and is recalculated
- **THEN** the grand total `I247` of the first sheet SHALL equal `Tabella!E9`
- **AND** the check cell `L250` of the first sheet SHALL read `esatto`

#### Scenario: Utility meters do not start at zero

- **WHEN** the three utility meters read 1000 at the start of the period
- **THEN** the grand total `I247` SHALL still equal `Tabella!E9`
- **AND** the check cell `L250` SHALL read `esatto`

### Requirement: Recalculated workbook has no error cells

A workbook generated from a report and complete static data SHALL contain no formula error after
a spreadsheet application recalculates it.

#### Scenario: Recalculation by a spreadsheet application

- **WHEN** the generated workbook is recalculated
- **THEN** no cell of the sheets `Ripartizione`, `Tabella` and `Tabelle millesimali` SHALL hold an
  error value such as `#DIV/0!`, `#REF!` or `#NAME?`

## ADDED Requirements

### Requirement: Sheet names do not depend on the exercise

The sheets of the output workbook SHALL have the fixed names `Ripartizione`, `Tabella` and
`Tabelle millesimali`, whatever the exercise, so that no formula or defined name has to be rewritten
for a run.

#### Scenario: Output sheet names

- **WHEN** a workbook is generated for any exercise
- **THEN** its first three sheets SHALL be named `Ripartizione`, `Tabella` and `Tabelle millesimali`
- **AND** every formula that refers to another sheet, and the defined names `costoafs` and
  `costoenegiaelettrica` scoped to a sheet, SHALL refer to those names

### Requirement: The template holds no exercise data

The workbook template SHALL contain no value that belongs to one exercise: no start reading, no
date, no year, no tenancy days. Its first sheet SHALL hold a number or a date as a literal only in
`C40` (9.45), `C48` (93), `C123` (13) and `C124` (50), the physical constants of the plant. Every
formula SHALL be kept.

#### Scenario: No literal outside the constants

- **WHEN** the first sheet of the template is inspected
- **THEN** no cell outside `C40`, `C48`, `C123` and `C124` SHALL hold a number or a date as a literal

#### Scenario: Run without exercise data

- **WHEN** a workbook is generated and the static data gives the exercise data
- **THEN** every start reading, date, year label and tenancy value of the output SHALL come from the static data and none from the template
