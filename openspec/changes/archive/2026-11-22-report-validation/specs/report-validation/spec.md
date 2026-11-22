# Spec Delta

## Purpose

Checks the converter runs on an FC report so that an incomplete report is refused and a report with
readings of different dates or devices in an abnormal state is converted with a warning the user
can read.

## ADDED Requirements

### Requirement: Incomplete reports are refused

The system SHALL refuse an FC report that does not contain both central meters (`name_device`
`contacalorie CT`, descriptions `Riscaldamento` and `Sanitario`), one heat meter (`device_detail`
`Contacalorie`) and one hot-water meter (`device_detail` `Acqua calda`) for each apartment 1 to 10.
The error SHALL list every missing device and no workbook SHALL be written.

#### Scenario: Complete report

- **WHEN** the report contains both central meters and a heat meter and a hot-water meter for each
  apartment 1 to 10
- **THEN** the conversion proceeds and the workbook is written

#### Scenario: Central meter missing

- **WHEN** the report has no `Sanitario` central meter
- **THEN** the conversion fails with a message that lists `contatore centrale «Sanitario»` as
  missing
- **AND** no workbook is written

#### Scenario: Apartment meters missing

- **WHEN** the report has no heat meter for apartment 3 and no hot-water meter for apartment 7
- **THEN** the conversion fails with one message that lists both
  `contacalorie dell'appartamento 3` and `contatore dell'acqua calda dell'appartamento 7`

### Requirement: Blank required readings are refused

The system SHALL refuse an FC report in which a reading the workbook needs is blank or is not a
number: `heat_energy` for central meters and heat meters, `aux1_volume` for heat meters,
`water_volume` for hot-water meters. The error SHALL name the device and the reading, and a blank
reading SHALL NOT be converted to 0.

#### Scenario: Blank central-meter reading

- **WHEN** the `heat_energy` cell of the `Riscaldamento` central meter is empty
- **THEN** the conversion fails with a message that names `contatore centrale «Riscaldamento»` and
  says its energy reading is empty

#### Scenario: Blank heat-meter readings

- **WHEN** the `heat_energy` or the `aux1_volume` cell of an apartment heat meter is empty
- **THEN** the conversion fails with a message that names that heat meter by its description and
  the empty reading

#### Scenario: Blank hot-water reading

- **WHEN** the `water_volume` cell of an apartment hot-water meter is empty
- **THEN** the conversion fails with a message that names that hot-water meter by its description

#### Scenario: Reading that is not a number

- **WHEN** a required reading contains text that is not a number
- **THEN** the conversion fails with a message that names the device and quotes the text found

#### Scenario: Missing devices and blank readings together

- **WHEN** the report has both a missing device and a blank required reading
- **THEN** the conversion fails with one message that lists both problems

### Requirement: Each date cell carries the readout date of its own meters

The system SHALL date each cell of the "Ripartizione" sheet from the meters it describes: `D28`
from the `Riscaldamento` central meter, `D31` from the `Sanitario` central meter, `C76` and `C178`
from the apartment heat meters, `C126` from the apartment hot-water meters. When apartment meters
of the same kind carry different readout dates, the cell SHALL carry the most recent one.

#### Scenario: Central meters read on different days

- **WHEN** the `Riscaldamento` central meter was read on 2026-01-10 and the `Sanitario` central
  meter on 2026-01-09
- **THEN** `D28` is 2026-01-10 and `D31` is 2026-01-09

#### Scenario: Apartment meters read after the central meters

- **WHEN** the central meters were read on 2026-01-03, the heat meters on 2026-01-15 and the
  hot-water meters on 2026-01-14
- **THEN** `C76` and `C178` are 2026-01-15 and `C126` is 2026-01-14

#### Scenario: Apartment meters of the same kind read on different days

- **WHEN** nine heat meters were read on 2026-01-15 and one on 2026-01-12
- **THEN** `C76` and `C178` are 2026-01-15

### Requirement: Readout date differences are reported as warnings

The system SHALL warn, without failing, when the `Riscaldamento` central meter and the apartment
heat meters, or the `Sanitario` central meter and the apartment hot-water meters, were read on
different days, stating both dates and the number of days between them. It SHALL also warn when
apartment meters of the same kind were read on different days, naming the meters read before the
most recent date.

#### Scenario: Central meter read before the apartment meters

- **WHEN** the `Riscaldamento` central meter was read on 2026-01-03 and the heat meters on
  2026-01-15
- **THEN** the conversion succeeds with a warning that names the central meter, both dates and
  `12 giorni prima`

#### Scenario: One day of difference

- **WHEN** a central meter was read one day after its apartment meters
- **THEN** the warning says `1 giorno dopo`

#### Scenario: One apartment meter left behind

- **WHEN** one heat meter was read on 2026-01-12 and the others on 2026-01-15
- **THEN** the conversion succeeds with a warning that names that heat meter by its description,
  its date and the date written in the workbook

#### Scenario: All readings taken on the same day

- **WHEN** every device in the report was read on the same day
- **THEN** no date warning is produced

### Requirement: Devices whose communication status is not Ok are reported as a warning

The system SHALL warn, without failing, when the `communication_status` of a central meter, a heat
meter or a hot-water meter is not `Ok`. The warning SHALL name each such device by its description
and report its status as written in the report.

#### Scenario: Central meter with an abnormal status

- **WHEN** the `communication_status` of the `Riscaldamento` central meter is
  `Portata negativa | Potenza negativa`
- **THEN** the conversion succeeds with a warning that names `contatore centrale «Riscaldamento»`
  and quotes that status

#### Scenario: Several devices with an abnormal status

- **WHEN** a heat meter and a hot-water meter of the same apartment both have a status other than
  `Ok`
- **THEN** one warning lists both, told apart as `contacalorie` and `contatore dell'acqua calda`

#### Scenario: Every status is Ok

- **WHEN** every device has `communication_status` `Ok`
- **THEN** no status warning is produced

### Requirement: Errors and warnings reach the user as readable messages

The system SHALL show an FC report error as a message and exit with code 1, never as a traceback:
on the console when there is one, in a message box otherwise. Warnings SHALL be printed on the
console before the final `Done` line and, when there is no console, appended to the final message
box. Messages SHALL be in Italian, name devices by their description in the report, and say what
to do next.

#### Scenario: Refused report on the console

- **WHEN** the converter is run from a console on a report that fails a check
- **THEN** the error message is printed, the exit code is 1 and no traceback is shown

#### Scenario: Refused report without a console

- **WHEN** the no-console build is run on a report that fails a check
- **THEN** the error message is shown in a message box and the exit code is 1

#### Scenario: Warnings on the console

- **WHEN** the converter is run from a console on a report that produces warnings
- **THEN** the warnings are printed before the `Done` line and the exit code is 0

#### Scenario: Warnings without a console

- **WHEN** the no-console build converts a report that produces warnings
- **THEN** the final message box contains the output file name followed by the warnings

#### Scenario: No warnings

- **WHEN** the report produces no warning
- **THEN** the console output and the final message box are the same as before this capability

### Requirement: A valid report is converted as before

The system SHALL write, for a report that passes every check, the same workbook it wrote before
this capability existed, except for the date cells `D28`, `D31`, `C76`, `C126` and `C178`.

#### Scenario: Report with every reading taken on the same day

- **WHEN** a complete report whose devices were all read on the same day is converted
- **THEN** every cell of every sheet has the same value as before, including the five date cells
