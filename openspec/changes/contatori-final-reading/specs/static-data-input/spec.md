## MODIFIED Requirements

### Requirement: Utility meter rows are required

The Contatori sheet SHALL contain each of the three required meter rows exactly once, and the final reading of a row SHALL NOT be lower than its initial reading.

#### Scenario: Missing meter row

- **WHEN** the Contatori sheet has no row for one of the three meters
- **THEN** the function SHALL raise a `StaticDataError` that names the file, the sheet and the missing meter

#### Scenario: Duplicate cost or meter row

- **WHEN** a cost item or a meter appears in two rows
- **THEN** the function SHALL raise a `StaticDataError` that names the file, the sheet, the item and both cells

#### Scenario: Final reading below the initial reading

- **WHEN** the final reading of a meter of the Contatori sheet is lower than its initial reading
- **THEN** the function SHALL raise a `StaticDataError` that names the file, the sheet and the cell of the final reading, and gives both readings

#### Scenario: Equal readings

- **WHEN** the final reading of a meter of the Contatori sheet equals its initial reading
- **THEN** the function SHALL accept the row
