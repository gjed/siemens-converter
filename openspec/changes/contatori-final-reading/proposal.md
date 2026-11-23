# Proposal

## Why

The check of final readings against start readings covers the meters of the FC_report only. On the
`Contatori` sheet of the static data both readings are typed by hand, and a final reading typed lower
than the initial one (1380 instead of 13800) gives a negative consumption and a negative unit cost
for gas and electricity, with no refusal. Raised in the review of
[PR 33](https://github.com/gjed/siemens-converter/pull/33) and decided by the maintainer.

## What Changes

- A `Contatori` row whose final reading is lower than its initial reading is refused with one Italian
  message naming the file, the sheet and the cell. Equal readings stay valid.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

- `static-data-input`: the utility meter rows requirement gains the refusal of a final reading below
  the initial one.

## Impact

- `src/siemens_converter/static_reader.py`: `_read_contatori`.
- `tests/test_static_reader.py`, `tests/test_cli.py`: a unit test and a CLI negative test.
- `docs/uso.md`: the problem guide lists the refusal.
