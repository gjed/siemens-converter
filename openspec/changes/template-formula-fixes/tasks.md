# Tasks

## 1. Recalculation test

- [x] 1.1 Add `tests/test_recalc.py`: build three workbooks through `parse_fc_report`, the static
  data fixture and `write_xlsx`, recalculate them with LibreOffice headless in one call with its
  own profile, read cached values with openpyxl. Verify: against the unfixed template the offset
  and millesimali assertions fail.
- [x] 1.2 Skip when `soffice` is missing, fail instead when `CI` is set. Verify: with `soffice`
  hidden from `PATH` the tests are skipped, and with `CI=true` as well they fail.

## 2. Template

- [x] 2.1 Change `C25`, `C23`, `C63` and `C64` on the first sheet of `template.xlsx`. Verify: a
  cell-by-cell comparison with the previous template shows four changed cells, unchanged sheet
  names, merged ranges, defined names and number formats.
- [x] 2.2 Assert the four new formulas in `tests/test_template.py`. Verify: the tests pass on the
  new template.
- [x] 2.3 Run the recalculation tests on the new template. Verify: all pass.

## 3. CI

- [x] 3.1 Install LibreOffice Calc before the test step in `.github/workflows/ci.yml` and in the
  `test` job of `.github/workflows/release.yml`. Verify: both files parse as YAML with the install
  step before the pytest step, and the CI run of the pull request reports no skipped test.

## Workflow follow-up

- Archive the change in the same pull request.
