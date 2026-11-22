# Design

## Context

The converter copies `template.xlsx`, writes the report and static data values into input cells
and saves. Every amount in the output is computed by formulas that live in the template, and
openpyxl never evaluates them: the saved file has no cached values until a spreadsheet application
opens it. The existing tests therefore see formula text only. See proposal.md for the three wrong
cells.

The cost chain on the first sheet is `C43 = C25`, `C44 = Tabella_2026!E12 = E3/H20`,
`C45 = C43*C44`, `C47 = C45 + C37`, `C51 = C47/C50`. `C51` prices every energy row, so `C25` reaches
every apartment total. `C23` feeds display cells only (`C24`, `C36`, `C38`, `C41`, `C49`).

## Goals / Non-Goals

**Goals:**

- Four cells changed in the template, nothing else in the file.
- A test that fails on the old template and passes on the new one, by evaluating formulas.
- CI that cannot go green without running that test.

**Non-Goals:**

- The other pasted 2026 values in the template (central meter start readings, previous readings,
  dates): [issue #24](https://github.com/gjed/siemens-converter/issues/24).
- Validating the static data or the report before writing.
- Shipping cached values in the output workbook.

## Decisions

### Recalculate with LibreOffice headless, not a Python formula library

The test converts the generated workbook with `soffice --headless --convert-to xlsx` and reads
the cached values back with `openpyxl.load_workbook(data_only=True)`.

- LibreOffice Calc is one of the two applications users open the workbook with, so the test
  checks what a user sees.
- The check cell `L250` compares two floating point sums with `=`. Spreadsheet applications treat
  values that differ in the last bits as equal; a Python library would have to copy that rule.
- The workbook uses sheet-scoped defined names (`Tabella_2026!costoafs`,
  `costoenegiaelettrica`), which the real engine resolves without help.
- A Python library (`formulas`, `pycel`) would be a new development dependency that has to
  install on Python 3.15, for an evaluator that is not the one users run.

The cost is a system package in CI and a test that needs LibreOffice locally.

### One LibreOffice start for all variants

A module-scoped fixture writes three workbooks (utility meters 0 → N, 1000 → 1000 + N, and
tripled heating millesimali) and converts them in a single `soffice` call. LibreOffice starts
once, so the recalculation tests add a few seconds to the suite.

The call passes `-env:UserInstallation=<tmp dir>` so it uses its own profile and does not attach
to, or get blocked by, a LibreOffice the developer has open.

### Static data variants come from the committed fixture

The base is `tests/fixtures/static_data_TEST.xlsx` read with the real static data reader. The
variants are derived in the test with `dataclasses.replace` and the `MeterReading` dataclass. No
fixture file is added or changed.

With the fixture values some per-apartment amounts are negative, because the template still holds
the 2026 start readings of the central meters and they are higher than the fixture's final
readings (issue #24). The properties under test (no error cell, distributed total equal to the
costs, independence from the meters' initial reading, shares equal to the millesimali shares) are
algebraic and hold regardless.

### Skip locally, fail in CI

The fixture skips when `soffice` is not on `PATH`, so a contributor without LibreOffice can still
run the suite. When the `CI` environment variable is set, which GitHub Actions does for every
job, a missing `soffice` is a test failure instead. No workflow has to opt in, and a workflow
that runs `pytest` without installing LibreOffice goes red.

### Both workflows install LibreOffice

`ci.yml` and the `test` job of `release.yml` install `libreoffice-calc` with apt before running
pytest. The release job is the last gate before binaries that embed `template.xlsx` are
published, so it runs the same tests as CI; leaving it out would also make it fail, because of
the `CI` rule above. The `build` jobs run no tests and are unchanged.

### Template edited with openpyxl

The template is already an openpyxl-written file, so loading it, setting four cells and saving
changes `xl/worksheets/sheet1.xml` (the four cells) and the modification date in
`docProps/core.xml`. Every other part of the package stays byte-identical. The four cells keep
their style index, so number formats are unchanged.

## Risks / Trade-offs

- [The apt install adds about a minute to each CI run] → Accepted: it is the only way the
  formulas get evaluated at all.
- [A LibreOffice release changes how it recalculates files without cached values] → The test
  would fail on `None` values rather than pass silently, and the assertions compare against
  values computed from the inputs.
- \[`L250` depends on the application's tolerant `=` on floating point sums\] → Out of scope; the
  test also asserts `I247` against `Tabella_2026!E9` with a tolerance, independently of `L250`.
