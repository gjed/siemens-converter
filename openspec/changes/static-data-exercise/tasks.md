# Tasks

## 1. Template

- [ ] 1.1 Clear every exercise literal of `template.xlsx`, keep every formula and rename the sheets to `Ripartizione` and `Tabella`, rewriting the cross-sheet formulas and defined names
- [ ] 1.2 Add `tests/test_template.py` tests for the literal whitelist and the neutral sheet names, and update the assertions on the old names

## 2. Reader and models

- [ ] 2.1 Add the `Esercizio` sheet to `static_data_template.xlsx` and the exercise and central meter models
- [ ] 2.2 Read the `Esercizio` sheet in `parse_static_data`
- [ ] 2.3 Refuse blank numbers, duplicate, missing and out-of-range apartments, stray rows and missing cost and meter rows
- [ ] 2.4 Add the report-versus-static cross-check
- [ ] 2.5 Test each refusal in `tests/test_static_reader.py`

## 3. Writer and CLI

- [ ] 3.1 Write the exercise data, the central meter overrides and the owner names in `write_xlsx`, require the static data and run the cross-check
- [ ] 3.2 Require the static data file in `_route_files`, the second dialog and the usage text
- [ ] 3.3 Regenerate `tests/fixtures/static_data_TEST.xlsx` with invented values and update `test_writer.py`, `test_cli.py` and `test_recalc.py`

## 4. Documentation

- [ ] 4.1 Update `docs/uso.md`, `README.md` and `docs/development.md`
- [ ] 4.2 Update the Purpose lines of the main specs `static-data-input` and `two-file-workflow`
- [ ] 4.3 Run the markdown linter and `openspec validate --all --strict`
