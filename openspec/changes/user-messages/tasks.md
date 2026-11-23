# Tasks

## 1. FC report

- [x] 1.1 Refuse files that are not an FC report in `parse_fc_report` with a `ReportError` naming the file
- [x] 1.2 Give `extract_apartment_number` Italian wording and refuse a device without an apartment number in the parser
- [x] 1.3 Test each refusal in `tests/test_parser.py`

## 2. Static data

- [x] 2.1 Add `StaticDataError`, refuse files that do not open or lack sheets, point to the template
- [x] 2.2 Rewrite the unknown-cost-label error and refuse non-numeric values with file, sheet and cell
- [x] 2.3 Test each message in `tests/test_static_reader.py`

## 3. CLI

- [x] 3.1 Route files by extension in one function and report a missing or doubled file
- [x] 3.2 Translate the usage text, progress lines, final line and file dialog failure
- [x] 3.3 Show only `ReportError` and `StaticDataError` as messages
- [x] 3.4 Lead the last-resort handler with an Italian sentence
- [x] 3.5 Test each message on the console and without a console in `tests/test_cli.py`

## 4. Documentation

- [x] 4.1 Rewrite "Problemi e soluzioni" in `docs/uso.md` and add the report refusal and warnings
- [x] 4.2 Update `README.md` and `docs/development.md` where they quote a message
- [x] 4.3 Run the markdown linter and `openspec validate --all --strict`
