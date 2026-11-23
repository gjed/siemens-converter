# Tasks

## 1. Pipeline

- [x] 1.1 Extract `route_files` and `convert` into `siemens_converter.pipeline`, raising `ConversionError`
- [x] 1.2 Make `main()` and `_route_files` call it, keeping arguments, dialogs, message boxes and exit codes
- [x] 1.3 Test the pipeline in `tests/test_pipeline.py`; `tests/test_cli.py` passes unchanged

## 2. Python 3.14

- [x] 2.1 Lower `requires-python` to `>=3.14` and update `openspec/config.yaml` and `docs/development.md`
- [x] 2.2 Test on Python 3.14 and 3.15 in `ci.yml`; the release build stays on 3.15

## 3. Page

- [x] 3.1 Write the page, its worker and styles in `web/`
- [x] 3.2 Write `web/build.py`: pinned Pyodide core, micropip and packaging, pinned `openpyxl` and `et_xmlfile` wheels, converter zip
- [x] 3.3 Test the built page with a headless browser: workbook equal to the CLI output, the three refusals, no external request

## 4. Workflows

- [x] 4.1 Build the site and run the browser test on pull requests in `ci.yml`
- [x] 4.2 Add `pages.yml`, publishing to GitHub Pages on `v*` tags

## 5. Documentation

- [x] 5.1 Describe the page in `docs/uso.md` and `README.md`
- [x] 5.2 Add the site build and test commands to `docs/development.md`
- [x] 5.3 Run the markdown linter and `openspec validate web-pyodide --strict`
