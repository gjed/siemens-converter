# Proposal

## Why

The converter ships only as a PyInstaller binary: a Windows `.exe` and a Linux binary. An unsigned
`.exe` triggers SmartScreen warnings and antivirus false positives, the main obstacle for the
non-technical users the tool targets. It runs only on Windows, and every fix means a new release
and a new download ([issue #32](https://github.com/gjed/siemens-converter/issues/32)).

## What Changes

- A static web page, published on GitHub Pages, converts the FC_report and the static data file in
  the browser with the same Python code as the `.exe`, run by Pyodide. The files are read on the
  user's computer and never uploaded.
- The page is plain Italian: a drop zone or file picker takes the two files in any order, a
  loading state shows while Pyodide loads, the result is a download link named
  `Riparto_<serial>_<dd-mm-yyyy>.xlsx`, the warnings appear under "Attenzione", and errors are the
  exact messages the program shows. It links `static_data_template.xlsx` on the latest release and
  states that the files stay in the browser.
- The conversion sequence that lived only in `main()` becomes one function with no console, message
  box or Windows dependency, called by both `main()` and the page, together with the extension
  routing of the two files. The CLI and the page cannot drift apart.
- Pyodide, `openpyxl` and the converter source are served from the site itself, from pinned
  versions; the page makes no request to another origin.
- A workflow publishes the site on `v*` tags, so the page matches a released version. Every pull
  request builds the site and runs a headless browser test that compares the downloaded workbook
  with the CLI output.
- `requires-python` drops from `>=3.15` to `>=3.14`, the Python that Pyodide 314 runs. CI tests on
  3.14 and 3.15; the release binaries stay on 3.15.
- `docs/uso.md`, the README and `docs/development.md` describe the page.

## Capabilities

### New Capabilities

- `web-page`: the browser page, its inputs, messages, privacy, and its deploy on release tags.

### Modified Capabilities

None.

## Impact

- New: `src/siemens_converter/pipeline.py`, `web/` (page, worker, build script, browser test),
  `.github/workflows/pages.yml`.
- `src/siemens_converter/__main__.py`: `main()` and `_route_files` call the pipeline; arguments,
  dialogs, message boxes and exit codes are unchanged.
- `.github/workflows/ci.yml`: Python matrix and a site job. `release.yml` is unchanged.
- `pyproject.toml`, `openspec/config.yaml`, `docs/`, `README.md`.
- Out of scope: a JavaScript rewrite of the converter, server-side processing, any change to the
  release binaries, and offline use (no service worker; users who need it keep the `.exe`).
