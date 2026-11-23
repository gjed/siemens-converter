# Development Guide

## Module Responsibilities

| Module             | Responsibility                                                                              |
| ------------------ | ------------------------------------------------------------------------------------------- |
| `models.py`        | Pure dataclasses for the report and the static data, no I/O                                 |
| `parser.py`        | FC_report parsing (HTML table saved as `.xls`)                                              |
| `static_reader.py` | Static data XLSX parsing and the check of its start readings against the report             |
| `writer.py`        | XLSX generation: copies `template.xlsx`, injects values, adds `Inquilini` and `Dati Report` |
| `pipeline.py`      | The conversion sequence and the file routing, shared by the CLI and the web page            |
| `__main__.py`      | CLI entry point (arguments, drag-and-drop, Windows dialogs), calls `pipeline.py`            |

## Setup

Python 3.14 or 3.15 is required. CI tests on both; the release binaries are built on 3.15.

```bash
pip install -e ".[dev]"
```

## Running Tests

```bash
pytest
pytest --cov=siemens_converter --cov-report=term-missing
```

## Web Page

`web/` holds a static page that runs the same pipeline in the browser with [Pyodide](https://pyodide.org/): `index.html`, `app.js` and `style.css` for the page, `worker.js` for the Web Worker that runs Python. The page and the CLI both call `siemens_converter.pipeline`, so they cannot drift apart. Pyodide 314 runs Python 3.14, which is why `requires-python` is `>=3.14`.

`web/build.py` builds the site into `build/site`: the page, the converter source and `template.xlsx` zipped from the checked-out commit, the pinned `pyodide-core` tarball (sha256 pinned in the script), the `micropip` and `packaging` wheels checked against the Pyodide lock file, and the `openpyxl` and `et_xmlfile` wheels pinned with hashes in `web/wheels.txt`. Nothing is loaded from a CDN when the page runs. To update Pyodide, change `PYODIDE_VERSION` and `CORE_SHA256` in `web/build.py`.

```bash
python web/build.py                       # writes build/site
python -m http.server -d build/site 8000  # then open http://localhost:8000
```

The browser test converts the test fixtures in headless Chromium and compares the downloaded workbook, cell by cell, with the one the CLI writes. It also checks the refusals (a missing file, a malformed static data file, a file that is not a report), the warnings, and that the page makes no request outside its own origin.

```bash
pip install -e ".[web]"
playwright install chromium
pytest web/tests                          # SITE_DIR=<dir> tests another build
```

CI builds the site and runs this test on every pull request. `.github/workflows/pages.yml` publishes the site to GitHub Pages on every stable `v*` tag; pre-release tags (`v1.2.3-rc.1`) do not publish. There is no service worker: the page needs a connection, and the browser's HTTP cache keeps Pyodide after the first visit.

## Building the Windows Executable

The release workflow runs this command on Windows. `--add-data` bundles the output template; `--noconsole` makes the program show dialogs and message boxes instead of a terminal.

```bash
pyinstaller --onefile --paths src --name siemens-converter --add-data "src/siemens_converter/template.xlsx;siemens_converter" --noconsole scripts/pyinstaller_entry.py
```

On Linux the `--add-data` separator is `:` instead of `;`, and the workflow omits `--noconsole`.

`static_data_template.xlsx` is not bundled: the workflow attaches it to the release as a separate asset.

## Releasing

A release is cut by merging to `main`; nobody tags by hand.

1. On every push to `main`, the `Semantic release` workflow reads the Conventional Commits since the last `v*` tag and computes the next version: a breaking change gives a major version, `feat` a minor one, and `fix`, `perf` and `revert` a patch. Any other type (`docs`, `chore`, `ci`, `test`, `refactor`, `style`, `build`) releases nothing.
1. It pushes the tag `v<version>` and creates the GitHub release with its notes as a draft. It does not commit to `main`: there is no changelog file and no version bump.
1. The tag starts the `Release` workflow, which tests, builds the Windows `.exe` and the Linux binary and attaches them, with `static_data_template.xlsx`, to that release. It waits for the release if the build is faster than semantic-release, and publishes it as the latest release once the assets are attached.

If a `Release` job fails, the release stays a draft and the previous release stays the latest. Fix the cause and re-run the failed `Release` run: it attaches the assets again and publishes the draft. The next push to `main` does not retry a tag that already exists.

The tag is pushed with a token of the `gjed-release` GitHub App, because a tag pushed with the default `GITHUB_TOKEN` does not start other workflows. The workflow runs in the `release` environment, which holds the app credentials and accepts only `main`. Only the 3 newest stable releases keep their assets. There are no pre-releases.

The package version comes from the tag through `setuptools-scm`: `0.3.0` when built at the tag `v0.3.0`, and a development version such as `0.3.1.dev4+g1a2b3c4` after it. A checkout needs its full history and tags (`fetch-depth: 0` in CI) for the version to be right.

## FC_report Format Reference

An FC_report is a UTF-8 HTML document saved with an `.xls` extension. `parser.py` reads the rows of its table and addresses them by zero-based index.

| Row | Content                            |
| --- | ---------------------------------- |
| 0   | Metadata labels (ignored)          |
| 1   | Metadata values                    |
| 2   | Blank separator                    |
| 3   | Column headers of the device table |
| 4+  | Device data, one row per device    |

Metadata columns read from row 1:

| Column | Field                      |
| ------ | -------------------------- |
| 0      | File name                  |
| 1      | Report date (`dd/mm/yyyy`) |
| 2      | Report time                |
| 3      | Plant reference            |
| 4      | Firmware version           |
| 5      | Total wired devices        |
| 7      | Serial number              |

Device columns read from rows 4 and below:

| Column | Header                 | Use                                                                                                             |
| ------ | ---------------------- | --------------------------------------------------------------------------------------------------------------- |
| 2      | `device_serial_number` | Device serial number                                                                                            |
| 3      | `name_device`          | `contacalorie CT` marks a central heat meter                                                                    |
| 4      | `device_description`   | `Riscaldamento` or `Sanitario` for a central meter; for an apartment it carries the number, as in `App, 01 ...` |
| 5      | `device_detail`        | Device type: `Acqua calda`, `Contacalorie`, or empty for a central meter                                        |
| 9      | `readout_date`         | Reading date (`yyyy/mm/dd`)                                                                                     |
| 15     | `heat_energy`          | Heat energy: kWh for a central meter, MWh for an apartment heat meter                                           |
| 24     | `water_volume`         | Hot water volume in m³, for `Acqua calda` rows                                                                  |
| 26     | `aux1_volume`          | Cold water volume in m³, for `Contacalorie` rows                                                                |

Each device row is classified in this order; rows that match none are ignored:

1. `name_device` is `contacalorie CT` and `device_detail` is empty: central heat meter.
1. `device_detail` is `Acqua calda`: apartment hot water meter.
1. `device_detail` is `Contacalorie`: apartment heat meter.

Decimal values use a comma as separator, and an empty decimal cell is read as 0. The header row and the device rows also feed the `Dati Report` sheet of the output.
