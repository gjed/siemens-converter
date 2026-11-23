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

Python 3.14 or 3.15 is required.

```bash
pip install -e ".[dev]"
```

## Running Tests

```bash
pytest
pytest --cov=siemens_converter --cov-report=term-missing
```

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
