# Development Guide

## Module Responsibilities

| Module             | Responsibility                                                                              |
| ------------------ | ------------------------------------------------------------------------------------------- |
| `models.py`        | Pure dataclasses for the report and the static data, no I/O                                 |
| `parser.py`        | FC_report parsing (HTML table saved as `.xls`)                                              |
| `static_reader.py` | Static data XLSX parsing                                                                    |
| `writer.py`        | XLSX generation: copies `template.xlsx`, injects values, adds `Inquilini` and `Dati Report` |
| `__main__.py`      | CLI entry point (arguments, drag-and-drop, Windows dialogs), wires the pipeline             |

## Setup

Python 3.15 is required.

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
