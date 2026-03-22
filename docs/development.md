# Development Guide

## Module Responsibilities

| Module        | Responsibility                                  |
| ------------- | ----------------------------------------------- |
| `models.py`   | Pure dataclasses, no I/O                        |
| `parser.py`   | CSV parsing                                     |
| `sorter.py`   | Sorting + pivot grouping                        |
| `styles.py`   | Reusable openpyxl styles                        |
| `writer.py`   | XLSX generation (two sheets, formulas, filters) |
| `__main__.py` | CLI entry point, wires the pipeline             |

## Setup

```bash
pip install -e ".[dev]"
```

## Running Tests

```bash
pytest
pytest --cov=siemens_converter --cov-report=term-missing
```

## Building the Windows Executable

```bash
pyinstaller --onefile --paths src --name siemens-converter scripts/pyinstaller_entry.py
```

## CSV Format Reference

<!-- TODO: document the Siemens CSV format once known -->
