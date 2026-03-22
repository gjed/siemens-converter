# siemens-converter

[![CI](https://github.com/gjed/siemens-converter/actions/workflows/ci.yml/badge.svg)](https://github.com/gjed/siemens-converter/actions/workflows/ci.yml) [![codecov](https://codecov.io/gh/gjed/siemens-converter/branch/main/graph/badge.svg)](https://codecov.io/gh/gjed/siemens-converter) [![GitHub Release](https://img.shields.io/github/v/release/gjed/siemens-converter)](https://github.com/gjed/siemens-converter/releases/latest) [![Python](https://img.shields.io/badge/python-3.10%20%7C%203.11%20%7C%203.12%20%7C%203.13-blue)](https://github.com/gjed/siemens-converter) [![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE.md)

[Siemens](https://www.siemens.com/) manufactures **thermal energy meters (heat cost allocators)** used in condominiums to measure individual apartment heating consumption and fairly split shared heating expenses.

This tool converts **Siemens report** CSV files into professionally formatted XLSX workbooks with a pivot summary sheet.

> [Siemens](https://www.siemens.com/) produce **contatori di energia termica (ripartitori di calore)** utilizzati nei condomini per misurare il consumo di riscaldamento dei singoli appartamenti e suddividere equamente le spese di riscaldamento.
>
> Questo strumento converte i file CSV dei **report Siemens** in cartelle di lavoro XLSX formattate professionalmente con un foglio di riepilogo pivot.

## Features / Funzionalita

- **Two-sheet output** -- PIVOT summary with apartment totals + full raw data sheet
  > **Output a due fogli** -- riepilogo PIVOT con i totali per appartamento + foglio dati grezzi completo
- **Professional formatting** -- dark headers, alternating rows, auto-fit columns, freeze panes
  > **Formattazione professionale** -- intestazioni scure, righe alternate, colonne adattate automaticamente, riquadri bloccati
- **Auto-filter** -- dropdown filters on every column of the data table
  > **Filtro automatico** -- filtri a tendina su ogni colonna della tabella dati
- **Drag-and-drop** -- Windows users drop a CSV onto the `.exe`, done
  > **Trascina e rilascia** -- gli utenti Windows trascinano un CSV sull'`.exe`, fatto
- **Cross-platform** -- runs on Linux and macOS for development, Windows for end users
  > **Multipiattaforma** -- funziona su Linux e macOS per lo sviluppo, Windows per gli utenti finali

## Quick Start / Avvio rapido

```bash
pip install -e .
python -m siemens_converter path/to/report.csv
```

## Windows (.exe)

Build a standalone executable for non-technical users:

> Crea un eseguibile standalone per utenti non tecnici:

```bash
pip install -e ".[dev]"
pyinstaller --onefile --paths src --name siemens-converter scripts/pyinstaller_entry.py
```

The `.exe` lands in `dist/`. Users drag a CSV onto it -- the XLSX appears next to the CSV.

> L'`.exe` viene creato in `dist/`. Gli utenti trascinano un CSV sull'eseguibile -- il file XLSX appare accanto al CSV.

## Development

```bash
pip install -e ".[dev]"
pytest
```

### Project Structure

```text
src/siemens_converter/
  __main__.py    -> CLI entry point (drag-and-drop)
  models.py      -> Data classes (no I/O)
  parser.py      -> CSV parsing
  sorter.py      -> Sorting + pivot grouping
  styles.py      -> XLSX formatting (fonts, fills, borders)
  writer.py      -> XLSX generation (two sheets, formulas, filters)
tests/
  fixtures/      -> Anonymized test CSV
```

See [docs/development.md](docs/development.md) for module responsibilities and CSV format reference.

## Documentation

- [docs/uso.md](docs/uso.md) -- Guida utente (italiano)
- [docs/development.md](docs/development.md) -- Developer guide
