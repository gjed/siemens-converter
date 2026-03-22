# siemens-converter

[![CI](https://github.com/gjed/siemens-converter/actions/workflows/ci.yml/badge.svg)](https://github.com/gjed/siemens-converter/actions/workflows/ci.yml) [![codecov](https://codecov.io/gh/gjed/siemens-converter/branch/main/graph/badge.svg)](https://codecov.io/gh/gjed/siemens-converter) [![GitHub Release](https://img.shields.io/github/v/release/gjed/siemens-converter)](https://github.com/gjed/siemens-converter/releases/latest) [![Python](https://img.shields.io/badge/python-3.10%20%7C%203.11%20%7C%203.12%20%7C%203.13-blue)](https://github.com/gjed/siemens-converter) [![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE.md)

[Siemens](https://www.siemens.com/) manufactures **thermal energy meters (heat cost allocators)** used in condominiums to measure individual apartment heating consumption and fairly split shared heating expenses.

This tool converts **Siemens FC_report** files (HTML-as-XLS exported by the Siemens software) into professionally formatted XLSX workbooks with heating and water cost apportionment calculations.

> [Siemens](https://www.siemens.com/) produce **contatori di energia termica (ripartitori di calore)** utilizzati nei condomini per misurare il consumo di riscaldamento dei singoli appartamenti e suddividere equamente le spese di riscaldamento.
>
> Questo strumento converte i file **FC_report Siemens** (HTML-come-XLS esportati dal software Siemens) in cartelle di lavoro XLSX formattate professionalmente con i calcoli di ripartizione riscaldamento e acqua.

## Features / Funzionalita

- **Three-sheet output** -- Ripartizione (cost apportionment with formulas), Tabella (cost table), Tabelle millesimali (proportional shares)
  > **Output a tre fogli** -- Ripartizione (ripartizione costi con formule), Tabella (tabella costi), Tabelle millesimali (quote proporzionali)
- **Template-based** -- embedded XLSX template with all formulas; the tool only injects meter readings
  > **Basato su template** -- template XLSX incorporato con tutte le formule; lo strumento inserisce solo le letture dei contatori
- **Drag-and-drop** -- Windows users drop an FC_report `.xls` onto the `.exe`, done
  > **Trascina e rilascia** -- gli utenti Windows trascinano un FC_report `.xls` sull'`.exe`, fatto
- **Cross-platform** -- runs on Linux and macOS for development, Windows for end users
  > **Multipiattaforma** -- funziona su Linux e macOS per lo sviluppo, Windows per gli utenti finali

## Quick Start / Avvio rapido

```bash
pip install -e .
python -m siemens_converter path/to/FC_report.xls
```

## Windows (.exe)

Build a standalone executable for non-technical users:

> Crea un eseguibile standalone per utenti non tecnici:

```bash
pip install -e ".[dev]"
pyinstaller --onefile --paths src --name siemens-converter scripts/pyinstaller_entry.py
```

The `.exe` lands in `dist/`. Users drag an FC_report `.xls` onto it -- the XLSX appears next to the input file.

> L'`.exe` viene creato in `dist/`. Gli utenti trascinano un FC_report `.xls` sull'eseguibile -- il file XLSX appare accanto al file di input.

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
  parser.py      -> FC_report HTML parsing
  writer.py      -> XLSX generation (inject readings into template)
  template.xlsx  -> Embedded output template
tests/
  fixtures/      -> Anonymized test FC_report
```

See [docs/development.md](docs/development.md) for module responsibilities.

## Documentation

- [docs/uso.md](docs/uso.md) -- Guida utente (italiano)
- [docs/development.md](docs/development.md) -- Developer guide
