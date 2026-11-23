# siemens-converter

[![CI](https://github.com/gjed/siemens-converter/actions/workflows/ci.yml/badge.svg)](https://github.com/gjed/siemens-converter/actions/workflows/ci.yml) [![codecov](https://codecov.io/gh/gjed/siemens-converter/branch/main/graph/badge.svg)](https://codecov.io/gh/gjed/siemens-converter) [![GitHub Release](https://img.shields.io/github/v/release/gjed/siemens-converter)](https://github.com/gjed/siemens-converter/releases/latest) [![Python](https://img.shields.io/badge/python-3.14%20%7C%203.15-blue)](https://github.com/gjed/siemens-converter) [![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE.md)

[Siemens](https://www.siemens.com/) manufactures **thermal energy meters (heat cost allocators)** used in condominiums to measure individual apartment heating consumption and fairly split shared heating expenses.

This tool converts **Siemens FC_report** files (HTML-as-XLS exported by the Siemens software) into professionally formatted XLSX workbooks with heating and water cost apportionment calculations.

> [Siemens](https://www.siemens.com/) produce **contatori di energia termica (ripartitori di calore)** utilizzati nei condomini per misurare il consumo di riscaldamento dei singoli appartamenti e suddividere equamente le spese di riscaldamento.
>
> Questo strumento converte i file **FC_report Siemens** (HTML-come-XLS esportati dal software Siemens) in cartelle di lavoro XLSX formattate professionalmente con i calcoli di ripartizione riscaldamento e acqua.

## Features / Funzionalita

- **Three-sheet output** -- Ripartizione (cost apportionment with formulas), Tabella (cost table), Tabelle millesimali (proportional shares)
  > **Output a tre fogli** -- Ripartizione (ripartizione costi con formule), Tabella (tabella costi), Tabelle millesimali (quote proporzionali)
- **Template-based** -- embedded XLSX template with all formulas and no exercise data; the tool only injects meter readings and the static data
  > **Basato su template** -- template XLSX incorporato con tutte le formule e senza dati di un esercizio; lo strumento inserisce solo le letture dei contatori e i dati statici
- **Drag-and-drop** -- Windows users drop an FC_report `.xls` and a static data `.xlsx` onto the `.exe`, done
  > **Trascina e rilascia** -- gli utenti Windows trascinano un FC_report `.xls` e un file di dati statici `.xlsx` sull'`.exe`, fatto
- **Cross-platform** -- runs on Linux and macOS for development, Windows for end users
  > **Multipiattaforma** -- funziona su Linux e macOS per lo sviluppo, Windows per gli utenti finali
- **In the browser** -- the same converter runs on a web page, with no install and no upload
  > **Nel browser** -- lo stesso convertitore gira su una pagina web, senza installazione e senza caricare i file

## Quick Start / Avvio rapido

```bash
pip install -e .
python -m siemens_converter path/to/FC_report.xls path/to/static_data.xlsx
```

The second argument is required: a static data file with the exercise year and start date, the start readings of the central heat meters, tenant names, millesimali, costs, meter readings, and previous-period readings. Without it the converter stops with an error. Download `static_data_template.xlsx` from the [latest release](https://github.com/gjed/siemens-converter/releases/latest) for the expected format.

> Il secondo argomento è obbligatorio: un file di dati statici con anno e data di inizio dell'esercizio, letture iniziali dei contatori di calore della centrale, nomi inquilini, millesimali, costi, letture contatori e letture precedenti. Senza, il convertitore si ferma con un errore. Scaricare il template `static_data_template.xlsx` dall'[ultima release](https://github.com/gjed/siemens-converter/releases/latest) per il formato atteso.

## Web page / Pagina web

Open <https://gjed.github.io/siemens-converter/>, choose the FC_report `.xls` and the static data `.xlsx` (in any order) and download the `Riparto_*.xlsx`. Pyodide runs the same Python code as the `.exe` inside the browser: the files are processed on your computer and never sent anywhere. The page matches the latest release and needs an Internet connection; the `.exe` and the Linux binary are still available on the [latest release](https://github.com/gjed/siemens-converter/releases/latest) page.

> Aprire <https://gjed.github.io/siemens-converter/>, scegliere l'FC_report `.xls` e il file di dati statici `.xlsx` (in qualsiasi ordine) e scaricare il `Riparto_*.xlsx`. Pyodide esegue nel browser lo stesso codice Python dell'`.exe`: i file sono elaborati sul computer in uso e non vengono inviati da nessuna parte. La pagina corrisponde all'ultima release e richiede una connessione a Internet; l'`.exe` e il binario Linux restano disponibili nella pagina dell'[ultima release](https://github.com/gjed/siemens-converter/releases/latest).

## Windows (.exe)

Build a standalone executable for non-technical users:

> Crea un eseguibile standalone per utenti non tecnici:

```bash
pip install -e ".[dev]"
pyinstaller --onefile --paths src --name siemens-converter \
  --add-data "src/siemens_converter/template.xlsx:siemens_converter" \
  scripts/pyinstaller_entry.py
```

The `.exe` lands in `dist/`. Users drag an FC_report `.xls` and a static data `.xlsx` onto it -- the XLSX appears next to the FC_report.

> L'`.exe` viene creato in `dist/`. Gli utenti trascinano un FC_report `.xls` e un file di dati statici `.xlsx` sull'eseguibile -- il file XLSX appare accanto al FC_report.

## Development

```bash
pip install -e ".[dev]"
pytest
```

### Project Structure

```text
src/siemens_converter/
  __main__.py                -> CLI entry point (drag-and-drop, two-file input)
  pipeline.py                -> Conversion sequence and file routing, shared with the web page
  models.py                  -> Data classes (no I/O)
  parser.py                  -> FC_report HTML parsing
  static_reader.py           -> Static data XLSX parsing
  writer.py                  -> XLSX generation (inject readings into template)
  template.xlsx              -> Embedded output template
  static_data_template.xlsx  -> Template for static condominium data input
web/                         -> Web page, its build script and its browser test
tests/
  fixtures/                  -> Anonymized test data (FC_report + static data)
```

See [docs/development.md](docs/development.md) for module responsibilities.

## Documentation

- [docs/uso.md](docs/uso.md) -- Guida utente (italiano)
- [docs/development.md](docs/development.md) -- Developer guide
