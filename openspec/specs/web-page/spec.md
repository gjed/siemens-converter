# web-page Specification

## Purpose

Define the static web page, published on GitHub Pages, that runs the converter in the browser with Pyodide as an alternative to the release binaries: which inputs it takes, which messages it shows, what stays on the user's computer, and when it is published.

## Requirements

### Requirement: Same pipeline as the CLI

The page SHALL convert the files by calling the same pipeline function as the command line, with the package source and `template.xlsx` taken from the same commit as the release, so that the workbook it offers is identical to the one the CLI writes for the same inputs. The file in the FC_report field SHALL be passed as the FC_report and the file in the static data field as the static data.

#### Scenario: Both files chosen in either order

- **WHEN** the user puts the FC_report in its field and the static data file in its field, filling either field first
- **THEN** the page SHALL offer a workbook with the same sheets and the same cell values as the CLI writes for those files

#### Scenario: Output name

- **WHEN** a conversion succeeds
- **THEN** the download SHALL be named `Riparto_<serial>_<dd-mm-yyyy>.xlsx`

### Requirement: Files never leave the browser

The page SHALL process the files only in the browser and send them nowhere. It SHALL serve Pyodide, `openpyxl` and the converter source from its own origin, load no third-party script, use no analytics and make no request to another origin. The page SHALL say in plain Italian that the files are processed only in this browser.

#### Scenario: Conversion makes no external request

- **WHEN** the page loads and converts a pair of files
- **THEN** every network request the browser makes SHALL go to the origin of the page

### Requirement: Same messages as the CLI

The page SHALL show the text of every error and warning exactly as the CLI shows it, with the same `Errore nel report FC:` and `Errore nel file dati statici:` prefixes, and SHALL offer no download when a conversion fails. Warnings of a successful conversion SHALL be listed under "Attenzione". Every other word on the page SHALL be plain Italian in the impersonal infinitive style of the CLI messages.

#### Scenario: Only the FC_report is chosen

- **WHEN** only the FC_report field holds a file
- **THEN** `Convertire` SHALL be disabled, the page SHALL start no conversion and SHALL offer no download

#### Scenario: Malformed static data

- **WHEN** the static data file has a blank cell the reader refuses
- **THEN** the page SHALL show the same message the CLI prints, naming file, sheet and cell, and offer no download

#### Scenario: File that is not an FC_report

- **WHEN** the file in the FC_report field is not an FC_report
- **THEN** the page SHALL show the same message the CLI prints and offer no download

#### Scenario: File that is not a static data file

- **WHEN** the file in the static data field is not a static data file, for example a Riparto written by an earlier conversion
- **THEN** the page SHALL show the same message the CLI prints for that file and offer no download

#### Scenario: Warnings

- **WHEN** a conversion succeeds with report warnings
- **THEN** the page SHALL list each warning under "Attenzione" next to the download link

### Requirement: Loading state and template link

The page SHALL show a loading state while Pyodide loads, run the conversion in a Web Worker so the page stays responsive, and link `static_data_template.xlsx` on the latest GitHub release.

#### Scenario: Page just opened

- **WHEN** Pyodide has not finished loading
- **THEN** the page SHALL say that the converter is loading and SHALL convert the chosen files once it is ready

### Requirement: Published on release tags

A workflow SHALL publish the site to GitHub Pages when a stable `v*` tag is pushed (a tag with a pre-release suffix such as `-rc.1` publishes nothing), with only the permissions `pages: write`, `id-token: write` and `contents: read`, so that the page always matches a released version. Pull requests SHALL build the site and run the browser test. The page is online only: there is no service worker.

#### Scenario: Release tag pushed

- **WHEN** a stable tag matching `v*` is pushed
- **THEN** the workflow SHALL build the site from that commit and deploy it to GitHub Pages

#### Scenario: Pull request

- **WHEN** a pull request is opened against `main`
- **THEN** CI SHALL build the site and run the headless browser test against it

### Requirement: A field for each file

The page SHALL take the FC_report and the static data file in two separate labelled fields, each holding at most one file, chosen with a file picker or dropped on the field, and showing its name. A new file SHALL replace the one the field held. The field SHALL decide the role of the file: the page SHALL NOT route the files by extension.

#### Scenario: A field's file replaced

- **WHEN** a conversion has been refused because of the static data file and the user chooses another file in the static data field
- **THEN** the page SHALL convert the FC_report with the new static data file and SHALL NOT use the earlier one

#### Scenario: Several files dropped on one field

- **WHEN** the user drops two files on the FC_report field
- **THEN** the page SHALL say that each field takes one file, keep the file the field held and start no conversion
