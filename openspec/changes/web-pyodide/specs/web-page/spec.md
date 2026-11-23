## ADDED Requirements

### Requirement: Same pipeline as the CLI

The page SHALL convert the files by calling the same pipeline function as the command line, with the package source and `template.xlsx` taken from the same commit as the release, so that the workbook it offers is identical to the one the CLI writes for the same inputs. The files SHALL be routed by extension like the CLI: the `.xlsx` is the static data and the other file is the FC_report.

#### Scenario: Both files chosen in either order

- **WHEN** the user chooses the FC_report and the static data file, in any order, on the page
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

- **WHEN** the user chooses only the FC_report and asks to convert
- **THEN** the page SHALL show the CLI message that the static data file is missing and offer no download

#### Scenario: Malformed static data

- **WHEN** the static data file has a blank cell the reader refuses
- **THEN** the page SHALL show the same message the CLI prints, naming file, sheet and cell, and offer no download

#### Scenario: File that is not an FC_report

- **WHEN** the file given as FC_report is not an FC_report
- **THEN** the page SHALL show the same message the CLI prints and offer no download

#### Scenario: Warnings

- **WHEN** a conversion succeeds with report warnings
- **THEN** the page SHALL list each warning under "Attenzione" next to the download link

### Requirement: Loading state and template link

The page SHALL show a loading state while Pyodide loads, run the conversion in a Web Worker so the page stays responsive, and link `static_data_template.xlsx` on the latest GitHub release.

#### Scenario: Page just opened

- **WHEN** Pyodide has not finished loading
- **THEN** the page SHALL say that the converter is loading and SHALL convert the chosen files once it is ready

### Requirement: Published on release tags

A workflow SHALL publish the site to GitHub Pages when a `v*` tag is pushed, with only the permissions `pages: write`, `id-token: write` and `contents: read`, so that the page always matches a released version. Pull requests SHALL build the site and run the browser test. The page is online only: there is no service worker.

#### Scenario: Release tag pushed

- **WHEN** a tag matching `v*` is pushed
- **THEN** the workflow SHALL build the site from that commit and deploy it to GitHub Pages

#### Scenario: Pull request

- **WHEN** a pull request is opened against `main`
- **THEN** CI SHALL build the site and run the headless browser test against it
