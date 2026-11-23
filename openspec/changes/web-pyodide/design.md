# Design

## Context

`parse_fc_report`, `parse_static_data` and `write_xlsx` have no OS dependency; only `__main__.py`
(argument handling, file dialogs, message boxes, exit codes) is desktop-only. The sequence that
joins them, and the extension routing of the two files, were written inline in `__main__.py`.

## Decisions

- **One pipeline module, `siemens_converter.pipeline`.** `route_files(paths)` and
  `convert(fc_path, static_path, output_dir, progress)` raise `ConversionError`, whose message is
  the finished text the CLI shows (including the `Errore nel report FC:` and
  `Errore nel file dati statici:` prefixes, and the unprefixed cross-check error). `main()` turns
  it into `_fail`; the page shows it. The existence check, dialogs and message boxes stay in
  `main()`.
- **Pyodide 314.0.8 runs Python 3.14.2** (checked in `Makefile.envs` and `pyodide-lock.json` at
  that tag), so `requires-python` becomes `>=3.14`. The suite already runs on 3.14.
- **Served from the site, not a CDN.** The build downloads the `pyodide-core` tarball (about 7 MB,
  sha256 pinned), `micropip` and `packaging` wheels checked against the checksums in its lock
  file, and `openpyxl` and `et_xmlfile` wheels from PyPI with pinned hashes. `micropip` installs
  `openpyxl` from the site's own wheels, so every request is same-origin and the privacy claim can
  be tested. The converter source and `template.xlsx` are zipped from the checked-out commit.
- **Web Worker, module type.** The conversion and Pyodide run off the main thread so the page does
  not freeze; the worker starts loading at page load and the page shows a loading state until it
  reports ready.
- **Files accumulate until two are chosen.** The page converts as soon as two files are chosen,
  and a "Convertire" button handles the single-file case so that the CLI's "missing file" message
  is shown instead of waiting forever. "Ricominciare" clears the choice.
- **Online only.** No service worker: the browser's HTTP cache keeps Pyodide after the first visit.
  A service worker adds cache invalidation risk for little gain; offline users keep the `.exe`.
- **Deploy on `v*` tags** with `actions/upload-pages-artifact` and `actions/deploy-pages`, so the
  page matches a released version. Pull requests build the site and run the browser test, which
  is what keeps the page working between releases.
- **Browser test in Python, outside `tests/`.** It lives in `web/tests/` with Playwright, an
  optional `web` dependency, so the main suite and its coverage run need no browser.
- **Unexpected exceptions** show the Italian sentence of the executable's last-resort handler
  followed by the technical details.

## Risks

- A Pyodide or wheel URL going away breaks the build, not the deployed page; hashes make a
  tampered download fail loudly.
- The first visit downloads about 12 MB; later visits use the HTTP cache.
