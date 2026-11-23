# Design

## Context

Messages reach the user through `__main__.py`, which catches every exception from the parser and
the static reader and prefixes it, and through the last-resort handler of
`scripts/pyinstaller_entry.py`. `ReportError` already carries an Italian message for incomplete
reports. The static reader raises plain `ValueError` with English text, and lets openpyxl and
`float()` errors through. The readers of issue #25 will rewrite the static data checks next, so the
message strings must be easy to extend.

## Decisions

- **One exception type per input, carrying the finished Italian message.** `ReportError` (parser)
  and the new `StaticDataError` (static reader) both subclass `ValueError`. `__main__.py` shows
  only these two as messages, prefixed `Errore nel report FC:` or `Errore nel file dati statici:`;
  any other exception is a bug and reaches the last-resort handler, which is the one place that
  still shows technical text. Catching `Exception` in `main()` would put English exception text
  in front of the user.
- **Messages are built next to the check that finds the problem**, as short sentences ending with
  an instruction in the impersonal infinitive ("Scegliere…", "Controllare…") already used by
  `ReportError`. Shared advice (what a static data file is, where to download the template) is a
  module constant, so #25 reuses it.
- **Cell positions use the spreadsheet's own notation** (`foglio «Costi», cella A4`), computed from
  the row counter and `get_column_letter`, never from a column index.
- **Routing is one function over the paths given**, used by drag-and-drop, CLI and the dialogs, so
  the same three mistakes get the same three messages everywhere.
- **The wording of the existing refusal and warnings is unchanged.** `models.extract_apartment_number`
  gets Italian wording; the parser turns its failure into a `ReportError` that names the device.
- **A blank cell still reads as 0, duplicate and missing rows are still accepted.** Only an
  unreadable value changes from an English conversion error to an Italian one; that behaviour is
  for #25.
- **The last-resort contact is "chi ha fornito il programma"**, the wording the guide already uses.

## Risks

- A bug that raises an unexpected exception on the CLI now shows a traceback instead of
  `Errore nel report FC: <English text>`. Both are failures; the traceback is more useful to the
  developers who use the CLI, and the executable wraps it in the Italian sentence.
