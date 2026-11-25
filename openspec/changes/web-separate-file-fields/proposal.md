# Proposal

## Why

The page takes the FC_report and the static data file in one drop zone and tells them apart by
extension. A user who picks the wrong files cannot see which file the page took for which role: a
real export for the Via Nicola Mazza 16 condominium came with an FC_report, an old Riparto and a
hand-made Riparto, two of them `.xlsx`. Chosen together, the page took whichever `.xlsx` came
first as the static data and refused it as a malformed static data file, without saying it had
ignored the other one. Separate fields make the role of each file explicit and visible
before the conversion starts.

## What Changes

- The single drop zone becomes two labelled fields, one for the FC_report and one for the static
  data file. Each takes one file, by click or by drag and drop, and shows the name of the file it
  holds; a new file replaces the previous one in that field.
- The field decides the role of the file: the page no longer routes the files by extension. A file
  in the wrong field is refused by the reader of that role, with the message the CLI prints for
  that file.
- The conversion starts on its own once both fields hold a file, as it does now with two files.
  `Convertire` stays disabled until both fields are filled, so the page no longer shows the CLI
  message for a missing file.
- Dropping more than one file on a field is refused with a message, and the field keeps its file.
- `docs/uso.md` describes the two fields.
- The CLI, the drag and drop on the `.exe` and the file dialogs are unchanged.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

- `web-page`: the page takes each file in its own field instead of routing two files by extension;
  the scenario for a missing static data file changes from the CLI message to a disabled
  `Convertire`.

## Impact

- `web/index.html`, `web/app.js`, `web/worker.js`, `web/style.css`
- `web/tests/test_page.py`
- `docs/uso.md`
- `openspec/specs/web-page/spec.md`
