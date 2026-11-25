# Design

## Context

`web/app.js` keeps one list of chosen files and `web/worker.js` passes their paths to
`route_files`, which splits them by extension, then to `convert`. The CLI keeps `route_files`
because files dropped on the `.exe` arrive in no fixed order; the page has no such constraint.

## Goals / Non-Goals

**Goals:**

- One field per role, each holding at most one file, with the file name visible.
- The worker receives the two files by role and calls `convert` directly.

**Non-Goals:**

- Changing the CLI, the `.exe` drag and drop or the file dialogs.
- Checking the extension or the content of a file on the page: the readers already refuse a file
  that is not of their kind, with the CLI message.

## Decisions

- **Two `<label>` drop zones, each wrapping its own `<input type="file">`** (`#report` with
  `accept=".xls"`, `#static` with `accept=".xlsx"`), no `multiple`. The `accept` filter only
  pre-selects the file type in the picker; a dropped file of another type still goes to the
  reader, which explains what is wrong. Rejected: inferring the role from the extension when a
  file is dropped anywhere, which is the ambiguity this change removes.
- **State is `{ report, static }`**, one `File` or `null` each, replacing the `files` array and
  the replace-by-kind logic of `addFiles`. A new file in a field overwrites that field only.
- **The worker message carries `report` and `static`**, each `{ name, buffer }`, and the Python
  glue calls `convert(fc_path, static_path, out_dir)`. `route_files` stays in `pipeline.py` for
  the CLI.
- **A drop of several files on a field is refused in the page** with one message: no Python
  message covers it and no conversion is started.

## Risks / Trade-offs

- [A user drops both files on one field, as on the old single zone] → The refusal says that each
  field takes one file; the old behaviour stays in `docs/uso.md` only for the `.exe`.
- [The two readers' messages assume a wrong file, not a wrong field] → A Riparto in the static data
  field gets "non è un file dei dati statici: mancano i fogli …", which already names the file and
  points to the template; no new message is needed.
