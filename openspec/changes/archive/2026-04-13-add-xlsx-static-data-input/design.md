## Context

The converter currently follows a single-input pipeline: parse FC_report → inject readings into template → save output. The template.xlsx has formulas referencing an `Inquilini` sheet (columns A and B) and `Tabelle millesimali` sheet, but the Inquilini column B (tenant names) is always empty and the millesimali energy values are never populated — only the formulas computing millesimi from those values exist.

The "2026 Riparto ACS CT AFS GEST.xlsx" in `assets/` shows what the complete output looks like: tenant names in Ripartizione labels, millesimali values filled in, costs in Tabella, previous readings in the heat/water/AFS sections. All of this data is "static" in the sense that it doesn't come from the Siemens FC_report — it comes from the condominium administrator's records.

## Goals / Non-Goals

**Goals:**
- Add a second xlsx input carrying static condominium data (tenant names, millesimali, costs, utility readings, previous-period allocator readings)
- Provide a committed blank template for this input that anyone can populate
- Create a filled-in test file (gitignored) from the real 2026 data for manual testing
- Support two-file input in CLI, drag-and-drop, and Windows no-console dialog
- Maintain full backward compatibility when no static data file is provided

**Non-Goals:**
- Changing the output template structure (formulas, sheet layout stay the same)
- Auto-computing millesimali from readings (the input provides pre-calculated values)
- Supporting multiple condominium configurations (one building, 10 apartments)
- Validating financial correctness of cost data
- GUI beyond the existing Windows dialog pattern

## Decisions

### 1. Static data as structured xlsx (not CSV/JSON)

**Decision**: Use xlsx with named sheets as the input format for static data.

**Rationale**: The target users are building administrators who work in Excel. A structured xlsx with labeled sheets (Inquilini, Millesimali, Costi, Contatori, Letture_precedenti) is the most natural format. It also allows us to provide a template that users fill in directly in Excel.

**Alternative considered**: JSON or YAML config file — rejected because the target audience is non-technical and works exclusively in Excel.

### 2. New `static_reader.py` module (not extending `parser.py`)

**Decision**: Create a separate `static_reader.py` module for reading the static data xlsx.

**Rationale**: The existing `parser.py` handles HTML-as-XLS parsing of Siemens FC_reports — a completely different format. Mixing xlsx reading logic into it would violate the single-responsibility principle. The new module follows the same pattern: `parse_static_data(path) -> StaticData`.

### 3. Optional `StaticData` parameter in `write_xlsx`

**Decision**: Add `static_data: StaticData | None = None` parameter to `write_xlsx` rather than creating a separate writer function.

**Rationale**: The static data populates cells in the same sheets that `write_xlsx` already creates. A single function with an optional parameter keeps the pipeline simple and maintains backward compatibility. When `static_data is None`, the behavior is identical to current.

### 4. Windows file dialog with sequential selection

**Decision**: For no-console mode, use two sequential `win32` file open dialogs — first for FC_report, second for static data (cancellable).

**Rationale**: Windows native file dialogs (via `ctypes` / `win32gui`) don't natively support multi-file-with-different-types in a single dialog. Sequential dialogs are simpler and let the user skip the second file. No additional dependencies needed — `ctypes.windll` is already used.

**Alternative considered**: `tkinter.filedialog` — rejected because it would add a large dependency to the PyInstaller bundle and `tkinter` is not always available in PyInstaller builds.

### 5. Template bundled as package resource

**Decision**: Place `static_data_template.xlsx` alongside `template.xlsx` in `src/siemens_converter/` and bundle it in PyInstaller builds.

**Rationale**: Follows the existing pattern for `template.xlsx`. Users can extract the template from the installation or use the one from the repo.

### 6. Compiled 2026 data file — manual creation, gitignored

**Decision**: Create `assets/static_data_2026.xlsx` by hand (or script), add to `.gitignore` pattern.

**Rationale**: This file contains real tenant names and financial data. The `assets/` directory is already fully gitignored. This file exists purely for manual integration testing.

## Risks / Trade-offs

- **[Risk] Sheet/column name mismatch between template and reader**: The reader validates sheet names on open. If the template evolves, the reader must be updated in sync. → Mitigation: Sheet names are constants shared between reader and template creation.
- **[Risk] User provides wrong file as static data**: Could silently produce garbage output. → Mitigation: Validate sheet names and raise clear errors for structure mismatches.
- **[Risk] Two-file drag-and-drop order ambiguity**: If both files are `.xls`/`.xlsx`, we can't tell which is which by extension alone. → Mitigation: The FC_report is `.xls` (HTML) and static data is `.xlsx` (real xlsx) — different extensions. Additionally, validate by trying to parse each and checking structure.
- **[Trade-off] No GUI toolkit**: Using sequential `ctypes` dialogs is simpler but less polished than a proper GUI. Acceptable for the target audience.
- **[Trade-off] Fixed 10-apartment structure**: The template and model assume exactly 10 apartments. This is correct for the target condominium but not generalizable. Out of scope for now.

## Open Questions

- Should the static data template include example/placeholder values in cells, or be completely blank? Leaning toward placeholder headers only with a comment explaining each field.
- Should the compiled 2026 file be created programmatically from the existing `2026 Riparto` xlsx, or manually? Programmatic extraction would be more reliable.
