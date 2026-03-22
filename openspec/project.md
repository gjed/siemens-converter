# Project Context

## Purpose

Convert Siemens heat cost allocator report CSV files into professionally formatted XLSX workbooks with a PIVOT summary sheet. Parallel project to [sinapsi-converter](https://github.com/gjed/sinapsi-converter) — same architecture, different device manufacturer.

## Tech Stack

- Python >=3.10
- setuptools (pyproject.toml, src-layout)
- openpyxl (XLSX generation)
- pytest + pytest-cov (testing)
- PyInstaller (Windows .exe distribution)
- GitHub Actions (CI matrix 3.10–3.13, tag-triggered releases)

## Project Conventions

### Code Style

- Modules are single-responsibility: models (dataclasses, no I/O), parser (CSV reading), sorter (ordering + pivot), styles (openpyxl formatting), writer (XLSX output)
- `from __future__ import annotations` in every module
- Type hints on public functions

### Architecture Patterns

- Pipeline: `parse_csv` -> `sort` + `build_pivot_groups` -> `write_xlsx`
- Pure dataclasses in models.py — no I/O, no business logic
- Styles module provides reusable openpyxl formatting functions
- Entry point (`__main__.py`) wires the pipeline — no business logic there

### Testing Strategy

- pytest with fixture CSV files in `tests/fixtures/`
- Unit tests for parser, sorter
- Integration tests for writer (write XLSX to tmpdir, read back with openpyxl)
- Coverage uploaded to Codecov on Python 3.13

### Git Workflow

- Conventional Commits: `feat(scope): description`, `fix(scope):`, `chore(scope):`
- Main branch protection with CI required
- Tag-triggered releases (`v*`) build Windows + Linux binaries

## Domain Context

Siemens manufactures thermal energy meters and heat cost allocators (HCA) used in Italian condominiums. Their reporting system exports CSV files that building administrators need converted to formatted Excel workbooks for billing and record-keeping.

## Important Constraints

- Must handle Italian locale CSV (semicolons, comma decimals, BOM)
- Windows drag-and-drop .exe is the primary distribution method for end users
- Output XLSX must be professionally formatted (freeze panes, auto-filter, styled headers)
- Real report data is confidential — test fixtures use anonymized data

## External Dependencies

- Runtime: openpyxl (XLSX read/write)
- Dev: pytest, pytest-cov, pyinstaller
- CI: GitHub Actions, Codecov
