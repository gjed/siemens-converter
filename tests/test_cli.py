"""Tests for CLI entry point."""

from __future__ import annotations

import pytest
import sys
from pathlib import Path


def test_main_no_args(monkeypatch):
    monkeypatch.setattr(sys, "argv", ["siemens-converter"])
    from siemens_converter.__main__ import main

    with pytest.raises(SystemExit):
        main()


def test_main_missing_file(monkeypatch):
    monkeypatch.setattr(sys, "argv", ["siemens-converter", "nonexistent.xls"])
    from siemens_converter.__main__ import main

    with pytest.raises(SystemExit):
        main()


def test_main_success(monkeypatch, tmp_path):
    fixture = Path(__file__).parent / "fixtures" / "FC_report_TEST_9999_2026-01-15.xls"
    # Copy fixture to tmp_path so output lands there
    import shutil

    input_file = tmp_path / fixture.name
    shutil.copy(fixture, input_file)

    monkeypatch.setattr(sys, "argv", ["siemens-converter", str(input_file)])
    from siemens_converter.__main__ import main

    main()

    # Output should exist in tmp_path
    outputs = list(tmp_path.glob("Riparto_*.xlsx"))
    assert len(outputs) == 1
