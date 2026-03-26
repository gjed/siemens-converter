"""CLI entry point for siemens-converter."""

from __future__ import annotations

import sys
from pathlib import Path


def _has_console() -> bool:
    """Check if we have a working console (False when --noconsole PyInstaller)."""
    return sys.stdin is not None and sys.stdout is not None


def _msgbox(title: str, text: str) -> None:
    """Show a Windows message box (fallback when no console)."""
    try:
        import ctypes

        ctypes.windll.user32.MessageBoxW(0, text, title, 0)  # type: ignore[union-attr]
    except Exception:
        pass


def main() -> None:
    """Run the converter from CLI or drag-and-drop."""
    if len(sys.argv) < 2:
        msg = "Usage: siemens-converter <path/to/FC_report.xls>"
        if _has_console():
            print(msg)
        else:
            _msgbox("siemens-converter", msg)
        sys.exit(1)

    csv_path = Path(sys.argv[1])
    if not csv_path.exists():
        msg = f"File not found: {csv_path}"
        if _has_console():
            print(msg)
        else:
            _msgbox("Errore", msg)
        sys.exit(1)

    from siemens_converter.parser import parse_fc_report
    from siemens_converter.writer import write_xlsx

    if _has_console():
        print(f"Parsing {csv_path.name} ...")

    report = parse_fc_report(csv_path)

    date_safe = report.header.report_date.replace("/", "-")
    out_name = f"Riparto_{report.header.serial}_{date_safe}.xlsx"
    output_path = csv_path.parent / out_name

    if _has_console():
        print(f"Writing {output_path.name} ...")

    write_xlsx(report, output_path)

    if _has_console():
        print(f"Done: {output_path}")
        if sys.platform == "win32":
            input("\nPress Enter to close...")
    else:
        _msgbox("siemens-converter", f"Fatto!\n\n{output_path.name}")


if __name__ == "__main__":
    main()
