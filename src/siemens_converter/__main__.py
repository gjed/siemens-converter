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


def _win_open_file_dialog(title: str, filetypes: str) -> str | None:
    """Show a Windows file-open dialog.  Returns path string or None if cancelled.

    *filetypes* is a null-separated filter string, e.g.
    ``"Excel files\\0*.xls;*.xlsx\\0All files\\0*.*\\0"``.
    Only available on Windows.
    """
    try:
        import ctypes
        import ctypes.wintypes

        buf = ctypes.create_unicode_buffer(1024)

        class OPENFILENAME(ctypes.Structure):  # noqa: N801
            _fields_ = [
                ("lStructSize", ctypes.wintypes.DWORD),
                ("hwndOwner", ctypes.wintypes.HWND),
                ("hInstance", ctypes.wintypes.HINSTANCE),
                ("lpstrFilter", ctypes.wintypes.LPCWSTR),
                ("lpstrCustomFilter", ctypes.c_wchar_p),
                ("nMaxCustFilter", ctypes.wintypes.DWORD),
                ("nFilterIndex", ctypes.wintypes.DWORD),
                ("lpstrFile", ctypes.c_wchar_p),
                ("nMaxFile", ctypes.wintypes.DWORD),
                ("lpstrFileTitle", ctypes.c_wchar_p),
                ("nMaxFileTitle", ctypes.wintypes.DWORD),
                ("lpstrInitialDir", ctypes.c_wchar_p),
                ("lpstrTitle", ctypes.wintypes.LPCWSTR),
                ("Flags", ctypes.wintypes.DWORD),
                ("nFileOffset", ctypes.wintypes.WORD),
                ("nFileExtension", ctypes.wintypes.WORD),
                ("lpstrDefExt", ctypes.c_wchar_p),
                ("lCustData", ctypes.wintypes.LPARAM),
                ("lpfnHook", ctypes.c_void_p),
                ("lpTemplateName", ctypes.c_wchar_p),
            ]

        ofn = OPENFILENAME()
        ofn.lStructSize = ctypes.sizeof(OPENFILENAME)
        ofn.lpstrFilter = filetypes
        ofn.lpstrFile = ctypes.cast(buf, ctypes.c_wchar_p)
        ofn.nMaxFile = 1024
        ofn.lpstrTitle = title
        ofn.Flags = 0x00080000 | 0x00001000  # OFN_EXPLORER | OFN_FILEMUSTEXIST

        if ctypes.windll.comdlg32.GetOpenFileNameW(ctypes.byref(ofn)):  # type: ignore[union-attr]
            return buf.value or None
        return None
    except Exception:
        return None


def main() -> None:
    """Run the converter from CLI or drag-and-drop."""
    fc_path: Path | None = None
    static_path: Path | None = None

    if len(sys.argv) >= 2:
        # CLI / drag-and-drop mode: first arg is FC_report
        fc_path = Path(sys.argv[1])
        if len(sys.argv) >= 3:
            static_path = Path(sys.argv[2])
    elif not _has_console():
        # No-console (PyInstaller --noconsole): use Windows file dialogs
        fc_str = _win_open_file_dialog(
            "Seleziona FC_report (.xls)",
            "FC Report\0*.xls\0All files\0*.*\0",
        )
        if fc_str is None:
            sys.exit(0)  # User cancelled
        fc_path = Path(fc_str)

        # Second dialog for static data (optional — user can cancel)
        static_str = _win_open_file_dialog(
            "Seleziona dati statici (.xlsx) — o Annulla per saltare",
            "Excel files\0*.xlsx\0All files\0*.*\0",
        )
        if static_str is not None:
            static_path = Path(static_str)
    else:
        msg = (
            "Usage: siemens-converter <FC_report.xls> [static_data.xlsx]\n"
            "\n"
            "Arguments:\n"
            "  FC_report.xls      Siemens FC_report file (required)\n"
            "  static_data.xlsx   Static condominium data file (optional)"
        )
        print(msg)
        sys.exit(1)

    # Validate FC_report
    if not fc_path.exists():
        msg = f"File not found: {fc_path}"
        if _has_console():
            print(msg)
        else:
            _msgbox("Errore", msg)
        sys.exit(1)

    # Validate static data file if provided
    if static_path is not None and not static_path.exists():
        msg = f"Static data file not found: {static_path}"
        if _has_console():
            print(msg)
        else:
            _msgbox("Errore", msg)
        sys.exit(1)

    from siemens_converter.parser import parse_fc_report
    from siemens_converter.writer import write_xlsx

    if _has_console():
        print(f"Parsing {fc_path.name} ...")

    report = parse_fc_report(fc_path)

    # Parse static data if provided
    static_data = None
    if static_path is not None:
        from siemens_converter.static_reader import parse_static_data

        if _has_console():
            print(f"Reading static data from {static_path.name} ...")
        try:
            static_data = parse_static_data(static_path)
        except Exception as exc:
            msg = f"Errore nel file dati statici: {exc}"
            if _has_console():
                print(msg)
            else:
                _msgbox("Errore", msg)
            sys.exit(1)

    date_safe = report.header.report_date.replace("/", "-")
    out_name = f"Riparto_{report.header.serial}_{date_safe}.xlsx"
    output_path = fc_path.parent / out_name

    if _has_console():
        print(f"Writing {output_path.name} ...")

    write_xlsx(report, output_path, static_data=static_data)

    if _has_console():
        print(f"Done: {output_path}")
        if sys.platform == "win32":
            input("\nPress Enter to close...")
    else:
        _msgbox("siemens-converter", f"Fatto!\n\n{output_path.name}")


if __name__ == "__main__":
    main()
