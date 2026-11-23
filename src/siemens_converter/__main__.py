"""CLI entry point for siemens-converter."""

from __future__ import annotations

import sys
from pathlib import Path
from typing import NoReturn


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
        _msgbox(
            "Errore",
            "Impossibile aprire la finestra di selezione dei file. "
            "Trascinare i file da convertire sopra l'icona del programma.",
        )
        return None


def _fail(message: str) -> NoReturn:
    """Show *message* (console, or message box without one) and exit with code 1."""
    if _has_console():
        print(message)
    else:
        _msgbox("Errore", message)
    sys.exit(1)


def _route_files(paths: list[Path]) -> tuple[Path, Path]:
    """Split the files given into (FC_report, static data), or exit with a message."""
    from siemens_converter.pipeline import ConversionError, route_files

    try:
        return route_files(paths)
    except ConversionError as exc:
        _fail(str(exc))


def main() -> None:
    """Run the converter from CLI or drag-and-drop."""
    if len(sys.argv) >= 2:
        # CLI / drag-and-drop mode
        fc_path, static_path = _route_files([Path(a) for a in sys.argv[1:3]])
    elif not _has_console():
        # No-console (PyInstaller --noconsole): use Windows file dialogs
        fc_str = _win_open_file_dialog(
            "Seleziona FC_report (.xls)",
            "FC Report\0*.xls\0All files\0*.*\0",
        )
        if fc_str is None:
            sys.exit(0)  # User cancelled

        # Second dialog for static data (required: cancelling is refused by _route_files)
        static_str = _win_open_file_dialog(
            "Seleziona dati statici (.xlsx)",
            "Excel files\0*.xlsx\0All files\0*.*\0",
        )
        picked = [Path(fc_str)] + ([Path(static_str)] if static_str else [])
        fc_path, static_path = _route_files(picked)
    else:
        print(
            "Uso: siemens-converter <FC_report.xls> <dati_statici.xlsx>\n"
            "\n"
            "Argomenti:\n"
            "  FC_report.xls      il file esportato dal sistema Siemens (obbligatorio)\n"
            "  dati_statici.xlsx  il file dei dati statici (obbligatorio)"
        )
        sys.exit(1)

    for path, what in ((fc_path, "FC_report"), (static_path, "dei dati statici")):
        if not path.exists():
            _fail(
                f"Il file {what} «{path}» non esiste. Controllare che non sia stato "
                "spostato o rinominato e ripetere la conversione."
            )

    from siemens_converter.pipeline import ConversionError, convert

    def progress(line: str) -> None:
        if _has_console():
            print(line)

    try:
        result = convert(fc_path, static_path, fc_path.parent, progress)
    except ConversionError as exc:
        _fail(str(exc))
    output_path = result.output_path

    warnings = ""
    if result.warnings:
        warnings = "\n\nAttenzione:\n" + "\n".join(f"- {w}" for w in result.warnings)

    if _has_console():
        if warnings:
            print(warnings.lstrip("\n") + "\n")
        print(f"Fatto: {output_path}")
        if sys.platform == "win32":
            input("\nPremere Invio per chiudere...")
    else:
        _msgbox("siemens-converter", f"Fatto!\n\n{output_path.name}{warnings}")


if __name__ == "__main__":
    main()
