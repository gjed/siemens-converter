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


_FIND_REPORT = (
    "il file FC_report esportato dal sistema Siemens (il file .xls, il cui nome "
    "comincia con FC_report)"
)


def _route_files(paths: list[Path]) -> tuple[Path, Path]:
    """Split the files given into (FC_report, static data).

    Explorer does not preserve drag order, so the .xlsx is the static data and
    the other file is the FC_report. Exits with a message when the files given
    are not exactly one of each.
    """
    static = [p for p in paths if p.suffix.lower() == ".xlsx"]
    reports = [p for p in paths if p.suffix.lower() != ".xlsx"]
    if not reports and len(static) == 1:
        _fail(
            f"Manca il file FC_report: è stato scelto solo «{static[0].name}», che è "
            "un file .xlsx e quindi viene usato come file dei dati statici. "
            f"Scegliere anche {_FIND_REPORT} e ripetere la conversione."
        )
    if not reports:
        _fail(
            f"Sono stati scelti due file .xlsx, «{static[0].name}» e "
            f"«{static[1].name}», ma ne serve uno solo: il file dei dati statici. "
            f"Manca il file FC_report. Scegliere {_FIND_REPORT}, insieme a un solo "
            "file dei dati statici, e ripetere la conversione."
        )
    if len(reports) > 1:
        _fail(
            f"Sono stati scelti due file che non sono .xlsx, «{reports[0].name}» e "
            f"«{reports[1].name}», ma ne serve uno solo: il file FC_report. Il file "
            "dei dati statici deve essere un file .xlsx. Scegliere il file "
            "FC_report insieme al file dei dati statici e ripetere la conversione."
        )
    if not static:
        from siemens_converter.static_reader import _TEMPLATE_ADVICE

        _fail(
            f"Manca il file dei dati statici: è stato scelto solo «{reports[0].name}», "
            f"il file FC_report. {_TEMPLATE_ADVICE}"
        )
    return reports[0], static[0]


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

    from siemens_converter.parser import ReportError, parse_fc_report
    from siemens_converter.static_reader import StaticDataError, parse_static_data
    from siemens_converter.writer import write_xlsx

    if _has_console():
        print(f"Lettura del report {fc_path.name} ...")

    try:
        report = parse_fc_report(fc_path)
    except ReportError as exc:
        _fail(f"Errore nel report FC: {exc}")

    if _has_console():
        print(f"Lettura dei dati statici da {static_path.name} ...")
    try:
        static_data = parse_static_data(static_path)
    except StaticDataError as exc:
        _fail(f"Errore nel file dati statici: {exc}")

    date_safe = report.header.report_date.replace("/", "-")
    out_name = f"Riparto_{report.header.serial}_{date_safe}.xlsx"
    output_path = fc_path.parent / out_name

    if _has_console():
        print(f"Scrittura di {output_path.name} ...")

    try:
        write_xlsx(report, output_path, static_data)
    except OSError:
        _fail(
            f"Impossibile creare il file «{output_path.name}» nella cartella "
            f"{output_path.parent}. Se è aperto in Excel, chiuderlo; altrimenti "
            "controllare di poter scrivere in quella cartella. Poi ripetere la "
            "conversione."
        )

    warnings = ""
    if report.warnings:
        warnings = "\n\nAttenzione:\n" + "\n".join(f"- {w}" for w in report.warnings)

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
