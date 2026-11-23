"""PyInstaller entry point for --onefile builds."""

from __future__ import annotations

import sys
import traceback

FAILURE = (
    "La conversione non è riuscita per un errore imprevisto del programma. "
    "Inviare i dettagli tecnici qui sotto a chi ha fornito il programma."
)


def main() -> None:
    from siemens_converter.__main__ import main as run

    run()


def report_failure(details: str) -> None:
    """Show the Italian sentence first, then the technical *details*."""
    message = f"{FAILURE}\n\nDettagli tecnici:\n{details}"
    if sys.stdout is not None:
        print(message, file=sys.stderr)
        if sys.platform == "win32" and sys.stdin is not None:
            input("\nPremere Invio per chiudere...")
    elif sys.platform == "win32":
        try:
            import ctypes

            ctypes.windll.user32.MessageBoxW(  # type: ignore[union-attr]
                0,
                message,
                "siemens-converter - Errore",
                0,
            )
        except Exception:
            pass


if __name__ == "__main__":
    try:
        main()
    except Exception:
        report_failure(traceback.format_exc())
        sys.exit(1)
