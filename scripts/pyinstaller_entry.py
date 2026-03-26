"""PyInstaller entry point for --onefile builds."""

from __future__ import annotations

import sys
import traceback


def main() -> None:
    from siemens_converter.__main__ import main as run

    run()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        msg = traceback.format_exc()
        if sys.stdout is not None:
            traceback.print_exc()
            if sys.platform == "win32" and sys.stdin is not None:
                input("\nPress Enter to close...")
        elif sys.platform == "win32":
            try:
                import ctypes

                ctypes.windll.user32.MessageBoxW(  # type: ignore[union-attr]
                    0,
                    msg,
                    "siemens-converter - Errore",
                    0,
                )
            except Exception:
                pass
        sys.exit(1)
