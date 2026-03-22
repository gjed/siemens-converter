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
        traceback.print_exc()
        if sys.platform == "win32":
            input("\nPress Enter to close...")
        sys.exit(1)
