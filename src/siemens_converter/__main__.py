"""CLI entry point for siemens-converter."""

from __future__ import annotations

import sys
from pathlib import Path


def main() -> None:
    """Run the converter from CLI or drag-and-drop."""
    # TODO: implement CLI logic
    if len(sys.argv) < 2:
        print("Usage: siemens-converter <path/to/report.csv>")
        sys.exit(1)

    csv_path = Path(sys.argv[1])
    if not csv_path.exists():
        print(f"File not found: {csv_path}")
        sys.exit(1)

    print(f"Converting {csv_path} ...")
    # TODO: wire parse -> sort -> write pipeline


if __name__ == "__main__":
    main()
