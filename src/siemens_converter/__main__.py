"""CLI entry point for siemens-converter."""

from __future__ import annotations

import sys
from pathlib import Path


def main() -> None:
    """Run the converter from CLI or drag-and-drop."""
    if len(sys.argv) < 2:
        print("Usage: siemens-converter <path/to/FC_report.xls>")
        sys.exit(1)

    csv_path = Path(sys.argv[1])
    if not csv_path.exists():
        print(f"File not found: {csv_path}")
        sys.exit(1)

    from siemens_converter.parser import parse_fc_report
    from siemens_converter.writer import write_xlsx

    print(f"Parsing {csv_path.name} ...")
    report = parse_fc_report(csv_path)

    date_safe = report.header.report_date.replace("/", "-")
    out_name = f"Riparto_{report.header.serial}_{date_safe}.xlsx"
    output_path = csv_path.parent / out_name

    print(f"Writing {output_path.name} ...")
    write_xlsx(report, output_path)
    print(f"Done: {output_path}")

    if sys.platform == "win32":
        input("\nPress Enter to close...")


if __name__ == "__main__":
    main()
