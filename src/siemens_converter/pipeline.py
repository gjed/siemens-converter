"""The conversion sequence shared by the command line and the web page.

No console, message box or Windows API here: callers decide how to show the
messages carried by ``ConversionError`` and the warnings of a ``Conversion``.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

from siemens_converter.parser import (
    ReportError,
    parse_fc_report,
    warnings_with_manual_readings,
)
from siemens_converter.static_reader import (
    _TEMPLATE_ADVICE,
    StaticDataError,
    parse_static_data,
)
from siemens_converter.writer import write_xlsx

_FIND_REPORT = (
    "il file FC_report esportato dal sistema Siemens (il file .xls, il cui nome "
    "comincia con FC_report)"
)


class ConversionError(Exception):
    """A conversion that cannot go on; the message is complete and in Italian."""


@dataclass(frozen=True)
class Conversion:
    """The workbook written and the warnings the report raised."""

    output_path: Path
    warnings: list[str]


def route_files(paths: list[Path]) -> tuple[Path, Path]:
    """Split the files given into (FC_report, static data).

    Explorer does not preserve drag order, so the .xlsx is the static data and
    the other file is the FC_report. Raises ``ConversionError`` when the files
    given are not exactly one of each.
    """
    static = [p for p in paths if p.suffix.lower() == ".xlsx"]
    reports = [p for p in paths if p.suffix.lower() != ".xlsx"]
    if not reports and len(static) == 1:
        raise ConversionError(
            f"Manca il file FC_report: è stato scelto solo «{static[0].name}», che è "
            "un file .xlsx e quindi viene usato come file dei dati statici. "
            f"Scegliere anche {_FIND_REPORT} e ripetere la conversione."
        )
    if not reports:
        raise ConversionError(
            f"Sono stati scelti due file .xlsx, «{static[0].name}» e "
            f"«{static[1].name}», ma ne serve uno solo: il file dei dati statici. "
            f"Manca il file FC_report. Scegliere {_FIND_REPORT}, insieme a un solo "
            "file dei dati statici, e ripetere la conversione."
        )
    if len(reports) > 1:
        raise ConversionError(
            f"Sono stati scelti due file che non sono .xlsx, «{reports[0].name}» e "
            f"«{reports[1].name}», ma ne serve uno solo: il file FC_report. Il file "
            "dei dati statici deve essere un file .xlsx. Scegliere il file "
            "FC_report insieme al file dei dati statici e ripetere la conversione."
        )
    if not static:
        raise ConversionError(
            f"Manca il file dei dati statici: è stato scelto solo «{reports[0].name}», "
            f"il file FC_report. {_TEMPLATE_ADVICE}"
        )
    return reports[0], static[0]


def convert(
    fc_path: Path,
    static_path: Path,
    output_dir: Path,
    progress: Callable[[str], None] | None = None,
) -> Conversion:
    """Convert the FC_report and the static data into ``Riparto_<serial>_<date>.xlsx``.

    The workbook is written in *output_dir*. *progress* receives a line before
    each step. Raises ``ConversionError`` with the message to show the user.
    """
    say = progress or (lambda _line: None)

    say(f"Lettura del report {fc_path.name} ...")
    try:
        report = parse_fc_report(fc_path)
    except ReportError as exc:
        raise ConversionError(f"Errore nel report FC: {exc}") from exc

    say(f"Lettura dei dati statici da {static_path.name} ...")
    try:
        static_data = parse_static_data(static_path)
    except StaticDataError as exc:
        raise ConversionError(f"Errore nel file dati statici: {exc}") from exc
    report.warnings = warnings_with_manual_readings(report, static_data.central_meters)

    date_safe = report.header.report_date.replace("/", "-")
    output_path = output_dir / f"Riparto_{report.header.serial}_{date_safe}.xlsx"

    say(f"Scrittura di {output_path.name} ...")
    try:
        write_xlsx(report, output_path, static_data)
    except StaticDataError as exc:
        raise ConversionError(str(exc)) from exc
    except OSError as exc:
        raise ConversionError(
            f"Impossibile creare il file «{output_path.name}» nella cartella "
            f"{output_path.parent}. Se è aperto in Excel, chiuderlo; altrimenti "
            "controllare di poter scrivere in quella cartella. Poi ripetere la "
            "conversione."
        ) from exc

    return Conversion(output_path, report.warnings)
