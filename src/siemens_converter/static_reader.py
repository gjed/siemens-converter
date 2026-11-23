"""Parse static condominium data from a user-populated XLSX file."""

from __future__ import annotations

from collections.abc import Callable, Iterable, Iterator
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.utils import get_column_letter

from siemens_converter.models import (
    ApartmentInfo,
    CostItem,
    MeterReading,
    Millesimali,
    PreviousReading,
    StaticData,
)

REQUIRED_SHEETS = frozenset(
    {"Inquilini", "Millesimali", "Costi", "Contatori", "Letture_precedenti"}
)

# Labels expected in the Costi sheet (order matches Tabella E3-E8)
COST_LABELS = (
    "Energia elettrica",
    "Gas metano",
    "Acqua condominio",
    "Conduzione e manutenzione",
    "Contabilizzazione",
    "Acqua sanitaria manutenzione",
)

_TEMPLATE_ADVICE = (
    "Scegliere il file dei dati statici, compilato a partire da "
    "static_data_template.xlsx (si scarica dalla pagina dell'ultima versione: "
    "https://github.com/gjed/siemens-converter/releases/latest), "
    "e ripetere la conversione."
)


class StaticDataError(ValueError):
    """A static data file that cannot be read; the message is for the user."""


def _quoted(names: Iterable[str]) -> str:
    return ", ".join(f"«{name}»" for name in names)


def parse_static_data(path: str | Path) -> StaticData:
    """Read a user-populated static data XLSX and return a StaticData object.

    Raises StaticDataError, with a message for the user, if the file is not a
    static data file or holds a value that cannot be read.
    Missing cost rows default to 0.0; fewer apartment rows are silently accepted.
    """
    path = Path(path)
    name = path.name
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as exc:  # openpyxl raises many types for a bad file
        raise StaticDataError(
            f"«{name}» non è un file dei dati statici: non si apre come file "
            f"Excel (.xlsx). {_TEMPLATE_ADVICE}"
        ) from exc

    missing = sorted(REQUIRED_SHEETS - set(wb.sheetnames))
    if missing:
        sheets = (
            f"manca il foglio {_quoted(missing)}"
            if len(missing) == 1
            else f"mancano i fogli {_quoted(missing)}"
        )
        raise StaticDataError(
            f"«{name}» non è un file dei dati statici: {sheets}. Se il file è "
            "quello giusto, controllare che nessun foglio sia stato rinominato "
            f"o eliminato. {_TEMPLATE_ADVICE}"
        )

    apartments = _read_inquilini(wb["Inquilini"])
    millesimali = _read_millesimali(name, wb["Millesimali"])
    costs = _read_costi(name, wb["Costi"])
    meters = _read_contatori(name, wb["Contatori"])
    previous_readings = _read_letture_precedenti(name, wb["Letture_precedenti"])

    return StaticData(
        apartments=apartments,
        millesimali=millesimali,
        costs=costs,
        meters=meters,
        previous_readings=previous_readings,
    )


_DATA_START_ROW = 3  # row 1 = section title, row 2 = column headers, row 3 = first data


def _rows(ws: openpyxl.worksheet.worksheet.Worksheet) -> Iterator[tuple[int, tuple]]:
    """Data rows of a sheet as (row number, values)."""
    return enumerate(
        ws.iter_rows(min_row=_DATA_START_ROW, values_only=True), start=_DATA_START_ROW
    )


def _number(
    file: str,
    ws: openpyxl.worksheet.worksheet.Worksheet,
    row_no: int,
    row: tuple,
    col: int,
    convert: Callable[[Any], Any],
) -> Any:
    """Read a number from a cell: blank reads as 0, unreadable text is refused."""
    raw = row[col] if len(row) > col else None
    if raw is None:
        return convert(0)
    try:
        return convert(raw)
    except (ValueError, TypeError) as exc:
        kind = "un numero intero" if convert is int else "un numero"
        raise StaticDataError(
            f"«{file}», foglio «{ws.title}», cella {get_column_letter(col + 1)}{row_no}: "
            f"il valore «{raw}» non è {kind}. Scrivere solo il numero, senza unità "
            "di misura e senza simboli, e ripetere la conversione."
        ) from exc


def _read_inquilini(ws: openpyxl.worksheet.worksheet.Worksheet) -> list[ApartmentInfo]:
    """Read Inquilini sheet: App.number, Proprietario, Inquilino."""
    result: list[ApartmentInfo] = []
    for row in ws.iter_rows(min_row=_DATA_START_ROW, values_only=True):
        apt_num = row[0] if len(row) > 0 else None
        if apt_num is None:
            break
        try:
            apt_num = int(apt_num)
        except (ValueError, TypeError):
            break
        result.append(
            ApartmentInfo(
                apartment_number=apt_num,
                proprietario=str(row[1]) if len(row) > 1 and row[1] else "",
                inquilino=str(row[2]) if len(row) > 2 and row[2] else "",
            )
        )
    return result


def _read_millesimali(
    file: str, ws: openpyxl.worksheet.worksheet.Worksheet
) -> list[Millesimali]:
    """Read Millesimali sheet: App., Subalterno, Energia riscaldamento kWh, Energia ACS kWh.

    Column layout: App(1), Subalterno(2), Heat kWh(3), Mill.risc formula(4), ACS kWh(5), Mill.ACS formula(6).
    Non-integer rows (e.g. TOTALE) are skipped rather than stopping iteration.
    """
    result: list[Millesimali] = []
    for row_no, row in _rows(ws):
        apt_num = row[0] if len(row) > 0 else None
        if apt_num is None:
            continue  # skip empty rows
        try:
            apt_num = int(apt_num)
        except (ValueError, TypeError):
            continue  # skip TOTALE and other non-data rows
        result.append(
            Millesimali(
                apartment_number=apt_num,
                subalterno=_number(file, ws, row_no, row, 1, int),
                heat_energy_kwh=_number(file, ws, row_no, row, 2, float),
                # Column 5 (index 4) = ACS kWh; column 4 (index 3) may be a formula cell
                water_energy_kwh=_number(file, ws, row_no, row, 4, float),
            )
        )
    return result


def _read_costi(
    file: str, ws: openpyxl.worksheet.worksheet.Worksheet
) -> list[CostItem]:
    """Read Costi sheet: Voce, Importo.

    Labels match case-insensitively, ignoring surrounding whitespace. Missing
    expected labels default to 0.0; unknown labels raise StaticDataError.
    """
    known = {lbl.casefold() for lbl in COST_LABELS}
    found: dict[str, float] = {}
    unknown: list[str] = []
    for row_no, row in _rows(ws):
        label = row[0] if len(row) > 0 else None
        if label is None:
            break
        # Skip the TOTALE row or any non-cost label
        if str(label).upper() == "TOTALE":
            continue
        found[str(label).strip().casefold()] = _number(file, ws, row_no, row, 1, float)
        if str(label).strip().casefold() not in known:
            unknown.append(f"- «{str(label).strip()}» (cella A{row_no})")

    if unknown:
        raise StaticDataError(
            f"«{file}», foglio «Costi»: queste voci di spesa non sono tra quelle "
            "accettate:\n" + "\n".join(unknown) + "\n\nLe voci accettate sono: "
            f"{_quoted(COST_LABELS)}. Scrivere ogni voce come nel modello, oppure "
            "eliminare la riga, e ripetere la conversione."
        )

    result: list[CostItem] = []
    for label in COST_LABELS:
        result.append(CostItem(label=label, amount=found.get(label.casefold(), 0.0)))
    return result


def _read_contatori(
    file: str, ws: openpyxl.worksheet.worksheet.Worksheet
) -> list[MeterReading]:
    """Read Contatori sheet: Contatore, Unità, Lettura iniziale, Lettura finale."""
    result: list[MeterReading] = []
    for row_no, row in _rows(ws):
        name = row[0] if len(row) > 0 else None
        if name is None:
            break
        result.append(
            MeterReading(
                name=str(name),
                unit=str(row[1]) if len(row) > 1 and row[1] else "",
                initial=_number(file, ws, row_no, row, 2, float),
                final=_number(file, ws, row_no, row, 3, float),
            )
        )
    return result


def _read_letture_precedenti(
    file: str, ws: openpyxl.worksheet.worksheet.Worksheet
) -> list[PreviousReading]:
    """Read Letture_precedenti sheet: App., Riscaldamento kWh, ACS m3, AFS m3."""
    result: list[PreviousReading] = []
    for row_no, row in _rows(ws):
        apt_num = row[0] if len(row) > 0 else None
        if apt_num is None:
            break
        try:
            apt_num = int(apt_num)
        except (ValueError, TypeError):
            break
        result.append(
            PreviousReading(
                apartment_number=apt_num,
                heat_kwh=_number(file, ws, row_no, row, 1, int),
                water_m3=_number(file, ws, row_no, row, 2, float),
                cold_water_m3=_number(file, ws, row_no, row, 3, float),
            )
        )
    return result
