"""Parse static condominium data from a user-populated XLSX file."""

from __future__ import annotations

from collections.abc import Callable, Iterable, Iterator
from datetime import date, datetime
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.utils import get_column_letter

from siemens_converter.models import (
    ApartmentInfo,
    CentralMeterStart,
    CostItem,
    ExerciseInfo,
    MeterReading,
    Millesimali,
    ParsedReport,
    PreviousReading,
    StaticData,
)

REQUIRED_SHEETS = frozenset(
    {
        "Inquilini",
        "Millesimali",
        "Costi",
        "Contatori",
        "Letture_precedenti",
        "Esercizio",
    }
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

# Labels expected in the Contatori sheet (order matches Tabella F20-G22)
METER_LABELS = (
    "ENERGIA ELETTRICA CT",
    "GAS METANO CT",
    "ACQUA GENERALE",
)

APARTMENTS = tuple(range(1, 11))

_TEMPLATE_ADVICE = (
    "Scegliere il file dei dati statici, compilato a partire da "
    "static_data_template.xlsx (si scarica dalla pagina dell'ultima versione: "
    "https://github.com/gjed/siemens-converter/releases/latest), "
    "e ripetere la conversione."
)

# Wording of the sheets whose rows are told apart by the label in column A
_LABELLED_SHEETS = {
    "Costi": {
        "labels": COST_LABELS,
        "unknown": "queste voci di spesa non sono tra quelle accettate",
        "accepted": "Le voci accettate sono",
        "each": "ogni voce",
        "one": "la voce",
        "missing_one": "manca la voce",
        "missing_many": "mancano le voci",
        "of_one": "della voce",
        "of_many": "delle voci",
    },
    "Contatori": {
        "labels": METER_LABELS,
        "unknown": "questi contatori non sono tra quelli accettati",
        "accepted": "I contatori accettati sono",
        "each": "ogni contatore",
        "one": "il contatore",
        "missing_one": "manca il contatore",
        "missing_many": "mancano i contatori",
        "of_one": "del contatore",
        "of_many": "dei contatori",
    },
}


class StaticDataError(ValueError):
    """A static data file that cannot be read; the message is for the user."""


def _quoted(names: Iterable[str]) -> str:
    return ", ".join(f"«{name}»" for name in names)


def _join(items: Iterable[Any]) -> str:
    """'3', '3 e 7', '3, 7 e 9'."""
    parts = [str(item) for item in items]
    return " e ".join([", ".join(parts[:-1]), parts[-1]] if len(parts) > 1 else parts)


def _sheet_error(file: str, ws: Any, problem: str) -> StaticDataError:
    return StaticDataError(f"«{file}», foglio «{ws.title}»: {problem}")


def _cell_error(file: str, ws: Any, cell: str, problem: str) -> StaticDataError:
    return StaticDataError(f"«{file}», foglio «{ws.title}», cella {cell}: {problem}")


def parse_static_data(path: str | Path) -> StaticData:
    """Read a user-populated static data XLSX and return a StaticData object.

    Raises StaticDataError, with a message for the user, if the file is not a
    static data file or is incomplete or inconsistent: a blank or non-numeric
    number cell, an apartment missing or listed twice, a missing cost or meter
    row, a row that is neither data nor TOTALE, or missing exercise data.
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

    apartments = _read_inquilini(name, wb["Inquilini"])
    millesimali = _read_millesimali(name, wb["Millesimali"])
    costs = _read_costi(name, wb["Costi"])
    meters = _read_contatori(name, wb["Contatori"])
    previous_readings = _read_letture_precedenti(name, wb["Letture_precedenti"])
    exercise, central_meters = _read_esercizio(name, wb["Esercizio"])

    return StaticData(
        apartments=apartments,
        millesimali=millesimali,
        costs=costs,
        meters=meters,
        previous_readings=previous_readings,
        exercise=exercise,
        central_meters=central_meters,
    )


_DATA_START_ROW = 3  # row 1 = section title, row 2 = column headers, row 3 = first data
_MILLESIMALI_START_ROW = 5  # rows 3 and 4 hold the units and the TOTALE formulas


def _is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _data_rows(
    ws: openpyxl.worksheet.worksheet.Worksheet, start_row: int = _DATA_START_ROW
) -> Iterator[tuple[int, tuple]]:
    """Rows of a sheet as (row number, values), without blank rows and TOTALE."""
    for row_no, row in enumerate(
        ws.iter_rows(min_row=start_row, values_only=True), start=start_row
    ):
        if all(_is_blank(value) for value in row):
            continue
        if isinstance(row[0], str) and row[0].strip().upper() == "TOTALE":
            continue
        yield row_no, row


def _integer(raw: Any) -> int:
    """An integer written as a number or as text; 3.5 and True are refused."""
    if isinstance(raw, bool):
        raise ValueError(raw)
    if isinstance(raw, float):
        if not raw.is_integer():
            raise ValueError(raw)
        return int(raw)
    return int(raw)


def _number(
    file: str,
    ws: openpyxl.worksheet.worksheet.Worksheet,
    row_no: int,
    row: tuple,
    col: int,
    convert: Callable[[Any], Any],
) -> Any:
    """Read a number from a cell: a blank or unreadable cell is refused."""
    raw = row[col] if len(row) > col else None
    cell = f"{get_column_letter(col + 1)}{row_no}"
    if _is_blank(raw):
        raise _cell_error(
            file,
            ws,
            cell,
            "la cella è vuota ma serve un numero. Scrivere il numero (se è "
            "davvero zero, scrivere 0) e ripetere la conversione.",
        )
    try:
        if isinstance(raw, bool):
            raise ValueError(raw)
        return convert(raw)
    except (ValueError, TypeError) as exc:
        kind = "un numero intero" if convert is int else "un numero"
        raise _cell_error(
            file,
            ws,
            cell,
            f"il valore «{raw}» non è {kind}. Scrivere solo il numero, senza "
            "unità di misura e senza simboli, e ripetere la conversione.",
        ) from exc


def _apartment_rows(
    file: str,
    ws: openpyxl.worksheet.worksheet.Worksheet,
    start_row: int = _DATA_START_ROW,
) -> list[tuple[int, int, tuple]]:
    """The rows of a per-apartment sheet as (apartment, row number, values).

    Every apartment from 1 to 10 must be there exactly once; the rows come
    back in apartment order.
    """
    found: dict[int, tuple[int, tuple]] = {}
    for row_no, row in _data_rows(ws, start_row):
        raw = row[0]
        if _is_blank(raw):
            raise _cell_error(
                file,
                ws,
                f"A{row_no}",
                "manca il numero dell'appartamento, ma la riga contiene altri "
                "valori. Scrivere il numero dell'appartamento (da 1 a 10) o "
                "cancellare la riga, e ripetere la conversione.",
            )
        try:
            apartment = _integer(raw)
        except (ValueError, TypeError) as exc:
            raise _cell_error(
                file,
                ws,
                f"A{row_no}",
                f"il valore «{raw}» non è un numero di appartamento. In questo "
                "foglio sono ammesse solo le righe degli appartamenti da 1 a 10 "
                "e la riga TOTALE: correggere il valore o cancellare la riga, e "
                "ripetere la conversione.",
            ) from exc
        if apartment not in APARTMENTS:
            raise _cell_error(
                file,
                ws,
                f"A{row_no}",
                f"l'appartamento «{apartment}» non esiste: gli appartamenti vanno "
                "da 1 a 10. Correggere il numero o cancellare la riga, e "
                "ripetere la conversione.",
            )
        if apartment in found:
            raise _sheet_error(
                file,
                ws,
                f"l'appartamento {apartment} compare due volte (celle "
                f"A{found[apartment][0]} e A{row_no}). Lasciare una sola riga "
                "per ogni appartamento e ripetere la conversione.",
            )
        found[apartment] = (row_no, row)

    absent = [apartment for apartment in APARTMENTS if apartment not in found]
    if absent:
        problem = (
            f"manca l'appartamento {absent[0]}. Aggiungere la riga "
            "dell'appartamento, con il suo numero nella colonna A, e ripetere "
            "la conversione."
            if len(absent) == 1
            else f"mancano gli appartamenti {_join(absent)}. Aggiungere la riga "
            "di ciascuno, con il suo numero nella colonna A, e ripetere la "
            "conversione."
        )
        raise _sheet_error(file, ws, problem)
    return [(apartment, *found[apartment]) for apartment in APARTMENTS]


def _read_inquilini(
    file: str, ws: openpyxl.worksheet.worksheet.Worksheet
) -> list[ApartmentInfo]:
    """Read Inquilini sheet: App.number, Proprietario, Inquilino."""
    return [
        ApartmentInfo(
            apartment_number=apartment,
            proprietario=str(row[1]) if len(row) > 1 and row[1] else "",
            inquilino=str(row[2]) if len(row) > 2 and row[2] else "",
        )
        for apartment, _, row in _apartment_rows(file, ws)
    ]


def _read_millesimali(
    file: str, ws: openpyxl.worksheet.worksheet.Worksheet
) -> list[Millesimali]:
    """Read Millesimali sheet: App., Subalterno, Energia riscaldamento kWh, Energia ACS kWh.

    Column layout: App(1), Subalterno(2), Heat kWh(3), Mill.risc formula(4), ACS kWh(5), Mill.ACS formula(6).
    Rows 3 and 4 hold the units and the TOTALE formulas; the data starts at row 5.
    """
    return [
        Millesimali(
            apartment_number=apartment,
            subalterno=_number(file, ws, row_no, row, 1, int),
            heat_energy_kwh=_number(file, ws, row_no, row, 2, float),
            # Column 5 (index 4) = ACS kWh; column 4 (index 3) may be a formula cell
            water_energy_kwh=_number(file, ws, row_no, row, 4, float),
        )
        for apartment, row_no, row in _apartment_rows(file, ws, _MILLESIMALI_START_ROW)
    ]


def _labelled_rows(
    file: str, ws: openpyxl.worksheet.worksheet.Worksheet
) -> list[tuple[str, int, tuple]]:
    """The rows of Costi or Contatori as (label, row number, values).

    Every expected label must be there exactly once, in any case and with any
    surrounding whitespace; the rows come back in the order of the labels.
    """
    wording = _LABELLED_SHEETS[ws.title]
    labels = wording["labels"]
    known = {label.casefold(): label for label in labels}
    found: dict[str, tuple[int, tuple]] = {}
    unknown: list[str] = []
    for row_no, row in _data_rows(ws):
        raw = row[0]
        if _is_blank(raw):
            raise _cell_error(
                file,
                ws,
                f"A{row_no}",
                f"manca il nome ({wording['one']}), ma la riga contiene altri "
                "valori. Scriverlo come nel modello o cancellare la riga, e "
                "ripetere la conversione.",
            )
        label = known.get(str(raw).strip().casefold())
        if label is None:
            unknown.append(f"- «{str(raw).strip()}» (cella A{row_no})")
        elif label in found:
            raise _sheet_error(
                file,
                ws,
                f"{wording['one']} «{label}» compare due volte (celle "
                f"A{found[label][0]} e A{row_no}). Lasciare una sola riga e "
                "ripetere la conversione.",
            )
        else:
            found[label] = (row_no, row)

    if unknown:
        raise _sheet_error(
            file,
            ws,
            f"{wording['unknown']}:\n"
            + "\n".join(unknown)
            + f"\n\n{wording['accepted']}: {_quoted(labels)}. Scrivere "
            f"{wording['each']} come nel modello, oppure eliminare la riga, e "
            "ripetere la conversione.",
        )
    absent = [label for label in labels if label not in found]
    if absent:
        many = len(absent) > 1
        raise _sheet_error(
            file,
            ws,
            f"{wording['missing_many' if many else 'missing_one']} "
            f"{_quoted(absent)}. Aggiungere la riga "
            f"{wording['of_many' if many else 'of_one']}, scritta come nel modello "
            "e con i suoi valori (se un valore è davvero zero, scrivere 0), e "
            "ripetere la conversione.",
        )
    return [(label, *found[label]) for label in labels]


def _read_costi(
    file: str, ws: openpyxl.worksheet.worksheet.Worksheet
) -> list[CostItem]:
    """Read Costi sheet: Voce, Importo.

    Every cost label is required, matched case-insensitively and ignoring
    surrounding whitespace; a cost that is really zero is entered as 0.
    """
    return [
        CostItem(label=label, amount=_number(file, ws, row_no, row, 1, float))
        for label, row_no, row in _labelled_rows(file, ws)
    ]


def _read_contatori(
    file: str, ws: openpyxl.worksheet.worksheet.Worksheet
) -> list[MeterReading]:
    """Read Contatori sheet: Contatore, Unità, Lettura iniziale, Lettura finale.

    A final reading lower than the initial one is refused; equal is valid.
    """
    meters = []
    for label, row_no, row in _labelled_rows(file, ws):
        initial = _number(file, ws, row_no, row, 2, float)
        final = _number(file, ws, row_no, row, 3, float)
        if final < initial:
            raise _cell_error(
                file,
                ws,
                f"D{row_no}",
                f"la lettura finale di «{label}» ({_amount(final)}) è più bassa "
                f"della lettura iniziale ({_amount(initial)}, cella C{row_no}). "
                "Controllare le due letture, scrivere quella sbagliata e ripetere "
                "la conversione.",
            )
        unit = str(row[1]) if len(row) > 1 and row[1] else ""
        meters.append(MeterReading(label, unit, initial, final))
    return meters


def _read_letture_precedenti(
    file: str, ws: openpyxl.worksheet.worksheet.Worksheet
) -> list[PreviousReading]:
    """Read Letture_precedenti sheet: App., Riscaldamento kWh, ACS m3, AFS m3."""
    return [
        PreviousReading(
            apartment_number=apartment,
            heat_kwh=_number(file, ws, row_no, row, 1, int),
            water_m3=_number(file, ws, row_no, row, 2, float),
            cold_water_m3=_number(file, ws, row_no, row, 3, float),
        )
        for apartment, row_no, row in _apartment_rows(file, ws)
    ]


# Esercizio sheet: year in C3, start date in C4, then one row per central meter,
# with the start reading in C, the optional manual final reading in D and its date in E
_YEAR_ROW, _START_DATE_ROW = 3, 4
_CENTRAL_METER_ROWS = {"Riscaldamento": 5, "Sanitario": 6}
_DATE_ADVICE = "Scrivere la data nel formato gg/mm/aaaa (per esempio 31/12/2025)"


def _date(
    file: str,
    ws: openpyxl.worksheet.worksheet.Worksheet,
    cell: str,
    raw: Any,
    what: str,
) -> datetime:
    """A date typed in a date cell; blank and text are refused."""
    if isinstance(raw, datetime):
        return raw
    if isinstance(raw, date):
        return datetime(raw.year, raw.month, raw.day)
    problem = (
        f"{what} è vuota" if _is_blank(raw) else f"il valore «{raw}» non è una data"
    )
    raise _cell_error(
        file, ws, cell, f"{problem}. {_DATE_ADVICE} e ripetere la conversione."
    )


def _read_esercizio(
    file: str, ws: openpyxl.worksheet.worksheet.Worksheet
) -> tuple[ExerciseInfo, list[CentralMeterStart]]:
    """Read the Esercizio sheet: year, start date and the two central heat meters."""
    rows = {
        row_no: row
        for row_no, row in enumerate(
            ws.iter_rows(min_row=_YEAR_ROW, max_row=6, values_only=True),
            start=_YEAR_ROW,
        )
    }

    def cell(row_no: int, col: int) -> Any:
        row = rows.get(row_no, ())
        return row[col] if len(row) > col else None

    raw_year = cell(_YEAR_ROW, 2)
    year_advice = "Scrivere l'anno con quattro cifre (per esempio 2026) e ripetere la conversione."
    if _is_blank(raw_year):
        raise _cell_error(
            file,
            ws,
            f"C{_YEAR_ROW}",
            f"l'anno dell'esercizio è vuoto. {year_advice}",
        )
    try:
        year = _integer(raw_year)
        if not 1000 <= year <= 9999:
            raise ValueError(raw_year)
    except (ValueError, TypeError) as exc:
        raise _cell_error(
            file,
            ws,
            f"C{_YEAR_ROW}",
            f"il valore «{raw_year}» non è un anno. {year_advice}",
        ) from exc
    start_date = _date(
        file,
        ws,
        f"C{_START_DATE_ROW}",
        cell(_START_DATE_ROW, 2),
        "la data di inizio dell'esercizio",
    )

    meters = []
    for description, row_no in _CENTRAL_METER_ROWS.items():
        row = rows.get(row_no, ())
        initial = _number(file, ws, row_no, row, 2, float)
        raw_final, raw_date = cell(row_no, 3), cell(row_no, 4)
        final = final_date = None
        if _is_blank(raw_final) != _is_blank(raw_date):
            if _is_blank(raw_date):
                raise _cell_error(
                    file,
                    ws,
                    f"E{row_no}",
                    f"manca la data della lettura finale scritta in D{row_no}. "
                    f"{_DATE_ADVICE}, oppure cancellare la lettura finale, e "
                    "ripetere la conversione.",
                )
            raise _cell_error(
                file,
                ws,
                f"D{row_no}",
                f"manca la lettura finale della data scritta in E{row_no}. "
                "Scrivere la lettura (solo il numero), oppure cancellare la "
                "data, e ripetere la conversione.",
            )
        if not _is_blank(raw_final):
            final = _number(file, ws, row_no, row, 3, float)
            final_date = _date(file, ws, f"E{row_no}", raw_date, "la data")
        meters.append(CentralMeterStart(description, initial, final, final_date))
    return ExerciseInfo(year, start_date), meters


def _amount(value: float) -> str:
    """A reading as an Italian number: 24000, 31,613."""
    return f"{value:.3f}".rstrip("0").rstrip(".").replace(".", ",")


def check_against_report(report: ParsedReport, static_data: StaticData) -> None:
    """Refuse static data whose start readings are above the final readings.

    Compares the final reading of every central heat meter (the manual one of
    the static data when given, otherwise the report's) and of every apartment
    heat, hot water and cold water meter with the matching start reading.
    Raises StaticDataError naming every meter that went backwards: it means the
    static data is of another exercise than the report. Equal readings are fine.
    """
    # (apartment, kind, line): central meters sort first, then apartment by apartment
    found: list[tuple[int, int, str]] = []

    def check(
        apartment: int,
        kind: int,
        name: str,
        unit: str,
        final: float,
        initial: float,
        note: str = "",
    ) -> None:
        if final < initial:
            found.append(
                (
                    apartment,
                    kind,
                    f"- {name}: lettura finale {_amount(final)} {unit}{note}, "
                    f"lettura iniziale {_amount(initial)} {unit}",
                )
            )

    reported = {m.description: m.heat_energy_kwh for m in report.central_meters}
    for meter in static_data.central_meters:
        name = f"contatore centrale «{meter.description}»"
        if meter.final_kwh is not None:
            check(
                0,
                0,
                name,
                "kWh",
                meter.final_kwh,
                meter.initial_kwh,
                " (dai dati statici)",
            )
        elif meter.description in reported:
            check(0, 0, name, "kWh", reported[meter.description], meter.initial_kwh)

    previous = {p.apartment_number: p for p in static_data.previous_readings}
    for ha in report.heat_allocators:
        if (prev := previous.get(ha.apartment_number)) is not None:
            apartment = ha.apartment_number
            name = f"«{ha.description}»"
            check(
                apartment,
                1,
                f"contacalorie {name}",
                "kWh",
                ha.heat_energy_kwh,
                prev.heat_kwh,
            )
            check(
                apartment,
                3,
                f"contatore dell'acqua fredda {name}",
                "m³",
                ha.aux1_volume_m3,
                prev.cold_water_m3,
            )
    for wm in report.water_meters:
        if (prev := previous.get(wm.apartment_number)) is not None:
            check(
                wm.apartment_number,
                2,
                f"contatore dell'acqua calda «{wm.description}»",
                "m³",
                wm.water_volume_m3,
                prev.water_m3,
            )
    lines = [line for _, _, line in sorted(found)]

    if lines:
        raise StaticDataError(
            "Alcune letture finali sono più basse delle letture iniziali dei "
            "dati statici, e non può essere:\n"
            + "\n".join(lines)
            + "\n\nControllare di aver scelto i dati statici dello stesso "
            "esercizio del report FC (le letture iniziali sono le letture finali "
            "dell'esercizio precedente) e ripetere la conversione."
        )
