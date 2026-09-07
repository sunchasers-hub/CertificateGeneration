"""
excel_reader.py
Reads the participants Excel file and yields cleaned, validated row dicts.
"""

import re
from dataclasses import dataclass
from datetime import datetime, date
from pathlib import Path
from typing import Iterator, Optional

from openpyxl import load_workbook

import config


@dataclass
class Participant:
    row_number: int          # 1-based Excel row, for error messages
    name: str
    distance: str
    time_value: float
    time_display: str        # original cell text/value formatted for display
    date_display: str
    email: str

    @property
    def has_time(self) -> bool:
        return self.time_value > 0


class ExcelValidationError(Exception):
    pass


def _to_float_safe(value) -> Optional[float]:
    """Best-effort conversion of a Time cell to a float for the >0 check.

    Accepts numeric cells, numeric strings, and Excel time/duration objects
    (datetime.time / timedelta) by falling back to a non-zero string check.
    """
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    # datetime.time or timedelta -> treat any non-midnight/non-zero value as > 0
    try:
        import datetime as _dt
        if isinstance(value, _dt.time):
            return 1.0 if (value.hour, value.minute, value.second) != (0, 0, 0) else 0.0
        if isinstance(value, _dt.timedelta):
            return value.total_seconds()
    except Exception:
        pass
    try:
        return float(str(value).strip())
    except ValueError:
        pass

    # Handle strings like "26 min", "53 minutes", "1:05:00", "25:30"
    text = str(value).strip().lower()

    # "hh:mm:ss" or "mm:ss" style durations
    if re.fullmatch(r"\d{1,2}(:\d{2}){1,2}", text):
        parts = [int(p) for p in text.split(":")]
        if len(parts) == 2:
            minutes, seconds = parts
            hours = 0
        else:
            hours, minutes, seconds = parts
        return hours * 3600 + minutes * 60 + seconds

    # "<number> min" / "<number> minutes" / "<number> sec" etc.
    match = re.search(r"[\d.]+", text)
    if match:
        return float(match.group())

    # Non-empty, non-numeric text with no extractable number (e.g. "DNF")
    return 0.0


def _format_display(value) -> str:
    """Render a cell value the way it should appear on the certificate / email."""
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%d-%b-%Y")
    return str(value).strip()


def read_participants(excel_path: Path = None) -> Iterator[Participant]:
    excel_path = excel_path or config.EXCEL_PATH
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    wb = load_workbook(excel_path, data_only=True)
    ws = wb.active

    header_row = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    header_index = {name: idx for idx, name in enumerate(header_row)}

    missing = [c for c in config.REQUIRED_COLUMNS if c not in header_index]
    if missing:
        raise ExcelValidationError(
            f"Excel is missing required column(s): {missing}. Found columns: {header_row}"
        )

    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        values = [cell.value for cell in row]
        if all(v is None for v in values):
            continue  # skip blank rows

        def get(col_name):
            return values[header_index[col_name]]

        name = _format_display(get(config.COL_NAME))
        email = _format_display(get(config.COL_EMAIL))
        if not name or not email:
            # Can't process a row with no name or no email address
            continue

        time_raw = get(config.COL_TIME)
        time_value = _to_float_safe(time_raw)

        yield Participant(
            row_number=row_idx,
            name=name,
            distance=_format_display(get(config.COL_DISTANCE)),
            time_value=time_value,
            time_display=_format_display(time_raw),
            date_display=_format_display(get(config.COL_DATE)),
            email=email,
        )
