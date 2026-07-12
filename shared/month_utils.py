"""Parse and validate month labels across workbook sheets."""
import re
from typing import List, Optional, Tuple

_MONTH_RE = re.compile(
    r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*'?\s*'?(\d{2})",
    re.IGNORECASE,
)

_MONTH_TO_NUM = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}

_NUM_TO_MONTH = {v: k.title() for k, v in _MONTH_TO_NUM.items()}


class MonthSequenceError(Exception):
    """Raised when workbook month order does not match the requested new month."""


def parse_month_label(text: str) -> Optional[Tuple[int, int]]:
    """
    Extract (month_number, full_year) from labels like:
      May'26, Apr'26 New, MoM (May'26), Apr'25 New (Anuj)
    """
    if not text:
        return None
    m = _MONTH_RE.search(str(text).strip())
    if not m:
        return None
    mon = _MONTH_TO_NUM.get(m.group(1).lower()[:3])
    if not mon:
        return None
    yr = int(m.group(2))
    year = 2000 + yr if yr < 100 else yr
    return mon, year


def format_month_label(month: int, year: int) -> str:
    return f"{_NUM_TO_MONTH[month]}'{year % 100:02d}"


def previous_month(month: int, year: int) -> Tuple[int, int]:
    if month == 1:
        return 12, year - 1
    return month - 1, year


def next_month(month: int, year: int) -> Tuple[int, int]:
    if month == 12:
        return 1, year + 1
    return month + 1, year


def prev_month_label_for(new_month: str) -> str:
    parsed = parse_month_label(new_month)
    if not parsed:
        raise ValueError(f"Could not parse new month label: {new_month!r}")
    pm, py = previous_month(*parsed)
    return format_month_label(pm, py)


def year_ago_label_for(new_month: str) -> str:
    """Same calendar month, one year earlier. May'26 -> May'25."""
    parsed = parse_month_label(new_month)
    if not parsed:
        raise ValueError(f"Could not parse new month label: {new_month!r}")
    mon, year = parsed
    return format_month_label(mon, year - 1)


def next_month_label_after(last_month_label: str) -> Optional[str]:
    parsed = parse_month_label(last_month_label)
    if not parsed:
        return None
    nm, ny = next_month(*parsed)
    return format_month_label(nm, ny)


def month_labels_same_month(found: str, target_month: int, target_year: int) -> bool:
    parsed = parse_month_label(found)
    return parsed == (target_month, target_year) if parsed else False


def anchor_search_token(new_month: str) -> str:
    """
    Substring to locate prior-month section ends in multi-section sheets.
    For new_month Jun'26 → returns May'26.
    """
    return prev_month_label_for(new_month)


def validate_new_month_against_last(
    last_label: str,
    new_month: str,
    *,
    sheet: str = "",
    header_row: int = 0,
) -> None:
    """
    Ensure the workbook's last contiguous month is the calendar month
    immediately before new_month, and new_month is not already present.
    Raises MonthSequenceError on failure.
    """
    new_parsed = parse_month_label(new_month)
    if not new_parsed:
        raise MonthSequenceError(
            f"Invalid new month label: {new_month!r}. Use format like May'26."
        )

    last_parsed = parse_month_label(last_label)
    if not last_parsed:
        raise MonthSequenceError(
            f"Could not read a month from the last header {last_label!r} "
            f"in sheet {sheet!r} row {header_row}."
        )

    expected_prev = previous_month(*new_parsed)
    if last_parsed != expected_prev:
        expected_label = format_month_label(*expected_prev)
        raise MonthSequenceError(
            f"Month sequence mismatch in {sheet!r} (row {header_row}).\n"
            f"  You asked to add:     {new_month!r}\n"
            f"  Workbook last month:  {last_label!r}\n"
            f"  Expected last month:  {expected_label!r}\n"
            f"  Upload the workbook that already has {expected_label!r} "
            f"as the latest month, then run again."
        )

    if month_labels_same_month(last_label, *new_parsed):
        raise MonthSequenceError(
            f"{new_month!r} already appears to be the latest month in {sheet!r}. "
            f"Use last month's output file as your master and pick the next month."
        )


def scan_contiguous_headers(
    ws, header_row: int, series_start: str,
) -> Tuple[str, str]:
    """
    Last month in the main data series (stops at first blank column).
    Also stops at MoM/YoY headers or formula-based headers (=…).
    """
    from shared.formula_utils import col_num, col_str

    skip_words = ("MoM", "YoY")
    start_n = col_num(series_start)
    last_col, last_label = series_start, ""
    for n in range(start_n, start_n + 400):
        v = ws.cell(header_row, n).value
        if v is None or str(v).strip() == "":
            break
        s = str(v).strip()
        if s.startswith("=") or any(w in s for w in skip_words):
            break
        if parse_month_label(s):
            last_col = col_str(n)
            last_label = s
    return last_col, last_label


def detect_suggested_new_month(wb) -> Optional[str]:
    """
    Read Raw Data row 24 and suggest the next month to process.
    Used by the web UI after master upload.
    """
    from shared.anchors import SHEET_ANCHORS

    if "Raw Data" not in wb.sheetnames:
        return None
    anchor = SHEET_ANCHORS["Raw Data"]
    ws = wb["Raw Data"]
    _, last_label = scan_contiguous_headers(
        ws, anchor["header_row"], anchor["series_start"],
    )
    return next_month_label_after(last_label)
