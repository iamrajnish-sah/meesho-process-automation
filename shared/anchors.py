"""Per-sheet anchor configuration and contiguous-series validation."""
import re
from typing import Dict, List, Optional, Tuple

from shared.formula_utils import col_num, col_str, next_col

# Per-sheet structural anchors (no month labels — those are detected from the workbook).
SHEET_ANCHORS: Dict[str, Dict] = {
    "Raw Working Sheet": {
        "mode": "append_right",
    },
    "Raw Data": {
        "mode": "contiguous",
        "series_start": "C",
        "header_row": 24,
        "header_row_orders": 64,
    },
    "Meesho Prelim View": {
        "mode": "contiguous",
        "series_start": "F",
        "header_row": 6,
    },
    "Error Margin - Expert vs Report": {
        "mode": "multi_section",
        "header_row": 5,
    },
    "FK Prelim View": {
        "mode": "multi_section",
        "header_row": 6,
    },
    "Client Prelim": {
        "mode": "contiguous_sparse",
        "header_row": 4,
    },
}


def find_contiguous_last_col(
    ws, header_row: int, series_start: str, new_month: str,
) -> str:
    """
    Walk right from series_start, stop at first blank header.
    Validates last month is the calendar month before new_month.
    """
    from shared.month_utils import scan_contiguous_headers, validate_new_month_against_last

    last_col, last_label = scan_contiguous_headers(ws, header_row, series_start)
    validate_new_month_against_last(
        last_label, new_month, sheet=ws.title, header_row=header_row,
    )
    return last_col


def find_next_blank_col(ws, header_row: int, start_col: str) -> str:
    """First blank column at header_row, scanning right from start_col."""
    n = col_num(start_col)
    while n < 500:
        v = ws.cell(header_row, n).value
        if v is None or str(v).strip() == "":
            return col_str(n)
        n += 1
    return col_str(n)


def find_all_month_cols_in_row(ws, row: int, search_label: str,
                                start_col_n: int = 1) -> List[str]:
    """All column letters in a row whose value contains search_label."""
    found = []
    for cell in ws[row]:
        if cell.column < start_col_n:
            continue
        v = cell.value
        if v and isinstance(v, str) and search_label in v:
            found.append(cell.column_letter)
    return found


def find_last_month_col_in_row(ws, row: int, skip_words=None) -> Optional[Tuple[str, str]]:
    """Return (col_letter, month_label) for rightmost month header in a row."""
    skip_words = skip_words or ("MoM", "YoY", "Deviation", "FORMULA", "=")
    last: Optional[Tuple[str, str]] = None
    for c in range(1, 200):
        cl = col_str(c)
        v = ws[f"{cl}{row}"].value
        if v is None:
            continue
        s = str(v).strip()
        if not s or any(w in s for w in skip_words):
            continue
        if "'" in s or re.search(r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)", s, re.I):
            last = (cl, s)
    return last
