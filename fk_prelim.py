"""
STEP 5 — FK Prelim View
Drags all parallel section end-columns forward.
Reads MPVContract from upstream for #REF! annotation context.
Returns FKContract listing new columns written.
"""
from openpyxl.comments import Comment as XLComment

from shared.anchors import SHEET_ANCHORS, find_all_month_cols_in_row
from shared.month_utils import anchor_search_token
from shared.contracts import FKContract, MPVContract
from shared.formula_utils import drag_column_forward, next_col
from shared.style_utils import copy_cell_style

FK_MONTH_HDR_ROW  = 6
FK_DATA_ROW_START = 7
FK_DATA_ROW_END   = 35


def _fix_fk_ref_errors(ws, mpv_expert_col: str) -> None:
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and "#REF!" in cell.value:
                try:
                    old = cell.comment.text if cell.comment else ""
                    cell.comment = XLComment(
                        f"⚠ #REF! needs manual fix. Original: {cell.value}. {old}",
                        "MeeshoBot",
                    )
                except Exception:
                    pass


def run(wb, new_month: str, mpv_contract: MPVContract, dry_run: bool) -> FKContract:
    """Entry point for Step 5."""
    ws = wb["FK Prelim View"]
    anchor = SHEET_ANCHORS["FK Prelim View"]
    print(f"\n  [FK] Processing FK Prelim View …")

    search = anchor_search_token(new_month)
    section_ends = find_all_month_cols_in_row(ws, anchor["header_row"], search)
    if not section_ends:
        print(f"  [FK] WARNING: No sections ending with {search!r}. Skipping.")
        return FKContract(section_new_cols=[])

    print(f"  [FK] Found {len(section_ends)} section(s): {section_ends}")
    new_cols: list = []
    for last_col in section_ends:
        new_col = next_col(last_col)
        new_cols.append(new_col)
        if dry_run:
            print(f"  [DRY-RUN] FK section: {last_col}→{new_col}")
            continue
        ws[f"{new_col}{FK_MONTH_HDR_ROW}"].value = new_month
        copy_cell_style(ws[f"{last_col}{FK_MONTH_HDR_ROW}"],
                        ws[f"{new_col}{FK_MONTH_HDR_ROW}"])
        n = drag_column_forward(ws, last_col, new_col,
                                FK_DATA_ROW_START, FK_DATA_ROW_END, dry_run=dry_run)
        print(f"  [FK] ✓ Section {last_col}→{new_col}: {n} cells.")

    if not dry_run:
        _fix_fk_ref_errors(ws, mpv_contract["expert_col"])

    return FKContract(section_new_cols=new_cols)
