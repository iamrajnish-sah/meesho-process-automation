"""
STEP 4 — Error Margin - Expert vs Report
Drags all parallel section end-columns forward.
Reads MPVContract from upstream (for reference; formulas shift via drag).
Returns EMContract listing new columns written.
"""
from shared.anchors import SHEET_ANCHORS, find_all_month_cols_in_row
from shared.month_utils import anchor_search_token
from shared.contracts import EMContract, MPVContract
from shared.formula_utils import drag_column_forward, next_col
from shared.style_utils import copy_cell_style

EM_DATA_ROW_START = 5
EM_DATA_ROW_END   = 110
EM_CHECKER_ROW    = 4


def run(wb, new_month: str, mpv_contract: MPVContract, dry_run: bool) -> EMContract:
    """Entry point for Step 4."""
    ws = wb["Error Margin - Expert vs Report"]
    anchor = SHEET_ANCHORS["Error Margin - Expert vs Report"]
    print(f"\n  [EM] Processing Error Margin sheet …")

    search = anchor_search_token(new_month)
    section_ends = find_all_month_cols_in_row(ws, anchor["header_row"], search)
    if not section_ends:
        print(f"  [EM] WARNING: No sections ending with {search!r}. Skipping.")
        return EMContract(section_new_cols=[])

    print(f"  [EM] Found {len(section_ends)} section(s): {section_ends}")
    new_cols: list = []
    for last_col in section_ends:
        new_col = next_col(last_col)
        new_cols.append(new_col)
        if dry_run:
            print(f"  [DRY-RUN] EM section: {last_col}→{new_col}")
            continue
        ws[f"{new_col}{anchor['header_row']}"].value = new_month
        copy_cell_style(ws[f"{last_col}{anchor['header_row']}"],
                        ws[f"{new_col}{anchor['header_row']}"])
        n = drag_column_forward(ws, last_col, new_col,
                                EM_DATA_ROW_START, EM_DATA_ROW_END, dry_run=dry_run)
        print(f"  [EM] ✓ Section {last_col}→{new_col}: {n} cells.")

    if not dry_run:
        print(f"  [EM] ⚠  Confirm row {EM_CHECKER_ROW} checkers = 0 in Excel.")

    return EMContract(section_new_cols=new_cols)
