"""
STEP 6 — Client Prelim
Copies prior month block (+6 cols) for top summary, Expert Data, Client RSC.
Reads MPVContract.col_map from upstream for cross-sheet formula remapping.
Returns CPContract listing new block columns.
"""
from shared.anchors import find_last_month_col_in_row
from shared.contracts import CPContract, MPVContract
from shared.formula_utils import next_col, remap_formula_refs, shift_formula
from shared.month_utils import validate_new_month_against_last
from shared.style_utils import copy_cell_style

CP_MONTH_HDR_ROW      = 4
CP_BLOCK_COL_DELTA    = 6
CP_APR_ABS_COL        = "AZ"
CP_APR_YOY_COL        = "BA"
CP_APR_EXPERT_YOY_COL = "BB"
CP_APR_CLIENT_MOM_COL = "BC"
CP_APR_MOM_REF_COL    = "AW"


def _find_cp_last_block(ws, new_month: str):
    """
    Locate the rightmost block in Client Prelim row 4 and validate it is the
    expected previous month. Raises MonthSequenceError if the anchor looks wrong.
    Uses find_last_month_col_in_row (full-row scan) because CP blocks are sparse
    (+6-col jumps with non-month values in between), which defeats the contiguous
    scanner used by Raw Data and Meesho Prelim View. The validation step provides
    the same safety guarantee: a wrong anchor raises an error instead of silently
    writing to the wrong column.
    """
    last = find_last_month_col_in_row(ws, CP_MONTH_HDR_ROW)
    if last:
        abs_col, last_label = last
        validate_new_month_against_last(
            last_label, new_month,
            sheet="Client Prelim", header_row=CP_MONTH_HDR_ROW,
        )
        return abs_col, next_col(abs_col), next_col(abs_col, 2), next_col(abs_col, 3)
    return CP_APR_ABS_COL, CP_APR_YOY_COL, CP_APR_EXPERT_YOY_COL, CP_APR_CLIENT_MOM_COL


def _copy_cp_cell(ws, src_col, dst_col, row, col_map, dry_run):
    src = ws[f"{src_col}{row}"]
    if src.value is None:
        return False
    if dry_run:
        return True
    dst = ws[f"{dst_col}{row}"]
    val = src.value
    if isinstance(val, str) and val.startswith("="):
        if "'Meesho Prelim View'!" in val or "'Raw Data'!" in val:
            dst.value = remap_formula_refs(val, col_map, CP_BLOCK_COL_DELTA)
        else:
            dst.value = shift_formula(val, CP_BLOCK_COL_DELTA)
    else:
        dst.value = val
    copy_cell_style(src, dst)
    return True


def run(wb, new_month: str, mpv_contract: MPVContract, dry_run: bool) -> CPContract:
    """Entry point for Step 6. Uses mpv_contract.col_map for formula remapping."""
    ws = wb["Client Prelim"]
    col_map = mpv_contract["col_map"]
    print(f"\n  [CP] Adding '{new_month}' block to Client Prelim …")

    abs_col, yoy_col, exp_yoy_col, client_mom_col = _find_cp_last_block(ws, new_month)
    new_abs        = next_col(abs_col,        CP_BLOCK_COL_DELTA)
    new_yoy        = next_col(yoy_col,        CP_BLOCK_COL_DELTA)
    new_exp_yoy    = next_col(exp_yoy_col,    CP_BLOCK_COL_DELTA)
    new_client_mom = next_col(client_mom_col, CP_BLOCK_COL_DELTA)
    label_col      = next_col(abs_col, -1)
    new_label      = next_col(new_abs, -1)

    print(f"  [CP] Block {abs_col}/{yoy_col} → {new_abs}/{new_yoy} "
          f"(+{CP_BLOCK_COL_DELTA} cols)")

    count = 0
    if not dry_run:
        ws[f"{new_abs}{CP_MONTH_HDR_ROW}"].value = new_month
        copy_cell_style(ws[f"{abs_col}{CP_MONTH_HDR_ROW}"], ws[f"{new_abs}{CP_MONTH_HDR_ROW}"])
        ws[f"{new_abs}{5}"].value = ws[f"{abs_col}{5}"].value
        ws[f"{new_yoy}{5}"].value = ws[f"{yoy_col}{5}"].value
        copy_cell_style(ws[f"{abs_col}{5}"], ws[f"{new_abs}{5}"])
        copy_cell_style(ws[f"{yoy_col}{5}"], ws[f"{new_yoy}{5}"])

    for r in range(6, 17):
        if _copy_cp_cell(ws, abs_col, new_abs, r, col_map, dry_run): count += 1
        if _copy_cp_cell(ws, yoy_col, new_yoy, r, col_map, dry_run): count += 1

    if not dry_run:
        ws[f"{new_label}{20}"].value = new_month
        copy_cell_style(ws[f"{label_col}{20}"], ws[f"{new_label}{20}"])
        ws[f"{new_label}{28}"].value = new_month
        copy_cell_style(ws[f"{label_col}{28}"], ws[f"{new_label}{28}"])

    for r in range(21, 26):
        for prev, new in [(abs_col, new_abs), (yoy_col, new_yoy), (exp_yoy_col, new_exp_yoy)]:
            if _copy_cp_cell(ws, prev, new, r, col_map, dry_run): count += 1

    for r in range(29, 36):
        for prev, new in [
            (abs_col, new_abs), (yoy_col, new_yoy),
            (exp_yoy_col, new_exp_yoy), (client_mom_col, new_client_mom),
        ]:
            if _copy_cp_cell(ws, prev, new, r, col_map, dry_run): count += 1

    for r in range(30, 35):
        src = ws[f"{CP_APR_MOM_REF_COL}{r}"]
        if src.value and isinstance(src.value, str) and src.value.startswith("="):
            if not dry_run:
                ws[f"{CP_APR_MOM_REF_COL}{r}"].value = remap_formula_refs(
                    src.value, col_map, 0
                )
            count += 1

    if dry_run:
        print(f"  [DRY-RUN] Would write ~{count} Client Prelim cells.")
    else:
        print(f"  [CP] ✓ Wrote {count} cells. May block at {new_abs}–{new_client_mom}.")

    return CPContract(
        abs_col=new_abs,
        yoy_col=new_yoy,
        expert_yoy_col=new_exp_yoy,
        client_mom_col=new_client_mom,
    )
