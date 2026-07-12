"""
STEP 3 — Meesho Prelim View
Adds Expert + RSC + YoY sections. Reads RDContract from upstream.
Returns MPVContract for downstream error_margin, fk_prelim, client_prelim.
"""
from typing import Dict

from shared.anchors import SHEET_ANCHORS, find_contiguous_last_col
from shared.month_utils import scan_contiguous_headers
from shared.contracts import MPVContract, RDContract
from shared.formula_utils import col_num, col_str, copy_column_block, next_col
from shared.style_utils import (
    YELLOW_FILL, add_comment, copy_cell_style,
    get_row_values_across_cols, trailing_avg,
)

# ── Constants ─────────────────────────────────────────────────────────────────
MPV_MONTH_HDR_ROW = 6
MPV_ORDERS_ROW    = 9
MPV_TOTAL_GMV_ROW = 10
MPV_CANCEL_ROW    = 11
MPV_GTN_ROW       = 14

MPV_MONTH_REMAP: Dict[str, str] = {
    "U": "V",   "AR": "AS",  "AM": "AN",  "AW": "AX",
    "AL": "AM", "AK": "AL",  "AA": "AB",
    "I": "J",   "J": "K",
    "T": "U",   "AQ": "AR",
}


def run(
    wb,
    new_month: str,
    rd_contract: RDContract,
    threshold_pp: float,
    dry_run: bool,
) -> MPVContract:
    """
    Entry point for Step 3.
    Uses rd_contract.prev_col / gmv_col for cross-sheet Raw Data refs.
    Returns MPVContract for downstream modules.
    """
    ws = wb["Meesho Prelim View"]
    anchor = SHEET_ANCHORS["Meesho Prelim View"]

    expert_prev = find_contiguous_last_col(
        ws, anchor["header_row"], anchor["series_start"], new_month,
    )
    _, last_label = scan_contiguous_headers(
        ws, anchor["header_row"], anchor["series_start"],
    )
    expert_new = next_col(expert_prev)
    print(f"\n  [MPV] Expert series ends at {expert_prev} "
          f"({last_label!r}). New col: {expert_new}")

    col_map = dict(MPV_MONTH_REMAP)
    col_map[rd_contract["prev_col"]] = rd_contract["gmv_col"]

    expert_offset = col_num(expert_prev) - col_num("U")
    rsc_prev     = col_str(col_num("AM") + expert_offset)
    rsc_new      = next_col(rsc_prev)
    yoy_exp_prev = col_str(col_num("AR") + expert_offset)
    yoy_exp_new  = next_col(yoy_exp_prev)
    yoy_rsc_prev = col_str(col_num("AW") + expert_offset)
    yoy_rsc_new  = next_col(yoy_rsc_prev)

    sections = [
        (expert_prev,  expert_new,   None),
        (rsc_prev,     rsc_new,      expert_new),
        (yoy_exp_prev, yoy_exp_new,  None),
        (yoy_rsc_prev, yoy_rsc_new,  None),
    ]
    print(f"  [MPV] Sections: Expert={expert_prev}→{expert_new}, "
          f"RSC={rsc_prev}→{rsc_new}, "
          f"YoY={yoy_exp_prev}→{yoy_exp_new}/{yoy_rsc_prev}→{yoy_rsc_new}")

    total = 0
    for prev_col, new_col, mirror in sections:
        n = copy_column_block(
            ws, prev_col, new_col, 7, 45,
            col_map=col_map,
            header_row=MPV_MONTH_HDR_ROW,
            month_label=new_month,
            mirror_expert_col=mirror,
            hardcoded_rows=[MPV_CANCEL_ROW] if prev_col == expert_prev else [],
            dry_run=dry_run,
        )
        total += n
        print(f"  [MPV]   {prev_col}→{new_col}: {n} cells")

    contract = MPVContract(
        expert_prev_col=expert_prev,
        expert_col=expert_new,
        rsc_prev_col=rsc_prev,
        rsc_col=rsc_new,
        yoy_expert_prev_col=yoy_exp_prev,
        yoy_expert_col=yoy_exp_new,
        yoy_rsc_prev_col=yoy_rsc_prev,
        yoy_rsc_col=yoy_rsc_new,
        col_map=col_map,
    )

    if dry_run:
        print(f"  [DRY-RUN] Would write {total} cells across MPV sections.")
        return contract

    rd_gmv_col = rd_contract["gmv_col"]
    ws[f"{expert_new}{MPV_ORDERS_ROW}"].value   = f"='Raw Data'!{rd_gmv_col}65"
    ws[f"{expert_new}{MPV_TOTAL_GMV_ROW}"].value = f"='Raw Data'!{rd_gmv_col}25"
    copy_cell_style(ws[f"{expert_prev}{MPV_ORDERS_ROW}"],
                    ws[f"{expert_new}{MPV_ORDERS_ROW}"])
    copy_cell_style(ws[f"{expert_prev}{MPV_TOTAL_GMV_ROW}"],
                    ws[f"{expert_new}{MPV_TOTAL_GMV_ROW}"])

    lookback_cols = [col_str(col_num(expert_prev) - i) for i in range(3, -1, -1)]
    for row, label in [(MPV_CANCEL_ROW, "Cancellation %"), (MPV_GTN_ROW, "Gross-to-Net %")]:
        nums = get_row_values_across_cols(ws, row, lookback_cols)
        suggested = trailing_avg(nums, 3)
        if suggested is not None:
            c = ws[f"{expert_new}{row}"]
            c.value = round(suggested, 4)
            c.fill = YELLOW_FILL
            c.number_format = "0.00%"
            add_comment(ws, f"{expert_new}{row}",
                        f"TRENDLINE SUGGESTION for {label} "
                        f"(trailing-3m avg = {suggested:.2%}). Review before locking.")

    print(f"  [MPV] ✓ {total} cells. Expert={expert_new}, RSC={rsc_new}, "
          f"YoY={yoy_exp_new}/{yoy_rsc_new}. RD refs: "
          f"{rd_contract['prev_col']}→{rd_gmv_col}.")
    return contract
