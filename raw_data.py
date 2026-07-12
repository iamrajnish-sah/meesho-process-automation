"""
STEP 2 — Raw Data
Adds data + MoM + YoY columns, placed immediately to the right of the
last column in each section (same as dragging a column to the right in Excel).

MoM/YoY cells with a change beyond threshold_pp get a red font and a
"Please re-check again." note so an analyst can spot them at a glance.
Flags are computed from the real category numbers (Raw Working Sheet),
not from unevaluated Excel formulas, so they are correct without needing
Excel to recalculate first.
"""
from typing import Dict, Optional, Tuple

from raw_working_sheet import RWS_LOOKUP_RANGE
from shared.anchors import SHEET_ANCHORS, find_contiguous_last_col
from shared.contracts import RDContract, RWSResult
from shared.formula_utils import col_num, col_str, next_col
from shared.month_utils import scan_contiguous_headers, year_ago_label_for
from shared.style_utils import add_comment, copy_cell_style, mark_anomaly

RD_GMV_HDR_ROW   = 24
RD_TOTAL_GMV_ROW = 25
RD_ORD_HDR_ROW   = 64
RD_TOTAL_ORD_ROW = 65

RD_GMV_ROLLUP_ROWS = {
    25: "SUM({{c}}26,{{c}}42,{{c}}43,{{c}}46,{{c}}47,{{c}}50)",
    26: "{{c}}27+{{c}}34+{{c}}36+{{c}}37+{{c}}38+{{c}}39+{{c}}40+{{c}}41",
    27: "SUM({{c}}28:{{c}}33)",
    43: "{{c}}44+{{c}}45",
    47: "{{c}}48+{{c}}49",
    50: "SUM({{c}}51:{{c}}60)",
    60: "SUM({{c}}57:{{c}}59)",
}
RD_ORD_ROLLUP_ROWS = {
    65: "SUM({{c}}66,{{c}}82,{{c}}83,{{c}}86,{{c}}87,{{c}}90)",
    66: "{{c}}67+{{c}}74+{{c}}76+{{c}}77+{{c}}78+{{c}}79+{{c}}80+{{c}}81",
    67: "SUM({{c}}68:{{c}}73)",
    83: "{{c}}84+{{c}}85",
    87: "{{c}}88+{{c}}89",
    90: "SUM({{c}}91:{{c}}100)",
}


def _make_gmv_formula(row: int, new_col: str, rws_gmv_cr: str) -> str:
    if row in RD_GMV_ROLLUP_ROWS:
        return "=" + RD_GMV_ROLLUP_ROWS[row].replace("{{c}}", new_col)
    return (
        f"=_xlfn.XLOOKUP(A{row},"
        f"'Raw Working Sheet'!{RWS_LOOKUP_RANGE},"
        f"'Raw Working Sheet'!${rws_gmv_cr}$5:${rws_gmv_cr}$34,"
        f'\"\",0)'
    )


def _make_ord_formula(row: int, new_col: str, rws_ord_mn: str) -> str:
    if row in RD_ORD_ROLLUP_ROWS:
        return "=" + RD_ORD_ROLLUP_ROWS[row].replace("{{c}}", new_col)
    return (
        f"=_xlfn.XLOOKUP(A{row},"
        f"'Raw Working Sheet'!{RWS_LOOKUP_RANGE},"
        f"'Raw Working Sheet'!${rws_ord_mn}$5:${rws_ord_mn}$34,"
        f'\"\",0)'
    )


def _find_section_end_after_gap(ws, header_row: int, start_col_n: int) -> str:
    """
    Starting from start_col_n, skip any blank cells (the gap/spacer column),
    then walk forward and return the last column of the next contiguous occupied block.

    Example — for May'26:
      Data ends at S (col 19). Call with start=20 (T).
      T is blank → skip. U has content (=F24) → section starts here.
      Walk forward: U, V, W … AA all have content, AB is blank → section ends at AA.
      Returns "AA" → caller puts MoM in next_col("AA") = "AB". ✓
    """
    n = start_col_n
    # Skip blank separator column(s)
    while n < 500:
        v = ws.cell(header_row, n).value
        if v is not None and str(v).strip() != "":
            break
        n += 1
    # Walk the occupied block and record last non-blank column
    last_n = n
    while n < 500:
        v = ws.cell(header_row, n).value
        if v is None or str(v).strip() == "":
            break
        last_n = n
        n += 1
    return col_str(last_n)


def _fix_ref_errors(ws) -> None:
    for r in range(25, 61):
        cell = ws[f"Z{r}"]
        if cell.value and "#REF!" in str(cell.value):
            cell.value = f"=R{r}/Q{r}-1"
            print(f"  [RD] Fixed #REF! at Z{r}")


def _write_pct_cell(ws, col: str, row: int, formula: str, style_src_col: str) -> None:
    ws[f"{col}{row}"].value = formula
    copy_cell_style(ws[f"{style_src_col}{row}"], ws[f"{col}{row}"])
    ws[f"{col}{row}"].number_format = "0.0%"


def _flag_from_known_values(
    ws, col: str, row: int, new_v: Optional[float], old_v: Optional[float],
    threshold_pp: float,
) -> bool:
    """Compare two real numbers (not Excel formulas) and flag if change > threshold."""
    if new_v is None or old_v is None or old_v == 0:
        return False
    change = new_v / old_v - 1
    if abs(change) <= threshold_pp / 100.0:
        return False
    mark_anomaly(ws, f"{col}{row}")
    return True


def _sum_matched(by_name: Dict[str, Tuple[float, float]], idx: int) -> Optional[float]:
    if not by_name:
        return None
    return sum(v[idx] for v in by_name.values())


def run(
    wb,
    new_month: str,
    rws_contract: RWSResult,
    threshold_pp: float,
    dry_run: bool,
) -> RDContract:
    ws = wb["Raw Data"]
    anchor = SHEET_ANCHORS["Raw Data"]

    prev_col = find_contiguous_last_col(
        ws, anchor["header_row"], anchor["series_start"], new_month,
    )
    _, last_label = scan_contiguous_headers(
        ws, anchor["header_row"], anchor["series_start"],
    )

    new_col = next_col(prev_col)

    # MoM: skip the blank spacer column after the new data column, then find the
    # end of the existing MoM section and append right after it (e.g. AB for May'26).
    mom_section_end = _find_section_end_after_gap(
        ws, anchor["header_row"], col_num(next_col(new_col))
    )
    mom_new = next_col(mom_section_end)

    # YoY: skip the blank spacer column after the MoM section, then find the
    # end of the existing YoY section and append right after it (e.g. AH for May'26).
    yoy_section_end = _find_section_end_after_gap(
        ws, anchor["header_row"], col_num(next_col(mom_new))
    )
    yoy_new = next_col(yoy_section_end)

    yoy_label_target = year_ago_label_for(new_month)
    yoy_base_col: Optional[str] = None
    for c in range(col_num(anchor["series_start"]), col_num(prev_col) + 1):
        v = ws.cell(anchor["header_row"], c).value
        if v and yoy_label_target in str(v).strip():
            yoy_base_col = col_str(c)
            break

    print(
        f"  [RD] Last month in workbook: {last_label!r} (col {prev_col}). "
        f"Adding {new_month!r}. New={new_col}, MoM={mom_new}, YoY={yoy_new} "
        f"(base={yoy_base_col or 'NOT FOUND'}, flag threshold={threshold_pp}%)"
    )

    contract = RDContract(
        prev_col=prev_col,
        gmv_col=new_col,
        mom_col=mom_new,
        yoy_col=yoy_new if yoy_base_col else None,
        yoy_base_col=yoy_base_col,
    )

    if dry_run:
        return contract

    rws_ord_mn = rws_contract["ord_mn_col"] or "BN"
    rws_gmv_cr = rws_contract["gmv_cr_col"] or "BO"
    matched_by_name = rws_contract.get("matched_by_name", {})
    prev_by_name = rws_contract.get("prev_by_name", {})
    yoy_by_name = rws_contract.get("yoy_by_name", {})
    flagged = 0

    for hdr_row in (anchor["header_row"], anchor["header_row_orders"]):
        ws.cell(hdr_row, col_num(new_col)).value = new_month
        copy_cell_style(ws.cell(hdr_row, col_num(prev_col)),
                        ws.cell(hdr_row, col_num(new_col)))
        ws.cell(hdr_row, col_num(mom_new)).value = f"MoM ({new_month})"
        if yoy_base_col:
            ws.cell(hdr_row, col_num(yoy_new)).value = f"YoY ({new_month})"

    for r in range(RD_TOTAL_GMV_ROW, 61):
        ws[f"{new_col}{r}"].value = _make_gmv_formula(r, new_col, rws_gmv_cr)
        copy_cell_style(ws[f"{prev_col}{r}"], ws[f"{new_col}{r}"])
        _write_pct_cell(ws, mom_new, r, f'=IFERROR({new_col}{r}/{prev_col}{r}-1,"")', prev_col)
        if yoy_base_col:
            _write_pct_cell(ws, yoy_new, r, f'=IFERROR({new_col}{r}/{yoy_base_col}{r}-1,"")', prev_col)

        cat = ws[f"A{r}"].value
        key = str(cat).strip() if cat else None
        if key and key in matched_by_name:
            new_v = matched_by_name[key][1]
            old_v = prev_by_name.get(key, (None, None))[1]
            yoy_v = yoy_by_name.get(key, (None, None))[1]
        elif r == RD_TOTAL_GMV_ROW:
            new_v = _sum_matched(matched_by_name, 1)
            old_v = _sum_matched(prev_by_name, 1)
            yoy_v = _sum_matched(yoy_by_name, 1)
        else:
            new_v = old_v = yoy_v = None

        if _flag_from_known_values(ws, mom_new, r, new_v, old_v, threshold_pp):
            flagged += 1
        if yoy_base_col and _flag_from_known_values(ws, yoy_new, r, new_v, yoy_v, threshold_pp):
            flagged += 1

    for r in range(RD_TOTAL_ORD_ROW, 101):
        if ws[f"A{r}"].value is None and ws[f"B{r}"].value is None:
            break
        ws[f"{new_col}{r}"].value = _make_ord_formula(r, new_col, rws_ord_mn)
        copy_cell_style(ws[f"{prev_col}{r}"], ws[f"{new_col}{r}"])
        _write_pct_cell(ws, mom_new, r, f'=IFERROR({new_col}{r}/{prev_col}{r}-1,"")', prev_col)
        if yoy_base_col:
            _write_pct_cell(ws, yoy_new, r, f'=IFERROR({new_col}{r}/{yoy_base_col}{r}-1,"")', prev_col)

        cat = ws[f"A{r}"].value
        key = str(cat).strip() if cat else None
        if key and key in matched_by_name:
            new_v = matched_by_name[key][0]
            old_v = prev_by_name.get(key, (None, None))[0]
            yoy_v = yoy_by_name.get(key, (None, None))[0]
        elif r == RD_TOTAL_ORD_ROW:
            new_v = _sum_matched(matched_by_name, 0)
            old_v = _sum_matched(prev_by_name, 0)
            yoy_v = _sum_matched(yoy_by_name, 0)
        else:
            new_v = old_v = yoy_v = None

        if _flag_from_known_values(ws, mom_new, r, new_v, old_v, threshold_pp):
            flagged += 1
        if yoy_base_col and _flag_from_known_values(ws, yoy_new, r, new_v, yoy_v, threshold_pp):
            flagged += 1

    _fix_ref_errors(ws)
    if flagged:
        print(f"  [RD] Flagged {flagged} MoM/YoY cells (>{threshold_pp}% change) in red with a note.")
    print(f"  [RD] Done. data={new_col}, MoM={mom_new}, YoY={yoy_new}.")
    return contract

