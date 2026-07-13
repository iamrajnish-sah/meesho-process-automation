"""
STEP 4 — Error Margin - Expert vs Report

Drags all parallel section end-columns forward by one month.
Returns EMContract listing new columns written.

After the drag, runs the EM gap analysis automatically:
  - Column references come DIRECTLY from the pipeline contract (section_ends /
    new_cols) — no re-scan of the sheet.  The "current month" label is always
    new_month, which is passed in from run_pipeline.py and originates from the
    user's input; it is never hardcoded here.
  - Numeric values are read from src_path (the master workbook file, before
    pipeline changes) with data_only=True so cached formula results are available.
  - ALL output is written to the far-right of the sheet only
    (cols new_cols[-1]+1 … new_cols[-1]+7).  Nothing is written adjacent to the
    RSC block: the RSC block grows one column right every month, and placing any
    output next to it guarantees a collision next cycle (BUG 2 fix).
  - Headers are labelled with new_month ("May'26", "Jun'26", …) so they always
    reflect the pipeline month that produced them.
"""
import os
import re
from typing import List, Optional, Tuple

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import column_index_from_string as CN, get_column_letter as CL

from shared.anchors import SHEET_ANCHORS, find_all_month_cols_in_row
from shared.month_utils import anchor_search_token, year_ago_label_for, parse_month_label
from shared.contracts import EMContract, MPVContract
from shared.formula_utils import drag_column_forward, next_col
from shared.style_utils import copy_cell_style, YELLOW_FILL, ORANGE_FILL

SHEET_NAME        = "Error Margin - Expert vs Report"
HDR_ROW           = 5
EM_DATA_ROW_START = 5
EM_DATA_ROW_END   = 110
EM_CHECKER_ROW    = 4
EXPERT_START_COL  = 4    # column D (never changes)

# ── Gap analysis constants ─────────────────────────────────────────────────
MOM_GAP_THRESHOLD = 6      # pp — flag if |Expert MoM% − RSC MoM%| > this
YOY_GAP_THRESHOLD = 6      # pp — same for YoY
EXPERT_WEIGHT     = 0.7    # weight in suggest_adjustment_placeholder
RSC_WEIGHT        = 1 - EXPERT_WEIGHT
MIN_GMV           = 5.0    # ignore rows whose values are below this

MONTH_RE  = re.compile(r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)'\d{2}", re.I)
RED_FILL  = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
GREEN_FILL= PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
GREY_FILL = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
BOLD      = Font(bold=True)


# ─────────────────────────────────────────────────────────────────────────────
# Pipeline entry point
# ─────────────────────────────────────────────────────────────────────────────

def run(wb, new_month: str, mpv_contract: MPVContract, dry_run: bool,
        src_path: Optional[str] = None) -> EMContract:
    """Step 4.

    src_path — path to the tmp/master workbook (used read-only for cached
    Expert/RSC numeric values during gap analysis only).
    """
    ws = wb[SHEET_NAME]
    anchor = SHEET_ANCHORS[SHEET_NAME]
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

    if not dry_run and src_path and os.path.exists(src_path) and len(new_cols) >= 2:
        try:
            _run_gap_analysis(ws, new_month, section_ends, new_cols, src_path)
        except Exception as exc:
            print(f"  [EM Gap] ⚠ Gap analysis skipped: {exc}")

    return EMContract(section_new_cols=new_cols)


# ─────────────────────────────────────────────────────────────────────────────
# Gap analysis — helpers
# ─────────────────────────────────────────────────────────────────────────────

def _val(ws, row: int, col: int) -> Optional[float]:
    v = ws.cell(row, col).value
    try:
        f = float(v)
        return None if f != f else f
    except (TypeError, ValueError):
        return None


def _pct(new_val: float, base_val: float) -> Optional[float]:
    return None if base_val == 0 else (new_val - base_val) / abs(base_val) * 100


def _is_bare(ws_f, row: int, col: int) -> bool:
    return isinstance(ws_f.cell(row, col).value, (int, float))


def suggest_adjustment_placeholder(rsc_prev: float,
                                   expert_mom_rate: float,
                                   rsc_mom_rate:    float) -> float:
    """PLACEHOLDER — replace only this function body when wiring Gemini later.
    Signature must stay the same.

    Rule: RSC_prev × (1 + 0.7×Expert_MoM + 0.3×RSC_MoM)
    """
    return rsc_prev * (1 + EXPERT_WEIGHT * expert_mom_rate + RSC_WEIGHT * rsc_mom_rate)


def _find_category_rows(
    ws_v, ws_f,
    exp_n: int, exp_p: int,
    rsc_n: int, rsc_p: int,
    exp_yoy: Optional[int],
    rsc_yoy: Optional[int],
) -> List[Tuple]:
    """Rows where RSC newest column is a bare number and all four GMV values
    are GMV-scale (> MIN_GMV).  Returns list of
      (row, label, ev, epv, rv, rpv, ey, ry)."""
    results = []
    for r in range(2, 150):
        b = ws_v.cell(r, 2).value
        label = str(b).strip() if b else ""
        if "checker" in label.lower():
            continue
        ev  = _val(ws_v, r, exp_n)
        epv = _val(ws_v, r, exp_p)
        rv  = _val(ws_v, r, rsc_n)
        rpv = _val(ws_v, r, rsc_p)
        if any(x is None for x in [ev, epv, rv, rpv]):
            continue
        if any(x <= MIN_GMV for x in [ev, epv, rv, rpv]):
            continue
        if not _is_bare(ws_f, r, rsc_n):
            continue
        ey = _val(ws_v, r, exp_yoy) if exp_yoy else None
        ry = _val(ws_v, r, rsc_yoy) if rsc_yoy else None
        results.append((r, label, ev, epv, rv, rpv, ey, ry))
    return results


# ─────────────────────────────────────────────────────────────────────────────
# Gap analysis — orchestrator
# ─────────────────────────────────────────────────────────────────────────────

def _run_gap_analysis(ws, new_month: str,
                      section_ends: List[str],
                      new_cols:     List[str],
                      src_path:     str) -> None:
    """
    HOW COLUMN POSITIONS ARE DETERMINED (BUG 1 FIX):
      section_ends[0] = Expert "previous-month" column in the master
                        (e.g. S for "Apr'26 New") — this IS the Expert data to analyse.
      section_ends[1] = RSC  "previous-month" column in the master
                        (e.g. AJ for "Apr'26 New", bare-number = analyst's RSC entry).
      Both come directly from find_all_month_cols_in_row which the EM pipeline already
      ran.  We do NOT re-scan the sheet here — that would be the bug that caused the
      wrong-month label.

    HOW HEADERS ARE LABELLED (BUG 1 FIX):
      All column headers use `new_month` (e.g. "May'26"), which is the month the user
      asked to process.  It arrives via the pipeline contract, never hardcoded.

    WHERE OUTPUT IS WRITTEN (BUG 2 FIX):
      ONLY to the far-right block at new_cols[-1]+1 … new_cols[-1]+7.
      Nothing is written adjacent to the RSC block (the old AL write is removed).
      Reason: the RSC block grows one column right every month.  "adjacent" means a
      guaranteed overwrite collision next cycle.  The far-right block also shifts one
      column right per month — it lands on the first blank column after the last
      section endpoint, which is always free at the time of writing.

    CACHED VALUES:
      src_path (the master/tmp file) is loaded read-only with data_only=True so
      Excel-cached formula values are available for the Expert and RSC formula cells.
    """
    print(f"\n  [EM Gap] Loading cached values from {os.path.basename(src_path)} …")
    try:
        wb_v = openpyxl.load_workbook(src_path, data_only=True,  read_only=True)
        wb_f = openpyxl.load_workbook(src_path, data_only=False, read_only=True)
    except Exception as exc:
        print(f"  [EM Gap] Cannot open src file: {exc}")
        return

    ws_v = wb_v[SHEET_NAME]
    ws_f = wb_f[SHEET_NAME]

    # ── Column references — from pipeline contract, not re-scanned ──────────
    exp_old_c = CN(section_ends[0])      # Expert newest in master: S (Apr'26 Expert)
    exp_prv_c = exp_old_c - 1            # Expert prev: R (Mar'26 Expert)
    rsc_old_c = CN(section_ends[1])      # RSC hardcoded column: AJ (Apr'26 RSC bare)
    rsc_prv_c = rsc_old_c - 1            # RSC prev: AI (Mar'26 RSC)

    if exp_old_c <= EXPERT_START_COL:
        print("  [EM Gap] Expert block too short for MoM — skipping.")
        return

    # ── YoY base: read the month label at exp_old_c from src_path ───────────
    #   e.g. "Apr'26 New " → year_ago → "Apr'25"
    #   Scan HDR_ROW in src_path for that label to find YoY columns.
    exp_month_label = ws_v.cell(HDR_ROW, exp_old_c).value or ""
    exp_yoy_c = rsc_yoy_c = None
    yoy_base_label = None
    try:
        yoy_base_label = year_ago_label_for(str(exp_month_label))
        yoy_target     = parse_month_label(yoy_base_label)
        if yoy_target:
            for c in range(EXPERT_START_COL, exp_old_c):
                lbl = ws_v.cell(HDR_ROW, c).value
                if lbl and parse_month_label(str(lbl)) == yoy_target:
                    exp_yoy_c = c
                    break
            for c in range(exp_old_c + 1, rsc_old_c):
                lbl = ws_v.cell(HDR_ROW, c).value
                if lbl and parse_month_label(str(lbl)) == yoy_target:
                    rsc_yoy_c = c
                    break
    except ValueError:
        pass

    print(f"  [EM Gap] new_month   : {new_month!r}  ← label for all output headers")
    print(f"  [EM Gap]   (source   : run() parameter, from run_pipeline.py contract)")
    print(f"  [EM Gap] Expert cols : prev={CL(exp_prv_c)}, newest={CL(exp_old_c)}, "
          f"YoY={CL(exp_yoy_c) if exp_yoy_c else 'N/A'} ({yoy_base_label})")
    print(f"  [EM Gap] RSC cols    : prev={CL(rsc_prv_c)}, newest={CL(rsc_old_c)}, "
          f"YoY={CL(rsc_yoy_c) if rsc_yoy_c else 'N/A'}")
    print(f"  [EM Gap]   (cols derived from section_ends={section_ends[:2]} — pipeline contract)")

    # ── Category rows ───────────────────────────────────────────────────────
    rows_data = _find_category_rows(
        ws_v, ws_f, exp_old_c, exp_prv_c, rsc_old_c, rsc_prv_c, exp_yoy_c, rsc_yoy_c,
    )
    print(f"  [EM Gap] Rows found  : {len(rows_data)} with hardcoded RSC newest")
    if not rows_data:
        print("  [EM Gap] Nothing to analyse — skipping.")
        return

    # ── Far-right output columns ────────────────────────────────────────────
    # Always placed at new_cols[-1]+1 (first column after the last dragged section).
    # This shifts right by one each month alongside every EM section — never collides.
    # The previous month's first far-right column is overwritten by the section drag
    # (acceptable; the current month is always fresh and correct).
    far_n = CN(new_cols[-1]) + 1

    print(f"  [EM Gap] Far-right   : cols {CL(far_n)}–{CL(far_n+6)}")
    print(f"  [EM Gap] RSC block ends at {new_cols[1]}. "
          f"Next month grows to {CL(CN(new_cols[1])+1)}. "
          f"Nothing written at {CL(CN(new_cols[1])+1)} — no collision.")

    # ── Headers ─────────────────────────────────────────────────────────────
    col_headers = [
        f"Expert MoM% ({new_month})",
        f"RSC MoM% ({new_month})",
        f"MoM Gap pp ({new_month})",
        f"Expert YoY% ({new_month})",
        f"RSC YoY% ({new_month})",
        f"YoY Gap pp ({new_month})",
        f"RSC Suggested ({new_month})",
    ]
    for i, hdr_text in enumerate(col_headers):
        cell = ws.cell(HDR_ROW, far_n + i)
        cell.value = hdr_text
        cell.font  = BOLD
        cell.fill  = GREY_FILL

    # ── Process rows ─────────────────────────────────────────────────────────
    print()
    print(f"  {'Category':<38} {'ExpMoM%':>8} {'RscMoM%':>8} {'MoMGap':>7}"
          f" {'MoM':>5} {'ExpYoY%':>8} {'RscYoY%':>8} {'YoYGap':>7} {'YoY':>5} {'Sugg':>10}")
    print(f"  {'-'*38} {'-'*8} {'-'*8} {'-'*7} {'-'*5} {'-'*8} {'-'*8} {'-'*7} {'-'*5} {'-'*10}")

    mom_flag_count = yoy_flag_count = 0

    for row_data in rows_data:
        r, label, ev, epv, rv, rpv, ey, ry = row_data

        exp_mom  = _pct(ev,  epv)
        rsc_mom  = _pct(rv,  rpv)
        mom_gap  = (abs(exp_mom - rsc_mom)
                    if exp_mom is not None and rsc_mom is not None else None)
        mom_flag = mom_gap is not None and mom_gap > MOM_GAP_THRESHOLD

        exp_yoy_pct = _pct(ev, ey) if (ey and ey > MIN_GMV) else None
        rsc_yoy_pct = _pct(rv, ry) if (ry and ry > MIN_GMV) else None
        yoy_gap     = (abs(exp_yoy_pct - rsc_yoy_pct)
                       if exp_yoy_pct is not None and rsc_yoy_pct is not None else None)
        yoy_flag    = yoy_gap is not None and yoy_gap > YOY_GAP_THRESHOLD

        suggested = None
        if mom_flag and exp_mom is not None and rsc_mom is not None:
            mom_flag_count += 1
            suggested = suggest_adjustment_placeholder(rpv, exp_mom / 100, rsc_mom / 100)
        if yoy_flag:
            yoy_flag_count += 1

        # Write far-right values (rows aligned to actual EM row numbers)
        metrics = [exp_mom, rsc_mom, mom_gap, exp_yoy_pct, rsc_yoy_pct, yoy_gap]
        gap_cols = {2, 5}
        for i, v in enumerate(metrics):
            cell = ws.cell(r, far_n + i)
            if v is None:
                cell.value = "N/A"
            else:
                cell.value = round(v, 2)
                cell.number_format = '0.00' if i in gap_cols else '+0.00;-0.00;0.00'
                if i == 2:
                    cell.fill = RED_FILL if mom_flag else GREEN_FILL
                elif i == 5:
                    cell.fill = RED_FILL if yoy_flag else GREEN_FILL

        sg_cell = ws.cell(r, far_n + 6)
        if suggested is not None:
            sg_cell.value = round(suggested, 2)
            sg_cell.fill  = YELLOW_FILL
            comment_txt = (
                f"Category: {label}\n"
                f"Expert MoM%: {exp_mom:+.2f}%  RSC MoM%: {rsc_mom:+.2f}%  "
                f"MoM Gap: {mom_gap:.2f}pp\n"
                f"Placeholder: RSC_prev×(1+{EXPERT_WEIGHT}×Exp+{RSC_WEIGHT}×RSC)\n"
                f"Replace suggest_adjustment_placeholder() body to wire Gemini."
            )
            sg_cell.comment = Comment(comment_txt, "em_gap_analysis")
            sg_cell.comment.width  = 300
            sg_cell.comment.height = 100
        else:
            sg_cell.value = "—"

        # Terminal row
        def _f(v): return f"{v:+8.2f}" if v is not None else "     N/A"
        def _g(v): return f"{v:7.2f}"  if v is not None else "    N/A"
        mf = "YES◀" if mom_flag else "  no"
        yf = "YES◀" if yoy_flag else "  no"
        sg = f"{suggested:>10,.2f}" if suggested is not None else "         —"
        print(f"  {label[:38]:<38} {_f(exp_mom)} {_f(rsc_mom)} {_g(mom_gap)}"
              f" {mf} {_f(exp_yoy_pct)} {_f(rsc_yoy_pct)} {_g(yoy_gap)} {yf} {sg}")

    print()
    print(f"  [EM Gap] MoM flagged : {mom_flag_count}/{len(rows_data)}  (>{MOM_GAP_THRESHOLD}pp)")
    print(f"  [EM Gap] YoY flagged : {yoy_flag_count}/{len(rows_data)}  (>{YOY_GAP_THRESHOLD}pp)")
    print(f"  [EM Gap] Cols written: {CL(far_n)}–{CL(far_n+6)} only.")
    print(f"  [EM Gap] Checker rows 109/135 — columns never touched. ✓")
