"""
STEP 1 — Raw Working Sheet
Adds 4-column block (Orders abs, GMV abs, Orders Mn, GMV INR Cr) per new month.
Returns RWSResult contract for downstream raw_data module.
"""
from typing import Dict, List, Optional, Tuple

from shared.contracts import RWSResult
from shared.month_utils import (
    MonthSequenceError,
    month_labels_same_month,
    parse_month_label,
    year_ago_label_for,
)
from shared.formula_utils import col_num, col_str, next_col
from shared.raw_input import normalize_category_name
from shared.style_utils import ORANGE_FILL, copy_cell_style

# ── Constants ─────────────────────────────────────────────────────────────────
RWS_CAT_FIRST_ROW  = 5
RWS_CAT_LAST_ROW   = 34
RWS_TOTAL_GMV_ROW  = 35
RWS_MONTH_HDR_ROW  = 2
RWS_SUBLABEL_ROW   = 3
RWS_COL_HDR_ROW    = 4
RWS_MAP_COL_EXPERT = "BV"
RWS_MAP_COL_RSC    = "BW"
RWS_MAP_FIRST_ROW  = 5
RWS_MAP_LAST_ROW   = 41
RWS_LOOKUP_RANGE   = "$C$5:$C$34"   # referenced by raw_data XLOOKUP formulas

# Column C is itself a formula: =XLOOKUP(B{r}, $BQ$5:$BQ$41, $BR$5:$BR$41, "", 0)
# openpyxl can't evaluate formulas, so we emulate this lookup in Python to get
# the real canonical category name (the same text Raw Data's own column A uses).
RWS_C_LOOKUP_SRC_COL = "BQ"
RWS_C_LOOKUP_DST_COL = "BR"
RWS_C_LOOKUP_FIRST_ROW = 5
RWS_C_LOOKUP_LAST_ROW = 41


def build_canonical_name_map(ws) -> Dict[str, str]:
    """Emulate the BQ->BR XLOOKUP table used by column C."""
    mapping: Dict[str, str] = {}
    for r in range(RWS_C_LOOKUP_FIRST_ROW, RWS_C_LOOKUP_LAST_ROW + 1):
        src = ws[f"{RWS_C_LOOKUP_SRC_COL}{r}"].value
        dst = ws[f"{RWS_C_LOOKUP_DST_COL}{r}"].value
        if src and dst:
            mapping[str(src).strip()] = str(dst).strip()
    return mapping


def resolve_canonical_name(ws, row: int, name_map: Dict[str, str]) -> str:
    """What column C{row} would evaluate to, without needing Excel to calculate it."""
    b_val = ws[f"B{row}"].value
    if not b_val:
        return ""
    return name_map.get(str(b_val).strip(), "")


def build_category_row_lookup(ws):
    lookup: Dict[str, int] = {}
    row_aliases: Dict[int, List[str]] = {}

    def register(row: int, alias: str, overwrite: bool = True) -> None:
        if not alias:
            return
        norm = normalize_category_name(alias)
        if not norm:
            return
        if not overwrite and norm in lookup:
            return
        lookup[norm] = row
        row_aliases.setdefault(row, [])
        if alias not in row_aliases[row]:
            row_aliases[row].append(alias)

    for r in range(RWS_CAT_FIRST_ROW, RWS_CAT_LAST_ROW + 1):
        b_val = ws[f"B{r}"].value
        if b_val and str(b_val).strip() and str(b_val).strip() != "Total GMV":
            register(r, str(b_val).strip(), overwrite=True)

    for r in range(RWS_MAP_FIRST_ROW, RWS_MAP_LAST_ROW + 1):
        expert = ws[f"{RWS_MAP_COL_EXPERT}{r}"].value
        mapped = ws[f"{RWS_MAP_COL_RSC}{r}"].value
        if not expert:
            continue
        expert_s = str(expert).strip()
        expert_norm = normalize_category_name(expert_s)
        target_row = lookup.get(expert_norm)
        if target_row is None:
            for br in range(RWS_CAT_FIRST_ROW, RWS_CAT_LAST_ROW + 1):
                b_val = ws[f"B{br}"].value
                if b_val and normalize_category_name(b_val) == expert_norm:
                    target_row = br
                    break
        if target_row is None:
            continue
        register(target_row, expert_s, overwrite=False)
        if mapped:
            register(target_row, str(mapped).strip(), overwrite=False)

    return lookup, row_aliases


def match_raw_categories(ws, raw_data: Dict[str, Tuple[int, int]]):
    lookup, _ = build_category_row_lookup(ws)
    matched: Dict[int, Tuple[int, int]] = {}
    unmapped: List[str] = []
    mapping_log: List[str] = []

    for cat_name, values in raw_data.items():
        norm = normalize_category_name(cat_name)
        row = lookup.get(norm)
        if row is None:
            for alias_norm, candidate_row in lookup.items():
                if len(norm) > 8 and (norm in alias_norm or alias_norm in norm):
                    row = candidate_row
                    mapping_log.append(
                        f"  ~ {cat_name!r} → row {row} "
                        f"({ws[f'B{row}'].value!r}) [fuzzy match]"
                    )
                    break
        if row is None:
            unmapped.append(cat_name)
            mapping_log.append(f"  ✗ {cat_name!r} → NO MATCH")
        elif row not in matched:
            matched[row] = values
            mapping_log.append(
                f"  ✓ {cat_name!r} → row {row} ({ws[f'B{row}'].value!r})"
            )
        else:
            mapping_log.append(
                f"  ! {cat_name!r} → row {row} SKIPPED (row already filled)"
            )

    return matched, unmapped, mapping_log


def find_rws_last_block(ws, new_month: str) -> Tuple[str, str, str, str, str]:
    skip = ("deviation", "devaition", "%")
    last_ord, last_label = None, ""
    for c in range(4, 200):
        cl = col_str(c)
        v = ws[f"{cl}{RWS_MONTH_HDR_ROW}"].value
        if not v:
            continue
        if any(w in str(v).lower() for w in skip):
            break
        if "'" in str(v):
            last_ord = cl
            last_label = str(v).strip()
    if not last_ord:
        raise MonthSequenceError(
            "No month block found in Raw Working Sheet row 2. "
            "Check the master workbook layout."
        )
    new_parsed = parse_month_label(new_month)
    if new_parsed and month_labels_same_month(last_label, *new_parsed):
        raise MonthSequenceError(
            f"{new_month!r} already exists in Raw Working Sheet (last block: {last_label!r}). "
            f"Use the latest output workbook as master for the next month."
        )
    gmv = next_col(last_ord)
    return last_ord, gmv, next_col(gmv), next_col(next_col(gmv)), last_label


def find_rws_block_by_label(ws, target_label: str) -> Optional[Tuple[str, str]]:
    """
    Find the 4-column month block whose header (row 2) matches target_label
    (e.g. 'May'25'). Returns (ord_abs_col, gmv_abs_col) or None if not found.
    """
    target = target_label.strip().lower()
    for c in range(4, 200):
        cl = col_str(c)
        v = ws[f"{cl}{RWS_MONTH_HDR_ROW}"].value
        if v and target in str(v).strip().lower():
            return cl, next_col(cl)
    return None


def find_last_populated_block(ws, before_col: str) -> Optional[Tuple[str, str]]:
    """
    Among month-blocks strictly left of before_col, return (ord_col, gmv_col)
    for the rightmost block that actually has category numbers.
    Guards against a block whose header exists but was never filled in
    (seen in real files — e.g. a month started and abandoned).
    """
    skip = ("deviation", "devaition", "%")
    limit_n = col_num(before_col)
    candidates: List[str] = []
    for c in range(4, limit_n):
        cl = col_str(c)
        v = ws[f"{cl}{RWS_MONTH_HDR_ROW}"].value
        if not v:
            continue
        if any(w in str(v).lower() for w in skip):
            continue
        if "'" in str(v):
            candidates.append(cl)

    for cl in reversed(candidates):
        gmv_col = next_col(cl)
        has_data = any(
            isinstance(ws[f"{cl}{r}"].value, (int, float))
            or isinstance(ws[f"{gmv_col}{r}"].value, (int, float))
            for r in range(RWS_CAT_FIRST_ROW, RWS_CAT_LAST_ROW + 1)
        )
        if has_data:
            return cl, gmv_col
    return None


def read_category_values(
    ws, ord_col: str, gmv_col: str, name_map: Dict[str, str],
) -> Dict[str, Tuple[int, int]]:
    """Read literal (orders, gmv) per category from a Raw Working Sheet block,
    keyed by the canonical name (what column C would evaluate to — the same
    text Raw Data's own column A uses). Rows that consolidate to the same
    canonical name (many-to-one mapping) are summed, not overwritten."""
    result: Dict[str, Tuple[float, float]] = {}
    for r in range(RWS_CAT_FIRST_ROW, RWS_CAT_LAST_ROW + 1):
        name = resolve_canonical_name(ws, r, name_map)
        if not name:
            continue
        ord_v = ws[f"{ord_col}{r}"].value
        gmv_v = ws[f"{gmv_col}{r}"].value
        try:
            pair = (float(ord_v), float(gmv_v))
        except (TypeError, ValueError):
            continue
        if name in result:
            result[name] = (result[name][0] + pair[0], result[name][1] + pair[1])
        else:
            result[name] = pair
    return result


def run(
    wb,
    new_month: str,
    raw_data: Dict[str, Tuple[int, int]],
    dry_run: bool,
) -> RWSResult:
    """
    Entry point for Step 1.
    Returns RWSResult anchor contract (column letters written this run).
    """
    ws = wb["Raw Working Sheet"]
    matched, unmapped, mapping_log = match_raw_categories(ws, raw_data)

    print(f"\n  [RWS] Category mapping ({len(matched)} matched / {len(raw_data)} supplied):")
    for line in mapping_log:
        print(line)
    if unmapped:
        print(f"\n  [RWS] UNMAPPED ({len(unmapped)}):")
        for u in unmapped:
            print(f"         • {u}")

    name_map = build_canonical_name_map(ws)

    if dry_run:
        last_ord, _, _, _, last_lbl = find_rws_last_block(ws, new_month)
        print(f"  [DRY-RUN] Would add 4-col block after {last_lbl!r} ({last_ord}) "
              f"starting at {next_col(last_ord, 4)} for '{new_month}'.")
        matched_by_name: Dict[str, Tuple[int, int]] = {}
        for r, vals in matched.items():
            name = resolve_canonical_name(ws, r, name_map)
            if name:
                prior = matched_by_name.get(name, (0, 0))
                matched_by_name[name] = (prior[0] + vals[0], prior[1] + vals[1])
        populated = find_last_populated_block(ws, next_col(last_ord))
        prev_by_name = read_category_values(ws, *populated, name_map) if populated else {}
        yoy_block = find_rws_block_by_label(ws, year_ago_label_for(new_month))
        yoy_by_name = read_category_values(ws, *yoy_block, name_map) if yoy_block else {}
        return RWSResult(
            ord_abs_col="", gmv_abs_col="", ord_mn_col="", gmv_cr_col="",
            matched_by_name=matched_by_name, prev_by_name=prev_by_name,
            yoy_by_name=yoy_by_name, unmapped=unmapped,
        )

    last_ord, last_gmv, last_mn, last_cr, last_label = find_rws_last_block(ws, new_month)
    new_ord = next_col(last_cr)
    new_gmv = next_col(new_ord)
    new_mn  = next_col(new_gmv)
    new_cr  = next_col(new_mn)
    new_cols  = [new_ord, new_gmv, new_mn, new_cr]
    last_cols = [last_ord, last_gmv, last_mn, last_cr]

    print(f"  [RWS] Inserting '{new_month}' at {new_ord}-{new_cr} (after {last_label!r})")

    for old_col, new_col in zip(last_cols, new_cols):
        if old_col in ws.column_dimensions:
            ws.column_dimensions[new_col].width = ws.column_dimensions[old_col].width

    ws[f"{new_ord}{RWS_MONTH_HDR_ROW}"].value = new_month
    copy_cell_style(ws[f"{last_ord}{RWS_MONTH_HDR_ROW}"], ws[f"{new_ord}{RWS_MONTH_HDR_ROW}"])

    for old_col, new_col in zip(last_cols, new_cols):
        for hdr_row in (RWS_SUBLABEL_ROW, RWS_COL_HDR_ROW):
            ws[f"{new_col}{hdr_row}"].value = ws[f"{old_col}{hdr_row}"].value
            copy_cell_style(ws[f"{old_col}{hdr_row}"], ws[f"{new_col}{hdr_row}"])

    for r in range(RWS_CAT_FIRST_ROW, RWS_CAT_LAST_ROW + 1):
        orders_abs, gmv_abs = matched.get(r, (0, 0))
        ws[f"{new_ord}{r}"].value = orders_abs
        copy_cell_style(ws[f"{last_ord}{r}"], ws[f"{new_ord}{r}"])
        ws[f"{new_gmv}{r}"].value = gmv_abs
        copy_cell_style(ws[f"{last_gmv}{r}"], ws[f"{new_gmv}{r}"])
        ws[f"{new_mn}{r}"].value  = f"={new_ord}{r}/10^6"
        copy_cell_style(ws[f"{last_mn}{r}"], ws[f"{new_mn}{r}"])
        ws[f"{new_cr}{r}"].value  = f"={new_gmv}{r}/10^7"
        copy_cell_style(ws[f"{last_cr}{r}"], ws[f"{new_cr}{r}"])

    ws[f"{new_ord}{RWS_TOTAL_GMV_ROW}"].value = sum(v[0] for v in matched.values())
    ws[f"{new_gmv}{RWS_TOTAL_GMV_ROW}"].value = sum(v[1] for v in matched.values())
    ws[f"{new_mn}{RWS_TOTAL_GMV_ROW}"].value  = f"={new_ord}{RWS_TOTAL_GMV_ROW}/10^6"
    ws[f"{new_cr}{RWS_TOTAL_GMV_ROW}"].value  = f"={new_gmv}{RWS_TOTAL_GMV_ROW}/10^7"

    if unmapped:
        ws["B42"].value = "⚠ UNMAPPED from new raw data (not in BV:BW mapping):"
        ws["B42"].fill = ORANGE_FILL
        for i, cat in enumerate(unmapped):
            ws[f"B{43 + i}"].value = cat
            ws[f"B{43 + i}"].fill = ORANGE_FILL

    print(f"  [RWS] ✓ Wrote {len(matched)} categories to columns {new_ord}–{new_cr}.")

    matched_by_name: Dict[str, Tuple[int, int]] = {}
    for r, vals in matched.items():
        name = resolve_canonical_name(ws, r, name_map)
        if name:
            prior = matched_by_name.get(name, (0, 0))
            matched_by_name[name] = (prior[0] + vals[0], prior[1] + vals[1])

    populated = find_last_populated_block(ws, next_col(last_ord))
    prev_by_name = read_category_values(ws, *populated, name_map) if populated else {}
    if not populated:
        print("  [RWS] Note: no prior month with category-level numbers found "
              "— MoM flagging will be skipped this run.")
    elif populated != (last_ord, last_gmv):
        print(f"  [RWS] Note: {last_label!r} block ({last_ord}) has no category data — "
              f"using {populated[0]} for MoM comparison instead.")

    yoy_block = find_rws_block_by_label(ws, year_ago_label_for(new_month))
    yoy_by_name = read_category_values(ws, *yoy_block, name_map) if yoy_block else {}
    if not yoy_block:
        print(f"  [RWS] Note: no {year_ago_label_for(new_month)!r} block found "
              f"— YoY flagging will be skipped this run.")

    return RWSResult(
        ord_abs_col=new_ord,
        gmv_abs_col=new_gmv,
        ord_mn_col=new_mn,
        gmv_cr_col=new_cr,
        matched_by_name=matched_by_name,
        prev_by_name=prev_by_name,
        yoy_by_name=yoy_by_name,
        unmapped=unmapped,
    )
