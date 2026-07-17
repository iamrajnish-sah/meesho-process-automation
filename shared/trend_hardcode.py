"""Generate next-month hardcoded values within ±1pp, following historical MoM trend."""
import random
from typing import Dict, List, Optional, Tuple

from shared.style_utils import YELLOW_FILL, add_comment


def cell_val(ws, row: int, col: int) -> Optional[float]:
    v = ws.cell(row, col).value
    try:
        f = float(v)
        return None if f != f else f
    except (TypeError, ValueError):
        return None


def is_bare(ws, row: int, col: int) -> bool:
    return isinstance(ws.cell(row, col).value, (int, float))


def hist_value_series(
    ws, row: int, end_col_n: int, months: int,
) -> List[Optional[float]]:
    series: List[Optional[float]] = []
    for i in range(months - 1, -1, -1):
        col_n = end_col_n - i
        series.append(cell_val(ws, row, col_n) if col_n >= 1 else None)
    return series


def pp_bound_for(prev_val: float) -> float:
    """±1 percentage point in the value's native storage units."""
    # Decimal fractions (e.g. 0.0444 = 4.44%)
    if abs(prev_val) < 0.15:
        return 0.01
    # Whole percent points (e.g. 1.42 = 1.42%, 4.623 = 4.623%)
    return 1.0


def pp_change_display(prev_val: float, delta: float) -> float:
    """Express MoM change in percentage points for validation output."""
    if abs(prev_val) < 0.15:
        return delta * 100.0
    return delta


def generate_pp_bounded(
    ws,
    row: int,
    prev_col_n: int,
    *,
    months: int = 6,
    rng: Optional[random.Random] = None,
    pp_bound: Optional[float] = None,
    floor: Optional[float] = None,
) -> Tuple[Optional[float], Dict]:
    """Return (new_value, meta) for one hardcoded newest-month cell."""
    prev_val = cell_val(ws, row, prev_col_n)
    if prev_val is None:
        return None, {}

    if pp_bound is None:
        pp_bound = pp_bound_for(prev_val)

    series = hist_value_series(ws, row, prev_col_n, months)
    numeric = [v for v in series if v is not None]
    deltas = [numeric[i] - numeric[i - 1] for i in range(1, len(numeric))]

    trend = sum(deltas) / len(deltas) if deltas else 0.0
    trend = max(-pp_bound, min(pp_bound, trend))
    if trend > 0:
        hist_dir = "UP"
    elif trend < 0:
        hist_dir = "DN"
    else:
        hist_dir = "FLAT"

    if rng is None:
        rng = random.Random()
    noise = rng.uniform(-0.2, 0.2) * pp_bound
    candidate = prev_val + trend + noise
    lo, hi = prev_val - pp_bound, prev_val + pp_bound
    new_val = max(lo, min(hi, candidate))
    if floor is not None:
        new_val = max(floor, new_val)
        new_val = min(hi, new_val)

    actual_delta = new_val - prev_val
    followed = (
        hist_dir == "FLAT"
        or (hist_dir == "UP" and actual_delta >= 0)
        or (hist_dir == "DN" and actual_delta <= 0)
    )

    return new_val, {
        "prev": prev_val,
        "delta": actual_delta,
        "pp_change": pp_change_display(prev_val, actual_delta),
        "pp_bound": pp_bound,
        "trend_dir": hist_dir,
        "followed_trend": followed,
    }


def write_hardcoded_cell(
    ws,
    row: int,
    col_n: int,
    new_val: float,
    *,
    tag: str,
    prev_col: str,
    new_col: str,
    meta: Dict,
) -> None:
    cell = ws.cell(row, col_n)
    cell.value = new_val
    cell.fill = YELLOW_FILL
    add_comment(
        ws, f"{new_col}{row}",
        f"[{tag}] Generated {new_col} from {prev_col}{row}.\n"
        f"Prior={meta.get('prev'):.6g}, MoM={meta.get('pp_change'):+.3f}pp "
        f"(±{meta.get('pp_bound'):.4g} cap), trend={meta.get('trend_dir')}.",
    )
