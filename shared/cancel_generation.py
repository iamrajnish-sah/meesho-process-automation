"""Cancellation % generation — same rules as Error Margin Cancellation block."""
import random
from typing import List, Optional, Tuple

CANCEL_HARD_CAP_PP = 2.0
CANCEL_HIST_MONTHS = 6


def hist_mom_swings_pp_from_col_n(
    ws, row: int, prev_col_n: int, months: int = CANCEL_HIST_MONTHS,
) -> List[float]:
    """MoM swings in percentage points over the last `months` columns."""
    start_n = prev_col_n - (months - 1)
    vals: List[Optional[float]] = []
    for c in range(start_n, prev_col_n + 1):
        if c < 1:
            vals.append(None)
            continue
        v = ws.cell(row, c).value
        try:
            vals.append(float(v))
        except (TypeError, ValueError):
            vals.append(None)
    deltas: List[float] = []
    for i in range(1, len(vals)):
        if vals[i] is not None and vals[i - 1] is not None:
            deltas.append((vals[i] - vals[i - 1]) * 100.0)
    return deltas


def generate_cancel_pct(
    apr_pct: float,
    hist_deltas_pp: List[float],
    rng: random.Random,
    *,
    gmv_apr: Optional[float] = None,
    gmv_may: Optional[float] = None,
    hard_cap_pp: float = CANCEL_HARD_CAP_PP,
) -> Tuple[float, float, str, float]:
    """Return (may_pct, delta_pp, direction_label, cap_pp_used).

    Same algorithm as ``error_margin._fill_cancellation_hardcoded`` per-row body.
    When GMV is omitted, direction follows the historical MoM average.
    """
    if gmv_apr is not None and gmv_may is not None and gmv_apr > 0:
        if gmv_may > gmv_apr * (1 + 1e-12):
            direction = 1
            dir_s = "UP"
        elif gmv_may < gmv_apr * (1 - 1e-12):
            direction = -1
            dir_s = "DN"
        else:
            direction = 0
            dir_s = "FLAT"
    else:
        avg = sum(hist_deltas_pp) / len(hist_deltas_pp) if hist_deltas_pp else 0.0
        if avg > 0.01:
            direction, dir_s = 1, "UP"
        elif avg < -0.01:
            direction, dir_s = -1, "DN"
        else:
            direction, dir_s = 0, "FLAT"

    hist_max = max((abs(d) for d in hist_deltas_pp), default=0.3)
    cap_pp = min(hard_cap_pp, max(hist_max, 0.05))

    if direction == 0:
        delta_pp = 0.0
    else:
        delta_pp = direction * rng.uniform(cap_pp * 0.25, cap_pp)

    may_pct = max(0.0, min(0.5, apr_pct + delta_pp / 100.0))
    return may_pct, delta_pp, dir_s, cap_pp
