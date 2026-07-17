"""Evaluate NMM / upstream formulas in-memory for Contribution row-73 balancing."""
import re
from typing import Dict, Optional, Set, Tuple

_CROSS_SHEET = re.compile(
    r"^'([^']+)'!(\$?)([A-Z]+)(\$?)(\d+)$",
)
_SAME_CELL = re.compile(r"^(\$?)([A-Z]+)(\$?)(\d+)$")
_MUL = re.compile(r"^(.+)\*(.+)$")
_SUM_RANGE = re.compile(
    r"^SUM\((\$?)([A-Z]+)(\$?)(\d+):(\$?)([A-Z]+)(\$?)(\d+)\)$",
    re.IGNORECASE,
)
_SUM_ARGS = re.compile(
    r"^SUM\((.+)\)$",
    re.IGNORECASE,
)


def _resolve_sheet(wb, name: str) -> Optional[str]:
    if name in wb.sheetnames:
        return name
    compact = name.replace(" ", "")
    for candidate in wb.sheetnames:
        if candidate.replace(" ", "") == compact:
            return candidate
    return None


def _eval_cell(
    wb,
    sheet: str,
    coord: str,
    cache: Dict[Tuple[str, str], Optional[float]],
    visiting: Set[Tuple[str, str]],
) -> Optional[float]:
    resolved = _resolve_sheet(wb, sheet)
    if resolved is None:
        return None
    sheet = resolved
    key = (sheet, coord.upper())
    if key in cache:
        return cache[key]
    if key in visiting:
        return None
    visiting.add(key)

    raw = wb[sheet][coord].value
    if isinstance(raw, (int, float)):
        cache[key] = float(raw)
        visiting.discard(key)
        return cache[key]
    if not isinstance(raw, str) or not raw.startswith("="):
        cache[key] = None
        visiting.discard(key)
        return None

    result = _eval_expr(wb, sheet, raw[1:].strip(), cache, visiting)
    cache[key] = result
    visiting.discard(key)
    return result


def _eval_expr(
    wb,
    sheet: str,
    expr: str,
    cache: Dict[Tuple[str, str], Optional[float]],
    visiting: Set[Tuple[str, str]],
) -> Optional[float]:
    expr = expr.replace(" ", "")

    m = _CROSS_SHEET.match(expr)
    if m:
        sh = _resolve_sheet(wb, m.group(1))
        if sh is None:
            return None
        return _eval_cell(wb, sh, f"{m.group(3)}{m.group(5)}", cache, visiting)

    m = _SUM_RANGE.match(expr)
    if m:
        c1, r1, c2, r2 = m.group(2), int(m.group(4)), m.group(6), int(m.group(8))
        if c1 != c2:
            return None
        total = 0.0
        ok = False
        for row in range(min(r1, r2), max(r1, r2) + 1):
            v = _eval_cell(wb, sheet, f"{c1}{row}", cache, visiting)
            if v is not None:
                total += v
                ok = True
            elif wb[sheet][f"{c1}{row}"].value is not None:
                return None
        return total if ok else None

    m = _SUM_ARGS.match(expr)
    if m:
        parts = _split_top_level(m.group(1), ",")
        total = 0.0
        for part in parts:
            v = _eval_sum_arg(wb, sheet, part.strip(), cache, visiting)
            if v is None:
                return None
            total += v
        return total

    if "*" in expr and "(" not in expr.split("*")[0]:
        m = _MUL.match(expr)
        if m:
            left = _eval_expr(wb, sheet, m.group(1), cache, visiting)
            right = _eval_expr(wb, sheet, m.group(2), cache, visiting)
            if left is not None and right is not None:
                return left * right

    if "*(" in expr:
        m = re.match(r"^([A-Z]+\d+)\*\(1-([A-Z]+\d+)\)$", expr)
        if m:
            a = _eval_expr(wb, sheet, m.group(1), cache, visiting)
            b = _eval_expr(wb, sheet, m.group(2), cache, visiting)
            if a is not None and b is not None:
                return a * (1.0 - b)

    parts = _split_top_level(expr, "+")
    if len(parts) > 1:
        total = 0.0
        for part in parts:
            for sub in _split_top_level(part, "-"):
                if sub == part:
                    v = _eval_expr(wb, sheet, sub, cache, visiting)
                else:
                    subs = _split_signed(part)
                    v = None
                    acc = 0.0
                    ok = False
                    for sign, term in subs:
                        t = _eval_expr(wb, sheet, term, cache, visiting)
                        if t is None:
                            return None
                        acc += sign * t
                        ok = True
                    v = acc if ok else None
                if v is None:
                    return None
                total += v
        return total

    subs = _split_signed(expr)
    if len(subs) > 1:
        acc = 0.0
        for sign, term in subs:
            v = _eval_expr(wb, sheet, term, cache, visiting)
            if v is None:
                return None
            acc += sign * v
        return acc

    m = _SAME_CELL.match(expr)
    if m:
        return _eval_cell(wb, sheet, f"{m.group(2)}{m.group(4)}", cache, visiting)

    return None


def _eval_sum_arg(
    wb,
    sheet: str,
    arg: str,
    cache: Dict[Tuple[str, str], Optional[float]],
    visiting: Set[Tuple[str, str]],
) -> Optional[float]:
    arg = arg.strip()
    m = _CROSS_SHEET.match(arg)
    if m:
        sh = _resolve_sheet(wb, m.group(1))
        if sh is None:
            return None
        return _eval_cell(wb, sh, f"{m.group(3)}{m.group(5)}", cache, visiting)
    m = _SUM_RANGE.match(arg)
    if m:
        return _eval_expr(wb, sheet, f"SUM({arg[4:-1]})", cache, visiting)
    return _eval_expr(wb, sheet, arg, cache, visiting)


def _split_top_level(text: str, sep: str) -> list:
    parts: list = []
    depth = 0
    start = 0
    for i, ch in enumerate(text):
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif ch == sep and depth == 0:
            parts.append(text[start:i])
            start = i + 1
    parts.append(text[start:])
    return [p for p in parts if p]


def _split_signed(expr: str) -> list:
    terms: list = []
    i = 0
    sign = 1
    while i < len(expr):
        if expr[i] == "+":
            sign = 1
            i += 1
            continue
        if expr[i] == "-":
            sign = -1
            i += 1
            continue
        depth = 0
        start = i
        while i < len(expr):
            if expr[i] == "(":
                depth += 1
            elif expr[i] == ")":
                depth -= 1
            elif expr[i] in "+-" and depth == 0:
                break
            i += 1
        terms.append((sign, expr[start:i]))
        sign = 1
    return terms


def contrib_formula_rows_share_sum(
    wb,
    sheet: str,
    metrics_col: str,
    rows: Tuple[int, ...] = (68, 69, 70, 71, 72),
    denom_row: int = 64,
) -> Optional[float]:
    """Sum of metrics_row / metrics$64 for formula-driven Contribution rows."""
    cache: Dict[Tuple[str, str], Optional[float]] = {}
    visiting: Set[Tuple[str, str]] = set()
    denom = _eval_cell(wb, sheet, f"{metrics_col}{denom_row}", cache, visiting)
    if denom is None or denom == 0:
        return None
    total = 0.0
    for row in rows:
        num = _eval_cell(wb, sheet, f"{metrics_col}{row}", cache, visiting)
        if num is None:
            return None
        total += num / denom
    return total
