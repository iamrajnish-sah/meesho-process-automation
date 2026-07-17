"""FK Prelim View manual inputs for ASP, AoV, and Cancellation % (newest month only)."""
from dataclasses import dataclass
from typing import Optional


@dataclass(frozen=True)
class FKManualInputs:
    """User-supplied values for FK Prelim rows 7, 8, 11 in the newest Expert column."""
    asp: float
    aov: float
    cancellation_pct: float  # decimal fraction, e.g. 0.1395 for 13.95%


FK_MANUAL_ROWS = {
    "asp": 7,
    "aov": 8,
    "cancellation_pct": 11,
}


def _parse_float(raw: str, field: str) -> float:
    text = raw.strip().replace(",", "")
    if not text:
        raise ValueError(f"{field} is required.")
    val = float(text)
    if field == "cancellation_pct" and val > 1.0:
        # Allow entry as whole percent (e.g. 13.95) — convert to decimal.
        if val <= 100.0:
            val = val / 100.0
    if field == "cancellation_pct" and not (0.0 <= val <= 1.0):
        raise ValueError(f"Cancellation % must be between 0 and 100% (got {val}).")
    if field in ("asp", "aov") and val <= 0:
        raise ValueError(f"{field} must be positive (got {val}).")
    return val


def parse_fk_manual_inputs(
    asp: str,
    aov: str,
    cancellation_pct: str,
) -> FKManualInputs:
    """Parse form/CLI strings into FKManualInputs."""
    return FKManualInputs(
        asp=_parse_float(asp, "asp"),
        aov=_parse_float(aov, "aov"),
        cancellation_pct=_parse_float(cancellation_pct, "cancellation_pct"),
    )


def prompt_fk_manual_inputs(new_month: str) -> FKManualInputs:
    """Terminal prompts — blocks until three valid values are entered."""
    print(f"\n  [FK Manual] Enter values for {new_month} "
          f"(FK Prelim rows 7/8/11 — not auto-generated):")
    while True:
        try:
            asp = _parse_float(
                input(f"  ASP for {new_month}: "), "asp",
            )
            aov = _parse_float(
                input(f"  AoV for {new_month}: "), "aov",
            )
            cancel = _parse_float(
                input(f"  Cancellation % for {new_month}: "), "cancellation_pct",
            )
            return FKManualInputs(asp=asp, aov=aov, cancellation_pct=cancel)
        except ValueError as exc:
            print(f"  ⚠ {exc} Try again.")


def resolve_fk_manual_inputs(
    new_month: str,
    fk_manual: Optional[FKManualInputs],
    *,
    interactive: bool = True,
) -> FKManualInputs:
    """Use provided values or prompt when interactive; raise if missing in non-interactive mode."""
    if fk_manual is not None:
        return fk_manual
    if interactive:
        return prompt_fk_manual_inputs(new_month)
    raise ValueError(
        "FK Prelim ASP, AoV, and Cancellation % are required for the new month. "
        "Pass fk_manual=FKManualInputs(...) or run interactively."
    )
