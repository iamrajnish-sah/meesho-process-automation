# Project Context — Meesho Monthly Update Pipeline

Read this file before making any change. It exists because this pipeline has
already broken twice in ways that were hard to track down — both times because
a fix made in one sheet's code silently didn't apply to, or got lost from,
another. This file is the fix for that: the persistent memory the codebase
doesn't otherwise have.

## What this is

A deterministic pipeline that takes one month's raw category-level GMV/Orders
data and propagates it across six Excel sheets in a fixed master workbook. It
is NOT a general-purpose app — the sheet layout, column meanings, and
dependency order are fixed and known. Treat any code that "detects" structure
at runtime with suspicion: detection is a workaround for not having a fixed
answer, and every bug so far has come from a detection heuristic guessing
wrong.

## The six sheets, in strict dependency order

1. **Raw Working Sheet** — raw category data lands here first (append-right).
2. **Raw Data** — reads Raw Working Sheet's new columns via XLOOKUP. Computes
   MoM and YoY.
3. **Meesho Prelim View** — reads Raw Data's new GMV/MoM/YoY columns.
4. **Error Margin - Expert vs Report** — reads upstream sheets; has checker
   rows that must evaluate to 0/TRUE.
5. **FK Prelim View** — parallel structure to Error Margin.
6. **Client Prelim** — reads Meesho Prelim View's columns (via
   `MPVContract.col_map`) and Raw Data's columns, copies a 6-column block per
   month (`CP_BLOCK_COL_DELTA = 6`).

**New_Meesho Masterfile is off-limits.** Never read from or write to it as
part of this pipeline.

This is a strict pipeline, not six independent modules. A change to any
sheet's *output shape* (which column it writes to, what a contract field
means) can silently break every sheet downstream of it. A change to a
sheet's *internal values only* (the numbers it computes) cannot break
anything downstream, as long as the contract it returns is unchanged.
Always know which kind of change you're making before you start.

## The contract system (`shared/contracts.py`, `shared/anchors.py`)

Each sheet module returns a `TypedDict` (`RWSResult`, `RDContract`,
`MPVContract`, etc.) describing exactly which columns it wrote for the new
month. `run_pipeline.py` passes each contract to the next module. **No sheet
module should ever re-detect a previous sheet's column by scanning that
sheet's cells for date-like text.** It should receive the column from the
contract instead. If a sheet needs to know something the contract doesn't
currently carry, add the field to the contract — don't work around it with a
scan.

Two anchor-detection functions exist in `shared/anchors.py`, and they are not
interchangeable safety-wise:
- `find_contiguous_last_col` — validates the label at the anchor column
  actually looks like the expected previous month before returning it. Safe.
  Used by Raw Data and Meesho Prelim View.
- `find_last_month_col_in_row` — no validation; returns whatever it scans
  into. This is the pattern that caused the original Raw Data bug (it landed
  on an unrelated scratch-note column). **Client Prelim still uses this one.**
  If Client Prelim ever misbehaves, this is the first place to look, and it
  should eventually be migrated to the validated pattern.

## Known bug classes — do not reintroduce these

### 1. Anchor detected by loose scanning instead of validated lookup
**Symptom:** new month's data lands in a disconnected column, or MoM/YoY
divides against the wrong "previous" column.
**Cause:** a function scans a whole row for anything date-like and takes the
rightmost match, which can land on an unrelated note/scratch cell instead of
the real end of the series.
**Fix pattern:** always use `find_contiguous_last_col` (or equivalent
validated lookup) — never a bare scan — for any sheet where the row may
contain non-series content.

### 2. Double-shifted formula references
**Symptom:** a formula that should reference column V instead references W
(off by exactly the block-delta amount); a cross-sheet reference points past
where the actual data is.
**Cause:** `remap_formula_refs(formula, col_map, col_shift)` applies the
explicit `col_map` substitution AND THEN applies `shift_formula` with
`col_shift` on top of the *already-substituted* result — so any column
present in `col_map` gets shifted twice.
**Fix pattern:** if a column is already being remapped via `col_map`, do not
also pass a nonzero `col_shift` that would apply to it. Either pass
`col_shift=0` (as `meesho_prelim.py` correctly does via `copy_column_block`),
or protect mapped columns from the shift (e.g. placeholder-substitute them out
before shifting, restore after).
**FIXED (Jul 2026).** `remap_formula_refs` in `shared/formula_utils.py` now
uses lowercase placeholder tokens (e.g. `zzmap0zz`) when `col_shift != 0`.
Columns substituted via `col_map` are swapped to a placeholder before
`shift_formula` runs; the regex in `shift_formula` only matches `[A-Z]+` so
placeholders survive unshifted, then are restored to their final values
after. Callers that pass `col_shift=0` (`copy_column_block`, the AW-column
call site in `client_prelim.py`) take the fast path and are unaffected.

**Verified before/after (May'26 run):**
| Cell | Before (wrong) | After (correct) |
|------|----------------|-----------------|
| BF6  | `='Meesho Prelim View'!AB8` | `='Meesho Prelim View'!V8` |
| BF7  | `='Meesho Prelim View'!AB10` | `='Meesho Prelim View'!V10` |
| BF9  | `='Raw Data'!Y25` | `='Raw Data'!S25` |
| BG6  | `='Meesho Prelim View'!AY8` | `='Meesho Prelim View'!AS8` |
| BF8  | `=BF7/BF6` (same-sheet — unchanged, correct) | `=BF7/BF6` |

Client Prelim anchor also migrated from bare `find_last_month_col_in_row` to
the validated pattern (adds `validate_new_month_against_last`) so a wrong
anchor now raises `MonthSequenceError` instead of silently writing to the
wrong column.

## Definition of "done" for any change

A sheet is not done just because it ran without a Python exception. Before
calling a change complete:
1. The sheet's contract (if any) is unchanged in *shape* — only values differ
   — unless you deliberately changed the contract, in which case every
   downstream consumer of that field must be checked, not assumed fine.
2. Re-run the known-good prior-month transition (currently March→April) and
   confirm formulas are structurally identical to before your change, for any
   rows you didn't intend to touch.
3. For Raw Data: the summed category-level GMV for the new month equals the
   raw input file's total GMV.
4. For Error Margin: checker rows evaluate to 0/TRUE after a manual
   recalculation in Excel (Ctrl+Alt+F9).
5. Spot-check at least one formula in the new month's column against the
   equivalent formula in the prior month's column, side by side, to confirm
   the column reference actually shifted by the amount you expect — not more,
   not less.
