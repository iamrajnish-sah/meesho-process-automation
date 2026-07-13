# Workflow — how to make a change without breaking something else

Read `PROJECT_CONTEXT.md` first if you haven't already. This file is the
procedure; that file is the background you need to follow it.

## When told "sheet X is not working properly, fix it"

Follow these steps in order. Do not skip to writing code.

**1. Identify what actually changed.**
Is the requested fix about *values* (a formula computes the wrong number, a
trendline looks off) or about *shape* (a column is in the wrong place, a
contract field is missing or wrong)? This determines your blast radius:
- Values-only fix → touch only that sheet's `.py` file. Nothing downstream can
  be affected, because the contract it returns doesn't change.
- Shape fix → touch that sheet's `.py` file AND check every downstream module
  that consumes its contract (see the dependency order in PROJECT_CONTEXT.md).
  Do not assume downstream is fine; check it.

**2. Read the contract, don't re-derive it.**
Before changing a sheet module, look at what `shared/contracts.py` says it
receives from upstream and what it's supposed to return. If the bug is "this
sheet used the wrong column," check first whether it's reading the contract
correctly, before assuming the upstream sheet produced the wrong contract.

**3. Make the change in exactly one sheet module (plus `shared/` only if the
bug is in shared logic).**
If you find yourself editing two sheet-specific files to fix one reported
problem, stop and check whether the real bug is in `shared/` and is showing up
in two places, versus you drifting into an unrelated file.

**4. If you touch `shared/formula_utils.py`, `shared/anchors.py`, or any other
shared file: this affects every sheet that imports it. List which sheets
import the function you're changing, and check each one's calling pattern is
still correct after your change** — not just the one sheet that reported the
bug. (This is exactly how the double-shift bug crept back into
`client_prelim.py`: a shared function's behavior assumption silently changed
without every caller being re-checked.)

**5. Regression check before declaring done.**
Re-run the pipeline for the known-good prior transition (March→April) and diff
the output formulas against a known-good reference. If you don't have a saved
known-good reference file yet, save one now (`Meesho_may_output_FIXED.xlsx` or
similar, from a run you've manually verified) and use it for this check going
forward. `compare_workbook_formulas.py` already exists in this repo for this
purpose — use it, don't skip it.

**6. Report back in terms of contract changes, not just "fixed."**
State explicitly: which file(s) changed, whether any contract's shape changed,
and which downstream sheets you checked as a result. If nothing downstream
needed checking because the fix was values-only, say so — that's the whole
point of the contract system, and confirming it out loud is what stops
"fix one thing, break three others."

## When told "build a new feature" (e.g. the web app, Gemini ingestion)

New features should not be added by editing pipeline sheet modules directly.
`web_app.py`, `web_uploads.py`, and `gemini_ingest.py` already exist as a
separate layer that calls into `run_pipeline.py` — keep it that way. If a new
feature seems to require changing `raw_data.py` or another sheet module, stop
and ask whether it actually needs to, or whether it belongs in the calling
layer instead.

## Immediate next step

No known unfixed bugs as of Jul 2026. The Client Prelim double-shift bug
(documented in PROJECT_CONTEXT.md §2) was fixed in `shared/formula_utils.py`.
All six sheet modules are considered correct. The next work item is adding a
regression-check comparison run to `compare_workbook_formulas.py` so future
changes have a concrete before/after diff to validate against.
