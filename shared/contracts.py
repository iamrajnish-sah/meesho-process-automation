"""
Anchor contracts passed between pipeline steps.

Each sheet module RETURNS a contract describing exactly which columns it wrote.
run_pipeline.py forwards these to downstream modules — no re-detection of
upstream columns via string scanning in downstream code.
"""
from typing import Dict, List, Optional, TypedDict


class RWSResult(TypedDict):
    """Output of raw_working_sheet.run()."""
    ord_abs_col: str
    gmv_abs_col: str
    ord_mn_col: str
    gmv_cr_col: str
    matched_by_name: Dict[str, tuple]   # category → (orders, gmv) for the NEW month
    prev_by_name: Dict[str, tuple]      # category → (orders, gmv) for the PREVIOUS month
    yoy_by_name: Dict[str, tuple]       # category → (orders, gmv) for same month LAST YEAR
    unmapped: List[str]


class RDContract(TypedDict):
    """Columns written by raw_data.run()."""
    prev_col: str
    gmv_col: str
    mom_col: str
    yoy_col: Optional[str]
    yoy_base_col: Optional[str]


class MPVContract(TypedDict):
    """Columns written by meesho_prelim.run()."""
    expert_prev_col: str
    expert_col: str
    rsc_prev_col: str
    rsc_col: str
    yoy_expert_prev_col: str
    yoy_expert_col: str
    yoy_rsc_prev_col: str
    yoy_rsc_col: str
    col_map: Dict[str, str]   # full remap dict (MPV + Raw Data cols) for downstream


class EMContract(TypedDict):
    """Columns written by error_margin.run()."""
    section_new_cols: List[str]


class FKContract(TypedDict):
    """Columns written by fk_prelim.run()."""
    section_new_cols: List[str]


class CPContract(TypedDict):
    """Columns written by client_prelim.run()."""
    abs_col: str
    yoy_col: str
    expert_yoy_col: str
    client_mom_col: str
