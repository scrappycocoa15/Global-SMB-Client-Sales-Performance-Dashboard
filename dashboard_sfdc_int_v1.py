"""
SAP Concur INT SMB Client Sales — Live SFDC Dashboard
Markets: Canada (CAD) · United Kingdom (GBP) · Australia (AUD)

Run:  streamlit run dashboard_sfdc_int_v1.py
"""

import io, json, re, time as _time, warnings
from calendar import monthrange
from datetime import datetime
from pathlib import Path

import requests as _requests

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="SAP Concur | INT SMB Client Sales Dashboard",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# BRAND
# ─────────────────────────────────────────────────────────────────────────────
C = dict(blue="#0070F2", light_blue="#4CB1FF", dark_blue="#00144A",
         green="#188918", red="#BB0000", amber="#E8A000", yellow="#F0AB00",
         grey="#E1E2E6", dark_grey="#6A6D73", bg="#FFFFFF", bg_subtle="#F5F6F7")

MARKET_COLORS = {
    "Canada":         C["blue"],
    "United Kingdom": C["dark_blue"],
    "Australia":      C["light_blue"],
}
BUCKET_ORDER  = ["<50%", "50-75%", "75-100%", "100-120%", "120%+"]
BUCKET_COLORS = {"<50%": C["red"], "50-75%": C["amber"], "75-100%": C["yellow"],
                 "100-120%": "#5DB533", "120%+": C["green"]}
MONTH_NAMES   = ["January","February","March","April","May","June",
                 "July","August","September","October","November","December"]
MONTH_ABBREV  = ["Jan","Feb","Mar","Apr","May","Jun",
                 "Jul","Aug","Sep","Oct","Nov","Dec"]

# ─────────────────────────────────────────────────────────────────────────────
# INT ORG STRUCTURE
# ─────────────────────────────────────────────────────────────────────────────

# Currency per market (used in SFDC reports and quota file)
MARKET_CURRENCY = {
    "Canada":         "CAD",
    "United Kingdom": "GBP",
    "Australia":      "AUD",
}
CURRENCY_SYMBOL = {"CAD": "C$", "GBP": "£", "AUD": "A$", "USD": "$"}

# Concur Team → FLSM (first-level sales manager / equivalent of RSD)
TEAM_LEADER_MAP = {
    "Canada SMB Client Sales Key":        "Lesley Nunes",
    "Canada SMB Client Sales Premier":    "Lesley Nunes",
    "Canada SMB Client Sales Strategic":  "Lesley Nunes",
    "UK SMB Client Sales Key":            "Katie Brown",
    "UK SMB Client Sales Strategic":      "Bret Edis",
    "UK SMB Client Sales Premier":        "Bret Edis",
    "UK Mid Market":                      "Jessica Brown",
    "UK National":                        "Ryan Headington",
    "UK Premier":                         "Ryan Headington",
    "UK Ireland":                         "Nic Henney",
    "Australia SMB Client Sales":         "Peter Soukos",
    "Australia Mid Market":               "Hamish Tebbutt",
}

# FLSM → Director (one level up)
FLSM_DIRECTOR = {
    "Lesley Nunes":     "Brian Veloso",
    "Katie Brown":      "Adam Bazeley",
    "Bret Edis":        "Adam Bazeley",
    "Jessica Brown":    "Angus Milledge",
    "Ryan Headington":  "Angus Milledge",
    "Nic Henney":       "Angus Milledge",
    "Peter Soukos":     "Fabian Calle",
    "Hamish Tebbutt":   "Fabian Calle",
}

# Director → Market
DIRECTOR_MARKET = {
    "Brian Veloso":   "Canada",
    "Adam Bazeley":   "United Kingdom",
    "Angus Milledge": "United Kingdom",
    "Fabian Calle":   "Australia",
}

# Concur Team → Market
TEAM_MARKET = {
    "Canada SMB Client Sales Key":        "Canada",
    "Canada SMB Client Sales Premier":    "Canada",
    "Canada SMB Client Sales Strategic":  "Canada",
    "UK SMB Client Sales Key":            "United Kingdom",
    "UK SMB Client Sales Strategic":      "United Kingdom",
    "UK SMB Client Sales Premier":        "United Kingdom",
    "UK Mid Market":                      "United Kingdom",
    "UK National":                        "United Kingdom",
    "UK Premier":                         "United Kingdom",
    "UK Ireland":                         "United Kingdom",
    "Australia SMB Client Sales":         "Australia",
    "Australia Mid Market":               "Australia",
}

# Market leaders (used for org-level attainment attribution)
MARKET_LEADER = {
    "Canada":         "Brian Veloso",
    "United Kingdom": "Cassie Petrie",
    "Australia":      "Fabian Calle",
}

# Name normalisations (add as needed when SFDC names differ from quota file)
NAME_MAP = {
    "Conor Tomlinson": "conor tomlinson",   # quota file uses lowercase
}

DEFAULT_TARGETS = Path(__file__).parent / "2026 SMB quotas in local currency 9.17.26.xlsx"

# ─────────────────────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
html,body,[class*="css"]{{font-family:'72','Helvetica Neue',Arial,sans-serif;}}
[data-testid="stSidebar"]{{background:{C['bg_subtle']};border-right:1px solid {C['grey']};}}
.kpi-card{{background:{C['bg']};border:1px solid {C['grey']};border-top:4px solid {C['blue']};
  border-radius:8px;padding:18px 20px 14px;text-align:center;
  box-shadow:0 1px 4px rgba(0,0,0,0.06);}}
.kpi-card.green{{border-top-color:{C['green']};}}
.kpi-card.amber{{border-top-color:{C['amber']};}}
.kpi-card.dark{{border-top-color:{C['dark_blue']};}}
.kpi-value{{font-size:26px;font-weight:700;color:{C['blue']};line-height:1.15;}}
.kpi-value.green{{color:{C['green']};}} .kpi-value.amber{{color:{C['amber']};}}
.kpi-value.dark{{color:{C['dark_blue']};}}
.kpi-label{{font-size:12px;color:{C['dark_grey']};margin-top:5px;
  text-transform:uppercase;letter-spacing:.03em;}}
.sh{{font-size:15px;font-weight:600;color:{C['dark_blue']};
  border-bottom:2px solid {C['blue']};padding-bottom:5px;margin:6px 0 10px;}}
</style>""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def normalize(name):
    if pd.isna(name): return name
    s = str(name).strip()
    return NAME_MAP.get(s, s)

def market_from_team(team):
    """Map Concur Team string → Market."""
    if not isinstance(team, str): return "Unknown"
    return TEAM_MARKET.get(team.strip(), "Unknown")

def leader_from_team(team):
    """Map Concur Team string → FLSM name."""
    if not isinstance(team, str): return ""
    return TEAM_LEADER_MAP.get(team.strip(), "")

def director_from_leader(leader):
    """Map FLSM name → Director name."""
    if not isinstance(leader, str): return "Unknown"
    return FLSM_DIRECTOR.get(leader.strip(), "Unknown")

def ltc_rate(term):
    try:    t = int(float(term))
    except: return 0.0
    if t >= 36: return 0.40
    if t >= 24: return 0.30
    if t >= 12: return 0.20
    return 0.0

def bucket(pct):
    if pct is None or (isinstance(pct, float) and np.isnan(pct)): return "No Quota"
    if pct < 0.50:  return "<50%"
    if pct < 0.75:  return "50-75%"
    if pct < 1.00:  return "75-100%"
    if pct < 1.20:  return "100-120%"
    return "120%+"

def kpi_html(label, value, accent=""):
    cls = f"kpi-card {accent}".strip()
    vcls = f"kpi-value {accent}".strip()
    return (f"<div class='{cls}'><div class='{vcls}'>{value}</div>"
            f"<div class='kpi-label'>{label}</div></div>")

def fmt_money(v, sym="$"):
    """Format a monetary value with appropriate scale."""
    if abs(v) >= 1e6: return f"{sym}{v/1e6:.2f}M"
    if abs(v) >= 1e3: return f"{sym}{v/1e3:.0f}K"
    return f"{sym}{v:,.0f}"

def fmt_pct(v):
    if v is None or (isinstance(v, float) and np.isnan(v)): return "—"
    return f"{v*100:.1f}%"

def currency_sym(market):
    """Return the currency symbol for a given market selection."""
    if market in MARKET_CURRENCY:
        return CURRENCY_SYMBOL.get(MARKET_CURRENCY[market], "$")
    return ""   # All Markets — mixed currency, no single symbol

# ─────────────────────────────────────────────────────────────────────────────
# SALESFORCE CONNECTION
# ─────────────────────────────────────────────────────────────────────────────
def sf_connect(username, password, security_token, domain="login"):
    from simple_salesforce import Salesforce
    return Salesforce(username=username, password=password,
                      security_token=security_token, domain=domain)

def sf_connect_session(session_id, instance_url="https://sapconcur.my.salesforce.com"):
    from simple_salesforce import Salesforce
    return Salesforce(session_id=session_id, instance_url=instance_url)

_meta_cache = {}
def _sf_get_meta(sf, report_id):
    if report_id not in _meta_cache:
        _meta_cache[report_id] = sf.restful(
            path=f"analytics/reports/{report_id}",
            method="GET",
            params={"includeDetails": "true"},
        )
    return _meta_cache[report_id]

def _cell_val(cell):
    if cell is None: return None
    v = cell.get("value")
    if v is None: v = cell.get("label")
    return v

# ─────────────────────────────────────────────────────────────────────────────
# SFDC REPORT EXECUTION  (identical to US version)
# ─────────────────────────────────────────────────────────────────────────────
def _post_report(sf, report_id, body_dict, params=None):
    import json as _json
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {sf.session_id}",
    }
    url = f"{sf.base_url}analytics/reports/{report_id}"
    resp = _requests.post(url, headers=headers,
                          json=body_dict, params=params or {}, timeout=120)
    resp.raise_for_status()
    return resp.json()

def _post_report_async(sf, report_id, body_dict, params=None):
    """POST to /instances to get an async instance, then poll until done."""
    import json as _json
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {sf.session_id}",
    }
    inst_url = f"{sf.base_url}analytics/reports/{report_id}/instances"
    resp = _requests.post(inst_url, headers=headers,
                          json=body_dict, params=params or {}, timeout=120)
    resp.raise_for_status()
    inst = resp.json()
    inst_id = inst.get("id") or inst.get("reportId")
    if inst_id is None:
        raise ValueError(f"No instance id in POST /instances response: {inst}")

    # Poll
    for attempt in range(30):
        _time.sleep(1.5 if attempt < 5 else 3)
        poll = sf.restful(
            path=f"analytics/reports/{report_id}/instances/{inst_id}",
            method="GET",
            params=params or {},
        )
        if poll.get("status") == "Success":
            return poll, inst_id
        if poll.get("status") in ("Error", "Failed"):
            raise RuntimeError(f"Async report failed: {poll}")
    raise TimeoutError(f"Report instance {inst_id} did not complete in time")

def _fetch_all_pages_async(sf, report_id, inst_id, first_result, base_params):
    rows = list(first_result.get("factMap", {}).get("T!T", {}).get("rows", []))
    if first_result.get("allData", True):
        return rows
    page, MAX = 1, 50
    while page < MAX:
        pg_params = {**base_params, "startRow": page * 2000}
        try:
            r = sf.restful(
                path=f"analytics/reports/{report_id}/instances/{inst_id}",
                method="GET", params=pg_params)
            new = r.get("factMap", {}).get("T!T", {}).get("rows", [])
            if not new: break
            rows.extend(new)
            if r.get("allData", True): break
            page += 1
        except Exception: break
    return rows

def _fetch_all_pages(sf, report_id, first_result, body_dict, base_params):
    rows = list(first_result.get("factMap", {}).get("T!T", {}).get("rows", []))
    if first_result.get("allData", True):
        return rows
    page, MAX = 1, 50
    while page < MAX:
        pg_params = {**base_params, "startRow": page * 2000}
        try:
            if body_dict is not None:
                current = _post_report(sf, report_id, body_dict, pg_params)
            else:
                current = sf.restful(
                    path=f"analytics/reports/{report_id}",
                    method="GET", params=pg_params)
            combined = current.get("factMap", {}).get("T!T", {}).get("rows", [])
            if not combined: break
            rows.extend(combined)
            page += 1
        except Exception: break
    return rows

def sf_run_report(sf, report_id, start_date=None, end_date=None):
    import copy as _copy
    params = {"includeDetails": "true"}
    result = None
    debug  = []
    _winning_body    = None
    _winning_params  = params
    _winning_inst_id = None

    if "sf_post_debug" not in st.session_state:
        st.session_state.sf_post_debug = {}

    _DATE_COLS = {"CLOSE_DATE","CLOSEDATE","CLOSED_DATE","CREATEDDATE",
                  "LASTMODIFIEDDATE","CLOSE_MONTH","CLOSEMONTH"}

    if start_date and end_date:
        _a1_bucket_err = False
        try:
            _a1_body = {"reportMetadata": {"standardDateFilter": {
                "column": "CLOSE_DATE", "durationValue": "CUSTOM",
                "startDate": start_date, "endDate": end_date}}}
            r1, _a1_inst_id = _post_report_async(sf, report_id, _a1_body, params)
            n1 = len(r1.get("factMap", {}).get("T!T", {}).get("rows", []))
            all1 = r1.get("allData", True)
            result = r1
            _winning_body    = _a1_body
            _winning_inst_id = _a1_inst_id
            debug.append(f"A1 OK → {n1} rows" + ("" if all1 else " [paginating…]"))
        except Exception as e1:
            if "BucketField" in str(e1):
                _a1_bucket_err = True
                debug.append("A1 ERR: BucketField — switching to A2")
            else:
                debug.append(f"A1 ERR: {e1}")

        if result is None:
            try:
                full_resp   = _sf_get_meta(sf, report_id)
                saved_meta  = full_resp.get("reportMetadata", {})
                std_info    = saved_meta.get("standardDateFilter") or {}
                std_col     = std_info.get("column", "CLOSE_DATE")
                patched_meta = _copy.deepcopy(saved_meta)
                patched_meta["standardDateFilter"] = {
                    "column": std_col, "durationValue": "CUSTOM",
                    "startDate": start_date, "endDate": end_date}
                existing_filters = patched_meta.get("reportFilters") or []
                non_date = [f for f in existing_filters
                            if f.get("column","").upper() not in _DATE_COLS]
                patched_meta["reportFilters"] = non_date + [
                    {"column": std_col, "operator": "greaterOrEqual", "value": start_date},
                    {"column": std_col, "operator": "lessOrEqual",    "value": end_date},
                ]
                _a2_body = {"reportMetadata": patched_meta}
                r2, _a2_inst_id = _post_report_async(sf, report_id, _a2_body, params)
                n2 = len(r2.get("factMap", {}).get("T!T", {}).get("rows", []))
                result = r2
                _winning_body    = _a2_body
                _winning_inst_id = _a2_inst_id
                debug.append(f"A2 OK → {n2} rows")
            except Exception as e2:
                debug.append(f"A2 ERR: {e2}")

    if result is None:
        result = sf.restful(path=f"analytics/reports/{report_id}",
                            method="GET", params=params)
        nG = len(result.get("factMap", {}).get("T!T", {}).get("rows", []))
        _winning_body    = None
        _winning_inst_id = None
        debug.append(f"GET-fallback → {nG} rows")

    if _winning_inst_id:
        all_rows = _fetch_all_pages_async(sf, report_id, _winning_inst_id,
                                          result, _winning_params)
    else:
        all_rows = _fetch_all_pages(sf, report_id, result,
                                    _winning_body, _winning_params)

    st.session_state.sf_post_debug[report_id] = " | ".join(debug)

    meta   = result.get("reportMetadata", {})
    ext    = result.get("reportExtendedMetadata", {})
    cols   = meta.get("detailColumns", [])
    cinfo  = ext.get("detailColumnInfo", {})
    labels = [cinfo.get(c, {}).get("label", c) for c in cols]
    records = []
    for row in all_rows:
        cells = row.get("dataCells", [])
        records.append({labels[i]: _cell_val(cells[i])
                        for i in range(min(len(labels), len(cells)))})
    df = pd.DataFrame(records)
    return df, len(records)


def sf_run_report_multi(sf, report_id, month_nums, year, full_year=False):
    """Run a report across the selected period, concatenating monthly calls.

    full_year=True: single call spanning Jan 1 – Dec 31 of year (used for
    Retention, whose date field is a text picklist not filterable by API).
    """
    if isinstance(month_nums, int):
        month_nums = [month_nums]

    if full_year:
        start = f"{year}-01-01"
        end   = f"{year}-12-31"
        df, n = sf_run_report(sf, report_id, start, end)
        return df, n

    frames = []
    for m in month_nums:
        _, last_day = monthrange(year, m)
        start = f"{year}-{m:02d}-01"
        end   = f"{year}-{m:02d}-{last_day:02d}"
        try:
            df_m, _ = sf_run_report(sf, report_id, start, end)
            frames.append(df_m)
        except Exception as e:
            st.warning(f"Report {report_id} month {m} error: {e}")

    if not frames:
        return pd.DataFrame(), 0
    combined = pd.concat(frames, ignore_index=True)
    return combined, len(combined)


# ─────────────────────────────────────────────────────────────────────────────
# QUOTA LOADING  (INT — reads from local-currency quota file)
# ─────────────────────────────────────────────────────────────────────────────
def load_quotas_int(tgt_bytes, month_nums):
    """Read the INT local-currency quota file and return quota maps.

    Returns:
        ldr_quota_map  : FLSM name  → period quota (sum of reps under FLSM)
        rep_quota_map  : rep name   → period quota (PR monthly columns summed)
        rep_leader_map : rep name   → FLSM name
        rep_market_map : rep name   → market string
        market_pl_map  : market     → total period quota
    """
    if isinstance(month_nums, int):
        month_nums = [month_nums]

    buf = io.BytesIO(tgt_bytes)
    df  = pd.read_excel(buf, sheet_name="GTM Ops Plan Summary Concur")

    # Filter to INT markets and quota-carrying named reps
    df = df[df["Concur Market"].isin(["Canada", "United Kingdom", "Australia"])].copy()
    df = df[(df["Quota Carrier"] == "Yes") & df["Employee Name"].notna()].copy()

    # PR monthly columns for selected months
    month_cols = [f"PR Concur Cloud Bookings ARR 501 - {MONTH_ABBREV[m-1]}"
                  for m in month_nums]

    rep_quota_map  = {}
    rep_leader_map = {}
    rep_market_map = {}
    ldr_quota_map  = {}
    market_pl_map  = {"Canada": 0.0, "United Kingdom": 0.0, "Australia": 0.0}

    for _, row in df.iterrows():
        raw_name = str(row["Employee Name"]).strip()
        nm       = normalize(raw_name)
        team     = str(row.get("Concur Team","")).strip()
        market   = str(row.get("Concur Market","")).strip()

        # Period quota: sum PR monthly amounts (already pro-rated for mid-year)
        period_q = 0.0
        for col in month_cols:
            val = row.get(col)
            if pd.notna(val):
                try: period_q += float(val)
                except: pass

        rep_quota_map[nm]  = period_q
        rep_market_map[nm] = market

        # Leader from static TEAM_LEADER_MAP
        leader = TEAM_LEADER_MAP.get(team, "")
        rep_leader_map[nm] = leader

        # Accumulate leader and market totals
        if leader:
            ldr_quota_map[leader] = ldr_quota_map.get(leader, 0.0) + period_q
        if market in market_pl_map:
            market_pl_map[market] += period_q

    return ldr_quota_map, rep_quota_map, rep_leader_map, rep_market_map, market_pl_map


# ─────────────────────────────────────────────────────────────────────────────
# LEADER ASSIGNMENT HELPER
# ─────────────────────────────────────────────────────────────────────────────
def _assign_leader(df, team_leader_map, mgr_col, rep_leader_map):
    """4-level leader assignment.  Priority (highest → lowest):
      1. Oppty Team → team_leader_map   (INT static map)
      2. Oppty Region → team_leader_map (fallback if region matches a team key)
      3. mgr_col (Opportunity Owner: Manager / Oppty Manager)
      4. rep_leader_map from quota file
    """
    _bad = {"", "nan", "none", "unknown", "n/a"}

    def _clean(s):
        v = normalize(str(s)) if pd.notna(s) else ""
        return "" if str(v).lower().strip() in _bad else str(v).strip()

    # Base: quota file (lowest priority)
    leader = df["Opportunity Owner"].apply(
        lambda r: rep_leader_map.get(_clean(r), "Unknown"))

    # Layer 3: SFDC manager hierarchy
    if mgr_col in df.columns:
        mgr  = df[mgr_col].apply(_clean)
        mask = mgr.ne("")
        leader = leader.where(~mask, mgr)

    if team_leader_map:
        # Layer 2: Oppty Region → team_leader_map
        if "Oppty Region" in df.columns:
            reg_ldr = df["Oppty Region"].apply(
                lambda r: _clean(team_leader_map.get(str(r).strip(), "")))
            mask = reg_ldr.ne("")
            leader = leader.where(~mask, reg_ldr)

        # Layer 1: Oppty Team → team_leader_map (highest priority)
        if "Oppty Team" in df.columns:
            team_ldr = df["Oppty Team"].apply(
                lambda t: _clean(team_leader_map.get(str(t).strip(), "")))
            mask = team_ldr.ne("")
            leader = leader.where(~mask, team_ldr)

    return leader


# ─────────────────────────────────────────────────────────────────────────────
# CALCULATION ENGINE
# ─────────────────────────────────────────────────────────────────────────────
def run_calc(cw_raw, ltc_raw, ret_raw, comp_raw,
             ldr_quota_map, rep_quota_map, rep_leader_map, rep_market_map,
             market_pl_map, month_nums, year):
    """Apply the INT SMB calculation engine.

    Key INT differences vs US:
    - rep_quota_map already holds period quota (no monthly_factor multiplication)
    - Segment → Market grouping
    - Leader assignment uses TEAM_LEADER_MAP (static) as primary source
    """
    if isinstance(month_nums, int):
        month_nums = [month_nums]
    month_names = [MONTH_NAMES[m - 1] for m in month_nums]

    # ── Filter by period ─────────────────────────────────────────────────────
    cw = cw_raw.copy()
    cw["Close Date"] = pd.to_datetime(cw["Close Date"], errors="coerce")
    cw = cw[(cw["Close Date"].dt.month.isin(month_nums)) &
            (cw["Close Date"].dt.year  == year)].copy()
    cw["Opportunity Owner"] = cw["Opportunity Owner"].apply(normalize)
    mgr_col_cw = ("Oppty Manager" if "Oppty Manager" in cw.columns
                  else "Opportunity Owner: Manager")
    if mgr_col_cw in cw.columns:
        cw[mgr_col_cw] = cw[mgr_col_cw].apply(normalize)

    ltc = ltc_raw.copy()
    ltc["Close Date"] = pd.to_datetime(ltc["Close Date"], errors="coerce")
    ltc = ltc[(ltc["Close Date"].dt.month.isin(month_nums)) &
              (ltc["Close Date"].dt.year  == year)].copy()
    ltc["Opportunity Owner"] = ltc["Opportunity Owner"].apply(normalize)
    mgr_col_ltc = ("Oppty Manager" if "Oppty Manager" in ltc.columns
                   else "Opportunity Owner: Manager")
    if mgr_col_ltc in ltc.columns:
        ltc[mgr_col_ltc] = ltc[mgr_col_ltc].apply(normalize)

    # ── Retention — year-aware FMC filter ────────────────────────────────────
    ret = ret_raw.copy()
    _fmc = ret["Final Month Closed"].fillna("").astype(str).str.strip()
    _month_year_explicit = [f"{mn} {year}"       for mn in month_names]
    _month_year_short    = [f"{mn[:3]} {year}"   for mn in month_names]
    _month_year_hyphen   = [f"{mn[:3]}-{year}"   for mn in month_names]
    _month_only          = month_names
    _match_explicit = _fmc.isin(_month_year_explicit + _month_year_short +
                                 _month_year_hyphen)
    if "Close Date" in ret.columns:
        _close_yr = pd.to_datetime(ret["Close Date"], errors="coerce").dt.year
        _match_month_only = _fmc.isin(_month_only) & _close_yr.eq(year).fillna(False)
    else:
        _match_month_only = _fmc.isin(_month_only)
    ret = ret[_match_explicit | _match_month_only].copy()
    ret["Opportunity Owner"] = ret["Opportunity Owner"].apply(normalize)
    mgr_col_ret = ("Oppty Manager" if "Oppty Manager" in ret.columns
                   else "Opportunity Owner: Manager")
    if mgr_col_ret in ret.columns:
        ret[mgr_col_ret] = ret[mgr_col_ret].apply(normalize)
    ret["BMI_ARR"]     = pd.to_numeric(ret["BMI Sales ARR"], errors="coerce").fillna(0)
    # Roll-up Sales Credit — try converted column first, fall back to plain column
    _sc_col = ("Roll-up Sales Credit Calculation (converted)"
               if "Roll-up Sales Credit Calculation (converted)" in ret.columns
               else "Roll-up Sales Credit Calculation"
               if "Roll-up Sales Credit Calculation" in ret.columns
               else None)
    ret["Roll_SC_Ret"] = (pd.to_numeric(ret[_sc_col], errors="coerce").fillna(0)
                          if _sc_col else 0.0)
    # Exclude Split Opportunity rows
    _split = ret.get("ARR Disputes", pd.Series([""] * len(ret))).fillna("")
    ret = ret[~_split.str.contains("Split Opportunity", case=False, na=False)].copy()

    # ── Complete ──────────────────────────────────────────────────────────────
    comp = comp_raw.copy()
    comp["_cm_parsed"] = pd.to_datetime(
        comp["Close Month"].astype(str).str.strip(), errors="coerce", dayfirst=False)
    comp = comp[(comp["_cm_parsed"].dt.year  == year) &
                (comp["_cm_parsed"].dt.month.isin(month_nums))].copy()
    comp.drop(columns=["_cm_parsed"], inplace=True)
    comp["Opportunity Owner"] = comp["Opportunity Owner"].apply(normalize)
    _sc_comp = ("Roll-up Sales Credit Calculation (converted)"
                if "Roll-up Sales Credit Calculation (converted)" in comp.columns
                else "Roll-up Sales Credit Calculation"
                if "Roll-up Sales Credit Calculation" in comp.columns
                else None)
    comp["Complete_Credit_Val"] = (pd.to_numeric(comp[_sc_comp], errors="coerce").fillna(0)
                                   if _sc_comp else 0.0)

    # ── Lookup tables ─────────────────────────────────────────────────────────
    complete_names  = set(comp["Opportunity Name"].str.strip())
    complete_lookup = dict(zip(comp["Opportunity Name"].str.strip(),
                               comp["Complete_Credit_Val"]))

    ltc["LTC_Uplift_Calc"] = ltc.apply(
        lambda r: pd.to_numeric(r["Forecast Amount"], errors="coerce")
                  * ltc_rate(r["Term (no. of months)"]), axis=1)
    ltc_lookup = dict(zip(ltc["Opportunity Name"].str.strip(),
                          ltc["LTC_Uplift_Calc"]))

    # ── Retention grouping ───────────────────────────────────────────────────
    _team_col_ret = ("Oppty Team" if "Oppty Team" in ret.columns
                     else "Opportunity Owner" if True else None)
    ret_grp_agg = {
        "BMI_SUM":  ("BMI_ARR",            "sum"),
        "SC_SUM":   ("Roll_SC_Ret",         "sum"),
        "Rep":      ("Opportunity Owner",   "first"),
        "OppName":  ("Opportunity Name",    "first"),
    }
    if "Oppty Team" in ret.columns:
        ret_grp_agg["Team"] = ("Oppty Team", "first")
    ret_grp = ret.groupby("Opportunity ID").agg(**ret_grp_agg).reset_index()
    ret_grp["Retention_Credit"] = (ret_grp["BMI_SUM"] - ret_grp["SC_SUM"]).clip(lower=0)

    # Assign leader to retention deals
    _rg = ret_grp.rename(columns={"Team": "Oppty Team"} if "Team" in ret_grp.columns else {})
    _rg["Opportunity Owner"] = _rg["Rep"]
    if mgr_col_ret in ret.columns:
        _first_mgr = (ret.groupby("Opportunity ID")[mgr_col_ret]
                      .first().reset_index()
                      .rename(columns={mgr_col_ret: "_mgr_ret"}))
        _rg = _rg.merge(_first_mgr, on="Opportunity ID", how="left")
        _rg[mgr_col_ret] = _rg.get("_mgr_ret", "")
    else:
        _rg[mgr_col_ret] = ""
    ret_grp["Leader"] = _assign_leader(_rg, TEAM_LEADER_MAP,
                                        mgr_col_ret, rep_leader_map).values

    ret_by_oppname = ret_grp.groupby("OppName")["Retention_Credit"].sum().to_dict()

    # ── Master CW dataset ─────────────────────────────────────────────────────
    master = cw.copy()
    master["Rep"] = master["Opportunity Owner"]
    master["Leader"] = _assign_leader(master, TEAM_LEADER_MAP,
                                       mgr_col_cw, rep_leader_map)

    # Team / Market columns
    team_col = ("Oppty Team" if "Oppty Team" in master.columns
                else "Owner Team" if "Owner Team" in master.columns
                else None)
    if team_col:
        master["Team"] = master[team_col].fillna("Unknown").str.strip()
    else:
        master["Team"] = "Unknown"
    master["Market"] = master["Team"].apply(market_from_team)
    master["Director"] = master["Leader"].apply(director_from_leader)

    master["Forecast_Amount_ARR"] = pd.to_numeric(
        master["Forecast Amount"], errors="coerce").fillna(0)
    master["_OppName"]       = master["Opportunity Name"].str.strip()
    master["In_Complete"]    = master["_OppName"].isin(complete_names).astype(int)
    master["Complete_Credit"]  = master["_OppName"].map(complete_lookup).fillna(0)
    master["CW_ARR_Adjusted"]  = np.where(
        master["In_Complete"] == 1, 0, master["Forecast_Amount_ARR"])
    master["LTC_Uplift"]       = master["_OppName"].map(ltc_lookup).fillna(0)
    master["Retention_Credit"] = master["_OppName"].map(ret_by_oppname).fillna(0)
    master["Total_Credited"]   = (master["CW_ARR_Adjusted"] + master["LTC_Uplift"]
                                   + master["Retention_Credit"]
                                   + master["Complete_Credit"])

    # ── Standalone retention (opps not matched to CW) ────────────────────────
    cw_names   = set(master["_OppName"])
    total_ret  = ret_grp["Retention_Credit"].sum()

    # ── Rep aggregation ───────────────────────────────────────────────────────
    master_s = master.sort_values("Close Date", na_position="first")
    cw_by_rep = master_s.groupby("Rep").agg(
        Leader_cw       = ("Leader",          "last"),
        Team_cw         = ("Team",            "last"),
        Market_cw       = ("Market",          "last"),
        Director_cw     = ("Director",        "last"),
        CW_ARR          = ("CW_ARR_Adjusted", "sum"),
        LTC_Credit      = ("LTC_Uplift",      "sum"),
        Complete_Credit = ("Complete_Credit", "sum"),
        CW_Units        = ("Opportunity Name","count"),
    ).reset_index()

    ret_by_rep = ret_grp.groupby("Rep").agg(
        Retention_Credit = ("Retention_Credit", "sum"),
        Leader_ret       = ("Leader",           "first"),
    ).reset_index()

    rep_grp = cw_by_rep.merge(
        ret_by_rep[["Rep","Retention_Credit","Leader_ret"]],
        on="Rep", how="outer")
    for col in ["CW_ARR","LTC_Credit","Complete_Credit","CW_Units","Retention_Credit"]:
        rep_grp[col] = rep_grp[col].fillna(0)

    def resolve_leader(row):
        for src in [row.get("Leader_cw",""), row.get("Leader_ret","")]:
            if isinstance(src, str) and src.strip().lower() not in ("","nan","unknown","n/a"):
                return src.strip()
        return rep_leader_map.get(row["Rep"], "Unknown")

    def resolve_team(row):
        for src in [row.get("Team_cw","")]:
            if isinstance(src, str) and src.strip().lower() not in ("","nan","unknown"):
                return src.strip()
        return "Unknown"

    def resolve_market(row):
        t = resolve_team(row)
        m = market_from_team(t)
        if m != "Unknown": return m
        # Fall back to rep_market_map from quota file
        return rep_market_map.get(row["Rep"], "Unknown")

    rep_grp["Leader"]   = rep_grp.apply(resolve_leader,  axis=1)
    rep_grp["Team"]     = rep_grp.apply(resolve_team,    axis=1)
    rep_grp["Market"]   = rep_grp.apply(resolve_market,  axis=1)
    rep_grp["Director"] = rep_grp["Leader"].apply(director_from_leader)

    for _c in ["CW_ARR","LTC_Credit","Retention_Credit","Complete_Credit","CW_Units"]:
        rep_grp[_c] = pd.to_numeric(rep_grp[_c], errors="coerce").fillna(0)
    rep_grp["Total_Credited"] = (rep_grp["CW_ARR"] + rep_grp["LTC_Credit"]
                                  + rep_grp["Retention_Credit"]
                                  + rep_grp["Complete_Credit"])
    # Period quota: already a dollar amount (no monthly_factor needed)
    rep_grp["Period_Quota"] = pd.to_numeric(
        rep_grp["Rep"].map(lambda r: rep_quota_map.get(r, 0)),
        errors="coerce").fillna(0)
    rep_grp["Pct_to_Quota"] = np.where(
        rep_grp["Period_Quota"] > 0,
        rep_grp["Total_Credited"] / rep_grp["Period_Quota"], np.nan)
    rep_grp = rep_grp[["Rep","Leader","Director","Team","Market",
                        "CW_ARR","LTC_Credit","Retention_Credit","Complete_Credit",
                        "Total_Credited","Period_Quota","Pct_to_Quota","CW_Units"]]

    # ── Leader aggregation ────────────────────────────────────────────────────
    ldr_grp = rep_grp.groupby("Leader").agg(
        Director        = ("Director",         "first"),
        Market          = ("Market",           "first"),
        CW_ARR          = ("CW_ARR",           "sum"),
        LTC_Credit      = ("LTC_Credit",       "sum"),
        Retention_Credit= ("Retention_Credit", "sum"),
        Complete_Credit = ("Complete_Credit",  "sum"),
        Total_Credited  = ("Total_Credited",   "sum"),
        CW_Units        = ("CW_Units",         "sum"),
    ).reset_index()
    ldr_grp["Period_Quota"] = pd.to_numeric(
        ldr_grp["Leader"].map(lambda l: ldr_quota_map.get(l, 0)),
        errors="coerce").fillna(0)
    ldr_grp["Pct_to_Quota"] = np.where(
        ldr_grp["Period_Quota"] > 0,
        ldr_grp["Total_Credited"] / ldr_grp["Period_Quota"], np.nan)

    # ── Market aggregation ────────────────────────────────────────────────────
    mkt_grp = rep_grp[rep_grp["Market"].isin(["Canada","United Kingdom","Australia"])
                      ].groupby("Market").agg(
        CW_ARR          = ("CW_ARR",           "sum"),
        LTC_Credit      = ("LTC_Credit",       "sum"),
        Retention_Credit= ("Retention_Credit", "sum"),
        Complete_Credit = ("Complete_Credit",  "sum"),
        Total_Credited  = ("Total_Credited",   "sum"),
        CW_Units        = ("CW_Units",         "sum"),
    ).reset_index()
    mkt_grp["PL_Quota"] = pd.to_numeric(
        mkt_grp["Market"].map(market_pl_map), errors="coerce").fillna(0)
    mkt_grp["Pct_to_PL"] = np.where(
        mkt_grp["PL_Quota"] > 0,
        mkt_grp["Total_Credited"] / mkt_grp["PL_Quota"], np.nan)

    # ── Attainment distribution ───────────────────────────────────────────────
    rq = rep_grp[rep_grp["Period_Quota"] > 0].copy()
    rq["Bucket"] = rq["Pct_to_Quota"].apply(bucket)
    dist = rq.groupby("Bucket").size().reset_index(name="Count")

    # ── Org summary ───────────────────────────────────────────────────────────
    org_total = rep_grp["Total_Credited"].sum()
    org_quota = sum(market_pl_map.values())
    org = {
        "Total_Credited":   org_total,
        "CW_ARR":           rep_grp["CW_ARR"].sum(),
        "LTC_Credit":       rep_grp["LTC_Credit"].sum(),
        "Retention_Credit": total_ret,
        "Complete_Credit":  rep_grp["Complete_Credit"].sum(),
        "PL_Quota":         org_quota,
        "Pct_to_PL":        org_total / org_quota if org_quota else 0,
        "CW_Units":         int(len(cw)),
        "Above_100":        int((rq["Pct_to_Quota"] >= 1.0).sum()),
        "Above_120":        int((rq["Pct_to_Quota"] >= 1.2).sum()),
        "Total_Reps":       int(len(rq)),
        "Row_Counts": {"CW": len(cw), "LTC": len(ltc),
                       "Retention": len(ret), "Complete": len(comp)},
    }

    return {"org": org, "market": mkt_grp, "leader": ldr_grp,
            "rep": rep_grp, "dist": dist, "master": master}


# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE INIT
# ─────────────────────────────────────────────────────────────────────────────
for k, v in [("sf", None), ("data", None), ("last_run", None),
              ("raw_cw", None), ("raw_ltc", None), ("raw_ret", None),
              ("raw_comp", None), ("tgt_bytes", None)]:
    if k not in st.session_state:
        st.session_state[k] = v

period_label = "—"
sel_market   = "All Markets"

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### SAP Concur")
    st.markdown("**INT SMB Client Sales Dashboard**")
    st.markdown("---")

    # ── Credentials ──────────────────────────────────────────────────────────
    with st.expander("Salesforce Connection", expanded=st.session_state.sf is None):
        secrets_sf = st.secrets.get("salesforce", {}) if hasattr(st, "secrets") else {}

        auth_mode = st.radio("Auth Method",
                             ["Session ID (SSO / Microsoft login)", "Username + Password"],
                             index=0, horizontal=True)

        if auth_mode == "Session ID (SSO / Microsoft login)":
            session_id = st.text_input("Session ID", type="password",
                                       value=secrets_sf.get("session_id", ""),
                                       placeholder="00D…  (paste from browser — see instructions below)",
                                       help="Log into Salesforce in your browser → F12 → "
                                            "Application → Cookies → sapconcur.my.salesforce.com → "
                                            "copy the value of the 'sid' cookie")
            instance_url = st.text_input("Instance URL",
                                         value=secrets_sf.get("instance_url",
                                                               "https://sapconcur.my.salesforce.com"))
            col_a, col_b = st.columns(2)
            if col_a.button("Connect", use_container_width=True):
                if not session_id:
                    st.error("Paste your Session ID first.")
                else:
                    try:
                        st.session_state.sf = sf_connect_session(session_id, instance_url)
                        st.success("Connected")
                    except Exception as e:
                        st.error(str(e))
        else:
            username  = st.text_input("Username (email)",
                                      value=secrets_sf.get("username",""),
                                      placeholder="you@company.com")
            password  = st.text_input("Password", type="password",
                                      value=secrets_sf.get("password",""))
            sec_token = st.text_input("Security Token", type="password",
                                      value=secrets_sf.get("security_token",""))
            domain    = st.text_input("Domain",
                                      value=secrets_sf.get("domain","login"))
            col_a, col_b = st.columns(2)
            if col_a.button("Connect", use_container_width=True):
                try:
                    st.session_state.sf = sf_connect(username, password, sec_token, domain)
                    st.success("Connected")
                except Exception as e:
                    st.error(str(e))

        if col_b.button("Disconnect", use_container_width=True):
            st.session_state.sf = None
            st.session_state.data = None

    conn_ok = st.session_state.sf is not None
    st.markdown(
        f"<span style='color:{'#188918' if conn_ok else '#BB0000'};font-weight:600'>"
        f"{'● Connected' if conn_ok else '○ Not connected'}</span>",
        unsafe_allow_html=True)

    st.markdown("---")

    # ── Report IDs ────────────────────────────────────────────────────────────
    with st.expander("Report IDs", expanded=True):
        secrets_rpt = st.secrets.get("reports", {}) if hasattr(st, "secrets") else {}
        rpt_cw   = st.text_input("CW ARR Report ID",
                                  value=secrets_rpt.get("cw_arr_report_id",   "00OPg00000Qf2dB"),
                                  placeholder="00O…")
        rpt_ltc  = st.text_input("LTC Report ID",
                                  value=secrets_rpt.get("ltc_report_id",      "00OPg00000Qf2i1"),
                                  placeholder="00O…")
        rpt_ret  = st.text_input("Retention Report ID",
                                  value=secrets_rpt.get("retention_report_id","00OPg00000Qf39R"),
                                  placeholder="00O…")
        rpt_comp = st.text_input("Complete / Referral Report ID",
                                  value=secrets_rpt.get("complete_report_id", "00OPg00000Qf2mr"),
                                  placeholder="00O…")

    st.markdown("---")

    # ── Reporting Period ──────────────────────────────────────────────────────
    st.markdown("**Reporting Period**")
    period_type = st.radio("Period Type", ["Monthly","Quarterly","YTD"],
                           horizontal=True)
    sel_year = st.number_input("Year", min_value=2020, max_value=2030,
                                value=2026, step=1)
    QUARTER_MONTHS = {"Q1":[1,2,3],"Q2":[4,5,6],"Q3":[7,8,9],"Q4":[10,11,12]}

    if period_type == "Monthly":
        sel_month    = st.selectbox("Month", MONTH_NAMES,
                                    index=MONTH_NAMES.index("July"))
        month_nums   = [MONTH_NAMES.index(sel_month) + 1]
        period_label = f"{sel_month} {int(sel_year)}"
    elif period_type == "Quarterly":
        sel_quarter  = st.selectbox("Quarter", ["Q1","Q2","Q3","Q4"], index=2)
        month_nums   = QUARTER_MONTHS[sel_quarter]
        period_label = f"{sel_quarter} {int(sel_year)}"
    else:
        sel_month    = st.selectbox("Through Month", MONTH_NAMES,
                                    index=MONTH_NAMES.index("July"))
        month_nums   = list(range(1, MONTH_NAMES.index(sel_month) + 2))
        period_label = f"YTD through {sel_month} {int(sel_year)}"

    st.markdown("---")

    # ── Quota file ────────────────────────────────────────────────────────────
    st.markdown("**Quota File**")
    tgt_upload = st.file_uploader("Upload quotas .xlsx", type=["xlsx"],
                                   help="2026 SMB quotas in local currency … .xlsx")
    if tgt_upload:
        st.session_state.tgt_bytes = tgt_upload.read()
    elif DEFAULT_TARGETS.exists() and st.session_state.tgt_bytes is None:
        st.session_state.tgt_bytes = DEFAULT_TARGETS.read_bytes()
        st.caption(f"Using: {DEFAULT_TARGETS.name}")

    st.markdown("---")

    # ── Run Reports ───────────────────────────────────────────────────────────
    run_btn = st.button("Run Reports", type="primary",
                        use_container_width=True, disabled=not conn_ok)
    if run_btn:
        if not all([rpt_cw, rpt_ltc, rpt_ret, rpt_comp]):
            st.error("All 4 Report IDs are required.")
        elif st.session_state.tgt_bytes is None:
            st.error("Please upload the Quota file.")
        else:
            n_months = len(month_nums)
            with st.spinner(f"Running Salesforce reports… "
                            f"({n_months} month{'s' if n_months>1 else ''})"):
                try:
                    sf   = st.session_state.sf
                    _yr  = int(sel_year)
                    cw_df,  n1 = sf_run_report_multi(sf, rpt_cw,   month_nums, _yr)
                    ltc_df, n2 = sf_run_report_multi(sf, rpt_ltc,  month_nums, _yr)
                    ret_df, n3 = sf_run_report_multi(sf, rpt_ret,  month_nums, _yr,
                                                     full_year=True)
                    comp_df,n4 = sf_run_report_multi(sf, rpt_comp, month_nums, _yr)
                    st.session_state.raw_cw   = cw_df
                    st.session_state.raw_ltc  = ltc_df
                    st.session_state.raw_ret  = ret_df
                    st.session_state.raw_comp = comp_df
                    st.caption(f"CW ARR: {n1} rows | LTC: {n2} rows | "
                               f"Retention: {n3} rows | Referral: {n4} rows")
                    for name, n in [("CW ARR",n1),("LTC",n2),
                                    ("Retention",n3),("Referral",n4)]:
                        if n == 0:
                            st.warning(f"**{name} returned 0 rows.** "
                                       f"Check the report date filter in Salesforce.")
                        elif n >= 2000:
                            st.warning(f"**{name} returned 2,000 rows** — API cap. "
                                       f"Some records may be missing.")
                except Exception as e:
                    st.error(f"Report error: {e}")
                    st.stop()

            with st.spinner("Calculating metrics…"):
                try:
                    lq, rq, rlm, rmm, mpl = load_quotas_int(
                        st.session_state.tgt_bytes, month_nums)
                    result = run_calc(
                        st.session_state.raw_cw,  st.session_state.raw_ltc,
                        st.session_state.raw_ret, st.session_state.raw_comp,
                        lq, rq, rlm, rmm, mpl, month_nums, int(sel_year))
                    st.session_state.data     = result
                    st.session_state.last_run = datetime.now()
                except Exception as e:
                    st.error(f"Calculation error: {e}\n\n"
                             f"Check the column preview in the diagnostic expander.")
                    st.stop()
            st.rerun()

    # ── Filters (shown after data loads) ─────────────────────────────────────
    if st.session_state.data:
        st.markdown("---")
        st.markdown("**Filters**")
        rep_df_all = st.session_state.data["rep"]

        mkts_avail = ["All Markets"] + sorted(
            rep_df_all[rep_df_all["Market"].isin(
                ["Canada","United Kingdom","Australia"])]["Market"].unique())
        sel_market = st.selectbox("Market", mkts_avail)

        _rep_mkt = (rep_df_all[rep_df_all["Market"] == sel_market]
                    if sel_market != "All Markets" else rep_df_all)
        teams_avail = ["All Teams"] + sorted([
            t for t in _rep_mkt["Team"].unique()
            if isinstance(t, str) and t.strip() not in ("", "Unknown", "nan")
        ])
        sel_team = st.selectbox("Team", teams_avail)

        _rep_team = (_rep_mkt[_rep_mkt["Team"] == sel_team]
                     if sel_team != "All Teams" else _rep_mkt)
        ldrs_avail = ["All Leaders"] + sorted(
            _rep_team["Leader"].dropna().unique().tolist())
        sel_ldr = st.selectbox("Leader", ldrs_avail)
    else:
        sel_market = "All Markets"
        sel_team   = "All Teams"
        sel_ldr    = "All Leaders"

    if st.session_state.last_run:
        st.caption(f"Last run: {st.session_state.last_run.strftime('%b %d %Y %H:%M')}")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN — SPLASH
# ─────────────────────────────────────────────────────────────────────────────
if st.session_state.data is None:
    st.markdown(
        f"<div style='text-align:center;padding:80px 0'>"
        f"<h1 style='color:{C['blue']};font-size:2rem;margin-bottom:.5rem'>"
        f"SAP Concur INT SMB Client Sales Dashboard</h1>"
        f"<p style='color:{C['dark_grey']};font-size:1rem'>"
        f"Markets: Canada (CAD) &nbsp;·&nbsp; United Kingdom (GBP) &nbsp;·&nbsp; Australia (AUD)</p>"
        f"<p style='color:{C['dark_grey']};font-size:.95rem;margin-top:20px'>"
        f"Connect to Salesforce and run reports using the sidebar to begin.</p>"
        f"</div>", unsafe_allow_html=True)
    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# DATA + FILTER HELPERS
# ─────────────────────────────────────────────────────────────────────────────
d        = st.session_state.data
org_d    = d["org"]
rep_df   = d["rep"].copy()
ldr_df   = d["leader"].copy()
mkt_df   = d["market"].copy()
dist_df  = d["dist"].copy()

# Apply market / team / leader filters
if sel_market != "All Markets":
    rep_df  = rep_df[rep_df["Market"] == sel_market].copy()
    ldr_df  = ldr_df[ldr_df["Market"] == sel_market].copy()
    mkt_df  = mkt_df[mkt_df["Market"] == sel_market].copy()

if "sel_team" in dir() and sel_team != "All Teams":
    rep_df = rep_df[rep_df["Team"] == sel_team].copy()
    ldr_df = ldr_df[ldr_df["Leader"].isin(
        [TEAM_LEADER_MAP.get(sel_team, "")])].copy()

if "sel_ldr" in dir() and sel_ldr != "All Leaders":
    rep_df = rep_df[rep_df["Leader"] == sel_ldr].copy()
    ldr_df = ldr_df[ldr_df["Leader"] == sel_ldr].copy()

# Determine currency symbol for KPI cards
curr_sym = currency_sym(sel_market)  # empty string for "All Markets"
fmt_m = lambda v: fmt_money(v, curr_sym) if curr_sym else fmt_money(v, "")

# Filtered totals
f_total    = rep_df["Total_Credited"].sum()
f_quota    = rep_df["Period_Quota"].sum()
f_cw       = rep_df["CW_ARR"].sum()
f_ltc      = rep_df["LTC_Credit"].sum()
f_ret      = rep_df["Retention_Credit"].sum()
f_comp     = rep_df["Complete_Credit"].sum()
f_units    = int(rep_df["CW_Units"].sum())
f_pct      = f_total / f_quota if f_quota else 0.0

# Rep attainment distribution (filtered)
rq_filt    = rep_df[rep_df["Period_Quota"] > 0].copy()
rq_filt["Bucket"] = rq_filt["Pct_to_Quota"].apply(bucket)
dist_filt  = rq_filt.groupby("Bucket").size().reset_index(name="Count")
above_100  = int((rq_filt["Pct_to_Quota"] >= 1.0).sum())
above_120  = int((rq_filt["Pct_to_Quota"] >= 1.2).sum())
total_reps = len(rq_filt)


# ─────────────────────────────────────────────────────────────────────────────
# CURRENCY WARNING (shown for "All Markets" view)
# ─────────────────────────────────────────────────────────────────────────────
CURRENCY_WARNING_REPS = {
    "Chelsea Salonek":  ("Canada",          "CA SMB CS Premier",    "USD",  "CAD"),
    "Loreena Maguet":   ("United Kingdom",  "UK Mid Market",        "EUR",  "GBP"),
    "Alan Donohoe":     ("United Kingdom",  "UK SMB CS Premier",    "EUR",  "GBP"),
    "Lydia Holloway":   ("United Kingdom",  "UK SMB CS Premier",    "AED",  "GBP"),
}

if sel_market == "All Markets":
    st.warning(
        "**Mixed-currency view:** Amounts for each market are in their local currency "
        "(CAD, GBP, AUD). Cross-market totals shown here are **not directly comparable** "
        "and should not be aggregated as USD equivalents. "
        "Select a single market in the sidebar for a clean single-currency view.")

# Flag mismatched-currency reps if any are in current filtered view
_visible_reps = set(rep_df["Rep"].tolist())
_warn_visible = {r: v for r, v in CURRENCY_WARNING_REPS.items()
                 if r in _visible_reps}
if _warn_visible:
    _lines = " &nbsp;|&nbsp; ".join(
        f"<b>{r}</b> ({v[0]}, {v[1]}): quota is in {v[2]} — should be {v[3]}"
        for r, v in _warn_visible.items())
    st.error(
        f"**Quota Currency Mismatch** — the following reps have incorrect quota currency "
        f"in Salesforce and may skew market totals. Fix needed in SFDC.<br>{_lines}",
        icon="⚠")


# ─────────────────────────────────────────────────────────────────────────────
# TITLE ROW
# ─────────────────────────────────────────────────────────────────────────────
mkt_label = (f" — {sel_market} ({MARKET_CURRENCY.get(sel_market,'')})"
             if sel_market != "All Markets" else " — All Markets (CAD · GBP · AUD)")
st.markdown(
    f"<h2 style='color:{C['dark_blue']};margin-bottom:4px'>"
    f"INT SMB Client Sales{mkt_label}</h2>"
    f"<p style='color:{C['dark_grey']};margin-top:0'>{period_label}</p>",
    unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# ROW 1 — KPI CARDS
# ─────────────────────────────────────────────────────────────────────────────
c1, c2, c3, c4, c5, c6 = st.columns(6)
pct_color = ("green" if f_pct >= 1.0 else "amber" if f_pct >= 0.75 else "")

c1.markdown(kpi_html("Total Credited",    fmt_m(f_total),   pct_color), unsafe_allow_html=True)
c2.markdown(kpi_html("Period Quota",      fmt_m(f_quota),   "dark"),    unsafe_allow_html=True)
c3.markdown(kpi_html("% to Quota",        fmt_pct(f_pct),   pct_color), unsafe_allow_html=True)
c4.markdown(kpi_html("CW Units",          str(f_units),     ""),        unsafe_allow_html=True)
c5.markdown(kpi_html("Reps ≥ 100%",
                      f"{above_100} / {total_reps}",          "green"),  unsafe_allow_html=True)
c6.markdown(kpi_html("Reps ≥ 120%",
                      f"{above_120} / {total_reps}",          "green"),  unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# ROW 2 — MARKET PERFORMANCE + CREDIT BREAKDOWN
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<div class='sh'>Market Performance</div>", unsafe_allow_html=True)
col_mkt, col_pie = st.columns([3, 2])

with col_mkt:
    if len(mkt_df) == 0:
        st.info("No market data for the current filter.")
    else:
        fig_mkt = go.Figure()
        # Stacked bars: CW ARR, LTC, Retention, Complete
        components = [
            ("CW ARR",         "CW_ARR",           C["blue"]),
            ("LTC Uplift",     "LTC_Credit",        C["light_blue"]),
            ("Retention",      "Retention_Credit",  C["dark_blue"]),
            ("Complete/Ref",   "Complete_Credit",   C["dark_grey"]),
        ]
        for lbl, col_name, clr in components:
            fig_mkt.add_trace(go.Bar(
                name=lbl, x=mkt_df["Market"], y=mkt_df[col_name],
                marker_color=clr, text=mkt_df[col_name].apply(
                    lambda v: f"{curr_sym}{v/1e3:.0f}K" if v >= 1000 else ""),
                textposition="inside"))

        # Quota line
        fig_mkt.add_trace(go.Scatter(
            name="Quota", x=mkt_df["Market"], y=mkt_df["PL_Quota"],
            mode="markers+lines",
            marker=dict(size=10, symbol="diamond", color=C["amber"]),
            line=dict(color=C["amber"], width=2, dash="dot")))

        fig_mkt.update_layout(
            barmode="stack", height=300, margin=dict(l=0, r=0, t=30, b=0),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
            xaxis_title="", yaxis_title="",
            plot_bgcolor=C["bg"], paper_bgcolor=C["bg"],
            font=dict(family="72,Helvetica Neue,Arial,sans-serif", size=12))
        fig_mkt.update_yaxes(tickprefix=curr_sym if curr_sym else "",
                              showgrid=True, gridcolor=C["grey"])
        st.plotly_chart(fig_mkt, use_container_width=True)

with col_pie:
    labels_pie = ["CW ARR", "LTC Uplift", "Retention", "Complete/Ref"]
    vals_pie   = [f_cw, f_ltc, f_ret, f_comp]
    vals_pie   = [max(0, v) for v in vals_pie]
    if sum(vals_pie) > 0:
        fig_pie = go.Figure(go.Pie(
            labels=labels_pie, values=vals_pie, hole=0.55,
            marker_colors=[C["blue"], C["light_blue"], C["dark_blue"], C["dark_grey"]],
            textinfo="percent", textfont_size=11))
        fig_pie.update_layout(
            height=300, margin=dict(l=0, r=0, t=30, b=0),
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.15, x=0.5,
                        xanchor="center", font=dict(size=11)),
            plot_bgcolor=C["bg"], paper_bgcolor=C["bg"],
            annotations=[dict(text=f"{fmt_m(f_total)}<br><span style='font-size:9px'>"
                                   f"Total</span>",
                              font_size=13, showarrow=False)])
        st.plotly_chart(fig_pie, use_container_width=True)
    else:
        st.info("No credited amounts for current filter.")


# ─────────────────────────────────────────────────────────────────────────────
# ROW 3 — LEADER ATTAINMENT + DISTRIBUTION
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<div class='sh'>Leader Attainment</div>", unsafe_allow_html=True)
col_ldr, col_dist = st.columns([3, 2])

with col_ldr:
    ldr_show = ldr_df[ldr_df["Leader"].ne("Unknown")].copy()
    if len(ldr_show) == 0:
        st.info("No leader data for the current filter.")
    else:
        ldr_show = ldr_show.sort_values("Total_Credited", ascending=True)
        ldr_show["Bar_Color"] = ldr_show["Market"].map(
            lambda m: MARKET_COLORS.get(m, C["blue"]))
        ldr_show["Pct_Label"] = ldr_show["Pct_to_Quota"].apply(
            lambda p: f"  {p*100:.0f}%" if pd.notna(p) else "")
        fig_ldr = go.Figure()
        for mkt, clr in MARKET_COLORS.items():
            sub = ldr_show[ldr_show["Market"] == mkt]
            if len(sub) == 0: continue
            fig_ldr.add_trace(go.Bar(
                name=mkt, y=sub["Leader"], x=sub["Total_Credited"],
                orientation="h", marker_color=clr,
                text=sub["Pct_Label"], textposition="outside"))
        # Quota markers
        ldr_q = ldr_show[ldr_show["Period_Quota"] > 0]
        fig_ldr.add_trace(go.Scatter(
            name="Quota", y=ldr_q["Leader"], x=ldr_q["Period_Quota"],
            mode="markers",
            marker=dict(size=8, symbol="line-ns", color=C["amber"],
                        line=dict(width=2, color=C["amber"]))))
        fig_ldr.update_layout(
            barmode="overlay", height=max(280, len(ldr_show) * 38),
            margin=dict(l=0, r=60, t=20, b=0),
            legend=dict(orientation="h", y=1.06, x=0),
            xaxis_title="", yaxis_title="",
            plot_bgcolor=C["bg"], paper_bgcolor=C["bg"],
            font=dict(family="72,Helvetica Neue,Arial,sans-serif", size=12))
        fig_ldr.update_xaxes(tickprefix=curr_sym if curr_sym else "",
                              showgrid=True, gridcolor=C["grey"])
        st.plotly_chart(fig_ldr, use_container_width=True)

with col_dist:
    if len(dist_filt) > 0:
        dist_sorted = pd.DataFrame({"Bucket": BUCKET_ORDER}).merge(
            dist_filt, on="Bucket", how="left").fillna(0)
        dist_sorted["Count"] = dist_sorted["Count"].astype(int)
        fig_dist = go.Figure(go.Bar(
            x=dist_sorted["Bucket"], y=dist_sorted["Count"],
            marker_color=[BUCKET_COLORS.get(b, C["grey"]) for b in dist_sorted["Bucket"]],
            text=dist_sorted["Count"], textposition="outside"))
        fig_dist.update_layout(
            height=280, margin=dict(l=0, r=0, t=30, b=0),
            xaxis_title="Attainment Band", yaxis_title="Reps",
            plot_bgcolor=C["bg"], paper_bgcolor=C["bg"],
            font=dict(family="72,Helvetica Neue,Arial,sans-serif", size=12))
        fig_dist.update_yaxes(showgrid=True, gridcolor=C["grey"])
        st.plotly_chart(fig_dist, use_container_width=True)
    else:
        st.info("No quota-carrying reps in current filter.")


# ─────────────────────────────────────────────────────────────────────────────
# ROW 4 — REP PERFORMANCE TABLE
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<div class='sh'>Rep Performance Detail</div>", unsafe_allow_html=True)

rep_show = rep_df.copy()
rep_show = rep_show.sort_values("Total_Credited", ascending=False)
_pct_disp = rep_show["Pct_to_Quota"].apply(fmt_pct)

rep_disp = pd.DataFrame({
    "Rep":              rep_show["Rep"],
    "Market":           rep_show["Market"],
    "Team":             rep_show["Team"],
    "Leader":           rep_show["Leader"],
    "CW ARR":           rep_show["CW_ARR"].round(0),
    "LTC":              rep_show["LTC_Credit"].round(0),
    "Retention":        rep_show["Retention_Credit"].round(0),
    "Complete/Ref":     rep_show["Complete_Credit"].round(0),
    "Total Credited":   rep_show["Total_Credited"].round(0),
    "Period Quota":     rep_show["Period_Quota"].round(0),
    "% to Quota":       _pct_disp,
    "CW Units":         rep_show["CW_Units"].astype(int),
})

def _style_row(row):
    """Colour-code % to Quota column."""
    styles = [""] * len(row)
    idx = rep_disp.columns.get_loc("% to Quota")
    pct_str = row["% to Quota"]
    try:
        pct = float(pct_str.replace("%","")) / 100
        if pct >= 1.0:     clr = "#E8F5E9"
        elif pct >= 0.75:  clr = "#FFF8E1"
        elif pct >= 0.50:  clr = "#FFF3E0"
        else:              clr = "#FFEBEE"
        styles[idx] = f"background-color:{clr}"
    except: pass
    return styles

styled = rep_disp.style.apply(_style_row, axis=1)
st.dataframe(styled, use_container_width=True, height=420, hide_index=True)


# ─────────────────────────────────────────────────────────────────────────────
# ROW 5 — MARKET DETAIL TABLE
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<div class='sh'>Market Summary</div>", unsafe_allow_html=True)

mkt_disp = mkt_df.copy() if len(mkt_df) > 0 else pd.DataFrame(
    columns=["Market","CW_ARR","LTC_Credit","Retention_Credit",
             "Complete_Credit","Total_Credited","PL_Quota","Pct_to_PL","CW_Units"])

mkt_table = pd.DataFrame({
    "Market":           mkt_disp["Market"],
    "Currency":         mkt_disp["Market"].map(MARKET_CURRENCY).fillna(""),
    "CW ARR":           mkt_disp["CW_ARR"].round(0),
    "LTC":              mkt_disp["LTC_Credit"].round(0),
    "Retention":        mkt_disp["Retention_Credit"].round(0),
    "Complete/Ref":     mkt_disp["Complete_Credit"].round(0),
    "Total Credited":   mkt_disp["Total_Credited"].round(0),
    "PL Quota":         mkt_disp["PL_Quota"].round(0),
    "% to Plan":        mkt_disp["Pct_to_PL"].apply(fmt_pct),
    "CW Units":         mkt_disp["CW_Units"].astype(int) if len(mkt_disp) else pd.Series(dtype=int),
})
st.dataframe(mkt_table, use_container_width=True, height=180, hide_index=True)


# ─────────────────────────────────────────────────────────────────────────────
# ROW 6 — EXCEL DOWNLOAD
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("<div class='sh'>Export</div>", unsafe_allow_html=True)

@st.cache_data(show_spinner=False)
def build_excel(rep_df_json, ldr_df_json, mkt_df_json, period_lbl):
    rep_df_e  = pd.read_json(rep_df_json,  orient="records")
    ldr_df_e  = pd.read_json(ldr_df_json,  orient="records")
    mkt_df_e  = pd.read_json(mkt_df_json,  orient="records")

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        rep_df_e.to_excel(writer,  sheet_name="Rep Detail",    index=False)
        ldr_df_e.to_excel(writer,  sheet_name="Leader Summary",index=False)
        mkt_df_e.to_excel(writer,  sheet_name="Market Summary",index=False)

        # Pivot: rep x month (if month_cols available)
        meta_df = pd.DataFrame({"Metric": ["Period", "Generated"],
                                 "Value":  [period_lbl,
                                            datetime.now().strftime("%Y-%m-%d %H:%M")]})
        meta_df.to_excel(writer, sheet_name="Metadata", index=False)

    return buf.getvalue()

col_dl, col_debug = st.columns([2, 3])
with col_dl:
    try:
        xl_bytes = build_excel(
            rep_df.to_json(orient="records"),
            ldr_df.to_json(orient="records"),
            mkt_df.to_json(orient="records"),
            period_label)
        safe_period = re.sub(r"[^A-Za-z0-9_\-]", "_", period_label)
        mkt_slug = re.sub(r"\s+", "_", sel_market.replace("/","_"))
        fname = f"INT_SMB_{mkt_slug}_{safe_period}.xlsx"
        st.download_button(
            label="Download Excel Report",
            data=xl_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    except Exception as e:
        st.error(f"Excel build error: {e}")

with col_debug:
    with st.expander("Diagnostics & Debug"):
        st.markdown(f"**Org row counts:** CW={org_d['Row_Counts']['CW']} | "
                    f"LTC={org_d['Row_Counts']['LTC']} | "
                    f"Retention={org_d['Row_Counts']['Retention']} | "
                    f"Complete={org_d['Row_Counts']['Complete']}")
        st.markdown(f"**Filtered reps:** {len(rep_df)} | "
                    f"**With quota:** {total_reps}")

        # Raw SFDC column names for CW
        if st.session_state.raw_cw is not None and len(st.session_state.raw_cw):
            st.markdown("**CW ARR columns:**")
            st.code(", ".join(st.session_state.raw_cw.columns.tolist()))

        # Report-level debug
        if "sf_post_debug" in st.session_state:
            for rid, msg in st.session_state.sf_post_debug.items():
                st.caption(f"{rid}: {msg}")

        # Currency warning detail
        st.markdown("**Quota currency outliers (fix in SFDC):**")
        for rep, (mkt, team, has_cur, should_cur) in CURRENCY_WARNING_REPS.items():
            st.caption(f"• {rep} ({mkt} / {team}): quota is {has_cur}, expected {should_cur}")
