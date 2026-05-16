
"""analyzer_core.py

Core analysis functions for Tech Efficiency Analyzer.

Supports:
- Column normalization via aliases
- Individual and team rollups
- Optional store/CID rollups when present
- Efficiency scoring + tiering
- Trend outputs when Date present (weekly/daily)
- Excel report generation with executive summary + charts

Designed to run in:
- Replit (Streamlit app)
- CLI scripts

"""

from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd
import re


# -----------------------------
# Config + column normalization
# -----------------------------

DEFAULT_COLUMN_ALIASES: Dict[str, List[str]] = {
    # identifiers
    "technician": ["technician", "tech", "installer", "contractor", "name", "tech name", "technician name"],
    "technician_team": ["technician team", "team", "region", "market", "group"],
    "role": ["role", "position", "type"],
    "date": ["date", "work date", "day", "service date", "report date"],
    # metrics
    "hours": ["hours", "hrs", "labor hours", "worked hours", "time", "total hours"],
    "units": ["units", "builds", "installs", "completed", "completed jobs", "jobs", "jobs completed", "qty"],
    "amount": ["amount", "revenue", "invoice amount", "sales", "gross", "total amount", "dollars"],
    "mileage": ["mileage", "miles", "mileage paid", "miles paid"],
    # optional store identifiers
    "cid": ["cid", "store", "store #", "store number", "location id", "site id"],
    "store_name": ["store name", "location", "site", "store location", "account"],
}

# Canonical columns we want to have
CANON = ["Technician Team", "Technician", "Role", "Date", "Hours", "Units", "Amount", "Mileage", "CID", "Store Name"]


@dataclass
class AnalyzerConfig:
    # weights for efficiency score components
    w_units_per_hr: float = 0.4
    w_revenue_per_job: float = 0.3
    w_revenue_per_hr: float = 0.3

    # tier cutoffs (percentiles; higher is better)
    tier_elite: float = 0.90
    tier_strong: float = 0.70
    tier_average: float = 0.30

    # mileage cost per mile (only used if Mileage exists and represents miles; if Mileage Paid exists, treat as cost)
    mileage_cost_per_mile: float = 0.0

    # output naming
    output_filename: str = "Tech_Efficiency_Report.xlsx"

    # column aliases
    column_aliases: Dict[str, List[str]] = None

    @staticmethod
    def from_json(path: Path) -> "AnalyzerConfig":
        cfg = AnalyzerConfig()
        data = {}
        try:
            if path.exists():
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f) or {}
        except Exception:
            data = {}

        for k, v in data.items():
            if hasattr(cfg, k):
                setattr(cfg, k, v)

        if cfg.column_aliases is None:
            cfg.column_aliases = DEFAULT_COLUMN_ALIASES
        return cfg


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip().lower())


def normalize_columns(df: pd.DataFrame, aliases: Dict[str, List[str]] = None) -> pd.DataFrame:
    """Rename columns in df to canonical names where possible.

    This is resilient to drift in report headers.
    """
    import re  # local import for lint friendliness
    if aliases is None:
        aliases = DEFAULT_COLUMN_ALIASES

    # build map of normalized existing columns -> actual
    existing = { _norm(c): c for c in df.columns }

    # map canonical target -> matched existing col
    mapping = {}

    def try_match(target_key: str, canon_name: str):
        for a in aliases.get(target_key, []):
            an = _norm(a)
            if an in existing:
                mapping[existing[an]] = canon_name
                return

    try_match("technician_team", "Technician Team")
    try_match("technician", "Technician")
    try_match("role", "Role")
    try_match("date", "Date")
    try_match("hours", "Hours")
    try_match("units", "Units")
    try_match("amount", "Amount")
    try_match("mileage", "Mileage")
    try_match("cid", "CID")
    try_match("store_name", "Store Name")

    df = df.rename(columns=mapping)

    # Ensure missing canon columns exist (filled with NaNs / defaults)
    for col in CANON:
        if col not in df.columns:
            df[col] = np.nan

    # Clean numeric columns
    for col in ["Hours", "Units", "Amount", "Mileage"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    # Clean date
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    # Clean identifiers
    for col in ["Technician Team", "Technician", "Role", "CID", "Store Name"]:
        if col in df.columns:
            df[col] = df[col].astype(str).replace({"nan": np.nan, "None": np.nan}).fillna("")

    return df


def load_excel_files(file_paths: Iterable[str]) -> pd.DataFrame:
    frames = []
    for p in file_paths:
        df = pd.read_excel(p, header=0)
        df["__source_file__"] = Path(p).name
        frames.append(df)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


# -----------------------------
# Metrics + scoring
# -----------------------------

def _safe_div(n: pd.Series, d: pd.Series) -> pd.Series:
    d2 = d.replace(0, np.nan)
    return (n / d2).replace([np.inf, -np.inf], np.nan)


def compute_individual_stats(df: pd.DataFrame, cfg: AnalyzerConfig) -> pd.DataFrame:
    agg = {
        "Date": "nunique",
        "Hours": "sum",
        "Units": "sum",
        "Amount": "sum",
        "Mileage": "sum",
    }
    out = (
        df.groupby(["Technician Team", "Technician", "Role"], dropna=False)
          .agg(agg)
          .reset_index()
          .rename(columns={"Date": "Days Worked", "Amount": "Revenue"})
    )

    out["Units/Hr"] = _safe_div(out["Units"], out["Hours"]).round(2)
    out["Revenue/Job"] = _safe_div(out["Revenue"], out["Units"]).round(2)
    out["Gross $/Hr"] = _safe_div(out["Revenue"], out["Hours"]).round(2)

    # Mileage cost handling (best-effort)
    if cfg.mileage_cost_per_mile and cfg.mileage_cost_per_mile > 0:
        out["Mileage Paid"] = (out["Mileage"] * cfg.mileage_cost_per_mile).round(2)
    else:
        # If Mileage looks like it is already paid cost (common), keep as-is but name it mileage paid
        out["Mileage Paid"] = out["Mileage"].round(2)

    out["Net Revenue"] = (out["Revenue"] - out["Mileage Paid"]).round(2)
    out["Net $/Hr"] = _safe_div(out["Net Revenue"], out["Hours"]).round(2)
    out["Mileage Cost %"] = (_safe_div(out["Mileage Paid"], out["Revenue"]) * 100).round(1)

    out["Avg Hrs/Day"] = _safe_div(out["Hours"], out["Days Worked"]).round(1)

    # Consistency Score (Std Dev of Daily Gross $/Hr)
    if "Date" in df.columns and not df["Date"].isna().all():
        tmp_cons = df.copy()
        tmp_cons["Date"] = pd.to_datetime(tmp_cons["Date"], errors="coerce")
        tmp_cons = tmp_cons.dropna(subset=["Date"])
        daily = tmp_cons.groupby(["Technician", "Date"]).agg({"Amount": "sum", "Hours": "sum"}).reset_index()
        daily["Gross $/Hr"] = _safe_div(daily["Amount"], daily["Hours"])
        consistency = daily.groupby("Technician")["Gross $/Hr"].std().fillna(0).round(2).reset_index(name="Consistency Score")
        out = out.merge(consistency, on="Technician", how="left")
        out["Consistency Score"] = out["Consistency Score"].fillna(0)
    else:
        out["Consistency Score"] = 0.0

    # Flight Risk (Recent 2 weeks avg vs historical avg)
    if "Date" in df.columns and not df["Date"].isna().all():
        tmp_flight = df.copy()
        tmp_flight["Date"] = pd.to_datetime(tmp_flight["Date"], errors="coerce")
        tmp_flight = tmp_flight.dropna(subset=["Date"])
        max_date = tmp_flight["Date"].max()
        cutoff_date = max_date - pd.Timedelta(days=14)

        tmp_flight["Week Start"] = tmp_flight["Date"] - pd.to_timedelta(tmp_flight["Date"].dt.dayofweek, unit="D")
        weekly_hrs = tmp_flight.groupby(["Technician", "Week Start"])["Hours"].sum().reset_index()

        hist_avg = weekly_hrs.groupby("Technician")["Hours"].mean().reset_index(name="Hist Avg Hrs")

        recent_mask = weekly_hrs["Week Start"] >= (cutoff_date - pd.to_timedelta(cutoff_date.dayofweek, unit="D"))
        recent_avg = weekly_hrs[recent_mask].groupby("Technician")["Hours"].mean().reset_index(name="Recent Avg Hrs")

        risk_df = hist_avg.merge(recent_avg, on="Technician", how="left").fillna(0)
        risk_df["Flight Risk Ratio"] = _safe_div(risk_df["Recent Avg Hrs"], risk_df["Hist Avg Hrs"]).fillna(0)

        def assign_risk(ratio):
            if ratio < 0.5:
                return "High Risk (Recent drop)"
            return "Low Risk"

        risk_df["Flight Risk"] = risk_df["Flight Risk Ratio"].apply(assign_risk)
        out = out.merge(risk_df[["Technician", "Flight Risk"]], on="Technician", how="left")
        out["Flight Risk"] = out["Flight Risk"].fillna("Unknown")
    else:
        out["Flight Risk"] = "Unknown"

    scored = add_efficiency_score(out, cfg)
    return scored


def add_efficiency_score(individual_df: pd.DataFrame, cfg: AnalyzerConfig) -> pd.DataFrame:
    df = individual_df.copy()

    # Normalize components to 0-1 via percentile ranks (robust to outliers)
    def pr(s: pd.Series) -> pd.Series:
        return s.rank(pct=True, method="average").fillna(0.0)

    df["_p_units_hr"] = pr(df["Units/Hr"])
    df["_p_rev_job"] = pr(df["Revenue/Job"])
    df["_p_gross_hr"] = pr(df["Gross $/Hr"])

    df["Efficiency Score"] = (
        df["_p_units_hr"] * cfg.w_units_per_hr
        + df["_p_rev_job"] * cfg.w_revenue_per_job
        + df["_p_gross_hr"] * cfg.w_revenue_per_hr
    ).round(4)

    df["Rank"] = df["Efficiency Score"].rank(ascending=False, method="min").astype(int)
    df["Percentile"] = df["Efficiency Score"].rank(pct=True).round(4)

    # Tiering
    def tier(p):
        if p >= cfg.tier_elite:
            return "Elite"
        if p >= cfg.tier_strong:
            return "Strong"
        if p >= cfg.tier_average:
            return "Average"
        return "Underperformer"

    df["Tier"] = df["Percentile"].apply(tier)

    # Coaching Profile
    def coaching_profile(row):
        p_units = row["_p_units_hr"]
        p_rev = row["_p_rev_job"]
        mileage_pct = row.get("Mileage Cost %", 0)

        if p_units > 0.8 and p_rev > 0.8:
            return "Mentor Candidate"
        if p_units < 0.5 and p_rev > 0.5:
            return "Needs Speed/Efficiency Training"
        if p_units > 0.5 and p_rev < 0.5:
            return "Needs Upsell/Value Training"
        if mileage_pct > 15:
            return "Needs Route Optimization Coaching"
        return "On Track"

    df["Coaching Profile"] = df.apply(coaching_profile, axis=1)

    # cleanup temp cols
    df = df.drop(columns=["_p_units_hr", "_p_rev_job", "_p_gross_hr"], errors="ignore")
    return df


def compute_tenure_stats(df: pd.DataFrame, core_threshold: int = 60) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Calculate tenure statistics by team.
    
    Args:
        df: Normalized dataframe with Technician Team, Technician, Date columns
        core_threshold: Days worked to qualify as a "Core Tech" (default 60)
    
    Returns:
        team_tenure: Team-level tenure stats (Core %, Avg Days, etc.)
        tech_tenure: Individual tech tenure (days worked per tech)
    """
    # Days worked per technician
    tech_tenure = df.groupby(['Technician Team', 'Technician']).agg({
        'Date': 'nunique'
    }).reset_index()
    tech_tenure.columns = ['Team', 'Technician', 'Days Worked']
    
    # Aggregate to team level
    def team_tenure_agg(x):
        return pd.Series({
            'Core Techs': (x['Days Worked'] >= core_threshold).sum(),
            'Total Techs': len(x),
            'Avg Days Worked': x['Days Worked'].mean()
        })
    
    team_tenure = tech_tenure.groupby('Team').apply(team_tenure_agg).reset_index()
    team_tenure['Core %'] = (team_tenure['Core Techs'] / team_tenure['Total Techs'] * 100).round(1)
    team_tenure['Avg Days Worked'] = team_tenure['Avg Days Worked'].round(1)
    
    return team_tenure, tech_tenure


def compute_diagnostic_summary(team_stats: pd.DataFrame, team_tenure: pd.DataFrame) -> pd.DataFrame:
    """Create diagnostic quadrant analysis for teams.
    
    Combines efficiency ranking with tenure ranking to diagnose team issues:
    - Strong Team: High tenure + High efficiency
    - Slow Techs: High tenure + Low efficiency (coaching opportunity)
    - Turnover Risk: Low tenure + High efficiency (retention issue)
    - Needs Attention: Low tenure + Low efficiency (rebuild needed)
    """
    # Merge stats and tenure
    merged = team_stats.merge(
        team_tenure[['Team', 'Core %', 'Avg Days Worked', 'Core Techs', 'Total Techs']], 
        left_on='Technician Team', 
        right_on='Team', 
        how='left'
    )
    merged = merged.drop(columns=['Team'], errors='ignore')
    
    # Add efficiency ranking (by Gross $/Hr)
    merged = merged.sort_values('Gross $/Hr', ascending=False).reset_index(drop=True)
    merged['Efficiency Rank'] = range(1, len(merged) + 1)
    
    # Add tenure ranking (by Core %)
    merged = merged.sort_values('Core %', ascending=False).reset_index(drop=True)
    merged['Tenure Rank'] = range(1, len(merged) + 1)
    
    # Calculate medians for diagnosis
    median_eff_rank = len(merged) / 2
    median_tenure_rank = len(merged) / 2
    
    # Diagnose each team
    def diagnose(row):
        high_tenure = row['Tenure Rank'] <= median_tenure_rank
        high_eff = row['Efficiency Rank'] <= median_eff_rank
        
        if high_tenure and high_eff:
            return "✓ Strong Team"
        elif high_tenure and not high_eff:
            return "⚠ Slow Techs (Coach/Train)"
        elif not high_tenure and high_eff:
            return "⚠ Turnover Risk (Retain)"
        else:
            return "✗ Needs Attention (Rebuild)"
    
    merged['Diagnosis'] = merged.apply(diagnose, axis=1)
    
    # Reorder columns for clarity
    priority_cols = ['Technician Team', 'Diagnosis', 'Efficiency Rank', 'Tenure Rank', 
                     'Gross $/Hr', 'Net $/Hr', 'Core %', 'Core Techs', 'Total Techs',
                     'Avg Days Worked', 'Tech Count', 'Hours', 'Units', 'Revenue']
    other_cols = [c for c in merged.columns if c not in priority_cols]
    merged = merged[[c for c in priority_cols if c in merged.columns] + other_cols]
    
    return merged.sort_values('Diagnosis')


def compute_mileage_analysis(team_stats: pd.DataFrame) -> pd.DataFrame:
    """Create standalone mileage analysis sheet."""
    mileage_df = team_stats[['Technician Team', 'Tech Count', 'Hours', 'Revenue', 
                              'Mileage Paid', 'Mileage Cost %', 'Gross $/Hr', 'Net $/Hr']].copy()
    mileage_df = mileage_df.sort_values('Mileage Paid', ascending=False).reset_index(drop=True)
    
    # Add miles per hour estimate (assuming ~$0.67/mile)
    mileage_df['Est Miles'] = (mileage_df['Mileage Paid'] / 0.67).round(0)
    mileage_df['Miles/Hour'] = _safe_div(mileage_df['Est Miles'], mileage_df['Hours']).round(1)
    
    # Add rank
    mileage_df.insert(0, 'Rank', range(1, len(mileage_df) + 1))
    
    return mileage_df


def compute_team_stats(df: pd.DataFrame, cfg: AnalyzerConfig) -> pd.DataFrame:
    agg = {
        "Technician": "nunique",
        "Role": "nunique",
        "Date": "nunique",
        "Hours": "sum",
        "Units": "sum",
        "Amount": "sum",
        "Mileage": "sum",
    }
    out = (
        df.groupby(["Technician Team"], dropna=False)
          .agg(agg)
          .reset_index()
          .rename(columns={"Technician": "Tech Count", "Role": "Role Types", "Date": "Days", "Amount": "Revenue"})
    )

    out["Units/Hr"] = _safe_div(out["Units"], out["Hours"]).round(2)
    out["Gross $/Hr"] = _safe_div(out["Revenue"], out["Hours"]).round(2)

    if cfg.mileage_cost_per_mile and cfg.mileage_cost_per_mile > 0:
        out["Mileage Paid"] = (out["Mileage"] * cfg.mileage_cost_per_mile).round(2)
    else:
        out["Mileage Paid"] = out["Mileage"].round(2)

    out["Net Revenue"] = (out["Revenue"] - out["Mileage Paid"]).round(2)
    out["Net $/Hr"] = _safe_div(out["Net Revenue"], out["Hours"]).round(2)
    out["Mileage Cost %"] = (_safe_div(out["Mileage Paid"], out["Revenue"]) * 100).round(1)

    return out.sort_values("Gross $/Hr", ascending=False)


def compute_store_stats(df: pd.DataFrame) -> Optional[pd.DataFrame]:
    # Only if CID or Store Name present and non-empty
    if ("CID" not in df.columns and "Store Name" not in df.columns):
        return None
    has_any = (df.get("CID", pd.Series([""])).astype(str).str.strip() != "").any() or (
        df.get("Store Name", pd.Series([""])).astype(str).str.strip() != ""
    ).any()
    if not has_any:
        return None

    group_cols = []
    if "CID" in df.columns:
        group_cols.append("CID")
    if "Store Name" in df.columns:
        group_cols.append("Store Name")
    if not group_cols:
        return None

    agg = {"Date": "nunique", "Hours": "sum", "Units": "sum", "Amount": "sum"}
    out = (
        df.groupby(group_cols, dropna=False)
          .agg(agg)
          .reset_index()
          .rename(columns={"Date": "Days", "Amount": "Revenue"})
    )
    out["Units/Hr"] = _safe_div(out["Units"], out["Hours"]).round(2)
    out["Gross $/Hr"] = _safe_div(out["Revenue"], out["Hours"]).round(2)
    return out.sort_values("Revenue", ascending=False)


def compute_trends(df: pd.DataFrame) -> Optional[pd.DataFrame]:
    # Requires a valid Date column
    if "Date" not in df.columns:
        return None
    if df["Date"].isna().all():
        return None

    tmp = df.copy()
    tmp["Date"] = pd.to_datetime(tmp["Date"], errors="coerce")
    tmp = tmp.dropna(subset=["Date"])
    if tmp.empty:
        return None

    # daily trend by technician (can be aggregated in pivots later)
    agg = {"Hours": "sum", "Units": "sum", "Amount": "sum"}
    trend = (
        tmp.groupby(["Date", "Technician Team", "Technician"], dropna=False)
           .agg(agg)
           .reset_index()
           .rename(columns={"Amount": "Revenue"})
    )
    trend["Units/Hr"] = _safe_div(trend["Units"], trend["Hours"]).round(2)
    trend["Gross $/Hr"] = _safe_div(trend["Revenue"], trend["Hours"]).round(2)
    return trend.sort_values(["Date", "Gross $/Hr"], ascending=[True, False])


def compute_weekly_staffing(df: pd.DataFrame) -> Optional[Tuple[pd.DataFrame, pd.DataFrame]]:
    """Compute unique techs working per week (Mon-Sun) by team.
    
    Returns:
        weekly_by_team: DataFrame with columns [Week Start, Team, Unique Techs, Hours, Units, Revenue]
        weekly_pivot: Pivot table with teams as rows, weeks as columns, values = unique tech count
    """
    if "Date" not in df.columns:
        return None
    if df["Date"].isna().all():
        return None
    
    tmp = df.copy()
    tmp["Date"] = pd.to_datetime(tmp["Date"], errors="coerce")
    tmp = tmp.dropna(subset=["Date"])
    if tmp.empty:
        return None
    
    # Create week start (Monday) column
    tmp["Week Start"] = tmp["Date"] - pd.to_timedelta(tmp["Date"].dt.dayofweek, unit="D")
    tmp["Week Start"] = tmp["Week Start"].dt.date
    
    # Aggregate by week and team
    weekly_by_team = tmp.groupby(["Week Start", "Technician Team"]).agg({
        "Technician": "nunique",
        "Hours": "sum",
        "Units": "sum",
        "Amount": "sum"
    }).reset_index()
    weekly_by_team.columns = ["Week Start", "Team", "Unique Techs", "Hours", "Units", "Revenue"]
    weekly_by_team["Avg Hrs/Tech"] = _safe_div(weekly_by_team["Hours"], weekly_by_team["Unique Techs"]).round(1)
    weekly_by_team["Units/Tech"] = _safe_div(weekly_by_team["Units"], weekly_by_team["Unique Techs"]).round(1)
    weekly_by_team = weekly_by_team.sort_values(["Team", "Week Start"])
    
    # Create pivot table for visualization (teams as rows, weeks as columns)
    weekly_pivot = weekly_by_team.pivot_table(
        index="Team",
        columns="Week Start",
        values="Unique Techs",
        aggfunc="sum",
        fill_value=0
    )
    
    # Calculate trend metrics per team
    team_trends = []
    for team in weekly_by_team["Team"].unique():
        team_data = weekly_by_team[weekly_by_team["Team"] == team].sort_values("Week Start")
        if len(team_data) >= 2:
            first_weeks = team_data.head(4)["Unique Techs"].mean()
            last_weeks = team_data.tail(4)["Unique Techs"].mean()
            change = last_weeks - first_weeks
            pct_change = (change / first_weeks * 100) if first_weeks > 0 else 0
            trend = "📈 Growing" if change > 0.5 else ("📉 Shrinking" if change < -0.5 else "➡️ Stable")
        else:
            first_weeks = last_weeks = team_data["Unique Techs"].mean()
            change = pct_change = 0
            trend = "➡️ Stable"
        
        team_trends.append({
            "Team": team,
            "Avg Techs (First 4 wks)": round(first_weeks, 1),
            "Avg Techs (Last 4 wks)": round(last_weeks, 1),
            "Change": round(change, 1),
            "% Change": round(pct_change, 1),
            "Trend": trend
        })
    
    staffing_trends = pd.DataFrame(team_trends).sort_values("Change", ascending=False)
    
    return weekly_by_team, weekly_pivot, staffing_trends


def compute_weekly_staffing_totals(df: pd.DataFrame) -> Optional[pd.DataFrame]:
    """Compute total unique techs working per week across all teams."""
    if "Date" not in df.columns:
        return None
    if df["Date"].isna().all():
        return None
    
    tmp = df.copy()
    tmp["Date"] = pd.to_datetime(tmp["Date"], errors="coerce")
    tmp = tmp.dropna(subset=["Date"])
    if tmp.empty:
        return None
    
    # Create week start (Monday) column
    tmp["Week Start"] = tmp["Date"] - pd.to_timedelta(tmp["Date"].dt.dayofweek, unit="D")
    tmp["Week Start"] = tmp["Week Start"].dt.date
    
    # Aggregate by week (total across all teams)
    weekly_totals = tmp.groupby("Week Start").agg({
        "Technician": "nunique",
        "Technician Team": "nunique",
        "Hours": "sum",
        "Units": "sum",
        "Amount": "sum"
    }).reset_index()
    weekly_totals.columns = ["Week Start", "Unique Techs", "Active Teams", "Hours", "Units", "Revenue"]
    weekly_totals["Avg Hrs/Tech"] = _safe_div(weekly_totals["Hours"], weekly_totals["Unique Techs"]).round(1)
    weekly_totals = weekly_totals.sort_values("Week Start")
    
    return weekly_totals


def compute_capacity_analysis(df: pd.DataFrame) -> Optional[pd.DataFrame]:
    """Analyze if teams are properly staffed for their workload.
    
    Compares output per tech across teams to identify:
    - Understaffed teams (high output per tech = techs overworked)
    - Overstaffed teams (low output per tech = techs underutilized)
    """
    if "Date" not in df.columns:
        return None
    
    tmp = df.copy()
    tmp["Date"] = pd.to_datetime(tmp["Date"], errors="coerce")
    tmp = tmp.dropna(subset=["Date"])
    if tmp.empty:
        return None
    
    # Create week start column
    tmp["Week Start"] = tmp["Date"] - pd.to_timedelta(tmp["Date"].dt.dayofweek, unit="D")
    tmp["Week Start"] = tmp["Week Start"].dt.date
    
    # Weekly stats by team
    weekly_team = tmp.groupby(["Week Start", "Technician Team"]).agg({
        "Technician": "nunique",
        "Hours": "sum",
        "Units": "sum",
        "Amount": "sum"
    }).reset_index()
    weekly_team.columns = ["Week Start", "Team", "Techs", "Hours", "Units", "Revenue"]
    
    # Calculate per-tech metrics
    weekly_team["Units/Tech"] = _safe_div(weekly_team["Units"], weekly_team["Techs"]).round(1)
    weekly_team["Revenue/Tech"] = _safe_div(weekly_team["Revenue"], weekly_team["Techs"]).round(0)
    weekly_team["Hours/Tech"] = _safe_div(weekly_team["Hours"], weekly_team["Techs"]).round(1)
    
    # Territory-wide benchmarks per week
    weekly_benchmarks = weekly_team.groupby("Week Start").agg({
        "Units/Tech": "median",
        "Revenue/Tech": "median",
        "Hours/Tech": "median"
    }).reset_index()
    weekly_benchmarks.columns = ["Week Start", "Benchmark Units/Tech", "Benchmark Rev/Tech", "Benchmark Hrs/Tech"]
    
    # Merge benchmarks back
    weekly_team = weekly_team.merge(weekly_benchmarks, on="Week Start", how="left")
    
    # Calculate variance from benchmark
    weekly_team["Units/Tech vs Benchmark"] = ((weekly_team["Units/Tech"] / weekly_team["Benchmark Units/Tech"] - 1) * 100).round(1)
    weekly_team["Rev/Tech vs Benchmark"] = ((weekly_team["Revenue/Tech"] / weekly_team["Benchmark Rev/Tech"] - 1) * 100).round(1)
    
    # Aggregate to team level (average across all weeks)
    team_capacity = weekly_team.groupby("Team").agg({
        "Techs": "mean",
        "Hours": "sum",
        "Units": "sum",
        "Revenue": "sum",
        "Units/Tech": "mean",
        "Revenue/Tech": "mean",
        "Hours/Tech": "mean",
        "Units/Tech vs Benchmark": "mean",
        "Rev/Tech vs Benchmark": "mean",
        "Benchmark Units/Tech": "mean"
    }).reset_index()
    
    team_capacity["Avg Techs/Week"] = team_capacity["Techs"].round(1)
    team_capacity["Avg Units/Tech/Week"] = team_capacity["Units/Tech"].round(1)
    team_capacity["Avg Rev/Tech/Week"] = team_capacity["Revenue/Tech"].round(0)
    team_capacity["Avg Hrs/Tech/Week"] = team_capacity["Hours/Tech"].round(1)
    team_capacity["Units/Tech vs Avg %"] = team_capacity["Units/Tech vs Benchmark"].round(1)
    team_capacity["Rev/Tech vs Avg %"] = team_capacity["Rev/Tech vs Benchmark"].round(1)
    
    # Target Techs calculations
    team_capacity["Target Techs"] = _safe_div(team_capacity["Units"], team_capacity["Benchmark Units/Tech"]).round(1)
    team_capacity["Techs to Hire/Transfer"] = (team_capacity["Target Techs"] - team_capacity["Avg Techs/Week"]).round(1)

    # Diagnose staffing
    def diagnose_capacity(row):
        units_var = row["Units/Tech vs Avg %"]
        rev_var = row["Rev/Tech vs Avg %"]
        avg_var = (units_var + rev_var) / 2
        
        if avg_var > 20:
            return "🔴 Understaffed (High output/tech)"
        elif avg_var > 10:
            return "🟡 Possibly Understaffed"
        elif avg_var < -20:
            return "🔴 Overstaffed (Low output/tech)"
        elif avg_var < -10:
            return "🟡 Possibly Overstaffed"
        else:
            return "🟢 Right-sized"
    
    team_capacity["Staffing Assessment"] = team_capacity.apply(diagnose_capacity, axis=1)
    
    # Clean up columns for output
    output_cols = ["Team", "Avg Techs/Week", "Target Techs", "Techs to Hire/Transfer",
                   "Avg Units/Tech/Week", "Avg Rev/Tech/Week", "Avg Hrs/Tech/Week",
                   "Units/Tech vs Avg %", "Rev/Tech vs Avg %",
                   "Staffing Assessment", "Hours", "Units", "Revenue"]
    team_capacity = team_capacity[output_cols].sort_values("Units/Tech vs Avg %", ascending=False)
    
    return team_capacity, weekly_team


def compute_seasonality_analysis(df: pd.DataFrame, cfg: AnalyzerConfig) -> Optional[Dict[str, pd.DataFrame]]:
    """Analyze how seasonality affects workload and efficiency.
    
    Returns dict with:
    - monthly_stats: Efficiency and volume by month
    - weekly_stats: Efficiency and volume by week
    - seasonality_summary: Which months/periods are high/low
    """
    if "Date" not in df.columns:
        return None
    
    tmp = df.copy()
    tmp["Date"] = pd.to_datetime(tmp["Date"], errors="coerce")
    tmp = tmp.dropna(subset=["Date"])
    if tmp.empty:
        return None
    
    # Add time period columns
    tmp["Week Start"] = tmp["Date"] - pd.to_timedelta(tmp["Date"].dt.dayofweek, unit="D")
    tmp["Week Start"] = tmp["Week Start"].dt.date
    tmp["Month"] = tmp["Date"].dt.to_period("M")
    tmp["Month Name"] = tmp["Date"].dt.strftime("%Y-%m")
    
    # Weekly efficiency stats
    weekly_eff = tmp.groupby("Week Start").agg({
        "Technician": "nunique",
        "Hours": "sum",
        "Units": "sum",
        "Amount": "sum",
        "Mileage": "sum"
    }).reset_index()
    weekly_eff.columns = ["Week Start", "Techs", "Hours", "Units", "Revenue", "Mileage"]
    weekly_eff["Units/Hr"] = _safe_div(weekly_eff["Units"], weekly_eff["Hours"]).round(2)
    weekly_eff["Gross $/Hr"] = _safe_div(weekly_eff["Revenue"], weekly_eff["Hours"]).round(2)
    weekly_eff["Net Revenue"] = weekly_eff["Revenue"] - weekly_eff["Mileage"]
    weekly_eff["Net $/Hr"] = _safe_div(weekly_eff["Net Revenue"], weekly_eff["Hours"]).round(2)
    weekly_eff["Units/Tech"] = _safe_div(weekly_eff["Units"], weekly_eff["Techs"]).round(1)
    weekly_eff["Revenue/Tech"] = _safe_div(weekly_eff["Revenue"], weekly_eff["Techs"]).round(0)
    
    # Calculate weekly index (compared to average)
    avg_units = weekly_eff["Units"].mean()
    avg_revenue = weekly_eff["Revenue"].mean()
    avg_efficiency = weekly_eff["Gross $/Hr"].mean()
    
    weekly_eff["Volume Index"] = ((weekly_eff["Units"] / avg_units) * 100).round(1)
    weekly_eff["Revenue Index"] = ((weekly_eff["Revenue"] / avg_revenue) * 100).round(1)
    weekly_eff["Efficiency Index"] = ((weekly_eff["Gross $/Hr"] / avg_efficiency) * 100).round(1)
    
    # Monthly stats
    monthly_eff = tmp.groupby("Month Name").agg({
        "Technician": "nunique",
        "Hours": "sum",
        "Units": "sum",
        "Amount": "sum",
        "Mileage": "sum"
    }).reset_index()
    monthly_eff.columns = ["Month", "Techs", "Hours", "Units", "Revenue", "Mileage"]
    monthly_eff["Units/Hr"] = _safe_div(monthly_eff["Units"], monthly_eff["Hours"]).round(2)
    monthly_eff["Gross $/Hr"] = _safe_div(monthly_eff["Revenue"], monthly_eff["Hours"]).round(2)
    monthly_eff["Net Revenue"] = monthly_eff["Revenue"] - monthly_eff["Mileage"]
    monthly_eff["Net $/Hr"] = _safe_div(monthly_eff["Net Revenue"], monthly_eff["Hours"]).round(2)
    monthly_eff["Units/Tech"] = _safe_div(monthly_eff["Units"], monthly_eff["Techs"]).round(1)
    
    # Monthly indices
    monthly_avg_units = monthly_eff["Units"].mean()
    monthly_avg_revenue = monthly_eff["Revenue"].mean()
    monthly_avg_efficiency = monthly_eff["Gross $/Hr"].mean()
    
    monthly_eff["Volume Index"] = ((monthly_eff["Units"] / monthly_avg_units) * 100).round(1)
    monthly_eff["Revenue Index"] = ((monthly_eff["Revenue"] / monthly_avg_revenue) * 100).round(1)
    monthly_eff["Efficiency Index"] = ((monthly_eff["Gross $/Hr"] / monthly_avg_efficiency) * 100).round(1)
    
    # Seasonality classification
    def classify_period(row):
        vol_idx = row["Volume Index"]
        eff_idx = row["Efficiency Index"]
        
        if vol_idx > 115 and eff_idx > 105:
            return "🔥 Peak Season (High Vol + High Eff)"
        elif vol_idx > 115 and eff_idx < 95:
            return "⚠️ Surge Period (High Vol, Eff Drops)"
        elif vol_idx > 115:
            return "📈 High Volume"
        elif vol_idx < 85 and eff_idx > 105:
            return "💪 Slow but Efficient"
        elif vol_idx < 85 and eff_idx < 95:
            return "📉 Off Season (Low Vol + Low Eff)"
        elif vol_idx < 85:
            return "📉 Low Volume"
        elif eff_idx > 110:
            return "⭐ High Efficiency Period"
        elif eff_idx < 90:
            return "⚠️ Low Efficiency Period"
        else:
            return "➡️ Normal"
    
    weekly_eff["Period Type"] = weekly_eff.apply(classify_period, axis=1)
    monthly_eff["Period Type"] = monthly_eff.apply(classify_period, axis=1)
    
    # Correlation analysis: Does higher volume hurt efficiency?
    if len(weekly_eff) >= 4:
        vol_eff_corr = weekly_eff["Units"].corr(weekly_eff["Gross $/Hr"])
        techs_eff_corr = weekly_eff["Techs"].corr(weekly_eff["Gross $/Hr"])
    else:
        vol_eff_corr = techs_eff_corr = np.nan
    
    # Summary insights
    summary_data = [
        ["Metric", "Value", "Interpretation"],
        ["Avg Weekly Units", f"{avg_units:,.0f}", "Baseline workload"],
        ["Avg Weekly Revenue", f"${avg_revenue:,.0f}", "Baseline revenue"],
        ["Avg Gross $/Hr", f"${avg_efficiency:.2f}", "Baseline efficiency"],
        ["", "", ""],
        ["Peak Volume Week", weekly_eff.loc[weekly_eff["Units"].idxmax(), "Week Start"], 
         f"{weekly_eff['Units'].max():,.0f} units"],
        ["Lowest Volume Week", weekly_eff.loc[weekly_eff["Units"].idxmin(), "Week Start"],
         f"{weekly_eff['Units'].min():,.0f} units"],
        ["Peak Efficiency Week", weekly_eff.loc[weekly_eff["Gross $/Hr"].idxmax(), "Week Start"],
         f"${weekly_eff['Gross $/Hr'].max():.2f}/hr"],
        ["Lowest Efficiency Week", weekly_eff.loc[weekly_eff["Gross $/Hr"].idxmin(), "Week Start"],
         f"${weekly_eff['Gross $/Hr'].min():.2f}/hr"],
        ["", "", ""],
        ["Volume-Efficiency Correlation", f"{vol_eff_corr:.2f}" if pd.notna(vol_eff_corr) else "N/A",
         "Negative = efficiency drops when busy" if pd.notna(vol_eff_corr) and vol_eff_corr < -0.3 else 
         ("Positive = efficiency improves when busy" if pd.notna(vol_eff_corr) and vol_eff_corr > 0.3 else "Weak relationship")],
        ["Staffing-Efficiency Correlation", f"{techs_eff_corr:.2f}" if pd.notna(techs_eff_corr) else "N/A",
         "Negative = more techs hurts efficiency" if pd.notna(techs_eff_corr) and techs_eff_corr < -0.3 else 
         ("Positive = more techs helps efficiency" if pd.notna(techs_eff_corr) and techs_eff_corr > 0.3 else "Weak relationship")],
    ]
    summary = pd.DataFrame(summary_data[1:], columns=summary_data[0])
    
    return {
        "weekly": weekly_eff.sort_values("Week Start"),
        "monthly": monthly_eff.sort_values("Month"),
        "summary": summary,
        "vol_eff_correlation": vol_eff_corr,
        "techs_eff_correlation": techs_eff_corr
    }


# -----------------------------
# Excel report generation
# -----------------------------

def write_excel_report(
    output_path: str,
    raw_df: pd.DataFrame,
    individual: pd.DataFrame,
    team: pd.DataFrame,
    store: Optional[pd.DataFrame] = None,
    trends: Optional[pd.DataFrame] = None,
    team_tenure: Optional[pd.DataFrame] = None,
    diagnostic: Optional[pd.DataFrame] = None,
    mileage_analysis: Optional[pd.DataFrame] = None,
    weekly_staffing: Optional[pd.DataFrame] = None,
    weekly_pivot: Optional[pd.DataFrame] = None,
    staffing_trends: Optional[pd.DataFrame] = None,
    weekly_totals: Optional[pd.DataFrame] = None,
    capacity_analysis: Optional[pd.DataFrame] = None,
    capacity_weekly: Optional[pd.DataFrame] = None,
    seasonality: Optional[Dict[str, Any]] = None,
) -> str:
    """Write a multi-sheet Excel report with an executive summary + charts."""
    from openpyxl import load_workbook
    from openpyxl.chart import BarChart, Reference, ScatterChart, Series, LineChart
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.worksheet.worksheet import Worksheet

    # Create ranked versions
    gross_eff_ranked = team.sort_values('Gross $/Hr', ascending=False).reset_index(drop=True)
    gross_eff_ranked.insert(0, 'Rank', range(1, len(gross_eff_ranked) + 1))
    
    net_eff_ranked = team.sort_values('Net $/Hr', ascending=False).reset_index(drop=True)
    net_eff_ranked.insert(0, 'Rank', range(1, len(net_eff_ranked) + 1))
    
    tenure_ranked = None
    if team_tenure is not None:
        tenure_ranked = team_tenure.sort_values('Core %', ascending=False).reset_index(drop=True)
        tenure_ranked.insert(0, 'Rank', range(1, len(tenure_ranked) + 1))

    # Create legend
    legend_data = [
        ["Metric", "Definition"],
        ["Gross $/Hr", "Total Revenue / Total Hours (before mileage)"],
        ["Net $/Hr", "(Revenue - Mileage Paid) / Total Hours (TRUE efficiency)"],
        ["Units/Hr", "Total Units / Total Hours"],
        ["Revenue/Job", "Total Revenue / Total Units"],
        ["Efficiency Score", "Weighted composite: 40% Units/Hr + 30% Revenue/Job + 30% Gross $/Hr (percentile-based)"],
        ["Core Tech", "Technician with 60+ days worked in the period"],
        ["Core %", "Percentage of team that are Core Techs"],
        ["Mileage Cost %", "Mileage Paid / Revenue * 100"],
        ["Miles/Hour", "Estimated miles driven per hour worked"],
        ["Unique Techs", "Count of distinct technicians who worked that week"],
        ["Avg Hrs/Tech", "Total Hours / Unique Techs"],
        ["Units/Tech", "Total Units / Unique Techs"],
        ["Revenue/Tech", "Total Revenue / Unique Techs"],
        ["Volume Index", "Week's units as % of average (100 = average)"],
        ["Efficiency Index", "Week's $/Hr as % of average (100 = average)"],
        ["Target Techs", "Number of techs required to match the benchmark units per tech"],
        ["Techs to Hire/Transfer", "Target Techs minus Average Techs. Positive means hire, negative means transfer/reduce."],
        ["Consistency Score", "Standard deviation of a technician's daily Gross $/Hr (lower means more consistent)"],
        ["Flight Risk", "High Risk if recent weeks' average hours are less than 50% of historical average"],
        ["Coaching Profile", "Data-driven training recommendation based on Units/Hr, Rev/Job, and Mileage"],
        ["", ""],
        ["Tier", "Meaning"],
        ["Elite", "Top 10% by Efficiency Score"],
        ["Strong", "Top 30% by Efficiency Score"],
        ["Average", "Middle 40% by Efficiency Score"],
        ["Underperformer", "Bottom 30% by Efficiency Score"],
        ["", ""],
        ["Diagnosis", "Meaning"],
        ["✓ Strong Team", "High tenure + High efficiency = Keep doing what you're doing"],
        ["⚠ Slow Techs (Coach/Train)", "High tenure + Low efficiency = Training/coaching opportunity"],
        ["⚠ Turnover Risk (Retain)", "Low tenure + High efficiency = Good techs leaving, retention issue"],
        ["✗ Needs Attention (Rebuild)", "Low tenure + Low efficiency = Fundamental rebuild needed"],
        ["", ""],
        ["Staffing Trend", "Meaning"],
        ["📈 Growing", "Team added techs over the period (last 4 weeks avg > first 4 weeks avg)"],
        ["📉 Shrinking", "Team lost techs over the period"],
        ["➡️ Stable", "Headcount roughly unchanged"],
        ["", ""],
        ["Staffing Assessment", "Meaning"],
        ["🔴 Understaffed", "Output per tech >20% above average - techs overworked"],
        ["🟡 Possibly Understaffed", "Output per tech 10-20% above average"],
        ["🟢 Right-sized", "Output per tech within ±10% of average"],
        ["🟡 Possibly Overstaffed", "Output per tech 10-20% below average"],
        ["🔴 Overstaffed", "Output per tech >20% below average - techs underutilized"],
        ["", ""],
        ["Seasonality Period", "Meaning"],
        ["🔥 Peak Season", "High volume + High efficiency - best weeks"],
        ["⚠️ Surge Period", "High volume but efficiency drops - capacity strained"],
        ["📈 High Volume", "Above average workload"],
        ["💪 Slow but Efficient", "Low volume but high efficiency"],
        ["📉 Off Season", "Low volume + Low efficiency"],
        ["⭐ High Efficiency", "Efficiency well above average"],
        ["", ""],
        ["Ranking Note", "Rankings are relative to other teams in your territory"],
        ["", "Teams in top half = 'High', bottom half = 'Low'"],
    ]
    legend = pd.DataFrame(legend_data[1:], columns=legend_data[0])

    # First write with pandas
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        individual.to_excel(writer, index=False, sheet_name="Tech Leaderboard")
        gross_eff_ranked.to_excel(writer, index=False, sheet_name="Ranked by Gross Efficiency")
        net_eff_ranked.to_excel(writer, index=False, sheet_name="Ranked by Net Efficiency")
        if tenure_ranked is not None:
            tenure_ranked.to_excel(writer, index=False, sheet_name="Ranked by Tenure")
        if diagnostic is not None:
            diagnostic.to_excel(writer, index=False, sheet_name="Diagnostic Summary")
        if mileage_analysis is not None:
            mileage_analysis.to_excel(writer, index=False, sheet_name="Mileage Analysis")
        
        # Capacity analysis sheets
        if capacity_analysis is not None:
            capacity_analysis.to_excel(writer, index=False, sheet_name="Capacity Analysis")
        
        # Seasonality sheets
        if seasonality is not None:
            seasonality["summary"].to_excel(writer, index=False, sheet_name="Seasonality Summary")
            seasonality["monthly"].to_excel(writer, index=False, sheet_name="Monthly Seasonality")
            seasonality["weekly"].to_excel(writer, index=False, sheet_name="Weekly Seasonality")
        
        # Weekly staffing sheets
        if staffing_trends is not None:
            staffing_trends.to_excel(writer, index=False, sheet_name="Staffing Trends")
        if weekly_totals is not None:
            weekly_totals.to_excel(writer, index=False, sheet_name="Weekly Totals")
        if weekly_staffing is not None:
            weekly_staffing.to_excel(writer, index=False, sheet_name="Weekly by Team")
        if weekly_pivot is not None:
            weekly_pivot.to_excel(writer, sheet_name="Staffing Pivot")
        
        team.to_excel(writer, index=False, sheet_name="Team Summary")
        if store is not None:
            store.to_excel(writer, index=False, sheet_name="Store Summary")
        if trends is not None:
            trends.to_excel(writer, index=False, sheet_name="Trends (Daily)")
        legend.to_excel(writer, index=False, sheet_name="Legend")
        # Keep a raw tab for auditing
        raw_df.head(200000).to_excel(writer, index=False, sheet_name="Raw (First 200k)")

    wb = load_workbook(output_path)

    def autosize(ws: Worksheet, max_width: int = 45):
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value is None:
                    continue
                max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = min(max_len + 2, max_width)

    # Executive summary
    summary = wb.create_sheet("Executive Summary", 0)
    bold = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="1F4E79")  # dark-ish blue
    header_font = Font(bold=True, color="FFFFFF")

    # KPIs
    total_revenue = float(pd.to_numeric(raw_df.get("Amount", pd.Series([0])), errors="coerce").fillna(0).sum())
    total_units = float(pd.to_numeric(raw_df.get("Units", pd.Series([0])), errors="coerce").fillna(0).sum())
    total_hours = float(pd.to_numeric(raw_df.get("Hours", pd.Series([0])), errors="coerce").fillna(0).sum())
    gross_hr = (total_revenue / total_hours) if total_hours else 0.0
    units_hr = (total_units / total_hours) if total_hours else 0.0

    summary["A1"] = "Executive Summary"
    summary["A1"].font = Font(bold=True, size=16)
    summary["A3"] = "Total Revenue"; summary["B3"] = round(total_revenue, 2)
    summary["A4"] = "Total Units"; summary["B4"] = round(total_units, 2)
    summary["A5"] = "Total Hours"; summary["B5"] = round(total_hours, 2)
    summary["A6"] = "Gross $/Hr"; summary["B6"] = round(gross_hr, 2)
    summary["A7"] = "Units/Hr"; summary["B7"] = round(units_hr, 2)
    for r in range(3, 8):
        summary[f"A{r}"].font = bold

    # Top/bottom techs
    summary["A9"] = "Top 10 Techs (by Efficiency Score)"; summary["A9"].font = bold
    top10 = individual.sort_values("Efficiency Score", ascending=False).head(10)[
        ["Rank", "Technician Team", "Technician", "Role", "Efficiency Score", "Tier", "Gross $/Hr", "Units/Hr", "Revenue"]
    ]
    for row in dataframe_to_rows(top10, index=False, header=True):
        summary.append(row)

    # Format header row for the top10 table (row 10)
    for cell in summary[10]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    # Add a simple bar chart for top10 efficiency
    chart = BarChart()
    chart.title = "Top 10 - Efficiency Score"
    # data range: Efficiency Score column in the appended table
    start_row = 10
    end_row = 10 + len(top10)
    # Efficiency Score column index within the table: find it
    headers = [c.value for c in summary[start_row]]
    score_col = headers.index("Efficiency Score") + 1
    name_col = headers.index("Technician") + 1

    data_ref = Reference(summary, min_col=score_col, min_row=start_row, max_row=end_row)
    cats_ref = Reference(summary, min_col=name_col, min_row=start_row+1, max_row=end_row)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.height = 9
    chart.width = 20
    summary.add_chart(chart, "A22")

    # Scatter plot: Gross $/Hr vs Units/Hr for all techs
    ws_lb = wb["Tech Leaderboard"]
    autosize(ws_lb)

    # Find columns
    lb_headers = [c.value for c in ws_lb[1]]
    if "Units/Hr" in lb_headers and "Gross $/Hr" in lb_headers:
        x_col = lb_headers.index("Units/Hr") + 1
        y_col = lb_headers.index("Gross $/Hr") + 1
        name_col_lb = lb_headers.index("Technician") + 1

        scatter = ScatterChart()
        scatter.title = "Techs: Gross $/Hr vs Units/Hr"
        scatter.x_axis.title = "Units/Hr"
        scatter.y_axis.title = "Gross $/Hr"

        xvalues = Reference(ws_lb, min_col=x_col, min_row=2, max_row=ws_lb.max_row)
        yvalues = Reference(ws_lb, min_col=y_col, min_row=2, max_row=ws_lb.max_row)
        series = Series(yvalues, xvalues, title_from_data=False)
        scatter.series.append(series)
        scatter.height = 12
        scatter.width = 20
        ws_lb.add_chart(scatter, "K2")

    # Autosize other sheets
    for name in wb.sheetnames:
        autosize(wb[name])

    wb.save(output_path)
    return output_path


def run_analysis(file_paths: List[str], cfg: AnalyzerConfig) -> Dict[str, Any]:
    """Run full analysis and return all computed dataframes.
    
    Returns a dict with keys:
        individual, team, store, trends, norm, team_tenure, tech_tenure, diagnostic, 
        mileage_analysis, weekly_staffing, weekly_pivot, staffing_trends, weekly_totals,
        capacity_analysis, capacity_weekly, seasonality
    """
    raw = load_excel_files(file_paths)
    if raw.empty:
        raise ValueError("No rows loaded from the provided files.")

    norm = normalize_columns(raw, cfg.column_aliases)

    # Remove blank tech rows
    norm = norm[norm["Technician"].astype(str).str.strip() != ""].copy()

    individual = compute_individual_stats(norm, cfg)
    team = compute_team_stats(norm, cfg)
    store = compute_store_stats(norm)
    trends = compute_trends(norm)
    
    # Tenure and diagnostic analysis
    team_tenure, tech_tenure = compute_tenure_stats(norm)
    diagnostic = compute_diagnostic_summary(team, team_tenure)
    mileage_analysis = compute_mileage_analysis(team)
    
    # Weekly staffing trends
    weekly_result = compute_weekly_staffing(norm)
    if weekly_result:
        weekly_staffing, weekly_pivot, staffing_trends = weekly_result
    else:
        weekly_staffing = weekly_pivot = staffing_trends = None
    
    weekly_totals = compute_weekly_staffing_totals(norm)
    
    # Capacity analysis
    capacity_result = compute_capacity_analysis(norm)
    if capacity_result:
        capacity_analysis, capacity_weekly = capacity_result
    else:
        capacity_analysis = capacity_weekly = None
    
    # Seasonality analysis
    seasonality = compute_seasonality_analysis(norm, cfg)

    return {
        'individual': individual,
        'team': team,
        'store': store,
        'trends': trends,
        'norm': norm,
        'team_tenure': team_tenure,
        'tech_tenure': tech_tenure,
        'diagnostic': diagnostic,
        'mileage_analysis': mileage_analysis,
        'weekly_staffing': weekly_staffing,
        'weekly_pivot': weekly_pivot,
        'staffing_trends': staffing_trends,
        'weekly_totals': weekly_totals,
        'capacity_analysis': capacity_analysis,
        'capacity_weekly': capacity_weekly,
        'seasonality': seasonality,
    }


# Legacy function signature for backward compatibility
def run_analysis_legacy(file_paths: List[str], cfg: AnalyzerConfig) -> Tuple[pd.DataFrame, pd.DataFrame, Optional[pd.DataFrame], Optional[pd.DataFrame], pd.DataFrame]:
    """Legacy function that returns a tuple for backward compatibility."""
    result = run_analysis(file_paths, cfg)
    return result['individual'], result['team'], result['store'], result['trends'], result['norm']
