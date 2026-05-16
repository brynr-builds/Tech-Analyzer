"""
Tech Efficiency Dashboard
=========================
Comprehensive visual dashboard summarizing all analysis with filtering capabilities.

Run with: streamlit run dashboard.py
"""

import io
import json
from pathlib import Path
import tempfile
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Any

import pandas as pd
import numpy as np
import streamlit as st

from analyzer_core import (
    AnalyzerConfig, load_excel_files, normalize_columns,
    compute_individual_stats, compute_team_stats, compute_tenure_stats,
    compute_diagnostic_summary, compute_mileage_analysis,
    compute_weekly_staffing, compute_weekly_staffing_totals,
    compute_capacity_analysis, compute_seasonality_analysis,
    compute_trends, write_excel_report, _safe_div
)

# Page config
st.set_page_config(
    page_title="Tech Efficiency Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better visuals
st.markdown("""
<style>
    .metric-card {
        background-color: #1e1e1e;
        border-radius: 10px;
        padding: 15px;
        margin: 5px;
    }
    .stMetric {
        background-color: #262730;
        padding: 10px;
        border-radius: 5px;
    }
    .flag-warning {
        background-color: #ff6b6b;
        color: white;
        padding: 5px 10px;
        border-radius: 5px;
        font-size: 12px;
    }
    .flag-info {
        background-color: #4ecdc4;
        color: white;
        padding: 5px 10px;
        border-radius: 5px;
        font-size: 12px;
    }
</style>
""", unsafe_allow_html=True)


def load_and_process_data(file_paths: List[str], cfg: AnalyzerConfig) -> pd.DataFrame:
    """Load and normalize data from files."""
    raw = load_excel_files(file_paths)
    if raw.empty:
        raise ValueError("No data loaded")
    
    norm = normalize_columns(raw, cfg.column_aliases)
    norm = norm[norm["Technician"].astype(str).str.strip() != ""].copy()
    
    # Ensure Date is datetime
    if "Date" in norm.columns:
        norm["Date"] = pd.to_datetime(norm["Date"], errors="coerce")
    
    return norm


def get_team_activity_flags(df: pd.DataFrame, min_weeks: int = 4) -> pd.DataFrame:
    """Flag teams with limited activity periods."""
    if "Date" not in df.columns or df["Date"].isna().all():
        return pd.DataFrame()
    
    team_activity = df.groupby("Technician Team").agg({
        "Date": ["min", "max", "nunique"]
    }).reset_index()
    team_activity.columns = ["Team", "First Active", "Last Active", "Days Active"]
    
    team_activity["First Active"] = pd.to_datetime(team_activity["First Active"])
    team_activity["Last Active"] = pd.to_datetime(team_activity["Last Active"])
    team_activity["Weeks Active"] = ((team_activity["Last Active"] - team_activity["First Active"]).dt.days / 7).round(1)
    
    # Calculate total date range in data
    total_weeks = (df["Date"].max() - df["Date"].min()).days / 7
    team_activity["Coverage %"] = (team_activity["Weeks Active"] / total_weeks * 100).round(1) if total_weeks > 0 else 100
    
    # Flag status
    def get_flag(row):
        if row["Weeks Active"] < min_weeks:
            return "🔴 New/Inactive"
        elif row["Coverage %"] < 50:
            return "🟡 Partial Period"
        elif row["Coverage %"] < 80:
            return "🟢 Most of Period"
        else:
            return "✅ Full Period"
    
    team_activity["Activity Status"] = team_activity.apply(get_flag, axis=1)
    
    return team_activity


def filter_data(df: pd.DataFrame, date_range: tuple, selected_teams: List[str]) -> pd.DataFrame:
    """Filter dataframe by date range and teams."""
    filtered = df.copy()
    
    # Date filter
    if "Date" in filtered.columns and date_range:
        start_date, end_date = date_range
        filtered = filtered[
            (filtered["Date"] >= pd.Timestamp(start_date)) & 
            (filtered["Date"] <= pd.Timestamp(end_date))
        ]
    
    # Team filter
    if selected_teams and len(selected_teams) > 0:
        filtered = filtered[filtered["Technician Team"].isin(selected_teams)]
    
    return filtered


def render_kpi_cards(df: pd.DataFrame, team_stats: pd.DataFrame, individual_stats: pd.DataFrame):
    """Render top-level KPI cards."""
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    
    total_revenue = df["Amount"].sum()
    total_units = df["Units"].sum()
    total_hours = df["Hours"].sum()
    avg_efficiency = (total_revenue / total_hours) if total_hours > 0 else 0
    
    with col1:
        st.metric("💰 Total Revenue", f"${total_revenue:,.0f}")
    with col2:
        st.metric("📦 Total Units", f"{total_units:,.0f}")
    with col3:
        st.metric("⏱️ Total Hours", f"{total_hours:,.0f}")
    with col4:
        st.metric("💵 Avg $/Hr", f"${avg_efficiency:.2f}")
    with col5:
        st.metric("👥 Teams", len(team_stats))
    with col6:
        st.metric("🔧 Technicians", df["Technician"].nunique())


def render_diagnostic_overview(diagnostic: pd.DataFrame):
    """Render diagnostic quadrant summary."""
    st.subheader("🎯 Team Diagnostics")
    
    if diagnostic is None or diagnostic.empty:
        st.info("No diagnostic data available")
        return
    
    # Count by diagnosis
    diag_counts = diagnostic["Diagnosis"].value_counts()
    
    col1, col2, col3, col4 = st.columns(4)
    
    strong = diag_counts.get("✓ Strong Team", 0)
    slow = diag_counts.get("⚠ Slow Techs (Coach/Train)", 0)
    turnover = diag_counts.get("⚠ Turnover Risk (Retain)", 0)
    attention = diag_counts.get("✗ Needs Attention (Rebuild)", 0)
    
    with col1:
        st.metric("✓ Strong Teams", strong, help="High efficiency + High tenure")
    with col2:
        st.metric("⚠ Slow Techs", slow, help="High tenure but low efficiency - coaching needed")
    with col3:
        st.metric("⚠ Turnover Risk", turnover, help="High efficiency but low tenure - retention issue")
    with col4:
        st.metric("✗ Needs Attention", attention, help="Low efficiency + Low tenure - rebuild needed")
    
    # Show teams needing attention
    problem_teams = diagnostic[diagnostic["Diagnosis"].str.contains("⚠|✗")]
    if not problem_teams.empty:
        with st.expander(f"🚨 Teams Needing Attention ({len(problem_teams)})", expanded=False):
            st.dataframe(
                problem_teams[["Technician Team", "Diagnosis", "Efficiency Rank", "Tenure Rank", "Gross $/Hr", "Core %"]],
                use_container_width=True,
                hide_index=True
            )


def render_capacity_overview(capacity_analysis: pd.DataFrame):
    """Render capacity analysis summary."""
    st.subheader("📊 Capacity Analysis")
    
    if capacity_analysis is None or capacity_analysis.empty:
        st.info("No capacity data available (requires Date column)")
        return
    
    # Count by assessment
    col1, col2, col3, col4, col5 = st.columns(5)
    
    understaffed = len(capacity_analysis[capacity_analysis["Staffing Assessment"].str.contains("Understaffed")])
    rightsized = len(capacity_analysis[capacity_analysis["Staffing Assessment"].str.contains("Right-sized")])
    overstaffed = len(capacity_analysis[capacity_analysis["Staffing Assessment"].str.contains("Overstaffed")])
    
    techs_to_hire = capacity_analysis[capacity_analysis["Techs to Hire/Transfer"] > 0]["Techs to Hire/Transfer"].sum()
    techs_to_transfer = abs(capacity_analysis[capacity_analysis["Techs to Hire/Transfer"] < 0]["Techs to Hire/Transfer"].sum())

    with col1:
        st.metric("🔴 Understaffed", understaffed, help="Output/tech >10% above average")
    with col2:
        st.metric("🟢 Right-sized", rightsized, help="Output/tech within ±10% of average")
    with col3:
        st.metric("🔴 Overstaffed", overstaffed, help="Output/tech >10% below average")
    with col4:
        st.metric("📈 Total Techs to Hire", round(techs_to_hire, 1), help="Total missing headcount across understaffed teams")
    with col5:
        st.metric("📉 Total Techs to Transfer", round(techs_to_transfer, 1), help="Total excess headcount across overstaffed teams")
    
    # Scatter plot: Techs vs Output/Tech
    if len(capacity_analysis) > 1:
        chart_data = capacity_analysis[["Team", "Avg Techs/Week", "Avg Units/Tech/Week"]].copy()
        chart_data = chart_data.set_index("Team")
        
        st.write("**Team Size vs Productivity**")
        st.scatter_chart(
            chart_data,
            x="Avg Techs/Week",
            y="Avg Units/Tech/Week",
            use_container_width=True
        )


def render_seasonality_overview(seasonality: Dict[str, Any]):
    """Render seasonality analysis summary."""
    st.subheader("📅 Seasonality Impact")
    
    if seasonality is None:
        st.info("No seasonality data available (requires Date column)")
        return
    
    # Correlation insight
    vol_corr = seasonality.get("vol_eff_correlation")
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        if pd.notna(vol_corr):
            st.metric("Volume-Efficiency Correlation", f"{vol_corr:.2f}")
            
            if vol_corr < -0.3:
                st.error("⚠️ Efficiency DROPS when busy")
                st.caption("Consider adding capacity during peak periods")
            elif vol_corr > 0.3:
                st.success("✅ Efficiency IMPROVES when busy")
                st.caption("Teams perform better under load")
            else:
                st.info("➡️ Weak relationship")
                st.caption("Seasonality has minimal impact")
    
    with col2:
        # Volume and efficiency trends
        weekly = seasonality.get("weekly")
        if weekly is not None and len(weekly) > 1:
            chart_data = weekly[["Week Start", "Volume Index", "Efficiency Index"]].copy()
            chart_data["Week Start"] = pd.to_datetime(chart_data["Week Start"])
            chart_data = chart_data.set_index("Week Start")
            
            st.line_chart(chart_data, use_container_width=True)
            st.caption("Index: 100 = average. Above 100 = above average.")


def render_staffing_trends(weekly_totals: pd.DataFrame, staffing_trends: pd.DataFrame):
    """Render staffing trends visualization."""
    st.subheader("👥 Staffing Trends")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        if weekly_totals is not None and len(weekly_totals) > 1:
            chart_data = weekly_totals[["Week Start", "Unique Techs", "Hours"]].copy()
            chart_data["Week Start"] = pd.to_datetime(chart_data["Week Start"])
            chart_data = chart_data.set_index("Week Start")
            
            st.line_chart(chart_data["Unique Techs"], use_container_width=True)
            st.caption("Unique technicians working per week")
    
    with col2:
        if staffing_trends is not None and not staffing_trends.empty:
            growing = len(staffing_trends[staffing_trends["Trend"].str.contains("Growing")])
            shrinking = len(staffing_trends[staffing_trends["Trend"].str.contains("Shrinking")])
            stable = len(staffing_trends[staffing_trends["Trend"].str.contains("Stable")])
            
            st.metric("📈 Growing Teams", growing)
            st.metric("📉 Shrinking Teams", shrinking)
            st.metric("➡️ Stable Teams", stable)


def render_efficiency_rankings(team_stats: pd.DataFrame, individual_stats: pd.DataFrame):
    """Render efficiency leaderboards."""
    st.subheader("🏆 Efficiency Rankings")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.write("**Top 10 Teams (by Gross $/Hr)**")
        top_teams = team_stats.nlargest(10, "Gross $/Hr")[
            ["Technician Team", "Gross $/Hr", "Net $/Hr", "Units/Hr", "Hours"]
        ].reset_index(drop=True)
        top_teams.index = top_teams.index + 1
        st.dataframe(top_teams, use_container_width=True)
    
    with col2:
        st.write("**Top 10 Technicians (by Efficiency Score)**")
        if "Efficiency Score" in individual_stats.columns:
            cols_to_show = ["Technician", "Technician Team", "Efficiency Score", "Tier", "Gross $/Hr"]
            if "Coaching Profile" in individual_stats.columns:
                cols_to_show.append("Coaching Profile")
            if "Flight Risk" in individual_stats.columns:
                cols_to_show.append("Flight Risk")

            top_techs = individual_stats.nlargest(10, "Efficiency Score")[cols_to_show].reset_index(drop=True)
            top_techs.index = top_techs.index + 1

            chart_data = top_techs[["Technician", "Efficiency Score"]].set_index("Technician")
            st.bar_chart(chart_data)

            st.dataframe(top_techs, use_container_width=True)


def render_team_activity_flags(activity_flags: pd.DataFrame):
    """Render team activity flag warnings."""
    if activity_flags is None or activity_flags.empty:
        return
    
    # Filter to only show flagged teams
    flagged = activity_flags[~activity_flags["Activity Status"].str.contains("Full Period")]
    
    if flagged.empty:
        return
    
    st.subheader("🚩 Team Activity Flags")
    st.caption("Teams with limited activity during the selected period")
    
    # Color code by status
    def highlight_status(row):
        if "🔴" in row["Activity Status"]:
            return ["background-color: #ff6b6b20"] * len(row)
        elif "🟡" in row["Activity Status"]:
            return ["background-color: #feca5720"] * len(row)
        else:
            return [""] * len(row)
    
    display_cols = ["Team", "Activity Status", "First Active", "Last Active", "Weeks Active", "Coverage %"]
    flagged_display = flagged[display_cols].copy()
    flagged_display["First Active"] = flagged_display["First Active"].dt.strftime("%Y-%m-%d")
    flagged_display["Last Active"] = flagged_display["Last Active"].dt.strftime("%Y-%m-%d")
    
    st.dataframe(
        flagged_display.style.apply(highlight_status, axis=1),
        use_container_width=True,
        hide_index=True
    )


def render_mileage_analysis(mileage_analysis: pd.DataFrame):
    """Render mileage analysis summary."""
    st.subheader("🚗 Mileage Analysis")
    
    if mileage_analysis is None or mileage_analysis.empty:
        st.info("No mileage data available")
        return
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        # Top 10 by mileage cost
        top_mileage = mileage_analysis.nlargest(10, "Mileage Paid")[
            ["Technician Team", "Mileage Paid", "Mileage Cost %", "Gross $/Hr", "Net $/Hr"]
        ].reset_index(drop=True)
        top_mileage.index = top_mileage.index + 1
        top_mileage["Mileage Paid"] = top_mileage["Mileage Paid"].apply(lambda x: f"${x:,.0f}")
        top_mileage["Mileage Cost %"] = top_mileage["Mileage Cost %"].apply(lambda x: f"{x:.1f}%")
        st.dataframe(top_mileage, use_container_width=True)
    
    with col2:
        total_mileage = mileage_analysis["Mileage Paid"].sum()
        avg_mileage_pct = mileage_analysis["Mileage Cost %"].mean()
        
        st.metric("Total Mileage Cost", f"${total_mileage:,.0f}")
        st.metric("Avg Mileage % of Revenue", f"{avg_mileage_pct:.1f}%")


def render_detailed_tables(
    team_stats: pd.DataFrame,
    individual_stats: pd.DataFrame,
    diagnostic: pd.DataFrame,
    capacity_analysis: pd.DataFrame
):
    """Render detailed data tables in expandable sections."""
    st.subheader("📋 Detailed Data")
    
    tab1, tab2, tab3, tab4 = st.tabs(["Team Summary", "Tech Leaderboard", "Diagnostics", "Capacity"])
    
    with tab1:
        st.dataframe(team_stats, use_container_width=True, hide_index=True)
    
    with tab2:
        st.dataframe(individual_stats.head(100), use_container_width=True, hide_index=True)
    
    with tab3:
        if diagnostic is not None:
            def highlight_diagnosis(val):
                if isinstance(val, str):
                    if "✗" in val or "⚠" in val:
                        return 'background-color: #ffcccc'
                    if "✓" in val:
                        return 'background-color: #ccffcc'
                return ''

            styled_diagnostic = diagnostic.style.map(highlight_diagnosis, subset=['Diagnosis'])
            st.dataframe(styled_diagnostic, use_container_width=True, hide_index=True)
    
    with tab4:
        if capacity_analysis is not None:
            def highlight_capacity(val):
                if isinstance(val, str):
                    if "Understaffed" in val or "Overstaffed" in val:
                        return 'background-color: #ffcccc'
                    if "Right-sized" in val:
                        return 'background-color: #ccffcc'
                return ''

            styled_capacity = capacity_analysis.style.map(highlight_capacity, subset=['Staffing Assessment'])
            st.dataframe(styled_capacity, use_container_width=True, hide_index=True)


def main():
    st.title("📊 Tech Efficiency Dashboard")
    st.caption("Comprehensive analysis of technician performance, capacity, and trends")
    
    # Sidebar - File Upload
    with st.sidebar:
        st.header("📁 Data Upload")
        uploads = st.file_uploader(
            "Upload Excel files or Zips",
            type=["xlsx", "xls", "zip"],
            accept_multiple_files=True,
            help="Upload Daily Tech Performance files or Zip files containing them"
        )
        
        if not uploads:
            st.info("👆 Upload files to get started")
        
        # Load config
        cfg = AnalyzerConfig()

    if not uploads:
        st.markdown("### Welcome to the Tech Efficiency Analyzer!")
        st.markdown("""
        This tool analyzes your daily technician performance data to provide actionable insights.

        **With this dashboard, you can:**
        - Identify understaffed teams that need to hire, and overstaffed teams that can reduce headcount.
        - Get data-driven coaching recommendations for every single technician (e.g., Needs Speed Training, Upsell Training).
        - Catch flight risks early before they churn.
        - Understand how seasonality impacts efficiency.

        **How to use:**
        1. Open the sidebar on the left.
        2. Upload your Daily Tech Performance Excel files (or a `.zip` file containing them).
        3. The dashboard will automatically calculate everything and generate a downloadable report!
        
        **Expected Excel Columns:**
        - `Technician` (or Tech/Name)
        - `Technician Team` (or Team/Region)
        - `Hours`
        - `Units`
        - `Amount` (or Revenue)
        - `Date` (Optional, required for capacity and flight risk)
        """)
        st.stop()

    # Process uploads
    with st.spinner("Loading data..."):
        import zipfile
        import os
        temp_dir = tempfile.mkdtemp(prefix="dashboard_")
        paths = []
        for i, u in enumerate(uploads):
            if u.name.endswith('.zip'):
                zip_dir = os.path.join(temp_dir, f"zip_{i}")
                os.makedirs(zip_dir, exist_ok=True)
                with zipfile.ZipFile(io.BytesIO(u.getvalue())) as z:
                    z.extractall(zip_dir)
                    # Add any extracted excel files to our paths list
                    for root, _, files in os.walk(zip_dir):
                        for file in files:
                            if file.endswith('.xlsx') or file.endswith('.xls'):
                                extracted_path = os.path.join(root, file)
                                if extracted_path not in paths:
                                    paths.append(extracted_path)
            else:
                p = Path(temp_dir) / f"{i}_{u.name}"
                p.write_bytes(u.getvalue())
                paths.append(str(p))
        
        try:
            df = load_and_process_data(paths, cfg)
        except Exception as e:
            st.error(f"Error loading data: {e}")
            st.stop()

    st.success(f"✅ Loaded {len(df):,} records")

    st.divider()
        
    # Date Filter
    with st.sidebar:
        st.header("📅 Date Filter")
        
        if "Date" in df.columns and df["Date"].notna().any():
            min_date = df["Date"].min().date()
            max_date = df["Date"].max().date()
            
            date_range = st.date_input(
                "Select date range",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date
            )
            
            if len(date_range) == 2:
                st.caption(f"Selected: {date_range[0]} to {date_range[1]}")
            else:
                date_range = (min_date, max_date)
        else:
            date_range = None
            st.info("No date column found")
        
        st.divider()
        
        # Team Filter
        st.header("👥 Team Filter")
        
        all_teams = sorted(df["Technician Team"].dropna().unique().tolist())
        
        select_all = st.checkbox("Select All Teams", value=True)
        
        if select_all:
            selected_teams = all_teams
        else:
            selected_teams = st.multiselect(
                "Select teams",
                options=all_teams,
                default=all_teams[:10] if len(all_teams) > 10 else all_teams
            )
        
        st.caption(f"Selected: {len(selected_teams)} of {len(all_teams)} teams")
        
        st.divider()
        
        # Activity flag threshold
        st.header("🚩 Activity Flags")
        min_weeks = st.slider(
            "Min weeks to be 'active'",
            min_value=1,
            max_value=12,
            value=4,
            help="Teams active less than this will be flagged"
        )
        
        st.divider()
        
        # Export button
        st.header("📥 Export")
        export_btn = st.button("Generate Full Excel Report", type="primary", use_container_width=True)
    
    # Apply filters
    filtered_df = filter_data(df, date_range, selected_teams)
    
    if filtered_df.empty:
        st.warning("No data matches the selected filters")
        st.stop()
    
    # Calculate all stats on filtered data
    with st.spinner("Calculating statistics..."):
        team_stats = compute_team_stats(filtered_df, cfg)
        individual_stats = compute_individual_stats(filtered_df, cfg)
        team_tenure, tech_tenure = compute_tenure_stats(filtered_df)
        diagnostic = compute_diagnostic_summary(team_stats, team_tenure)
        mileage_analysis = compute_mileage_analysis(team_stats)
        
        # Weekly analyses
        weekly_result = compute_weekly_staffing(filtered_df)
        if weekly_result:
            weekly_staffing, weekly_pivot, staffing_trends = weekly_result
        else:
            weekly_staffing = weekly_pivot = staffing_trends = None
        
        weekly_totals = compute_weekly_staffing_totals(filtered_df)
        
        # Capacity
        capacity_result = compute_capacity_analysis(filtered_df)
        if capacity_result:
            capacity_analysis, capacity_weekly = capacity_result
        else:
            capacity_analysis = capacity_weekly = None
        
        # Seasonality
        seasonality = compute_seasonality_analysis(filtered_df, cfg)
        
        # Activity flags
        activity_flags = get_team_activity_flags(filtered_df, min_weeks)
    
    # Render dashboard
    st.markdown("---")
    
    # Row 1: KPIs
    render_kpi_cards(filtered_df, team_stats, individual_stats)
    
    st.markdown("---")
    
    # Row 2: Activity Flags (if any)
    render_team_activity_flags(activity_flags)
    
    # Row 3: Diagnostics and Capacity
    col1, col2 = st.columns(2)
    with col1:
        render_diagnostic_overview(diagnostic)
    with col2:
        render_capacity_overview(capacity_analysis)
    
    st.markdown("---")
    
    # Row 4: Seasonality and Staffing
    col1, col2 = st.columns(2)
    with col1:
        render_seasonality_overview(seasonality)
    with col2:
        render_staffing_trends(weekly_totals, staffing_trends)
    
    st.markdown("---")
    
    # Row 5: Rankings
    render_efficiency_rankings(team_stats, individual_stats)
    
    st.markdown("---")
    
    # Row 6: Mileage
    render_mileage_analysis(mileage_analysis)
    
    st.markdown("---")
    
    # Row 7: Detailed Tables
    render_detailed_tables(team_stats, individual_stats, diagnostic, capacity_analysis)
    
    # Handle export
    if export_btn:
        with st.spinner("Generating report..."):
            output_name = f"Tech_Efficiency_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            out_path = str(Path(temp_dir) / output_name)
            
            write_excel_report(
                out_path, filtered_df, individual_stats, team_stats,
                store=None, trends=compute_trends(filtered_df),
                team_tenure=team_tenure, diagnostic=diagnostic,
                mileage_analysis=mileage_analysis,
                weekly_staffing=weekly_staffing, weekly_pivot=weekly_pivot,
                staffing_trends=staffing_trends, weekly_totals=weekly_totals,
                capacity_analysis=capacity_analysis, capacity_weekly=capacity_weekly,
                seasonality=seasonality
            )
            
            data = Path(out_path).read_bytes()
            
            st.download_button(
                label="📥 Download Excel Report",
                data=data,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


if __name__ == "__main__":
    main()
