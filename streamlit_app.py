
import io
import json
from pathlib import Path
import tempfile
import zipfile

import pandas as pd
import streamlit as st

from analyzer_core import AnalyzerConfig, run_analysis, write_excel_report

st.set_page_config(page_title="Tech Efficiency Analyzer", layout="wide")

st.title("Tech Efficiency Analyzer")
st.caption("Upload one or more Daily Tech Performance Excel files to generate a leaderboard, team/store rollups, tenure analysis, diagnostics, and an executive summary.")

# Load config.json if present
cfg_path = Path("config.json")
cfg = AnalyzerConfig.from_json(cfg_path)

with st.sidebar:
    st.header("Scoring / Tuning")
    cfg.w_units_per_hr = st.slider("Weight: Units/Hr", 0.0, 1.0, float(cfg.w_units_per_hr), 0.05)
    cfg.w_revenue_per_job = st.slider("Weight: Revenue/Job", 0.0, 1.0, float(cfg.w_revenue_per_job), 0.05)
    cfg.w_revenue_per_hr = st.slider("Weight: Gross $/Hr", 0.0, 1.0, float(cfg.w_revenue_per_hr), 0.05)

    st.divider()
    st.subheader("Tier Cutoffs (Percentiles)")
    cfg.tier_elite = st.slider("Elite (>=)", 0.5, 0.99, float(cfg.tier_elite), 0.01)
    cfg.tier_strong = st.slider("Strong (>=)", 0.2, 0.95, float(cfg.tier_strong), 0.01)
    cfg.tier_average = st.slider("Average (>=)", 0.0, 0.80, float(cfg.tier_average), 0.01)

    st.divider()
    cfg.mileage_cost_per_mile = st.number_input("Mileage cost per mile (optional)", value=float(cfg.mileage_cost_per_mile), step=0.05)

    st.divider()
    st.subheader("Output")
    output_name = st.text_input("Output filename", value=cfg.output_filename)

st.write("### Upload files")
uploads = st.file_uploader("Excel files (.xlsx, .xls)", type=["xlsx", "xls"], accept_multiple_files=True)

run_btn = st.button("Run Analysis", type="primary", disabled=(not uploads))

if run_btn and uploads:
    with st.spinner("Reading files and running analysis..."):
        temp_dir = tempfile.mkdtemp(prefix="tech_analyzer_")
        paths = []
        for u in uploads:
            p = Path(temp_dir) / u.name
            p.write_bytes(u.getvalue())
            paths.append(str(p))

        results = run_analysis(paths, cfg)
        
        individual = results['individual']
        team = results['team']
        store = results['store']
        trends = results['trends']
        norm = results['norm']
        team_tenure = results['team_tenure']
        diagnostic = results['diagnostic']
        mileage_analysis = results['mileage_analysis']
        weekly_staffing = results['weekly_staffing']
        weekly_pivot = results['weekly_pivot']
        staffing_trends = results['staffing_trends']
        weekly_totals = results['weekly_totals']
        capacity_analysis = results['capacity_analysis']
        seasonality = results['seasonality']

        out_path = str(Path(temp_dir) / output_name)
        write_excel_report(
            out_path, norm, individual, team, 
            store=store, trends=trends,
            team_tenure=team_tenure, diagnostic=diagnostic, 
            mileage_analysis=mileage_analysis,
            weekly_staffing=weekly_staffing, weekly_pivot=weekly_pivot,
            staffing_trends=staffing_trends, weekly_totals=weekly_totals,
            capacity_analysis=capacity_analysis,
            capacity_weekly=results['capacity_weekly'],
            seasonality=seasonality
        )

        st.success("Report generated!")

        # Preview tabs
        tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
            "Tech Leaderboard", "Diagnostic Summary", "Capacity Analysis", 
            "Seasonality", "Weekly Staffing", "Team Tenure", "Mileage Analysis"
        ])
        
        with tab1:
            st.subheader("Tech Leaderboard")
            st.dataframe(individual.head(50), use_container_width=True)
        
        with tab2:
            st.subheader("Diagnostic Summary")
            st.caption("Teams diagnosed by Efficiency + Tenure quadrant")
            st.dataframe(diagnostic, use_container_width=True)
        
        with tab3:
            st.subheader("Capacity Analysis")
            st.caption("Does each team have the right number of techs for their workload?")
            
            if capacity_analysis is not None:
                # Show assessment summary
                col1, col2, col3 = st.columns(3)
                understaffed = len(capacity_analysis[capacity_analysis["Staffing Assessment"].str.contains("Understaffed")])
                overstaffed = len(capacity_analysis[capacity_analysis["Staffing Assessment"].str.contains("Overstaffed")])
                rightsized = len(capacity_analysis[capacity_analysis["Staffing Assessment"].str.contains("Right-sized")])
                
                with col1:
                    st.metric("🔴 Understaffed", understaffed)
                with col2:
                    st.metric("🟢 Right-sized", rightsized)
                with col3:
                    st.metric("🔴 Overstaffed", overstaffed)
                
                st.write("#### Team Capacity Assessment")
                st.dataframe(capacity_analysis, use_container_width=True)
            else:
                st.info("Capacity analysis requires Date column in data")
        
        with tab4:
            st.subheader("Seasonality Analysis")
            st.caption("Does workload seasonality affect efficiency?")
            
            if seasonality is not None:
                # Key insights
                st.write("#### Key Insights")
                st.dataframe(seasonality["summary"], use_container_width=True)
                
                # Correlation interpretation
                vol_corr = seasonality.get("vol_eff_correlation")
                if pd.notna(vol_corr):
                    if vol_corr < -0.3:
                        st.warning(f"⚠️ Volume-Efficiency Correlation: {vol_corr:.2f} — Efficiency DROPS when volume is high. Consider adding capacity during peak periods.")
                    elif vol_corr > 0.3:
                        st.success(f"✅ Volume-Efficiency Correlation: {vol_corr:.2f} — Efficiency IMPROVES when busy. Teams perform better under load.")
                    else:
                        st.info(f"➡️ Volume-Efficiency Correlation: {vol_corr:.2f} — Weak relationship. Seasonality has minimal impact on efficiency.")
                
                # Charts
                st.write("#### Weekly Volume & Efficiency Trends")
                weekly_data = seasonality["weekly"].set_index("Week Start")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.write("**Volume Index** (100 = average)")
                    st.line_chart(weekly_data["Volume Index"])
                with col2:
                    st.write("**Efficiency Index** (100 = average)")
                    st.line_chart(weekly_data["Efficiency Index"])
                
                st.write("#### Monthly Breakdown")
                st.dataframe(seasonality["monthly"], use_container_width=True)
            else:
                st.info("Seasonality analysis requires Date column in data")
        
        with tab5:
            st.subheader("Weekly Staffing Analysis")
            st.caption("Unique techs working per week (Mon-Sun) by team")
            
            if weekly_totals is not None:
                # Chart: Total unique techs per week
                st.write("#### Territory-Wide Staffing")
                st.line_chart(weekly_totals.set_index("Week Start")["Unique Techs"])
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Peak Week Techs", int(weekly_totals["Unique Techs"].max()))
                with col2:
                    st.metric("Avg Weekly Techs", round(weekly_totals["Unique Techs"].mean(), 1))
            
            if staffing_trends is not None:
                st.write("#### Team Staffing Trends")
                st.caption("Comparing first 4 weeks vs last 4 weeks")
                st.dataframe(staffing_trends, use_container_width=True)
            
            if weekly_pivot is not None:
                st.write("#### Unique Techs by Team per Week")
                st.dataframe(weekly_pivot, use_container_width=True)
        
        with tab6:
            st.subheader("Team Tenure (Ranked by Core %)")
            st.dataframe(team_tenure.sort_values('Core %', ascending=False), use_container_width=True)
        
        with tab7:
            st.subheader("Mileage Analysis")
            st.dataframe(mileage_analysis, use_container_width=True)

        if store is not None:
            st.subheader("Store Summary (preview)")
            st.dataframe(store.head(50), use_container_width=True)

        if trends is not None:
            st.subheader("Trends (Daily) (preview)")
            st.dataframe(trends.head(100), use_container_width=True)

        # Download
        data = Path(out_path).read_bytes()
        st.download_button(
            label="Download Excel Report",
            data=data,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

st.write("---")
st.write("#### Notes")
st.write("- The analyzer auto-normalizes column headers (e.g., 'Revenue' vs 'Amount').")
st.write("- **Diagnostic Summary** identifies teams by quadrant: Strong Team, Slow Techs, Turnover Risk, Needs Attention.")
st.write("- **Core Tech** = 60+ days worked in the period. **Core %** = retention indicator.")
st.write("- If your source files include Store/CID columns, you'll get a Store Summary automatically.")
st.write("- Trends require a usable Date column.")
