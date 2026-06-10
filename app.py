import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import json
import os
from datetime import date, datetime, timedelta
import sys
sys.path.append(os.path.dirname(__file__))
from utils.pdf_generator import generate_pdf_report
from utils.data_store import load_data, save_entry, delete_entry, get_week_ending

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="RMCHCS Spend Tracker",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=DM+Serif+Display&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    /* Dark sidebar */
    [data-testid="stSidebar"] {
        background: #0f1724;
        border-right: 1px solid #1e2d40;
    }
    [data-testid="stSidebar"] * {
        color: #cbd5e1 !important;
    }
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stDateInput label,
    [data-testid="stSidebar"] h1, 
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 {
        color: #f1f5f9 !important;
    }

    /* Main background */
    .main .block-container {
        background: #f8fafc;
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }

    /* KPI cards */
    .kpi-card {
        background: #ffffff;
        border-radius: 12px;
        padding: 1.4rem 1.6rem;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    }
    .kpi-label {
        font-size: 0.72rem;
        font-weight: 600;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        color: #64748b;
        margin-bottom: 0.3rem;
    }
    .kpi-value {
        font-family: 'DM Serif Display', serif;
        font-size: 2rem;
        color: #0f172a;
        line-height: 1.1;
    }
    .kpi-sub {
        font-size: 0.78rem;
        color: #94a3b8;
        margin-top: 0.25rem;
    }
    .kpi-delta-up {
        color: #10b981;
        font-weight: 600;
        font-size: 0.82rem;
    }
    .kpi-delta-neutral {
        color: #64748b;
        font-weight: 500;
        font-size: 0.82rem;
    }

    /* Section headers */
    .section-header {
        font-family: 'DM Serif Display', serif;
        font-size: 1.25rem;
        color: #0f172a;
        margin-bottom: 0.2rem;
        margin-top: 0.5rem;
    }
    .section-sub {
        font-size: 0.82rem;
        color: #64748b;
        margin-bottom: 1rem;
    }

    /* Page header banner */
    .page-header {
        background: linear-gradient(135deg, #0f1724 0%, #1a3a5c 100%);
        border-radius: 14px;
        padding: 1.6rem 2rem;
        margin-bottom: 1.5rem;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .page-header-title {
        font-family: 'DM Serif Display', serif;
        font-size: 1.6rem;
        color: #f1f5f9;
        margin: 0;
    }
    .page-header-sub {
        font-size: 0.82rem;
        color: #94a3b8;
        margin-top: 0.2rem;
    }
    .page-header-badge {
        background: #1e3a5f;
        border: 1px solid #2d5a8e;
        border-radius: 8px;
        padding: 0.5rem 1rem;
        font-size: 0.78rem;
        color: #7dd3fc;
        font-weight: 500;
    }

    /* Form card */
    .form-card {
        background: #ffffff;
        border-radius: 12px;
        padding: 1.5rem;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 4px rgba(0,0,0,0.05);
    }

    /* Dataframe styling */
    .stDataFrame {
        border-radius: 10px;
        overflow: hidden;
    }

    /* Metric override */
    [data-testid="metric-container"] {
        background: white;
        border-radius: 12px;
        padding: 1rem;
        border: 1px solid #e2e8f0;
    }

    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        background: #f1f5f9;
        border-radius: 10px;
        padding: 4px;
        gap: 2px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        font-weight: 500;
        font-size: 0.88rem;
        color: #64748b;
    }
    .stTabs [aria-selected="true"] {
        background: white !important;
        color: #0f172a !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }

    /* Button */
    .stButton > button {
        border-radius: 8px;
        font-weight: 600;
        font-size: 0.88rem;
        transition: all 0.15s ease;
    }
    .stButton > button[kind="primary"] {
        background: #1a3a5c;
        border: none;
        color: white;
    }
    .stButton > button[kind="primary"]:hover {
        background: #0f2640;
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(26,58,92,0.3);
    }

    /* Success/error messages */
    .stSuccess {
        border-radius: 8px;
    }
    .stError {
        border-radius: 8px;
    }

    /* Hide default streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ── Load data ─────────────────────────────────────────────────────────────────
df = load_data()

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🏥 RMCHCS")
    st.markdown("**Spend Intelligence**")
    st.markdown("---")
    
    page = st.radio(
        "Navigate",
        ["Dashboard", "Log Spend", "Manage Entries", "Generate Report"],
        index=0
    )
    
    st.markdown("---")
    
    if not df.empty:
        st.markdown("**Quick Stats**")
        total_spend = df['total_spend'].sum()
        st.markdown(f"Total Tracked: **${total_spend:,.0f}**")
        st.markdown(f"Entries: **{len(df)}**")
        weeks = df['week_ending'].nunique() if 'week_ending' in df.columns else 0
        st.markdown(f"Weeks Tracked: **{weeks}**")
    
    st.markdown("---")
    st.markdown("<small style='color:#475569'>Vista Staffing Solutions<br>Internal Use Only</small>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
if page == "Dashboard":

    st.markdown("""
    <div class="page-header">
        <div>
            <div class="page-header-title">RMCHCS Spend Dashboard</div>
            <div class="page-header-sub">Rehoboth McKinley Christian Health Care Services · Gallup, NM</div>
        </div>
        <div class="page-header-badge">Vista Staffing Solutions</div>
    </div>
    """, unsafe_allow_html=True)

    if df.empty:
        st.info("No spend data yet. Go to **Log Spend** to add your first entry.")
        st.stop()

    # ── Filters row ──
    col_f1, col_f2, col_f3 = st.columns([2, 2, 2])
    with col_f1:
        all_weeks = sorted(df['week_ending'].unique(), reverse=True)
        week_options = ["All Weeks"] + list(all_weeks)
        selected_week = st.selectbox("Filter by Week", week_options)
    with col_f2:
        all_specialties = ["All Specialties"] + sorted(df['specialty'].unique().tolist())
        selected_specialty = st.selectbox("Filter by Specialty", all_specialties)
    with col_f3:
        all_service_lines = ["All Service Lines"] + sorted(df['service_line'].unique().tolist())
        selected_service_line = st.selectbox("Filter by Service Line", all_service_lines)

    # Apply filters
    filtered = df.copy()
    if selected_week != "All Weeks":
        filtered = filtered[filtered['week_ending'] == selected_week]
    if selected_specialty != "All Specialties":
        filtered = filtered[filtered['specialty'] == selected_specialty]
    if selected_service_line != "All Service Lines":
        filtered = filtered[filtered['service_line'] == selected_service_line]

    # ── KPI Cards ──
    st.markdown("<div style='margin-top:0.5rem'></div>", unsafe_allow_html=True)
    k1, k2, k3, k4 = st.columns(4)

    total = filtered['total_spend'].sum()
    providers = filtered['provider_name'].nunique()
    specialties = filtered['specialty'].nunique()
    avg_rate = filtered['daily_rate'].mean() if not filtered.empty else 0

    # Week-over-week delta
    if len(all_weeks) >= 2 and selected_week == "All Weeks":
        this_week = df[df['week_ending'] == all_weeks[0]]['total_spend'].sum()
        last_week = df[df['week_ending'] == all_weeks[1]]['total_spend'].sum()
        delta_pct = ((this_week - last_week) / last_week * 100) if last_week > 0 else 0
        delta_str = f"{'▲' if delta_pct >= 0 else '▼'} {abs(delta_pct):.1f}% vs last week"
        delta_class = "kpi-delta-up" if delta_pct >= 0 else "kpi-delta-neutral"
    else:
        delta_str = f"{filtered['week_ending'].nunique()} week(s) shown"
        delta_class = "kpi-delta-neutral"

    with k1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Total Spend</div>
            <div class="kpi-value">${total:,.0f}</div>
            <div class="{delta_class}">{delta_str}</div>
        </div>""", unsafe_allow_html=True)

    with k2:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Active Providers</div>
            <div class="kpi-value">{providers}</div>
            <div class="kpi-sub">Unique providers in view</div>
        </div>""", unsafe_allow_html=True)

    with k3:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Specialties Covered</div>
            <div class="kpi-value">{specialties}</div>
            <div class="kpi-sub">Across all service lines</div>
        </div>""", unsafe_allow_html=True)

    with k4:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Avg Daily Rate</div>
            <div class="kpi-value">${avg_rate:,.0f}</div>
            <div class="kpi-sub">Per provider per day</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<div style='margin-top:1.5rem'></div>", unsafe_allow_html=True)

    # ── Charts Row 1: Weekly Spend + Specialty Breakdown ──
    chart1, chart2 = st.columns([3, 2])

    with chart1:
        st.markdown('<div class="section-header">Weekly Spend Trend</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-sub">Total spend logged per week ending date</div>', unsafe_allow_html=True)

        weekly = df.groupby('week_ending')['total_spend'].sum().reset_index()
        weekly = weekly.sort_values('week_ending')

        fig_weekly = go.Figure()
        fig_weekly.add_trace(go.Bar(
            x=weekly['week_ending'],
            y=weekly['total_spend'],
            marker_color='#1a3a5c',
            marker_line_width=0,
            hovertemplate='<b>Week Ending:</b> %{x}<br><b>Spend:</b> $%{y:,.0f}<extra></extra>'
        ))
        fig_weekly.add_trace(go.Scatter(
            x=weekly['week_ending'],
            y=weekly['total_spend'],
            mode='lines+markers',
            line=dict(color='#3b82f6', width=2.5),
            marker=dict(size=6, color='#3b82f6'),
            hoverinfo='skip'
        ))
        fig_weekly.update_layout(
            plot_bgcolor='white',
            paper_bgcolor='white',
            margin=dict(l=0, r=0, t=10, b=0),
            height=280,
            xaxis=dict(showgrid=False, tickfont=dict(size=11)),
            yaxis=dict(showgrid=True, gridcolor='#f1f5f9', tickprefix='$', tickfont=dict(size=11)),
            showlegend=False
        )
        st.plotly_chart(fig_weekly, use_container_width=True)

    with chart2:
        st.markdown('<div class="section-header">Spend by Specialty</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-sub">Share of total spend per specialty</div>', unsafe_allow_html=True)

        spec_spend = filtered.groupby('specialty')['total_spend'].sum().reset_index()
        spec_spend = spec_spend.sort_values('total_spend', ascending=False)

        colors = ['#0f1724', '#1a3a5c', '#2d5a8e', '#3b82f6', '#60a5fa', '#93c5fd', '#bfdbfe']
        fig_pie = go.Figure(go.Pie(
            labels=spec_spend['specialty'],
            values=spec_spend['total_spend'],
            hole=0.52,
            marker=dict(colors=colors[:len(spec_spend)]),
            textinfo='percent',
            hovertemplate='<b>%{label}</b><br>$%{value:,.0f}<br>%{percent}<extra></extra>'
        ))
        fig_pie.update_layout(
            plot_bgcolor='white',
            paper_bgcolor='white',
            margin=dict(l=0, r=0, t=10, b=0),
            height=280,
            legend=dict(font=dict(size=11), orientation='v'),
            showlegend=True
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    # ── Charts Row 2: Provider spend bar + Service line breakdown ──
    chart3, chart4 = st.columns([3, 2])

    with chart3:
        st.markdown('<div class="section-header">Spend by Provider</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-sub">Total spend per provider in current view</div>', unsafe_allow_html=True)

        prov_spend = filtered.groupby('provider_name')['total_spend'].sum().reset_index()
        prov_spend = prov_spend.sort_values('total_spend', ascending=True)

        fig_prov = go.Figure(go.Bar(
            x=prov_spend['total_spend'],
            y=prov_spend['provider_name'],
            orientation='h',
            marker_color='#2d5a8e',
            marker_line_width=0,
            hovertemplate='<b>%{y}</b><br>$%{x:,.0f}<extra></extra>'
        ))
        fig_prov.update_layout(
            plot_bgcolor='white',
            paper_bgcolor='white',
            margin=dict(l=0, r=10, t=10, b=0),
            height=max(200, len(prov_spend) * 38),
            xaxis=dict(showgrid=True, gridcolor='#f1f5f9', tickprefix='$', tickfont=dict(size=11)),
            yaxis=dict(showgrid=False, tickfont=dict(size=11))
        )
        st.plotly_chart(fig_prov, use_container_width=True)

    with chart4:
        st.markdown('<div class="section-header">Service Line Breakdown</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-sub">Physician vs APP vs other lines</div>', unsafe_allow_html=True)

        sl_spend = filtered.groupby('service_line')['total_spend'].sum().reset_index()
        sl_spend = sl_spend.sort_values('total_spend', ascending=False)

        sl_colors = ['#1a3a5c', '#3b82f6', '#60a5fa', '#93c5fd']
        fig_sl = go.Figure(go.Bar(
            x=sl_spend['service_line'],
            y=sl_spend['total_spend'],
            marker_color=sl_colors[:len(sl_spend)],
            marker_line_width=0,
            hovertemplate='<b>%{x}</b><br>$%{y:,.0f}<extra></extra>'
        ))
        fig_sl.update_layout(
            plot_bgcolor='white',
            paper_bgcolor='white',
            margin=dict(l=0, r=0, t=10, b=0),
            height=280,
            xaxis=dict(showgrid=False, tickfont=dict(size=11)),
            yaxis=dict(showgrid=True, gridcolor='#f1f5f9', tickprefix='$', tickfont=dict(size=11))
        )
        st.plotly_chart(fig_sl, use_container_width=True)

    # ── Detail table ──
    st.markdown('<div class="section-header" style="margin-top:0.5rem">Entry Detail</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">All logged entries in current filter view</div>', unsafe_allow_html=True)

    display_cols = ['week_ending', 'provider_name', 'specialty', 'service_line', 'days_worked', 'daily_rate', 'total_spend', 'notes']
    display_df = filtered[display_cols].sort_values('week_ending', ascending=False).copy()
    display_df['daily_rate'] = display_df['daily_rate'].apply(lambda x: f"${x:,.2f}")
    display_df['total_spend'] = display_df['total_spend'].apply(lambda x: f"${x:,.2f}")
    display_df.columns = ['Week Ending', 'Provider', 'Specialty', 'Service Line', 'Days', 'Daily Rate', 'Total Spend', 'Notes']

    st.dataframe(display_df, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: LOG SPEND
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Log Spend":

    st.markdown("""
    <div class="page-header">
        <div>
            <div class="page-header-title">Log Weekly Spend</div>
            <div class="page-header-sub">Add a single entry manually or upload multiple rows at once</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    tab_manual, tab_bulk = st.tabs(["Single Entry", "Bulk Upload (CSV)"])

    # ── Tab 1: Manual entry ──────────────────────────────────────────────────
    with tab_manual:
        st.markdown("<div style='height:0.75rem'></div>", unsafe_allow_html=True)
        with st.container():
            st.markdown('<div class="form-card">', unsafe_allow_html=True)

            col1, col2 = st.columns(2)

            with col1:
                st.markdown("**Week & Provider**")
                week_ending = st.date_input(
                    "Week Ending Date",
                    value=get_week_ending(),
                    help="Select the Saturday or Sunday the work week ends on"
                )
                provider_name = st.text_input("Provider Name", placeholder="Dr. Jane Smith")
                provider_type = st.selectbox("Provider Type", ["Physician", "APP", "CRNA", "NP", "PA", "Other"])

            with col2:
                st.markdown("**Assignment Details**")
                specialty_options = [
                    "Emergency Medicine", "Hospitalist/Internal Medicine", "Family Medicine",
                    "OB/GYN", "Surgery", "Radiology", "Anesthesiology", "Psychiatry",
                    "Pediatrics", "Cardiology", "Orthopedics", "Neurology", "Other"
                ]
                specialty = st.selectbox("Specialty", specialty_options)
                service_line = st.selectbox("Service Line", ["Physician", "APP", "Nursing", "Allied Health", "Other"])
                department = st.text_input("Department / Unit", placeholder="e.g., ED, ICU, L&D")

            st.markdown("---")
            col3, col4, col5 = st.columns(3)

            with col3:
                days_worked = st.number_input("Days Worked", min_value=0.5, max_value=7.0, value=5.0, step=0.5)
            with col4:
                daily_rate = st.number_input("Daily Rate ($)", min_value=0.0, value=1800.0, step=50.0)
            with col5:
                auto_total = days_worked * daily_rate
                st.metric("Calculated Total", f"${auto_total:,.2f}")

            notes = st.text_area("Notes (optional)", placeholder="Contract type, extension, specialty notes...", height=80)

            st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)

            col_btn1, col_btn2 = st.columns([1, 4])
            with col_btn1:
                submit = st.button("Save Entry", type="primary", use_container_width=True)

            if submit:
                if not provider_name.strip():
                    st.error("Provider name is required.")
                else:
                    entry = {
                        "week_ending": str(week_ending),
                        "provider_name": provider_name.strip(),
                        "provider_type": provider_type,
                        "specialty": specialty,
                        "service_line": service_line,
                        "department": department.strip(),
                        "days_worked": days_worked,
                        "daily_rate": daily_rate,
                        "total_spend": round(days_worked * daily_rate, 2),
                        "notes": notes.strip(),
                        "logged_at": datetime.now().isoformat()
                    }
                    save_entry(entry)
                    st.success(f"Entry saved. {provider_name} | Week ending {week_ending} | ${entry['total_spend']:,.2f}")
                    st.rerun()

            st.markdown('</div>', unsafe_allow_html=True)

    # ── Tab 2: Bulk CSV upload ───────────────────────────────────────────────
    with tab_bulk:
        st.markdown("<div style='height:0.75rem'></div>", unsafe_allow_html=True)

        # Template download
        st.markdown('<div class="section-header">Step 1: Download the Template</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-sub">Fill this out in Excel or Google Sheets. Do not rename the column headers.</div>', unsafe_allow_html=True)

        template_data = {
            "week_ending":   ["2025-06-07", "2025-06-07"],
            "provider_name": ["Dr. Jane Smith", "Sarah Jones NP"],
            "provider_type": ["Physician", "NP"],
            "specialty":     ["Emergency Medicine", "Family Medicine"],
            "service_line":  ["Physician", "APP"],
            "department":    ["ED", "Clinic"],
            "days_worked":   [5, 3],
            "daily_rate":    [1800, 950],
            "notes":         ["Extended contract", ""]
        }
        template_df = pd.DataFrame(template_data)
        csv_template = template_df.to_csv(index=False).encode('utf-8')

        st.download_button(
            label="Download CSV Template",
            data=csv_template,
            file_name="RMCHCS_SpendUpload_Template.csv",
            mime="text/csv",
            type="primary"
        )

        st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)

        # Upload
        st.markdown('<div class="section-header">Step 2: Upload Your Filled File</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-sub">CSV files only. Totals are calculated automatically from days worked x daily rate.</div>', unsafe_allow_html=True)

        uploaded_file = st.file_uploader("Choose CSV file", type=["csv"], label_visibility="collapsed")

        if uploaded_file is not None:
            try:
                upload_df = pd.read_csv(uploaded_file)

                # Validate required columns
                required_cols = ["week_ending", "provider_name", "provider_type",
                                  "specialty", "service_line", "days_worked", "daily_rate"]
                missing = [c for c in required_cols if c not in upload_df.columns]

                if missing:
                    st.error(f"Missing required columns: {', '.join(missing)}. Download the template above and use it as your starting point.")
                else:
                    # Fill optional cols
                    if "department" not in upload_df.columns:
                        upload_df["department"] = ""
                    if "notes" not in upload_df.columns:
                        upload_df["notes"] = ""

                    upload_df["days_worked"] = pd.to_numeric(upload_df["days_worked"], errors="coerce")
                    upload_df["daily_rate"]  = pd.to_numeric(upload_df["daily_rate"],  errors="coerce")
                    upload_df["total_spend"] = (upload_df["days_worked"] * upload_df["daily_rate"]).round(2)

                    # Flag bad rows
                    bad_rows = upload_df[upload_df[["days_worked", "daily_rate", "total_spend"]].isnull().any(axis=1)]
                    clean_df = upload_df[~upload_df.index.isin(bad_rows.index)].copy()

                    st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)
                    st.markdown('<div class="section-header">Step 3: Preview Before Saving</div>', unsafe_allow_html=True)

                    if not bad_rows.empty:
                        st.warning(f"{len(bad_rows)} row(s) skipped due to missing or non-numeric days/rate values. They will not be imported.")
                        with st.expander("Show skipped rows"):
                            st.dataframe(bad_rows, use_container_width=True, hide_index=True)

                    if clean_df.empty:
                        st.error("No valid rows to import after validation.")
                    else:
                        preview_cols = ["week_ending", "provider_name", "provider_type",
                                        "specialty", "service_line", "days_worked", "daily_rate", "total_spend", "notes"]
                        preview_show = clean_df[preview_cols].copy()
                        preview_show["daily_rate"]  = preview_show["daily_rate"].apply(lambda x: f"${x:,.2f}")
                        preview_show["total_spend"] = preview_show["total_spend"].apply(lambda x: f"${x:,.2f}")
                        preview_show.columns = ["Week Ending", "Provider", "Type", "Specialty",
                                                 "Service Line", "Days", "Daily Rate", "Total Spend", "Notes"]

                        st.dataframe(preview_show, use_container_width=True, hide_index=True)

                        bulk_total = clean_df["total_spend"].sum()
                        st.markdown(f"**{len(clean_df)} rows ready to import | Combined total: ${bulk_total:,.2f}**")

                        st.markdown("<div style='height:0.4rem'></div>", unsafe_allow_html=True)
                        col_b1, col_b2 = st.columns([1, 4])
                        with col_b1:
                            confirm_import = st.button("Save All Entries", type="primary", use_container_width=True)

                        if confirm_import:
                            saved_count = 0
                            for _, row in clean_df.iterrows():
                                entry = {
                                    "week_ending":   str(row["week_ending"]).strip(),
                                    "provider_name": str(row["provider_name"]).strip(),
                                    "provider_type": str(row.get("provider_type", "")).strip(),
                                    "specialty":     str(row["specialty"]).strip(),
                                    "service_line":  str(row["service_line"]).strip(),
                                    "department":    str(row.get("department", "")).strip(),
                                    "days_worked":   float(row["days_worked"]),
                                    "daily_rate":    float(row["daily_rate"]),
                                    "total_spend":   float(row["total_spend"]),
                                    "notes":         str(row.get("notes", "")).strip(),
                                    "logged_at":     datetime.now().isoformat()
                                }
                                save_entry(entry)
                                saved_count += 1
                            st.success(f"{saved_count} entries saved successfully. Total spend logged: ${bulk_total:,.2f}")
                            st.rerun()

            except Exception as e:
                st.error(f"Could not read file: {e}. Make sure it is a valid CSV.")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: MANAGE ENTRIES
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Manage Entries":

    st.markdown("""
    <div class="page-header">
        <div>
            <div class="page-header-title">Manage Entries</div>
            <div class="page-header-sub">Review and delete logged spend entries</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if df.empty:
        st.info("No entries logged yet.")
        st.stop()

    # Filter by week
    all_weeks = sorted(df['week_ending'].unique(), reverse=True)
    selected_week = st.selectbox("Select Week", ["All Weeks"] + list(all_weeks))

    view_df = df.copy() if selected_week == "All Weeks" else df[df['week_ending'] == selected_week].copy()
    view_df = view_df.sort_values('week_ending', ascending=False).reset_index(drop=True)

    st.markdown(f"**{len(view_df)} entries** in current view")

    for idx, row in view_df.iterrows():
        with st.expander(f"{row['week_ending']} | {row['provider_name']} | {row['specialty']} | ${row['total_spend']:,.2f}"):
            c1, c2, c3 = st.columns(3)
            c1.write(f"**Provider Type:** {row.get('provider_type', 'N/A')}")
            c2.write(f"**Service Line:** {row['service_line']}")
            c3.write(f"**Department:** {row.get('department', 'N/A')}")
            c1.write(f"**Days Worked:** {row['days_worked']}")
            c2.write(f"**Daily Rate:** ${row['daily_rate']:,.2f}")
            c3.write(f"**Total Spend:** ${row['total_spend']:,.2f}")
            if row.get('notes'):
                st.write(f"**Notes:** {row['notes']}")

            if st.button(f"Delete Entry", key=f"del_{row['id']}", type="secondary"):
                delete_entry(row['id'])
                st.success("Entry deleted.")
                st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: GENERATE REPORT
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Generate Report":

    st.markdown("""
    <div class="page-header">
        <div>
            <div class="page-header-title">Generate Client Report</div>
            <div class="page-header-sub">Build a PDF spend summary to send to RMCHCS</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if df.empty:
        st.info("No data available to report on yet.")
        st.stop()

    col1, col2 = st.columns(2)

    with col1:
        all_weeks = sorted(df['week_ending'].unique(), reverse=True)
        report_week = st.selectbox("Select Week to Report", all_weeks)
        report_title = st.text_input("Report Title", value=f"Weekly Spend Report - Week Ending {report_week}")

    with col2:
        include_detail = st.checkbox("Include Provider Detail Table", value=True)
        include_notes = st.checkbox("Include Notes Column", value=True)
        prepared_by = st.text_input("Prepared By", value="Vista Staffing Solutions")

    st.markdown("---")

    week_data = df[df['week_ending'] == report_week]
    week_total = week_data['total_spend'].sum()
    week_providers = week_data['provider_name'].nunique()
    week_specialties = week_data['specialty'].nunique()

    st.markdown("**Report Preview**")
    m1, m2, m3 = st.columns(3)
    m1.metric("Week Total Spend", f"${week_total:,.2f}")
    m2.metric("Providers", week_providers)
    m3.metric("Specialties", week_specialties)

    st.dataframe(
        week_data[['provider_name', 'specialty', 'service_line', 'days_worked', 'daily_rate', 'total_spend']].rename(
            columns={'provider_name': 'Provider', 'specialty': 'Specialty', 'service_line': 'Service Line',
                     'days_worked': 'Days', 'daily_rate': 'Daily Rate', 'total_spend': 'Total Spend'}
        ),
        use_container_width=True, hide_index=True
    )

    st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)

    col_gen1, col_gen2 = st.columns([1, 4])
    with col_gen1:
        gen_button = st.button("Generate PDF", type="primary", use_container_width=True)

    if gen_button:
        with st.spinner("Building report..."):
            pdf_path = generate_pdf_report(
                df=week_data,
                week_ending=report_week,
                title=report_title,
                prepared_by=prepared_by,
                include_detail=include_detail,
                include_notes=include_notes
            )
        with open(pdf_path, "rb") as f:
            st.download_button(
                label="Download PDF Report",
                data=f,
                file_name=f"RMCHCS_SpendReport_{report_week}.pdf",
                mime="application/pdf",
                type="primary"
            )
        st.success("Report ready. Click above to download.")
