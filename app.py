import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import json
import os
import uuid
import io
from datetime import date, datetime, timedelta
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

# ══════════════════════════════════════════════════════════════════════════════
# DATA STORE  (no external files needed — JSON lives next to app.py)
# ══════════════════════════════════════════════════════════════════════════════
BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "spend_log.json")

def _ensure_data_file():
    if not os.path.exists(DATA_FILE):
        with open(DATA_FILE, "w") as f:
            json.dump([], f)

# Service lines billed hourly vs daily
HOURLY_SERVICE_LINES = {"Nursing", "Allied Health"}

def load_data() -> pd.DataFrame:
    _ensure_data_file()
    with open(DATA_FILE, "r") as f:
        data = json.load(f)
    if not data:
        return pd.DataFrame(columns=[
            "id", "week_ending", "provider_name", "provider_type",
            "specialty", "service_line", "department",
            "days_worked", "daily_rate",
            "hours_worked", "bill_rate",
            "total_spend", "notes", "logged_at"
        ])
    df = pd.DataFrame(data)
    for col in ["days_worked", "daily_rate", "hours_worked", "bill_rate", "total_spend"]:
        if col not in df.columns:
            df[col] = None
        df[col] = pd.to_numeric(df[col], errors="coerce")
    return df

def save_entry(entry: dict):
    _ensure_data_file()
    with open(DATA_FILE, "r") as f:
        data = json.load(f)
    entry["id"] = str(uuid.uuid4())
    data.append(entry)
    with open(DATA_FILE, "w") as f:
        json.dump(data, f, indent=2)

def delete_entry(entry_id: str):
    _ensure_data_file()
    with open(DATA_FILE, "r") as f:
        data = json.load(f)
    data = [e for e in data if e.get("id") != entry_id]
    with open(DATA_FILE, "w") as f:
        json.dump(data, f, indent=2)

def get_week_ending() -> date:
    today = date.today()
    days_until_saturday = (5 - today.weekday()) % 7
    return today if days_until_saturday == 0 else today + timedelta(days=days_until_saturday)

# ══════════════════════════════════════════════════════════════════════════════
# PDF GENERATOR
# ══════════════════════════════════════════════════════════════════════════════
NAVY      = colors.HexColor("#0f1724")
BLUE      = colors.HexColor("#1a3a5c")
SLATE     = colors.HexColor("#64748b")
LIGHT     = colors.HexColor("#f1f5f9")
WHITE     = colors.white
BORDER    = colors.HexColor("#e2e8f0")

def generate_pdf_report(df, week_ending, title, prepared_by="Vista Staffing Solutions",
                         include_detail=True, include_notes=True) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter,
                             rightMargin=0.65*inch, leftMargin=0.65*inch,
                             topMargin=0.65*inch, bottomMargin=0.65*inch)

    header_style   = ParagraphStyle("h",  fontName="Helvetica-Bold", fontSize=18, textColor=WHITE,   leading=22, alignment=TA_LEFT)
    sub_style      = ParagraphStyle("s",  fontName="Helvetica",      fontSize=9,  textColor=colors.HexColor("#94a3b8"), leading=13)
    section_style  = ParagraphStyle("sc", fontName="Helvetica-Bold", fontSize=11, textColor=NAVY,    leading=14, spaceBefore=14, spaceAfter=6)
    kpi_lbl_style  = ParagraphStyle("kl", fontName="Helvetica",      fontSize=7,  textColor=SLATE,   leading=10, alignment=TA_CENTER)
    kpi_val_style  = ParagraphStyle("kv", fontName="Helvetica-Bold", fontSize=16, textColor=NAVY,    leading=20, alignment=TA_CENTER)
    small_style    = ParagraphStyle("sm", fontName="Helvetica",      fontSize=8,  textColor=SLATE,   leading=11)
    right_style    = ParagraphStyle("r",  fontName="Helvetica",      fontSize=8,  textColor=SLATE,   leading=11, alignment=TA_RIGHT)

    story = []

    # Header banner
    hdr = Table([[
        Paragraph("<b>RMCHCS Weekly Spend Report</b>", header_style),
        Paragraph(f"Week Ending: {week_ending}<br/>Generated: {datetime.now().strftime('%B %d, %Y')}<br/>Prepared by: {prepared_by}", sub_style)
    ]], colWidths=[4.2*inch, 2.8*inch])
    hdr.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), NAVY),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING",(0,0), (0,-1),  18),
        ("RIGHTPADDING",(-1,0),(-1,-1),16),
        ("TOPPADDING", (0,0), (-1,-1), 16),
        ("BOTTOMPADDING",(0,0),(-1,-1),16),
    ]))
    story += [hdr, Spacer(1, 0.2*inch)]

    # KPI row
    total_spend     = df["total_spend"].sum()
    total_providers = df["provider_name"].nunique()
    total_days      = df["days_worked"].sum()
    total_specs     = df["specialty"].nunique()

    kpi_items = [("TOTAL SPEND", f"${total_spend:,.2f}"),
                 ("PROVIDERS",   str(total_providers)),
                 ("SPECIALTIES", str(total_specs)),
                 ("TOTAL DAYS",  f"{total_days:.1f}")]
    kpi_cells = [[Paragraph(l, kpi_lbl_style) for l,_ in kpi_items],
                 [Paragraph(v, kpi_val_style) for _,v in kpi_items]]
    kpi_tbl = Table(kpi_cells, colWidths=[1.74*inch]*4)
    kpi_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0,0),(-1,-1), LIGHT),
        ("BOX",        (0,0),(-1,-1), 1, BORDER),
        ("LINEAFTER",  (0,0),(2,-1),  0.5, BORDER),
        ("TOPPADDING", (0,0),(-1,-1), 10),
        ("BOTTOMPADDING",(0,0),(-1,-1),10),
        ("VALIGN",     (0,0),(-1,-1), "MIDDLE"),
    ]))
    story += [kpi_tbl, Spacer(1, 0.18*inch)]

    # Specialty summary
    story.append(Paragraph("Spend by Specialty", section_style))
    spec = df.groupby("specialty").agg(
        Providers=("provider_name","nunique"),
        Days=("days_worked","sum"),
        Spend=("total_spend","sum")
    ).reset_index().sort_values("Spend", ascending=False)

    spec_rows = [["Specialty","Providers","Days Worked","Total Spend","% of Week"]]
    for _, row in spec.iterrows():
        pct = (row["Spend"]/total_spend*100) if total_spend else 0
        spec_rows.append([row["specialty"], str(int(row["Providers"])),
                          f"{row['Days']:.1f}", f"${row['Spend']:,.2f}", f"{pct:.1f}%"])
    spec_rows.append(["TOTAL", str(total_providers), f"{total_days:.1f}", f"${total_spend:,.2f}", "100%"])

    st_tbl = Table(spec_rows, colWidths=[2.4*inch,1.0*inch,1.0*inch,1.3*inch,0.95*inch])
    st_tbl.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),BLUE), ("TEXTCOLOR",(0,0),(-1,0),WHITE),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"), ("FONTSIZE",(0,0),(-1,0),8),
        ("FONTNAME",(0,1),(-1,-1),"Helvetica"),     ("FONTSIZE",(0,1),(-1,-1),8),
        ("TEXTCOLOR",(0,1),(-1,-1),NAVY),
        ("ROWBACKGROUNDS",(0,1),(-1,-2),[WHITE,LIGHT]),
        ("BACKGROUND",(0,-1),(-1,-1),colors.HexColor("#dbeafe")),
        ("FONTNAME",(0,-1),(-1,-1),"Helvetica-Bold"),
        ("ALIGN",(1,0),(-1,-1),"CENTER"), ("ALIGN",(3,0),(3,-1),"RIGHT"),
        ("TOPPADDING",(0,0),(-1,-1),6), ("BOTTOMPADDING",(0,0),(-1,-1),6),
        ("LEFTPADDING",(0,0),(0,-1),10), ("GRID",(0,0),(-1,-1),0.4,BORDER),
    ]))
    story += [st_tbl, Spacer(1, 0.15*inch)]

    # Service line summary
    story.append(Paragraph("Spend by Service Line", section_style))
    sl = df.groupby("service_line").agg(
        Providers=("provider_name","nunique"),
        Spend=("total_spend","sum")
    ).reset_index().sort_values("Spend", ascending=False)

    sl_rows = [["Service Line","Providers","Total Spend","% of Week"]]
    for _, row in sl.iterrows():
        pct = (row["Spend"]/total_spend*100) if total_spend else 0
        sl_rows.append([row["service_line"], str(int(row["Providers"])),
                        f"${row['Spend']:,.2f}", f"{pct:.1f}%"])
    sl_tbl = Table(sl_rows, colWidths=[2.4*inch,1.1*inch,1.5*inch,1.1*inch])
    sl_tbl.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),BLUE), ("TEXTCOLOR",(0,0),(-1,0),WHITE),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"), ("FONTSIZE",(0,0),(-1,0),8),
        ("FONTNAME",(0,1),(-1,-1),"Helvetica"),     ("FONTSIZE",(0,1),(-1,-1),8),
        ("TEXTCOLOR",(0,1),(-1,-1),NAVY),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[WHITE,LIGHT]),
        ("ALIGN",(1,0),(-1,-1),"CENTER"), ("ALIGN",(2,0),(2,-1),"RIGHT"),
        ("TOPPADDING",(0,0),(-1,-1),6), ("BOTTOMPADDING",(0,0),(-1,-1),6),
        ("LEFTPADDING",(0,0),(0,-1),10), ("GRID",(0,0),(-1,-1),0.4,BORDER),
    ]))
    story.append(sl_tbl)

    # Provider detail
    if include_detail:
        story += [Spacer(1, 0.1*inch), Paragraph("Provider Detail", section_style)]
        hdr_row = [["Provider","Type","Specialty","Svc Line","Days","Rate/Day","Total"]]
        if include_notes:
            hdr_row[0].append("Notes")
        det_rows = []
        for _, row in df.sort_values("specialty").iterrows():
            sl = row.get("service_line","")
            if sl in HOURLY_SERVICE_LINES:
                qty  = f"{row['hours_worked']:.1f} hrs" if pd.notna(row.get("hours_worked")) else ""
                rate = f"${row['bill_rate']:,.2f}/hr"   if pd.notna(row.get("bill_rate"))    else ""
            else:
                qty  = f"{row['days_worked']:.1f} days" if pd.notna(row.get("days_worked")) else ""
                rate = f"${row['daily_rate']:,.2f}/day"  if pd.notna(row.get("daily_rate"))  else ""
            r = [row["provider_name"], row.get("provider_type",""),
                 row["specialty"], sl, qty, rate, f"${row['total_spend']:,.2f}"]
            if include_notes:
                r.append(str(row.get("notes","") or ""))
            det_rows.append(r)
        cw = [1.4*inch,0.7*inch,1.1*inch,0.75*inch,0.45*inch,0.75*inch,0.8*inch,1.2*inch] if include_notes \
             else [1.7*inch,0.8*inch,1.4*inch,0.9*inch,0.5*inch,0.85*inch,1.0*inch]
        det_tbl = Table(hdr_row+det_rows, colWidths=cw)
        det_tbl.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),BLUE), ("TEXTCOLOR",(0,0),(-1,0),WHITE),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"), ("FONTSIZE",(0,0),(-1,0),7.5),
            ("FONTNAME",(0,1),(-1,-1),"Helvetica"),     ("FONTSIZE",(0,1),(-1,-1),7.5),
            ("TEXTCOLOR",(0,1),(-1,-1),NAVY),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[WHITE,LIGHT]),
            ("ALIGN",(4,0),(-1,-1),"CENTER"), ("ALIGN",(6,0),(6,-1),"RIGHT"),
            ("TOPPADDING",(0,0),(-1,-1),5), ("BOTTOMPADDING",(0,0),(-1,-1),5),
            ("LEFTPADDING",(0,0),(0,-1),8), ("GRID",(0,0),(-1,-1),0.3,BORDER),
        ]))
        story.append(det_tbl)

    # Footer
    story += [Spacer(1,0.25*inch), HRFlowable(width="100%",thickness=0.5,color=BORDER), Spacer(1,0.08*inch)]
    ft = Table([[Paragraph(f"Confidential | {prepared_by}", small_style),
                 Paragraph(f"Report generated {datetime.now().strftime('%B %d, %Y at %I:%M %p')}", right_style)]],
               colWidths=[3.5*inch, 3.5*inch])
    ft.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),
                             ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0)]))
    story.append(ft)

    doc.build(story)
    return buf.getvalue()

# ══════════════════════════════════════════════════════════════════════════════
# PAGE CONFIG & CSS
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="RMCHCS Spend Tracker", page_icon="🏥",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=DM+Serif+Display&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    [data-testid="stSidebar"] { background: #0f1724; border-right: 1px solid #1e2d40; }
    [data-testid="stSidebar"] * { color: #cbd5e1 !important; }
    .main .block-container { background: #f8fafc; padding-top: 1.5rem; padding-bottom: 2rem; }
    .kpi-card { background:#fff; border-radius:12px; padding:1.4rem 1.6rem; border:1px solid #e2e8f0; box-shadow:0 1px 4px rgba(0,0,0,0.06); }
    .kpi-label { font-size:0.72rem; font-weight:600; letter-spacing:0.08em; text-transform:uppercase; color:#64748b; margin-bottom:0.3rem; }
    .kpi-value { font-family:'DM Serif Display',serif; font-size:2rem; color:#0f172a; line-height:1.1; }
    .kpi-delta-up { color:#10b981; font-weight:600; font-size:0.82rem; }
    .kpi-delta-neutral { color:#64748b; font-weight:500; font-size:0.82rem; }
    .section-header { font-family:'DM Serif Display',serif; font-size:1.25rem; color:#0f172a; margin-bottom:0.2rem; margin-top:0.5rem; }
    .section-sub { font-size:0.82rem; color:#64748b; margin-bottom:1rem; }
    .page-header { background:linear-gradient(135deg,#0f1724 0%,#1a3a5c 100%); border-radius:14px; padding:1.6rem 2rem; margin-bottom:1.5rem; display:flex; align-items:center; justify-content:space-between; }
    .page-header-title { font-family:'DM Serif Display',serif; font-size:1.6rem; color:#f1f5f9; margin:0; }
    .page-header-sub { font-size:0.82rem; color:#94a3b8; margin-top:0.2rem; }
    .page-header-badge { background:#1e3a5f; border:1px solid #2d5a8e; border-radius:8px; padding:0.5rem 1rem; font-size:0.78rem; color:#7dd3fc; font-weight:500; }
    .form-card { background:#fff; border-radius:12px; padding:1.5rem; border:1px solid #e2e8f0; box-shadow:0 1px 4px rgba(0,0,0,0.05); }
    [data-testid="metric-container"] { background:white; border-radius:12px; padding:1rem; border:1px solid #e2e8f0; }
    .stTabs [data-baseweb="tab-list"] { background:#f1f5f9; border-radius:10px; padding:4px; gap:2px; }
    .stTabs [data-baseweb="tab"] { border-radius:8px; font-weight:500; font-size:0.88rem; color:#64748b; }
    .stTabs [aria-selected="true"] { background:white !important; color:#0f172a !important; box-shadow:0 1px 3px rgba(0,0,0,0.1); }
    .stButton > button { border-radius:8px; font-weight:600; font-size:0.88rem; }
    #MainMenu {visibility:hidden;} footer {visibility:hidden;} header {visibility:hidden;}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# LOAD DATA & SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
df = load_data()

with st.sidebar:
    st.markdown("### 🏥 RMCHCS")
    st.markdown("**Spend Intelligence**")
    st.markdown("---")
    page = st.radio("Navigate", ["Dashboard", "Log Spend", "Manage Entries", "Generate Report"], index=0)
    st.markdown("---")
    if not df.empty:
        st.markdown("**Quick Stats**")
        st.markdown(f"Total Tracked: **${df['total_spend'].sum():,.0f}**")
        st.markdown(f"Entries: **{len(df)}**")
        st.markdown(f"Weeks Tracked: **{df['week_ending'].nunique()}**")
    st.markdown("---")
    st.markdown("<small style='color:#475569'>Vista Staffing Solutions<br>Internal Use Only</small>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
if page == "Dashboard":
    st.markdown("""
    <div class="page-header">
        <div>
            <div class="page-header-title">RMCHCS Spend Dashboard</div>
            <div class="page-header-sub">Rehoboth McKinley Christian Health Care Services · Gallup, NM</div>
        </div>
        <div class="page-header-badge">Vista Staffing Solutions</div>
    </div>""", unsafe_allow_html=True)

    if df.empty:
        st.info("No spend data yet. Go to **Log Spend** to add your first entry.")
        st.stop()

    col_f1, col_f2, col_f3 = st.columns(3)
    all_weeks = sorted(df["week_ending"].unique(), reverse=True)
    with col_f1:
        selected_week = st.selectbox("Filter by Week", ["All Weeks"] + list(all_weeks))
    with col_f2:
        selected_specialty = st.selectbox("Filter by Specialty", ["All Specialties"] + sorted(df["specialty"].unique().tolist()))
    with col_f3:
        selected_sl = st.selectbox("Filter by Service Line", ["All Service Lines"] + sorted(df["service_line"].unique().tolist()))

    filtered = df.copy()
    if selected_week != "All Weeks":      filtered = filtered[filtered["week_ending"] == selected_week]
    if selected_specialty != "All Specialties": filtered = filtered[filtered["specialty"] == selected_specialty]
    if selected_sl != "All Service Lines": filtered = filtered[filtered["service_line"] == selected_sl]

    total      = filtered["total_spend"].sum()
    providers  = filtered["provider_name"].nunique()
    specialties= filtered["specialty"].nunique()
    avg_rate   = filtered["daily_rate"].mean() if not filtered.empty else 0

    if len(all_weeks) >= 2 and selected_week == "All Weeks":
        tw = df[df["week_ending"]==all_weeks[0]]["total_spend"].sum()
        lw = df[df["week_ending"]==all_weeks[1]]["total_spend"].sum()
        dpct = ((tw-lw)/lw*100) if lw else 0
        dstr = f"{'▲' if dpct>=0 else '▼'} {abs(dpct):.1f}% vs last week"
        dcls = "kpi-delta-up" if dpct>=0 else "kpi-delta-neutral"
    else:
        dstr = f"{filtered['week_ending'].nunique()} week(s) shown"
        dcls = "kpi-delta-neutral"

    k1,k2,k3,k4 = st.columns(4)
    with k1:
        st.markdown(f'<div class="kpi-card"><div class="kpi-label">Total Spend</div><div class="kpi-value">${total:,.0f}</div><div class="{dcls}">{dstr}</div></div>', unsafe_allow_html=True)
    with k2:
        st.markdown(f'<div class="kpi-card"><div class="kpi-label">Active Providers</div><div class="kpi-value">{providers}</div><div class="kpi-delta-neutral">Unique in view</div></div>', unsafe_allow_html=True)
    with k3:
        st.markdown(f'<div class="kpi-card"><div class="kpi-label">Specialties Covered</div><div class="kpi-value">{specialties}</div><div class="kpi-delta-neutral">Across all service lines</div></div>', unsafe_allow_html=True)
    with k4:
        st.markdown(f'<div class="kpi-card"><div class="kpi-label">Avg Daily Rate</div><div class="kpi-value">${avg_rate:,.0f}</div><div class="kpi-delta-neutral">Per provider per day</div></div>', unsafe_allow_html=True)

    st.markdown("<div style='margin-top:1.5rem'></div>", unsafe_allow_html=True)

    # Row 1: weekly trend + specialty donut
    c1, c2 = st.columns([3,2])
    with c1:
        st.markdown('<div class="section-header">Weekly Spend Trend</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-sub">Total spend logged per week ending date</div>', unsafe_allow_html=True)
        weekly = df.groupby("week_ending")["total_spend"].sum().reset_index().sort_values("week_ending")
        fig = go.Figure()
        fig.add_trace(go.Bar(x=weekly["week_ending"], y=weekly["total_spend"], marker_color="#1a3a5c", marker_line_width=0,
                             hovertemplate="<b>Week Ending:</b> %{x}<br><b>Spend:</b> $%{y:,.0f}<extra></extra>"))
        fig.add_trace(go.Scatter(x=weekly["week_ending"], y=weekly["total_spend"], mode="lines+markers",
                                  line=dict(color="#3b82f6", width=2.5), marker=dict(size=6), hoverinfo="skip"))
        fig.update_layout(plot_bgcolor="white", paper_bgcolor="white", margin=dict(l=0,r=0,t=10,b=0), height=280,
                          xaxis=dict(showgrid=False, tickfont=dict(size=11)),
                          yaxis=dict(showgrid=True, gridcolor="#f1f5f9", tickprefix="$", tickfont=dict(size=11)), showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.markdown('<div class="section-header">Spend by Specialty</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-sub">Share of total spend per specialty</div>', unsafe_allow_html=True)
        spec_spend = filtered.groupby("specialty")["total_spend"].sum().reset_index().sort_values("total_spend", ascending=False)
        fig2 = go.Figure(go.Pie(labels=spec_spend["specialty"], values=spec_spend["total_spend"], hole=0.52,
                                 marker=dict(colors=["#0f1724","#1a3a5c","#2d5a8e","#3b82f6","#60a5fa","#93c5fd","#bfdbfe"][:len(spec_spend)]),
                                 textinfo="percent",
                                 hovertemplate="<b>%{label}</b><br>$%{value:,.0f}<br>%{percent}<extra></extra>"))
        fig2.update_layout(plot_bgcolor="white", paper_bgcolor="white", margin=dict(l=0,r=0,t=10,b=0), height=280,
                           legend=dict(font=dict(size=11)))
        st.plotly_chart(fig2, use_container_width=True)

    # Row 2: provider bar + service line bar
    c3, c4 = st.columns([3,2])
    with c3:
        st.markdown('<div class="section-header">Spend by Provider</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-sub">Total spend per provider in current view</div>', unsafe_allow_html=True)
        prov_spend = filtered.groupby("provider_name")["total_spend"].sum().reset_index().sort_values("total_spend")
        fig3 = go.Figure(go.Bar(x=prov_spend["total_spend"], y=prov_spend["provider_name"], orientation="h",
                                 marker_color="#2d5a8e", marker_line_width=0,
                                 hovertemplate="<b>%{y}</b><br>$%{x:,.0f}<extra></extra>"))
        fig3.update_layout(plot_bgcolor="white", paper_bgcolor="white", margin=dict(l=0,r=10,t=10,b=0),
                           height=max(200, len(prov_spend)*38),
                           xaxis=dict(showgrid=True, gridcolor="#f1f5f9", tickprefix="$", tickfont=dict(size=11)),
                           yaxis=dict(showgrid=False, tickfont=dict(size=11)))
        st.plotly_chart(fig3, use_container_width=True)

    with c4:
        st.markdown('<div class="section-header">Service Line Breakdown</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-sub">Physician vs APP vs other lines</div>', unsafe_allow_html=True)
        sl_spend = filtered.groupby("service_line")["total_spend"].sum().reset_index().sort_values("total_spend", ascending=False)
        fig4 = go.Figure(go.Bar(x=sl_spend["service_line"], y=sl_spend["total_spend"],
                                 marker_color=["#1a3a5c","#3b82f6","#60a5fa","#93c5fd"][:len(sl_spend)],
                                 marker_line_width=0,
                                 hovertemplate="<b>%{x}</b><br>$%{y:,.0f}<extra></extra>"))
        fig4.update_layout(plot_bgcolor="white", paper_bgcolor="white", margin=dict(l=0,r=0,t=10,b=0), height=280,
                           xaxis=dict(showgrid=False, tickfont=dict(size=11)),
                           yaxis=dict(showgrid=True, gridcolor="#f1f5f9", tickprefix="$", tickfont=dict(size=11)))
        st.plotly_chart(fig4, use_container_width=True)

    st.markdown('<div class="section-header" style="margin-top:0.5rem">Entry Detail</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">All logged entries in current filter view</div>', unsafe_allow_html=True)
    disp = filtered[["week_ending","provider_name","specialty","service_line",
                       "days_worked","daily_rate","hours_worked","bill_rate",
                       "total_spend","notes"]].sort_values("week_ending", ascending=False).copy()
    def fmt_unit(row):
        if row["service_line"] in HOURLY_SERVICE_LINES:
            hrs  = row["hours_worked"] if pd.notna(row["hours_worked"]) else ""
            rate = f"${row['bill_rate']:,.2f}/hr" if pd.notna(row["bill_rate"]) else ""
            return pd.Series([hrs, rate])
        else:
            days = row["days_worked"] if pd.notna(row["days_worked"]) else ""
            rate = f"${row['daily_rate']:,.2f}/day" if pd.notna(row["daily_rate"]) else ""
            return pd.Series([days, rate])
    disp[["_qty","_rate"]] = disp.apply(fmt_unit, axis=1)
    disp["total_spend"] = disp["total_spend"].apply(lambda x: f"${x:,.2f}" if pd.notna(x) else "")
    disp_final = disp[["week_ending","provider_name","specialty","service_line","_qty","_rate","total_spend","notes"]].copy()
    disp_final.columns = ["Week Ending","Provider","Specialty","Service Line","Days / Hours","Rate","Total Spend","Notes"]
    st.dataframe(disp_final, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# LOG SPEND
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Log Spend":
    st.markdown("""
    <div class="page-header">
        <div>
            <div class="page-header-title">Log Weekly Spend</div>
            <div class="page-header-sub">Add a single entry manually or upload multiple rows at once</div>
        </div>
    </div>""", unsafe_allow_html=True)

    tab_manual, tab_bulk = st.tabs(["Single Entry", "Bulk Upload (CSV)"])

    with tab_manual:
        st.markdown("<div style='height:0.75rem'></div>", unsafe_allow_html=True)
        specialty_options = ["Emergency Medicine","Hospitalist/Internal Medicine","Family Medicine",
                             "OB/GYN","Surgery","Radiology","Anesthesiology","Psychiatry",
                             "Pediatrics","Cardiology","Orthopedics","Neurology","Other"]
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**Week & Provider**")
            week_ending   = st.date_input("Week Ending Date", value=get_week_ending())
            provider_name = st.text_input("Provider Name", placeholder="Dr. Jane Smith")
            provider_type = st.selectbox("Provider Type", ["Physician","APP","CRNA","NP","PA","RN","LPN","Tech","Other"])
        with col2:
            st.markdown("**Assignment Details**")
            specialty    = st.selectbox("Specialty", specialty_options)
            service_line = st.selectbox("Service Line", ["Physician","APP","Nursing","Allied Health","Other"])
            department   = st.text_input("Department / Unit", placeholder="e.g., ED, ICU, L&D")

        st.markdown("---")
        is_hourly = service_line in HOURLY_SERVICE_LINES

        if is_hourly:
            st.markdown("**Billing: Hourly** (Nursing / Allied Health)")
            c3,c4,c5 = st.columns(3)
            with c3: hours_worked = st.number_input("Hours Worked", min_value=0.5, value=40.0, step=0.5)
            with c4: bill_rate    = st.number_input("Bill Rate ($/hr)", min_value=0.0, value=75.0, step=1.0)
            with c5: st.metric("Calculated Total", f"${hours_worked*bill_rate:,.2f}")
            days_worked = None; daily_rate = None
            calc_total  = round(hours_worked * bill_rate, 2)
        else:
            st.markdown("**Billing: Daily** (Physician / APP / NP / PA)")
            c3,c4,c5 = st.columns(3)
            with c3: days_worked = st.number_input("Days Worked", min_value=0.5, max_value=7.0, value=5.0, step=0.5)
            with c4: daily_rate  = st.number_input("Daily Rate ($)", min_value=0.0, value=1800.0, step=50.0)
            with c5: st.metric("Calculated Total", f"${days_worked*daily_rate:,.2f}")
            hours_worked = None; bill_rate = None
            calc_total   = round(days_worked * daily_rate, 2)

        notes = st.text_area("Notes (optional)", placeholder="Contract type, extension, specialty notes...", height=80)
        st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)
        cb1, _ = st.columns([1,4])
        with cb1:
            if st.button("Save Entry", type="primary", use_container_width=True):
                if not provider_name.strip():
                    st.error("Provider name is required.")
                else:
                    save_entry({
                        "week_ending":  str(week_ending),
                        "provider_name":provider_name.strip(),
                        "provider_type":provider_type,
                        "specialty":    specialty,
                        "service_line": service_line,
                        "department":   department.strip(),
                        "days_worked":  days_worked,
                        "daily_rate":   daily_rate,
                        "hours_worked": hours_worked,
                        "bill_rate":    bill_rate,
                        "total_spend":  calc_total,
                        "notes":        notes.strip(),
                        "logged_at":    datetime.now().isoformat()
                    })
                    st.success(f"Saved. {provider_name} | {week_ending} | ${calc_total:,.2f}")
                    st.rerun()

    with tab_bulk:
        st.markdown("<div style='height:0.75rem'></div>", unsafe_allow_html=True)
        st.markdown('<div class="section-header">Step 1: Download the Template</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-sub">Fill this out in Excel or Google Sheets. Do not rename the column headers.</div>', unsafe_allow_html=True)
        template_df = pd.DataFrame({
            "week_ending":  ["2025-06-07","2025-06-07","2025-06-07"],
            "provider_name":["Dr. Jane Smith","Sarah Jones NP","Mary RN"],
            "provider_type":["Physician","NP","RN"],
            "specialty":    ["Emergency Medicine","Family Medicine","Med/Surg"],
            "service_line": ["Physician","APP","Nursing"],
            "department":   ["ED","Clinic","3 West"],
            "days_worked":  [5,3,""],
            "daily_rate":   [1800,950,""],
            "hours_worked": ["","",36],
            "bill_rate":    ["","",75],
            "notes":        ["Extended contract","",""]
        })
        st.download_button("Download CSV Template", template_df.to_csv(index=False).encode("utf-8"),
                           "RMCHCS_SpendUpload_Template.csv", "text/csv", type="primary")
        st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)
        st.markdown('<div class="section-header">Step 2: Upload Your Filled File</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-sub">CSV only. Total is calculated automatically from days x rate.</div>', unsafe_allow_html=True)
        uploaded = st.file_uploader("Choose CSV", type=["csv"], label_visibility="collapsed")
        if uploaded:
            try:
                udf = pd.read_csv(uploaded)
                required = ["week_ending","provider_name","provider_type","specialty","service_line"]
                missing  = [c for c in required if c not in udf.columns]
                if missing:
                    st.error(f"Missing columns: {', '.join(missing)}. Use the template above.")
                else:
                    for col in ["department","notes","days_worked","daily_rate","hours_worked","bill_rate"]:
                        if col not in udf.columns: udf[col] = ""
                    for col in ["days_worked","daily_rate","hours_worked","bill_rate"]:
                        udf[col] = pd.to_numeric(udf[col], errors="coerce")

                    def calc_row_total(row):
                        sl = str(row.get("service_line",""))
                        if sl in HOURLY_SERVICE_LINES:
                            if pd.notna(row["hours_worked"]) and pd.notna(row["bill_rate"]):
                                return round(row["hours_worked"] * row["bill_rate"], 2)
                        else:
                            if pd.notna(row["days_worked"]) and pd.notna(row["daily_rate"]):
                                return round(row["days_worked"] * row["daily_rate"], 2)
                        return None

                    udf["total_spend"] = udf.apply(calc_row_total, axis=1)
                    bad   = udf[udf["total_spend"].isnull()]
                    clean = udf[udf["total_spend"].notna()].copy()

                    st.markdown('<div class="section-header">Step 3: Preview Before Saving</div>', unsafe_allow_html=True)
                    if not bad.empty:
                        st.warning(f"{len(bad)} row(s) skipped. Physician/APP rows need days_worked + daily_rate. Nursing/Allied rows need hours_worked + bill_rate.")
                    if clean.empty:
                        st.error("No valid rows to import.")
                    else:
                        def fmt_prev_row(row):
                            sl = str(row.get("service_line",""))
                            if sl in HOURLY_SERVICE_LINES:
                                qty  = f"{row['hours_worked']:.1f} hrs" if pd.notna(row["hours_worked"]) else ""
                                rate = f"${row['bill_rate']:,.2f}/hr"   if pd.notna(row["bill_rate"])    else ""
                            else:
                                qty  = f"{row['days_worked']:.1f} days" if pd.notna(row["days_worked"]) else ""
                                rate = f"${row['daily_rate']:,.2f}/day"  if pd.notna(row["daily_rate"])  else ""
                            return pd.Series([qty, rate])
                        prev = clean.copy()
                        prev[["_qty","_rate"]] = prev.apply(fmt_prev_row, axis=1)
                        prev["total_spend"] = prev["total_spend"].apply(lambda x: f"${x:,.2f}")
                        prev_show = prev[["week_ending","provider_name","provider_type","specialty","service_line","_qty","_rate","total_spend","notes"]].copy()
                        prev_show.columns = ["Week Ending","Provider","Type","Specialty","Service Line","Qty","Rate","Total Spend","Notes"]
                        st.dataframe(prev_show, use_container_width=True, hide_index=True)
                        bulk_total = clean["total_spend"].sum()
                        st.markdown(f"**{len(clean)} rows ready | Combined total: ${bulk_total:,.2f}**")
                        cb2, _ = st.columns([1,4])
                        with cb2:
                            if st.button("Save All Entries", type="primary", use_container_width=True):
                                for _, row in clean.iterrows():
                                    sl = str(row.get("service_line",""))
                                    save_entry({
                                        "week_ending":  str(row["week_ending"]).strip(),
                                        "provider_name":str(row["provider_name"]).strip(),
                                        "provider_type":str(row.get("provider_type","")).strip(),
                                        "specialty":    str(row["specialty"]).strip(),
                                        "service_line": sl,
                                        "department":   str(row.get("department","")).strip(),
                                        "days_worked":  float(row["days_worked"]) if pd.notna(row.get("days_worked")) else None,
                                        "daily_rate":   float(row["daily_rate"])  if pd.notna(row.get("daily_rate"))  else None,
                                        "hours_worked": float(row["hours_worked"])if pd.notna(row.get("hours_worked"))else None,
                                        "bill_rate":    float(row["bill_rate"])   if pd.notna(row.get("bill_rate"))   else None,
                                        "total_spend":  float(row["total_spend"]),
                                        "notes":        str(row.get("notes","")).strip(),
                                        "logged_at":    datetime.now().isoformat()
                                    })
                                st.success(f"{len(clean)} entries saved. Total: ${bulk_total:,.2f}")
                                st.rerun()
            except Exception as e:
                st.error(f"Could not read file: {e}")

# ══════════════════════════════════════════════════════════════════════════════
# MANAGE ENTRIES
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Manage Entries":
    st.markdown("""
    <div class="page-header">
        <div>
            <div class="page-header-title">Manage Entries</div>
            <div class="page-header-sub">Review and delete logged spend entries</div>
        </div>
    </div>""", unsafe_allow_html=True)

    if df.empty:
        st.info("No entries logged yet.")
        st.stop()

    all_weeks = sorted(df["week_ending"].unique(), reverse=True)
    sel_week  = st.selectbox("Select Week", ["All Weeks"] + list(all_weeks))
    view_df   = df.copy() if sel_week=="All Weeks" else df[df["week_ending"]==sel_week].copy()
    view_df   = view_df.sort_values("week_ending", ascending=False).reset_index(drop=True)
    st.markdown(f"**{len(view_df)} entries** in current view")

    for _, row in view_df.iterrows():
        with st.expander(f"{row['week_ending']} | {row['provider_name']} | {row['specialty']} | ${row['total_spend']:,.2f}"):
            c1,c2,c3 = st.columns(3)
            c1.write(f"**Provider Type:** {row.get('provider_type','N/A')}")
            c2.write(f"**Service Line:** {row['service_line']}")
            c3.write(f"**Department:** {row.get('department','N/A')}")
            sl = row.get("service_line","")
            if sl in HOURLY_SERVICE_LINES:
                hrs  = row["hours_worked"] if pd.notna(row.get("hours_worked")) else "N/A"
                rate = f"${row['bill_rate']:,.2f}/hr" if pd.notna(row.get("bill_rate")) else "N/A"
                c1.write(f"**Hours Worked:** {hrs}")
                c2.write(f"**Bill Rate:** {rate}")
            else:
                days = row["days_worked"] if pd.notna(row.get("days_worked")) else "N/A"
                rate = f"${row['daily_rate']:,.2f}/day" if pd.notna(row.get("daily_rate")) else "N/A"
                c1.write(f"**Days Worked:** {days}")
                c2.write(f"**Daily Rate:** {rate}")
            c3.write(f"**Total Spend:** ${row['total_spend']:,.2f}")
            if row.get("notes"): st.write(f"**Notes:** {row['notes']}")
            if st.button("Delete Entry", key=f"del_{row['id']}", type="secondary"):
                delete_entry(row["id"])
                st.success("Entry deleted.")
                st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# GENERATE REPORT
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Generate Report":
    st.markdown("""
    <div class="page-header">
        <div>
            <div class="page-header-title">Generate Client Report</div>
            <div class="page-header-sub">Build a PDF spend summary to send to RMCHCS</div>
        </div>
    </div>""", unsafe_allow_html=True)

    if df.empty:
        st.info("No data available to report on yet.")
        st.stop()

    all_weeks = sorted(df["week_ending"].unique(), reverse=True)
    c1,c2 = st.columns(2)
    with c1:
        report_week  = st.selectbox("Select Week to Report", all_weeks)
        report_title = st.text_input("Report Title", value=f"Weekly Spend Report - Week Ending {report_week}")
    with c2:
        include_detail = st.checkbox("Include Provider Detail Table", value=True)
        include_notes  = st.checkbox("Include Notes Column", value=True)
        prepared_by    = st.text_input("Prepared By", value="Vista Staffing Solutions")

    st.markdown("---")
    week_data = df[df["week_ending"]==report_week]
    m1,m2,m3 = st.columns(3)
    m1.metric("Week Total Spend", f"${week_data['total_spend'].sum():,.2f}")
    m2.metric("Providers",  week_data["provider_name"].nunique())
    m3.metric("Specialties",week_data["specialty"].nunique())
    rpt_disp = week_data.copy()
    def fmt_report_row(row):
        sl = row.get("service_line","")
        if sl in HOURLY_SERVICE_LINES:
            qty  = f"{row['hours_worked']:.1f} hrs" if pd.notna(row.get("hours_worked")) else ""
            rate = f"${row['bill_rate']:,.2f}/hr"   if pd.notna(row.get("bill_rate"))    else ""
        else:
            qty  = f"{row['days_worked']:.1f} days" if pd.notna(row.get("days_worked")) else ""
            rate = f"${row['daily_rate']:,.2f}/day"  if pd.notna(row.get("daily_rate"))  else ""
        return pd.Series([qty, rate])
    rpt_disp[["_qty","_rate"]] = rpt_disp.apply(fmt_report_row, axis=1)
    rpt_disp["total_spend"] = rpt_disp["total_spend"].apply(lambda x: f"${x:,.2f}")
    st.dataframe(rpt_disp[["provider_name","specialty","service_line","_qty","_rate","total_spend"]].rename(
        columns={"provider_name":"Provider","specialty":"Specialty","service_line":"Service Line",
                 "_qty":"Days / Hours","_rate":"Rate","total_spend":"Total Spend"}),
        use_container_width=True, hide_index=True)

    st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)
    cg1, _ = st.columns([1,4])
    with cg1:
        if st.button("Generate PDF", type="primary", use_container_width=True):
            with st.spinner("Building report..."):
                pdf_bytes = generate_pdf_report(week_data, report_week, report_title,
                                                 prepared_by, include_detail, include_notes)
            st.download_button("Download PDF Report", pdf_bytes,
                               f"RMCHCS_SpendReport_{report_week}.pdf", "application/pdf", type="primary")
            st.success("Report ready. Click above to download.")
