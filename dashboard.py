"""
Attendance Dashboard  (eSSL Biometric)
========================================
Streamlit app — monthly attendance overview, top performers,
and individual employee deep-dive.

Run:  python -m streamlit run dashboard.py
"""

import json
import calendar
from datetime import datetime, date
from pathlib import Path

import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px

import config

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Attendance Dashboard",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    /* metric cards */
    .emp-card {
        background:#ffffff !important; border-radius:12px;
        padding:16px 18px; margin-bottom:10px;
        border-top:4px solid #0066cc;
        box-shadow:0 1px 4px rgba(0,0,0,.10);
        text-align:center;
    }
    .emp-card.green  { border-color:#28a745; }
    .emp-card.red    { border-color:#dc3545; }
    .emp-card.orange { border-color:#fd7e14; }
    .emp-card.purple { border-color:#6f42c1; }
    .emp-card.teal   { border-color:#20c997; }
    .emp-card.blue   { border-color:#0066cc; }

    .emp-val { font-size:26px; font-weight:800; margin:4px 0 2px; color:#1a1a2e !important; }
    .emp-lbl { font-size:12px; color:#555555 !important; margin:0; }
    .emp-sub { font-size:11px; color:#888888 !important; margin-top:2px; }

    /* leaderboard rows */
    .lb-row {
        display:flex; align-items:center; justify-content:space-between;
        padding:8px 12px; border-radius:8px; margin-bottom:5px;
        background:#f0f2f6 !important;
    }
    .lb-rank { font-size:18px; font-weight:700; color:#888888 !important; width:30px; }
    .lb-name { font-size:14px; font-weight:600; color:#1a1a2e !important; flex:1; padding:0 10px; }
    .lb-val  { font-size:14px; font-weight:700; color:#0066cc !important; white-space:nowrap; }

    [data-testid="stDataFrame"] { border-radius:8px; }
    div[data-testid="stTabs"] button { font-size:15px; font-weight:600; }
</style>
""", unsafe_allow_html=True)


# ── Constants ─────────────────────────────────────────────────────────────────
STATUS_LABELS = {
    "P":         "Present",
    "A":         "Absent",
    "WO":        "Week Off",
    "WOP":       "Week Off (Worked)",
    "HP":        "Half Day",
    "HD":        "Half Day",
    "\u00bdP":   "Half Day",
    "\ufffdP":   "Half Day",
    "L":         "Leave",
    "CL":        "Casual Leave",
    "EL":        "Earned Leave",
    "ML":        "Medical Leave",
    "SL":        "Sick Leave",
    "OD":        "On Duty",
    "H":         "Holiday",
}

STATUS_COLOR = {
    "Present":           "#28a745",
    "Absent":            "#dc3545",
    "Week Off":          "#adb5bd",
    "Week Off (Worked)": "#20c997",
    "Half Day":          "#17a2b8",
    "Leave":             "#ffc107",
    "Casual Leave":      "#ffc107",
    "Earned Leave":      "#20c997",
    "Medical Leave":     "#fd7e14",
    "Sick Leave":        "#fd7e14",
    "Holiday":           "#6f42c1",
    "On Duty":           "#0dcaf0",
    "Unknown":           "#888888",
}

MEDAL = ["🥇", "🥈", "🥉", "4️⃣", "5️⃣", "6️⃣", "7️⃣", "8️⃣", "9️⃣", "🔟"]


# ── Data helpers ──────────────────────────────────────────────────────────────
@st.cache_data(ttl=60)
def load_departments() -> dict:
    """Load Emp Code → Department mapping from departments.csv."""
    path = Path("departments.csv")
    if not path.exists():
        return {}
    try:
        df = pd.read_csv(path, dtype=str)
        df = df.fillna("").rename(columns=str.strip)
        mapping = {}
        for _, row in df.iterrows():
            code = str(row.get("Emp Code", "")).strip()
            dept = str(row.get("Department", "")).strip()
            if code:
                mapping[code] = dept if dept else "Unassigned"
        return mapping
    except Exception:
        return {}


def load_data() -> dict:
    """Load the full multi-month store from JSON."""
    path = Path(config.DATA_FILE)
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def load_month(year: int, month: int) -> dict:
    """Return the data dict for a specific year-month, or {} if not scraped yet."""
    store = load_data()
    if not store:
        return {}
    key = f"{year}-{month:02d}"
    # Multi-month format
    if "months" in store:
        return store["months"].get(key, {})
    # Legacy single-month format
    if store.get("year") == year and store.get("month") == month:
        return store
    return {}


def normalize_status(raw: str) -> str:
    if not raw:
        return "Unknown"
    if raw.strip() in STATUS_LABELS:
        return STATUS_LABELS[raw.strip()]
    r = raw.strip().upper()
    for code, label in STATUS_LABELS.items():
        if r == code.upper() or r == label.upper():
            return label
    if "PRESENT" in r:  return "Present"
    if "ABSENT"  in r:  return "Absent"
    if "HALF"    in r:  return "Half Day"
    if "WEEK"    in r:  return "Week Off"
    if "HOLIDAY" in r:  return "Holiday"
    if "LEAVE"   in r:  return "Leave"
    return raw.strip()


def prepare_df(records: list) -> pd.DataFrame:
    df = pd.DataFrame(records)
    if df.empty:
        return df
    if "Status" in df.columns:
        df["Status_Label"] = df["Status"].apply(normalize_status)
    if "Attendance Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Attendance Date"], format="%d %b %Y", errors="coerce")
        df["Day"]  = df["Date"].dt.day
    for col in ["Duration", "Over Time", "LateBy", "EarlyBy"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
    # Add Department from mapping
    dept_map = load_departments()
    if dept_map and "Emp Code" in df.columns:
        df["Department"] = df["Emp Code"].astype(str).map(dept_map).fillna("Unassigned")
    else:
        df["Department"] = "Unassigned"
    return df


SUMMARY_STATUS_COLS = [
    "Present", "Absent", "Week Off", "Week Off (Worked)",
    "Half Day", "Leave", "Casual Leave", "Earned Leave",
    "Medical Leave", "Sick Leave", "Holiday", "On Duty",
]

def build_summary(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "Emp Name" not in df.columns or "Status_Label" not in df.columns:
        return pd.DataFrame()
    grp = (
        df.groupby(["Emp Code", "Emp Name", "Status_Label"])
        .size().unstack(fill_value=0).reset_index()
    )
    for s in SUMMARY_STATUS_COLS:
        if s not in grp.columns:
            grp[s] = 0
    keep = ["Emp Code", "Emp Name"] + [c for c in SUMMARY_STATUS_COLS if c in grp.columns]
    grp = grp[keep]

    ot = (
        df.groupby(["Emp Code", "Emp Name"])
        .agg(
            Total_OT    =("Over Time", "sum"),
            Total_Late  =("LateBy",    "sum"),
            Late_Days   =("LateBy",    lambda x: (x > 0).sum()),
            Working_Days=("Duration",  lambda x: (x > 0).sum()),
            Avg_Duration=("Duration",  lambda x: int(x[x > 0].mean()) if (x > 0).any() else 0),
        )
        .reset_index()
    )
    grp = grp.merge(ot, on=["Emp Code", "Emp Name"], how="left")

    # Total calendar days in data
    total_days = df["Date"].nunique() if "Date" in df.columns else 1

    # Working days = total days - week offs (this is what attendance % should be based on)
    wo_days  = grp["Week Off"] if "Week Off" in grp.columns else 0
    wop_days = grp["Week Off (Worked)"] if "Week Off (Worked)" in grp.columns else 0
    grp["Total_Days"]        = total_days
    grp["Working_Days_Net"]  = (total_days - wo_days).clip(lower=1)  # exclude pure week offs
    grp["Attendance_Pct"]    = (grp["Present"] * 100 / grp["Working_Days_Net"]).round(1).clip(upper=100)
    grp["WO_Pct"]            = (wo_days * 100 / total_days).round(1)
    return grp


# ── Overall KPI row ───────────────────────────────────────────────────────────
def render_kpis(df: pd.DataFrame, summary: pd.DataFrame):
    total_emp    = df["Emp Name"].nunique() if "Emp Name" in df.columns else 0
    present      = (df["Status_Label"] == "Present").sum() if "Status_Label" in df.columns else 0
    absent       = (df["Status_Label"] == "Absent").sum()  if "Status_Label" in df.columns else 0
    wo           = (df["Status_Label"] == "Week Off").sum() if "Status_Label" in df.columns else 0
    total_ot_h   = round(df["Over Time"].sum() / 60, 1)    if "Over Time" in df.columns else 0
    perfect      = int((summary["Absent"] == 0).sum())      if not summary.empty else 0

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Total Employees",     total_emp)
    c2.metric("Present Entries",     present,  help="All present-day entries this month")
    c3.metric("Absent Entries",      absent)
    c4.metric("Week Offs",           wo)
    c5.metric("Total OT",            f"{total_ot_h} hrs")
    c6.metric("Perfect Attendance",  perfect,  help="Employees with 0 absents")


# ── Monthly summary table ─────────────────────────────────────────────────────
def render_summary_table(summary: pd.DataFrame):
    if summary.empty:
        st.warning("No data.")
        return
    show = ["Emp Code", "Emp Name", "Present", "Absent", "Half Day",
            "Leave", "Week Off", "Working_Days", "Attendance_Pct",
            "Total_OT", "Total_Late", "Late_Days"]
    cols = [c for c in show if c in summary.columns]
    disp = summary[cols].rename(columns={
        "Working_Days":   "Worked Days",
        "Attendance_Pct": "Attend %",
        "Total_OT":       "OT (min)",
        "Total_Late":     "Late (min)",
        "Late_Days":      "Late Days",
    }).sort_values("Emp Name", ignore_index=True)
    st.dataframe(disp, use_container_width=True, height=min(60 + 36 * len(disp), 680))


# ── Charts tab ────────────────────────────────────────────────────────────────
_DARK = dict(color="#333333")
_LAYOUT_BASE = dict(
    plot_bgcolor="white", paper_bgcolor="white",
    font=dict(color="#333333"),
    title_font=dict(color="#333333"),
    legend=dict(font=dict(color="#333333"), bgcolor="white",
                title_font=dict(color="#333333")),
)

def render_bar_chart(summary: pd.DataFrame):
    if summary.empty:
        return
    status_cols = [c for c in ["Present", "Absent", "Leave", "Half Day",
                                "Week Off", "Week Off (Worked)", "Holiday"]
                   if c in summary.columns]
    fig = go.Figure()
    for col in status_cols:
        fig.add_trace(go.Bar(
            name=col, x=summary["Emp Name"], y=summary[col],
            marker_color=STATUS_COLOR.get(col, "#888"),
        ))
    fig.update_layout(
        **_LAYOUT_BASE,
        barmode="stack", title="Monthly Attendance Breakdown",
        xaxis_tickangle=-40, yaxis_title="Days",
        height=420, margin=dict(b=130), legend_title="Status",
        xaxis=dict(tickfont=_DARK, title_font=_DARK, color="#333333"),
        yaxis=dict(tickfont=_DARK, title_font=_DARK, color="#333333"),
    )
    st.plotly_chart(fig, use_container_width=True)


def render_ot_chart(summary: pd.DataFrame):
    if summary.empty or "Total_OT" not in summary.columns:
        return
    s = summary[summary["Total_OT"] > 0].nlargest(15, "Total_OT")
    if s.empty:
        return
    s = s.copy()
    s["OT_hrs"] = (s["Total_OT"] / 60).round(1)
    fig = px.bar(s, x="Emp Name", y="OT_hrs",
                 title="Overtime by Employee (hours)",
                 color_discrete_sequence=["#fd7e14"],
                 labels={"OT_hrs": "OT (hrs)", "Emp Name": ""})
    fig.update_layout(
        **_LAYOUT_BASE,
        height=360, xaxis_tickangle=-35,
        xaxis=dict(tickfont=_DARK, title_font=_DARK, color="#333333"),
        yaxis=dict(tickfont=_DARK, title_font=_DARK, color="#333333"),
    )
    st.plotly_chart(fig, use_container_width=True)


def render_attendance_donut(df: pd.DataFrame):
    if "Status_Label" not in df.columns:
        return
    counts = df["Status_Label"].value_counts()
    fig = px.pie(
        names=counts.index, values=counts.values,
        color=counts.index,
        color_discrete_map=STATUS_COLOR,
        hole=0.5,
        title="Overall Status Distribution",
    )
    fig.update_traces(textfont=dict(color="#333333"))
    fig.update_layout(
        **_LAYOUT_BASE,
        height=340, margin=dict(t=50, b=10),
    )
    st.plotly_chart(fig, use_container_width=True)


# ── Top Performers tab ────────────────────────────────────────────────────────
def _leaderboard(title: str, icon: str, df_sorted: pd.DataFrame,
                 name_col: str, val_col: str, val_fmt=None, n: int = 10):
    """Render a ranked leaderboard card."""
    st.markdown(f"#### {icon} {title}")
    rows = df_sorted.head(n)
    if rows.empty:
        st.caption("No data.")
        return
    for i, (_, row) in enumerate(rows.iterrows()):
        medal = MEDAL[i] if i < len(MEDAL) else f"{i+1}."
        val   = row[val_col]
        val_s = val_fmt(val) if val_fmt else str(val)
        st.markdown(
            f'<div class="lb-row">'
            f'<span class="lb-rank">{medal}</span>'
            f'<span class="lb-name">{row[name_col]}</span>'
            f'<span class="lb-val">{val_s}</span>'
            f'</div>',
            unsafe_allow_html=True,
        )
    st.write("")


def render_top_performers(summary: pd.DataFrame):
    if summary.empty:
        st.info("No data.")
        return

    # ── Row 1 ─────────────────────────────────────────────────────────────────
    c1, c2, c3 = st.columns(3)

    with c1:
        # Best attendance %
        s = summary.sort_values("Attendance_Pct", ascending=False)
        _leaderboard("Best Attendance", "🏆", s, "Emp Name", "Attendance_Pct",
                     val_fmt=lambda v: f"{v}%")

    with c2:
        # Highest OT
        s = summary[summary["Total_OT"] > 0].sort_values("Total_OT", ascending=False)
        _leaderboard("Highest Overtime", "⏰", s, "Emp Name", "Total_OT",
                     val_fmt=lambda v: f"{v//60}h {v%60}m")

    with c3:
        # Most punctual (present days with 0 late minutes among those who came)
        s = summary[summary["Present"] > 0].sort_values(
            ["Total_Late", "Present"], ascending=[True, False]
        )
        _leaderboard("Most Punctual", "⚡", s, "Emp Name", "Total_Late",
                     val_fmt=lambda v: "Perfect" if v == 0 else f"{v} min late")

    st.divider()

    # ── Row 2 ─────────────────────────────────────────────────────────────────
    c4, c5, c6 = st.columns(3)

    with c4:
        # Perfect attendance (0 absents)
        perfect = summary[summary["Absent"] == 0].sort_values(
            "Present", ascending=False
        )
        _leaderboard("Zero Absents 🎯", "✅", perfect, "Emp Name", "Present",
                     val_fmt=lambda v: f"{v} days present")

    with c5:
        # Worked on week offs (dedication)
        if "Week Off (Worked)" in summary.columns:
            s = summary[summary["Week Off (Worked)"] > 0].sort_values(
                "Week Off (Worked)", ascending=False
            )
            _leaderboard("Worked on Week Off", "💪", s, "Emp Name",
                         "Week Off (Worked)",
                         val_fmt=lambda v: f"{v} day{'s' if v > 1 else ''}")
        else:
            st.caption("No week-off-worked data.")

    with c6:
        # Needs attention — most absent
        s = summary[summary["Absent"] > 0].sort_values("Absent", ascending=False)
        _leaderboard("Most Absent", "⚠️", s, "Emp Name", "Absent",
                     val_fmt=lambda v: f"{v} days")

    st.divider()

    # ── Scatter: Attendance % vs OT ───────────────────────────────────────────
    st.markdown("#### 📊 Attendance % vs Overtime")
    scatter_df = summary[summary["Present"] > 0].copy()
    scatter_df["OT_hrs"] = (scatter_df["Total_OT"] / 60).round(1)
    fig = px.scatter(
        scatter_df,
        x="Attendance_Pct", y="OT_hrs",
        text="Emp Name",
        size="Present",
        color="Absent",
        color_continuous_scale="RdYlGn_r",
        labels={"Attendance_Pct": "Attendance %", "OT_hrs": "Overtime (hrs)",
                "Absent": "Absent Days"},
        title="Each dot = one employee  |  Size = present days  |  Color = absent days",
        height=420,
    )
    fig.update_traces(
        textposition="top center",
        textfont=dict(size=9, color="#333333"),
    )
    fig.update_layout(
        **_LAYOUT_BASE,
        coloraxis_colorbar=dict(
            title=dict(text="Absent Days", font=dict(color="#333333")),
            tickfont=dict(color="#333333"),
        ),
        xaxis=dict(gridcolor="#eeeeee", title_font=_DARK, tickfont=_DARK, color="#333333"),
        yaxis=dict(gridcolor="#eeeeee", title_font=_DARK, tickfont=_DARK, color="#333333"),
    )
    st.plotly_chart(fig, use_container_width=True)


# ── Individual employee metrics ───────────────────────────────────────────────
def _metric_card(col, label: str, value: str, sub: str = "", color: str = "blue"):
    sub_html = f'<p class="emp-sub">{sub}</p>' if sub else ""
    col.markdown(
        f'<div class="emp-card {color}">'
        f'<p class="emp-val">{value}</p>'
        f'<p class="emp-lbl">{label}</p>'
        f'{sub_html}'
        f'</div>',
        unsafe_allow_html=True,
    )


def render_employee_metrics(summary: pd.DataFrame, df: pd.DataFrame, emp_name: str,
                             data_year: int, data_month: int):
    row = summary[summary["Emp Name"] == emp_name]
    if row.empty:
        return
    r = row.iloc[0]

    total_days   = int(r.get("Total_Days", df["Date"].nunique() if "Date" in df.columns else 1))
    present      = int(r.get("Present", 0))
    absent       = int(r.get("Absent", 0))
    wo           = int(r.get("Week Off", 0))
    wop          = int(r.get("Week Off (Worked)", 0))
    leave        = int(r.get("Leave", 0)) + int(r.get("Casual Leave", 0)) + \
                   int(r.get("Earned Leave", 0)) + int(r.get("Medical Leave", 0))
    ot_min       = int(r.get("Total_OT", 0))
    late_min     = int(r.get("Total_Late", 0))
    late_days    = int(r.get("Late_Days", 0))
    avg_dur      = int(r.get("Avg_Duration", 0))

    # Working days = total - week offs (correct base for attendance %)
    working_days = max(total_days - wo, 1)
    att_pct      = round(min(present * 100 / working_days, 100), 1)
    wo_pct       = round(wo * 100 / total_days, 1) if total_days else 0
    absent_pct   = round(absent * 100 / working_days, 1) if working_days else 0

    ot_str  = f"{ot_min//60}h {ot_min%60}m" if ot_min else "0h 0m"
    dur_str = f"{avg_dur//60}h {avg_dur%60}m" if avg_dur else "--"

    # ── Row 1: Attendance metrics ─────────────────────────────────────────────
    st.caption(f"📅 Total days: **{total_days}**  |  Week Offs: **{wo}**  |  "
               f"Working Days (excl. WO): **{working_days}**")

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    _metric_card(c1, "Attendance %",
                 f"{att_pct}%",
                 f"{present}/{working_days} working days",
                 "green" if att_pct >= 90 else "orange" if att_pct >= 75 else "red")
    _metric_card(c2, "Present Days",
                 str(present),
                 f"out of {working_days} working days", "green")
    _metric_card(c3, "Absent Days",
                 str(absent),
                 "🎯 Zero!" if absent == 0 else f"{absent_pct}% of working days",
                 "blue" if absent == 0 else "red")
    _metric_card(c4, "Week Off %",
                 f"{wo_pct}%",
                 f"{wo} days out of {total_days}", "blue")
    _metric_card(c5, "Worked on WO",
                 str(wop),
                 "extra dedication 💪" if wop > 0 else "none", "teal")
    _metric_card(c6, "Leave Days",
                 str(leave),
                 "days on leave", "orange")

    # ── Row 2: Performance metrics ────────────────────────────────────────────
    c7, c8, c9, c10 = st.columns(4)
    _metric_card(c7, "Overtime",
                 ot_str,
                 "No OT" if ot_min == 0 else f"{ot_min} min total", "orange")
    _metric_card(c8, "Late Arrivals",
                 str(late_days),
                 "Perfect ✅" if late_min == 0 else f"{late_min} min total",
                 "teal" if late_min == 0 else "orange")
    _metric_card(c9, "Avg Work Time",
                 dur_str,
                 "per present day", "purple")
    _metric_card(c10, "Half Days",
                 str(int(r.get("Half Day", 0))),
                 "half day entries", "blue")

    st.write("")


# ── Attendance Calendar (HTML grid) ──────────────────────────────────────────
def render_heatmap(df: pd.DataFrame, emp_name: str, year: int, month: int):
    emp_df = df[df["Emp Name"] == emp_name].copy() if "Emp Name" in df.columns else pd.DataFrame()
    if emp_df.empty or "Day" not in emp_df.columns or "Status_Label" not in emp_df.columns:
        return

    day_status = dict(zip(emp_df["Day"].astype(int), emp_df["Status_Label"]))
    _, days_in = calendar.monthrange(year, month)
    first_wd   = calendar.monthrange(year, month)[0]   # 0=Mon … 6=Sun

    # Status → background color (Present=green, Absent=red, etc.)
    CELL_BG = {
        "Present":           "#28a745",
        "Absent":            "#dc3545",
        "Week Off":          "#6c757d",
        "Week Off (Worked)": "#20c997",
        "Half Day":          "#17a2b8",
        "Leave":             "#ffc107",
        "Casual Leave":      "#ffc107",
        "Earned Leave":      "#6f9",
        "Medical Leave":     "#fd7e14",
        "Sick Leave":        "#fd7e14",
        "Holiday":           "#6f42c1",
        "On Duty":           "#0dcaf0",
    }
    CELL_TXT = {   # readable text color per background
        "Leave": "#333333", "Casual Leave": "#333333",
        "Earned Leave": "#333333",
    }

    short = {       # 2-letter abbreviation
        "Present": "P", "Absent": "A", "Week Off": "WO",
        "Week Off (Worked)": "WOP", "Half Day": "HD",
        "Leave": "L", "Casual Leave": "CL", "Earned Leave": "EL",
        "Medical Leave": "ML", "Sick Leave": "SL",
        "Holiday": "HO", "On Duty": "OD",
    }

    # ── Build HTML ────────────────────────────────────────────────────────────
    grid_style = (
        "display:grid;grid-template-columns:repeat(7,1fr);"
        "gap:5px;margin-bottom:10px;"
    )
    cell_base = (
        "border-radius:8px;padding:6px 2px;text-align:center;"
        "min-height:52px;display:flex;flex-direction:column;"
        "align-items:center;justify-content:center;"
    )

    html = f'<p style="font-weight:700;color:#333;font-size:15px;margin-bottom:6px;">'
    html += f'📅 {calendar.month_name[month]} {year}</p>'
    html += f'<div style="{grid_style}">'

    # Day headers
    for d in ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]:
        html += (f'<div style="text-align:center;font-weight:700;'
                 f'color:#555;font-size:12px;padding-bottom:4px;">{d}</div>')

    # Empty cells before first day
    for _ in range(first_wd):
        html += '<div></div>'

    # Day cells
    for d in range(1, days_in + 1):
        st_label = day_status.get(d, "")
        bg       = CELL_BG.get(st_label, "#e9ecef")
        tc       = CELL_TXT.get(st_label, "#ffffff") if st_label else "#999999"
        abbr     = short.get(st_label, st_label[:2] if st_label else "")
        html += (
            f'<div style="{cell_base}background:{bg};">'
            f'<span style="font-weight:700;font-size:14px;color:{tc};">{d}</span>'
            f'<span style="font-size:10px;color:{tc};margin-top:2px;">{abbr}</span>'
            f'</div>'
        )

    html += '</div>'

    # Legend — only show statuses that actually appear this month
    used = set(day_status.values())
    html += '<div style="display:flex;flex-wrap:wrap;gap:6px;margin-top:8px;">'
    for label, bg in CELL_BG.items():
        if label in used:
            tc = CELL_TXT.get(label, "#ffffff")
            html += (f'<span style="background:{bg};color:{tc};'
                     f'padding:3px 10px;border-radius:12px;font-size:11px;'
                     f'font-weight:600;">{label}</span>')
    html += '</div>'

    st.markdown(html, unsafe_allow_html=True)


# ── Daily detail table ────────────────────────────────────────────────────────
def render_daily_detail(df: pd.DataFrame, emp_name: str):
    emp_df = df[df["Emp Name"] == emp_name].copy() if "Emp Name" in df.columns else pd.DataFrame()
    if emp_df.empty:
        st.info("No records found.")
        return
    show   = ["Attendance Date", "Status_Label", "InTime", "OutTime",
              "Shift", "Duration", "Over Time", "LateBy", "EarlyBy", "PunchRecords"]
    cols   = [c for c in show if c in emp_df.columns]
    rename = {
        "Status_Label":    "Status",
        "Attendance Date": "Date",
        "Over Time":       "OT (min)",
        "LateBy":          "Late (min)",
        "EarlyBy":         "Early (min)",
        "PunchRecords":    "Punches",
    }
    disp = emp_df[cols].rename(columns=rename)
    if "Date" in disp.columns:
        disp = disp.sort_values("Date")
    st.dataframe(disp.reset_index(drop=True), use_container_width=True,
                 height=min(60 + 36 * len(disp), 600))


# ── Department tab ────────────────────────────────────────────────────────────
def render_department_tab(df: pd.DataFrame, summary: pd.DataFrame):
    if "Department" not in df.columns or df["Department"].eq("Unassigned").all():
        st.warning("No department data found. Please fill in **departments.csv** and push to GitHub.")
        st.info("Open `D:\\Spine HR\\departments.csv` in Excel, fill the Department column, save, then run:\n```\ngit add departments.csv\ngit commit -m 'add departments'\ngit push\n```")
        return

    dept_df = df.copy()

    # ── Dept summary table ────────────────────────────────────────────────────
    st.markdown("### 🏢 Department Summary")
    dept_grp = (
        dept_df.groupby("Department")
        .agg(
            Employees   =("Emp Name",    "nunique"),
            Present     =("Status_Label", lambda x: (x == "Present").sum()),
            Absent      =("Status_Label", lambda x: (x == "Absent").sum()),
            Half_Day    =("Status_Label", lambda x: (x == "Half Day").sum()),
            Leave       =("Status_Label", lambda x: x.isin(["Leave","Casual Leave","Earned Leave","Medical Leave","Sick Leave"]).sum()),
            Week_Off    =("Status_Label", lambda x: (x == "Week Off").sum()),
            Total_OT_min=("Over Time",   "sum"),
        )
        .reset_index()
    )
    dept_grp["OT (hrs)"]    = (dept_grp["Total_OT_min"] / 60).round(1)
    dept_grp["Attend %"]    = (
        dept_grp["Present"] * 100 /
        (dept_grp["Present"] + dept_grp["Absent"] + dept_grp["Half_Day"]).replace(0, 1)
    ).round(1)
    show = ["Department","Employees","Present","Absent","Half_Day","Leave","Week_Off","OT (hrs)","Attend %"]
    st.dataframe(dept_grp[show].sort_values("Department"), use_container_width=True)
    st.divider()

    # ── Chart row 1: Attendance % by dept ────────────────────────────────────
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### Attendance % by Department")
        fig = px.bar(
            dept_grp.sort_values("Attend %", ascending=True),
            x="Attend %", y="Department", orientation="h",
            color="Attend %", color_continuous_scale="RdYlGn",
            range_color=[50, 100], text="Attend %",
        )
        fig.update_traces(texttemplate="%{text}%", textposition="outside",
                          textfont=dict(color="#333333"))
        fig.update_layout(
            **_LAYOUT_BASE, height=max(300, len(dept_grp) * 45),
            xaxis=dict(tickfont=_DARK, title_font=_DARK),
            yaxis=dict(tickfont=_DARK, title_font=_DARK),
            coloraxis_colorbar=dict(title=dict(text="Attend %", font=_DARK), tickfont=_DARK),
            margin=dict(l=10, r=80, t=20, b=20),
        )
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.markdown("#### Employee Count by Department")
        fig2 = px.pie(
            dept_grp, names="Department", values="Employees",
            hole=0.4,
        )
        fig2.update_traces(textfont=dict(color="#333333"))
        fig2.update_layout(**_LAYOUT_BASE, height=max(300, len(dept_grp) * 45),
                           margin=dict(t=20, b=20))
        st.plotly_chart(fig2, use_container_width=True)

    st.divider()

    # ── Chart row 2: Stacked present/absent/leave by dept ────────────────────
    st.markdown("#### Attendance Breakdown by Department")
    melt_cols = ["Department", "Present", "Absent", "Half_Day", "Leave", "Week_Off"]
    melt_df = dept_grp[melt_cols].melt(id_vars="Department", var_name="Status", value_name="Days")
    color_map = {
        "Present": "#28a745", "Absent": "#dc3545", "Half_Day": "#17a2b8",
        "Leave": "#ffc107", "Week_Off": "#6c757d",
    }
    fig3 = px.bar(
        melt_df, x="Department", y="Days", color="Status",
        color_discrete_map=color_map, barmode="stack",
    )
    fig3.update_layout(
        **_LAYOUT_BASE, height=380, xaxis_tickangle=-30,
        xaxis=dict(tickfont=_DARK, title_font=_DARK),
        yaxis=dict(tickfont=_DARK, title_font=_DARK),
    )
    st.plotly_chart(fig3, use_container_width=True)

    st.divider()

    # ── OT by dept ────────────────────────────────────────────────────────────
    ot_df = dept_grp[dept_grp["Total_OT_min"] > 0].sort_values("OT (hrs)", ascending=False)
    if not ot_df.empty:
        st.markdown("#### Overtime by Department (hrs)")
        fig4 = px.bar(ot_df, x="Department", y="OT (hrs)",
                      color_discrete_sequence=["#fd7e14"], text="OT (hrs)")
        fig4.update_traces(textposition="outside", textfont=dict(color="#333333"))
        fig4.update_layout(
            **_LAYOUT_BASE, height=320, xaxis_tickangle=-20,
            xaxis=dict(tickfont=_DARK, title_font=_DARK),
            yaxis=dict(tickfont=_DARK, title_font=_DARK),
        )
        st.plotly_chart(fig4, use_container_width=True)


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    st.title("📋 Attendance Dashboard")
    today = date.today()

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("Controls")
        sel_month = st.selectbox(
            "Month", range(1, 13), index=today.month - 1,
            format_func=lambda m: calendar.month_name[m],
        )
        sel_year = st.number_input("Year", 2020, 2030, today.year)

        # Department filter
        dept_map = load_departments()
        all_depts = sorted(set(d for d in dept_map.values() if d and d != "Unassigned"))
        if all_depts:
            dept_options = ["All Departments"] + all_depts + ["Unassigned"]
            sel_dept = st.selectbox("Department", dept_options)
        else:
            sel_dept = "All Departments"
        st.divider()

        # ── Show Refresh only if eSSL machine is reachable (local only) ─────
        import socket as _sock
        def _essl_reachable():
            try:
                _sock.setdefaulttimeout(2)
                _sock.create_connection(("192.168.10.105", 89), timeout=2).close()
                return True
            except Exception:
                return False
        _local_mode = _essl_reachable()

        # ── Background scrape state ───────────────────────────────────────────
        if "scrape_proc" not in st.session_state:
            st.session_state.scrape_proc = None

        scraping = (
            st.session_state.scrape_proc is not None
            and st.session_state.scrape_proc.poll() is None
        )

        if not _local_mode:
            st.caption("📊 View-only mode — data updated by admin")

        elif scraping:
            st.info("Scraping in progress…")
            prog = load_data()
            done = len(prog.get("records", [])) if prog else 0
            st.caption(f"Records fetched so far: **{done}**")
            col1, col2 = st.columns(2)
            with col1:
                if st.button("Check progress"):
                    st.rerun()
            with col2:
                if st.button("Stop scrape", type="secondary"):
                    try:
                        st.session_state.scrape_proc.terminate()
                    except Exception:
                        pass
                    st.session_state.scrape_proc = None
                    st.rerun()
        else:
            # Check if just finished
            if st.session_state.scrape_proc is not None:
                rc = st.session_state.scrape_proc.poll()
                st.session_state.scrape_proc = None
                if rc == 0:
                    st.success("Scrape complete!")
                else:
                    st.warning("Scrape ended (check data below).")

            if st.button("🔄 Refresh from Biometric", type="primary"):
                import subprocess, sys
                from pathlib import Path as _P
                cmd = [
                    sys.executable,
                    str(_P(__file__).parent / "read_essl_data.py"),
                    "--year",  str(int(sel_year)),
                    "--month", str(int(sel_month)),
                    "--headless",
                ]
                proc = subprocess.Popen(cmd, cwd=str(_P(__file__).parent))
                st.session_state.scrape_proc = proc
                st.rerun()

        st.divider()
        st.caption(
            f"Source: eSSL eTimeTrackLite\n"
            f"File: `{config.DATA_FILE}`"
        )

    # ── Load data for selected month ──────────────────────────────────────────
    store = load_data()
    if not store:
        st.warning("No data found. Run **scrape.bat** or click Refresh.")
        st.stop()

    raw = load_month(int(sel_year), int(sel_month))

    # Show which months are available
    if "months" in store:
        available = sorted(store["months"].keys())
        st.caption(f"📁 Available months: **{', '.join(available)}**")

    if not raw:
        st.warning(
            f"No data for **{calendar.month_name[int(sel_month)]} {int(sel_year)}**. "
            f"Please scrape this month first."
        )
        st.stop()

    if "error" in raw and not raw.get("records"):
        st.error(f"Last fetch failed: {raw['error']}")
        st.stop()

    records = raw.get("records", [])
    if not records:
        st.warning("No records found for this month.")
        st.stop()

    data_month = int(raw.get("month", sel_month))
    data_year  = int(raw.get("year",  sel_year))
    fetched_at = raw.get("fetched_at", "")[:19].replace("T", " ")

    st.caption(
        f"**{calendar.month_name[data_month]} {data_year}**  ·  "
        f"{len(records)} records  ·  Last fetched: {fetched_at}  ·  "
        f"Source: {raw.get('source', 'eSSL')}"
    )

    df      = prepare_df(records)

    # Apply department filter
    if sel_dept != "All Departments" and "Department" in df.columns:
        df = df[df["Department"] == sel_dept].copy()

    summary = build_summary(df)

    # ── KPI bar ───────────────────────────────────────────────────────────────
    render_kpis(df, summary)
    st.divider()

    # ── Tabs ──────────────────────────────────────────────────────────────────
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 Monthly Summary",
        "📈 Charts",
        "🏆 Top Performers",
        "👤 Employee Detail",
        "🏢 Department",
    ])

    with tab1:
        render_summary_table(summary)
        with st.expander("Raw Records"):
            st.dataframe(
                df.drop(columns=["Date", "Day"], errors="ignore"),
                use_container_width=True,
            )

    with tab2:
        col_l, col_r = st.columns([2, 1])
        with col_l:
            render_bar_chart(summary)
        with col_r:
            render_attendance_donut(df)
        st.divider()
        render_ot_chart(summary)

    with tab3:
        render_top_performers(summary)

    with tab4:
        if "Emp Name" not in df.columns:
            st.info("No employee data.")
        else:
            employees = sorted(df["Emp Name"].dropna().unique().tolist())
            sel_emp   = st.selectbox("Select Employee", employees, key="emp_sel")

            # Employee metric cards
            render_employee_metrics(summary, df, sel_emp, data_year, data_month)
            st.divider()

            # Calendar + pie side by side
            left, right = st.columns([3, 2])
            with left:
                render_heatmap(df, sel_emp, data_year, data_month)
            with right:
                row = summary[summary["Emp Name"] == sel_emp]
                if not row.empty:
                    st.markdown("#### Status Breakdown")
                    status_cols = [c for c in SUMMARY_STATUS_COLS if c in row.columns
                                   and row.iloc[0][c] > 0]
                    if status_cols:
                        vals = row[status_cols].iloc[0]
                        fig  = px.pie(
                            names=status_cols, values=vals.tolist(),
                            color=status_cols,
                            color_discrete_map=STATUS_COLOR,
                            hole=0.45,
                        )
                        fig.update_traces(textfont=dict(color="#333333"))
                        fig.update_layout(
                            **_LAYOUT_BASE,
                            height=270, margin=dict(t=10, b=10, l=10, r=10),
                        )
                        st.plotly_chart(fig, use_container_width=True)

            st.markdown("#### Daily Records")
            render_daily_detail(df, sel_emp)

    with tab5:
        render_department_tab(df, summary)


if __name__ == "__main__":
    main()
