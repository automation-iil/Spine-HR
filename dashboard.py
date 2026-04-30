"""
Attendance Dashboard  (Spine HR)
=====================================
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
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    /* ── Sidebar ── */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1a1a2e 0%, #16213e 60%, #0f3460 100%) !important;
    }
    [data-testid="stSidebar"] * { color: #e0e0e0 !important; }
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stNumberInput label { color: #a0aec0 !important; font-size:12px !important; }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p { color:#a0aec0 !important; }
    [data-testid="stSidebar"] hr { border-color: #2d3748 !important; }

    /* ── Header banner ── */
    .dash-header {
        background: linear-gradient(135deg, #1a1a2e 0%, #0f3460 50%, #533483 100%);
        padding: 24px 32px; border-radius: 16px; margin-bottom: 20px;
        display: flex; align-items: center; justify-content: space-between;
        box-shadow: 0 4px 20px rgba(0,0,0,0.15);
    }
    .dash-title { font-size: 28px; font-weight: 800; color: #ffffff !important; margin: 0; }
    .dash-sub   { font-size: 13px; color: #a0c4ff !important; margin: 4px 0 0; }
    .dash-badge {
        background: rgba(255,255,255,0.15); border-radius: 20px;
        padding: 6px 16px; font-size: 13px; color: #ffffff !important;
        backdrop-filter: blur(10px); border: 1px solid rgba(255,255,255,0.2);
    }

    /* ── Tabs ── */
    div[data-testid="stTabs"] button {
        font-size: 14px; font-weight: 600;
        border-radius: 8px 8px 0 0;
        padding: 10px 20px;
        color: #4a5568 !important;
    }
    div[data-testid="stTabs"] button[aria-selected="true"] {
        color: #0f3460 !important;
        border-bottom: 3px solid #0f3460 !important;
        background: #eef2ff !important;
    }

    /* ── Employee metric cards ── */
    .emp-card {
        background: #ffffff; border-radius: 14px;
        padding: 18px 16px; margin-bottom: 10px;
        border-top: 4px solid #0f3460;
        box-shadow: 0 2px 12px rgba(0,0,0,0.07);
        text-align: center; transition: transform 0.2s;
    }
    .emp-card:hover { transform: translateY(-2px); box-shadow: 0 6px 20px rgba(0,0,0,0.12); }
    .emp-card.green  { border-color: #38a169; }
    .emp-card.red    { border-color: #e53e3e; }
    .emp-card.orange { border-color: #ed8936; }
    .emp-card.purple { border-color: #805ad5; }
    .emp-card.teal   { border-color: #319795; }
    .emp-card.blue   { border-color: #3182ce; }

    .emp-val { font-size: 28px; font-weight: 800; margin: 4px 0 2px; color: #1a1a2e !important; }
    .emp-lbl { font-size: 11px; font-weight: 600; text-transform: uppercase;
               letter-spacing: 0.5px; color: #718096 !important; margin: 0; }
    .emp-sub { font-size: 11px; color: #a0aec0 !important; margin-top: 3px; }

    /* ── Leaderboard rows ── */
    .lb-row {
        display: flex; align-items: center; justify-content: space-between;
        padding: 10px 14px; border-radius: 10px; margin-bottom: 6px;
        background: linear-gradient(135deg, #f7fafc 0%, #edf2f7 100%);
        border-left: 3px solid #e2e8f0; transition: all 0.2s;
    }
    .lb-row:hover { background: linear-gradient(135deg, #edf2f7 0%, #e2e8f0 100%);
                    border-left-color: #0f3460; }
    .lb-rank { font-size: 18px; width: 32px; }
    .lb-name { font-size: 14px; font-weight: 600; color: #2d3748 !important; flex: 1; padding: 0 12px; }
    .lb-val  { font-size: 13px; font-weight: 700; color: #0f3460 !important;
               background: #ebf8ff; padding: 3px 10px; border-radius: 20px; white-space: nowrap; }

    /* ── Section headers ── */
    .section-title {
        font-size: 16px; font-weight: 700; color: #1a1a2e !important;
        padding: 8px 0 6px; border-bottom: 2px solid #e2e8f0;
        margin-bottom: 14px;
    }

    /* ── DataFrames ── */
    [data-testid="stDataFrame"] { border-radius: 12px; overflow: hidden;
                                   box-shadow: 0 2px 8px rgba(0,0,0,0.06); }

    /* ── Warning / info boxes ── */
    [data-testid="stAlert"] { border-radius: 10px !important; }

    /* ── General ── */
    .block-container { padding-top: 1rem !important; }
    hr { border-color: #e2e8f0 !important; }
</style>
""", unsafe_allow_html=True)


# ── Constants ─────────────────────────────────────────────────────────────────
STATUS_LABELS = {
    # Standard codes
    "P":    "Present",
    "A":    "Absent",
    "WO":   "Week Off",
    "WOP":  "Week Off (Worked)",
    "HP":   "Half Day",
    "HD":   "Half Day",
    "\u00bdP":   "Half Day",
    "L":    "Leave",
    "CL":   "Casual Leave",
    "EL":   "Earned Leave",
    "ML":   "Medical Leave",
    "SL":   "Sick Leave",
    "OD":   "On Duty",
    "H":    "Holiday",
    # Spine HR codes
    "DP":   "Present",
    "ABS":  "Absent",
    "FP":   "Present",
    "LWP":  "Leave",
    "PL":   "Leave",
    "CO":   "On Duty",
    "WH":   "Holiday",
    "PH":   "Holiday",
    "NH":   "Holiday",
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
    """
    Load Emp Code → Department mapping.
    Primary source: spine_attendance.json employees list (scraped from Spine HR).
    Fallback: departments.csv (manual override).
    """
    mapping = {}

    # ── Primary: JSON employees list (always up-to-date from Spine HR) ────────
    try:
        store = json.loads(Path(config.DATA_FILE).read_text(encoding="utf-8"))
        employees = store.get("employees", [])
        # Multi-month format stores employees inside each month
        if not employees and "months" in store:
            for month_data in store["months"].values():
                employees = month_data.get("employees", [])
                if employees:
                    break
        for emp in employees:
            code = str(emp.get("code", "")).strip()
            dept = str(emp.get("department", "")).strip()
            if code:
                mapping[code] = dept if dept else "Unassigned"
    except Exception:
        pass

    # ── Fallback / override: departments.csv ──────────────────────────────────
    csv_path = Path("departments.csv")
    if csv_path.exists():
        try:
            csv_df = pd.read_csv(csv_path, dtype=str).fillna("").rename(columns=str.strip)
            for _, row in csv_df.iterrows():
                code = str(row.get("Emp Code", "")).strip()
                dept = str(row.get("Department", "")).strip()
                if code and dept:
                    mapping[code] = dept   # CSV overrides JSON
        except Exception:
            pass

    return mapping


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


def _hhmm_to_minutes(val) -> int:
    """Convert Spine HR HH.MM decimal string (e.g. '9.38' = 9h 38m) to minutes."""
    try:
        v = float(val)
        return int(v) * 60 + round((v % 1) * 100)
    except Exception:
        return 0


def _time_to_minutes(t_str: str):
    """Convert time string to minutes since midnight. Returns None on failure.
    Handles: '8:22 AM', '1:30 PM', '13:00', '09:22' (both 12h and 24h formats)."""
    try:
        t = str(t_str).strip().upper()
        if not t or t in ("", "---", "N/A", "0:00", "00:00"):
            return None
        if "AM" in t:
            h, m = map(int, t.replace("AM", "").strip().split(":"))
            return m if h == 12 else h * 60 + m
        if "PM" in t:
            h, m = map(int, t.replace("PM", "").strip().split(":"))
            return 12 * 60 + m if h == 12 else (h + 12) * 60 + m
        # 24-hour format: HH:MM or H:MM
        if ":" in t:
            parts = t.split(":")
            h, m = int(parts[0]), int(parts[1][:2])
            return h * 60 + m
    except Exception:
        return None


HALF_DAY_CUTOFF_MINUTES = 13 * 60  # 1:00 PM

# Office timing
OFFICE_START_MINUTES  = 9 * 60        # 9:00 AM
OFFICE_END_MINUTES    = 17 * 60 + 30  # 5:30 PM
LATE_GRACE_MINUTES    = 5             # 9:05 AM — grace period before marking late


def prepare_df(records: list) -> pd.DataFrame:
    df = pd.DataFrame(records)
    if df.empty:
        return df

    # ── Drop garbage header rows (rows without a real Attendance Date) ──────
    if "Attendance Date" in df.columns:
        df = df[df["Attendance Date"].notna() & (df["Attendance Date"] != "")].copy()
    if df.empty:
        return df

    # ── Derive Status from First Half + Second Half + Portion ───────────────
    # Spine HR stores each day as two halves; Portion=0.50 means mixed halves.
    if "Status" not in df.columns and "First Half" in df.columns:
        def _derive_status(row):
            fh = str(row.get("First Half", "")).strip()
            sh = str(row.get("Second Half", "")).strip()
            portion = str(row.get("Portion", "1.00")).strip()

            # Single-status day (second half is filler or same)
            if not sh or sh == "---" or sh == fh:
                return fh

            # Mixed halves → Half Day; prefer the "worked" status for reporting
            try:
                is_half = float(portion) <= 0.5
            except ValueError:
                is_half = True

            if not is_half:
                return fh  # treat full-day with mismatch as first half

            # Half-day: pick the most meaningful label
            halves = {fh, sh} - {"---", ""}
            if "WOP" in halves:
                return "WOP"   # worked on week off (even half day)
            if "DP" in halves:
                return "HD"    # half present
            if "CL" in halves or "PL" in halves or "LWP" in halves or "EL" in halves:
                return fh      # leave type
            return "HD"        # generic half day

        df["Status"] = df.apply(_derive_status, axis=1)

    # Duration: "Tot. Hrs." in HH.MM format → minutes
    if "Duration" not in df.columns and "Tot. Hrs." in df.columns:
        df["Duration"] = df["Tot. Hrs."].apply(_hhmm_to_minutes)

    # Over Time: "OT Hrs." in HH.MM format → minutes
    if "Over Time" not in df.columns and "OT Hrs." in df.columns:
        df["Over Time"] = df["OT Hrs."].apply(_hhmm_to_minutes)

    # LateBy: "LateMark" in HH.MM format → minutes
    if "LateBy" not in df.columns and "LateMark" in df.columns:
        df["LateBy"] = df["LateMark"].apply(_hhmm_to_minutes)

    # ── Half-day business rules ───────────────────────────────────────────────
    # Rule 1: InTime 1:00 PM – 2:00 PM → Half Day (late arrival)
    # Rule 2: OutTime ≤ 2:00 PM → Half Day (left early)
    # Rule 3: InTime 2:00 PM – 2:30 PM AND OutTime ≥ 5:15 PM → Half Day
    #         (came very late but stayed till end — still counts as half day)
    EARLY_OUT_CUTOFF  = 14 * 60        # 2:00 PM
    LATE_IN_START     = 14 * 60        # 2:00 PM
    LATE_IN_END       = 14 * 60 + 30   # 2:30 PM
    LATE_STAY_OUT_MIN = 17 * 60 + 15   # 5:15 PM
    PRESENT_STATUSES  = {"DP", "P", "FP", "PR", "PRES", "PRESENT"}
    ABS_STATUSES      = {"ABS", "A", "ABSENT"}
    if "InTime" in df.columns and "Status" in df.columns:
        def _apply_halfday_rules(row):
            status = str(row["Status"]).strip().upper()
            in_min  = _time_to_minutes(row.get("InTime",  ""))
            out_min = _time_to_minutes(row.get("OutTime", ""))

            # For ABS records: Spine HR can mark absent due to cutoff rules even
            # when the employee physically worked a half day. Override to HD if
            # the punch data shows a genuine half-day pattern.
            if status in ABS_STATUSES:
                # Rule A: morning arrival + left around 1–2 PM (morning half-day)
                if (in_min is not None and in_min < 12 * 60          # came before noon
                        and out_min is not None
                        and HALF_DAY_CUTOFF_MINUTES <= out_min <= EARLY_OUT_CUTOFF):
                    return "HD"
                # Rule B: came 1 PM–2 PM and has an out-punch (afternoon half-day)
                if (in_min is not None and HALF_DAY_CUTOFF_MINUTES <= in_min < LATE_IN_START
                        and out_min is not None and out_min > 0):
                    return "HD"
                # Rule C: came 2:00–2:30 PM and stayed past 5:15 PM
                if (in_min is not None and LATE_IN_START <= in_min <= LATE_IN_END
                        and out_min is not None and out_min >= LATE_STAY_OUT_MIN):
                    return "HD"
                return row["Status"]

            if status not in PRESENT_STATUSES:
                return row["Status"]

            shift_in = _time_to_minutes(row.get("Shift InTime", ""))
            if shift_in is not None and shift_in >= 12 * 60:
                return row["Status"]
            try:
                hrs = float(row.get("Tot. Hrs.", "0") or "0")
            except ValueError:
                hrs = 0
            if hrs <= 0:
                return row["Status"]

            # Rule 1: came between 1 PM and 2 PM
            if in_min is not None and HALF_DAY_CUTOFF_MINUTES <= in_min < LATE_IN_START:
                return "HD"
            # Rule 2: left at or before 2 PM
            if out_min is not None and out_min > 0 and out_min <= EARLY_OUT_CUTOFF:
                return "HD"
            # Rule 3: came 2:00–2:30 PM and stayed past 5:15 PM
            if (in_min is not None and LATE_IN_START <= in_min <= LATE_IN_END
                    and out_min is not None and out_min >= LATE_STAY_OUT_MIN):
                return "HD"
            return row["Status"]
        df["Status"] = df.apply(_apply_halfday_rules, axis=1)

    # ── Recalculate LateBy from InTime (office start = 9:00 AM, grace 5 min) ──
    # Spine HR's LateMark is often 0 or missing; we derive it ourselves.
    # Only applies to regular working days (not WO/Holiday rows).
    if "InTime" in df.columns:
        def _calc_late(row):
            status = str(row.get("Status", "")).upper()
            # Skip week-offs, holidays, absences
            if status in ("WO", "WOP", "WH", "PH", "NH", "ABS", ""):
                return 0
            in_min = _time_to_minutes(row.get("InTime", ""))
            if in_min is None or in_min == 0:
                return int(row.get("LateBy", 0) or 0)
            late = in_min - (OFFICE_START_MINUTES + LATE_GRACE_MINUTES)
            return max(late, 0)
        df["LateBy"] = df.apply(_calc_late, axis=1)

    # ── Normalise status ─────────────────────────────────────────────────────
    if "Status" in df.columns:
        df["Status_Label"] = df["Status"].apply(normalize_status)

    # ── Parse date — handle both 'dd-Mon-yy' (Spine) and 'dd Mon yyyy' ──────
    if "Attendance Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Attendance Date"], format="%d-%b-%y", errors="coerce")
        mask = df["Date"].isna()
        if mask.any():
            df.loc[mask, "Date"] = pd.to_datetime(
                df.loc[mask, "Attendance Date"], format="%d %b %Y", errors="coerce"
            )
        df["Day"] = df["Date"].dt.day

    # ── Numeric columns (already in minutes after mapping above) ────────────
    for col in ["Duration", "Over Time", "LateBy", "EarlyBy"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    # ── Department mapping ────────────────────────────────────────────────────
    # Records scraped after the dept-stamping fix already carry a Department
    # column from Spine HR. For older records (or gaps), fall back to the
    # mapping built from the employees list / departments.csv.
    dept_map = load_departments()
    if "Department" not in df.columns:
        df["Department"] = ""
    if dept_map and "Emp Code" in df.columns:
        mask = df["Department"].isna() | (df["Department"] == "")
        df.loc[mask, "Department"] = (
            df.loc[mask, "Emp Code"].astype(str).map(dept_map)
        )
    df["Department"] = df["Department"].fillna("Unassigned").replace("", "Unassigned")

    # ── Fill missing working days as Absent ──────────────────────────────────
    # A company working day is any date where at least one employee has a
    # regular attendance record (Present/Absent/Half Day/Leave/On Duty).
    # For each employee, any such date that falls within their active range
    # (first record ↔ last record) but has no record is treated as Absent.
    if "Date" in df.columns and "Status_Label" in df.columns and "Emp Code" in df.columns:
        regular_statuses = {
            "Present", "Absent", "Half Day", "Leave", "Casual Leave",
            "Earned Leave", "Medical Leave", "Sick Leave", "On Duty",
        }
        company_working_dates = set(
            df.loc[df["Status_Label"].isin(regular_statuses), "Date"].dropna().unique()
        )
        if company_working_dates:
            dept_map_fill = load_departments()
            missing_rows = []
            for emp_code, grp in df.groupby("Emp Code"):
                emp_name = grp["Emp Name"].iloc[0] if "Emp Name" in grp.columns else ""
                # Prefer dept already on the record; fall back to map
                dept_vals = grp["Department"].replace("Unassigned", "").replace("", pd.NA).dropna()
                dept = dept_vals.iloc[0] if not dept_vals.empty else \
                       dept_map_fill.get(str(emp_code), "Unassigned")
                emp_dates = set(grp["Date"].dropna())
                first_date = grp["Date"].min()
                last_date  = grp["Date"].max()
                for d in company_working_dates:
                    if first_date <= d <= last_date and d not in emp_dates:
                        missing_rows.append({
                            "Emp Code":       emp_code,
                            "Emp Name":       emp_name,
                            "Attendance Date": pd.Timestamp(d).strftime("%d-%b-%y"),
                            "Date":           d,
                            "Day":            pd.Timestamp(d).day,
                            "Status":         "ABS",
                            "Status_Label":   "Absent",
                            "Duration":       0,
                            "Over Time":      0,
                            "LateBy":         0,
                            "EarlyBy":        0,
                            "Department":     dept,
                        })
            if missing_rows:
                df = pd.concat([df, pd.DataFrame(missing_rows)], ignore_index=True)

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

    # Ensure numeric columns exist before aggregation
    for col in ["Over Time", "LateBy", "Duration"]:
        if col not in df.columns:
            df[col] = 0

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
    # ── Working days: days company was OPEN (exclude Sundays/holidays = days
    # where every employee either has no record or only WO/WOP records) ────────
    non_off_statuses = {"Present", "Absent", "Half Day", "Leave", "Casual Leave",
                        "Earned Leave", "Medical Leave", "Sick Leave", "On Duty"}
    company_open_days = (
        df[df["Status_Label"].isin(non_off_statuses)]["Date"].nunique()
        if "Date" in df.columns and "Status_Label" in df.columns
        else total_days
    )
    company_open_days = max(company_open_days, 1)

    half_col = grp["Half Day"] if "Half Day" in grp.columns else 0
    wop_col  = grp["Week Off (Worked)"] if "Week Off (Worked)" in grp.columns else 0

    grp["Total_Days"]        = company_open_days
    grp["Working_Days_Net"]  = (company_open_days - wo_days).clip(lower=1)
    # Present + 0.5 × Half Day (half days count as half attendance)
    grp["Effective_Present"] = (grp["Present"] + half_col * 0.5).round(1)
    grp["Attendance_Pct"]    = (
        grp["Effective_Present"] * 100 / grp["Working_Days_Net"]
    ).round(1).clip(upper=100)
    grp["WO_Pct"]            = (wo_days * 100 / company_open_days).round(1)
    return grp


# ── Overall KPI row ───────────────────────────────────────────────────────────
def _kpi_card(col, icon, label, value, color="#0f3460", bg="#f0f4ff"):
    """Render a compact KPI card with auto-sizing value text."""
    val_str = str(value)
    # Shrink font for longer values
    font_size = "22px" if len(val_str) <= 5 else "18px" if len(val_str) <= 8 else "15px"
    col.markdown(
        f'<div style="background:{bg};border-radius:14px;padding:16px 14px;'
        f'text-align:center;border-top:4px solid {color};'
        f'box-shadow:0 2px 10px rgba(0,0,0,0.07);height:90px;'
        f'display:flex;flex-direction:column;align-items:center;justify-content:center;">'
        f'<div style="font-size:20px;margin-bottom:4px;">{icon}</div>'
        f'<div style="font-size:{font_size};font-weight:800;color:#1a1a2e;line-height:1.2;">{val_str}</div>'
        f'<div style="font-size:11px;font-weight:600;color:#718096;text-transform:uppercase;'
        f'letter-spacing:0.4px;margin-top:3px;">{label}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

def render_kpis(df: pd.DataFrame, summary: pd.DataFrame):
    total_emp  = df["Emp Name"].nunique() if "Emp Name" in df.columns else 0
    present    = int((df["Status_Label"] == "Present").sum()) if "Status_Label" in df.columns else 0
    absent     = int((df["Status_Label"] == "Absent").sum())  if "Status_Label" in df.columns else 0
    wo         = int((df["Status_Label"] == "Week Off").sum()) if "Status_Label" in df.columns else 0
    total_ot_h = round(df["Over Time"].sum() / 60, 1) if "Over Time" in df.columns else 0
    perfect    = int((summary["Absent"] == 0).sum()) if not summary.empty else 0

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    _kpi_card(c1, "👥", "Total Employees",    total_emp,          "#0f3460", "#f0f4ff")
    _kpi_card(c2, "✅", "Present Entries",    present,            "#38a169", "#f0fff4")
    _kpi_card(c3, "❌", "Absent Entries",     absent,             "#e53e3e", "#fff5f5")
    _kpi_card(c4, "📅", "Week Offs",          wo,                 "#805ad5", "#faf5ff")
    _kpi_card(c5, "⏱️", "Total OT",          f"{total_ot_h} hrs","#ed8936", "#fffaf0")
    _kpi_card(c6, "🎯", "Perfect Attendance", perfect,            "#319795", "#e6fffa")


# ── Monthly summary table ─────────────────────────────────────────────────────
def render_summary_table(summary: pd.DataFrame):
    if summary.empty:
        st.warning("No data.")
        return
    show = ["Emp Code", "Emp Name", "Present", "Half Day", "Absent",
            "Leave", "Casual Leave", "Week Off", "Week Off (Worked)",
            "Effective_Present", "Working_Days_Net", "Attendance_Pct",
            "Total_OT", "Total_Late", "Late_Days"]
    cols = [c for c in show if c in summary.columns]
    disp = summary[cols].rename(columns={
        "Effective_Present":  "Eff. Present",
        "Working_Days_Net":   "Working Days",
        "Attendance_Pct":     "Attend %",
        "Week Off (Worked)":  "WOP",
        "Total_OT":           "OT (min)",
        "Total_Late":         "Late (min)",
        "Late_Days":          "Late Days",
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

def render_charts_tab(df: pd.DataFrame, summary: pd.DataFrame):
    if df.empty or summary.empty:
        st.info("No data available.")
        return

    # ── Row 1: Daily Trend + Status Donut ────────────────────────────────────
    c1, c2 = st.columns([3, 2])

    with c1:
        st.markdown("#### 📈 Daily Present Count")
        st.caption("How many employees were present each day this month")
        if "Day" in df.columns:
            daily = (
                df[df["Status_Label"] == "Present"]
                .groupby("Day")["Emp Name"].nunique()
                .reset_index()
                .rename(columns={"Emp Name": "Present"})
            )
            if not daily.empty:
                avg_p = int(daily["Present"].mean())
                fig = go.Figure()
                fig.add_trace(go.Scatter(
                    x=daily["Day"], y=daily["Present"],
                    mode="lines+markers",
                    line=dict(color="#0f3460", width=3),
                    marker=dict(size=7, color="#0f3460"),
                    fill="tozeroy",
                    fillcolor="rgba(15,52,96,0.08)",
                    name="Present",
                ))
                fig.add_hline(y=avg_p, line_dash="dot",
                              line_color="#ed8936", line_width=2,
                              annotation_text=f"Avg: {avg_p}",
                              annotation_font=dict(color="#ed8936", size=12))
                fig.update_layout(
                    **_LAYOUT_BASE, height=280,
                    xaxis=dict(title="Day of Month", tickfont=_DARK, dtick=2),
                    yaxis=dict(title="Employees Present", tickfont=_DARK),
                    margin=dict(t=20, b=40, l=50, r=20),
                )
                st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.markdown("#### 🍩 Status Distribution")
        st.caption("Overall breakdown of all attendance entries")
        counts = df["Status_Label"].value_counts()
        fig = px.pie(
            names=counts.index, values=counts.values,
            color=counts.index, color_discrete_map=STATUS_COLOR, hole=0.45,
        )
        fig.update_traces(
            textinfo="percent",
            textposition="inside",
            textfont=dict(color="#ffffff", size=12),
            insidetextorientation="radial",
        )
        fig.update_layout(
            **_LAYOUT_BASE, height=340,
            showlegend=True,
            margin=dict(t=20, b=20, l=10, r=130),
        )
        fig.update_layout(legend=dict(
            orientation="v",
            x=1.02, y=0.5,
            xanchor="left", yanchor="middle",
            font=dict(color="#333333", size=11),
            bgcolor="white",
        ))
        st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # ── Row 2: Attendance % brackets + Top Absent ─────────────────────────────
    c3, c4 = st.columns(2)

    with c3:
        st.markdown("#### 🎯 Attendance % Distribution")
        st.caption("How many employees fall in each attendance bracket")
        bins   = [0, 50, 70, 80, 90, 95, 101]
        labels = ["<50%", "50–70%", "70–80%", "80–90%", "90–95%", "95–100%"]
        colors = ["#e53e3e","#fc8181","#ed8936","#f6e05e","#68d391","#38a169"]
        s = summary.copy()
        s["Bracket"] = pd.cut(s["Attendance_Pct"], bins=bins, labels=labels, right=False)
        bracket_counts = s["Bracket"].value_counts().reindex(labels, fill_value=0).reset_index()
        bracket_counts.columns = ["Bracket", "Employees"]
        fig = px.bar(
            bracket_counts, x="Bracket", y="Employees",
            color="Bracket", color_discrete_sequence=colors,
            text="Employees",
        )
        fig.update_traces(textposition="outside", textfont=dict(color="#333333", size=13))
        fig.update_layout(
            **_LAYOUT_BASE, height=300, showlegend=False,
            xaxis=dict(tickfont=_DARK, title="Attendance %"),
            yaxis=dict(tickfont=_DARK, title="No. of Employees"),
            margin=dict(t=20, b=40),
        )
        st.plotly_chart(fig, use_container_width=True)

    with c4:
        st.markdown("#### ⚠️ Top Absentees")
        st.caption("Employees with highest absent days this month")
        top_absent = summary[summary["Absent"] > 0].nlargest(10, "Absent")[["Emp Name","Absent"]].copy()
        if not top_absent.empty:
            top_absent = top_absent.sort_values("Absent")
            fig = px.bar(
                top_absent, x="Absent", y="Emp Name", orientation="h",
                color="Absent", color_continuous_scale=["#fc8181","#e53e3e","#9b2335"],
                text="Absent",
            )
            fig.update_traces(textposition="outside", textfont=dict(color="#333333"))
            fig.update_layout(
                **_LAYOUT_BASE, height=300, showlegend=False,
                xaxis=dict(tickfont=_DARK, title="Absent Days"),
                yaxis=dict(tickfont=_DARK, title=""),
                coloraxis_showscale=False,
                margin=dict(t=20, b=40, l=10, r=60),
            )
            st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # ── Row 3: OT leaders + Late arrivals ────────────────────────────────────
    c5, c6 = st.columns(2)

    with c5:
        st.markdown("#### ⏱️ Overtime Leaders")
        st.caption("Top 10 employees with highest overtime this month")
        ot = summary[summary["Total_OT"] > 0].nlargest(10, "Total_OT").copy()
        if not ot.empty:
            ot["OT_hrs"] = (ot["Total_OT"] / 60).round(1)
            ot = ot.sort_values("OT_hrs")
            fig = px.bar(
                ot, x="OT_hrs", y="Emp Name", orientation="h",
                color="OT_hrs", color_continuous_scale=["#fbd38d","#ed8936","#c05621"],
                text="OT_hrs",
            )
            fig.update_traces(texttemplate="%{text}h", textposition="outside",
                              textfont=dict(color="#333333"))
            fig.update_layout(
                **_LAYOUT_BASE, height=300, showlegend=False,
                xaxis=dict(tickfont=_DARK, title="Hours"),
                yaxis=dict(tickfont=_DARK, title=""),
                coloraxis_showscale=False,
                margin=dict(t=20, b=40, l=10, r=60),
            )
            st.plotly_chart(fig, use_container_width=True)

    with c6:
        st.markdown("#### 🐢 Most Late Arrivals")
        st.caption("Top 10 employees by total late minutes")
        late = summary[summary["Total_Late"] > 0].nlargest(10, "Total_Late").copy()
        if not late.empty:
            late["Late_hrs"] = (late["Total_Late"] / 60).round(1)
            late = late.sort_values("Total_Late")
            fig = px.bar(
                late, x="Total_Late", y="Emp Name", orientation="h",
                color="Total_Late",
                color_continuous_scale=["#bee3f8","#4299e1","#2b6cb0"],
                text="Late_hrs",
            )
            fig.update_traces(texttemplate="%{text}h late", textposition="outside",
                              textfont=dict(color="#333333"))
            fig.update_layout(
                **_LAYOUT_BASE, height=300, showlegend=False,
                xaxis=dict(tickfont=_DARK, title="Total Late (min)"),
                yaxis=dict(tickfont=_DARK, title=""),
                coloraxis_showscale=False,
                margin=dict(t=20, b=40, l=10, r=80),
            )
            st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # ── Row 4: Weekday analysis ───────────────────────────────────────────────
    st.markdown("#### 📆 Attendance by Day of Week")
    st.caption("Which weekdays have the most absences vs presence")
    if "Date" in df.columns:
        wd = df.copy()
        wd["Weekday"] = wd["Date"].dt.day_name()
        order = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
        wd_grp = (
            wd.groupby(["Weekday","Status_Label"])
            .size().unstack(fill_value=0).reset_index()
        )
        for col in ["Present","Absent","Week Off","Half Day"]:
            if col not in wd_grp.columns:
                wd_grp[col] = 0
        wd_grp["Weekday"] = pd.Categorical(wd_grp["Weekday"], categories=order, ordered=True)
        wd_grp = wd_grp.sort_values("Weekday")

        fig = go.Figure()
        for col, color in [("Present","#38a169"),("Absent","#e53e3e"),
                            ("Half Day","#17a2b8"),("Week Off","#a0aec0")]:
            if col in wd_grp.columns:
                fig.add_trace(go.Bar(
                    name=col, x=wd_grp["Weekday"], y=wd_grp[col],
                    marker_color=color,
                ))
        fig.update_layout(
            **_LAYOUT_BASE, barmode="group", height=300,
            xaxis=dict(tickfont=_DARK),
            yaxis=dict(tickfont=_DARK, title="Count"),
            margin=dict(t=40, b=40),
        )
        fig.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02,
                                      font=dict(color="#333333")))
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
    if summary.empty or "Emp Name" not in summary.columns:
        return
    row = summary[summary["Emp Name"] == emp_name]
    if row.empty:
        return
    r = row.iloc[0]

    total_days   = int(r.get("Total_Days", df["Date"].nunique() if "Date" in df.columns else 1))
    present      = int(r.get("Present", 0))
    half_day     = int(r.get("Half Day", 0))
    absent       = int(r.get("Absent", 0))
    wo           = int(r.get("Week Off", 0))
    wop          = int(r.get("Week Off (Worked)", 0))
    leave        = int(r.get("Leave", 0)) + int(r.get("Casual Leave", 0)) + \
                   int(r.get("Earned Leave", 0)) + int(r.get("Medical Leave", 0))
    ot_min       = int(r.get("Total_OT", 0))
    late_min     = int(r.get("Total_Late", 0))
    late_days    = int(r.get("Late_Days", 0))
    avg_dur      = int(r.get("Avg_Duration", 0))

    # Working days = company open days − employee's pure week offs
    working_days  = int(r.get("Working_Days_Net", max(total_days - wo, 1)))
    eff_present   = round(present + half_day * 0.5, 1)   # half days count as 0.5
    att_pct       = round(min(eff_present * 100 / working_days, 100), 1)
    wo_pct        = round(wo * 100 / total_days, 1) if total_days else 0
    absent_pct    = round(absent * 100 / working_days, 1) if working_days else 0

    ot_str  = f"{ot_min//60}h {ot_min%60}m" if ot_min else "0h 0m"
    dur_str = f"{avg_dur//60}h {avg_dur%60}m" if avg_dur else "--"

    # ── Row 1: Attendance metrics ─────────────────────────────────────────────
    st.caption(f"📅 Company working days: **{total_days}**  |  Half Days: **{half_day}**  |  "
               f"Worked on Week Off: **{wop}**  |  Eff. Present: **{eff_present}**")

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    _metric_card(c1, "Attendance %",
                 f"{att_pct}%",
                 f"{eff_present}/{working_days} working days",
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
def render_heatmap(df: pd.DataFrame, emp_name: str, year: int, month: int,
                   missing_days: list = None):
    emp_df = df[df["Emp Name"] == emp_name].copy() if "Emp Name" in df.columns else pd.DataFrame()
    if emp_df.empty or "Day" not in emp_df.columns or "Status_Label" not in emp_df.columns:
        return

    day_status = dict(zip(emp_df["Day"].astype(int), emp_df["Status_Label"]))
    missing_set = set(missing_days or [])
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
        if d in missing_set:
            bg, tc, abbr = "#ff9800", "#ffffff", "?"
        else:
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
    if missing_set:
        html += (f'<span style="background:#ff9800;color:#ffffff;'
                 f'padding:3px 10px;border-radius:12px;font-size:11px;'
                 f'font-weight:600;">⚠️ No Data</span>')
    html += '</div>'

    st.markdown(html, unsafe_allow_html=True)


# ── Daily detail table ────────────────────────────────────────────────────────
def render_daily_detail(df: pd.DataFrame, emp_name: str):
    emp_df = df[df["Emp Name"] == emp_name].copy() if "Emp Name" in df.columns else pd.DataFrame()
    if emp_df.empty:
        st.info("No records found.")
        return
    show   = ["Attendance Date", "Day", "Status_Label", "InTime", "OutTime",
              "Shift Code", "Shift", "Duration", "Over Time", "LateBy", "EarlyBy",
              "Tot. Hrs.", "OT Hrs.", "LateMark", "Remarks", "PunchRecords"]
    cols   = [c for c in show if c in emp_df.columns]
    rename = {
        "Status_Label":    "Status",
        "Attendance Date": "Date",
        "Over Time":       "OT (min)",
        "OT Hrs.":         "OT",
        "LateBy":          "Late (min)",
        "LateMark":        "Late Mark",
        "Tot. Hrs.":       "Total Hrs",
        "EarlyBy":         "Early (min)",
        "PunchRecords":    "Punches",
        "Shift Code":      "Shift",
    }
    disp = emp_df[cols].rename(columns=rename)
    if "Date" in disp.columns:
        disp = disp.sort_values("Date")
    st.dataframe(disp.reset_index(drop=True), use_container_width=True,
                 height=min(60 + 36 * len(disp), 600))


# ── Department tab ────────────────────────────────────────────────────────────
def render_department_tab(df: pd.DataFrame, summary: pd.DataFrame):
    if "Department" not in df.columns or df["Department"].eq("Unassigned").all():
        st.warning(
            "No department data found in the scraped data. "
            "Re-scrape using **🔄 Refresh from Spine HR** — department names are "
            "fetched automatically from Spine HR's employee list."
        )
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
    today = date.today()

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown(
            '<p style="font-size:20px;font-weight:800;color:#ffffff !important;'
            'margin-bottom:4px;">⚙️ Controls</p>'
            '<p style="font-size:11px;color:#a0aec0 !important;margin-bottom:16px;">'
            'Indian Inovatix — Attendance</p>',
            unsafe_allow_html=True,
        )
        sel_month = st.selectbox(
            "Month", range(1, 13), index=today.month - 1,
            format_func=lambda m: calendar.month_name[m],
        )
        sel_year = st.number_input("Year", 2020, 2030, today.year)

        # Department filter — reads directly from Spine HR JSON (no CSV needed)
        dept_map = load_departments()
        all_depts = sorted(set(d for d in dept_map.values() if d and d != "Unassigned"))
        if all_depts:
            dept_options = ["All Departments"] + all_depts + ["Unassigned"]
            sel_dept = st.selectbox("🏢 Department", dept_options)
        else:
            sel_dept = "All Departments"
            st.caption("Departments load after first scrape.")
        st.divider()

        # ── Background scrape state ───────────────────────────────────────────
        if "scrape_proc" not in st.session_state:
            st.session_state.scrape_proc = None

        scraping = (
            st.session_state.scrape_proc is not None
            and st.session_state.scrape_proc.poll() is None
        )

        if scraping:
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

            if st.button("🔄 Refresh from Spine HR", type="primary"):
                import subprocess, sys
                from pathlib import Path as _P
                cmd = [
                    sys.executable,
                    str(_P(__file__).parent / "spine_scraper.py"),
                ]
                proc = subprocess.Popen(cmd, cwd=str(_P(__file__).parent))
                st.session_state.scrape_proc = proc
                st.rerun()

        st.divider()
        st.caption(
            f"Source: Spine HR (inovatix.spinehrm.in)\n"
            f"File: `{config.DATA_FILE}`"
        )

    # ── Load data for selected month ──────────────────────────────────────────
    store = load_data()
    if not store:
        st.warning("No data found. Click **Refresh from Spine HR** in the sidebar to scrape attendance data.")
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

    # ── Header banner ─────────────────────────────────────────────────────────
    import base64
    logo_path = Path(__file__).parent / "logo.png"
    logo_html = ""
    if logo_path.exists():
        logo_b64 = base64.b64encode(logo_path.read_bytes()).decode()
        logo_html = (
            f'<img src="data:image/png;base64,{logo_b64}" '
            f'style="width:72px;height:72px;border-radius:50%;'
            f'background:white;padding:4px;margin-right:18px;'
            f'box-shadow:0 2px 10px rgba(0,0,0,0.25);" />'
        )

    st.markdown(
        f'<div class="dash-header">'
        f'<div style="display:flex;align-items:center;">'
        f'{logo_html}'
        f'<div>'
        f'<p class="dash-title">Indian Inovatix Limited</p>'
        f'<p class="dash-sub">📋 Attendance Dashboard · Spine HR · {len(records)} records · Last fetched: {fetched_at}</p>'
        f'</div>'
        f'</div>'
        f'<div class="dash-badge">📅 {calendar.month_name[data_month]} {data_year}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    df      = prepare_df(records)

    # ── Detect missing dates ──────────────────────────────────────────────────
    def get_missing_dates(df, year, month):
        """Return list of dates in the month that have zero records."""
        today_d = date.today()
        _, days_in_month = calendar.monthrange(year, month)
        # Only check up to today if current month
        last_day = today_d.day if (year == today_d.year and month == today_d.month) \
                   else days_in_month
        expected = set(range(1, last_day + 1))
        if "Day" not in df.columns or df.empty:
            return sorted(expected)
        present_days = set(df["Day"].dropna().astype(int).unique())
        missing = expected - present_days
        return sorted(missing)

    all_df_full = prepare_df(records)   # unfiltered — check missing on full data
    missing_days = get_missing_dates(all_df_full, data_year, data_month)

    if missing_days:
        # Format as readable date strings e.g. "17 Apr, 18 Apr"
        missing_strs = [
            f"{d} {calendar.month_abbr[data_month]}"
            for d in missing_days
        ]
        st.warning(
            f"⚠️ **Missing Data Alert** — No attendance records found for "
            f"**{len(missing_days)} date(s)** in "
            f"{calendar.month_name[data_month]} {data_year}: "
            f"**{', '.join(missing_strs)}**  \n"
            f"These dates may be holidays, machine offline days, or not yet scraped."
        )

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
        render_charts_tab(df, summary)

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
                render_heatmap(df, sel_emp, data_year, data_month, missing_days)
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
