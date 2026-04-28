"""
Spine HR Attendance Scraper  (Production)
==========================================
Logs in, gets all employees via the GenericSearch popup, then for each employee
uses the popup-select flow (exactly as the user does manually) to load their
monthly attendance, and saves all records to attendance_data.json.

Run daily via Windows Task Scheduler using run_daily.bat
"""

import sys, io, json, time, traceback
from datetime import datetime, date
from pathlib import Path

# Fix Windows terminal encoding — only when run directly (not imported by the UI)
if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

import config

HR_ATTEN_URL = (
    "https://inovatix.spinehrm.in/Atten/MyAttendanceReport.aspx"
    "?Callfrom=pOGsF-4npfwEJZxkumNNVg&mnusr=menu__10104"
)
POPUP_URL = "https://inovatix.spinehrm.in/GenericSerach.aspx?CallFrom=Employee&reqfor=HR"

MONTH_NAMES = ["January","February","March","April","May","June",
               "July","August","September","October","November","December"]


# ── Driver setup ──────────────────────────────────────────────────────────
def make_driver():
    opts = Options()
    if config.HEADLESS:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--log-level=3")
    return webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=opts,
    )

def wait_for(driver, by, sel, timeout=None):
    return WebDriverWait(driver, timeout or config.CHROME_TIMEOUT).until(
        EC.presence_of_element_located((by, sel))
    )

def js_click(driver, el):
    driver.execute_script("arguments[0].click();", el)


# ── Login ─────────────────────────────────────────────────────────────────
def login(driver):
    print("[INFO] Logging in...")
    driver.get(config.SPINE_URL + "/login.aspx")
    time.sleep(3)
    driver.execute_script(
        "document.querySelectorAll('.Rsmodal,[id*=\"modal\"],[id*=\"Modal\"]')"
        ".forEach(e => e.style.display='none');"
    )
    time.sleep(0.5)
    wait_for(driver, By.ID, "txtUser").send_keys(config.USERNAME)
    driver.find_element(By.ID, "txtPassword").send_keys(config.PASSWORD)

    for sel_id, val in [("dpCompanyCodeList", None), ("dpConnectAs", config.LOGIN_FOR)]:
        try:
            s = Select(driver.find_element(By.ID, sel_id))
            opts = [o.text.strip() for o in s.options]
            if val and val in opts:
                s.select_by_visible_text(val)
            elif opts:
                s.select_by_index(0)
        except NoSuchElementException:
            pass

    driver.execute_script("document.getElementById('btnLogin').click();")
    time.sleep(5)
    if "login" in driver.current_url.lower():
        raise RuntimeError("Login failed — still on login page.")
    print(f"[INFO] Login OK  URL: {driver.current_url}")


# ── Get all employees ─────────────────────────────────────────────────────
def get_all_employees(driver):
    """
    Navigate to the GenericSearch popup URL (same session) and parse GridView1.
    Returns list of dicts with keys: name, ticket, code, branch, department, grade
    """
    print("[INFO] Fetching employee list from GenericSearch...")
    main_url = driver.current_url
    driver.get(POPUP_URL)
    time.sleep(3)

    employees = []
    try:
        table = wait_for(driver, By.ID, "GridView1", timeout=15)
        rows  = table.find_elements(By.TAG_NAME, "tr")

        headers = [th.text.strip() for th in rows[0].find_elements(By.TAG_NAME, "th")]
        if not headers:
            headers = [td.text.strip() for td in rows[0].find_elements(By.TAG_NAME, "td")]

        print(f"[INFO] Columns: {headers}")

        def ci(kw):
            for i, h in enumerate(headers):
                if kw.lower() in h.lower(): return i
            return -1

        ic_name  = ci("employee_name")
        ic_tkt   = ci("ticket")
        ic_code  = ci("employee_code")
        ic_br    = ci("branch")
        ic_dept  = ci("department")
        ic_grade = ci("grade")

        for row in rows[1:]:
            cells = [td.text.strip() for td in row.find_elements(By.TAG_NAME, "td")]
            if not cells or not any(cells):
                continue
            code = cells[ic_code] if 0 <= ic_code < len(cells) else ""
            if not code:
                continue
            employees.append({
                "name":       cells[ic_name]  if 0 <= ic_name  < len(cells) else "",
                "ticket":     cells[ic_tkt]   if 0 <= ic_tkt   < len(cells) else "",
                "code":       code,
                "branch":     cells[ic_br]    if 0 <= ic_br    < len(cells) else "",
                "department": cells[ic_dept]  if 0 <= ic_dept  < len(cells) else "",
                "grade":      cells[ic_grade] if 0 <= ic_grade < len(cells) else "",
            })
    except Exception as e:
        print(f"[ERROR] Employee list: {e}")

    print(f"[INFO] Found {len(employees)} employees.")
    # Return to attendance page
    driver.get(HR_ATTEN_URL)
    time.sleep(3)
    return employees


# ── Select employee via popup ─────────────────────────────────────────────
def select_employee_via_popup(driver, emp_index, main_window):
    """
    Open the GenericSearch popup, click the row button for emp_index,
    then click Apply — popup submits the main form and closes.
    Returns True if successful.
    """
    # Open popup (triggers window.open)
    driver.execute_script("openSearch('HdnEmp','txtHFld');")
    time.sleep(3)

    all_wins = driver.window_handles
    if len(all_wins) < 2:
        print("  [WARN] Popup did not open!")
        return False

    popup_win = [w for w in all_wins if w != main_window][0]
    driver.switch_to.window(popup_win)
    time.sleep(2)

    # Click the row's Select button
    try:
        driver.execute_script(
            f"__doPostBack('GridView1','Select${emp_index}');"
        )
        time.sleep(2)
    except Exception as e:
        print(f"  [WARN] Row select postback failed: {e}")

    # Click Apply
    try:
        apply_btn = driver.find_element(By.ID, "btnApply0")
        apply_btn.click()
        time.sleep(1)
    except Exception:
        try:
            driver.execute_script("btnApply_onclick();")
        except Exception as e2:
            print(f"  [WARN] Apply click failed: {e2}")

    # Popup should close — switch back to main window
    time.sleep(3)
    driver.switch_to.window(main_window)
    time.sleep(4)   # wait for main page to reload with employee data

    # Verify employee loaded
    try:
        nm = driver.find_element(By.ID, "ctl00_BodyContentPlaceHolder_lEmpName")
        loaded = nm.get_attribute("value") or ""
        return bool(loaded.strip())
    except Exception:
        return True  # can't verify but continue


# ── Set month / year and refresh ──────────────────────────────────────────
def set_month_and_refresh(driver, year, month):
    try:
        Select(driver.find_element(By.ID,
            "ctl00_BodyContentPlaceHolder_drpFromMonth")
        ).select_by_visible_text(MONTH_NAMES[month - 1])
    except Exception:
        pass
    try:
        Select(driver.find_element(By.ID,
            "ctl00_BodyContentPlaceHolder_drpFromYear")
        ).select_by_value(str(year))
    except Exception:
        pass
    # Refresh — validateData triggers __doPostBack for the Refresh button
    driver.execute_script(
        "if (typeof validateData==='function') { validateData(); } "
        "else { __doPostBack('ctl00$BodyContentPlaceHolder$btnRefresh',''); }"
    )
    time.sleep(5)


# ── Column name normalisation ─────────────────────────────────────────────
# Maps whatever Spine HR uses in its HTML table headers → dashboard names
_COL_MAP = {
    # Date
    "date":                    "Attendance Date",
    "attendance date":         "Attendance Date",
    "att date":                "Attendance Date",
    "attn date":               "Attendance Date",
    # Day name
    "day":                     "Day",
    "day name":                "Day",
    # Status
    "status":                  "Status",
    "attendance status":       "Status",
    "att status":              "Status",
    "present status":          "Status",
    # In / Out
    "in time":                 "InTime",
    "intime":                  "InTime",
    "punch in":                "InTime",
    "first in":                "InTime",
    "out time":                "OutTime",
    "outtime":                 "OutTime",
    "punch out":               "OutTime",
    "last out":                "OutTime",
    # Duration
    "duration":                "Duration",
    "working hours":           "Duration",
    "work hrs":                "Duration",
    "work duration":           "Duration",
    "hrs":                     "Duration",
    # Overtime
    "over time":               "Over Time",
    "overtime":                "Over Time",
    "ot":                      "Over Time",
    "ot hrs":                  "Over Time",
    "extra time":              "Over Time",
    # Late
    "late by":                 "LateBy",
    "lateby":                  "LateBy",
    "late":                    "LateBy",
    "late min":                "LateBy",
    # Early
    "early by":                "EarlyBy",
    "earlyby":                 "EarlyBy",
    "early":                   "EarlyBy",
    # Shift
    "shift":                   "Shift",
    "shift name":              "Shift",
    # Punch records
    "punch records":           "PunchRecords",
    "punch detail":            "PunchRecords",
    "punch details":           "PunchRecords",
    "punches":                 "PunchRecords",
    # Emp fields (added by scraper itself)
    "_emp_code":               "Emp Code",
    "_emp_name":               "Emp Name",
    "_emp_dept":               "Department",
    "_emp_branch":             "Branch",
}

def _normalize_columns(rec: dict) -> dict:
    """Rename raw Spine HR column keys to standard dashboard column names."""
    out = {}
    for k, v in rec.items():
        mapped = _COL_MAP.get(k.strip().lower(), k.strip())
        out[mapped] = v
    # Always ensure Emp Code / Emp Name are present under standard names
    if "Emp Code" not in out and "_emp_code" in rec:
        out["Emp Code"] = rec["_emp_code"]
    if "Emp Name" not in out and "_emp_name" in rec:
        out["Emp Name"] = rec["_emp_name"]
    return out


# ── Parse attendance table ────────────────────────────────────────────────
def parse_attendance_table(driver, emp_code, emp_name):
    """Parse tblPrintContent and return list of daily attendance dicts."""
    records = []
    try:
        table = driver.find_element(By.ID,
            "ctl00_BodyContentPlaceHolder_tblPrintContent")
    except NoSuchElementException:
        Path(f"debug_{emp_code}.html").write_text(
            driver.page_source, encoding="utf-8")
        return records

    rows = table.find_elements(By.TAG_NAME, "tr")
    headers = []

    for row in rows:
        ths  = row.find_elements(By.TAG_NAME, "th")
        tds  = row.find_elements(By.TAG_NAME, "td")

        if ths and not headers:
            headers = [h.text.strip() for h in ths]
            continue

        cells = [td.text.strip() for td in tds]
        if not cells or not any(cells) or len(cells) < 3:
            continue

        # Detect embedded header row
        if cells[0].lower() in ("date", "day", "sno", "sr.no", "sr no", "#"):
            headers = cells
            continue

        rec = ({headers[i]: cells[i]
                for i in range(min(len(headers), len(cells)))}
               if headers else
               {f"col_{i}": v for i, v in enumerate(cells)})

        rec["_emp_code"] = emp_code
        rec["_emp_name"] = emp_name
        records.append(_normalize_columns(rec))

    return records


# ── Master fetch function ─────────────────────────────────────────────────
def fetch_attendance(year=None, month=None):
    if year is None or month is None:
        today = date.today()
        year, month = today.year, today.month

    print(f"\n{'='*60}")
    print(f"  Spine HR Scraper  {year}-{month:02d}  ({MONTH_NAMES[month-1]} {year})")
    print(f"{'='*60}\n")

    driver = make_driver()
    output = {
        "fetched_at": datetime.now().isoformat(),
        "year":  year,
        "month": month,
        "employees": [],
        "records":   [],
        "errors":    [],
    }

    try:
        login(driver)

        # ── Get employee list ─────────────────────────────────────────
        employees = get_all_employees(driver)
        output["employees"] = employees

        if not employees:
            print("[ERROR] No employees found. Stopping.")
            return output

        # ── Navigate to HR attendance page ────────────────────────────
        driver.get(HR_ATTEN_URL)
        time.sleep(4)
        main_window = driver.current_window_handle

        total = len(employees)
        for idx, emp in enumerate(employees):
            code = emp["code"]
            name = emp["name"]
            print(f"\n[{idx+1:3d}/{total}] {code:10s}  {name}")

            try:
                # Select this employee via popup (exact same flow as manual)
                ok = select_employee_via_popup(driver, idx, main_window)
                if not ok:
                    print(f"         Employee may not have loaded correctly")

                # Set month / year and refresh
                set_month_and_refresh(driver, year, month)

                # Parse the attendance table
                records = parse_attendance_table(driver, code, name)
                # Stamp department & branch onto every record
                dept   = emp.get("department", "")
                branch = emp.get("branch", "")
                for rec in records:
                    if dept:
                        rec["Department"] = dept
                    if branch:
                        rec["Branch"] = branch
                if records:
                    output["records"].extend(records)
                    print(f"         {len(records)} day-records scraped")
                else:
                    print(f"         No records found")
                    output["errors"].append({
                        "code": code, "name": name, "error": "no_records"
                    })

            except Exception as e:
                msg = str(e)[:150]
                print(f"         ERROR: {msg}")
                output["errors"].append({"code": code, "name": name, "error": msg})
                # Re-navigate after error
                try:
                    driver.get(HR_ATTEN_URL)
                    time.sleep(3)
                    main_window = driver.current_window_handle
                except Exception:
                    pass

    except Exception as e:
        print(f"\n[FATAL] {e}")
        traceback.print_exc()
        output["fatal_error"] = str(e)
    finally:
        try:
            driver.quit()
        except Exception:
            pass

    # ── Save output (multi-month store — never overwrites other months) ────────
    out_path = Path(config.DATA_FILE)
    month_key = f"{year}-{month:02d}"

    # Load existing store
    store = {}
    if out_path.exists():
        try:
            store = json.loads(out_path.read_text(encoding="utf-8"))
        except Exception:
            store = {}

    # Migrate legacy single-month format → multi-month
    if store and "months" not in store:
        old_key = f"{store.get('year', year)}-{store.get('month', month):02d}"
        store = {"months": {old_key: store}}

    if "months" not in store:
        store["months"] = {}

    # Upsert this month; keep all other months untouched
    store["months"][month_key] = output
    store["last_updated"] = datetime.now().isoformat()

    out_path.write_text(
        json.dumps(store, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    months_saved = sorted(store["months"].keys())
    print(f"\n[DONE] Saved: {out_path.resolve()}")
    print(f"       Month     : {month_key}")
    print(f"       All months: {', '.join(months_saved)}")
    print(f"       Employees : {len(output['employees'])}")
    print(f"       Records   : {len(output['records'])}")
    print(f"       Errors    : {len(output['errors'])}")
    return output


if __name__ == "__main__":
    fetch_attendance()
