import os
import openpyxl
from datetime import datetime, date

# ── Paths ──────────────────────────────────────────────────────
BASE_DIR      = os.path.dirname(__file__)
USERS_PATH    = os.path.join(BASE_DIR, "users.xlsx")
PROJECTS_DIR  = os.path.join(BASE_DIR, "..", "data")

# ── Column mapping for project Excel header ────────────────────
# Label in col C/F, data in col D/G
HDR_READ = {
    "or_number":      "D1",
    "client_po":      "D3",
    "quote_number":   "G1",
    "sales_engineer": "G2",
    "sales_manager":  "G3",
    "po_value_lacs":  "D9",
    "customer_name":  "D10",
    "end_customer":   "D11",
    "consultant":     "D12",
    "project_desc":   "D13",
    "section":        "D14",
    "mfg_location":   "D15",
    "hw_efforts":     "D17",
    "std_panels":     "D18",
    "act_panels":     "D19",
    "sw_efforts":     "D20",
    "mfg_efforts":    "D21",
    "days_shop":      "AB2",
}

# Task rows: 9..55 (47 tasks)
TASK_START_ROW = 9
TASK_END_ROW   = 55

# Per-task columns
TASK_COL = {
    "phase":             "F",
    "owner":             "G",
    "task_code":         "H",
    "task_name":         "I",
    "rev1_start":        "J",
    "rev1_end":          "K",
    "rev2_start":        "L",
    "rev2_end":          "M",
    "rev3_start":        "N",
    "rev3_end":          "O",
    "orig_start":        "P",
    "orig_end":          "Q",
    "lead_time_days":    "R",
    "interlock":         "T",
    "effort_days":       "U",
    "current_start":     "V",
    "current_end":       "W",
    "actual_start":      "X",
    "actual_end":        "Y",
    "percent_complete":  "Z",
    "expected_percent":  "AA",
    "remark":            "AF",
}

OWNER_MAP = {
    "hw_tl":  "HW",
    "sw_tl":  "SW",
    "mfg_tl": "MFG",
    "pm":     None,   # PM sees all tasks
    "admin":  None,   # Admin sees all tasks
    "head":   None,   # Head sees all tasks
}

PHASE_NAMES = {
    "P1": "Engineering",
    "P2": "Procurement",
    "P3": "Manufacturing",
    "P4": "Software",
    "P5": "E&C"
}

def _fmt_date(val):
    if val is None:
        return None
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, str) and val.strip():
        return val.strip()
    return None

def _pct_to_int(val):
    """Convert Excel % (0.0–1.0 or 0–100) to int 0–100."""
    if val is None:
        return 0
    try:
        f = float(val)
        return int(f * 100) if f <= 1.0 else int(f)
    except (TypeError, ValueError):
        return 0

def _cell(ws, ref):
    v = ws[ref].value
    if isinstance(v, str):
        v = v.strip()
        return v if v else None
    return v


import json

INDEX_PATH = os.path.join(BASE_DIR, "..", "data", "index.json")

# ── Task JSON sidecar cache ────────────────────────────────────
def _tasks_cache_path(project_id):
    return os.path.join(PROJECTS_DIR, project_id + "_tasks.json")

def _write_tasks_cache(project_id, tasks):
    """Write tasks to JSON sidecar file."""
    try:
        with open(_tasks_cache_path(project_id), "w") as f:
            json.dump(tasks, f)
    except Exception as e:
        print(f"Task cache write failed for {project_id}: {e}")

def _read_tasks_cache(project_id):
    """Read tasks from JSON sidecar. Returns None if not found."""
    cache_path = _tasks_cache_path(project_id)
    if not os.path.exists(cache_path):
        return None
    try:
        with open(cache_path, "r") as f:
            return json.load(f)
    except Exception:
        return None

def build_index():
    """Scan all Excel files, write index.json and regenerate all task JSON caches."""
    projects = []
    for fname in list_project_files():
        fpath = os.path.join(PROJECTS_DIR, fname)
        project_id = os.path.splitext(fname)[0]
        try:
            p = read_project_header(fpath)
            projects.append(p)
            # Regenerate task cache
            tasks = get_tasks(project_id)
            print(f"  Indexed: {fname}")
        except Exception as e:
            print(f"  Skipped {fname}: {e}")
    with open(INDEX_PATH, "w") as f:
        json.dump(projects, f, indent=2)
    print(f"Index built: {len(projects)} projects")
    return projects

def read_index():
    """Read index.json. Returns list of project summaries."""
    if not os.path.exists(INDEX_PATH):
        return build_index()
    with open(INDEX_PATH, "r") as f:
        return json.load(f)

def update_index_entry(project_id):
    """Re-read one project Excel and update its entry in index.json."""
    fname = project_id + ".xlsx"
    fpath = os.path.join(PROJECTS_DIR, fname)
    if not os.path.exists(fpath):
        return
    try:
        updated = read_project_header(fpath)
        projects = read_index()
        found = False
        for i, p in enumerate(projects):
            if p["id"] == project_id:
                projects[i] = updated
                found = True
                break
        if not found:
            projects.append(updated)
        with open(INDEX_PATH, "w") as f:
            json.dump(projects, f, indent=2)
    except Exception as e:
        print(f"Index update failed for {project_id}: {e}")

# ── Users ──────────────────────────────────────────────────────
def get_all_users():
    if not os.path.exists(USERS_PATH):
        return []
    wb = openpyxl.load_workbook(USERS_PATH, data_only=True)
    ws = wb.active
    users = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0]:
            users.append({
                "name":          row[0],
                "short_name":    (row[1] or "").strip().upper(),
                "username":      row[2],
                "password_hash": row[3],
                "role":          row[4],
            })
    return users

def get_user_by_username(username):
    for u in get_all_users():
        if u["username"] == username:
            return u
    return None

def get_user_by_short_name(short_name):
    """Match Excel short form (e.g. KDM) to a user."""
    key = (short_name or "").strip().upper()
    for u in get_all_users():
        if u["short_name"] == key:
            return u
    return None

# ── Projects ───────────────────────────────────────────────────
def list_project_files():
    """Return list of .xlsx filenames in projects dir."""
    if not os.path.exists(PROJECTS_DIR):
        return []
    return [f for f in os.listdir(PROJECTS_DIR) if f.endswith(".xlsx")]

def read_project_header(filepath):
    """Read project header fields from Excel file."""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active

    p = {}
    for field, ref in HDR_READ.items():
        p[field] = _cell(ws, ref)

    # Read TL/PM assignments from dedicated rows
    # Based on Excel structure: row 7 col D = PM label area
    # We store PM name, HW TL, SW TL, MFG TL in specific cells
    # Using rows that have the name data
    # Read TL/PM assignments from stakeholder rows in col C
    p["pm_name"]     = _cell(ws, "C37") or ""
    p["hw_tl_name"]  = _cell(ws, "C38") or ""
    p["mfg_tl_name"] = _cell(ws, "C40") or ""
    p["sw_tl_name"]  = _cell(ws, "C41") or ""

    # Header info
    p["or_number"]       = _cell(ws, "D1") or ""
    p["master_or"]       = _cell(ws, "D2") or ""
    p["client_po"]       = _cell(ws, "D3") or ""
    p["sales_engineer"]  = _cell(ws, "I2") or ""
    p["sales_manager"]   = _cell(ws, "I3") or ""
    p["start_date_hdr"]  = _fmt_date(_cell(ws, "W1"))
    p["days_shop"]       = _cell(ws, "AB2")

    # Project info (col C rows 9-21)
    p["po_value"]        = _cell(ws, "C9")
    p["customer_name2"]  = _cell(ws, "C10") or ""
    p["end_customer"]    = _cell(ws, "C11") or ""
    p["consultant"]      = _cell(ws, "C12") or ""
    p["project_desc2"]   = _cell(ws, "C13") or ""
    p["section2"]        = _cell(ws, "C14") or ""
    p["mfg_loc"]         = _cell(ws, "C15") or "GON"
    p["hw_efforts"]      = _cell(ws, "C17")
    p["std_panels2"]     = _cell(ws, "C18")
    p["act_panels2"]     = _cell(ws, "C19")
    p["sw_efforts"]      = _cell(ws, "C20")
    p["mfg_efforts"]     = _cell(ws, "C21")

    # Customer & PM dates (col A=label, col C=value, rows 23-34)
    p["po_date"]         = _fmt_date(_cell(ws, "C23"))
    p["opf_recpt"]       = _fmt_date(_cell(ws, "C24"))
    p["hw_input"]        = _fmt_date(_cell(ws, "C25"))
    p["dwg_sub"]         = _fmt_date(_cell(ws, "C26"))
    p["dwg_appr"]        = _fmt_date(_cell(ws, "C27"))
    p["hw_fat"]          = _fmt_date(_cell(ws, "C28"))
    p["dispatch"]        = _fmt_date(_cell(ws, "C29"))
    p["sw_input"]        = _fmt_date(_cell(ws, "C30"))
    p["sw_fat"]          = _fmt_date(_cell(ws, "C31"))
    p["install"]         = _fmt_date(_cell(ws, "C32"))
    p["precomm"]         = _fmt_date(_cell(ws, "C33"))
    p["comm"]            = _fmt_date(_cell(ws, "C34"))

    # Stakeholders (col B=role, col C=initials, rows 36-42)
    p["sh_sales"]        = _cell(ws, "C36") or ""
    p["sh_hw"]           = _cell(ws, "C37") or ""
    p["sh_sw"]           = _cell(ws, "C38") or ""
    p["sh_byr"]          = _cell(ws, "C39") or ""
    p["sh_mfg"]          = _cell(ws, "C40") or ""
    p["sh_ec"]           = _cell(ws, "C41") or ""
    p["sh_ac"]           = _cell(ws, "C42") or ""

    # Misc
    p["balance_panels"]  = _cell(ws, "B43")
    p["panel_disp_act"]  = _cell(ws, "B44")
    p["est_va_pct"]      = _cell(ws, "B45")
    p["est_va"]          = _cell(ws, "B46")
    p["est_sm_pct"]      = _cell(ws, "B47")
    p["est_sm"]          = _cell(ws, "B48")
    p["act_va_pct"]      = _cell(ws, "B49")
    p["act_va"]          = _cell(ws, "B50")
    p["act_sm_pct"]      = _cell(ws, "B51")
    p["act_sm"]          = _cell(ws, "B52")
    p["reason_remark"]   = _cell(ws, "C53") or ""

    # LD date
    p["ld_date"]      = _fmt_date(_cell(ws, "AF1"))
    p["start_date"]   = _fmt_date(_cell(ws, "V9"))

    # Filename as ID
    p["filename"] = os.path.basename(filepath)
    p["id"]       = os.path.splitext(p["filename"])[0]

    # Overall % complete across all tasks
    p["overall_percent"] = calc_overall_percent(filepath)

    return p

def calc_overall_percent(filepath):
    """Read all task % complete values and return average as int 0-100."""
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.active
        values = []
        for row in range(TASK_START_ROW, TASK_END_ROW + 1):
            code = ws[f"H{row}"].value
            name = ws[f"I{row}"].value
            if not code or not name:
                continue
            values.append(_pct_to_int(ws[f"Z{row}"].value))
        if not values:
            return 0
        return round(sum(values) / len(values))
    except Exception:
        return 0

def get_all_projects():
    return read_index()

def get_projects_for_user(user):
    all_projects = get_all_projects()
    role = user["role"]
    name = user["name"]

    if role in ("admin", "head"):
        return all_projects

    short = user.get("short_name", "").upper()
    result = []
    for p in all_projects:
        if role == "pm"     and p.get("pm_name",     "").strip().upper() == short:
            result.append(p)
        elif role == "hw_tl"  and p.get("hw_tl_name",  "").strip().upper() == short:
            result.append(p)
        elif role == "sw_tl"  and p.get("sw_tl_name",  "").strip().upper() == short:
            result.append(p)
        elif role == "mfg_tl" and p.get("mfg_tl_name", "").strip().upper() == short:
            result.append(p)
    return result

def get_project_by_id(project_id):
    """Read project header from index.json instead of opening Excel."""
    projects = read_index()
    for p in projects:
        if p["id"] == project_id:
            return p
    return None

# ── Tasks ──────────────────────────────────────────────────────
def get_tasks(project_id, owner_filter=None):
    """Read all tasks from project Excel. Uses JSON cache if available."""
    fname = project_id + ".xlsx"
    fpath = os.path.join(PROJECTS_DIR, fname)
    if not os.path.exists(fpath):
        return []

    # Always try cache first — cache stores full task list, filter in memory
    cached = _read_tasks_cache(project_id)
    if cached is not None:
        if owner_filter:
            return [t for t in cached if t.get("owner") == owner_filter]
        return cached

    wb = openpyxl.load_workbook(fpath, data_only=True)
    ws = wb.active

    tasks = []
    for row in range(TASK_START_ROW, TASK_END_ROW + 1):
        phase     = ws[f"F{row}"].value
        owner     = ws[f"G{row}"].value
        task_code = ws[f"H{row}"].value
        task_name = ws[f"I{row}"].value

        if not task_code or not task_name:
            continue

        if owner_filter and owner != owner_filter:
            continue

        pct     = _pct_to_int(ws[f"Z{row}"].value)
        exp_pct = _pct_to_int(ws[f"AA{row}"].value)

        tasks.append({
            "row":              row,
            "phase":            phase,
            "phase_name":       PHASE_NAMES.get(phase, phase),
            "owner":            owner,
            "task_code":        task_code,
            "task_name":        str(task_name).strip(),
            "rev1_start":       _fmt_date(ws[f"J{row}"].value),
            "rev1_end":         _fmt_date(ws[f"K{row}"].value),
            "rev2_start":       _fmt_date(ws[f"L{row}"].value),
            "rev2_end":         _fmt_date(ws[f"M{row}"].value),
            "rev3_start":       _fmt_date(ws[f"N{row}"].value),
            "rev3_end":         _fmt_date(ws[f"O{row}"].value),
            "orig_start":       _fmt_date(ws[f"P{row}"].value),
            "orig_end":         _fmt_date(ws[f"Q{row}"].value),
            "current_start":    _fmt_date(ws[f"V{row}"].value),
            "current_end":      _fmt_date(ws[f"W{row}"].value),
            "actual_start":     _fmt_date(ws[f"X{row}"].value),
            "actual_end":       _fmt_date(ws[f"Y{row}"].value),
            "percent_complete": pct,
            "expected_percent": exp_pct,
            "lead_time_days":   ws[f"R{row}"].value,
            "effort_days":      ws[f"U{row}"].value,
            "interlock":        ws[f"T{row}"].value,
            "remark":           ws[f"AF{row}"].value,
        })

    # Always cache the full task list for future reads
    _write_tasks_cache(project_id, tasks)

    # Apply filter in memory if needed
    if owner_filter:
        return [t for t in tasks if t.get("owner") == owner_filter]
    return tasks


def update_task(project_id, task_code, percent_complete, remark):
    """Update % complete and remark for a specific task in the Excel file."""
    fname = project_id + ".xlsx"
    fpath = os.path.join(PROJECTS_DIR, fname)
    if not os.path.exists(fpath):
        return False, "Project file not found"

    wb = openpyxl.load_workbook(fpath)
    ws = wb.active

    found = False
    for row in range(TASK_START_ROW, TASK_END_ROW + 1):
        code = ws[f"H{row}"].value
        name = ws[f"I{row}"].value
        # Match by task_code + task_name combo (since codes repeat)
        row_key = f"{code}_{str(name).strip()}" if name else code
        if row_key == task_code:
            # Write % as decimal (0.0–1.0) to match Excel format
            ws[f"Z{row}"] = percent_complete / 100.0
            ws[f"AF{row}"] = remark or ""
            found = True
            break

    if not found:
        return False, "Task not found"

    wb.save(fpath)
    return True, "Updated"

def update_tasks_bulk(project_id, updates):
    """
    Bulk update multiple tasks at once.
    Supports two update formats:
    1. Task-key based: {task_key, percent_complete, remark}
    2. Row-based (from sheet inputs): {_row, actual_start?, actual_end?, percent_complete?, help_required?}
    """
    fname = project_id + ".xlsx"
    fpath = os.path.join(PROJECTS_DIR, fname)
    if not os.path.exists(fpath):
        return False, "Project file not found"

    wb = openpyxl.load_workbook(fpath)
    ws = wb.active

    # Separate row-based updates from task-key updates
    row_updates  = {u["_row"]: u for u in updates if "_row" in u}
    task_updates = {u["task_key"]: u for u in updates if "task_key" in u}
    updated = 0

    for row in range(TASK_START_ROW, TASK_END_ROW + 1):
        code = ws[f"H{row}"].value
        name = ws[f"I{row}"].value

        # ── Row-based update (from sheet input cells) ──
        if row in row_updates:
            u = row_updates[row]
            if "actual_start" in u and u["actual_start"]:
                ws[f"X{row}"] = u["actual_start"]
            if "actual_end" in u and u["actual_end"]:
                ws[f"Y{row}"] = u["actual_end"]
            if "percent_complete" in u and u["percent_complete"] is not None:
                try:
                    ws[f"Z{row}"] = float(u["percent_complete"]) / 100.0
                except (ValueError, TypeError):
                    pass
            if "help_required" in u:
                ws[f"AD{row}"] = u["help_required"] or ""
            updated += 1

        # ── Task-key based update ──
        if not code:
            continue
        row_key = f"{code}_{str(name).strip()}" if name else str(code)
        if row_key in task_updates:
            u = task_updates[row_key]
            ws[f"Z{row}"] = u["percent_complete"] / 100.0
            if "actual_start" in u and u["actual_start"]:
                ws[f"X{row}"] = u["actual_start"]
            if "actual_end" in u and u["actual_end"]:
                ws[f"Y{row}"] = u["actual_end"]
            ws[f"AF{row}"] = u.get("remark") or ""
            updated += 1

    wb.save(fpath)

    # Update JSON cache in-place
    try:
        cache_path = _tasks_cache_path(project_id)
        if os.path.exists(cache_path):
            with open(cache_path, "r") as f:
                cached_tasks = json.load(f)
            for t in cached_tasks:
                key = t.get("task_code", "") + "_" + t.get("task_name", "").strip()
                if key in task_updates:
                    u = task_updates[key]
                    t["percent_complete"] = u["percent_complete"]
                    t["remark"]           = u.get("remark") or ""
            with open(cache_path, "w") as f:
                json.dump(cached_tasks, f)
    except Exception as e:
        print(f"Cache update failed for {project_id}: {e}")

    update_index_entry(project_id)
    return True, f"Updated {updated} tasks"


def create_user(name, short_name, username, password, role):
    """Add a new user to users.xlsx."""
    import bcrypt
    # Check duplicate username
    for u in get_all_users():
        if u["username"].lower() == username.lower():
            return False, "Username already exists"
        if u["short_name"].upper() == short_name.strip().upper():
            return False, "Short name already exists"

    pw_hash = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
    wb = openpyxl.load_workbook(USERS_PATH)
    ws = wb.active
    ws.append([name, short_name.strip().upper(), username, pw_hash, role])
    wb.save(USERS_PATH)
    return True, "User created"


def update_user(username, updates):
    """Update an existing user. updates = dict of fields to change."""
    import bcrypt
    wb = openpyxl.load_workbook(USERS_PATH)
    ws = wb.active

    for row in ws.iter_rows(min_row=2):
        if row[2].value == username:  # col C = username
            if "name"       in updates: row[0].value = updates["name"]
            if "short_name" in updates: row[1].value = updates["short_name"].strip().upper()
            if "role"       in updates: row[4].value = updates["role"]
            if "password"   in updates and updates["password"]:
                row[3].value = bcrypt.hashpw(updates["password"].encode(), bcrypt.gensalt()).decode()
            wb.save(USERS_PATH)
            return True, "User updated"

    return False, "User not found"


def get_master_projects(username):
    """
    Read the user's master file from data/master/{username}.xlsb (or .xlsx fallback).
    Sheet: SWMon
    Col F  (index 5)  = project IDs (start row 2)
    Col BA (index 52) = last-edited date; red fill = stale (>2 days), grey = recent
    Returns list of {project_id, stale (bool), file_exists (bool)}
    """
    base = os.path.join(PROJECTS_DIR, "master", username)
    # Prefer .xlsb, fall back to .xlsx
    if os.path.exists(base + ".xlsb"):
        return _read_master_xlsb(base + ".xlsb")
    elif os.path.exists(base + ".xlsx"):
        return _read_master_xlsx(base + ".xlsx")
    return []


def _master_row_to_entry(pid, ba_val, ba_rgb=None):
    """Convert raw col-F / col-BA values into a result dict."""
    if not pid:
        return None
    pid = str(pid).strip()
    if not pid:
        return None

    # Only consider project IDs starting with FSL
    if not pid.startswith("FSL"):
        return None

    today = date.today()
    stale = False

    if ba_val is not None:
        if isinstance(ba_val, (datetime, date)):
            last_edited = ba_val.date() if isinstance(ba_val, datetime) else ba_val
            stale = (today - last_edited).days > 2
        # numeric serial date (Excel stores dates as floats in xlsb)
        elif isinstance(ba_val, (int, float)):
            try:
                from datetime import timedelta
                epoch = date(1899, 12, 30)
                last_edited = epoch + timedelta(days=int(ba_val))
                stale = (today - last_edited).days > 2
            except Exception:
                stale = False
    elif ba_rgb:
        # Fall back to cell colour: reddish = stale
        try:
            rgb = ba_rgb[-6:]
            r = int(rgb[0:2], 16)
            g = int(rgb[2:4], 16)
            b = int(rgb[4:6], 16)
            stale = r > 150 and g < 100 and b < 100
        except Exception:
            stale = False

    # Normalize: FSL/2122/CHN/OR004_PLC → SWESch_FSL_2122_CHN_OR004_PLC
    normalized_id = "SWESch_" + pid.replace("/", "_")
    file_exists = os.path.exists(os.path.join(PROJECTS_DIR, normalized_id + ".xlsx"))

    return {
        "project_id":    pid,            # display value e.g. FSL/2122/CHN/OR004_PLC
        "file_id":       normalized_id,  # actual file id e.g. SWESch_FSL_2122_CHN_OR004_PLC
        "stale":         stale,
        "file_exists":   file_exists,
    }


def _read_master_xlsb(path):
    """Read master file in .xlsb format using pyxlsb."""
    try:
        from pyxlsb import open_workbook
    except ImportError:
        print("pyxlsb not installed — cannot read .xlsb master file")
        return []

    results = []
    try:
        with open_workbook(path) as wb:
            sheet_names = wb.sheets
            if "SWMon" not in sheet_names:
                print(f"SWMon sheet not found in {path}. Sheets: {sheet_names}")
                return []
            with wb.get_sheet("SWMon") as ws:
                for i, row in enumerate(ws.rows()):
                    if i == 0:
                        continue  # skip header
                    # col F = index 5, col BA = index 52
                    f_val  = row[5].v  if len(row) > 5  else None
                    ba_val = row[52].v if len(row) > 52 else None
                    entry = _master_row_to_entry(f_val, ba_val)
                    if entry:
                        results.append(entry)
    except Exception as e:
        print(f"Master .xlsb read error for {path}: {e}")
    return results


def _read_master_xlsx(path):
    """Read master file in .xlsx format using openpyxl."""
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        if "SWMon" not in wb.sheetnames:
            return []
        ws = wb["SWMon"]
    except Exception as e:
        print(f"Master .xlsx read error for {path}: {e}")
        return []

    results = []
    for row in ws.iter_rows(min_row=2, values_only=False):
        f_cell  = row[5]  if len(row) > 5  else None
        ba_cell = row[52] if len(row) > 52 else None
        pid    = f_cell.value if f_cell else None
        ba_val = ba_cell.value if ba_cell else None
        # Try to get fill colour for fallback staleness check
        ba_rgb = None
        if ba_cell:
            try:
                fg = ba_cell.fill.fgColor
                if fg and fg.type == "rgb" and fg.rgb not in ("00000000", "FFFFFFFF"):
                    ba_rgb = fg.rgb
            except Exception:
                pass
        entry = _master_row_to_entry(pid, ba_val, ba_rgb)
        if entry:
            results.append(entry)
    return results


_THEME_COLORS = {
    0: "#FFFFFF", 1: "#000000", 2: "#EEECE1", 3: "#DDD9C3",
    4: "#C4BD97", 5: "#938953", 6: "#494429", 7: "#1F1A0E",
    8: "#C6D9F0", 9: "#8DB3E2",
}

def _resolve_color(color_obj):
    if color_obj is None:
        return None
    try:
        if color_obj.type == "rgb":
            raw = color_obj.rgb
            if raw in ("00000000", "FFFFFFFF"):
                return None
            return "#" + raw[2:]
        if color_obj.type == "theme":
            return _THEME_COLORS.get(color_obj.theme)
    except Exception:
        pass
    return None

def get_raw_sheet(filepath, max_col=30):
    """Read cell values + basic formatting only. Fast version — no borders, no alignment."""
    from openpyxl.utils import get_column_letter as gcl
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    max_row = ws.max_row
    max_col = min(ws.max_column, max_col)
    cols = [gcl(i) for i in range(1, max_col + 1)]

    # ── Merged cells ──────────────────────────────────────────
    merged_map = {}
    for mc in ws.merged_cells.ranges:
        master = f"{gcl(mc.min_col)}{mc.min_row}"
        rs = mc.max_row - mc.min_row + 1
        cs = mc.max_col - mc.min_col + 1
        for r in range(mc.min_row, mc.max_row + 1):
            for c in range(mc.min_col, mc.max_col + 1):
                coord = f"{gcl(c)}{r}"
                if coord == master:
                    merged_map[coord] = {"master": True, "rowspan": rs, "colspan": cs}
                else:
                    merged_map[coord] = {"skip": True}

    # ── Column widths + hidden flags ─────────────────────────
    # A column is "grouped/hidden" if:
    #   (a) it has outline_level > 0 and hidden=True, OR
    #   (b) it has no dimension entry (uses sheet default width ~0.88) meaning
    #       Excel collapsed it by making it near-zero width
    DEFAULT_COL_WIDTH = getattr(ws.sheet_format, 'defaultColWidth', None) or 0.88
    col_widths  = {}
    col_hidden  = {}   # True if this col is part of a collapsed group
    col_outline = {}

    for col in cols:
        try:
            cd = ws.column_dimensions.get(col)
            if cd is None:
                # No explicit dimension — uses sheet default (near-zero = collapsed group member)
                col_widths[col]  = max(30, round((DEFAULT_COL_WIDTH or 8) * 7.5))
                col_hidden[col]  = DEFAULT_COL_WIDTH < 2.0
                col_outline[col] = 0
            else:
                col_widths[col]  = max(30, round((cd.width or 8) * 7.5))
                # Only count as group-hidden if outline_level > 0 (not plain hidden cols)
                col_hidden[col]  = bool(cd.hidden) and int(cd.outline_level or 0) > 0
                col_outline[col] = int(cd.outline_level or 0)
        except Exception:
            col_widths[col]  = 64
            col_hidden[col]  = False
            col_outline[col] = 0

    # ── Manual column width adjustments ──────────────────────
    _col_overrides = {"B": 0.90, "C": 0.80, "D": 0.50, "I": 0.60, "V": 0.90, "W": 0.90}
    for _col, _factor in _col_overrides.items():
        if _col in col_widths:
            col_widths[_col] = max(30, round(col_widths[_col] * _factor))

    # ── Build col_groups ──────────────────────────────────────
    # A group is a run of consecutive hidden/near-zero cols bounded by a
    # visible col that has outline_level > 0 (the group "anchor").
    # Strategy: find cols that are hidden or have outline_level>0, group
    # consecutive ones together.
    col_groups = []
    try:
        i = 0
        while i < len(cols):
            col = cols[i]
            # A col belongs to a group if it has outline_level>0, OR it has no
            # dimension entry (tiny default width) meaning it was collapsed inline
            cd_check = ws.column_dimensions.get(col)
            is_grouped = (
                col_outline.get(col, 0) > 0 or
                (cd_check is None and DEFAULT_COL_WIDTH < 2.0)
            )
            if is_grouped:
                group_cols = [col]
                j = i + 1
                while j < len(cols):
                    next_col = cols[j]
                    cd_next = ws.column_dimensions.get(next_col)
                    next_grouped = (
                        col_outline.get(next_col, 0) > 0 or
                        (cd_next is None and DEFAULT_COL_WIDTH < 2.0)
                    )
                    if next_grouped:
                        group_cols.append(next_col)
                        j += 1
                    else:
                        break
                col_groups.append({
                    "cols":      group_cols,
                    "level":     1,
                    "collapsed": all(col_hidden.get(c, False) for c in group_cols),
                })
                i = j
            else:
                i += 1
    except Exception:
        col_groups = []

    # ── Row heights ───────────────────────────────────────────
    row_heights = {}
    for rn in range(1, max_row + 1):
        rd = ws.row_dimensions.get(rn)
        row_heights[str(rn)] = max(18, round((rd.height or 15) * 1.33)) if rd else 20

    # ── Cells ─────────────────────────────────────────────────
    cells = {}
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            coord = cell.coordinate
            mi = merged_map.get(coord, {})

            if mi.get("skip"):
                cells[coord] = {"skip": True}
                continue

            # Value
            v = cell.value
            if isinstance(v, (datetime, date)):
                if isinstance(v, datetime): v = v.date()
                v = v.strftime("%d-%b-%y")
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                fmt = cell.number_format or ""
                if "%" in fmt:
                    v = f"{int(round(float(v) * 100))}%"
            elif v is not None:
                v = str(v) if not isinstance(v, (int, str, bool)) else v

            c = {"v": v}

            # Merge spans
            if mi.get("master"):
                if mi["rowspan"] > 1: c["rowspan"] = mi["rowspan"]
                if mi["colspan"] > 1: c["colspan"] = mi["colspan"]

            # Font — bold and size only (skip color/italic/underline)
            try:
                f = cell.font
                font = {}
                if f.bold: font["bold"] = True
                if f.size and f.size != 11: font["size"] = f.size
                if font: c["font"] = font
            except Exception:
                pass

            # Fill — background color only
            try:
                fill = cell.fill
                if fill and fill.patternType not in (None, "none"):
                    fg = _resolve_color(fill.fgColor)
                    if fg and fg != "#FFFFFF": c["fill"] = fg
            except Exception:
                pass

            cells[coord] = c

    # ── Detect and mark editable cells ───────────────────────
    # Read the fill fingerprint from cell X at TASK_START_ROW — that is the
    # reference editable cell. Any cell in EDITABLE_COLS with the same
    # theme+tint in the data rows is marked editable.
    EDITABLE_COLS  = {"X", "Y", "Z", "AD"}
    TINT_TOLERANCE = 0.001
    editable_fill  = None
    editable_theme = None
    editable_tint  = None

    try:
        ref = ws[f"X{TASK_START_ROW}"]
        fg  = ref.fill.fgColor
        if ref.fill.patternType == "solid" and fg.type == "theme":
            editable_theme = fg.theme
            editable_tint  = fg.tint
            editable_fill  = "#BDD7EE"  # visual approximation sent to frontend
    except Exception:
        pass

    def _is_editable_fill(cell):
        if editable_theme is None:
            return False
        try:
            fg = cell.fill.fgColor
            return (
                cell.fill.patternType == "solid" and
                fg.type == "theme" and
                fg.theme == editable_theme and
                abs(fg.tint - editable_tint) < TINT_TOLERANCE
            )
        except Exception:
            return False

    for row in ws.iter_rows(min_row=TASK_START_ROW, max_row=max_row):
        for cell in row:
            try:
                col_letter = cell.column_letter
            except AttributeError:
                continue  # skip MergedCell objects
            if col_letter in EDITABLE_COLS and _is_editable_fill(cell):
                coord = cell.coordinate
                if coord in cells and not cells[coord].get("skip"):
                    cells[coord]["editable"] = True
                    cells[coord]["editable_col"] = col_letter

    return {
        "cells":         cells,
        "col_widths":    col_widths,
        "row_heights":   row_heights,
        "max_row":       max_row,
        "max_col":       max_col,
        "cols":          cols,
        "col_groups":    col_groups,
        "editable_fill": editable_fill,
    }

def delete_user(username):
    """Remove a user from users.xlsx."""
    wb = openpyxl.load_workbook(USERS_PATH)
    ws = wb.active
    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        if row[2].value == username:
            ws.delete_rows(i)
            wb.save(USERS_PATH)
            return True, "User deleted"
    return False, "User not found"
