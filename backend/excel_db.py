import os
import openpyxl
from datetime import datetime, date
import threading
import queue


# ── Paths ──────────────────────────────────────────────────────
BASE_DIR        = os.path.dirname(__file__)
USERS_PATH      = os.path.join(BASE_DIR, "users.xlsx")
DATA_DIR        = os.path.join(BASE_DIR, "..", "data")

# ── Role to discipline folder mapping ─────────────────────────
ROLE_DISCIPLINE = {
    "sw_tl":  "SW",
    "hw_tl":  "HW",
    "mfg_tl": "MFG",
    "pm":     "PM",
    "admin":  "SW",   # admin defaults to SW
    "head":   "SW",   # head defaults to SW
}

def get_discipline_dirs(role):
    """Return (projects_dir, monitoring_dir, schedules_cache, monitoring_cache)
    for the given role."""
    disc = ROLE_DISCIPLINE.get(role, "SW")
    disc_dir       = os.path.join(DATA_DIR, disc)
    projects_dir   = os.path.join(disc_dir, disc + "ESch")       # e.g. SW/SWESch
    monitoring_dir = disc_dir                                      # monitoring files at root of discipline
    cache_dir      = os.path.join(disc_dir, "cache")
    sched_cache    = os.path.join(cache_dir, "schedules")
    mon_cache      = os.path.join(cache_dir, "monitoring")
    os.makedirs(sched_cache, exist_ok=True)
    os.makedirs(mon_cache,   exist_ok=True)
    return projects_dir, monitoring_dir, sched_cache, mon_cache

# ── Default dirs (SW) for functions that don't have user context ──
PROJECTS_DIR, MONITORING_DIR, SWESCH_CACHE, MON_CACHE = get_discipline_dirs("sw_tl")

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

INDEX_PATH = os.path.join(DATA_DIR, "index.json")

# ── Task JSON sidecar cache ────────────────────────────────────
def _sheet_cache_path(project_id, sched_cache=None):
    return os.path.join(sched_cache or SWESCH_CACHE, project_id + "_sheet.json")

def _write_sheet_cache(project_id, data, sched_cache=None):
    """Write sheet data to JSON sidecar file."""
    try:
        with open(_sheet_cache_path(project_id, sched_cache), "w") as f:
            json.dump(data, f)
    except Exception as e:
        print(f"Sheet cache write failed for {project_id}: {e}")

def _read_sheet_cache(project_id, sched_cache=None):
    """Read sheet data from JSON sidecar. Auto-invalidates if date format is stale."""
    cache_path = _sheet_cache_path(project_id, sched_cache)
    if not os.path.exists(cache_path):
        return None
    try:
        with open(cache_path, "r") as f:
            data = json.load(f)
        # Invalidate if dates still use old single-value format (missing "pm" key)
        dates = data.get("left_panel", {}).get("dates", [])
        if dates and "pm" not in dates[0]:
            os.remove(cache_path)
            return None
        return data
    except Exception:
        return None

def _update_sheet_cache(project_id, updates, sched_cache=None):
    """
    Apply updates directly to the sheet JSON cache without reading Excel.
    Updates should be in the same format as received from frontend.
    """
    cache_path = _sheet_cache_path(project_id, sched_cache)
    
    if not os.path.exists(cache_path):
        # No cache to update, nothing to do
        return False
    
    try:
        with open(cache_path, 'r') as f:
            sheet_data = json.load(f)
        
        cells = sheet_data.get('cells', {})
        
        # Process each update
        for update in updates:
            if '_row' in update:
                # Row-based update (from sheet inputs)
                row = update['_row']
                
                # Map field to column
                field_to_col = {
                    'actual_start': 'X',
                    'actual_end': 'Y',
                    'percent_complete': 'Z',
                    'help_required': 'AD',
                    'remark': 'AF'
                }
                
                for field, value in update.items():
                    if field == '_row':
                        continue
                    
                    col = field_to_col.get(field)
                    if not col:
                        continue
                    
                    coord = f"{col}{row}"
                    
                    if coord in cells:
                        # Update the cell value
                        if field == 'percent_complete':
                            # Format as percentage string
                            cells[coord]['v'] = f"{int(value)}%"
                        else:
                            cells[coord]['v'] = value
                        # Mark that this cell was edited (optional, for debugging)
                        cells[coord]['edited'] = True
                        
            elif 'task_key' in update:
                # Task-key based update (from old task system)
                # Find the row by task_key
                task_key = update['task_key']
                percent = update.get('percent_complete')
                remark = update.get('remark')
                
                for coord, cell_data in cells.items():
                    # Check if this cell is in the task rows (H or I columns)
                    # H = task_code, I = task_name
                    if coord[0] in ('H', 'I'):
                        row_num = int(coord[1:])
                        if row_num >= 9 and row_num <= 55:
                            # Build task key from H and I cells
                            h_coord = f"H{row_num}"
                            i_coord = f"I{row_num}"
                            h_val = cells.get(h_coord, {}).get('v', '')
                            i_val = cells.get(i_coord, {}).get('v', '')
                            current_key = f"{h_val}_{i_val}".strip()
                            
                            if current_key == task_key:
                                if percent is not None:
                                    z_coord = f"Z{row_num}"
                                    if z_coord in cells:
                                        cells[z_coord]['v'] = f"{int(percent)}%"
                                if remark is not None:
                                    af_coord = f"AF{row_num}"
                                    if af_coord in cells:
                                        cells[af_coord]['v'] = remark
                                break
        
        # Write updated cache back
        with open(cache_path, 'w') as f:
            json.dump(sheet_data, f, indent=2)
        
        return True
        
    except Exception as e:
        print(f"Failed to update sheet cache for {project_id}: {e}")
        return False

def _invalidate_sheet_cache(project_id, sched_cache=None):
    """Delete sheet JSON sidecar so next open re-reads from Excel."""
    try:
        cache_path = _sheet_cache_path(project_id, sched_cache)
        if os.path.exists(cache_path):
            os.remove(cache_path)
    except Exception as e:
        print(f"Sheet cache invalidation failed for {project_id}: {e}")

def _monitoring_cache_path(short_name, mon_cache=None):
    """Path to monitoring JSON cache for a user."""
    return os.path.join(mon_cache or MON_CACHE, f"Monitor_{short_name}.json")

def _write_monitoring_cache(short_name, data, mon_cache=None):
    try:
        with open(_monitoring_cache_path(short_name, mon_cache), "w") as f:
            json.dump(data, f)
    except Exception as e:
        print(f"Monitoring cache write failed for {short_name}: {e}")

def _read_monitoring_cache(short_name, mon_cache=None):
    path = _monitoring_cache_path(short_name, mon_cache)
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r") as f:
            return json.load(f)
    except Exception:
        return None

def build_index():
    """Scan all Excel files and write index.json."""
    projects = []
    for fname in list_project_files():
        fpath = os.path.join(PROJECTS_DIR, fname)
        project_id = os.path.splitext(fname)[0]
        try:
            p = read_project_header(fpath)
            projects.append(p)
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
    return [f for f in os.listdir(PROJECTS_DIR)
            if f.endswith(".xlsx") or f.endswith(".xlsb")]

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

def get_all_projects(role="sw_tl"):
    """Read project list from all Excel files that exist in discipline SWESch folder."""
    projects_dir, _, _, _ = get_discipline_dirs(role)
    projects = []
    for fname in list_project_files(projects_dir):
        fpath = os.path.join(projects_dir, fname)
        project_id = os.path.splitext(fname)[0]
        try:
            p = read_project_header(fpath)
            projects.append(p)
        except Exception as e:
            print(f"Skipped {fname}: {e}")
    return projects

def get_projects_for_user(user):
    """Get projects for user.
    For non-admin: returns monitoring file entries directly (fast, no Excel open).
    For admin/head: scans discipline SWESch folder."""
    role = user["role"]

    if role in ("admin", "head"):
        return get_all_projects()

    master_list = get_master_projects(user["username"])
    result = []
    for m in master_list:
        result.append({
            "id":            m.get("file_id") or m.get("project_id"),
            "or_number":     m.get("project_id"),
            "customer_name": "",
            "file_exists":   m.get("file_exists", False),
            "stale":         m.get("stale", False),
        })
    return result

def get_project_by_id(project_id, role="sw_tl"):
    """Read project header. Uses sheet JSON cache if available, else reads Excel."""
    projects_dir, _, sched_cache, _ = get_discipline_dirs(role)
    # Try to get header from sheet JSON cache first (fast)
    sheet = _read_sheet_cache(project_id, sched_cache)
    if sheet and sheet.get("project_banner"):
        b = sheet["project_banner"]
        return {
            "id":            project_id,
            "or_number":     b.get("or_number", ""),
            "customer_name": b.get("customer_name", ""),
            "section":       b.get("section", ""),
            "pm_name":       sheet.get("project_banner", {}).get("sales_manager", ""),
            "sw_tl_name":    b.get("sales_engineer", ""),
        }
    # Fall back to reading Excel directly
    fpath = os.path.join(projects_dir, project_id + ".xlsx")
    if not os.path.exists(fpath):
        fpath = os.path.join(projects_dir, project_id + ".xlsb")
        if not os.path.exists(fpath):
            return None
    try:
        return read_project_header(fpath)
    except Exception:
        return None

# ── Tasks ──────────────────────────────────────────────────────
def _cell_val(cells, col, row):
    """Get raw value from sheet JSON cells dict."""
    info = cells.get(f"{col}{row}")
    if not info or info.get("skip"):
        return None
    return info.get("v")

def _pct_from_sheet(cells, col, row):
    """Extract percent as int 0-100 from sheet JSON cell."""
    v = _cell_val(cells, col, row)
    if v is None:
        return 0
    s = str(v).replace("%", "").strip()
    try:
        f = float(s)
        return int(f * 100) if f <= 1.0 else int(f)
    except (ValueError, TypeError):
        return 0

def get_tasks(project_id, owner_filter=None, role="sw_tl"):
    """Read tasks from sheet JSON sidecar — fast disk read, no Excel open."""
    projects_dir, _, sched_cache, _ = get_discipline_dirs(role)
    # Try sheet JSON first
    sheet = _read_sheet_cache(project_id, sched_cache)

    # Fall back to Excel if no sheet JSON yet
    if sheet is None:
        fpath = os.path.join(projects_dir, project_id + ".xlsx")
        if not os.path.exists(fpath):
            return []
        sheet = get_raw_sheet(fpath)  # this will also write the JSON cache

    cells = sheet.get("cells", {})
    tasks = []

    for row in range(TASK_START_ROW, TASK_END_ROW + 1):
        task_code = _cell_val(cells, "H", row)
        task_name = _cell_val(cells, "I", row)

        if not task_code or not task_name:
            continue

        phase = _cell_val(cells, "F", row)
        owner = _cell_val(cells, "G", row)

        if owner_filter and owner != owner_filter:
            continue

        pct     = _pct_from_sheet(cells, "Z", row)
        exp_pct = _pct_from_sheet(cells, "AA", row)

        tasks.append({
            "row":              row,
            "phase":            phase,
            "phase_name":       PHASE_NAMES.get(phase, phase),
            "owner":            owner,
            "task_code":        str(task_code),
            "task_name":        str(task_name).strip(),
            "rev1_start":       _cell_val(cells, "J", row),
            "rev1_end":         _cell_val(cells, "K", row),
            "rev2_start":       _cell_val(cells, "L", row),
            "rev2_end":         _cell_val(cells, "M", row),
            "rev3_start":       _cell_val(cells, "N", row),
            "rev3_end":         _cell_val(cells, "O", row),
            "orig_start":       _cell_val(cells, "P", row),
            "orig_end":         _cell_val(cells, "Q", row),
            "current_start":    _cell_val(cells, "V", row),
            "current_end":      _cell_val(cells, "W", row),
            "actual_start":     _cell_val(cells, "X", row),
            "actual_end":       _cell_val(cells, "Y", row),
            "percent_complete": pct,
            "expected_percent": exp_pct,
            "lead_time_days":   _cell_val(cells, "R", row),
            "effort_days":      _cell_val(cells, "U", row),
            "interlock":        _cell_val(cells, "T", row),
            "remark":           _cell_val(cells, "AF", row),
        })

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

def update_tasks_bulk(project_id, updates, role="sw_tl"):
    """
    Bulk update multiple tasks.
    - Updates JSON cache immediately (fast)
    - Queues Excel write for background (slow)
    """
    projects_dir, _, sched_cache, _ = get_discipline_dirs(role)
    
    # Step 1: Update JSON cache immediately (fast)
    _update_sheet_cache(project_id, updates, sched_cache)
    
    # Step 2: Queue Excel write to background thread
    queue_excel_write(project_id, updates, role)
    
    # Step 3: Count how many changes were made
    updated_count = len(updates)
    
    return True, f"Updated {updated_count} tasks (Excel syncing in background)"

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
    Find the user's monitoring file in their discipline folder root.
    File naming: SWMonitor_{INITIALS}_V{x.x}_{date}.xlsb
    Uses JSON cache in discipline/cache/monitoring/
    """
    user = get_user_by_username(username)
    if not user:
        return []

    short = user.get("short_name", "").strip().upper()
    role  = user.get("role", "sw_tl")
    _, monitoring_dir, _, mon_cache = get_discipline_dirs(role)

    if not os.path.exists(monitoring_dir):
        return []

    # Check monitoring cache first
    cached = _read_monitoring_cache(short, mon_cache)
    if cached is not None:
        return cached

    # Find monitoring file by initials
    for fname in os.listdir(monitoring_dir):
        if not (fname.endswith(".xlsb") or fname.endswith(".xlsx")):
            continue
        parts = fname.split("_")
        if len(parts) >= 2 and parts[1].upper() == short:
            fpath = os.path.join(monitoring_dir, fname)
            result = _read_master_xlsb(fpath) if fname.endswith(".xlsb") else _read_master_xlsx(fpath)
            _write_monitoring_cache(short, result, mon_cache)
            return result

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
    file_exists = (os.path.exists(os.path.join(PROJECTS_DIR, normalized_id + ".xlsx")) or
                   os.path.exists(os.path.join(PROJECTS_DIR, normalized_id + ".xlsb")))

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

def get_raw_sheet(filepath, max_col=32, sched_cache=None):
    """Read cell values + basic formatting. Uses JSON sidecar cache for speed."""
    from openpyxl.utils import get_column_letter as gcl

    # Check JSON sidecar cache first
    project_id = os.path.splitext(os.path.basename(filepath))[0]
    cached = _read_sheet_cache(project_id, sched_cache)
    if cached is not None:
        return cached

    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    max_row = ws.max_row
    max_col = min(ws.max_column, max_col)
    cols = [gcl(i) for i in range(1, max_col + 1) if gcl(i) not in ("A","B","C","D")]  # hide cols A-D

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
    _col_overrides = {"B": 0.90, "C": 0.80, "D": 0.35, "I": 0.60, "V": 0.90, "W": 0.90}
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
                # If group starts at J, trim to stop at P (exclude Q)
                if group_cols[0] == "J" and "Q" in group_cols:
                    group_cols = group_cols[:group_cols.index("Q")]

                # If group starts at X, extend to include AB explicitly
                # (AB has outline=0 in Excel but logically belongs to this group)
                if group_cols[0] == "X" and "AB" in cols and "AB" not in group_cols:
                    # Add any missing cols between last group col and AB
                    from openpyxl.utils import column_index_from_string as col2idx
                    last_idx = col2idx(group_cols[-1])
                    ab_idx   = col2idx("AB")
                    for extra_i in range(last_idx + 1, ab_idx + 1):
                        extra_col = gcl(extra_i)
                        if extra_col in cols and extra_col not in group_cols:
                            group_cols.append(extra_col)
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
    EDITABLE_COLS  = {"X", "Y", "Z", "AD", "AF"}
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

    # Rows to exclude from grid display (AE6/AF6 warranty shown in left panel)

    for row in ws.iter_rows(min_row=TASK_START_ROW, max_row=max_row):
        for cell in row:
            try:
                col_letter = cell.column_letter
            except AttributeError:
                continue  # skip MergedCell objects
            if col_letter in EDITABLE_COLS:
                coord = cell.coordinate
                if coord in cells and not cells[coord].get("skip"):
                    # AF (remarks) is always editable in task rows regardless of fill
                    # Other cols require the blue fill color
                    if col_letter == "AF" or _is_editable_fill(cell):
                        # For AD column: add department color + only editable if Z < 100
                        if col_letter == "AD":
                            AD_COLORS = {
                                "engineering":       "#ffb3b3",
                                "purchase":          "#5f933c",
                                "software":          "#0096cc",
                                "project management":"#005fa3",
                                "manufacturing":     "#2f491e",
                                "sales":             "#00d9d9",
                                "client":            "#6d006d",
                            }
                            val = cell.value
                            if val:
                                color = AD_COLORS.get(str(val).strip().lower())
                                if color:
                                    cells[coord]["fill"] = color
                                    # Set text color white for dark backgrounds
                                    dark_bgs = {"#5f933c","#0096cc","#005fa3","#2f491e","#6d006d"}
                                    cells[coord]["font"] = cells[coord].get("font", {})
                                    cells[coord]["font"]["color"] = "#ffffff" if color in dark_bgs else "#000000"
                            # Only editable if Z column (% complete) < 100
                            z_coord = "Z" + str(cell.row)
                            z_cell = ws[z_coord]
                            z_val = z_cell.value
                            z_pct = 0
                            try:
                                if z_val is not None:
                                    z_pct = float(z_val)
                                    if z_pct <= 1.0:
                                        z_pct = z_pct * 100
                            except (TypeError, ValueError):
                                z_pct = 0
                            if z_pct < 100:
                                cells[coord]["editable"] = True
                                cells[coord]["editable_col"] = col_letter
                        else:
                            cells[coord]["editable"] = True
                            cells[coord]["editable_col"] = col_letter

    # Read project info from header rows for banner
    def _v(ref):
        v = ws[ref].value
        return str(v).strip() if v else ""

    def _date(ref):
        v = ws[ref].value
        if v is None: return ""
        if isinstance(v, (datetime, date)):
            d = v.date() if isinstance(v, datetime) else v
            return d.strftime("%d %b %Y")
        return str(v).strip()

    # W2 may be empty — fall back to X9 (first task actual start date)
    w2_val = _date("W2") or _date("X9")

    project_banner = {
        "or_number":      _v("D1"),
        "section":        _v("D7"),
        "sales_engineer": _v("I2"),
        "sales_manager":  _v("I3"),
        "customer_name":  _v("D10"),
        "start_date_lbl": _v("W1"),
        "start_date_val": w2_val,
        "days_swe_lbl":   _v("AB1"),
        "days_swe_val":   _v("AB2"),
        "ld_date_lbl":    _v("AE1"),
        "ld_date_val":    _v("AF1"),
        "ld_maxwk_lbl":   _v("AE2"),
        "ld_maxwk_val":   _v("AF2"),
        "ld_maxov_lbl":   _v("AE3"),
        "ld_maxov_val":   _v("AF3"),
        "ld_remarks_lbl": _v("AE4"),
        "ld_remarks_val": _v("AF4"),
    }

    # Vertical left panel data
    left_panel = {
        "project_info": [
            {"label": _v("C9")  or "PO Value",      "value": str(_v("D9") or "")},
            {"label": _v("C10") or "Customer",      "value": _v("D10")},
            {"label": _v("C11") or "End Customer",  "value": _v("D11")},
            {"label": _v("C12") or "Consultant",    "value": _v("D12")},
            {"label": _v("C13") or "Project Desc.", "value": _v("D13")},
            {"label": _v("C14") or "Section",       "value": _v("D14")},
            {"label": _v("C17") or "SW Efforts",    "value": str(_v("D17") or "")},
        ],
        "dates": [
            {"label": _v("B28") or "B28", "pm": _date("C28"), "swe": _date("D28"), "pm_fill": _resolve_color(ws["C28"].fill.fgColor) if ws["C28"].fill.patternType not in (None,"none") else None, "pm_font": _resolve_color(ws["C28"].font.color), "swe_fill": _resolve_color(ws["D28"].fill.fgColor) if ws["D28"].fill.patternType not in (None,"none") else None, "swe_font": _resolve_color(ws["D28"].font.color)},
            {"label": _v("B29") or "B29", "pm": _date("C29"), "swe": _date("D29"), "pm_fill": _resolve_color(ws["C29"].fill.fgColor) if ws["C29"].fill.patternType not in (None,"none") else None, "pm_font": _resolve_color(ws["C29"].font.color), "swe_fill": _resolve_color(ws["D29"].fill.fgColor) if ws["D29"].fill.patternType not in (None,"none") else None, "swe_font": _resolve_color(ws["D29"].font.color)},
            {"label": _v("B30") or "B30", "pm": _date("C30"), "swe": _date("D30"), "pm_fill": _resolve_color(ws["C30"].fill.fgColor) if ws["C30"].fill.patternType not in (None,"none") else None, "pm_font": _resolve_color(ws["C30"].font.color), "swe_fill": _resolve_color(ws["D30"].fill.fgColor) if ws["D30"].fill.patternType not in (None,"none") else None, "swe_font": _resolve_color(ws["D30"].font.color)},
            {"label": _v("B31") or "B31", "pm": _date("C31"), "swe": _date("D31"), "pm_fill": _resolve_color(ws["C31"].fill.fgColor) if ws["C31"].fill.patternType not in (None,"none") else None, "pm_font": _resolve_color(ws["C31"].font.color), "swe_fill": _resolve_color(ws["D31"].fill.fgColor) if ws["D31"].fill.patternType not in (None,"none") else None, "swe_font": _resolve_color(ws["D31"].font.color)},
            {"label": _v("B32") or "B32", "pm": _date("C32"), "swe": _date("D32"), "pm_fill": _resolve_color(ws["C32"].fill.fgColor) if ws["C32"].fill.patternType not in (None,"none") else None, "pm_font": _resolve_color(ws["C32"].font.color), "swe_fill": _resolve_color(ws["D32"].fill.fgColor) if ws["D32"].fill.patternType not in (None,"none") else None, "swe_font": _resolve_color(ws["D32"].font.color)},
            {"label": _v("B33") or "B33", "pm": _date("C33"), "swe": _date("D33"), "pm_fill": _resolve_color(ws["C33"].fill.fgColor) if ws["C33"].fill.patternType not in (None,"none") else None, "pm_font": _resolve_color(ws["C33"].font.color), "swe_fill": _resolve_color(ws["D33"].fill.fgColor) if ws["D33"].fill.patternType not in (None,"none") else None, "swe_font": _resolve_color(ws["D33"].font.color)},
        ],
        "warranty": [
            {"label": _v("AE6") or "Warranty", "value": _v("AF6")},
        ],
        "stakeholders": [
            {"label": _v("B36") or "Sales",  "value": _v("C36")},
            {"label": _v("B37") or "PM",     "value": _v("C37")},
            {"label": _v("B38") or "HW",     "value": _v("C38")},
            {"label": _v("B41") or "SW",     "value": _v("C41")},
            {"label": _v("B40") or "MFG",    "value": _v("C40")},
            {"label": _v("B42") or "E&C",    "value": _v("C42")},
            {"label": _v("B39") or "BYR",    "value": _v("C39")},
            {"label": _v("B43") or "A/C",    "value": _v("C43")},
        ],
    }

    result = {
        "cells":          cells,
        "col_widths":     col_widths,
        "row_heights":    row_heights,
        "max_row":        max_row,
        "max_col":        max_col,
        "cols":           cols,
        "col_groups":     col_groups,
        "editable_fill":  editable_fill,
        "project_banner": project_banner,
        "info_rows":      list(range(1, 6)),
        "header_rows":    [],
        "left_panel":     left_panel,
    }

    # Write JSON sidecar cache for fast future loads
    _write_sheet_cache(project_id, result, sched_cache)
    return result

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


# Add this code at the VERY END of the file

# Global queue for Excel write operations
_excel_write_queue = queue.Queue()
_background_thread_started = False

def _start_background_worker():
    """Start a background thread that processes Excel writes"""
    global _background_thread_started
    if _background_thread_started:
        return
    
    def worker():
        while True:
            try:
                # Wait for a task
                project_id, updates, role = _excel_write_queue.get(timeout=1)
                try:
                    _do_excel_write(project_id, updates, role)
                    print(f"[Background] Excel write completed for {project_id}")
                except Exception as e:
                    print(f"[Background] Error writing {project_id}: {e}")
                finally:
                    _excel_write_queue.task_done()
            except queue.Empty:
                continue
            except Exception as e:
                print(f"[Background] Worker error: {e}")
    
    thread = threading.Thread(target=worker, daemon=True)
    thread.start()
    _background_thread_started = True
    print("[Background] Excel write worker started")

def _do_excel_write(project_id, updates, role):
    """Actually write to Excel file (this is the slow part)"""
    from openpyxl import load_workbook
    import os
    
    projects_dir, _, _, _ = get_discipline_dirs(role)
    fpath = os.path.join(projects_dir, project_id + ".xlsx")
    
    if not os.path.exists(fpath):
        fpath = os.path.join(projects_dir, project_id + ".xlsb")
        if not os.path.exists(fpath):
            print(f"[Background] Excel file not found for {project_id}")
            return
    
    # Load workbook
    wb = load_workbook(fpath)
    ws = wb.active
    
    # Apply updates
    row_updates = {u["_row"]: u for u in updates if "_row" in u}
    task_updates = {u["task_key"]: u for u in updates if "task_key" in u}
    
    TASK_START_ROW = 9
    TASK_END_ROW = 55
    
    for row in range(TASK_START_ROW, TASK_END_ROW + 1):
        # Row-based update
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
            if "remark" in u:
                ws[f"AF{row}"] = u["remark"] or ""
        
        # Task-key based update
        else:
            code = ws[f"H{row}"].value
            name = ws[f"I{row}"].value
            if code and name:
                row_key = f"{code}_{str(name).strip()}"
                if row_key in task_updates:
                    u = task_updates[row_key]
                    ws[f"Z{row}"] = u["percent_complete"] / 100.0
                    if "actual_start" in u and u["actual_start"]:
                        ws[f"X{row}"] = u["actual_start"]
                    if "actual_end" in u and u["actual_end"]:
                        ws[f"Y{row}"] = u["actual_end"]
                    ws[f"AF{row}"] = u.get("remark") or ""
    
    # Save workbook
    wb.save(fpath)
    print(f"[Background] Excel file saved: {fpath}")

def queue_excel_write(project_id, updates, role):
    """Queue a project for background Excel write"""
    _start_background_worker()  # Ensure worker is running
    _excel_write_queue.put((project_id, updates, role))