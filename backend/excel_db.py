import os
import openpyxl
from datetime import datetime, date
import threading
import queue
from openpyxl.utils import get_column_letter as gcl
import pycel
from pycel.excelcompiler import ExcelCompiler
from sysmemory_compute import compute_sysmemory_json, queue_sysmemory_compute
from openpyxl.styles.numbers import is_date_format

try:
    from pycel.excelcompiler import ExcelCompiler
    PY_CEL_AVAILABLE = True
    print("PyCel is available for formula evaluation")
except ImportError:
    PY_CEL_AVAILABLE = False
    print("PyCel not installed, falling back to openpyxl only")


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
                task_key = update['task_key']
                percent = update.get('percent_complete')
                remark = update.get('remark')

                # Direct row lookup — O(n rows) instead of O(n cells)
                for row_num in range(9, 56):
                    h_val = cells.get(f"H{row_num}", {}).get('v', '') or ''
                    i_val = cells.get(f"I{row_num}", {}).get('v', '') or ''
                    current_key = f"{h_val}_{str(i_val).strip()}"
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
        
        # Stamp last_modified with current time — JSON mtime is the source of truth
        # for "last edited via the app", more accurate than the Excel file mtime
        # because the Excel write happens in a background thread.
        now_str = datetime.now().strftime("%d %b %Y, %I:%M %p")
        sheet_data['last_modified'] = now_str

        # Write updated cache back
        with open(cache_path, 'w') as f:
            json.dump(sheet_data, f, indent=2)

        return now_str  # return so caller can propagate to in-memory cache + response
        
    except Exception as e:
        print(f"Failed to update sheet cache for {project_id}: {e}")
        return None

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
    - For admin/head: returns ALL projects from department monitoring file
    - For TL: returns only projects assigned to them
    """
    role = user["role"]
    short_name = user.get("short_name", "").strip().upper()

    print(f"[get_projects_for_user] role: {role}, short_name: {short_name}")

    if role in ("admin", "head"):
        # Admin/Head: get all projects from department monitoring file
        print("[get_projects_for_user] Using get_all_monitor_projects")
        master_list = get_all_monitor_projects(role)
        print(f"[get_projects_for_user] Found {len(master_list)} projects from monitor file")
    else:
        # TL: get projects where they are assigned
        print("[get_projects_for_user] Using get_master_projects")
        master_list = get_master_projects(user["username"])
        print(f"[get_projects_for_user] Found {len(master_list)} projects")

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
    
    # ALWAYS try sheet JSON cache first
    sheet = _read_sheet_cache(project_id, sched_cache)
    
    if sheet and sheet.get("project_banner"):
        # Fast path - use cached banner (THIS SHOULD BE FAST!)
        b = sheet["project_banner"]
        return {
            "id":            project_id,
            "or_number":     b.get("or_number", ""),
            "customer_name": b.get("customer_name", ""),
            "section":       b.get("section", ""),
            "pm_name":       sheet.get("project_banner", {}).get("sales_manager", ""),
            "sw_tl_name":    b.get("sales_engineer", ""),
            "overall_percent": 0,  # Not in cache, but not critical for first load
        }
    
    # If cache missing, generate sheet data first (this will create the cache)
    # Then read from cache
    fpath = os.path.join(projects_dir, project_id + ".xlsx")
    if not os.path.exists(fpath):
        fpath = os.path.join(projects_dir, project_id + ".xlsb")
        if not os.path.exists(fpath):
            return None
    
    # Generate the sheet cache first (this is slow but only once)
    try:
        from excel_db import get_raw_sheet
        sheet_data = get_raw_sheet(fpath, sched_cache=sched_cache)
        # Now read from cache
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
                "overall_percent": 0,
            }
    except Exception as e:
        print(f"Failed to generate sheet cache for {project_id}: {e}")
        return None
    
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
        sheet = get_raw_sheet(fpath, sched_cache=sched_cache)  # pass correct cache dir

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
    - Updates monitor timestamp
    - Regenerates system memory JSON in background
    """
    projects_dir, _, sched_cache, _ = get_discipline_dirs(role)

    # Step 1: Update JSON cache immediately (fast) — returns stamped timestamp
    now_str = _update_sheet_cache(project_id, updates, sched_cache)

    # Step 2: Update monitor timestamp with the same timestamp (JSON cache + queue Excel write)
    if now_str:
        update_monitor_timestamp(project_id, now_str, role)

    # Step 3: Queue Excel write to background thread
    queue_excel_write(project_id, updates, role)

    # Step 4: Regenerate sysmemory in background
    sysmemory_cache = os.path.join(DATA_DIR, ROLE_DISCIPLINE.get(role, "SW"), "cache", "system_memory")
    queue_sysmemory_compute(project_id, sched_cache, sysmemory_cache)
    
    updated_count = len(updates)
    return True, f"Updated {updated_count} tasks (Excel syncing in background)", now_str

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
    Return the project list for a user from the shared SW_Monitor file.
    - Admin/Head: all projects unfiltered
    - SW TL: only projects where their initials appear in col C (SWH HEAD) or col D (SWE NAME)
    """
    user = get_user_by_username(username)
    if not user:
        return []

    short = user.get("short_name", "").strip().upper()
    role  = user.get("role", "sw_tl")

    # Admin and head see all projects unfiltered
    if role in ("admin", "head"):
        return get_all_monitor_projects(role)

    # SW TL: read all projects from the shared monitor file then filter
    all_projects = get_all_monitor_projects(role)
    filtered = [
        m for m in all_projects
        if m.get("swh_head") == short or m.get("swe_name") == short
    ]
    print(f"[get_master_projects] SW TL {short}: {len(filtered)}/{len(all_projects)} projects")
    return filtered

def get_all_monitor_projects(role="sw_tl"):
    """Get ALL projects from the department monitoring file (using JSON cache)"""
    from datetime import date, datetime
    import json
    import os
    
    # Map role to department
    role_to_dept = {
        "sw_tl": "SW",
        "hw_tl": "HW", 
        "mfg_tl": "MFG",
        "pm": "PM",
        "admin": "SW",
        "head": "SW"
    }
    
    department = role_to_dept.get(role, "SW")
    
    # FIRST: Try to read from JSON cache
    monitor_cache_path = os.path.join(DATA_DIR, department, "cache", "monitoring", f"{department}_Monitor.json")
    
    if os.path.exists(monitor_cache_path):
        try:
            with open(monitor_cache_path, 'r') as f:
                monitor_data = json.load(f)
            
            cells = monitor_data.get('cells', {})
            results = []
            today = date.today()
            
            # Find all rows with project IDs (column I)
            for coord, cell_info in cells.items():
                if coord.startswith('F') and cell_info.get('v'):
                    row_num = int(coord[1:])
                    project_id = str(cell_info.get('v')).strip()
                    
                    # Only process FSL projects
                    if not project_id.startswith("FSL"):
                        continue
                    
                    # Get timestamp from column BA
                    ba_coord = f"BA{row_num}"
                    timestamp_str = cells.get(ba_coord, {}).get('v')
                    
                    # Calculate stale: more than 2 days old
                    stale = False
                    if timestamp_str:
                        try:
                            # Parse timestamp like "29-03-2026 20:04:00"
                            last_edited = datetime.strptime(timestamp_str, "%d-%m-%Y %H:%M:%S").date()
                            stale = (today - last_edited).days > 2
                            print(f"[Monitor] Project {project_id}: last_edited={last_edited}, stale={stale}")
                        except Exception as e:
                            print(f"[Monitor] Could not parse timestamp for {project_id}: {timestamp_str}, error={e}")
                    
                    # Normalize project ID for file lookup
                    normalized_id = "SWESch_" + project_id.replace("/", "_")
                    file_exists = (os.path.exists(os.path.join(PROJECTS_DIR, normalized_id + ".xlsx")) or
                                   os.path.exists(os.path.join(PROJECTS_DIR, normalized_id + ".xlsb")))
                    
                    # Get SW Head (col C) and SWE Name (col D)
                    swh_head = cells.get(f"C{row_num}", {}).get('v', '')
                    swe_name = cells.get(f"D{row_num}", {}).get('v', '')
                    
                    results.append({
                        "project_id": project_id,
                        "file_id": normalized_id,
                        "stale": stale,
                        "file_exists": file_exists,
                        "swh_head": (swh_head or "").strip().upper(),
                        "swe_name": (swe_name or "").strip().upper(),
                    })
            
            print(f"[Monitor] Loaded {len(results)} projects from JSON cache")
            return results
            
        except Exception as e:
            print(f"[Monitor] Error reading from JSON cache: {e}, falling back to Excel")
    
    # FALLBACK: Read from Excel if JSON cache doesn't exist
    monitor_path = os.path.join(DATA_DIR, department, f"{department}_Monitor.xlsx")
    print(f"[Monitor] Reading from Excel: {monitor_path}")
    
    if not os.path.exists(monitor_path):
        print(f"[Monitor] File not found: {monitor_path}")
        return []
    
    # ... your existing Excel reading code here (the original _read_master_xlsx or _read_master_xlsb) ...
    if monitor_path.endswith(".xlsb"):
        return _read_master_xlsb(monitor_path)
    else:
        return _read_master_xlsx(monitor_path)

def _master_row_to_entry(pid, ba_val, ba_rgb=None, swh_head=None, swe_name=None):
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
        "swh_head":      (swh_head or "").strip().upper(),
        "swe_name":      (swe_name or "").strip().upper(),
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
                    # col C = index 2, col D = index 3, col F = index 5, col BA = index 52
                    f_val    = row[5].v  if len(row) > 5  else None
                    ba_val   = row[52].v if len(row) > 52 else None
                    swh_head = row[2].v  if len(row) > 2  else None
                    swe_name = row[3].v  if len(row) > 3  else None
                    entry = _master_row_to_entry(f_val, ba_val, swh_head=swh_head, swe_name=swe_name)
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
        c_cell  = row[2]  if len(row) > 2  else None
        f_cell  = row[5]  if len(row) > 5  else None
        d_cell  = row[3]  if len(row) > 3  else None
        ba_cell = row[52] if len(row) > 52 else None
        pid      = f_cell.value if f_cell else None
        ba_val   = ba_cell.value if ba_cell else None
        swh_head = c_cell.value if c_cell else None
        swe_name = d_cell.value if d_cell else None
        # Try to get fill colour for fallback staleness check
        ba_rgb = None
        if ba_cell:
            try:
                fg = ba_cell.fill.fgColor
                if fg and fg.type == "rgb" and fg.rgb not in ("00000000", "FFFFFFFF"):
                    ba_rgb = fg.rgb
            except Exception:
                pass
        entry = _master_row_to_entry(pid, ba_val, ba_rgb, swh_head=swh_head, swe_name=swe_name)
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

    # Load workbook with openpyxl for structure, formatting, and non-formula cells
    wb = openpyxl.load_workbook(filepath, data_only=False)  # Load with formulas
    ws = wb.active
    max_row = ws.max_row
    max_col = min(ws.max_column, max_col)
    cols = [gcl(i) for i in range(1, max_col + 1) if gcl(i) not in ("A","B","C","D")]  # hide cols A-D

    # Initialize PyCel for formula evaluation
    excel_compiler = None
    if PY_CEL_AVAILABLE:
        try:
            excel_compiler = ExcelCompiler(filepath)
            print("PyCel initialized for formula evaluation")
        except Exception as e:
            print(f"PyCel initialization failed: {e}, falling back to openpyxl only")
            excel_compiler = None
    else:
        print("PyCel not available, using openpyxl only")

    # DEBUG: Print all columns being sent to frontend
    print(f"DEBUG: Columns being sent to frontend: {cols}")
    print(f"DEBUG: Does 'R' in cols? {'R' in cols}")
    print(f"DEBUG: Column index of R: {cols.index('R') if 'R' in cols else 'NOT FOUND'}")
    
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
    DEFAULT_COL_WIDTH = getattr(ws.sheet_format, 'defaultColWidth', None) or 0.88
    col_widths  = {}
    col_hidden  = {}
    col_outline = {}

    for col in cols:
        try:
            cd = ws.column_dimensions.get(col)
            if cd is None:
                col_widths[col]  = max(30, round((DEFAULT_COL_WIDTH or 8) * 7.5))
                col_hidden[col]  = DEFAULT_COL_WIDTH < 2.0
                col_outline[col] = 0
            else:
                col_widths[col]  = max(30, round((cd.width or 8) * 7.5))
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
    col_groups = []
    try:
        i = 0
        while i < len(cols):
            col = cols[i]
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
                if group_cols[0] == "J" and "Q" in group_cols:
                    group_cols = group_cols[:group_cols.index("Q")]
                if group_cols[0] == "X" and "AB" in cols and "AB" not in group_cols:
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
    
    # Pre-calculate sheet name for PyCel
    sheet_name = ws.title
    
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            coord = cell.coordinate
            mi = merged_map.get(coord, {})

            if mi.get("skip"):
                cells[coord] = {"skip": True}
                continue

            # Get value using PyCel for formulas, openpyxl for static values
            v = None
            if excel_compiler and cell.data_type == 'f':  # Cell contains a formula
                try:
                    # Use PyCel to evaluate the formula
                    cell_ref = f"{sheet_name}!{coord}"
                    v = excel_compiler.evaluate(cell_ref)

                    # DEBUG for column AA
                    if cell.column_letter == "AA" and cell.row >= 9:
                        print(f"DEBUG AA: {coord} raw value: {v}, format: {cell.number_format}")
                    
                     # Handle date values
                    fmt = cell.number_format or ""
                    
                    # Check if the value is a date serial number (Excel stores dates as numbers)
                    if isinstance(v, (int, float)) and v > 1 and any(pattern in fmt.lower() for pattern in ['yy', 'mm', 'dd', 'mmm']):
                        # Convert Excel date serial to proper date
                        try:
                            from openpyxl.utils.datetime import from_excel
                            date_val = from_excel(v)
                            v = date_val.strftime("%d-%b-%y")
                        except Exception:
                            pass
                    elif isinstance(v, (datetime, date)):
                        if isinstance(v, datetime):
                            v = v.date()
                        v = v.strftime("%d-%b-%y")

                    elif isinstance(v, (int, float)):
                        if "%" in fmt:
                            v = f"{int(round(float(v) * 100))}%"
                            print(f"DEBUG AA: converted to: {v}")


                except Exception as e:
                    print(f"PyCel evaluation failed for {coord}: {e}, falling back to openpyxl")
                    # Fall back to openpyxl's value (might be None)
                    v = cell.value
                    if isinstance(v, (datetime, date)):
                        if isinstance(v, datetime): v = v.date()
                        v = v.strftime("%d-%b-%y")
            else:
                # Static value - use openpyxl
                v = cell.value
                # DEBUG for column AA
                if cell.column_letter == "AA" and cell.row >= 9:
                    print(f"DEBUG AA static: {coord} value: {v}, format: {cell.number_format}")
                
                if isinstance(v, (datetime, date)):
                    if isinstance(v, datetime): v = v.date()
                    v = v.strftime("%d-%b-%y")
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    fmt = cell.number_format or ""
                    if "%" in fmt:
                        v = f"{int(round(float(v) * 100))}%"
                    elif fmt == "00":
                        v = f"{int(v):02d}"
                    elif fmt == "0.0":
                        v = f"{float(v):.1f}"
                    elif fmt == "0.00":
                        v = f"{float(v):.2f}"
                elif v is not None:
                    v = str(v) if not isinstance(v, (int, str, bool)) else v

            c = {"v": v}

             # DEBUG: Check what's being stored for column R
            if cell.column_letter == "R" and cell.row >= 9 and cell.row <= 16:
                print(f"DEBUG STORE: {coord} stored value: {v}")

            # DEBUG: Check column R
            if cell.column_letter == "R":
                print(f"DEBUG: Column R, Row {cell.row}, Value: {v}")

            # Merge spans
            if mi.get("master"):
                if mi["rowspan"] > 1: c["rowspan"] = mi["rowspan"]
                if mi["colspan"] > 1: c["colspan"] = mi["colspan"]

            # Font — bold and size only
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

    # ... Continue with the rest of your existing get_raw_sheet() code
    # (editable cells detection, project_banner, left_panel, etc.)
    # Keep everything from here unchanged

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
                    # Make X, Y, Z always editable (regardless of fill)
                    # AF is always editable, AD uses separate logic
                    is_always_editable = col_letter in ["X", "Y", "Z"] or col_letter == "AF"
                    
                    if is_always_editable or _is_editable_fill(cell):
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
        cell = ws[ref]
        v = cell.value
        print(f"DEBUG _date: {ref} raw value: {v}, data_type: {cell.data_type}")
        
        if v is None:
            return ""
        
        # If it's a formula, try to get the calculated value
        if cell.data_type == 'f' and excel_compiler:
            try:
                cell_ref = f"{sheet_name}!{ref}"
                v = excel_compiler.evaluate(cell_ref)
                print(f"DEBUG _date: {ref} evaluated to: {v}")
            except Exception as e:
                print(f"DEBUG _date: {ref} evaluation failed: {e}")
        
        # Convert Excel date serial number to date string
        if isinstance(v, (int, float)):
            # Check if this is a date serial number (Excel dates are numbers > 1)
            if v > 1:
                try:
                    from openpyxl.utils.datetime import from_excel
                    date_val = from_excel(v)
                    v = date_val.strftime("%d %b %Y")
                except Exception:
                    pass
        
        if isinstance(v, (datetime, date)):
            d = v.date() if isinstance(v, datetime) else v
            return d.strftime("%d %b %Y")
        return str(v).strip() if v else ""

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

     # DEBUG: Test _date function for SWE cells
    print("=== DEBUG: SWE DATE CELL VALUES ===")
    for row_num in [28, 29, 30, 31, 32, 33]:
        swe_cell = ws[f"D{row_num}"]
        print(f"D{row_num}: raw value = {swe_cell.value}, data_type = {swe_cell.data_type}")
        swe_value = _date(f"D{row_num}")
        print(f"  _date(D{row_num}) returned: {swe_value}")
    print("==================================")

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

    # Determine last_modified:
    # Prefer the JSON sidecar mtime — it reflects the last app-side edit and is
    # written synchronously, so it's always accurate even while the background
    # Excel write is still pending.
    # Fall back to the Excel file mtime if no sidecar exists yet.
    last_modified_str = None
    try:
        json_path = _sheet_cache_path(project_id, sched_cache)
        if os.path.exists(json_path):
            ts = os.path.getmtime(json_path)
        else:
            ts = os.path.getmtime(filepath)
        last_modified_str = datetime.fromtimestamp(ts).strftime("%d %b %Y, %I:%M %p")
    except Exception:
        pass

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
        "last_modified":  last_modified_str,
    }

        # Write JSON sidecar cache for fast future loads
    _write_sheet_cache(project_id, result, sched_cache)
    
    # Generate system memory JSON when sheet cache is first created
    if not cached:  # This means we just created a new cache
        try:
            from sysmemory_compute import compute_sysmemory_json
            
            # Get the sheet data we just wrote
            sheet_data = _read_sheet_cache(project_id, sched_cache)
            if sheet_data:
                # Determine the system memory cache directory
                # sched_cache is like: data/SW/cache/schedules
                # We want: data/SW/cache/system_memory
                sysmemory_cache = os.path.join(os.path.dirname(sched_cache), "system_memory") if sched_cache else None
                if not sysmemory_cache:
                    # Fallback: derive from PROJECTS_DIR
                    base_dir = os.path.dirname(PROJECTS_DIR)
                    sysmemory_cache = os.path.join(base_dir, "cache", "system_memory")
                
                compute_sysmemory_json(project_id, sheet_data, sysmemory_cache)
                print(f"[Auto] System memory JSON created for {project_id}")
        except Exception as e:
            print(f"[Auto] Failed to create system memory JSON for {project_id}: {e}")
    
    return result

def get_monitor_sheet_data(filepath, monitor_type="SW"):
    """
    Read monitoring file without PyCel, just basic cell values.
    Creates JSON cache for fast subsequent loads.
    """
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter as gcl
    from datetime import datetime, date, time
    import json
    import os
    
        # Define cache path
    # filepath is: .../data/SW/SW_Monitor.xlsx
    # We want: .../data/SW/cache/monitoring/SW_Monitor.json
    base_dir = os.path.dirname(filepath)  # .../data/SW
    cache_dir = os.path.join(base_dir, "cache", "monitoring")
    os.makedirs(cache_dir, exist_ok=True)
    cache_path = os.path.join(cache_dir, f"{monitor_type}_Monitor.json")
    
    # Return cache if exists
    if os.path.exists(cache_path):
        try:
            with open(cache_path, 'r') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error reading monitor cache: {e}")
    
    # Read Excel file
    wb = load_workbook(filepath, data_only=True)
    ws = wb.active
    max_row = ws.max_row
    max_col = ws.max_column
    
    # Get all columns (A, B, C, ... up to max_col)
    cols = [gcl(i) for i in range(1, max_col + 1)]
    
    # Read cells
    cells = {}
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            v = cell.value
            # Handle datetime
            if isinstance(v, datetime):
                v = v.strftime("%d-%b-%y")
            # Handle date (without time)
            elif isinstance(v, date):
                v = v.strftime("%d-%b-%y")
            # Handle time (without date)
            elif isinstance(v, time):
                v = v.strftime("%H:%M")
            # Keep other values as they are
            cells[cell.coordinate] = {"v": v}
    
    result = {
        "cells": cells,
        "col_widths": {col: 64 for col in cols},
        "row_heights": {str(r): 20 for r in range(1, max_row + 1)},
        "max_row": max_row,
        "max_col": max_col,
        "cols": cols,
        "col_groups": [],
        "editable_fill": None,
        "project_banner": {},
        "info_rows": [],
        "header_rows": [],
        "left_panel": {},
        "last_modified": None
    }
    
    # Write cache
    with open(cache_path, 'w') as f:
        json.dump(result, f)
    
    wb.close()
    return result
    
def generate_sysmemory_json(filepath, project_id, sched_cache=None):
    """
    Generate system memory JSON from columns CY to DJ (rows 1-55).
    Uses PyCel to evaluate formulas.
    Stores in data/[discipline]/cache/system_memory/[project_id]_sysmemory.json
    """
    import os
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter as gcl
    
    # Determine discipline from the filepath
    # Path format: .../data/SW/SWESch/file.xlsx
    # We want: .../data/SW/cache/system_memory/
    parts = filepath.split(os.sep)
    
    # Find where 'data' is in the path
    data_index = None
    for i, part in enumerate(parts):
        if part == 'data':
            data_index = i
            break
    
    if data_index is not None and data_index + 1 < len(parts):
        discipline = parts[data_index + 1]
        base_path = os.sep.join(parts[:data_index + 1])
        sysmemory_cache = os.path.join(base_path, discipline, "cache", "system_memory")
    else:
        # Fallback to using DATA_DIR
        sysmemory_cache = os.path.join(DATA_DIR, "SW", "cache", "system_memory")
    
    os.makedirs(sysmemory_cache, exist_ok=True)
    
    sysmemory_path = os.path.join(sysmemory_cache, f"{project_id}_sysmemory.json")
    
    # Load workbook with formulas
    wb = load_workbook(filepath, data_only=False)
    ws = wb.active
    sheet_name = ws.title
    
    # Get column letters from CY to DJ
    # CY = 103rd column, DJ = 114th column
    start_col_idx = 103  # CY
    end_col_idx = 114    # DJ
    
    sysmemory_cols = [gcl(i) for i in range(start_col_idx, end_col_idx + 1)]
    
    # Initialize PyCel
    excel_compiler = None
    if PY_CEL_AVAILABLE:
        try:
            excel_compiler = ExcelCompiler(filepath)
            print(f"PyCel initialized for system memory generation: {project_id}")
        except Exception as e:
            print(f"PyCel initialization failed for system memory {project_id}: {e}")
    
    # Prepare system memory data
    sysmemory_data = {
        "project_id": project_id,
        "columns": sysmemory_cols,
        "rows": {},
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    
    # Loop through rows 1 to 55
    for row in range(1, 56):
        row_data = {}
        for col in sysmemory_cols:
            coord = f"{col}{row}"
            cell = ws[coord]
            value = None
            
            # If cell has formula and PyCel available, evaluate it
            if excel_compiler and cell.data_type == 'f':
                try:
                    cell_ref = f"{sheet_name}!{coord}"
                    value = excel_compiler.evaluate(cell_ref)
                except Exception as e:
                    print(f"System memory eval failed for {coord}: {e}")
                    # Fallback to openpyxl value
                    value = cell.value
            else:
                value = cell.value
            
            # Convert to appropriate type for JSON
            if isinstance(value, (datetime, date)):
                if isinstance(value, datetime):
                    value = value.date()
                value = value.strftime("%Y-%m-%d")
            elif isinstance(value, float):
                # Keep as float
                pass
            elif value is None:
                value = None
            
            row_data[col] = value
        
        sysmemory_data["rows"][str(row)] = row_data
    
    wb.close()
    
    # Write to JSON file
    try:
        with open(sysmemory_path, 'w') as f:
            json.dump(sysmemory_data, f, indent=2)
        print(f"System memory JSON saved: {sysmemory_path}")
        return sysmemory_path
    except Exception as e:
        print(f"Failed to save system memory JSON for {project_id}: {e}")
        return None

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
    from openpyxl.styles.numbers import is_date_format
    from datetime import datetime as _dt, timedelta
    import os
    
    projects_dir, _, _, _ = get_discipline_dirs(role)
    fpath = os.path.join(projects_dir, project_id + ".xlsx")
    
    if not os.path.exists(fpath):
        fpath = os.path.join(projects_dir, project_id + ".xlsb")
        if not os.path.exists(fpath):
            print(f"[Background] Excel file not found for {project_id}")
            return
    
    # Load workbook
    wb = load_workbook(fpath, keep_links=False)
    ws = wb.active

    # Fix openpyxl bug: integer values in date-formatted cells cause 'int has no attr year'
    # Scan ALL cells and fix any ints in date-formatted cells
    for row in ws.iter_rows():
        for cell in row:
            # Check if cell has an integer value and date format
            if isinstance(cell.value, int) and cell.number_format:
                try:
                    if is_date_format(cell.number_format):
                        # Convert Excel serial date integer to Python datetime
                        try:
                            # Excel serial date: 1 = 1900-01-01
                            # openpyxl uses 1899-12-30 as epoch
                            from openpyxl.utils.datetime import from_excel
                            cell.value = from_excel(cell.value)
                        except Exception:
                            # Fallback manual conversion
                            epoch = _dt(1899, 12, 30)
                            cell.value = epoch + timedelta(days=cell.value)
                except Exception:
                    pass
            # Also handle float dates that might be integers
            elif isinstance(cell.value, float) and cell.number_format:
                try:
                    if is_date_format(cell.number_format) and cell.value > 1:
                        from openpyxl.utils.datetime import from_excel
                        cell.value = from_excel(cell.value)
                except Exception:
                    pass
    
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
    
    # One more pass to catch any remaining ints in date cells before saving
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, int) and cell.number_format:
                try:
                    if is_date_format(cell.number_format):
                        from openpyxl.utils.datetime import from_excel
                        cell.value = from_excel(cell.value)
                except Exception:
                    pass
    
    # Save workbook
    wb.save(fpath)
    print(f"[Background] Excel file saved: {fpath}")

def queue_excel_write(project_id, updates, role):
    """Queue a project for background Excel write"""
    _start_background_worker()  # Ensure worker is running
    _excel_write_queue.put((project_id, updates, role))

def _do_monitor_excel_write(department, updates):
    """Write monitor cache updates to actual Excel file in background"""
    import os
    from openpyxl import load_workbook

    
    monitor_path = os.path.join(DATA_DIR, department, f"{department}_Monitor.xlsx")
    
    if not os.path.exists(monitor_path):
        print(f"[Monitor Background] File not found: {monitor_path}")
        return
    
    try:
        # Load workbook
        wb = load_workbook(monitor_path)
        ws = wb["SWMon"] if "SWMon" in wb.sheetnames else wb.active
        
        # Apply updates
        for coord, value in updates.items():
            if value is not None:
                ws[coord] = value
        
        # Save workbook
        wb.save(monitor_path)
        print(f"[Monitor Background] Excel saved: {monitor_path}")
        
    except Exception as e:
        print(f"[Monitor Background] Error writing to Excel: {e}")


def queue_monitor_excel_write(department, updates):
    """Queue monitor Excel write to background thread"""
    import threading
    
    def _run():
        _do_monitor_excel_write(department, updates)
    
    thread = threading.Thread(target=_run, daemon=True)
    thread.start()


def update_monitor_timestamp(project_id, timestamp, role="sw_tl"):
    """Update the BA column timestamp for a project in monitor JSON cache and queue Excel write"""
    from datetime import datetime
    import json
    import os
    
    # Map role to department
    role_to_dept = {
        "sw_tl": "SW",
        "hw_tl": "HW", 
        "mfg_tl": "MFG",
        "pm": "PM",
        "admin": "SW",
        "head": "SW"
    }
    
    department = role_to_dept.get(role, "SW")
    monitor_cache_dir = os.path.join(DATA_DIR, department, "cache", "monitoring")
    monitor_cache_path = os.path.join(monitor_cache_dir, f"{department}_Monitor.json")
    
    if not os.path.exists(monitor_cache_path):
        print(f"[Monitor] Cache not found: {monitor_cache_path}")
        return False
    
    try:
        # Load monitor cache
        with open(monitor_cache_path, 'r') as f:
            monitor_data = json.load(f)
        
        cells = monitor_data.get('cells', {})
        
        # Find the row with matching project ID in COLUMN B
        found_row = None
        for coord, cell_info in cells.items():
            if coord.startswith('B') and cell_info.get('v') == project_id:
                found_row = int(coord[1:])
                break
        
        if found_row:
            # Update column BA in cache
            ba_coord = f"BA{found_row}"
            if ba_coord not in cells:
                cells[ba_coord] = {}
            
            cells[ba_coord]['v'] = timestamp
            cells[ba_coord]['updated'] = True
            
            # Write back to JSON cache
            with open(monitor_cache_path, 'w') as f:
                json.dump(monitor_data, f, indent=2)
            
            print(f"[Monitor] Updated cache for {project_id} at row {found_row}: {timestamp}")
            
            # Queue Excel write to background
            queue_monitor_excel_write(department, {ba_coord: timestamp})
            
            return True
        else:
            print(f"[Monitor] Project {project_id} not found in monitor cache (searched column F)")
            return False
            
    except Exception as e:
        print(f"[Monitor] Failed to update timestamp: {e}")
        return False