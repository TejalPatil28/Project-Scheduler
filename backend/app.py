from flask import Flask, request, jsonify, session, send_from_directory
from flask_cors import CORS
import bcrypt
import os
from datetime import datetime, date
from excel_db import (
    get_all_users,
    get_user_by_username,
    get_all_monitor_projects,
    get_projects_for_user,
    get_project_by_id,
    get_tasks,
    update_tasks_bulk,
    create_user,
    update_user,
    delete_user,
    get_master_projects,
    get_discipline_dirs,
    OWNER_MAP,
    update_monitor_cell,
    ROLE_TO_DEPT,      
    OWNER_MAP,
    update_monitor_timestamp,
    load_user_overdue_status,
)

app = Flask(__name__, static_folder=None)

# In-memory sheet cache (JSON sidecar handles persistence across restarts)
_sheet_cache = {}
# In-memory cache for master project lists (monitor file is read-only from outside)
_master_projects_cache = {}
app.secret_key = "scheduler_excel_secret_2024"
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_HTTPONLY"] = True

CORS(app, supports_credentials=True)

# Simple in-memory sheet cache
# ── Startup: build index ───────────────────────────────────────




FRONTEND_DIR = os.path.join(os.path.dirname(__file__), "..", "frontend")

# ── Serve frontend ─────────────────────────────────────────────
@app.route("/")
def index():
    return send_from_directory(FRONTEND_DIR, "index.html")

@app.route("/<path:filename>")
def static_files(filename):
    return send_from_directory(FRONTEND_DIR, filename)

# ── Auth helpers ───────────────────────────────────────────────
def get_current_user():
    username = session.get("username")
    if not username:
        return None
    return get_user_by_username(username)

def login_required(fn):
    from functools import wraps
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not get_current_user():
            return jsonify({"error": "Not authenticated"}), 401
        return fn(*args, **kwargs)
    return wrapper

# ── Auth routes ────────────────────────────────────────────────
@app.route("/api/auth/login", methods=["POST"])
def login():
    data = request.get_json()
    if not data or not data.get("username") or not data.get("password"):
        return jsonify({"error": "Username and password required"}), 400

    user = get_user_by_username(data["username"])
    if not user:
        return jsonify({"error": "Invalid credentials"}), 401

    if not bcrypt.checkpw(data["password"].encode(), user["password_hash"].encode()):
        return jsonify({"error": "Invalid credentials"}), 401

    session["username"] = user["username"]

    # Load overdue status for this user and store in session
    from excel_db import load_user_overdue_status
    session["overdue_map"] = load_user_overdue_status(user)

    return jsonify({
        "name":     user["name"],
        "username": user["username"],
        "role":     user["role"],
    })

@app.route("/api/auth/logout", methods=["POST"])
def logout():
    session.clear()
    return jsonify({"message": "Logged out"})

@app.route("/api/auth/me", methods=["GET"])
@login_required
def me():
    u = get_current_user()
    return jsonify({
        "name":       u["name"],
        "short_name": u["short_name"],
        "username":   u["username"],
        "role":       u["role"],
    })

# ── Projects routes ────────────────────────────────────────────
@app.route("/api/projects", methods=["GET"])
@login_required
def list_projects():
    user = get_current_user()
    projects = get_projects_for_user(user)  # role-aware internally
    return jsonify(projects)

@app.route("/api/projects/<project_id>", methods=["GET"])
@login_required
def get_project(project_id):
    user = get_current_user()
    project = get_project_by_id(project_id, role=user["role"])
    if not project:
        return jsonify({"error": "Project not found"}), 404

    # Access check — if project is in user's monitoring file they have access
    role = user["role"]
    if role != "admin" and not role.endswith("_head"):
        cache_key = f"all_master_projects_{role}"
        if cache_key not in _master_projects_cache:
            _master_projects_cache[cache_key] = get_all_monitor_projects(role)
        master_list = _master_projects_cache[cache_key]
        allowed_ids = {m.get("file_id") for m in master_list}
        if project_id not in allowed_ids:
            return jsonify({"error": "Access denied"}), 403

    return jsonify(project)



@app.route("/api/projects/<project_id>/tasks", methods=["GET"])
@login_required
def get_project_tasks(project_id):
    user = get_current_user()
    role = user["role"]
    owner_filter = OWNER_MAP.get(role)
    tasks = get_tasks(project_id, owner_filter=owner_filter, role=role)
    return jsonify(tasks)


@app.route("/api/projects/<project_id>/tasks", methods=["PUT"])
@login_required
def save_tasks(project_id):
    user = get_current_user()
    role = user["role"]
    data = request.get_json()

    if not data or not isinstance(data, list):
        return jsonify({"error": "Expected list of task updates"}), 400

    # Verify each task's owner matches user role (unless admin/head/pm)
    if role not in ("admin", "head", "pm"):
        allowed_owner = OWNER_MAP.get(role)
        for item in data:
            if "_row" in item:
                continue
            if item.get("owner") != allowed_owner:
                return jsonify({"error": f"You can only update {allowed_owner} tasks"}), 403

    ok, msg, now_str = update_tasks_bulk(project_id, data, role=role)
    if not ok:
        return jsonify({"error": msg}), 404

    # AFTER SAVE: Refresh the overdue_map for this project in session
    from excel_db import read_sysmemory_json
    overdue_map = session.get("overdue_map", {})
    
    # Refresh this project's data
    fresh_task_map = read_sysmemory_json(project_id, role)
    if fresh_task_map:
        overdue_map[project_id] = fresh_task_map
    else:
        # If no data, remove from map
        overdue_map.pop(project_id, None)
    
    session["overdue_map"] = overdue_map

    # Reload the in-memory cache
    from excel_db import _read_sheet_cache, get_discipline_dirs
    _, _, sched_cache, _ = get_discipline_dirs(role)
    fresh = _read_sheet_cache(project_id, sched_cache)

    sheet_cache_key = f"{role}:{project_id}" 

    if fresh:
        _sheet_cache[sheet_cache_key] = fresh
    elif sheet_cache_key in _sheet_cache:
        if now_str:
            _sheet_cache[sheet_cache_key]["last_modified"] = now_str

    return jsonify({"message": msg, "last_modified": now_str})

@app.route("/api/projects/<project_id>/sheet", methods=["GET"])
@login_required
def get_sheet_data(project_id):
    """Return raw cell data for the Excel-mirror UI. Cached in memory."""
    from excel_db import get_raw_sheet, get_discipline_dirs
    user = get_current_user()
    role = user["role"]
    is_readonly = role in ("admin", "head")

    # NEW: role-aware in-memory key
    sheet_cache_key = f"{role}:{project_id}"

    # Check in-memory cache first (fastest)
    if sheet_cache_key in _sheet_cache:
        data = _sheet_cache[sheet_cache_key]
        if is_readonly:
            # Return a copy with editable flags stripped — don't mutate the cache
            import copy
            data = copy.deepcopy(data)
            for cell in data.get("cells", {}).values():
                cell.pop("editable", None)
        return jsonify(data)

    # CHANGED: capture sched_cache too
    projects_dir, _, sched_cache, _ = get_discipline_dirs(role)
    
    fpath = os.path.join(projects_dir, project_id + ".xlsx")
    if not os.path.exists(fpath):
        fpath_b = os.path.join(projects_dir, project_id + ".xlsb")
        if os.path.exists(fpath_b):
            fpath = fpath_b
        else:
            return jsonify({"error": "Project file not found"}), 404
    try:
        data = get_raw_sheet(fpath, role=role, sched_cache=sched_cache)
        # Inject last_modified from file mtime if get_raw_sheet didn't already
        if not data.get("last_modified"):
            from datetime import datetime as _dt
            try:
                ts = os.path.getmtime(fpath)
                data["last_modified"] = _dt.fromtimestamp(ts).strftime("%d %b %Y, %I:%M %p")
            except Exception:
                data["last_modified"] = None
        _sheet_cache[sheet_cache_key] = data  # store in memory
        if is_readonly:
            import copy
            data = copy.deepcopy(data)
            for cell in data.get("cells", {}).values():
                cell.pop("editable", None)
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/master/projects", methods=["GET"])
@login_required
def master_projects():
    user = get_current_user()
    role = user["role"]
    cache_key = f"all_master_projects_{role}"   # <-- role-specific key
    if cache_key not in _master_projects_cache:
        from excel_db import get_all_monitor_projects
        _master_projects_cache[cache_key] = get_all_monitor_projects(role)
    return jsonify(_master_projects_cache[cache_key])

@app.route("/api/monitor/sheet", methods=["GET"])
@login_required
def get_monitor_sheet():
    """Return the raw sheet data for the department monitoring file."""
    from excel_db import get_monitor_sheet_data, find_monitor_file, load_user_overdue_status
    import os
    
    user = get_current_user()
    role = user.get("role", "sw_tl")
    
    # Map role to department
    department = ROLE_TO_DEPT.get(role, "SW")
    
    # Use find_monitor_file instead
    fpath = find_monitor_file(department)
    
    print(f"[Monitor] Looking for: {fpath}")
    
    if not os.path.exists(fpath):
        print(f"[Monitor] File not found: {fpath}")
        return jsonify({"error": f"No monitoring file found for {department}"}), 404
    
    try:
        data = get_monitor_sheet_data(fpath, department)
        
        # ALWAYS refresh overdue_map when monitor loads
        fresh_overdue_map = load_user_overdue_status(user)
        session["overdue_map"] = fresh_overdue_map
        
        return jsonify({
            "sheet": data,
            "overdue": fresh_overdue_map
        })
    except Exception as e:
        print(f"[Monitor] Error: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500
        
@app.route("/api/projects/<project_id>/monitor-timestamp", methods=["POST"])
@login_required
def save_monitor_timestamp(project_id):
    """Write the save timestamp into the BA column of the monitor JSON/Excel."""
    user = get_current_user()
    role = user["role"]

    data = request.get_json()
    display_ts = (data or {}).get("timestamp", "")

    # Convert display format "29 Mar 2026, 08:04 PM" -> raw "29-03-2026 20:04:00"
    # to match the Excel cell format dd-mm-yyyy hh:mm:ss
    raw_ts = display_ts
    try:
        raw_ts = datetime.strptime(display_ts, "%d %b %Y, %I:%M %p").strftime("%d-%m-%Y %H:%M:%S")
    except Exception:
        raw_ts = datetime.now().strftime("%d-%m-%Y %H:%M:%S")

    ok = update_monitor_timestamp(project_id, raw_ts, role=role)
    if not ok:
        # Non-fatal — project may not be in this user's monitor file
        return jsonify({"message": "Project not found in monitor, skipped"}), 200

    # Patch the in-memory master projects cache in-place so the sidebar
    # reflects the updated stale status immediately — no disk read needed.
    # With:
    cache_key = f"all_master_projects_{role}"
    cached = _master_projects_cache.get(cache_key, [])
    for entry in cached:
        if entry.get("file_id") == project_id:
            entry["stale"] = False
            break

    return jsonify({"message": "Monitor timestamp updated", "timestamp": raw_ts})


@app.route("/api/debug", methods=["GET"])
@login_required
def debug():
    user = get_current_user()
    from excel_db import get_all_projects
    projects = get_all_projects()
    return jsonify({
        "logged_in_user": {
            "name":       user["name"],
            "username":   user["username"],
            "short_name": user.get("short_name", "MISSING"),
            "role":       user["role"],
        },
        "projects": [{
            "id":           p["id"],
            "pm_name":      p.get("pm_name", ""),
            "hw_tl_name":   p.get("hw_tl_name", ""),
            "sw_tl_name":   p.get("sw_tl_name", ""),
            "mfg_tl_name":  p.get("mfg_tl_name", ""),
        } for p in projects]
    })


@app.route("/api/admin/rebuild-index", methods=["POST"])
@login_required
def rebuild_index():
    user = get_current_user()
    if user["role"] != "admin":
        return jsonify({"error": "Access denied"}), 403
    # Clear in-memory sheet cache so all projects reload fresh
    _sheet_cache.clear()
    _master_projects_cache.clear()
    return jsonify({"message": "Cache cleared successfully"})

@app.route("/api/admin/clear-monitor-cache", methods=["POST"])
@login_required
def clear_monitor_cache():
    user = get_current_user()
    if user["role"] != "admin":
        return jsonify({"error": "Access denied"}), 403
    
    # Clear in-memory monitor cache
    global _sheet_cache
    keys_to_remove = [k for k in _sheet_cache.keys() if k.startswith("monitor_")]
    for key in keys_to_remove:
        del _sheet_cache[key]
    
    # Also delete JSON cache files for monitor
    from excel_db import DATA_DIR
    import os
    
    # Delete monitor cache for all departments
    departments = ["SW", "HW", "MFG", "PM"]
    for dept in departments:
        from excel_db import get_cache_path
        monitor_cache_dir = get_cache_path(dept, "Monitoring")
        if os.path.exists(monitor_cache_dir):
            for filename in os.listdir(monitor_cache_dir):
                if filename.endswith(".json"):
                    filepath = os.path.join(monitor_cache_dir, filename)
                    try:
                        os.remove(filepath)
                        print(f"Deleted monitor cache: {filepath}")
                    except Exception as e:
                        print(f"Could not delete {filepath}: {e}")
    
    return jsonify({"message": "Monitor cache cleared successfully"})

# __NEW PROJECT CREATION LOGIC__________________________________

@app.route("/api/projects/create", methods=["POST"])
@login_required
def create_project():
    """Create a new project from template with form data"""
    from excel_db import DATA_DIR, get_cache_path, generate_sysmemory_json, get_discipline_dirs
    import shutil
    import openpyxl
    import os
    from datetime import datetime
    
    user = get_current_user()
    # Only PM or pm_head can create projects
    if user["role"] not in ("pm", "pm_head"):
        return jsonify({"error": "Access denied"}), 403
    
    data = request.get_json()
    
    # Get OR Number for filename
    or_number = data.get("or_number", "")
    if not or_number:
        return jsonify({"error": "OR Number is required"}), 400
    
    # Get section for filename
    section = data.get("section", "")
    
    # Generate filename: PrjSch_{or_number}_{section}.xlsx
    # Replace any slashes with underscores
    or_number_clean = or_number.replace("/", "_")
    section_clean = section.replace("/", "_") if section else ""
    
    if section_clean:
        filename = f"PrjSch_{or_number_clean}_{section_clean}.xlsx"
    else:
        filename = f"PrjSch_{or_number_clean}.xlsx"
    
    # Save to PM schedule folder
    projects_dir, _, _, _ = get_discipline_dirs("pm_head")
    new_file_path = os.path.join(projects_dir, filename)
    
    # Check if file already exists
    if os.path.exists(new_file_path):
        return jsonify({"error": f"Project already exists: {filename}"}), 409
    
    # Template path
    template_path = os.path.join(DATA_DIR, "Template", "PrjSch_Template_V2.0.xlsx")
    if not os.path.exists(template_path):
        return jsonify({"error": "Template file not found"}), 500
    
    # Copy template to new location
    shutil.copy2(template_path, new_file_path)
    
    # Open and populate the new file
    wb = openpyxl.load_workbook(new_file_path)
    ws = wb.active  # PrjSch sheet
    
    # Helper function to clean empty strings
    def clean_value(value):
        return value if value and value != "" else None
    
    def clean_number(value):
        if not value or value == "":
            return 0
        try:
            return float(value)
        except (ValueError, TypeError):
            return 0
    
    # ============================================
    # MAP FORM DATA TO EXCEL CELLS
    # ============================================
    
    # Header Section (Rows 1-3)
    ws["D1"] = clean_value(data.get("or_number"))
    ws["D2"] = clean_value(data.get("master_or"))
    ws["D3"] = clean_value(data.get("client_po"))
    ws["I1"] = clean_value(data.get("quote_number"))
    ws["I2"] = clean_value(data.get("sales_engineer"))
    ws["I3"] = clean_value(data.get("sales_manager"))
    
    # Project Details (Rows 9-15)
    ws["D9"] = clean_number(data.get("po_value"))
    ws["D10"] = clean_value(data.get("customer_name"))
    ws["D11"] = clean_value(data.get("end_customer"))
    ws["D12"] = clean_value(data.get("consultant"))
    ws["D13"] = clean_value(data.get("project_desc"))
    ws["D14"] = clean_value(data.get("section"))
    ws["D15"] = clean_value(data.get("mfg_loc", "GON"))
    
    # Efforts (Rows 17-21)
    ws["D17"] = clean_number(data.get("hw_efforts"))
    ws["D18"] = clean_number(data.get("std_panels"))
    ws["D19"] = clean_number(data.get("act_panels"))
    ws["D20"] = clean_number(data.get("sw_efforts"))
    ws["D21"] = clean_number(data.get("mfg_efforts"))
    
    # Actual Efforts (Rows 17,20,21 - Column B)
    ws["B17"] = clean_number(data.get("actual_hw_efforts"))
    ws["B20"] = clean_number(data.get("actual_sw_efforts"))
    ws["B21"] = clean_number(data.get("actual_mfg_efforts"))
    
    # Dates (Rows 23-34, Column C)
    ws["C23"] = clean_value(data.get("po_date"))
    ws["C24"] = clean_value(data.get("opf_recpt"))
    ws["C25"] = clean_value(data.get("hw_input"))
    ws["C26"] = clean_value(data.get("dwg_sub"))
    ws["C27"] = clean_value(data.get("dwg_appr"))
    ws["C28"] = clean_value(data.get("hw_fat"))
    ws["C29"] = clean_value(data.get("dispatch"))
    ws["C30"] = clean_value(data.get("sw_input"))
    ws["C31"] = clean_value(data.get("sw_fat"))
    ws["C32"] = clean_value(data.get("install"))
    ws["C33"] = clean_value(data.get("precomm"))
    ws["C34"] = clean_value(data.get("comm"))
    
    # Stakeholders (Rows 36-42, Column C)
    ws["C36"] = clean_value(data.get("sh_sales"))
    ws["C37"] = clean_value(data.get("sh_hw"))
    ws["C38"] = clean_value(data.get("sh_sw"))
    ws["C39"] = clean_value(data.get("sh_byr"))
    ws["C40"] = clean_value(data.get("sh_mfg"))
    ws["C41"] = clean_value(data.get("sh_ec"))
    ws["C42"] = clean_value(data.get("sh_ac"))
    
    # Scope Selection (Row 9, Columns AG-AK)
    ws["AG9"] = clean_value(data.get("scope_hw", "NO"))
    ws["AH9"] = clean_value(data.get("scope_sw", "NO"))
    ws["AI9"] = clean_value(data.get("scope_mfg", "NO"))
    ws["AJ9"] = clean_value(data.get("scope_inst", "NO"))
    ws["AK9"] = clean_value(data.get("scope_com", "NO"))
    
    # LD Fields (Row AF1-AF4)
    ws["AF1"] = clean_value(data.get("ld_date"))
    ws["AF2"] = clean_number(data.get("ld_maxwk"))
    ws["AF3"] = clean_number(data.get("ld_maxov"))
    ws["AF4"] = clean_value(data.get("ld_remarks"))
    
    # Warranty (Row AF6)
    ws["AF6"] = clean_value(data.get("warranty"))
    
    # Actuals (Rows 43-52)
    ws["B43"] = clean_number(data.get("balance_panels"))
    ws["B44"] = clean_number(data.get("panel_disp_act"))
    ws["B45"] = clean_number(data.get("est_va_pct"))
    ws["B46"] = clean_number(data.get("est_va"))
    ws["B47"] = clean_number(data.get("est_sm_pct"))
    ws["B48"] = clean_number(data.get("est_sm"))
    ws["B49"] = clean_number(data.get("act_va_pct"))
    ws["B50"] = clean_number(data.get("act_va"))
    ws["B51"] = clean_number(data.get("act_sm_pct"))
    ws["B52"] = clean_number(data.get("act_sm"))
    ws["C53"] = clean_value(data.get("reason_remark"))
    
    # Save the file
    wb.save(new_file_path)
    wb.close()
    
    # Generate JSON cache for the new project
    project_id = filename.replace(".xlsx", "")
    _, _, sched_cache, _ = get_discipline_dirs("pm_head")
    generate_sysmemory_json(new_file_path, project_id, sched_cache)
    
    print(f"[Create Project] Created: {filename}")
    
    return jsonify({
        "message": "Project created successfully",
        "filename": filename,
        "project_id": project_id
    }), 201

# ── User management (admin only) ──────────────────────────────
VALID_ROLES = ["admin", "sw_head", "hw_head", "mfg_head", "pm_head", "pm", "hw_tl", "sw_tl", "mfg_tl"]

@app.route("/api/users", methods=["GET"])
@login_required
def list_users():
    user = get_current_user()
    if user["role"] != "admin":
        return jsonify({"error": "Access denied"}), 403
    users = get_all_users()
    # Don't send password hashes to frontend
    return jsonify([{
        "name":       u["name"],
        "short_name": u["short_name"],
        "username":   u["username"],
        "role":       u["role"],
    } for u in users])

@app.route("/api/users", methods=["POST"])
@login_required
def add_user():
    user = get_current_user()
    if user["role"] != "admin":
        return jsonify({"error": "Access denied"}), 403

    data = request.get_json()
    for f in ["name", "short_name", "username", "password", "role"]:
        if not data.get(f):
            return jsonify({"error": f"{f} is required"}), 400
    if data["role"] not in VALID_ROLES:
        return jsonify({"error": "Invalid role"}), 400

    ok, msg = create_user(data["name"], data["short_name"], data["username"], data["password"], data["role"])
    if not ok:
        return jsonify({"error": msg}), 409
    return jsonify({"message": msg}), 201

@app.route("/api/users/<username>", methods=["PUT"])
@login_required
def edit_user(username):
    user = get_current_user()
    if user["role"] != "admin":
        return jsonify({"error": "Access denied"}), 403

    data = request.get_json()
    if "role" in data and data["role"] not in VALID_ROLES:
        return jsonify({"error": "Invalid role"}), 400

    ok, msg = update_user(username, data)
    if not ok:
        return jsonify({"error": msg}), 404
    return jsonify({"message": msg})

@app.route("/api/users/<username>", methods=["DELETE"])
@login_required
def remove_user(username):
    user = get_current_user()
    if user["role"] != "admin":
        return jsonify({"error": "Access denied"}), 403
    if username == user["username"]:
        return jsonify({"error": "Cannot delete your own account"}), 400

    ok, msg = delete_user(username)
    if not ok:
        return jsonify({"error": msg}), 404
    return jsonify({"message": msg})

@app.route("/api/refresh-file-status", methods=["POST"])
def refresh_file_status():
    from excel_db import get_all_monitor_projects, find_schedule_folder, DEPT_CONFIG
    import os
    
    user = get_current_user()
    if not user:
        return jsonify({"error": "Not authenticated"}), 401
    
    role = user.get("role", "sw_tl")
    projects = get_all_monitor_projects(role)
    
    # Map role to department
    
    dept = ROLE_TO_DEPT.get(role, "SW")
    projects_dir = find_schedule_folder(dept)
    
    # Get the correct prefix for this department
    prefix = DEPT_CONFIG[dept]["sched_prefix"]
    
    print("=== REFRESH DEBUG ===")
    
    for project in projects:
        project_id = project.get("project_id")
        if project_id:
            # Use department-specific prefix instead of hardcoded "SWESch_"
            normalized_id = prefix + "_" + project_id.replace("/", "_")
            excel_path_xlsx = os.path.join(projects_dir, normalized_id + ".xlsx")
            excel_path_xlsb = os.path.join(projects_dir, normalized_id + ".xlsb")
            file_exists = os.path.exists(excel_path_xlsx) or os.path.exists(excel_path_xlsb)
            
            print(f"Project ID: {project_id}")
            print(f"  Normalized: {normalized_id}")
            print(f"  Looking for: {excel_path_xlsx}")
            print(f"  Exists: {file_exists}")
            
            project["file_exists"] = file_exists
    
    return jsonify({"projects": projects})

@app.route("/api/monitor/cell", methods=["POST"])
@login_required
def update_monitor_cell():
    """Save a single cell edit in the monitor file"""
    user = get_current_user()
    role = user["role"]
    
    data = request.get_json()
    if not data:
        return jsonify({"error": "No data provided"}), 400
    
    col = data.get("col")
    row = data.get("row")
    value = data.get("value")
    
    if not col or not row:
        return jsonify({"error": "Column and row required"}), 400
    
    department = ROLE_TO_DEPT.get(role, "SW")
    
    # Update monitor cache and queue Excel write
    coord = f"{col}{row}"
    
    try:
        from excel_db import update_monitor_cell
        result = update_monitor_cell(department, coord, value, role)
        if result:
            return jsonify({"message": "Saved", "coord": coord, "value": value})
        else:
            return jsonify({"error": "Failed to save"}), 500
    except Exception as e:
        print(f"Error saving monitor cell: {e}")
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    print("Starting Project Scheduler on http://localhost:5000")
    app.run(debug=True, host="0.0.0.0", port=5000)
