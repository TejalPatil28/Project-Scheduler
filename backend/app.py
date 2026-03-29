from flask import Flask, request, jsonify, session, send_from_directory
from flask_cors import CORS
import bcrypt
import os
from datetime import datetime, date
from excel_db import (
    get_all_users,
    get_user_by_username,
    get_projects_for_user,
    get_project_by_id,
    get_tasks,
    update_tasks_bulk,
    create_user,
    update_user,
    delete_user,
    update_index_entry,
    get_master_projects,
    get_discipline_dirs,
    OWNER_MAP,
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
    if role not in ("admin", "head"):
        if "all_master_projects" not in _master_projects_cache:
            from excel_db import get_all_monitor_projects
            _master_projects_cache["all_master_projects"] = get_all_monitor_projects(role)
        master_list = _master_projects_cache["all_master_projects"]
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
    # Row-based updates (from sheet input cells, have _row key) bypass this check
    if role not in ("admin", "head", "pm"):
        allowed_owner = OWNER_MAP.get(role)
        for item in data:
            if "_row" in item:
                continue  # row-based sheet edits skip owner check
            if item.get("owner") != allowed_owner:
                return jsonify({"error": f"You can only update {allowed_owner} tasks"}), 403

    ok, msg, now_str = update_tasks_bulk(project_id, data, role=role)
    if not ok:
        return jsonify({"error": msg}), 404

    # Reload the in-memory cache from the JSON sidecar that _update_sheet_cache
    # just wrote. This ensures the next sheet fetch returns the updated cell values,
    # not the stale pre-save data that was sitting in _sheet_cache.
    from excel_db import _read_sheet_cache, get_discipline_dirs
    _, _, sched_cache, _ = get_discipline_dirs(role)
    fresh = _read_sheet_cache(project_id, sched_cache)
    if fresh:
        _sheet_cache[project_id] = fresh
    elif project_id in _sheet_cache:
        # JSON sidecar missing (edge case) — at least patch timestamp so it's not wrong
        if now_str:
            _sheet_cache[project_id]["last_modified"] = now_str

    return jsonify({"message": msg, "last_modified": now_str})

@app.route("/api/projects/<project_id>/sheet", methods=["GET"])
@login_required
def get_sheet_data(project_id):
    """Return raw cell data for the Excel-mirror UI. Cached in memory."""
    from excel_db import get_raw_sheet, PROJECTS_DIR
    user = get_current_user()
    role = user["role"]
    is_readonly = role in ("admin", "head")

    # Check in-memory cache first (fastest)
    if project_id in _sheet_cache:
        data = _sheet_cache[project_id]
        if is_readonly:
            # Return a copy with editable flags stripped — don't mutate the cache
            import copy
            data = copy.deepcopy(data)
            for cell in data.get("cells", {}).values():
                cell.pop("editable", None)
        return jsonify(data)

    fpath = os.path.join(PROJECTS_DIR, project_id + ".xlsx")
    if not os.path.exists(fpath):
        fpath_b = os.path.join(PROJECTS_DIR, project_id + ".xlsb")
        if os.path.exists(fpath_b):
            fpath = fpath_b
        else:
            return jsonify({"error": "Project file not found"}), 404
    try:
        data = get_raw_sheet(fpath)
        # Inject last_modified from file mtime if get_raw_sheet didn't already
        if not data.get("last_modified"):
            from datetime import datetime as _dt
            try:
                ts = os.path.getmtime(fpath)
                data["last_modified"] = _dt.fromtimestamp(ts).strftime("%d %b %Y, %I:%M %p")
            except Exception:
                data["last_modified"] = None
        _sheet_cache[project_id] = data  # store in memory
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
    if "all_master_projects" not in _master_projects_cache:
        from excel_db import get_all_monitor_projects
        _master_projects_cache["all_master_projects"] = get_all_monitor_projects(user["role"])
    return jsonify(_master_projects_cache["all_master_projects"])


@app.route("/api/monitor/sheet", methods=["GET"])
@login_required
def get_monitor_sheet():
    """Return the raw sheet data for the department monitoring file."""
    from excel_db import get_monitor_sheet_data, PROJECTS_DIR
    import os
    
    user = get_current_user()
    role = user.get("role", "sw_tl")
    
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
    
    # Construct path to department monitoring file
    monitor_filename = f"{department}_Monitor.xlsx"
    dept_dir = os.path.dirname(PROJECTS_DIR)  # This gives data/SW/
    fpath = os.path.join(dept_dir, monitor_filename)
    
    print(f"[Monitor] Looking for: {fpath}")
    
    if not os.path.exists(fpath):
        print(f"[Monitor] File not found: {fpath}")
        return jsonify({"error": f"No monitoring file found for {department}"}), 404
    
    # Check in-memory cache
    cache_key = "monitor_" + department
    if cache_key in _sheet_cache:
        return jsonify(_sheet_cache[cache_key])
    
    try:
        data = get_monitor_sheet_data(fpath, department)
        _sheet_cache[cache_key] = data
        return jsonify(data)
    except Exception as e:
        print(f"[Monitor] Error: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

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
        monitor_cache_dir = os.path.join(DATA_DIR, dept, "cache", "monitoring")
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

# ── User management (admin only) ──────────────────────────────
VALID_ROLES = ["admin", "head", "pm", "hw_tl", "sw_tl", "mfg_tl"]

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


if __name__ == "__main__":
    print("Starting Project Scheduler on http://localhost:5000")
    app.run(debug=True, host="0.0.0.0", port=5000)
