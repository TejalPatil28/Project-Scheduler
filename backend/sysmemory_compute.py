"""
compute_sysmemory_json()

Replaces generate_sysmemory_json() — no Excel open, no PyCel.
All values are read from the existing _sheet.json cache and computed
in pure Python. Results are identical to what the Excel formulas produce.

Columns computed (rows 9-55):
  CY  — PMREDActivity    (red alert flag: 0/1)
  CZ  — Exptd Compl Days (AA * U)
  DA  — Act Compl Days   (Z * U)
  DB  — Exptd%           (date-progress 0.0-1.0)
  DC  — PLREDActivity    (orange alert flag: 0/1)
  DD  — Progress Flag    (department numeric code)
  DF  — BalEfforts       (Z * U, same as DA)
  DG  — AssignedEfforts  (U)
  DH  — AlertDt Yellow   (yellow alert flag: 0/1)

Header rows (non-task):
  DE1  = "PrjComl%"  (label, static)
  DE2  = LeadT_Setting AR10 value  → hardcoded: "Comm"
  DF2  = LeadT_Setting AR24 value  → hardcoded: "Comm"
  DE3  = IFERROR(DE7/DE6, 0)  → SW actual / SW total efforts
  DF3  = IFERROR(DF7/DF6, 0)  → EC actual / EC total efforts
  DE4  = "TEfforts"  (label, static)
  DF4  = "TEfforts"  (label, static)
  DE6  = SUMIF(G:G, "SW", DG:DG)  → sum of U for SW-owner rows
  DF6  = SUMIF(G:G, "EC", DG:DG)  → sum of U for EC-owner rows
  DE7  = SUMIF(G9:G55, "SW", DF9:DF55) → sum of DF (Z*U) for SW rows
  DF7  = SUMIF(G9:G55, "EC", DF9:DF55) → sum of DF (Z*U) for EC rows
  DF1  = IFERROR((DE7+DF7)/(DE6+DF6), 0)  → overall completion %

Thresholds (from @LeadT_Setting, hardcoded — they never change per project):
  AK6 = 0.8   (red/orange threshold for expected %)
  AI6 = 0.2   (red alert gap threshold)
  AJ6 = 0.1   (orange alert gap threshold)
"""

from datetime import date, datetime


# ── Threshold constants from @LeadT_Setting ───────────────────────────────────
_THRESHOLD_AK6 = 0.8   # expected% level that triggers alerts
_THRESHOLD_AI6 = 0.2   # red alert: actual lags expected by this much
_THRESHOLD_AJ6 = 0.1   # orange alert: actual lags expected by this much

# Department code map (DD column)
_DEPT_CODE = {
    "engineering":        1,
    "purchase":           2,
    "software":           3,
    "project management": 4,
    "manufacturing":      5,
    "sales":              6,
    "client":             7,
}


# ── Value parsers (matching what get_raw_sheet stores in cells dict) ──────────

def _pct(v) -> float:
    """Parse a stored percent value to float 0.0-1.0.
    Handles: "75%", "0%", "100%", 0.75, 1, None → float"""
    if v is None:
        return 0.0
    s = str(v).replace("%", "").strip()
    try:
        f = float(s)
        return f / 100.0 if f > 1.0 else f
    except (ValueError, TypeError):
        return 0.0


def _float(v) -> float:
    """Parse effort_days — stored as "5.0" string or numeric."""
    if v is None:
        return 0.0
    try:
        return float(str(v).strip())
    except (ValueError, TypeError):
        return 0.0


def _date(v) -> date | None:
    """Parse a stored date string "23-Feb-26" → date object, or None."""
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    if not s:
        return None
    for fmt in ("%d-%b-%y", "%Y-%m-%d", "%d-%b-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _str(v) -> str:
    """Raw string value, lowercased for comparisons."""
    if v is None:
        return ""
    return str(v).strip()


# ── Per-row formula implementations ───────────────────────────────────────────

def _calc_DB(U: float, V: date | None, W: date | None, today: date) -> float:
    """DB = expected progress % based on today's position between V and W."""
    if U <= 0 or V is None or today < V:
        return 0.0
    if today >= W:
        return 1.0
    try:
        return (today - V).days / (W - V).days
    except Exception:
        return 0.0


def _calc_CY(today: date, W: date | None, Z: float, DB: float, U: float) -> int:
    """CY — Red alert flag.
    Fires if:
      - Task is overdue (today > W) AND not complete AND expected=100%
      - OR expected% >= 80% but actual <= 80%
      - OR actual lags expected by >20%, expected <80%, task has effort
    """
    if W is None:
        return 0
    cond1 = today > W and Z != 1.0 and DB == 1.0
    cond2 = DB >= _THRESHOLD_AK6 and Z <= _THRESHOLD_AK6
    cond3 = Z < (DB - _THRESHOLD_AI6) and DB > Z and U > 0 and DB < _THRESHOLD_AK6
    return 1 if (cond1 or cond2 or cond3) else 0


def _calc_DC(today: date, W: date | None, Z: float, DB: float, U: float) -> int:
    """DC — Orange alert flag. Same logic as CY but gap threshold is 0.1 not 0.2."""
    if W is None:
        return 0
    cond1 = today > W and Z != 1.0 and DB == 1.0
    cond2 = DB >= _THRESHOLD_AK6 and Z <= _THRESHOLD_AK6
    cond3 = Z < (DB - _THRESHOLD_AJ6) and DB > Z and U > 0
    return 1 if (cond1 or cond2 or cond3) else 0


def _calc_DD(Z: float, AD: str) -> int:
    """DD — Department numeric code (0 if task complete)."""
    if Z >= 1.0:
        return 0
    return _DEPT_CODE.get(AD.lower(), 0)


def _calc_DH(today: date, U: float, AB: date | None, Z: float,
             V: date | None, Y: date | None) -> int:
    """DH — Yellow alert flag.
    Fires if:
      - Alert date (AB) has passed, task has effort, not 80% done
      - OR task started (V < today), has effort, not complete, no actual end logged
    """
    cond1 = (U > 0 and AB is not None and today >= AB
             and Z < 0.8 and AB is not None)
    cond2 = (Z != 1.0 and U > 0 and V is not None
             and V < today and Y is None)
    return 1 if (cond1 or cond2) else 0


# ── Main function ─────────────────────────────────────────────────────────────

def compute_sysmemory_json(project_id: str, sheet_data: dict,
                           sysmemory_cache: str) -> str | None:
    """
    Compute system memory values from _sheet.json cache and write to
    {sysmemory_cache}/{project_id}_sysmemory.json.

    Args:
        project_id:       e.g. "SWESch_FSL_2122_CHN_OR004_PLC"
        sheet_data:       the dict returned by _read_sheet_cache()
        sysmemory_cache:  full path to the system_memory cache dir

    Returns:
        path to the written JSON file, or None on failure.
    """
    import os
    import json

    cells = sheet_data.get("cells", {})
    today = date.today()

    def cv(col, row):
        """Raw cell value from cells dict."""
        info = cells.get(f"{col}{row}")
        if not info or info.get("skip"):
            return None
        return info.get("v")

    TASK_START = 9
    TASK_END   = 55

    rows = {}

    # ── Per-task rows (9-55) ──────────────────────────────────────────────────
    for r in range(TASK_START, TASK_END + 1):
        U  = _float(cv("U", r))
        V  = _date(cv("V", r))
        W  = _date(cv("W", r))
        Y  = _date(cv("Y", r))
        Z  = _pct(cv("Z", r))
        AA = _pct(cv("AA", r))
        AB = _date(cv("AB", r))
        AD = _str(cv("AD", r))

        DB = _calc_DB(U, V, W, today)

        rows[str(r)] = {
            "CY": _calc_CY(today, W, Z, DB, U),
            "CZ": round(AA * U, 4) if U else 0,
            "DA": round(Z * U, 4) if U else 0,
            "DB": round(DB, 6),
            "DC": _calc_DC(today, W, Z, DB, U),
            "DD": _calc_DD(Z, AD),
            "DF": round(Z * U, 4) if U else 0,   # same as DA
            "DG": U,
            "DH": _calc_DH(today, U, AB, Z, V, Y),
        }

    # ── Header / summary rows ─────────────────────────────────────────────────
    # DE6 = SUMIF owner=="SW" → DG (assigned efforts)
    # DF6 = SUMIF owner=="EC" → DG
    # DE7 = SUMIF owner=="SW" → DF (actual completed effort)
    # DF7 = SUMIF owner=="EC" → DF
    de6 = df6 = de7 = df7 = 0.0
    for r in range(TASK_START, TASK_END + 1):
        owner = _str(cv("G", r)).upper()
        row_r = rows.get(str(r), {})
        dg_val = row_r.get("DG", 0) or 0
        df_val = row_r.get("DF", 0) or 0
        if owner == "SW":
            de6 += dg_val
            de7 += df_val
        elif owner == "EC":
            df6 += dg_val
            df7 += df_val

    de3 = round(de7 / de6, 6) if de6 else 0
    df3 = round(df7 / df6, 6) if df6 else 0
    df1 = round((de7 + df7) / (de6 + df6), 6) if (de6 + df6) else 0

    sysmemory_data = {
        "project_id":   project_id,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "header": {
            "DF1":  df1,          # overall project completion %
            "DE2":  "Comm",       # from @LeadT_Setting!AR10 (constant)
            "DF2":  "Comm",       # from @LeadT_Setting!AR24 (constant)
            "DE3":  de3,          # SW actual/total efforts ratio
            "DF3":  df3,          # EC actual/total efforts ratio
            "DE4":  "TEfforts",
            "DF4":  "TEfforts",
            "DE6":  round(de6, 4),
            "DF6":  round(df6, 4),
            "DE7":  round(de7, 4),
            "DF7":  round(df7, 4),
        },
        "rows": rows,
    }

    os.makedirs(sysmemory_cache, exist_ok=True)
    sysmemory_path = os.path.join(sysmemory_cache, f"{project_id}_sysmemory.json")

    try:
        with open(sysmemory_path, "w") as f:
            json.dump(sysmemory_data, f, indent=2)
        print(f"[sysmemory] Written: {sysmemory_path}")
        return sysmemory_path
    except Exception as e:
        print(f"[sysmemory] Write failed for {project_id}: {e}")
        return None


# ── Drop-in for update_tasks_bulk integration ─────────────────────────────────

def queue_sysmemory_compute(project_id: str, sched_cache: str,
                             sysmemory_cache: str) -> None:
    """
    Regenerate sysmemory JSON in a background thread from the sheet cache.
    Call this after _update_sheet_cache() in update_tasks_bulk().

    Usage in update_tasks_bulk():
        from sysmemory_compute import queue_sysmemory_compute
        queue_sysmemory_compute(project_id, sched_cache, sysmemory_cache)
    """
    import threading
    from excel_db import _read_sheet_cache  # adjust import path as needed

    def _run():
        try:
            sheet = _read_sheet_cache(project_id, sched_cache)
            if sheet is None:
                print(f"[sysmemory] No sheet cache for {project_id}, skipping")
                return
            compute_sysmemory_json(project_id, sheet, sysmemory_cache)
        except Exception as e:
            print(f"[sysmemory] Background compute failed for {project_id}: {e}")

    t = threading.Thread(target=_run, daemon=True)
    t.start()
