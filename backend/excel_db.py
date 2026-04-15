import os
import openpyxl
from datetime import datetime, date
import threading
import queue
from openpyxl.utils import get_column_letter as gcl
import pycel
from pycel.excelcompiler import ExcelCompiler
from openpyxl.styles.numbers import is_date_format
from openpyxl.utils import column_index_from_string as col2idx

try:
    from pycel.excelcompiler import ExcelCompiler
    PY_CEL_AVAILABLE = True
    print("PyCel is available for formula evaluation")
except ImportError:
    PY_CEL_AVAILABLE = False
    print("PyCel not installed, falling back to openpyxl only")

# ===== TOP OF excel_db.py (after imports, before any functions) =====
def _patch_pycel_globals():
    """
    Inject COUNTA and PRODUCT into pycel.excellib BEFORE ExcelCompiler() runs.
    pycel's load_functions() uses getattr(module, name) across default_modules,
    and pycel.excellib is first in that list — so this is where to inject.
    Must be called at module load time, not after ExcelCompiler() is instantiated.
    """
    if not PY_CEL_AVAILABLE:
        return
    try:
        import pycel.excellib as excellib

        def counta_func(*args):
            count = 0
            for arg in args:
                if arg is None:
                    continue
                if isinstance(arg, (list, tuple)):
                    for item in arg:
                        if item is not None and str(item).strip() != "":
                            count += 1
                else:
                    if str(arg).strip() != "":
                        count += 1
            return count

        def product_func(*args):
            result = 1
            has_numbers = False
            for arg in args:
                if arg is None:
                    continue
                if isinstance(arg, (list, tuple)):
                    for item in arg:
                        if item is not None:
                            try:
                                result *= float(item)
                                has_numbers = True
                            except (ValueError, TypeError):
                                pass
                else:
                    try:
                        result *= float(arg)
                        has_numbers = True
                    except (ValueError, TypeError):
                        pass
            return result if has_numbers else 0

        setattr(excellib, 'COUNTA', counta_func)
        setattr(excellib, 'PRODUCT', product_func)
        setattr(excellib, 'counta', counta_func)
        setattr(excellib, 'product', product_func)

        # Verify it landed
        assert getattr(excellib, 'counta', None) is counta_func
        assert getattr(excellib, 'product', None) is product_func
        print("[PyCel] COUNTA and PRODUCT patched into pycel.excellib successfully")

    except Exception as e:
        print(f"[PyCel] Global patch failed: {e}")

# Run once at module load — BEFORE any ExcelCompiler() is ever called
_patch_pycel_globals()

def safe_load_workbook(filepath, data_only=True):
    """Safely load an Excel workbook with error handling for corrupted files"""
    import os
    from openpyxl import load_workbook
    
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"File not found: {filepath}")
    
    try:
        if os.path.getsize(filepath) == 0:
            raise ValueError(f"File is empty: {filepath}")
    except OSError:
        pass
    
    try:
        # Try normal load first
        return load_workbook(filepath, data_only=data_only)
    except EOFError:
        print(f"EOFError loading {filepath}, trying recovery mode...")
        try:
            # Try read-only mode
            wb = load_workbook(filepath, read_only=True, data_only=data_only)
            # If successful, close it and try normal load again
            wb.close()
            return load_workbook(filepath, data_only=data_only)
        except Exception as e:
            raise Exception(f"Cannot load corrupted file {filepath}: {e}")

# ── Paths ──────────────────────────────────────────────────────
BASE_DIR        = os.path.dirname(__file__)
USERS_PATH      = os.path.join(BASE_DIR, "users.xlsx")
DATA_DIR        = os.path.join(BASE_DIR, "..", "data")

# ── Role to discipline folder mapping ─────────────────────────
ROLE_TO_DEPT = {
    "sw_tl":    "SW",
    "hw_tl":    "HW",
    "mfg_tl":   "MFG",
    "pm":       "PM",
    "admin":    "SW",      
    "sw_head":  "SW",
    "hw_head":  "HW",
    "mfg_head": "MFG",
    "pm_head":  "PM",
}

# ── Department configuration ───────────────────────────────────
DEPT_CONFIG = {
    "SW": {
        "sched_prefix":   "SWESch",
        "monitor_prefix": "SW_Monitor",
        "sched_folder":   "SWESch",
        "monitor_sheet":  "SWMon",
        "project_id_col": "F",
        "file_col":       "B",
        "head_col":       "C",
        "tl_col":         "D",
        "timestamp_col":  "BA",
        "data_start_row": 12,
        "date_rows":      [28, 29, 30, 31, 32, 33],
        "extra_sections": [],
        "editable_cols": {"X", "Y", "Z", "AD", "AF"},
        "task_cols": {
            "actual_start": "X",
            "actual_end": "Y",
            "percent_complete": "Z",
            "help_required": "AD",
            "remark": "AF",
        },
    },
    "HW": {
        "sched_prefix":   "HWESch",
        "monitor_prefix": "HW_Monitor",
        "sched_folder":   "HWESch",
        "monitor_sheet":  "HWMon",
        "project_id_col": "F",
        "head_col":       "C",
        "tl_col":         "D",
        "timestamp_col":  "BA",
        "data_start_row": 12,
    },
    "MFG": {
        "sched_prefix":   "MFGSch",
        "monitor_prefix": "MFG_Monitor",
        "sched_folder":   "MFGSch",
        "monitor_sheet":  "MFGMon",
        "project_id_col": "F",
        "head_col":       "C",
        "tl_col":         "D",
        "timestamp_col":  "BA",
        "data_start_row": 12,
    },
    "PM": {
        "sched_prefix":   "PrjSch",
        "monitor_prefix": "PL_Monitor",
        "sched_folder":   "PrjSch",
        "monitor_sheet":  "PLMon",
        "project_id_col": "E",
        "file_col":       "B",
        "head_col":       "C",
        "tl_col":         "D",
        "timestamp_col":  "AZ",
        "data_start_row": 12,
        "date_rows":      list(range(23, 35)),
        "extra_sections": [
            {
                "title": "Panels & Value",
                "rows": [
                    {"label": "Balance Panels",  "col": "C", "row": 43},
                    {"label": "Panel Disp Act",  "col": "C", "row": 44},
                    {"label": "Estimated VA%",   "col": "C", "format": "percent", "row": 45},
                    {"label": "Estimated VA",    "col": "C", "row": 46},
                    {"label": "Estimated SM%",   "col": "C", "format": "percent", "row": 47},
                    {"label": "Estimated SM",    "col": "C", "row": 48},
                    {"label": "Actual VA%",      "col": "C", "format": "percent", "row": 49},
                    {"label": "Actual VA",       "col": "C", "row": 50},
                    {"label": "Actual SM%",      "col": "C", "format": "percent", "row": 51},
                    {"label": "Actual SM",       "col": "C", "row": 52},
                    {"label": "Reason / Remark", "col": "C", "row": 53},
                ],
            },
        ],
        "editable_cols": {"X", "Y", "Z", "AF"},  # AD is NOT editable for PM
        "task_cols": {
            "actual_start": "X",
            "actual_end": "Y",
            "percent_complete": "Z",
            "remark": "AF",
            # No help_required for PM
        },
        "scopeConfig": {
            "enabled": True,
            "cells": {
                "hw":   {"col": "AG", "row": 9},
                "sw":   {"col": "AH", "row": 9},
                "mfg":  {"col": "AI", "row": 9},
                "inst": {"col": "AJ", "row": 9},
                "com":  {"col": "AK", "row": 9},
            },
            "labels": {
                "hw":   "HW",
                "sw":   "SW",
                "mfg":  "MFG",
                "inst": "INST",
                "com":  "COMM",
            },
            "displayOrder": [
                ["hw", "sw", "mfg"],
                ["inst", "com"],
            ],
        },
    },
}

# ── Project Setup Cell Mapping ────────────────────────────────────
# Maps form fields to Excel cells for each department
# Used when a new project (without JSON cache and timestamp) is opened for setup

PROJECT_SETUP_CELL_MAP = {
    "PM": {
        # Header Fields (always editable - blue cells)
        "master_or":        {"sheet": "PrjSch", "cell": "D2", "type": "text"},
        "client_po":        {"sheet": "PrjSch", "cell": "D3", "type": "text"},
        "quote_number":     {"sheet": "PrjSch", "cell": "I1", "type": "text"},
        "sales_engineer":   {"sheet": "PrjSch", "cell": "I2", "type": "text"},
        "sales_manager":    {"sheet": "PrjSch", "cell": "I3", "type": "text"},
        
        # Project Details (blue cells)
        "customer_name":    {"sheet": "PrjSch", "cell": "D10", "type": "text"},
        "end_customer":     {"sheet": "PrjSch", "cell": "D11", "type": "text"},
        "consultant":       {"sheet": "PrjSch", "cell": "D12", "type": "text"},
        "project_desc":     {"sheet": "PrjSch", "cell": "D13", "type": "text"},
        "mfg_loc":          {"sheet": "PrjSch", "cell": "D15", "type": "dropdown", "options": ["GON", "DUB"]},
        
        # Efforts (blue cells)
        "hw_efforts":       {"sheet": "PrjSch", "cell": "D17", "type": "number"},
        "std_panels":       {"sheet": "PrjSch", "cell": "D18", "type": "number"},
        "act_panels":       {"sheet": "PrjSch", "cell": "D19", "type": "number"},
        "sw_efforts":       {"sheet": "PrjSch", "cell": "D20", "type": "number"},
        "mfg_efforts":      {"sheet": "PrjSch", "cell": "D21", "type": "number"},
        
        # Actual Efforts (blue cells - B17, B20, B21)
        "actual_hw_efforts":   {"sheet": "PrjSch", "cell": "B17", "type": "number"},
        "actual_sw_efforts":   {"sheet": "PrjSch", "cell": "B20", "type": "number"},
        "actual_mfg_efforts":  {"sheet": "PrjSch", "cell": "B21", "type": "number"},
        
        # Dates (Customer Dates - Column C, blue cells)
        "po_date":          {"sheet": "PrjSch", "cell": "C23", "type": "date"},
        "opf_recpt":        {"sheet": "PrjSch", "cell": "C24", "type": "date"},
        "hw_input":         {"sheet": "PrjSch", "cell": "C25", "type": "date"},
        "dwg_sub":          {"sheet": "PrjSch", "cell": "C26", "type": "date"},
        "dwg_appr":         {"sheet": "PrjSch", "cell": "C27", "type": "date"},
        "hw_fat":           {"sheet": "PrjSch", "cell": "C28", "type": "date"},
        "dispatch":         {"sheet": "PrjSch", "cell": "C29", "type": "date"},
        "sw_input":         {"sheet": "PrjSch", "cell": "C30", "type": "date"},
        "sw_fat":           {"sheet": "PrjSch", "cell": "C31", "type": "date"},
        "install":          {"sheet": "PrjSch", "cell": "C32", "type": "date"},
        "precomm":          {"sheet": "PrjSch", "cell": "C33", "type": "date"},
        "comm":             {"sheet": "PrjSch", "cell": "C34", "type": "date"},
        
        # Stakeholders (Column C, blue cells)
        "sh_sales":         {"sheet": "PrjSch", "cell": "C36", "type": "text"},
        "sh_hw":            {"sheet": "PrjSch", "cell": "C37", "type": "text"},
        "sh_sw":            {"sheet": "PrjSch", "cell": "C38", "type": "text"},
        "sh_byr":           {"sheet": "PrjSch", "cell": "C39", "type": "text"},
        "sh_mfg":           {"sheet": "PrjSch", "cell": "C40", "type": "text"},
        "sh_ec":            {"sheet": "PrjSch", "cell": "C41", "type": "text"},
        "sh_ac":            {"sheet": "PrjSch", "cell": "C42", "type": "text"},
        
        # Actuals (Column C - shifted from B, blue cells)
        "panel_disp_act":   {"sheet": "PrjSch", "cell": "C44", "type": "number"},
        "est_va_pct":       {"sheet": "PrjSch", "cell": "C45", "type": "number"},
        "est_va":           {"sheet": "PrjSch", "cell": "C46", "type": "number"},
        "est_sm_pct":       {"sheet": "PrjSch", "cell": "C47", "type": "number"},
        "est_sm":           {"sheet": "PrjSch", "cell": "C48", "type": "number"},
        "act_va_pct":       {"sheet": "PrjSch", "cell": "C49", "type": "number"},
        "act_va":           {"sheet": "PrjSch", "cell": "C50", "type": "number"},
        "act_sm_pct":       {"sheet": "PrjSch", "cell": "C51", "type": "number"},
        "act_sm":           {"sheet": "PrjSch", "cell": "C52", "type": "number"},
        "reason_remark":    {"sheet": "PrjSch", "cell": "C53", "type": "text"},
        
        # LD Details
        "ld_date":          {"sheet": "PrjSch", "cell": "AF1", "type": "date"},
        "ld_maxwk":         {"sheet": "PrjSch", "cell": "AF2", "type": "number"},
        "ld_maxov":         {"sheet": "PrjSch", "cell": "AF3", "type": "number"},
        "ld_remarks":       {"sheet": "PrjSch", "cell": "AF4", "type": "text"},
        
        # Warranty
        "warranty":         {"sheet": "PrjSch", "cell": "AF6", "type": "text"},
        
        # Scope (YES/NO dropdowns at row 9, columns AG-AK)
        "scope_hw":         {"sheet": "PrjSch", "cell": "AG9", "type": "dropdown", "options": ["YES", "NO"]},
        "scope_sw":         {"sheet": "PrjSch", "cell": "AH9", "type": "dropdown", "options": ["YES", "NO"]},
        "scope_mfg":        {"sheet": "PrjSch", "cell": "AI9", "type": "dropdown", "options": ["YES", "NO"]},
        "scope_inst":       {"sheet": "PrjSch", "cell": "AJ9", "type": "dropdown", "options": ["YES", "NO"]},
        "scope_com":        {"sheet": "PrjSch", "cell": "AK9", "type": "dropdown", "options": ["YES", "NO"]},
        
        # Editable Number Fields (Column E - blue cells, only write if changed)
        "internal_kom":        {"sheet": "PrjSch", "cell": "E10", "type": "number"},
        "engg_ip_collection":  {"sheet": "PrjSch", "cell": "E12", "type": "number"},
        "design_preparation":  {"sheet": "PrjSch", "cell": "E13", "type": "number"},
        "electrical_drawings": {"sheet": "PrjSch", "cell": "E17", "type": "number"},
        
        # Lead Time Fields (Column S3, S4)
        "critical_lead_time":  {"sheet": "PrjSch", "cell": "S3", "type": "number"},
        "normal_lead_time":    {"sheet": "PrjSch", "cell": "S4", "type": "number"},
        
        # Column S (Lead Time) - per row, with task name from Column I
        "s9":  {"sheet": "PrjSch", "cell": "S9",  "row": 9,  "task_row": 9,  "type": "number"},
        "s10": {"sheet": "PrjSch", "cell": "S10", "row": 10, "task_row": 10, "type": "number"},
        "s11": {"sheet": "PrjSch", "cell": "S11", "row": 11, "task_row": 11, "type": "number"},
        "s12": {"sheet": "PrjSch", "cell": "S12", "row": 12, "task_row": 12, "type": "number"},
        "s13": {"sheet": "PrjSch", "cell": "S13", "row": 13, "task_row": 13, "type": "number"},
        "s16": {"sheet": "PrjSch", "cell": "S16", "row": 16, "task_row": 16, "type": "number"},
        "s17": {"sheet": "PrjSch", "cell": "S17", "row": 17, "task_row": 17, "type": "number"},
        "s22": {"sheet": "PrjSch", "cell": "S22", "row": 22, "task_row": 22, "type": "number"},
        "s23": {"sheet": "PrjSch", "cell": "S23", "row": 23, "task_row": 23, "type": "number"},
        "s24": {"sheet": "PrjSch", "cell": "S24", "row": 24, "task_row": 24, "type": "number"},
        "s25": {"sheet": "PrjSch", "cell": "S25", "row": 25, "task_row": 25, "type": "number"},
        "s26": {"sheet": "PrjSch", "cell": "S26", "row": 26, "task_row": 26, "type": "number"},
        "s27": {"sheet": "PrjSch", "cell": "S27", "row": 27, "task_row": 27, "type": "number"},
        "s28": {"sheet": "PrjSch", "cell": "S28", "row": 28, "task_row": 28, "type": "number"},
        "s29": {"sheet": "PrjSch", "cell": "S29", "row": 29, "task_row": 29, "type": "number"},
        "s30": {"sheet": "PrjSch", "cell": "S30", "row": 30, "task_row": 30, "type": "number"},
        "s31": {"sheet": "PrjSch", "cell": "S31", "row": 31, "task_row": 31, "type": "number"},
        "s32": {"sheet": "PrjSch", "cell": "S32", "row": 32, "task_row": 32, "type": "number"},
        "s33": {"sheet": "PrjSch", "cell": "S33", "row": 33, "task_row": 33, "type": "number"},
        "s34": {"sheet": "PrjSch", "cell": "S34", "row": 34, "task_row": 34, "type": "number"},
        "s35": {"sheet": "PrjSch", "cell": "S35", "row": 35, "task_row": 35, "type": "number"},
        "s36": {"sheet": "PrjSch", "cell": "S36", "row": 36, "task_row": 36, "type": "number"},
        "s37": {"sheet": "PrjSch", "cell": "S37", "row": 37, "task_row": 37, "type": "number"},
        "s38": {"sheet": "PrjSch", "cell": "S38", "row": 38, "task_row": 38, "type": "number"},
        "s39": {"sheet": "PrjSch", "cell": "S39", "row": 39, "task_row": 39, "type": "number"},
        "s40": {"sheet": "PrjSch", "cell": "S40", "row": 40, "task_row": 40, "type": "number"},
        "s41": {"sheet": "PrjSch", "cell": "S41", "row": 41, "task_row": 41, "type": "number"},
        "s46": {"sheet": "PrjSch", "cell": "S46", "row": 46, "task_row": 46, "type": "number"},
        "s47": {"sheet": "PrjSch", "cell": "S47", "row": 47, "task_row": 47, "type": "number"},
        "s48": {"sheet": "PrjSch", "cell": "S48", "row": 48, "task_row": 48, "type": "number"},
        "s49": {"sheet": "PrjSch", "cell": "S49", "row": 49, "task_row": 49, "type": "number"},
        "s50": {"sheet": "PrjSch", "cell": "S50", "row": 50, "task_row": 50, "type": "number"},
        "s51": {"sheet": "PrjSch", "cell": "S51", "row": 51, "task_row": 51, "type": "number"},
        "s52": {"sheet": "PrjSch", "cell": "S52", "row": 52, "task_row": 52, "type": "number"},
        "s53": {"sheet": "PrjSch", "cell": "S53", "row": 53, "task_row": 53, "type": "number"},
        "s54": {"sheet": "PrjSch", "cell": "S54", "row": 54, "task_row": 54, "type": "number"},
        "s55": {"sheet": "PrjSch", "cell": "S55", "row": 55, "task_row": 55, "type": "number"},
        
        # Column U (Effort Days) - per row, with task name from Column I
        "u12": {"sheet": "PrjSch", "cell": "U12", "row": 12, "task_row": 12, "type": "number"},
        "u13": {"sheet": "PrjSch", "cell": "U13", "row": 13, "task_row": 13, "type": "number"},
        "u14": {"sheet": "PrjSch", "cell": "U14", "row": 14, "task_row": 14, "type": "number"},
        "u15": {"sheet": "PrjSch", "cell": "U15", "row": 15, "task_row": 15, "type": "number"},
        "u16": {"sheet": "PrjSch", "cell": "U16", "row": 16, "task_row": 16, "type": "number"},
        "u17": {"sheet": "PrjSch", "cell": "U17", "row": 17, "task_row": 17, "type": "number"},
        "u18": {"sheet": "PrjSch", "cell": "U18", "row": 18, "task_row": 18, "type": "number"},
        "u19": {"sheet": "PrjSch", "cell": "U19", "row": 19, "task_row": 19, "type": "number"},
        "u20": {"sheet": "PrjSch", "cell": "U20", "row": 20, "task_row": 20, "type": "number"},
        "u21": {"sheet": "PrjSch", "cell": "U21", "row": 21, "task_row": 21, "type": "number"},
        "u22": {"sheet": "PrjSch", "cell": "U22", "row": 22, "task_row": 22, "type": "number"},
        "u23": {"sheet": "PrjSch", "cell": "U23", "row": 23, "task_row": 23, "type": "number"},
        "u31": {"sheet": "PrjSch", "cell": "U31", "row": 31, "task_row": 31, "type": "number"},
        "u35": {"sheet": "PrjSch", "cell": "U35", "row": 35, "task_row": 35, "type": "number"},
        "u36": {"sheet": "PrjSch", "cell": "U36", "row": 36, "task_row": 36, "type": "number"},
        "u37": {"sheet": "PrjSch", "cell": "U37", "row": 37, "task_row": 37, "type": "number"},
        "u38": {"sheet": "PrjSch", "cell": "U38", "row": 38, "task_row": 38, "type": "number"},
        "u39": {"sheet": "PrjSch", "cell": "U39", "row": 39, "task_row": 39, "type": "number"},
        "u40": {"sheet": "PrjSch", "cell": "U40", "row": 40, "task_row": 40, "type": "number"},
        "u41": {"sheet": "PrjSch", "cell": "U41", "row": 41, "task_row": 41, "type": "number"},
        "u42": {"sheet": "PrjSch", "cell": "U42", "row": 42, "task_row": 42, "type": "number"},
        "u43": {"sheet": "PrjSch", "cell": "U43", "row": 43, "task_row": 43, "type": "number"},
        "u44": {"sheet": "PrjSch", "cell": "U44", "row": 44, "task_row": 44, "type": "number"},
        "u45": {"sheet": "PrjSch", "cell": "U45", "row": 45, "task_row": 45, "type": "number"},
        "u46": {"sheet": "PrjSch", "cell": "U46", "row": 46, "task_row": 46, "type": "number"},
        "u47": {"sheet": "PrjSch", "cell": "U47", "row": 47, "task_row": 47, "type": "number"},
        "u48": {"sheet": "PrjSch", "cell": "U48", "row": 48, "task_row": 48, "type": "number"},
        "u49": {"sheet": "PrjSch", "cell": "U49", "row": 49, "task_row": 49, "type": "number"},
        "u50": {"sheet": "PrjSch", "cell": "U50", "row": 50, "task_row": 50, "type": "number"},
        "u51": {"sheet": "PrjSch", "cell": "U51", "row": 51, "task_row": 51, "type": "number"},
        "u52": {"sheet": "PrjSch", "cell": "U52", "row": 52, "task_row": 52, "type": "number"},
        "u54": {"sheet": "PrjSch", "cell": "U54", "row": 54, "task_row": 54, "type": "number"},
        
        # Column AD (Payment %) - specific rows only
        "ad9":  {"sheet": "PrjSch", "cell": "AD9",  "row": 9,  "task_row": 9,  "type": "number"},
        "ad10": {"sheet": "PrjSch", "cell": "AD10", "row": 10, "task_row": 10, "type": "number"},
        "ad11": {"sheet": "PrjSch", "cell": "AD11", "row": 11, "task_row": 11, "type": "number"},
        "ad12": {"sheet": "PrjSch", "cell": "AD12", "row": 12, "task_row": 12, "type": "number"},
        "ad13": {"sheet": "PrjSch", "cell": "AD13", "row": 13, "task_row": 13, "type": "number"},
        "ad14": {"sheet": "PrjSch", "cell": "AD14", "row": 14, "task_row": 14, "type": "number"},
        "ad15": {"sheet": "PrjSch", "cell": "AD15", "row": 15, "task_row": 15, "type": "number"},
        "ad16": {"sheet": "PrjSch", "cell": "AD16", "row": 16, "task_row": 16, "type": "number"},
        "ad18": {"sheet": "PrjSch", "cell": "AD18", "row": 18, "task_row": 18, "type": "number"},
        "ad19": {"sheet": "PrjSch", "cell": "AD19", "row": 19, "task_row": 19, "type": "number"},
        "ad22": {"sheet": "PrjSch", "cell": "AD22", "row": 22, "task_row": 22, "type": "number"},
        "ad24": {"sheet": "PrjSch", "cell": "AD24", "row": 24, "task_row": 24, "type": "number"},
        "ad25": {"sheet": "PrjSch", "cell": "AD25", "row": 25, "task_row": 25, "type": "number"},
        "ad31": {"sheet": "PrjSch", "cell": "AD31", "row": 31, "task_row": 31, "type": "number"},
        "ad32": {"sheet": "PrjSch", "cell": "AD32", "row": 32, "task_row": 32, "type": "number"},
        "ad33": {"sheet": "PrjSch", "cell": "AD33", "row": 33, "task_row": 33, "type": "number"},
        "ad34": {"sheet": "PrjSch", "cell": "AD34", "row": 34, "task_row": 34, "type": "number"},
        "ad35": {"sheet": "PrjSch", "cell": "AD35", "row": 35, "task_row": 35, "type": "number"},
        "ad40": {"sheet": "PrjSch", "cell": "AD40", "row": 40, "task_row": 40, "type": "number"},
        "ad43": {"sheet": "PrjSch", "cell": "AD43", "row": 43, "task_row": 43, "type": "number"},
        "ad48": {"sheet": "PrjSch", "cell": "AD48", "row": 48, "task_row": 48, "type": "number"},
        "ad49": {"sheet": "PrjSch", "cell": "AD49", "row": 49, "task_row": 49, "type": "number"},
        "ad50": {"sheet": "PrjSch", "cell": "AD50", "row": 50, "task_row": 50, "type": "number"},
        "ad51": {"sheet": "PrjSch", "cell": "AD51", "row": 51, "task_row": 51, "type": "number"},
        "ad52": {"sheet": "PrjSch", "cell": "AD52", "row": 52, "task_row": 52, "type": "number"},
        "ad53": {"sheet": "PrjSch", "cell": "AD53", "row": 53, "task_row": 53, "type": "number"},
        "ad54": {"sheet": "PrjSch", "cell": "AD54", "row": 54, "task_row": 54, "type": "number"},
        "ad55": {"sheet": "PrjSch", "cell": "AD55", "row": 55, "task_row": 55, "type": "number"},
    }
}

def get_setup_fields(dept):
    """Return the field mapping for a department, or empty dict if not found"""
    return PROJECT_SETUP_CELL_MAP.get(dept, {})

def detect_dept_from_path(filepath):
    """Detect department by matching monitor_prefix against filepath."""
    for dept, cfg in DEPT_CONFIG.items():
        if cfg["monitor_prefix"] in filepath:
            return dept
    return "SW"  # safe fallback

def find_monitor_file(dept):
    """Scan data/ for first file starting with department's monitor prefix."""
    prefix = DEPT_CONFIG[dept]["monitor_prefix"]
    if not os.path.exists(DATA_DIR):
        return None
    for f in os.listdir(DATA_DIR):
        if f.startswith(prefix) and f.endswith((".xlsx", ".xlsb")):
            return os.path.join(DATA_DIR, f)
    return None

def find_schedule_folder(dept):
    """Return path to schedule folder for department."""
    folder_name = DEPT_CONFIG[dept]["sched_folder"]
    path = os.path.join(DATA_DIR, folder_name)
    os.makedirs(path, exist_ok=True)
    return path

def get_cache_path(dept, cache_type):
    """
    Return cache folder path.
    cache_type = 'Schedules' | 'Monitoring' | 'SystemMemory'
    """
    path = os.path.join(DATA_DIR, "cache", cache_type, dept)
    os.makedirs(path, exist_ok=True)
    return path

def get_discipline_dirs(role):
    """
    Return (projects_dir, monitor_dir, sched_cache, mon_cache) for role.
    monitor_dir is DATA_DIR (where monitor files live).
    """
    dept = ROLE_TO_DEPT.get(role, "SW")
    projects_dir = find_schedule_folder(dept)
    monitor_dir = DATA_DIR
    sched_cache = get_cache_path(dept, "Schedules")
    mon_cache = get_cache_path(dept, "Monitoring")
    return projects_dir, monitor_dir, sched_cache, mon_cache

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


# ── Task JSON sidecar cache ────────────────────────────────────
def _sheet_cache_path(project_id, sched_cache=None):
    if sched_cache is None:
        # Fallback - get default cache path for SW
        _, _, sched_cache, _ = get_discipline_dirs("sw_tl")
    return os.path.join(sched_cache, project_id + "_sheet.json")

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

def _update_sheet_cache(project_id, updates, sched_cache=None, role="sw_tl"):
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
                # Get department config for column mapping
                dept = ROLE_TO_DEPT.get(role, "SW")  # Need to pass role to this function
                cfg = DEPT_CONFIG.get(dept, DEPT_CONFIG["SW"])
                task_cols = cfg.get("task_cols", {})

                # Build field_to_col from task_cols
                field_to_col = {}
                for field, col in task_cols.items():
                    if field in ["actual_start", "actual_end", "percent_complete", "help_required", "remark"]:
                        field_to_col[field] = col
                
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
    if mon_cache is None:
        # Fallback - get default monitoring cache path for SW
        _, _, _, mon_cache = get_discipline_dirs("sw_tl")
    return os.path.join(mon_cache, f"Monitor_{short_name}.json")

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
def list_project_files(projects_dir=None):
    """Return list of .xlsx filenames in projects dir."""
    if projects_dir is None:
        # Fallback for calls without context
        projects_dir = find_schedule_folder("SW")
    if not os.path.exists(projects_dir):
        return []
    return [f for f in os.listdir(projects_dir)
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

    if role in ("admin", "head", "sw_head", "hw_head", "mfg_head", "pm_head"):
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
        sheet_data = get_raw_sheet(fpath, sched_cache=sched_cache, role=role)
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

def get_task_data_from_cache(project_id, role="sw_tl"):
    """
    Get task data (task name and completion %) from project JSON cache.
    Returns dict: {task_name: percent, ...}
    Stops when task name is empty.
    """
    _, _, sched_cache, _ = get_discipline_dirs(role)
    
    # Read sheet cache
    sheet = _read_sheet_cache(project_id, sched_cache)
    if not sheet:
        print(f"[Monitor] No sheet cache found for {project_id}")
        return {}
    
    cells = sheet.get("cells", {})
    tasks = {}
    
    for row in range(9, 56):  # rows 9-55
        # Get task name from column I
        task_name = _cell_val(cells, "I", row)
        
        # Stop if task name is empty (no more tasks)
        if not task_name or not str(task_name).strip():
            break
        
        # Get completion % from column Z
        percent_raw = _cell_val(cells, "Z", row)
        percent = 0
        if percent_raw is not None:
            s = str(percent_raw).replace("%", "").strip()
            try:
                p = float(s)
                percent = int(p * 100) if p <= 1.0 else int(p)
            except (ValueError, TypeError):
                percent = 0
        
        tasks[str(task_name).strip()] = percent
    
    return tasks

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
        sheet = get_raw_sheet(fpath, sched_cache=sched_cache, role=role)  # pass correct cache dir

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


def update_task(project_id, task_code, percent_complete, remark, role="sw_tl"):
    """Update % complete and remark for a specific task in the Excel file."""
    projects_dir, _, _, _ = get_discipline_dirs(role)
    fname = project_id + ".xlsx"
    fpath = os.path.join(projects_dir, fname)
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
    - Updates monitor task percentages
    - Regenerates system memory JSON in background
    """
    projects_dir, _, sched_cache, _ = get_discipline_dirs(role)

    # Step 1: Update JSON cache immediately (fast) — returns stamped timestamp
    now_str = _update_sheet_cache(project_id, updates, sched_cache, role=role)

    # Step 2: Prepare monitor updates
    monitor_updates = {}
    
    # 2a: Update timestamp
    if now_str:
        timestamp_updates = update_monitor_timestamp(project_id, now_str, role)
        monitor_updates.update(timestamp_updates)
    
    # 2b: Update task percentages (always, not just when tasks changed)
    task_updates = update_monitor_task_percentages(project_id, role)
    monitor_updates.update(task_updates)
    
    # 2c: Queue monitor Excel write if there are updates
    if monitor_updates:
        
        department = ROLE_TO_DEPT.get(role, "SW")
        queue_monitor_excel_write(department, monitor_updates)

    # Step 3: Queue Excel write to background thread
    queue_excel_write(project_id, updates, role)

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
    user = get_user_by_username(username)
    if not user:
        return []

    short = user.get("short_name", "").strip().upper()
    role = user.get("role", "sw_tl")
    
    # Get department from role
    department = ROLE_TO_DEPT.get(role, "SW")
    cfg = DEPT_CONFIG.get(department)
    
    # Determine if this role is a head or a tl
    is_head_role = role.endswith("_head") or role in ("admin", "head")
    
    if is_head_role:
        return get_all_monitor_projects(role)
    
    # For TL roles, filter by tl_col (which is "swe_name" in the returned dict)
    all_projects = get_all_monitor_projects(role)
    # ADD THIS:
    print(f"[DEBUG] tl_col = {cfg['tl_col']}, short = {short}")
    for p in all_projects[:5]:  # print first 5 to see what's in them
        print(f"[DEBUG] project entry: {p}")
    filtered = [m for m in all_projects if m.get(cfg["tl_col"]) == short]
    
    return filtered

def get_all_monitor_projects(role="sw_tl"):
    """Get ALL projects from the department monitoring file (using JSON cache)"""
    from datetime import date, datetime
    import json
    import os

    department = ROLE_TO_DEPT.get(role, "SW")
    cfg = DEPT_CONFIG[department]
    pid_col = cfg["project_id_col"]

    monitor_cache_path = os.path.join(get_cache_path(department, "Monitoring"), f"{department}_Monitor.json")

    if os.path.exists(monitor_cache_path):
        try:
            with open(monitor_cache_path, 'r') as f:
                monitor_data = json.load(f)

            cells = monitor_data.get('cells', {})
            results = []
            today = date.today()

            for coord, cell_info in cells.items():
                if coord.startswith(pid_col) and coord[len(pid_col):].isdigit():
                    row_num = int(coord[len(pid_col):])
                    project_id = str(cell_info.get('v')).strip()

                    if not project_id.startswith("FSL"):
                        continue

                    ts_col = cfg["timestamp_col"]
                    timestamp_str = cells.get(f"{ts_col}{row_num}", {}).get('v')

                    stale = False
                    if timestamp_str:
                        try:
                            last_edited = datetime.strptime(timestamp_str, "%d-%m-%Y %H:%M:%S").date()
                            stale = (today - last_edited).days > 2
                        except Exception as e:
                            print(f"[Monitor] Could not parse timestamp for {project_id}: {timestamp_str}, error={e}")

                    normalized_id = cfg["sched_prefix"] + "_" + project_id.replace("/", "_")
                    sched_folder = find_schedule_folder(department)
                    file_exists = (os.path.exists(os.path.join(sched_folder, normalized_id + ".xlsx")) or
                                   os.path.exists(os.path.join(sched_folder, normalized_id + ".xlsb")))

                    head_col = cfg["head_col"]
                    tl_col = cfg["tl_col"]
                    swh_head = cells.get(f"{head_col}{row_num}", {}).get('v', '')
                    swe_name = cells.get(f"{tl_col}{row_num}", {}).get('v', '')

                    # Inside the JSON cache block, after getting swe_name, add:
                    da_value = cells.get(f"DA{row_num}", {}).get('v', '')

                    # Inside the loop where results are appended, after calculating file_exists

                    results.append({
                        "project_id": project_id,
                        "file_id": normalized_id,
                        "stale": stale,
                        "file_exists": file_exists,
                        "da_status": da_value,  
                        cfg["head_col"]: (swh_head or "").strip().upper(),  # "C" for SW, "C" for PM
                        cfg["tl_col"]: (swe_name or "").strip().upper(),    # "D" for SW, "D" for PM
                    })

            print(f"[Monitor] Loaded {len(results)} projects from JSON cache")
            return results

        except Exception as e:
            print(f"[Monitor] Error reading from JSON cache: {e}, falling back to Excel")

    monitor_path = find_monitor_file(department)
    if not monitor_path:
        print(f"[Monitor] No monitor file found for {department}")
        return []

    if monitor_path.endswith(".xlsb"):
        return _read_master_xlsb(monitor_path, department)
    else:
        return _read_master_xlsx(monitor_path, department)

def _master_row_to_entry(pid, ba_val, ba_rgb=None, swh_head=None, swe_name=None, projects_dir=None,dept="SW"):
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
    prefix = DEPT_CONFIG[dept]["sched_prefix"]
    normalized_id = prefix + "_" + pid.replace("/", "_")
    folder = projects_dir if projects_dir else find_schedule_folder("SW")
    file_exists = (os.path.exists(os.path.join(folder, normalized_id + ".xlsx")) or
                   os.path.exists(os.path.join(folder, normalized_id + ".xlsb")))

    return {
        "project_id":    pid,            # display value e.g. FSL/2122/CHN/OR004_PLC
        "file_id":       normalized_id,  # actual file id e.g. SWESch_FSL_2122_CHN_OR004_PLC
        "stale":         stale,
        "file_exists":   file_exists,
        "da_status": "",
        "swh_head":      (swh_head or "").strip().upper(),
        "swe_name":      (swe_name or "").strip().upper(),
    }


def _read_master_xlsb(path, department=None):
    if department is None:
        department = detect_dept_from_path(path)
    cfg = DEPT_CONFIG[department]
    sheet_name = cfg["monitor_sheet"]
    pid_col_idx = col2idx(cfg["project_id_col"]) - 1
    head_idx = col2idx(cfg["head_col"]) - 1
    tl_idx = col2idx(cfg["tl_col"]) - 1
    ts_idx = col2idx(cfg["timestamp_col"]) - 1

    try:
        from pyxlsb import open_workbook
    except ImportError:
        print("pyxlsb not installed — cannot read .xlsb master file")
        return []

    projects_dir = find_schedule_folder(department)
    results = []
    try:
        with open_workbook(path) as wb:
            if sheet_name not in wb.sheets:
                print(f"[Monitor] Sheet '{sheet_name}' not found. Available: {wb.sheets}")
                return []
            with wb.get_sheet(sheet_name) as ws:
                for i, row in enumerate(ws.rows()):
                    if i < cfg["data_start_row"] - 1:
                        continue
                    pid = row[pid_col_idx].v if len(row) > pid_col_idx else None
                    ba_val = row[ts_idx].v if len(row) > ts_idx else None
                    swh_head = row[head_idx].v if len(row) > head_idx else None
                    swe_name = row[tl_idx].v if len(row) > tl_idx else None
                    entry = _master_row_to_entry(pid, ba_val, swh_head=swh_head, swe_name=swe_name, projects_dir=projects_dir, dept=department)
                    if entry:
                        results.append(entry)
    except Exception as e:
        print(f"Master .xlsb read error for {path}: {e}")
    return results

def _read_master_xlsx(path, department=None):
    if department is None:
        department = detect_dept_from_path(path)
    cfg = DEPT_CONFIG[department]
    sheet_name = cfg["monitor_sheet"]
    pid_col_idx = col2idx(cfg["project_id_col"]) - 1
    head_idx = col2idx(cfg["head_col"]) - 1
    tl_idx = col2idx(cfg["tl_col"]) - 1
    ts_idx = col2idx(cfg["timestamp_col"]) - 1

    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        if sheet_name not in wb.sheetnames:
            print(f"[Monitor] Sheet '{sheet_name}' not found in {path}. Available: {wb.sheetnames}")
            return []
        ws = wb[sheet_name]
    except Exception as e:
        print(f"Master .xlsx read error for {path}: {e}")
        return []

    projects_dir = find_schedule_folder(department)
    results = []
    for row in ws.iter_rows(min_row=cfg["data_start_row"], values_only=False):
        pid_cell = row[pid_col_idx] if len(row) > pid_col_idx else None
        ba_cell = row[ts_idx] if len(row) > ts_idx else None
        c_cell = row[head_idx] if len(row) > head_idx else None
        d_cell = row[tl_idx] if len(row) > tl_idx else None

        pid = pid_cell.value if pid_cell else None
        ba_val = ba_cell.value if ba_cell else None
        swh_head = c_cell.value if c_cell else None
        swe_name = d_cell.value if d_cell else None

        ba_rgb = None
        if ba_cell:
            try:
                fg = ba_cell.fill.fgColor
                if fg and fg.type == "rgb" and fg.rgb not in ("00000000", "FFFFFFFF"):
                    ba_rgb = fg.rgb
            except Exception:
                pass

        entry = _master_row_to_entry(pid, ba_val, ba_rgb, swh_head=swh_head, swe_name=swe_name, projects_dir=projects_dir, dept=department)
        if entry:
            results.append(entry)
    return results

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


def get_raw_sheet(filepath, max_col=32, sched_cache=None, role=None):
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

    # Determine department and get config
    if role:
        dept = ROLE_TO_DEPT.get(role, "SW")
    else:
        dept = detect_dept_from_path(filepath)
    cfg = DEPT_CONFIG.get(dept, DEPT_CONFIG["SW"])
    date_rows = cfg.get("date_rows", [28, 29, 30, 31, 32, 33])  # default to SW rows

    max_row = ws.max_row
    # Expand max_col if this dept has scope cells beyond the default cap
    scope_config = cfg.get("scopeConfig")
    if scope_config and scope_config.get("enabled"):
        from openpyxl.utils import column_index_from_string
        for cell_info in scope_config.get("cells", {}).values():
            col_letter = cell_info.get("col", "")
            if col_letter:
                col_idx = column_index_from_string(col_letter)
                max_col = max(max_col, col_idx)

    max_col = min(ws.max_column, max_col)
    cols = [gcl(i) for i in range(1, max_col + 1) if gcl(i) not in ("A","B","C","D")]

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
    EDITABLE_COLS = cfg.get("editable_cols", {"X", "Y", "Z", "AD", "AF"})
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
    # Build dates array dynamically from DEPT_CONFIG
    dates_list = []
    for row_num in date_rows:
        dates_list.append({
            "label": _v(f"B{row_num}") or f"B{row_num}",
            "pm": _date(f"C{row_num}"),
            "swe": _date(f"D{row_num}"),
            "pm_fill": _resolve_color(ws[f"C{row_num}"].fill.fgColor) if ws[f"C{row_num}"].fill.patternType not in (None, "none") else None,
            "pm_font": _resolve_color(ws[f"C{row_num}"].font.color),
            "swe_fill": _resolve_color(ws[f"D{row_num}"].fill.fgColor) if ws[f"D{row_num}"].fill.patternType not in (None, "none") else None,
            "swe_font": _resolve_color(ws[f"D{row_num}"].font.color),
        })

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
        "dates": dates_list,

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
    # Build extra_sections from config
    extra_sections_list = []
    for section in cfg.get("extra_sections", []):
        section_data = {
            "title": section.get("title", ""),
            "rows": []
        }
        for row_cfg in section.get("rows", []):
            row_num = row_cfg.get("row")
            col = row_cfg.get("col")
            
            # Read value from column C (where actual data is)
            value_coord = f"C{row_num}"
            value = cells.get(value_coord, {}).get("v") if value_coord in cells else None
            
            # Read label from column B
            label_coord = f"B{row_num}"
            label_from_excel = cells.get(label_coord, {}).get("v") if label_coord in cells else None
            
            row_data = {
                "label": label_from_excel or row_cfg.get("label", ""),
                "value": str(value) if value is not None else "",
            }
            
            # Handle format
            if row_cfg.get("format") == "percent" and value is not None:
                try:
                    num = float(value)
                    if 0 <= num <= 1:
                        row_data["value"] = f"{int(num * 100)}%"
                except (ValueError, TypeError):
                    pass
            
            # Handle actual_col for estimated vs actual
            if "actual_col" in row_cfg:
                actual_coord = f"{row_cfg['actual_col']}{row_num}"
                actual_value = cells.get(actual_coord, {}).get("v") if actual_coord in cells else None
                row_data["actual"] = str(actual_value) if actual_value is not None else ""
                if row_cfg.get("format") == "percent" and actual_value is not None:
                    try:
                        num = float(actual_value)
                        if 0 <= num <= 1:
                            row_data["actual"] = f"{int(num * 100)}%"
                    except (ValueError, TypeError):
                        pass
            
            section_data["rows"].append(row_data)
        
        if section_data["rows"]:
            extra_sections_list.append(section_data)

    left_panel["extra_sections"] = extra_sections_list

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

    # ── Scope configuration (if defined for this department) ────
    scope_data = None
    scope_config = cfg.get("scopeConfig")
    if scope_config and scope_config.get("enabled"):
        scope_data = {
            "enabled": True,
            "cells": {},
            "labels": scope_config.get("labels", {}),
            "displayOrder": scope_config.get("displayOrder", []),
        }
        for key, cell_info in scope_config.get("cells", {}).items():
            col = cell_info.get("col")
            row = cell_info.get("row")
            if col and row:
                coord = f"{col}{row}"
                # Get cell value from cells dict (already processed earlier)
                cell_val = cells.get(coord, {}).get("v") if coord in cells else None
                # Check if value is "YES" (case-insensitive)
                is_checked = False
                if cell_val is not None:
                    val_upper = str(cell_val).strip().upper()
                    is_checked = val_upper in ("YES", "Y", "TRUE", "1")
                scope_data["cells"][key] = {
                    "coord": coord,
                    "value": cell_val,
                    "checked": is_checked,
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
        "last_modified":  last_modified_str,
        "scope":          scope_data,
    }

        # Write JSON sidecar cache for fast future loads
    _write_sheet_cache(project_id, result, sched_cache)
    
    return result

def get_monitor_sheet_data(filepath, monitor_type="SW"):
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter as gcl
    from datetime import datetime, date, time
    import json
    import os

    # Detect department from filepath using config
    dept = detect_dept_from_path(filepath)
    cfg = DEPT_CONFIG[dept]

    cache_dir = get_cache_path(dept, "Monitoring")
    cache_path = os.path.join(cache_dir, f"{dept}_Monitor.json")

    if os.path.exists(cache_path):
        try:
            with open(cache_path, 'r') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error reading monitor cache: {e}")

    try:
        wb = safe_load_workbook(filepath, data_only=True)
    except Exception as e:
        print(f"Error loading monitor file {filepath}: {e}")
        return None

    ws = wb.active
    max_col = ws.max_column
    pid_col = cfg["project_id_col"]
    data_start = cfg["data_start_row"]

    # Find true last row with data
    max_row = data_start
    for row in range(data_start, ws.max_row + 1):
        cell_value = ws[f"{pid_col}{row}"].value
        if cell_value is None or str(cell_value).strip() == "" or str(cell_value).strip() == "_":
            break
        max_row = row

    cols = [gcl(i) for i in range(1, max_col + 1)]

    cells = {}
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            v = cell.value
            if v is None:
                continue
            if isinstance(v, datetime):
                v = v.strftime("%d-%b-%y")
            elif isinstance(v, date):
                v = v.strftime("%d-%b-%y")
            elif isinstance(v, time):
                v = v.strftime("%H:%M")
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

    with open(cache_path, 'w') as f:
        json.dump(result, f)

    wb.close()
    return result   

def generate_sysmemory_json(filepath, project_id, sched_cache=None):
    """
    Generate system memory JSON.
    - Rows 1-7: DE and DF columns only (skip DE2)
    - Row 8: CY to DJ headers + I header
    - Rows 9+: Read task name from I using openpyxl (no PyCel)
    - Stops after first empty task name
    """
    import os
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter as gcl
    
    # Determine discipline from filepath
    if "SWESch" in filepath:
        discipline = "SW"
    elif "HWESch" in filepath:
        discipline = "HW"
    elif "MFGSch" in filepath:
        discipline = "MFG"
    elif "PrjSch" in filepath:
        discipline = "PM"
    else:
        discipline = "SW"

    # Use central cache
    sysmemory_cache = os.path.join(DATA_DIR, "cache", "SystemMemory", discipline)
    os.makedirs(sysmemory_cache, exist_ok=True)
    sysmemory_path = os.path.join(sysmemory_cache, f"{project_id}_sysmemory.json")

    # Load workbook with data_only=False to keep formulas for CY-DJ
    wb = load_workbook(filepath, data_only=False)
    ws = wb.active
    sheet_name = ws.title
    
    # Columns from CY to DJ
    start_col_idx = 103  # CY
    end_col_idx = 114    # DJ
    sysmemory_cols = [gcl(i) for i in range(start_col_idx, end_col_idx + 1)]
    
    # Header columns for rows 1-7 (DE and DF)
    header_cols_1_7 = ["DE", "DF"]
    
    # Initialize PyCel for CY-DJ formula evaluation only
    excel_compiler = None
    if PY_CEL_AVAILABLE:
        try:
            excel_compiler = ExcelCompiler(filepath)
            print(f"PyCel initialized for system memory generation: {project_id}")
        except Exception as e:
            print(f"PyCel initialization failed: {e}")
    
    # Prepare system memory data
    sysmemory_data = {
        "project_id": project_id,
        "columns": sysmemory_cols,
        "rows": {},
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    
    # Step 1: Read rows 1-7 (only DE and DF columns)
    for row in range(1, 8):
        row_data = {}
        for col in header_cols_1_7:
            if col == "DE" and row == 2:
                continue
            
            coord = f"{col}{row}"
            cell = ws[coord]
            value = None
            
            if excel_compiler and cell.data_type == 'f':
                try:
                    cell_ref = f"{sheet_name}!{coord}"
                    value = excel_compiler.evaluate(cell_ref)
                except Exception:
                    value = cell.value
            else:
                value = cell.value
            
            if isinstance(value, (datetime, date)):
                if isinstance(value, datetime):
                    value = value.date()
                value = value.strftime("%Y-%m-%d")
            elif value is None:
                value = None
            
            row_data[col] = value
        
        sysmemory_data["rows"][str(row)] = row_data
    
    # Step 2: Read row 8 headers
    row8_data = {}
    for col in sysmemory_cols:
        coord = f"{col}8"
        cell = ws[coord]
        row8_data[col] = cell.value if cell.value else col
    
    # Read column I header directly from openpyxl
    i_cell_8 = ws["I8"]
    row8_data["I"] = i_cell_8.value if i_cell_8.value else "I"
    
    sysmemory_data["rows"]["8"] = row8_data
    
    # Step 3: Read rows 9 onwards
    for row in range(9, 56):
        row_data = {}
        
        # Read task name from column I - DIRECTLY from openpyxl, NO PyCel
        i_cell = ws[f"I{row}"]
        task_name = i_cell.value
        
        # Convert to string and clean
        if task_name is not None:
            task_name = str(task_name).strip()
        else:
            task_name = ""
        
        
        # Store task name if not empty
        if task_name:
            row_data["I"] = task_name
        
        # Read CY to DJ columns (use PyCel for formulas)
        for col in sysmemory_cols:
            coord = f"{col}{row}"
            cell = ws[coord]
            value = None
            
            if excel_compiler and cell.data_type == 'f':
                try:
                    cell_ref = f"{sheet_name}!{coord}"
                    value = excel_compiler.evaluate(cell_ref)
                except Exception:
                    value = cell.value
            else:
                value = cell.value
            
            if isinstance(value, (datetime, date)):
                if isinstance(value, datetime):
                    value = value.date()
                value = value.strftime("%Y-%m-%d")
            elif value is None:
                value = None
            
            row_data[col] = value
        
        sysmemory_data["rows"][str(row)] = row_data
        
        # Stop if no task name
        if not task_name:
            break
    
    wb.close()
    
    # Write to JSON
    try:
        with open(sysmemory_path, 'w') as f:
            json.dump(sysmemory_data, f, indent=2)
        print(f"System memory JSON saved: {sysmemory_path}")
        return sysmemory_path
    except Exception as e:
        print(f"Failed to save system memory JSON: {e}")
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

# Dedicated queue + worker for monitor Excel writes — prevents race conditions
# when multiple saves happen close together (each used to spin a new thread,
# causing one thread to open the file while another was mid-write → EOFError)
_monitor_write_queue = queue.Queue()
_monitor_thread_started = False

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

def _read_sysmemory_from_excel(project_id, fpath, role):
    """
    Evaluate sysmemory columns from the Excel file using Pycel (formula-aware)
    and write results to the sysmemory JSON cache.

    Called from _do_excel_write after the Excel save is fully complete so that
    Pycel always sees the latest written values.

    Column scope:
      - Rows 1–7  : DF, DG only (summary/header values)
      - Rows 8–55 : CY, CZ, DA, DB, DC, DD, DF, DG, DH (all computed task cols)

    Sheet name:
      - SW discipline → always "PrjSch"
      - Other disciplines (HW, MFG, PM) → detected from wb.active at runtime
    """
    import warnings
    warnings.filterwarnings("ignore")

    try:
        from pycel.excelcompiler import ExcelCompiler
    except ImportError:
        print("[Sysmemory] Pycel not installed — skipping sysmemory compute.")
        return

    # Determine which sheet to evaluate against
    discipline = ROLE_TO_DEPT.get(role, "SW")
    if discipline == "SW":
        sheet_name = "PrjSch"
    else:
        # Detect active sheet name dynamically for other departments
        try:
            from openpyxl import load_workbook as _lw
            _wb = _lw(fpath, read_only=True)
            sheet_name = _wb.active.title
            _wb.close()
        except Exception:
            sheet_name = "PrjSch"  # safe fallback

    # Rows 1-7: only DE and DF (DE2 is explicitly skipped)
    HEADER_COLS = ["DE", "DF"]

    # Rows 8-55: task cols — no DE, DG used instead
    TASK_COLS = ["CY", "CZ", "DA", "DB", "DC", "DD", "DF", "DG", "DH", "DI", "DJ"]

    # Percentage cells: raw 0.0-1.0 float → converted to 0-100
    PCT_CELLS = {
        ("DF", 1),   # DF1 — project completion %
        ("DE", 3),   # DE3 — SW completion %
        ("DF", 3),   # DF3 — Mfg completion %
    }
    # DB column (Exptd%) in task rows is also 0-1 float — handled inline

    print(f"[Sysmemory] Compiling {fpath} with Pycel (sheet: {sheet_name})...")
    try:
        excel = ExcelCompiler(filename=fpath)
    except Exception as e:
        print(f"[Sysmemory] Pycel compile failed for {project_id}: {e}")
        return

    # Excel date serial range — Pycel sometimes returns integers instead of
    # datetime objects when evaluating date formulas. We detect and convert them.
    _EXCEL_DATE_MIN = 30000   # ~1982
    _EXCEL_DATE_MAX = 100000  # ~2173

    def _excel_serial_to_str(v):
        """Convert an Excel date serial integer/float to a date string."""
        try:
            from openpyxl.utils.datetime import from_excel
            dt = from_excel(int(v))
            return dt.strftime("%Y-%m-%d")
        except Exception:
            return None

    def _evaluate(col, row):
        """Evaluate a single cell; return None on any error.
        DE2 is always skipped — do not evaluate or process it."""
        if col == "DE" and row == 2:
            return None
        try:
            v = excel.evaluate(f"{sheet_name}!{col}{row}")
            # Handle datetime/date objects
            if isinstance(v, datetime):
                return v.strftime("%Y-%m-%d %H:%M:%S")
            if isinstance(v, date):
                return v.strftime("%Y-%m-%d")
            if isinstance(v, (int, float)):
                # Detect Excel date serial numbers returned by Pycel
                if _EXCEL_DATE_MIN <= v <= _EXCEL_DATE_MAX:
                    converted = _excel_serial_to_str(v)
                    if converted:
                        return converted
                # Percentage conversion for known cells
                if (col, row) in PCT_CELLS or col == "DB":
                    return round(float(v) * 100, 2)
            return v
        except Exception:
            return None

    rows = {}

    # Rows 1-7: only DE and DF; DE2 is silently skipped inside _evaluate
    for row in range(1, 8):
        row_data = {col: _evaluate(col, row) for col in HEADER_COLS}
        rows[str(row)] = row_data

    # Row 8: label/header row
    rows["8"] = {col: _evaluate(col, 8) for col in TASK_COLS}

    # Rows 9-55: task rows — no DE
    for row in range(9, 56):
        row_data = {col: _evaluate(col, row) for col in TASK_COLS}
        rows[str(row)] = row_data

        # Persist to JSON cache
    sysmemory_cache = os.path.join(DATA_DIR, "cache", "SystemMemory", discipline)
    os.makedirs(sysmemory_cache, exist_ok=True)
    sysmemory_path = os.path.join(sysmemory_cache, f"{project_id}_sysmemory.json")

    sysmemory_data = {
        "project_id":   project_id,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "sheet":        sheet_name,
        "header_cols":  HEADER_COLS,
        "task_cols":    TASK_COLS,
        "rows":         rows,
    }

    with open(sysmemory_path, "w") as f:
        json.dump(sysmemory_data, f, indent=2)

    print(f"[Sysmemory] Written: {sysmemory_path}")

def update_monitor_cell(department, coord, value, role="sw_tl"):
    """Update a single cell in monitor JSON cache and queue Excel write"""
    import json
    import os
    from datetime import datetime
    
    monitor_cache_path = os.path.join(get_cache_path(department, "Monitoring"), f"{department}_Monitor.json")
    
    if not os.path.exists(monitor_cache_path):
        print(f"[Monitor] Cache not found: {monitor_cache_path}")
        return False
    
    try:
        # Load monitor cache
        with open(monitor_cache_path, 'r') as f:
            monitor_data = json.load(f)
        
        cells = monitor_data.get('cells', {})
        
        # Update the cell
        if coord not in cells:
            cells[coord] = {}
        
        cells[coord]['v'] = value
        cells[coord]['updated'] = True
        
        # NO timestamp update for monitor edits
        
        # Write back to JSON cache
        with open(monitor_cache_path, 'w') as f:
            json.dump(monitor_data, f, indent=2)
        
        # Queue Excel write
        updates = {coord: value}
        queue_monitor_excel_write(department, updates)
        
        print(f"[Monitor] Updated {coord} = {value}")
        return True
        
    except Exception as e:
        print(f"[Monitor] Failed to update cell: {e}")
        import traceback
        traceback.print_exc()
        return False

def read_sysmemory_json(project_id, role="sw_tl"):
    """
    Read sysmemory JSON for a project and extract task data.
    Hardcoded column mapping for speed and reliability.
    Returns dict: { task_name: { "PLRedActivity": 0/1, "PMRedActivity": 0/1, 
                                  "AlertDtYellow": 0/1, "ProgressFlag": 0-7 } }
    """
    import json
    import os
    
    department = ROLE_TO_DEPT.get(role, "SW")
    
    # Hardcoded column letters
    TASK_NAME_COL = "I"
    PL_RED_COL = "DC"
    PM_RED_COL = "CY"
    ALERT_DT_COL = "DH"
    PROGRESS_FLAG_COL = "DD"
       
    # Path to sysmemory JSON
    sysmemory_path = os.path.join(DATA_DIR, "cache", "SystemMemory", department, f"{project_id}_sysmemory.json")
    
    if not os.path.exists(sysmemory_path):
        print(f"[Sysmemory] File not found: {sysmemory_path}")
        return {}
    
    try:
        with open(sysmemory_path, 'r') as f:
            data = json.load(f)
        
        rows = data.get("rows", {})
        task_map = {}
        
        # Loop through task rows (9 to 55)
        for row_num in range(9, 56):
            row_key = str(row_num)
            if row_key not in rows:
                continue
            
            row_data = rows[row_key]
            
            # Get task name from hardcoded column I
            task_name = row_data.get(TASK_NAME_COL, "")
            
            if not task_name or not str(task_name).strip():
                continue
            
            task_name = str(task_name).strip()
            
            # Get values from hardcoded columns
            pl_red = row_data.get(PL_RED_COL, 0)
            pm_red = row_data.get(PM_RED_COL, 0)
            alert_dt = row_data.get(ALERT_DT_COL, 0)
            progress_flag = row_data.get(PROGRESS_FLAG_COL, 0)
            
            # Convert to int
            try:
                pl_red = int(pl_red) if pl_red else 0
                pm_red = int(pm_red) if pm_red else 0
                alert_dt = int(alert_dt) if alert_dt else 0
                progress_flag = int(progress_flag) if progress_flag else 0
            except (ValueError, TypeError):
                pl_red = 0
                pm_red = 0
                alert_dt = 0
                progress_flag = 0
            
            task_map[task_name] = {
                "PLRedActivity": pl_red,
                "PMRedActivity": pm_red,
                "AlertDtYellow": alert_dt,
                "ProgressFlag": progress_flag
            }
            
            #if pl_red == 1 or pm_red == 1 or alert_dt == 1 or progress_flag > 0:
                #print(f"[Sysmemory] Task '{task_name}' - PLRed: {pl_red}, PMRed: {pm_red}, AlertDt: {alert_dt}, ProgressFlag: {progress_flag}")
        
        #print(f"[Sysmemory] Loaded {len(task_map)} tasks for {project_id}")
        return task_map
        
    except Exception as e:
        print(f"[Sysmemory] Error reading {project_id}: {e}")
        import traceback
        traceback.print_exc()
        return {}

def load_user_overdue_status(user):
    """
    Load overdue status for all projects assigned to the user.
    If sysmemory JSON doesn't exist, generate it.
    """
    from datetime import datetime
    import os
    
    #print(f"[Overdue] Loading overdue status for user: {user.get('username')}")
    
    # Get projects for this user
    projects = get_projects_for_user(user)
    #print(f"[Overdue] Found {len(projects)} projects for user")
    
    overdue_map = {}
    role = user.get("role", "sw_tl")
    
    # Get discipline directories
    projects_dir, _, sched_cache, _ = get_discipline_dirs(role)
    
    department = ROLE_TO_DEPT.get(role, "SW")
    
    for project in projects:
        project_id = project.get("file_id") or project.get("id")
        #print(f"[DEBUG] project_id key: '{project_id}'")  # <-- ADD THIS LINE
        if not project_id:
            continue
        
        # NEW: Skip if project has no sheet cache (new project)
        sheet_cache_path = os.path.join(sched_cache, project_id + "_sheet.json")
        if not os.path.exists(sheet_cache_path):
            print(f"[Overdue] Skipping {project_id} - no sheet cache yet")
            continue
        
        # Path to sysmemory JSON
        sysmemory_path = os.path.join(DATA_DIR, "cache", "SystemMemory", department, f"{project_id}_sysmemory.json")
        
        # If sysmemory JSON doesn't exist, generate it
        if not os.path.exists(sysmemory_path):
            print(f"[Overdue] Generating sysmemory for {project_id}")
            # Find the Excel file path
            excel_path = os.path.join(projects_dir, project_id + ".xlsx")
            if not os.path.exists(excel_path):
                excel_path = os.path.join(projects_dir, project_id + ".xlsb")
            
            if os.path.exists(excel_path):
                from excel_db import generate_sysmemory_json
                generate_sysmemory_json(excel_path, project_id, sched_cache)
            else:
                print(f"[Overdue] Excel file not found for {project_id}")
                continue
        
        # Read sysmemory for this project
        task_map = read_sysmemory_json(project_id, role)
        
        if task_map:
            overdue_map[project_id] = task_map
            #print(f"[Overdue] Loaded {len(task_map)} tasks for {project_id}")
        else:
            print(f"[Overdue] No tasks loaded for {project_id}")
    
    #print(f"[Overdue] Total projects loaded: {len(overdue_map)}")
    return overdue_map

def _parse_date_value(val):
    """
    Convert any date-like value (int serial, float serial, or date string) to a
    Python datetime.  Returns None if conversion fails.
    Used by _fix_int_dates and _do_excel_write to keep date cells clean.
    """
    from openpyxl.utils.datetime import from_excel as _fxl
    from datetime import datetime as _dt, timedelta

    _DATE_FMTS = (
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%Y/%m/%d",
        "%d-%m-%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%d %b %Y, %I:%M %p",
    )

    if val is None:
        return None
    if isinstance(val, (_dt,)):
        return val
    # date (not datetime) — promote to datetime
    try:
        from datetime import date as _date
        if isinstance(val, _date):
            return _dt(val.year, val.month, val.day)
    except Exception:
        pass
    if isinstance(val, int):
        try:
            if 1 <= val <= 2958465:
                return _fxl(val)
            return None
        except Exception:
            try:
                return _dt(1899, 12, 30) + timedelta(days=val)
            except Exception:
                return None
    if isinstance(val, float):
        try:
            if 1 < val < 2958465:
                return _fxl(val)
            return None
        except Exception:
            return None
    if isinstance(val, str):
        s = val.strip()
        if not s:
            return None
        for fmt in _DATE_FMTS:
            try:
                return _dt.strptime(s, fmt)
            except ValueError:
                continue
        return None
    return None

def _safe_fix_schedule_dates(ws):
    """
    Full sheet scan: fix every date-formatted cell holding a raw int/float
    so openpyxl does not crash with 'int has no attribute year' on save.
    Formula cells (data_type == 'f') are always skipped.
    """
    from openpyxl.styles.numbers import is_date_format as _idf
    from openpyxl.utils.datetime import from_excel
    from datetime import datetime as _dt, timedelta

    fixed = 0
    for row in ws.iter_rows():
        for cell in row:
            # Only care about raw int or float
            if not isinstance(cell.value, (int, float)):
                continue
            # Never touch formula cells
            if cell.data_type == 'f':
                continue
            # Only act on date-formatted cells
            try:
                if not cell.number_format or not _idf(cell.number_format):
                    continue
            except Exception:
                continue
            # Convert valid serial to datetime; clear anything invalid
            if 1 <= cell.value <= 2958465:
                try:
                    cell.value = from_excel(cell.value)
                    fixed += 1
                except Exception:
                    try:
                        cell.value = _dt(1899, 12, 30) + timedelta(days=int(cell.value))
                        fixed += 1
                    except Exception:
                        cell.value = None
            else:
                cell.value = None

    if fixed:
        print(f"[Background] Fixed {fixed} date-serial cell(s) before save")
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
 
    try:
        # IMPORTANT: Do NOT use data_only=True or keep_links=False here.
        # Loading with data_only=True and then saving permanently destroys
        # all formula cells (openpyxl replaces them with cached values or None).
        # keep_links=False can also corrupt external references.
        wb = load_workbook(fpath)
    except Exception as e:
        print(f"[Background] Error loading {project_id}: {e}")
        return
 
    ws = wb.active
 
    # Fix all int/float values sitting in date-formatted cells BEFORE updates.
    # _safe_fix_schedule_dates only touches known date columns (X, Y, J-Q, V, W)
    # and skips formula cells (data_type == 'f'), so formulas are never overwritten.
    _safe_fix_schedule_dates(ws)
 
    # Apply updates
    row_updates  = {u["_row"]:      u for u in updates if "_row"      in u}
    task_updates = {u["task_key"]:  u for u in updates if "task_key"  in u}
 
    TASK_START_ROW = 9
    TASK_END_ROW   = 55
 
    for row in range(TASK_START_ROW, TASK_END_ROW + 1):
        if row in row_updates:
            u = row_updates[row]
            if "actual_start" in u and u["actual_start"]:
                # Always convert to datetime — writing a raw string/int into a
                # date-formatted cell causes 'int has no attribute year' on save.
                ws[f"X{row}"] = _parse_date_value(u["actual_start"])
            if "actual_end" in u and u["actual_end"]:
                ws[f"Y{row}"] = _parse_date_value(u["actual_end"])
            if "percent_complete" in u and u["percent_complete"] is not None:
                try:
                    ws[f"Z{row}"] = float(u["percent_complete"]) / 100.0
                except (ValueError, TypeError):
                    pass
            if "help_required" in u:
                ws[f"AD{row}"] = u["help_required"] or ""
            if "remark" in u:
                ws[f"AF{row}"] = u["remark"] or ""
        else:
            code = ws[f"H{row}"].value
            name = ws[f"I{row}"].value
            if code and name:
                row_key = f"{code}_{str(name).strip()}"
                if row_key in task_updates:
                    u = task_updates[row_key]
                    ws[f"Z{row}"] = u["percent_complete"] / 100.0
                    if "actual_start" in u and u["actual_start"]:
                        ws[f"X{row}"] = _parse_date_value(u["actual_start"])
                    if "actual_end" in u and u["actual_end"]:
                        ws[f"Y{row}"] = _parse_date_value(u["actual_end"])
                    ws[f"AF{row}"] = u.get("remark") or ""
 
    # Second pass — catch any ints that crept in during the update writes
    _safe_fix_schedule_dates(ws)

    # DEBUG: find exactly which cell still has int in date-format before save
    from openpyxl.styles.numbers import is_date_format as _idf2
    for _row in ws.iter_rows():
        for _cell in _row:
            if isinstance(_cell.value, (int, float)) and _cell.number_format:
                try:
                    if _idf2(_cell.number_format):
                        print(f"[DEBUG] PROBLEM CELL: {_cell.coordinate} value={_cell.value} data_type={_cell.data_type} fmt={_cell.number_format}")
                except Exception:
                    pass

    try:
        wb.save(fpath)
        print(f"[Background] Excel file saved: {fpath}")
    except Exception as e:
        print(f"[Background] Error writing {project_id}: {e}")
        return
 
    # Regenerate sysmemory from the freshly saved file
    try:
        _, _, sched_cache, _ = get_discipline_dirs(role)
        generate_sysmemory_json(fpath, project_id, sched_cache)
    except Exception as e:
        print(f"[Background] Sysmemory read failed for {project_id}: {e}")

def queue_excel_write(project_id, updates, role):
    """Queue a project for background Excel write"""
    _start_background_worker()  # Ensure worker is running
    _excel_write_queue.put((project_id, updates, role))

def _safe_fix_monitor_dates(ws, department):
    """
    Scan ALL cells in the monitor sheet and fix any date-formatted cell that
    holds a raw int/float (Excel serial) instead of a proper datetime object.
 
    openpyxl's serialiser calls `to_excel(value, epoch)` on every cell whose
    number_format is a date pattern.  If the value is a plain int/float it
    crashes with:
        AttributeError: 'int' object has no attribute 'year'
 
    Restricting to a single column (BA) was not enough because the monitor
    sheet can have date-formatted cells in other columns too (e.g. columns
    inherited from the source xlsx that already had date formats applied before
    any data was written).  We now scan the entire used range so no stray
    int/float in a date cell can slip through.
 
    Formula cells are always skipped — we never destroy a formula.
    """
    from openpyxl.styles.numbers import is_date_format as _idf
    from openpyxl.utils.datetime import from_excel
    from datetime import datetime
 
    fixed = 0
    for row in ws.iter_rows():
        for cell in row:
            # Skip empty cells
            if cell.value is None:
                continue
 
            # CRITICAL: Never overwrite formula cells.
            if cell.data_type == 'f' or (
                isinstance(cell.value, str) and str(cell.value).startswith('=')
            ):
                continue
 
            # Only act on date-formatted cells that hold a raw number
            if not isinstance(cell.value, (int, float)):
                continue
 
            try:
                if not cell.number_format or not _idf(cell.number_format):
                    continue
            except Exception:
                continue
 
            # Valid Excel date serial range (1 = 1900-01-01, 2958465 = 9999-12-31)
            if 1 <= cell.value <= 2958465:
                try:
                    cell.value = from_excel(cell.value)
                    fixed += 1
                except Exception:
                    # If conversion fails, clear the value so openpyxl won't
                    # crash trying to serialise a bare int as a date.
                    cell.value = None
            else:
                # Out-of-range numeric in a date cell — clear it to avoid crash.
                cell.value = None
 
    if fixed:
        print(f"[Monitor] Fixed {fixed} date-serial cells before save")                            

def _do_monitor_excel_write(department, updates):
    """Write monitor cache updates to actual Excel file in background"""
    import os
    from openpyxl import load_workbook
    from openpyxl.styles.numbers import is_date_format as _is_date_fmt
    from openpyxl.utils.datetime import from_excel as _from_excel
    from datetime import datetime as _dt
    
    monitor_path = find_monitor_file(department)
    if not monitor_path:
        print(f"[Monitor] No monitor file found for {department}")
        return []
        
    if not os.path.exists(monitor_path):
        print(f"[Monitor Background] File not found: {monitor_path}")
        return
    
    try:
        # Load workbook
        try:
            # CRITICAL: Do NOT use data_only=True here.
            # Loading with data_only=True and then saving permanently destroys
            # all formula cells — openpyxl replaces them with their cached
            # values (or None), corrupting the file on every write.
            wb = load_workbook(monitor_path)
        except EOFError:
            print(f"[Monitor Background] EOFError - file may be corrupted: {monitor_path}")
            return
        
        ws = wb["SWMon"] if "SWMon" in wb.sheetnames else wb.active
 
        # Use the correct department sheet name from config
        dept_sheet = DEPT_CONFIG.get(department, {}).get("monitor_sheet", "SWMon")
        if dept_sheet in wb.sheetnames:
            ws = wb[dept_sheet]
        elif "SWMon" in wb.sheetnames:
            ws = wb["SWMon"]
        else:
            ws = wb.active
 
        # Process all updates
        for coord, value in updates.items():
            if value is None:
                continue
            cell = ws[coord]
 
            # CRITICAL: Never overwrite formula cells — this would destroy the
            # formula and replace it permanently with a raw value.
            if cell.data_type == 'f' or (isinstance(cell.value, str) and str(cell.value).startswith('=')):
                continue
            
            # Handle date format cells (BA column)
            cell_is_date_fmt = False
            try:
                if cell.number_format:
                    cell_is_date_fmt = _is_date_fmt(cell.number_format)
            except Exception:
                pass
            
            # Convert timestamp string to datetime if needed
            if cell_is_date_fmt and isinstance(value, str):
                for fmt in ("%d-%m-%Y %H:%M:%S", "%d %b %Y, %I:%M %p",
                            "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
                    try:
                        value = _dt.strptime(value, fmt)
                        break
                    except ValueError:
                        continue
            
            # For percentage cells, they are stored as strings with %
            # No conversion needed for non-date cells
            
            cell.value = value
 
        # Fix any pre-existing int/float values in date-formatted cells before saving
        _safe_fix_monitor_dates(ws, department)
 
        wb.save(monitor_path)
        print(f"[Monitor Background] Excel saved: {monitor_path}")
 
    except Exception as e:
        import traceback
        print(f"[Monitor Background] Error writing to Excel: {e}")
        traceback.print_exc()

def _start_monitor_worker():
    """Start a single persistent background thread for monitor Excel writes.
    Using a queue (not a new thread per write) prevents the race condition where
    two near-simultaneous writes both try to open the file — which caused
    EOFError / 'File is not a zip file' because one thread read a half-written file."""
    global _monitor_thread_started
    if _monitor_thread_started:
        return

    def worker():
        while True:
            try:
                department, updates = _monitor_write_queue.get(timeout=1)
                try:
                    _do_monitor_excel_write(department, updates)
                except Exception as e:
                    print(f"[Monitor Background] Worker error: {e}")
                finally:
                    _monitor_write_queue.task_done()
            except queue.Empty:
                continue
            except Exception as e:
                print(f"[Monitor Background] Worker fatal error: {e}")

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()
    _monitor_thread_started = True
    print("[Monitor Background] Monitor write worker started")

def queue_monitor_excel_write(department, updates):
    """Queue a monitor Excel write to the single persistent background worker.
    Never spawns a new thread — all writes are serialised through the queue."""
    _start_monitor_worker()
    _monitor_write_queue.put((department, updates))

def update_monitor_timestamp(project_id, timestamp, role="sw_tl"):
    from datetime import datetime
    import json
    import os

    department = ROLE_TO_DEPT.get(role, "SW")
    cfg = DEPT_CONFIG[department]
    pid_col = cfg["file_col"]  # Use config instead of hardcoded 'B'

    monitor_cache_path = os.path.join(get_cache_path(department, "Monitoring"), f"{department}_Monitor.json")

    if not os.path.exists(monitor_cache_path):
        print(f"[Monitor] Cache not found: {monitor_cache_path}")
        return {}

    try:
        with open(monitor_cache_path, 'r') as f:
            monitor_data = json.load(f)

        cells = monitor_data.get('cells', {})

        found_row = None
        for coord, cell_info in cells.items():
            if coord.startswith(pid_col) and cell_info.get('v') == project_id:
                found_row = int(coord[len(pid_col):])
                break

        if found_row:
            ba_coord = f"{cfg['timestamp_col']}{found_row}"
            if ba_coord not in cells:
                cells[ba_coord] = {}
            cells[ba_coord]['v'] = timestamp
            cells[ba_coord]['updated'] = True

            with open(monitor_cache_path, 'w') as f:
                json.dump(monitor_data, f, indent=2)

            print(f"[Monitor] Updated cache for {project_id} at row {found_row}: {timestamp}")
            return {ba_coord: timestamp}
        else:
            print(f"[Monitor] Project {project_id} not found in monitor cache (searched column {pid_col})")
            return {}

    except Exception as e:
        print(f"[Monitor] Failed to update timestamp: {e}")
        return {}

def update_monitor_task_percentages(project_id, role="sw_tl"):
    """
    Update task completion percentages in monitor JSON cache.
    Reads all tasks from project cache and writes them to monitor columns.
    Returns dict of updates applied.
    """
    from datetime import datetime
    import json
    import os
    from openpyxl.utils import get_column_letter as gcl
    
    department = ROLE_TO_DEPT.get(role, "SW")
    cfg = DEPT_CONFIG[department]
    monitor_cache_path = os.path.join(get_cache_path(department, "Monitoring"), f"{department}_Monitor.json")
    
    if not os.path.exists(monitor_cache_path):
        print(f"[Monitor] Cache not found: {monitor_cache_path}")
        return {}
    
    try:
        # Load monitor cache
        with open(monitor_cache_path, 'r') as f:
            monitor_data = json.load(f)
        
        cells = monitor_data.get('cells', {})
        
        # Find the row with matching project ID in COLUMN F
        pid_col = cfg.get("file_col", cfg.get("project_id_col", "B"))
        found_row = None
        for coord, cell_info in cells.items():
            if coord.startswith(pid_col) and cell_info.get('v') == project_id:
                found_row = int(coord[len(pid_col):])
                break
        
        if not found_row:
            print(f"[Monitor] Project {project_id} not found in monitor cache")
            return {}
        
        # Get column headers from row 10 (BB to CV)
        # BB = column 54, let's go up to column 100 (CV)
        col_headers = {}
        start_col_idx = 54  # BB
        end_col_idx = 100   # CV
        
        header_row = cfg["data_start_row"] - 2
        for col_idx in range(start_col_idx, end_col_idx + 1):
            col_letter = gcl(col_idx)
            coord = f"{col_letter}{header_row}"  # row 10 is header row
            header_val = cells.get(coord, {}).get('v')
            if header_val and str(header_val).strip():
                col_headers[str(header_val).strip()] = col_letter
        
        # Get task data from project cache
        task_data = get_task_data_from_cache(project_id, role)
        
        if not task_data:
            print(f"[Monitor] No task data found for {project_id}")
            return {}
        
        updates = {}
        
        # Match and update
        for task_name, percent in task_data.items():
            if task_name in col_headers:
                col_letter = col_headers[task_name]
                coord = f"{col_letter}{found_row}"
                
                # Format percent as string with % (e.g., "75%")
                percent_str = f"{percent}%"
                
                # Update JSON cache immediately
                if coord not in cells:
                    cells[coord] = {}
                cells[coord]['v'] = percent_str
                cells[coord]['updated'] = True
                
                # Add to updates dict for Excel write
                updates[coord] = percent_str
                
                print(f"[Monitor] Updated {task_name} at {coord}: {percent_str}")
            else:
                print(f"[Monitor] No matching column for task: {task_name}")
        
        # Write back to JSON cache
        with open(monitor_cache_path, 'w') as f:
            json.dump(monitor_data, f, indent=2)
        
        return updates
        
    except Exception as e:
        print(f"[Monitor] Failed to update task percentages: {e}")
        import traceback
        traceback.print_exc()
        return {}

def create_new_project_from_monitor(or_number, section, ov_value, assign_to, user):
    """Create new project from monitor UI"""
    import json
    import os
    import shutil
    from datetime import datetime
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter as gcl
    
    department = "PM"
    cfg = DEPT_CONFIG[department]
    
    # Generate file name: PrjSch_{OR_number}_{section}.xlsx
    file_name = f"PrjSch_{or_number.replace('/', '_')}_{section}.xlsx"
    projects_dir, _, _, _ = get_discipline_dirs("pm_head")
    file_path = os.path.join(projects_dir, file_name)
    
    # Check if file already exists
    if os.path.exists(file_path):
        return {'error': f'Project file already exists: {file_name}'}
    
    # Copy template file
    template_folder = os.path.join(DATA_DIR, "Template")
    template_files = [f for f in os.listdir(template_folder) if f.startswith("PrjSch_Template") and f.endswith(".xlsx")]

    if not template_files:
        return {'error': 'No PM template found in Template folder'}

    template_path = os.path.join(template_folder, template_files[0])
    shutil.copy2(template_path, file_path)
    
    # Load workbook with data_only=False to preserve formulas
    wb = load_workbook(file_path, data_only=False)
    ws = wb.active
    
    # ========== FIX DATE CELLS BEFORE ANY MODIFICATIONS ==========
    _safe_fix_schedule_dates(ws)
    
    # Now write the new values
    ws['D1'] = or_number
    ws['D14'] = section
    ws['D9'] = float(ov_value) if ov_value else 0
    
    # Second pass - fix any dates that might have been affected
    _safe_fix_schedule_dates(ws)
    
    wb.save(file_path)
    wb.close()
    
    # Append new row to monitor JSON
    monitor_cache_path = os.path.join(get_cache_path(department, "Monitoring"), f"{department}_Monitor.json")
    
    if os.path.exists(monitor_cache_path):
        with open(monitor_cache_path, 'r') as f:
            monitor_data = json.load(f)
        
        cells = monitor_data.get('cells', {})
        
        # Find next available row
        max_row = monitor_data.get('max_row', cfg['data_start_row'])
        new_row = max_row + 1
        
        # Set values for new row
        cells[f"{cfg['file_col']}{new_row}"] = {'v': file_name.replace('.xlsx', '')}
        cells[f"{cfg['head_col']}{new_row}"] = {'v': assign_to.upper()}
        cells[f"G{new_row}"] = {'v': or_number}
        cells[f"H{new_row}"] = {'v': section}

        normalized = or_number + "_" + section
        cells[f"E{new_row}"] = {'v': normalized}
        cells[f"K{new_row}"] = {'v': float(ov_value) if ov_value else 0}
        cells[f"DA{new_row}"] = {'v': 'NEW'}
        
        monitor_data['max_row'] = new_row
        
        with open(monitor_cache_path, 'w') as f:
            json.dump(monitor_data, f, indent=2)

        try:
            import sys
            if 'app' in sys.modules:
                app_module = sys.modules['app']
                if hasattr(app_module, '_master_projects_cache'):
                    app_module._master_projects_cache.clear()
                    print(f"[Cache] Cleared master projects cache for: {file_name}")
        except Exception as e:
            print(f"[Cache] Could not clear master cache: {e}")
    
    return {'success': True, 'file_name': file_name, 'project_id': file_name.replace('.xlsx', '')}

def _generate_cache_with_pycel(project_id, file_path, sched_cache=None, role=None):
    """
    Generate the sheet JSON cache after a setup save, using PyCel to evaluate
    formula cells and openpyxl (data_only=True) for plain cell values.

    This is called instead of get_raw_sheet() after write_project_setup_data()
    saves the Excel file, because get_raw_sheet() opens with data_only=False
    which returns raw formula strings instead of computed values.

    Two workbook handles are used from the same already-saved file:
      - wb_vals : data_only=True  → correct plain cell values + formatting
      - wb_fmt  : data_only=False → correct data_type flags for formula detection
    PyCel compiles the formula graph from the saved file, so it sees the
    newly written form values when evaluating dependent formulas.
    """
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter as gcl

    if not PY_CEL_AVAILABLE:
        print("[SetupCache] PyCel not available, falling back to get_raw_sheet")
        return get_raw_sheet(file_path, sched_cache=sched_cache, role=role)

    dept = ROLE_TO_DEPT.get(role, "SW") if role else detect_dept_from_path(file_path)
    cfg  = DEPT_CONFIG.get(dept, DEPT_CONFIG["SW"])
    date_rows = cfg.get("date_rows", [28, 29, 30, 31, 32, 33])

    # ── Two workbook handles ───────────────────────────────────
    # wb_vals: plain values + formatting (data_only=True)
    # wb_fmt:  formula strings + data_type flags (data_only=False)
    wb_vals = load_workbook(file_path, data_only=True)
    wb_fmt  = load_workbook(file_path, data_only=False)
    ws_vals = wb_vals.active
    ws_fmt  = wb_fmt.active
    sheet_name = ws_fmt.title

    # ── PyCel compiler ─────────────────────────────────────────
    excel = None
    try:
        excel = ExcelCompiler(file_path)
        print(f"[SetupCache] PyCel compiled successfully for {file_path}")
    except Exception as e:
        print(f"[SetupCache] PyCel compile failed ({e}), all cells will use data_only values")
        excel = None

    max_row = ws_vals.max_row
    max_col = 32
    scope_config = cfg.get("scopeConfig")
    if scope_config and scope_config.get("enabled"):
        from openpyxl.utils import column_index_from_string
        for cell_info in scope_config.get("cells", {}).values():
            col_letter = cell_info.get("col", "")
            if col_letter:
                max_col = max(max_col, column_index_from_string(col_letter))

    max_col = min(ws_vals.max_column, max_col)
    cols = [gcl(i) for i in range(1, max_col + 1) if gcl(i) not in ("A", "B", "C", "D")]

    # ── Helpers ────────────────────────────────────────────────
    def _resolve_color(color):
        try:
            if color is None:
                return None
            if color.type == "rgb":
                rgb = color.rgb
                if rgb and rgb != "00000000":
                    return f"#{rgb[2:]}"
            return None
        except Exception:
            return None

    def _eval_cell(coord):
        """
        Evaluate a cell. Uses PyCel for formula cells, wb_vals for plain cells.
        Silently falls back to wb_vals value on any PyCel error.
        """
        cell_fmt  = ws_fmt[coord]
        cell_vals = ws_vals[coord]
        fmt = cell_fmt.number_format or ""

        # Determine if it's a formula
        is_formula = (cell_fmt.data_type == "f") or (
            isinstance(cell_fmt.value, str) and cell_fmt.value.startswith("=")
        )

        v = None
        if is_formula and excel is not None:
            try:
                v = excel.evaluate(f"{sheet_name}!{coord}")
            except Exception:
                # Silently fall back to cached value
                v = cell_vals.value if cell_vals else None
        else:
            v = cell_vals.value if cell_vals else ""
        
        # If still None, use empty string (not formula text)
        if v is None:
            v = ""

        # ── Type normalisations ────────────────────────────────
        if isinstance(v, (datetime, date)):
            d = v.date() if isinstance(v, datetime) else v
            return d.strftime("%d-%b-%y").upper() 

        if isinstance(v, (int, float)) and not isinstance(v, bool):
            # Excel date serial check
            if 30000 <= v <= 100000 and any(
                p in fmt.lower() for p in ["yy", "mm", "dd", "mmm"]
            ):
                try:
                    from openpyxl.utils.datetime import from_excel
                    return from_excel(int(v)).strftime("%d-%b-%y").upper()  # 16-JUN-26
                except Exception:
                    pass
            if "%" in fmt:
                return f"{int(round(float(v) * 100))}%"
            if fmt == "00":
                return f"{int(v):02d}"
            if fmt == "0.0":
                return f"{float(v):.1f}"
            if fmt == "0.00":
                return f"{float(v):.2f}"

        if v is not None and not isinstance(v, (int, float, bool, str)):
            v = str(v)

        return v

    def _plain_val(ref):
        """Read a plain string value from ws_vals, no formula evaluation."""
        v = ws_vals[ref].value
        return str(v).strip() if v else ""

    def _date_val(ref):
        """Read a date cell, evaluating via PyCel if it is a formula."""
        v = _eval_cell(ref)
        if v is None:
            return ""
        if isinstance(v, str):
            return v  # already formatted by _eval_cell
        if isinstance(v, (int, float)) and 30000 <= v <= 100000:
            try:
                from openpyxl.utils.datetime import from_excel
                return from_excel(int(v)).strftime("%d %b %Y")
            except Exception:
                pass
        if isinstance(v, (datetime, date)):
            d = v.date() if isinstance(v, datetime) else v
            return d.strftime("%d %b %Y")
        return str(v).strip() if v else ""

    # ── Merged cells (from ws_fmt) ─────────────────────────────
    merged_map = {}
    for mc in ws_fmt.merged_cells.ranges:
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

    # ── Column widths + hidden + outline ──────────────────────
    DEFAULT_COL_WIDTH = getattr(ws_fmt.sheet_format, "defaultColWidth", None) or 0.88
    col_widths  = {}
    col_hidden  = {}
    col_outline = {}
    for col in cols:
        cd = ws_fmt.column_dimensions.get(col)
        if cd is None:
            col_widths[col]  = max(30, round((DEFAULT_COL_WIDTH or 8) * 7.5))
            col_hidden[col]  = DEFAULT_COL_WIDTH < 2.0
            col_outline[col] = 0
        else:
            col_widths[col]  = max(30, round((cd.width or 8) * 7.5))
            col_hidden[col]  = bool(cd.hidden) and int(cd.outline_level or 0) > 0
            col_outline[col] = int(cd.outline_level or 0)

    _col_overrides = {"B": 0.90, "C": 0.80, "D": 0.35, "I": 0.60, "V": 0.90, "W": 0.90}
    for _col, _factor in _col_overrides.items():
        if _col in col_widths:
            col_widths[_col] = max(30, round(col_widths[_col] * _factor))

    # ── col_groups ────────────────────────────────────────────
    col_groups = []
    try:
        i = 0
        while i < len(cols):
            col = cols[i]
            cd_check = ws_fmt.column_dimensions.get(col)
            is_grouped = col_outline.get(col, 0) > 0 or (
                cd_check is None and DEFAULT_COL_WIDTH < 2.0
            )
            if is_grouped:
                group_cols = [col]
                j = i + 1
                while j < len(cols):
                    next_col = cols[j]
                    cd_next  = ws_fmt.column_dimensions.get(next_col)
                    next_grouped = col_outline.get(next_col, 0) > 0 or (
                        cd_next is None and DEFAULT_COL_WIDTH < 2.0
                    )
                    if next_grouped:
                        group_cols.append(next_col)
                        j += 1
                    else:
                        break
                if group_cols[0] == "J" and "Q" in group_cols:
                    group_cols = group_cols[: group_cols.index("Q")]
                if group_cols[0] == "X" and "AB" in cols and "AB" not in group_cols:
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
        rd = ws_fmt.row_dimensions.get(rn)
        row_heights[str(rn)] = max(18, round((rd.height or 15) * 1.33)) if rd else 20

    # ── Cells ─────────────────────────────────────────────────
    cells = {}
    for row in ws_fmt.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell_fmt in row:
            coord = cell_fmt.coordinate
            mi    = merged_map.get(coord, {})

            if mi.get("skip"):
                cells[coord] = {"skip": True}
                continue

            v = _eval_cell(coord)
            c = {"v": v}

            if mi.get("master"):
                if mi["rowspan"] > 1: c["rowspan"] = mi["rowspan"]
                if mi["colspan"] > 1: c["colspan"] = mi["colspan"]

            try:
                f = ws_vals[coord].font
                font = {}
                if f.bold: font["bold"] = True
                if f.size and f.size != 11: font["size"] = f.size
                if font: c["font"] = font
            except Exception:
                pass

            try:
                fill = ws_vals[coord].fill
                if fill and fill.patternType not in (None, "none"):
                    fg = _resolve_color(fill.fgColor)
                    if fg and fg != "#FFFFFF": c["fill"] = fg
            except Exception:
                pass

            cells[coord] = c

    # ── Editable cells ────────────────────────────────────────
    TASK_START_ROW = 9
    EDITABLE_COLS  = cfg.get("editable_cols", {"X", "Y", "Z", "AD", "AF"})
    editable_fill  = None
    editable_theme = None
    editable_tint  = None
    TINT_TOLERANCE = 0.001

    try:
        ref = ws_vals[f"X{TASK_START_ROW}"]
        fg  = ref.fill.fgColor
        if ref.fill.patternType == "solid" and fg.type == "theme":
            editable_theme = fg.theme
            editable_tint  = fg.tint
            editable_fill  = "#BDD7EE"
    except Exception:
        pass

    def _is_editable_fill(cell):
        if editable_theme is None:
            return False
        try:
            fg = cell.fill.fgColor
            return (
                cell.fill.patternType == "solid"
                and fg.type == "theme"
                and fg.theme == editable_theme
                and abs(fg.tint - editable_tint) < TINT_TOLERANCE
            )
        except Exception:
            return False

    for row in ws_vals.iter_rows(min_row=TASK_START_ROW, max_row=max_row):
        for cell in row:
            try:
                col_letter = cell.column_letter
            except AttributeError:
                continue
            if col_letter not in EDITABLE_COLS:
                continue
            coord = cell.coordinate
            if coord not in cells or cells[coord].get("skip"):
                continue

            is_always_editable = col_letter in ("X", "Y", "Z", "AF")
            if is_always_editable or _is_editable_fill(cell):
                if col_letter == "AD":
                    AD_COLORS = {
                        "engineering":        "#ffb3b3",
                        "purchase":           "#5f933c",
                        "software":           "#0096cc",
                        "project management": "#005fa3",
                        "manufacturing":      "#2f491e",
                        "sales":              "#00d9d9",
                        "client":             "#6d006d",
                    }
                    val = cell.value
                    if val:
                        color = AD_COLORS.get(str(val).strip().lower())
                        if color:
                            cells[coord]["fill"] = color
                            dark_bgs = {"#5f933c", "#0096cc", "#005fa3", "#2f491e", "#6d006d"}
                            cells[coord]["font"] = cells[coord].get("font", {})
                            cells[coord]["font"]["color"] = "#ffffff" if color in dark_bgs else "#000000"
                    z_coord = f"Z{cell.row}"
                    z_val   = ws_vals[z_coord].value
                    z_pct   = 0
                    try:
                        if z_val is not None:
                            z_pct = float(z_val)
                            if z_pct <= 1.0:
                                z_pct *= 100
                    except (TypeError, ValueError):
                        z_pct = 0
                    if z_pct < 100:
                        cells[coord]["editable"]     = True
                        cells[coord]["editable_col"] = col_letter
                else:
                    cells[coord]["editable"]     = True
                    cells[coord]["editable_col"] = col_letter

    # ── project_banner ────────────────────────────────────────
    w2_val = _date_val("W2") or _date_val("X9")
    project_banner = {
        "or_number":      _plain_val("D1"),
        "section":        _plain_val("D7"),
        "sales_engineer": _plain_val("I2"),
        "sales_manager":  _plain_val("I3"),
        "customer_name":  _plain_val("D10"),
        "start_date_lbl": _plain_val("W1"),
        "start_date_val": w2_val,
        "days_swe_lbl":   _plain_val("AB1"),
        "days_swe_val":   _plain_val("AB2"),
        "ld_date_lbl":    _plain_val("AE1"),
        "ld_date_val":    _plain_val("AF1"),
        "ld_maxwk_lbl":   _plain_val("AE2"),
        "ld_maxwk_val":   _plain_val("AF2"),
        "ld_maxov_lbl":   _plain_val("AE3"),
        "ld_maxov_val":   _plain_val("AF3"),
        "ld_remarks_lbl": _plain_val("AE4"),
        "ld_remarks_val": _plain_val("AF4"),
    }

    # ── left_panel ────────────────────────────────────────────
    dates_list = []
    for row_num in date_rows:
        dates_list.append({
            "label":    _plain_val(f"B{row_num}") or f"B{row_num}",
            "pm":       _date_val(f"C{row_num}"),
            "swe":      _date_val(f"D{row_num}"),
            "pm_fill":  _resolve_color(ws_vals[f"C{row_num}"].fill.fgColor)
                        if ws_vals[f"C{row_num}"].fill.patternType not in (None, "none") else None,
            "pm_font":  _resolve_color(ws_vals[f"C{row_num}"].font.color),
            "swe_fill": _resolve_color(ws_vals[f"D{row_num}"].fill.fgColor)
                        if ws_vals[f"D{row_num}"].fill.patternType not in (None, "none") else None,
            "swe_font": _resolve_color(ws_vals[f"D{row_num}"].font.color),
        })

    left_panel = {
        "project_info": [
            {"label": _plain_val("C9")  or "PO Value",      "value": str(_plain_val("D9") or "")},
            {"label": _plain_val("C10") or "Customer",      "value": _plain_val("D10")},
            {"label": _plain_val("C11") or "End Customer",  "value": _plain_val("D11")},
            {"label": _plain_val("C12") or "Consultant",    "value": _plain_val("D12")},
            {"label": _plain_val("C13") or "Project Desc.", "value": _plain_val("D13")},
            {"label": _plain_val("C14") or "Section",       "value": _plain_val("D14")},
            {"label": _plain_val("C17") or "SW Efforts",    "value": str(_plain_val("D17") or "")},
        ],
        "dates": dates_list,
        "warranty": [
            {"label": _plain_val("AE6") or "Warranty", "value": _plain_val("AF6")},
        ],
        "stakeholders": [
            {"label": _plain_val("B36") or "Sales", "value": _plain_val("C36")},
            {"label": _plain_val("B37") or "PM",    "value": _plain_val("C37")},
            {"label": _plain_val("B38") or "HW",    "value": _plain_val("C38")},
            {"label": _plain_val("B41") or "SW",    "value": _plain_val("C41")},
            {"label": _plain_val("B40") or "MFG",   "value": _plain_val("C40")},
            {"label": _plain_val("B42") or "E&C",   "value": _plain_val("C42")},
            {"label": _plain_val("B39") or "BYR",   "value": _plain_val("C39")},
            {"label": _plain_val("B43") or "A/C",   "value": _plain_val("C43")},
        ],
    }

    # extra_sections from config
    extra_sections_list = []
    for section in cfg.get("extra_sections", []):
        section_data = {"title": section.get("title", ""), "rows": []}
        for row_cfg in section.get("rows", []):
            row_num     = row_cfg.get("row")
            value_coord = f"C{row_num}"
            label_coord = f"B{row_num}"
            value       = cells.get(value_coord, {}).get("v") if value_coord in cells else None
            label_excel = cells.get(label_coord, {}).get("v") if label_coord in cells else None
            row_data = {
                "label": label_excel or row_cfg.get("label", ""),
                "value": str(value) if value is not None else "",
            }
            if row_cfg.get("format") == "percent" and value is not None:
                try:
                    num = float(value)
                    if 0 <= num <= 1:
                        row_data["value"] = f"{int(num * 100)}%"
                except (ValueError, TypeError):
                    pass
            if "actual_col" in row_cfg:
                actual_coord = f"{row_cfg['actual_col']}{row_num}"
                actual_value = cells.get(actual_coord, {}).get("v") if actual_coord in cells else None
                row_data["actual"] = str(actual_value) if actual_value is not None else ""
                if row_cfg.get("format") == "percent" and actual_value is not None:
                    try:
                        num = float(actual_value)
                        if 0 <= num <= 1:
                            row_data["actual"] = f"{int(num * 100)}%"
                    except (ValueError, TypeError):
                        pass
            section_data["rows"].append(row_data)
        if section_data["rows"]:
            extra_sections_list.append(section_data)

    left_panel["extra_sections"] = extra_sections_list

    # ── last_modified ─────────────────────────────────────────
    last_modified_str = None
    try:
        ts = os.path.getmtime(file_path)
        last_modified_str = datetime.fromtimestamp(ts).strftime("%d %b %Y, %I:%M %p")
    except Exception:
        pass

    # ── Scope (same logic as get_raw_sheet) ──────────────────────
    scope_data = None
    scope_config = cfg.get("scopeConfig")
    if scope_config and scope_config.get("enabled"):
        scope_data = {
            "enabled": True,
            "cells": {},
            "labels": scope_config.get("labels", {}),
            "displayOrder": scope_config.get("displayOrder", []),
        }
        for key, cell_info in scope_config.get("cells", {}).items():
            col = cell_info.get("col")
            row = cell_info.get("row")
            if col and row:
                coord = f"{col}{row}"
                cell_val = cells.get(coord, {}).get("v") if coord in cells else None
                is_checked = False
                if cell_val is not None:
                    val_upper = str(cell_val).strip().upper()
                    is_checked = val_upper in ("YES", "Y", "TRUE", "1")
                scope_data["cells"][key] = {
                    "coord": coord,
                    "value": cell_val,
                    "checked": is_checked,
                }
                
    # ── Assemble result ───────────────────────────────────────
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
        "scope":          scope_data,
    }

    wb_vals.close()
    wb_fmt.close()

    _write_sheet_cache(project_id, result, sched_cache)
    print(f"[SetupCache] Cache written for {project_id}")

    return result

def write_project_setup_data(project_id, form_data, role="pm_head"):
    """Write setup form data to Excel file cells."""
    from openpyxl import load_workbook
    import os
    from datetime import datetime
    
    dept = ROLE_TO_DEPT.get(role, "PM")
    cell_map = PROJECT_SETUP_CELL_MAP.get(dept, {})
    cfg = DEPT_CONFIG.get(dept, DEPT_CONFIG["PM"])
    
    if not cell_map:
        return False, f"No setup cell mapping found for {dept}", None
    
    projects_dir, _, sched_cache, _ = get_discipline_dirs(role)
    file_path = os.path.join(projects_dir, project_id + ".xlsx")
    
    if not os.path.exists(file_path):
        file_path = os.path.join(projects_dir, project_id + ".xlsb")
        if not os.path.exists(file_path):
            return False, f"Project file not found: {project_id}", None
    
    try:
        wb = load_workbook(file_path)
        
        for field, value in form_data.items():
            # Handle scope checkboxes - convert "1" to "YES", empty to "NO"
            if field in ["scope_hw", "scope_sw", "scope_mfg", "scope_inst", "scope_com"]:
                value = "YES" if value == "1" else "NO"
            
            mapping = cell_map[field]
            sheet_name = mapping.get("sheet", "PrjSch")
            cell_ref = mapping["cell"]
            field_type = mapping.get("type", "text")
            
            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
            else:
                ws = wb.active
            
            if value is None or value == "":
                cell_value = None
            elif field_type == "number":
                try:
                    cell_value = float(value)
                except (ValueError, TypeError):
                    cell_value = 0
            elif field_type == "date":
                if value:
                    try:
                        # Parse from form's YYYY-MM-DD
                        dt = datetime.strptime(value, "%Y-%m-%d")
                        # Write as datetime object - Excel will format it
                        cell_value = dt
                    except:
                        cell_value = value
                else:
                    cell_value = None
            else:
                cell_value = str(value)
            
            ws[cell_ref] = cell_value
        
        wb.save(file_path)
        
        # Invalidate any stale cache before regenerating
        _invalidate_sheet_cache(project_id, sched_cache)

        # Generate JSON cache
        sheet_data = _generate_cache_with_pycel(project_id, file_path, sched_cache, role)
        last_modified = sheet_data.get('last_modified') if sheet_data else None
        
        # After generating JSON cache, update monitor DA column to "CONFIGURED"
        try:
            monitor_cache_path = os.path.join(get_cache_path(dept, "Monitoring"), f"{dept}_Monitor.json")
            
            if os.path.exists(monitor_cache_path):
                with open(monitor_cache_path, 'r') as f:
                    monitor_data = json.load(f)
                
                cells = monitor_data.get('cells', {})
                
                # Find the row with matching file_col
                found_row = None
                for coord, cell_info in cells.items():
                    if coord.startswith(cfg['file_col']) and cell_info.get('v') == project_id:
                        found_row = int(coord[len(cfg['file_col']):])
                        break
                
                if found_row:
                    cells[f"DA{found_row}"] = {'v': 'CONFIGURED', 'updated': True}
                    
                    with open(monitor_cache_path, 'w') as f:
                        json.dump(monitor_data, f, indent=2)
                    
                    # Queue monitor Excel write
                    queue_monitor_excel_write(dept, {f"DA{found_row}": 'CONFIGURED'})
        except Exception as e:
            print(f"[Setup] Failed to update monitor DA column: {e}")
        
        return True, "Setup data saved successfully", last_modified
        
    except Exception as e:
        print(f"[Setup] Error writing to {project_id}: {e}")
        return False, str(e), None
