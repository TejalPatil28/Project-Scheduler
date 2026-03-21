"""
ADD THIS FUNCTION TO excel_db.py
Add get_raw_sheet() anywhere after the imports.

ADD THIS ROUTE TO app.py
Add get_sheet_data() after the existing task routes.
"""

# ══════════════════════════════════════════════════════════════
# 1.  ADD TO excel_db.py  (paste after the existing imports)
# ══════════════════════════════════════════════════════════════

from openpyxl.utils import get_column_letter as _gcl

_THEME = {
    0: "#FFFFFF", 1: "#000000", 2: "#EEECE1", 3: "#DDD9C3",
    4: "#C4BD97", 5: "#938953", 6: "#494429", 7: "#1F1A0E",
    8: "#C6D9F0", 9: "#8DB3E2",
}

def _resolve(color_obj):
    if color_obj is None:
        return None
    try:
        t = color_obj.type
        if t == "rgb":
            raw = color_obj.rgb  # 8-char ARGB
            if raw in ("00000000", "FFFFFFFF"): return None
            return "#" + raw[2:]
        if t == "theme":
            return _THEME.get(color_obj.theme)
    except Exception:
        pass
    return None

def get_raw_sheet(filepath, max_col=38):
    """
    Read every cell up to `max_col` columns (default 38 = A..AL).
    Returns a dict suitable for JSON serialisation and the Excel-mirror UI.
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    max_row = ws.max_row
    max_col = min(ws.max_column, max_col)
    cols = [_gcl(i) for i in range(1, max_col + 1)]

    # ── merged cells ────────────────────────────────────────
    merged_map = {}
    for mc in ws.merged_cells.ranges:
        master_coord = f"{_gcl(mc.min_col)}{mc.min_row}"
        rs = mc.max_row - mc.min_row + 1
        cs = mc.max_col - mc.min_col + 1
        for r in range(mc.min_row, mc.max_row + 1):
            for c in range(mc.min_col, mc.max_col + 1):
                coord = f"{_gcl(c)}{r}"
                if coord == master_coord:
                    merged_map[coord] = {"master": True, "rowspan": rs, "colspan": cs}
                else:
                    merged_map[coord] = {"skip": True}

    # ── column widths (px) ──────────────────────────────────
    col_widths = {}
    for col in cols:
        cd = ws.column_dimensions.get(col)
        col_widths[col] = max(30, round((cd.width or 8) * 7.5)) if cd else 64

    # ── row heights (px) ────────────────────────────────────
    row_heights = {}
    for rn in range(1, max_row + 1):
        rd = ws.row_dimensions.get(rn)
        row_heights[str(rn)] = max(18, round((rd.height or 15) * 1.33)) if rd else 20

    # ── cells ───────────────────────────────────────────────
    cells = {}
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            coord = cell.coordinate
            mi = merged_map.get(coord, {})

            if mi.get("skip"):
                cells[coord] = {"skip": True}
                continue

            # value
            v = cell.value
            if isinstance(v, (datetime, date)):
                if isinstance(v, datetime): v = v.date()
                v = v.strftime("%d-%b-%y")
            elif isinstance(v, float):
                fmt = cell.number_format or ""
                if "%" in fmt:
                    v = f"{int(round(v * 100))}%"
            elif v is not None:
                v = str(v) if not isinstance(v, (int, str, bool)) else v

            c = {"row": cell.row, "col": cell.column, "v": v}

            # merge spans
            if mi.get("master"):
                if mi["rowspan"] > 1: c["rowspan"] = mi["rowspan"]
                if mi["colspan"] > 1: c["colspan"] = mi["colspan"]

            # font
            try:
                f = cell.font
                font = {}
                if f.bold:      font["bold"]      = True
                if f.italic:    font["italic"]    = True
                if f.underline: font["underline"] = True
                if f.size and f.size != 11: font["size"] = f.size
                fc = _resolve(f.color)
                if fc and fc not in ("#FFFFFF", "#000000"): font["color"] = fc
                if font: c["font"] = font
            except Exception:
                pass

            # fill
            try:
                fill = cell.fill
                if fill and fill.patternType not in (None, "none"):
                    fg = _resolve(fill.fgColor)
                    if fg and fg != "#FFFFFF": c["fill"] = fg
            except Exception:
                pass

            # alignment
            try:
                a = cell.alignment
                if a:
                    al = {}
                    if a.horizontal and a.horizontal not in ("general", None): al["h"] = a.horizontal
                    if a.vertical   and a.vertical   not in ("bottom",  None): al["v"] = a.vertical
                    if a.wrap_text: al["wrap"] = True
                    if al: c["align"] = al
            except Exception:
                pass

            # borders
            try:
                b = cell.border
                if b:
                    bd = {}
                    for side in ("left", "right", "top", "bottom"):
                        s = getattr(b, side)
                        if s and s.border_style and s.border_style != "none":
                            bd[side] = s.border_style
                    if bd: c["border"] = bd
            except Exception:
                pass

            cells[coord] = c

    return {
        "cells":       cells,
        "col_widths":  col_widths,
        "row_heights": row_heights,
        "max_row":     max_row,
        "max_col":     max_col,
        "cols":        cols,
    }


# ══════════════════════════════════════════════════════════════
# 2.  ADD TO app.py  (paste after the existing task routes)
# ══════════════════════════════════════════════════════════════

@app.route("/api/projects/<project_id>/sheet", methods=["GET"])
@login_required
def get_sheet_data(project_id):
    """Return raw cell data for the Excel-mirror UI."""
    import os
    from excel_db import get_raw_sheet, PROJECTS_DIR
    fpath = os.path.join(PROJECTS_DIR, project_id + ".xlsx")
    if not os.path.exists(fpath):
        return jsonify({"error": "Project file not found"}), 404
    try:
        data = get_raw_sheet(fpath)
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
