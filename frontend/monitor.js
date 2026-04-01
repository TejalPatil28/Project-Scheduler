// ─────────────────────────────────────────────────────────────
//  monitor.js  —  Tabbed Monitor Screen
//  Called from app.js:  renderMonitor(container, data, user)
//  Data format:  { cells: { "F10": {v:"..."}, "F12": {v:"..."}, ... }, ... }
//  Row 10  = column headers
//  Row 11  = ignored
//  Row 12+ = project data rows
//
//  Editable cells:
//    Dropdowns : N  → [SYSTEM, SERVICE, E & C]
//                AP → [RUN, HOLD, CLOSED]
//                C  → sw_head / head / admin sees dropdown of all SW_TL values
//                D  → sw_tl sees dropdown of all SW_TL values (to assign)
//    Free-text : AR → inline text input (read + write)
//    Progress  : Milestone cols BB–BW → 0–100 number input + colour bar
//
//  Save pattern: define window.__monitorSaveCell(col, row, value)
//    in your app.js.  monitor.js calls it on every change.
// ─────────────────────────────────────────────────────────────

(function (global) {
  "use strict";

  // ── Constants ───────────────────────────────────────────────
  var HEADER_ROW = 10;
  var DATA_START = 12;
  var ID_COL     = "F";

  // ═══════════════════════════════════════════════════════════
  //  EDITABLE COLUMN CONFIG
  //  ─ Add / remove options or columns here only.
  // ═══════════════════════════════════════════════════════════

  // Dropdown columns → their option lists (first entry = blank/reset)
  var DROPDOWN_OPTIONS = {
    "N":  ["", "SYSTEM", "SERVICE", "E & C"],
    "AP": ["", "RUN", "HOLD", "CLOSED"],
  };

  // Free-text editable columns
  var TEXT_EDIT_COLS = ["AR"];

  // Milestone progress columns (0–100 bar)
  var MILESTONE_COLS = [
    "BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK",
    "BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV","BW",
  ];

  // ── Column group definitions ─────────────────────────────────
  var COL_GROUPS = [
    { key: "info1",       label: "Project Info",      cols: ["C","D"] },
    { key: "info",        label: "Project Info",      cols: ["H","I","J","K","L","M"] },
    { key: "ld",          label: "LD Details",        cols: ["P","Q","R","S","T"] },
    { key: "ld1",         label: "LD Details",        cols: ["U","V"] },
    { key: "ld2",         label: "LD Details",        cols: ["W","X","Y"] },
    { key: "additional",  label: "Additional Fields", cols: ["AJ","AK","AL","AM"] },
    { key: "additional1", label: "Additional Fields", cols: ["AN","AO","AP"] },
    { key: "additional2", label: "Additional Fields", cols: ["AQ","AR"] },
  ];

  // ── Column width definitions ─────────────────────────────────
  // TO CHANGE: find the tab key, edit the px value.
  // "_default" = fallback for any unlisted column.
  var TAB_WIDTHS = {
    overview: {
      _default: "90px",
      "C": "80px",  "D": "80px",  "E": "80px",
      "F": "40px",  "G": "80px",  "H": "50px",
      "I": "90px",  "J": "90px",  "K": "90px",  "L": "90px",
      "M": "80px",  "N": "90px",  "O": "50px",
      "P": "90px",  "Q": "90px",  "R": "90px",  "S": "90px",
      "T": "70px", "U": "100px", "V": "70px", "W": "100px",
      "X": "100px", "Y": "70px",
      "AG": "90px", "AH": "50px", "AI": "40px", "AJ": "90px",
      "AK": "90px", "AL": "90px", "AM": "90px", "AN": "90px", "AO": "90px",
      "AP": "90px", "AQ": "100px","AR": "120px","BA": "90px",
    },
    // % widths for overview — used always, table expands beyond 100% when groups open
    overview_pct: {
      _default: "5%",
      "F": "15%",
      "C": "4%", "D": "4%", "E": "6%", "G": "5%",
      "H": "4%", "I": "5%", "J": "5%", "K": "5%", "L": "5%",
      "M": "5%", "N": "6%", "O": "5%",
      "P": "5%", "Q": "5%", "R": "5%", "S": "5%", "T": "5%",
      "U": "5%", "V": "5%", "W": "5%", "X": "5%", "Y": "5%",
      "AG": "5%", "AH": "4%", "AI": "4%",
      "AJ": "5%", "AK": "5%", "AL": "5%", "AM": "6%",
      "AN": "5%", "AO": "5%", "AP": "6%",
      "AQ": "5%", "AR": "6%",
      "BA": "7%"
    },
    team: {
      _default: "9.44%",
      "F": "15%", "C": "9.44%", "D": "9.44%",
      "Z": "9.44%","AA": "9.44%","AB": "9.44%","AC": "9.44%",
      "AD": "9.44%","AE": "9.44%","AF": "9.44%",
    },
    milestone: {
      _default: "3.86%",
      "F": "15%",
    },
  };

  // ── Tab definitions ─────────────────────────────────────────
  var TABS = [
    {
      id:        "overview",
      label:     "Project Overview",
      icon:      "🗂",
      scroll:    true,
      colGroups: COL_GROUPS,
      cols: ["C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","BA"],
    },
    {
      id:    "team",
      label: "Team & Roles",
      icon:  "👥",
      fluid: true,
      cols:  ["F","C","D","Z","AA","AB","AC","AD","AE","AF"],
    },
    {
      id:     "milestone",
      label:  "Milestone Progress",
      icon:   "🏁",
      fluid:  true,
      cols:   ["F","BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK","BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV","BW"],
    },
  ];

  // ── Helpers ──────────────────────────────────────────────────
  function cv(cells, col, row) {
    var key  = col + row;
    var info = cells[key];
    if (!info || info.v === null || info.v === undefined) return "";
    return String(info.v);
  }

  function esc(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  // ── Save helper ───────────────────────────────────────────────
  function saveCell(col, row, value) {
    if (typeof window.__monitorSaveCell === "function") {
      window.__monitorSaveCell(col, row, value);
    } else {
      console.warn("[monitor] window.__monitorSaveCell not defined — col=" + col + " row=" + row + " val=" + value);
    }
  }

  // ── Build header map from row 10 ─────────────────────────────
  function buildHeaderMap(cells) {
    var map = {};
    var allCols = [];
    TABS.forEach(function (tab) {
      tab.cols.forEach(function (col) {
        if (allCols.indexOf(col) === -1) allCols.push(col);
      });
    });
    allCols.forEach(function (col) {
      map[col] = cv(cells, col, HEADER_ROW) || col;
    });
    return map;
  }

  // ── Extract visible data rows (role-filtered) ─────────────────
  function extractRows(cells, maxRow, user) {
    var role     = user && user.role;
    var initials = user && user.short_name ? user.short_name.trim().toUpperCase() : "";
    var isHead   = (role === "head" || role === "admin");

    var rows = [];
    for (var r = DATA_START; r <= maxRow; r++) {
      var id = cv(cells, ID_COL, r);
      if (!id || id.trim() === "") continue;
      rows.push(r);
    }
    if (isHead) return rows;

    var swhRows = rows.filter(function (r) {
      return cv(cells, "C", r).trim().toUpperCase() === initials;
    });
    if (swhRows.length > 0) return swhRows;

    return rows.filter(function (r) {
      return cv(cells, "D", r).trim().toUpperCase() === initials;
    });
  }

  // ── Collect all unique SW_TL values from col D (all rows) ────
  // Used to populate the C and D dropdowns.
  function collectTLValues(cells, allRows) {
    var seen = {};
    var vals = [];
    allRows.forEach(function (r) {
      var v = cv(cells, "D", r).trim();
      if (v && !seen[v]) { seen[v] = true; vals.push(v); }
    });
    return vals.sort();
  }

  // ── col → groupIndex lookup ──────────────────────────────────
  function buildColToGroup(colGroups) {
    var map = {};
    if (!colGroups) return map;
    colGroups.forEach(function (grp, gi) {
      grp.cols.forEach(function (col) { map[col] = gi; });
    });
    return map;
  }

  // ── Render one tab's full table HTML ─────────────────────────
  function renderTabHTML(cells, rows, allRows, tab, headerMap, collapseState, user) {
    var cols        = tab.cols;
    var scroll      = tab.scroll;
    var colGroups   = tab.colGroups || null;
    var colToGroup  = buildColToGroup(colGroups);
    var role        = (user && user.role) || "";
    var isMilestone = (tab.id === "milestone");

    // TL list for C / D dropdowns — derived from ALL rows, not just visible
    var tlValues = collectTLValues(cells, allRows);

    // ── Width helper ──────────────────────────────────────────
    // For overview: use % widths only when all groups are collapsed.
    // If any group is expanded, fall back to px widths (horizontal scroll).
    // Overview always uses % widths. When groups expand, table width grows beyond 100%
    // so new columns push to the right without shifting existing columns.
    var isOverview = (tab.id === "overview");
    var widthMap = isOverview
      ? (TAB_WIDTHS["overview_pct"] || {})
      : (TAB_WIDTHS[tab.id] || {});
    function colWidth(col) {
      return widthMap[col] || widthMap["_default"] || (fluid ? "auto" : "40px");
    }

    // Count total % width of all currently visible columns to set table width
    // so expanding a group pushes right rather than squeezing existing columns.
    function computeOverviewTableWidth() {
      var total = 0;
      cols.forEach(function(col) {
        if (hidden(col)) return;
        var w = colWidth(col);
        var val = parseFloat(w);
        if (!isNaN(val) && w.indexOf("%") !== -1) total += val;
      });
      total += 2; // account for the fixed 36px # col (~2%)
      return total <= 100 ? "100%" : total + "%";
    }

    function hidden(col) {
      if (!colGroups) return false;
      var gi = colToGroup[col];
      if (gi === undefined) return false;
      if (!collapseState[gi]) return false;
      var lastCol = colGroups[gi].cols[colGroups[gi].cols.length - 1];
      return col !== lastCol;
    }

    // ── Color rules ───────────────────────────────────────────
    var isDark = document.documentElement.getAttribute("data-theme") !== "light";

    var colorRules = isDark ? [
      { columns: ["H","I","J","K","L"], color: "#2a2a2a" },
      { columns: ["P","Q","R","S"],     color: "#152030" },
      { columns: ["M","N","O"],         color: "#1f2a2a" },
    ] : [
      { columns: ["M","G","K","L","F","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AQ","BA"], color: "#747070" },
      { columns: ["E","N","I","J","C","D","AR","AP"], color: "#90b4df" },
      { columns: ["H","BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK","BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV","BW"], color: "#000000" },
    ];

    var textColorRules = isDark ? [
      { columns: ["S"], color: "#ff9999" },
      { columns: ["R"], color: "#66ccff" },
      { columns: ["M"], color: "#ffcc66" },
    ] : [
      { columns: ["F","H","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AQ"], color: "#ffba26" },
      { columns: ["M","G","BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK","BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV","BW"], color: "#ffffff" },
    ];

    function getCellBg(col) {
      for (var i = 0; i < colorRules.length; i++) {
        if (colorRules[i].columns.indexOf(col) !== -1) return colorRules[i].color;
      }
      return null;
    }
    function getCellFg(col) {
      for (var i = 0; i < textColorRules.length; i++) {
        if (textColorRules[i].columns.indexOf(col) !== -1) return textColorRules[i].color;
      }
      return null;
    }

    // ── Button style ──────────────────────────────────────────
    var btnStyle = [
      "cursor:pointer","font-size:11px","font-weight:700",
      "padding:1px 7px","border-radius:3px",
      "border:1px solid " + (isDark ? "#555555" : "#aaaaaa"),
      "background:" + (isDark ? "#2a2a2a" : "#e8e8e8"),
      "color:" + (isDark ? "#dddddd" : "#333333"),
      "line-height:1.6",
    ].join(";");

    // ── Base style shared by all editable inputs/selects ──────
    var inputBase = [
      "width:100%","box-sizing:border-box",
      "border:none","outline:none","background:transparent",
      "font-family:var(--font-body)","font-size:12px",
      "color:inherit","padding:0",
    ].join(";");

    var fluid     = tab.fluid;

    // ── Scroll wrapper ────────────────────────────────────────
    var wrapStyle = [
      "flex:1","min-height:0","overflow-y:auto",
      (isOverview || scroll && !fluid) ? "overflow-x:auto" : "overflow-x:hidden",
    ].join(";");

    // Overview: always fixed layout with dynamic % width — expands right when groups open
    var overviewTableWidth = isOverview ? computeOverviewTableWidth() : null;
    var tableClass = "monitor-table"
      + ((!isOverview && (scroll && !fluid)) ? " monitor-table-scroll" : "")
      + (tab.id === "milestone" ? " monitor-table-milestone" : "")
      + (isOverview ? " monitor-table-pct" : "");
    var tableStyle = (fluid || isOverview)
      ? "table-layout:fixed;width:" + (isOverview ? overviewTableWidth : "100%") + ";"
      : "";
    
    var html = '<div class="monitor-table-wrap" style="' + wrapStyle + '">';
    html += '<table class="' + tableClass + '"' + (tableStyle ? ' style="' + tableStyle + '"' : "") + '>';
    
    // ── <colgroup> ────────────────────────────────────────────
    // ── <colgroup> ────────────────────────────────────────────
    html += "<colgroup>";
    html += '<col style="width:36px;">';
    cols.forEach(function (col) {
      if (hidden(col)) return;
      var w = colWidth(col);
      html += '<col style="width:' + w + ';">';
    });
    html += "</colgroup><thead>";

    // ── ROW 1: Group button row ───────────────────────────────
    if (colGroups && colGroups.length) {
      html += '<tr class="mgrp-btn-row" style="height:24px;">';
      html += '<th class="monitor-th monitor-th-num" rowspan="2">#</th>';

      var ci = 0;
      while (ci < cols.length) {
        var col = cols[ci];
        var gi  = colToGroup[col];

        if (gi !== undefined && col === colGroups[gi].cols[0]) {
          var grp       = colGroups[gi];
          var collapsed = collapseState[gi];
          var btnLabel  = collapsed ? "+" : "–";
          var onclickFn = "window.__monitorToggleGroup_" + esc(tab.id) + "(" + gi + ")";

          if (!collapsed) {
            // Expanded: span all columns in the group so button stays above its columns
            html += '<th colspan="' + grp.cols.length + '" style="position:sticky;top:0;z-index:5;background:rgba(0,0,0,0.3);border:none;text-align:left;padding:2px 0 0 4px;box-sizing:border-box;border-radius:4px 4px 0 0;">'
                  + '<button style="' + btnStyle + '" onclick="' + onclickFn + '">' + btnLabel + '</button></th>';
            ci += grp.cols.length;
          } else {
            // Collapsed: single cell (only last col visible)
            html += '<th style="position:sticky;top:0;z-index:5;background:transparent;border:none;text-align:left;padding:2px 0 0 4px;">'
                  + '<button style="' + btnStyle + '" onclick="' + onclickFn + '">' + btnLabel + '</button></th>';
            ci += grp.cols.length;
          }
        } else if (gi !== undefined) {
          ci++;
        } else {
          html += '<th style="position:sticky;top:0;z-index:5;background:transparent;border:none;"></th>';
          ci++;
        }
      }
      html += '</tr>';
    }

    // ── ROW 2: Column name row ────────────────────────────────
    html += '<tr>';
    if (!colGroups || !colGroups.length) {
      html += '<th class="monitor-th monitor-th-num">#</th>';
    }
    cols.forEach(function (col) {
      if (hidden(col)) return;
      var gi = colToGroup[col];
      var isFirst = (gi !== undefined && colGroups && colGroups[gi].cols[0] === col);
      var extraStyle = isFirst ? "border-left:2px solid var(--accent);" : "";
      var w = colWidth(col);
      extraStyle += "width:" + w + " !important;max-width:" + w + " !important;";
      var headerText = headerMap[col] || col;
      html += '<th class="monitor-th" style="' + extraStyle + '" title="' + esc(headerText) + '">' + esc(headerText) + '</th>';
    });
    html += '</tr></thead><tbody>';

    // ── TBODY ─────────────────────────────────────────────────
    var visibleCols = cols.filter(function (c) { return !hidden(c); });

    if (rows.length === 0) {
      html += '<tr><td colspan="' + (visibleCols.length + 1) + '" class="monitor-empty-row">No data found</td></tr>';
    } else {
      rows.forEach(function (r, idx) {
        var rowCls = "monitor-tr" + (idx % 2 !== 0 ? " monitor-tr-alt" : "");
        html += '<tr class="' + rowCls + '">';
        html += '<td class="monitor-td monitor-td-num">' + (idx + 1) + '</td>';

        visibleCols.forEach(function (col) {
          var val       = cv(cells, col, r);
          var bgColor   = getCellBg(col);
          var textColor = getCellFg(col);
          var style = "";
          if (bgColor)   style += "background-color:" + bgColor + ";";
          if (textColor) style += "color:" + textColor + ";";

          var gi2 = colToGroup[col];
          if (gi2 !== undefined && colGroups && colGroups[gi2].cols[0] === col) {
            style += "border-left:2px solid var(--accent);";
          }

          var styleAttr = style ? ' style="' + style + '"' : "";
          var cellContent = "";

          // ── 1. MILESTONE PROGRESS BAR (BB–BW) — DISPLAY ONLY ──
          // ── 1. MILESTONE PROGRESS (BB–BW) — DISPLAY ONLY ──
          if (isMilestone && MILESTONE_COLS.indexOf(col) !== -1) {
            var raw = parseFloat(val);
            // Auto-detect scale: if stored as 0–1 (e.g. 0.75), multiply to get %
            // If stored as 0–100 (e.g. 75), use directly.
            var pct = isNaN(raw) ? 0 : (raw <= 1 && raw > 0 ? Math.round(raw * 100) : Math.round(raw));
            if (pct < 0)   pct = 0;
            if (pct > 100) pct = 100;
            cellContent = pct + '%';

          // ── 2. DROPDOWN: N and AP ───────────────────────────
          } else if (DROPDOWN_OPTIONS[col]) {
            var opts = DROPDOWN_OPTIONS[col];
            var ocDD = "window.__monitorSaveCellEvt('" + col + "'," + r + ",this.value)";
            cellContent = '<select style="' + inputBase + ';cursor:pointer;" onchange="' + ocDD + '" onclick="event.stopPropagation()">';
            // If current value exists but isn't in the list, show it first so it's not lost
            var valInOpts = opts.indexOf(val) !== -1;
            if (val && !valInOpts) {
              cellContent += '<option value="' + esc(val) + '" selected>' + esc(val) + '</option>';
            }
            opts.forEach(function (opt) {
              var sel = (opt === val) ? " selected" : "";
              cellContent += '<option value="' + esc(opt) + '"' + sel + '>' + esc(opt || "—") + '</option>';
            });
            cellContent += '</select>';

          // ── 3. DROPDOWN C: visible to sw_head / head / admin ─
          //       Shows all SW_TL values from col D
          } else if (col === "C" && (role === "sw_head" || role === "head" || role === "admin")) {
            var ocC = "window.__monitorSaveCellEvt('C'," + r + ",this.value)";
            cellContent = '<select style="' + inputBase + ';cursor:pointer;" onchange="' + ocC + '" onclick="event.stopPropagation()">';
            cellContent += '<option value="">—</option>';
            // If current value exists but isn't in tlValues, show it so it's not lost
            if (val && tlValues.indexOf(val) === -1) {
              cellContent += '<option value="' + esc(val) + '" selected>' + esc(val) + '</option>';
            }
            tlValues.forEach(function (v) {
              cellContent += '<option value="' + esc(v) + '"' + (v === val ? " selected" : "") + '>' + esc(v) + '</option>';
            });
            cellContent += '</select>';

          // ── 4. DROPDOWN D: visible to sw_tl ─────────────────
          //       Shows all SW_TL values so a TL can assign
          //       another TL (or themselves) as the sub-TL
          } else if (col === "D" && role === "sw_tl") {
            var ocD = "window.__monitorSaveCellEvt('D'," + r + ",this.value)";
            cellContent = '<select style="' + inputBase + ';cursor:pointer;" onchange="' + ocD + '" onclick="event.stopPropagation()">';
            cellContent += '<option value="">—</option>';
            // If current value exists but isn't in tlValues, show it so it's not lost
            if (val && tlValues.indexOf(val) === -1) {
              cellContent += '<option value="' + esc(val) + '" selected>' + esc(val) + '</option>';
            }
            tlValues.forEach(function (v) {
              cellContent += '<option value="' + esc(v) + '"' + (v === val ? " selected" : "") + '>' + esc(v) + '</option>';
            });
            cellContent += '</select>';

          // ── 5. FREE TEXT: AR ─────────────────────────────────
          } else if (TEXT_EDIT_COLS.indexOf(col) !== -1) {
            var ocTxt = "window.__monitorSaveCellEvt('" + col + "'," + r + ",this.value)";
            cellContent = '<input type="text" value="' + esc(val) + '" '
              + 'style="' + inputBase + ';cursor:text;" '
              + 'onchange="' + ocTxt + '" '
              + 'onclick="event.stopPropagation()" '
              + 'placeholder="—" />';

          // ── READ-ONLY ─────────────────────────────────────────
          } else {
            cellContent = esc(val);
          }

          html += '<td class="monitor-td"' + styleAttr + ' title="' + esc(val) + '">' + cellContent + '</td>';
        });

        html += "</tr>";
      });
    }

    html += "</tbody></table></div>";
    return html;
  }

  // ── Main render function ─────────────────────────────────────
  function renderMonitor(container, data, user) {
    var cells  = data.cells   || {};
    var maxRow = data.max_row || 0;

    var headerMap = buildHeaderMap(cells);
    var rows      = extractRows(cells, maxRow, user);
    var activeTab = 0;

    // allRows = unfiltered rows — needed to build full TL dropdown list
    var allRows = [];
    for (var r = DATA_START; r <= maxRow; r++) {
      if (cv(cells, ID_COL, r).trim() !== "") allRows.push(r);
    }

    container.innerHTML = "";
    container.className = "monitor-shell";

    // ── collapseState ─────────────────────────────────────────
    var collapseStates = {};
    TABS.forEach(function (tab) {
      if (tab.colGroups && tab.colGroups.length) {
        collapseStates[tab.id] = tab.colGroups.map(function () { return true; });
      }
    });

    // ── Global cell-save event handler ────────────────────────
    // Inline onchange attributes call this.
    // Updates the in-memory cells so re-renders reflect the change,
    // then calls the host app's persistence function.
    window.__monitorSaveCellEvt = function (col, row, value) {
      var key = col + row;
      if (!cells[key]) cells[key] = {};
      cells[key].v = value;
      saveCell(col, row, value);
      // Re-render the active panel to update bars / dependent cells
      var activePanel  = panelEls[activeTab];
      var activeTabDef = TABS[activeTab];
      if (activePanel && activeTabDef) {
        activePanel.innerHTML = renderTabHTML(
          cells, rows, allRows, activeTabDef,
          headerMap, collapseStates[activeTabDef.id] || [], user
        );
      }
    };

    // ── Build shell DOM ───────────────────────────────────────
    var tabBar = document.createElement("div");
    tabBar.className = "monitor-tabbar";

    var panelsWrap = document.createElement("div");
    panelsWrap.className = "monitor-panels";

    var badge = document.createElement("div");
    badge.className = "monitor-badge";
    badge.textContent = rows.length + " project" + (rows.length !== 1 ? "s" : "");

    var topRow = document.createElement("div");
    topRow.className = "monitor-toprow";
    topRow.appendChild(tabBar);
    topRow.appendChild(badge);

    container.appendChild(topRow);
    container.appendChild(panelsWrap);

    // ── Create panels & toggle handlers ──────────────────────
    var panelEls = [];

    TABS.forEach(function (tab, idx) {
      var tbtn = document.createElement("button");
      tbtn.className = "monitor-tab-btn" + (idx === activeTab ? " active" : "");
      tbtn.setAttribute("data-tab", idx);
      tbtn.innerHTML =
        '<span class="monitor-tab-icon">' + tab.icon + "</span>" +
        '<span class="monitor-tab-label">' + esc(tab.label) + "</span>";
      tbtn.onclick = function () { switchTab(idx); };
      tabBar.appendChild(tbtn);

      var panel = document.createElement("div");
      panel.className = "monitor-panel" + (idx === activeTab ? " active" : "");
      panel.setAttribute("data-panel", idx);
      panelEls.push(panel);
      panelsWrap.appendChild(panel);

      if (tab.colGroups && tab.colGroups.length) {
        (function (t, p) {
          window["__monitorToggleGroup_" + t.id] = function (gi) {
            collapseStates[t.id][gi] = !collapseStates[t.id][gi];
            p.innerHTML = renderTabHTML(cells, rows, allRows, t, headerMap, collapseStates[t.id], user);
          };
        })(tab, panel);
      }

      panel.innerHTML = renderTabHTML(cells, rows, allRows, tab, headerMap, collapseStates[tab.id] || [], user);
    });

    // ── Tab switching ─────────────────────────────────────────
    function switchTab(idx) {
      activeTab = idx;
      tabBar.querySelectorAll(".monitor-tab-btn").forEach(function (b, i) {
        b.classList.toggle("active", i === idx);
      });
      panelEls.forEach(function (p, i) {
        p.classList.toggle("active", i === idx);
      });
    }
  }

  global.renderMonitor = renderMonitor;

})(window);
