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
  var ID_COL     = "B";

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

  // Date picker columns
  var DATE_EDIT_COLS = ["E"];

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
      "H": "4%", "I": "8%", "J": "5%", "K": "10%", "L": "10%",
      "M": "5%", "N": "6%", "O": "5%",
      "P": "5%", "Q": "5%", "R": "5%", "S": "10%", "T": "5%",
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

  // In monitor.js, modify the saveCell function or add an API call

  function saveMonitorCell(col, row, value) {
      // Call backend API to save
      fetch('/api/monitor/cell', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ col: col, row: row, value: value }),
          credentials: 'include'
      })
      .then(response => response.json())
      .then(data => {
          if (data.error) {
              console.error('Save failed:', data.error);
          } else {
              console.log('Saved:', data);
          }
      })
      .catch(err => console.error('Error saving monitor cell:', err));
  }

  // ── Save helper ───────────────────────────────────────────────
  function saveCell(col, row, value) {
      // Check if this is monitor context (no project_id)
      if (typeof window.__monitorSaveCell === "function") {
          window.__monitorSaveCell(col, row, value);
      } else {
          // Fallback to direct API call
          fetch('/api/monitor/cell', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ col: col, row: row, value: value }),
              credentials: 'include'
          }).catch(err => console.error('Save failed:', err));
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
    var isHead = (role === "head" || role === "admin" || role.endsWith("_head"));

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

    // ── Collect all unique SW_TL values from monitor data column D ────
  function collectMonitorTLValues(cells, allRows) {
    var seen = {};
    var vals = [];
    allRows.forEach(function (r) {
      var v = cv(cells, "D", r).trim();
      if (v && !seen[v]) { seen[v] = true; vals.push(v); }
    });
    return vals.sort();
  }

  // ── Calculate column sums for monitor ──
function calculateColumnSums(cells, rows, cols) {
  var sums = {
    "M": 0,
    "Y": 0,
    "AH": 0,
    "AI": 0,
    "AJ":0,
    "R":0
  };
  
  rows.forEach(function(row) {
    // For each column we want to sum
    for (var col in sums) {
      var val = cv(cells, col, row);
      var num = parseFloat(val);
      if (!isNaN(num)) {
        sums[col] += num;
      }
    }
  });
  
  return sums;
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
  function renderTabHTML(cells, rows, allRows, tab, headerMap, collapseState, user, overdueMap) {
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
      { columns: ["H"], color: "#191616" },
    ];

    var textColorRules = isDark ? [
      { columns: ["S"], color: "#ff9999" },
      { columns: ["R"], color: "#66ccff" },
      { columns: ["M"], color: "#ffcc66" },
    ] : [
      { columns: ["F","H","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AQ"], color: "#ffba26" },
      { columns: ["M","G"], color: "#ffffff" },
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
      html += '<th class="monitor-th monitor-th-num" rowspan="3">#</th>';

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


  // ========== ADD THIS NEW ROW HERE ==========
    // ── ROW 1.5: Sum row (only for overview tab) ──────────────
    var isOverview = (tab.id === "overview");
    if (isOverview) {
      // Calculate sums for columns M, Y, AH, AI
      var sums = calculateColumnSums(cells, rows, cols);
      
      // Start the sum row
      html += '<tr style="background:var(--bg4);">';
      
      
      // Loop through each column to add sum values
      cols.forEach(function(col) {
        if (hidden(col)) return;
        
        // Check if this column should show a sum
        if (col === "M" || col === "Y" || col === "AH" || col === "AI" || col === "AJ" || col === "R") {
          var sumValue = sums[col] || 0;
          var bgColor = isDark ? "#4a4a4a" : "#818181";
          var textColor = "#ffffff";
          var borderColor = isDark ? "#666666" : "#555555";
          html += '<th class="monitor-th" style="background:' + bgColor + ';color:' + textColor + ';font-weight:700;text-align:center;border-bottom:2px solid ' + borderColor + ';">' + sumValue.toFixed(2) + '</th>';
        } else {
          // Empty cell for non-sum columns
          html += '<th class="monitor-th" style="background:var(--bg4);"></th>';
        }
      });
      
      html += '</tr>';
    }
    // ========== END OF SUM ROW ==========

    // ── ROW 2: Column name row ────────────────────────────────
    html += '<tr>';
    
    // Overview already has # with rowspan="3" from group button row — don't add again
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
      var wrapStyle = isMilestone ? "white-space:normal;word-break:break-word;line-height:1.2;padding:4px 2px;text-align:center;font-size:9px;" : "";
      html += '<th class="monitor-th" style="' + extraStyle + wrapStyle + '" title="' + esc(headerText) + '">' + esc(headerText) + '</th>';
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

          // ── MONITOR SCREEN EDITABLE COLUMNS (Role-based) ──
          // This block runs for all tabs EXCEPT milestone (overview, team)
          var userRole = (user && user.role) || "";
          var isMonitorTab = (tab.id !== "milestone");

          if (isMonitorTab) {
            
            // ── COLUMN C: SW Head (editable by HEAD only) ──
            if (col === "C" && userRole === "sw_head") {
              var tlValues = collectMonitorTLValues(cells, allRows);
              var ocC = "window.__monitorSaveCellEvt('" + col + "'," + r + ",this.value)";
              cellContent = '<select style="' + inputBase + ';cursor:pointer;" onchange="' + ocC + '" onclick="event.stopPropagation()">';
              cellContent += '<option value="">—</option>';
              if (val && tlValues.indexOf(val) === -1) {
                cellContent += '<option value="' + esc(val) + '" selected>' + esc(val) + '</option>';
              }
              tlValues.forEach(function (v) {
                cellContent += '<option value="' + esc(v) + '"' + (v === val ? " selected" : "") + '>' + esc(v) + '</option>';
              });
              cellContent += '</select>';
            }
            
            // ── COLUMN D: SWE Name (editable by SW_TL only) ──
            else if (col === "D" && userRole === "sw_tl") {
              var tlValues2 = collectMonitorTLValues(cells, allRows);
              var ocD = "window.__monitorSaveCellEvt('" + col + "'," + r + ",this.value)";
              cellContent = '<select style="' + inputBase + ';cursor:pointer;" onchange="' + ocD + '" onclick="event.stopPropagation()">';
              cellContent += '<option value="">—</option>';
              if (val && tlValues2.indexOf(val) === -1) {
                cellContent += '<option value="' + esc(val) + '" selected>' + esc(val) + '</option>';
              }
              tlValues2.forEach(function (v) {
                cellContent += '<option value="' + esc(v) + '"' + (v === val ? " selected" : "") + '>' + esc(v) + '</option>';
              });
              cellContent += '</select>';
            }
            
            // ── COLUMN E: Date field (editable by SW_TL only) ──
            else if (col === "E" && userRole === "sw_tl") {
              var dateVal = val || "";
              var dateInputVal = "";
              if (dateVal) {
                try {
                  var d = new Date(dateVal);
                  if (!isNaN(d)) dateInputVal = d.toISOString().split("T")[0];
                } catch(e) {}
              }
              var ocDate = "window.__monitorSaveCellEvt('" + col + "'," + r + ",this.value)";
              cellContent = '<input type="date" style="' + inputBase + ';cursor:text;" value="' + dateInputVal + '" onchange="' + ocDate + '" onclick="event.stopPropagation()" />';
            }
            
            // ── COLUMN N: Project Type dropdown (editable by SW_TL only) ──
            else if (col === "N" && userRole === "sw_tl") {
              var optsN = ["", "SYSTEM", "SERVICE", "E & C"];
              var ocN = "window.__monitorSaveCellEvt('" + col + "'," + r + ",this.value)";
              cellContent = '<select style="' + inputBase + ';cursor:pointer;" onchange="' + ocN + '" onclick="event.stopPropagation()">';
              var valInOpts = optsN.indexOf(val) !== -1;
              if (val && !valInOpts) {
                cellContent += '<option value="' + esc(val) + '" selected>' + esc(val) + '</option>';
              }
              optsN.forEach(function (opt) {
                var sel = (opt === val) ? " selected" : "";
                cellContent += '<option value="' + esc(opt) + '"' + sel + '>' + esc(opt || "—") + '</option>';
              });
              cellContent += '</select>';
            }
            
            // ── COLUMN AP: Status dropdown (editable by SW_TL only) ──
            else if (col === "AP" && userRole === "sw_tl") {
              var optsAP = ["", "RUN", "HOLD", "CLOSED"];
              var ocAP = "window.__monitorSaveCellEvt('" + col + "'," + r + ",this.value)";
              cellContent = '<select style="' + inputBase + ';cursor:pointer;" onchange="' + ocAP + '" onclick="event.stopPropagation()">';
              var valInOptsAP = optsAP.indexOf(val) !== -1;
              if (val && !valInOptsAP) {
                cellContent += '<option value="' + esc(val) + '" selected>' + esc(val) + '</option>';
              }
              optsAP.forEach(function (opt) {
                var sel = (opt === val) ? " selected" : "";
                cellContent += '<option value="' + esc(opt) + '"' + sel + '>' + esc(opt || "—") + '</option>';
              });
              cellContent += '</select>';
            }
            
            // ── COLUMN AR: Remarks text field (editable by SW_TL only) ──
            else if (col === "AR" && userRole === "sw_tl") {
              var ocAR = "window.__monitorSaveCellEvt('" + col + "'," + r + ",this.value)";
              cellContent = '<input type="text" value="' + esc(val) + '" '
                + 'style="' + inputBase + ';cursor:text;" '
                + 'onchange="' + ocAR + '" '
                + 'onclick="event.stopPropagation()" '
                + 'placeholder="—" />';
            }
            
            // ── For all other columns in monitor tab, show formatted text ──
            else {
              var displayValue = val;
              
              // Column M: 2 decimal places
              if (col === "M") {
                var numM = parseFloat(val);
                if (!isNaN(numM)) {
                  displayValue = numM.toFixed(2);
                }
              }
              // Columns P and Q: convert decimal to percentage
              else if (col === "P" || col === "Q") {
                var numPQ = parseFloat(val);
                if (!isNaN(numPQ)) {
                  displayValue = Math.round(numPQ * 100) + '%';
                }
              }
              // Columns R and AJ: 2 decimal places
              else if (col === "R" || col === "AJ") {
                var numRAJ = parseFloat(val);
                if (!isNaN(numRAJ)) {
                  displayValue = numRAJ.toFixed(2);
                }
              }
              // Columns AK and AL: convert decimal to percentage
              else if (col === "AK" || col === "AL" || col === "AO") {
                var numAKAL = parseFloat(val);
                if (!isNaN(numAKAL)) {
                  displayValue = Math.round(numAKAL * 100) + '%';
                }
              }
              
              cellContent = esc(displayValue);
            }
            
          } 
          // ── END OF MONITOR EDITABLE BLOCK ──

          // ── 1. MILESTONE PROGRESS BAR (BB–BW) — DISPLAY ONLY ──
          // ── 1. MILESTONE PROGRESS (BB–BW) — DISPLAY ONLY ──
          else if (isMilestone && MILESTONE_COLS.indexOf(col) !== -1) {
            var raw = parseFloat(val);
            var pct = isNaN(raw) ? 0 : (raw <= 1 && raw > 0 ? Math.round(raw * 100) : Math.round(raw));
            if (pct < 0) pct = 0;
            if (pct > 100) pct = 100;
            
            var taskName = (headerMap[col] || "").trim();
            var projectId = cv(cells, "B", r);
            var userRole = (user && user.role) || "";
            var textColor = "#ffffff";
            var bgColor = "#000000";
            
            if (overdueMap && projectId && overdueMap[projectId]) {
              var projectTasks = overdueMap[projectId];
              if (projectTasks && projectTasks[taskName]) {
                var taskStatus = projectTasks[taskName];
                
                // Text color — independent of bg
                if (userRole === "sw_tl" && taskStatus.PLRedActivity === 1) {
                    textColor = "#ff4444";
                } else if (userRole === "sw_head" && taskStatus.PMRedActivity === 1) {
                    textColor = "#ff4444";
                } else if (taskStatus.AlertDtYellow === 1) {
                    textColor = "#ffcc00";
                }
                
                // Priority 3: Progress Flag background color
                var progressFlag = taskStatus.ProgressFlag || 0;
                if (progressFlag >= 1 && progressFlag <= 7) {
                  var progressColors = {
                    1: "#ffb3b3",  // Engineering
                    2: "#5f933c",  // Purchase
                    3: "#0096cc",  // Software
                    4: "#005fa3",  // Project Management
                    5: "#2f491e",  // Manufacturing
                    6: "#00d9d9",  // Sales
                    7: "#6d006d"   // Client
                  };
                  bgColor = progressColors[progressFlag];                      
                }
              }
            }
            
            // Build style string and emit TD immediately — cannot reuse styleAttr
            // because it was already frozen before this block ran
            var cellStyle = "";
            if (bgColor)   cellStyle += "background-color:" + bgColor + ";";
            if (textColor) cellStyle += "color:" + textColor + ";font-weight:bold;";

            html += '<td class="monitor-td"' + (cellStyle ? ' style="' + cellStyle + '"' : '') + ' title="' + esc(val) + '">' + pct + '%</td>';
            return; // skip the generic html+= at the bottom of the loop
          }
          // ── 2. DROPDOWN: N and AP ───────────────────────────
           else if (DROPDOWN_OPTIONS[col]) {
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

                    // ── READ-ONLY with formatting ─────────────────────────
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
  function renderMonitor(container, data, user, overdueMap) {
    var cells  = data.cells   || {};
    var maxRow = data.max_row || 0;

    // Store overdueMap in a closure for use in renderTabHTML
    window.__overdueMap = overdueMap || {};
    
    var headerMap = buildHeaderMap(cells);
    var rows      = extractRows(cells, maxRow, user);
    var activeTab = 0;

    // allRows = unfiltered rows — needed to build full TL dropdown list
    var allRows = [];
    for (var r = DATA_START; r <= maxRow; r++) {
      if (cv(cells, ID_COL, r).trim() !== "") allRows.push(r);
    }

    // Remove existing monitor save bar if any
    var existingBar = document.getElementById("monitor-save-bar");
    if (existingBar) existingBar.remove();
    container.innerHTML = "";
    container.className = "monitor-shell";

    // ── collapseState ─────────────────────────────────────────
    var collapseStates = {};
    TABS.forEach(function (tab) {
      if (tab.colGroups && tab.colGroups.length) {
        collapseStates[tab.id] = tab.colGroups.map(function () { return true; });
      }
    });

    // ── Pending changes store ─────────────────────────────────
    var pendingChanges = {};

    function updateMonitorSaveBar() {
      var count   = Object.keys(pendingChanges).length;
      var bar     = document.getElementById("monitor-save-bar");
      var countEl = document.getElementById("monitor-change-count");
      if (bar)     bar.classList.toggle("visible", count > 0);
      if (countEl) countEl.textContent = count;
    }

    // ── Global cell-save event handler ────────────────────────
    // Inline onchange attributes call this.
    // Stages the change into pendingChanges and shows the save bar.
    // Does NOT fire any API call — that happens on Save.
    window.__monitorSaveCellEvt = function (col, row, value) {
      var key = col + row;
      if (!cells[key]) cells[key] = {};
      // Store original value on first edit so Discard can restore it
      if (cells[key]._orig === undefined) cells[key]._orig = cells[key].v;
      cells[key].v = value;

      // Stage the change
      pendingChanges[key] = { col: col, row: row, value: value };
      updateMonitorSaveBar();

      // Re-render the active panel to update bars / dependent cells
      var activePanel  = panelEls[activeTab];
      var activeTabDef = TABS[activeTab];
      if (activePanel && activeTabDef) {
        activePanel.innerHTML = renderTabHTML(
          cells, rows, allRows, activeTabDef,
          headerMap, collapseStates[activeTabDef.id] || [], user, overdueMap
        );
      }
    };

    // ── Save all pending monitor changes ──────────────────────
    window.saveMonitorChanges = function () {
      var btn     = document.getElementById("monitor-save-btn");
      var entries = Object.values(pendingChanges);
      if (!entries.length) return;

      btn.disabled = true;
      btn.innerHTML = '<span class="spinner" style="width:13px;height:13px;border-width:2px"></span>';

      // Fire sequentially so the server never races on the same file
      var chain = Promise.resolve();
      entries.forEach(function (e) {
        chain = chain.then(function () {
          return fetch('/api/monitor/cell', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ col: e.col, row: e.row, value: e.value }),
            credentials: 'include'
          }).then(function (r) { return r.json(); });
        });
      });

      chain.then(function () {
        pendingChanges = {};
        updateMonitorSaveBar();
        // Show a brief success toast if the host app provides one
        if (typeof window.toast === 'function') window.toast('Monitor saved \u2713');
        btn.disabled = false;
        btn.textContent = 'Save Changes';
      }).catch(function (err) {
        if (typeof window.toast === 'function') window.toast('Save failed', 'error');
        console.error('[Monitor] Save failed:', err);
        btn.disabled = false;
        btn.textContent = 'Save Changes';
      });
    };

    // ── Discard all pending monitor changes ───────────────────
    window.discardMonitorChanges = function () {
      // Revert the in-memory cells back to original values
      Object.keys(pendingChanges).forEach(function (key) {
        var e = pendingChanges[key];
        var origKey = e.col + e.row;
        // The original value was overwritten in cells — re-fetch from server
        // by simply clearing the v so it renders blank until next full load.
        // Better: store originals on first edit.
        if (cells[origKey] && cells[origKey]._orig !== undefined) {
          cells[origKey].v = cells[origKey]._orig;
          delete cells[origKey]._orig;
        }
      });
      pendingChanges = {};
      updateMonitorSaveBar();

      // Re-render active panel with reverted values
      var activePanel  = panelEls[activeTab];
      var activeTabDef = TABS[activeTab];
      if (activePanel && activeTabDef) {
        activePanel.innerHTML = renderTabHTML(
          cells, rows, allRows, activeTabDef,
          headerMap, collapseStates[activeTabDef.id] || [], user, overdueMap
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

    // ── Save bar (same structure as project save-bar) ─────────
    // ── Save bar (full width, fixed at bottom) ─────────
      var saveBar = document.createElement("div");
      saveBar.id = "monitor-save-bar";
      saveBar.className = "save-bar";
      // Inline styles to override sidebar offset and make it full width
      saveBar.style.cssText = "position:fixed !important;bottom:0 !important;left:0 !important;right:0 !important;width:100% !important;z-index:1000 !important;margin:0 !important;border-radius:0 !important;";
      saveBar.innerHTML =
          '<div class="save-bar-left">'
        +   '<div class="save-count" id="monitor-change-count">0</div>'
        +   '<div class="save-msg">unsaved changes</div>'
        + '</div>'
        + '<div class="save-actions">'
        +   '<button class="btn btn-secondary btn-sm" onclick="discardMonitorChanges()">Discard</button>'
        +   '<button class="btn btn-primary btn-sm" id="monitor-save-btn" onclick="saveMonitorChanges()">Save Changes</button>'
        + '</div>';

      container.appendChild(topRow);
      container.appendChild(panelsWrap);
      // Append to body instead of container so it stays fixed at bottom
      document.body.appendChild(saveBar);

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
            p.innerHTML = renderTabHTML(cells, rows, allRows, t, headerMap, collapseStates[t.id], user, overdueMap);

            // Re-attach resize handles after re-render
            var table = p.querySelector('.monitor-table');
            if (table) {
                setTimeout(function() {
                    makeColumnsResizable(table, t.id);
                }, 50);
            }
          };
        })(tab, panel);
      }

      panel.innerHTML = renderTabHTML(cells, rows, allRows, tab, headerMap, collapseStates[tab.id] || [], user, overdueMap);

      // Attach resize handles after rendering
      var table = panel.querySelector('.monitor-table');
        if (table) {
          setTimeout(function() {
              makeColumnsResizable(table, tab.id);
        }, 50);
      }
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

// Add this function to refresh monitor data
function refreshMonitorData(container, currentUser) {
    fetch('/api/monitor/sheet?t=' + Date.now(), { credentials: 'include' })
        .then(response => response.json())
        .then(result => {
            if (container && typeof renderMonitor === 'function') {
                // The response now has { sheet, overdue }
                renderMonitor(container, result.sheet, currentUser, result.overdue);
            }
        })
        .catch(err => console.error('Failed to refresh monitor:', err));
}

// Expose it globally so app.js can call it
global.refreshMonitorData = refreshMonitorData;

// ── Make columns resizable (temporary, no persistence) ──
function makeColumnsResizable(tableElement, tabId) {
    if (!tableElement) return;
    
    var ths = tableElement.querySelectorAll('.monitor-th');
    ths.forEach(function(th, index) {
        // Skip the # column (first th)
        if (index === 0 && th.classList.contains('monitor-th-num')) return;
        
        // Remove existing handle if any
        var existingHandle = th.querySelector('.resize-handle');
        if (existingHandle) existingHandle.remove();
        
        // Add resize handle on the LEFT edge
        var handle = document.createElement('div');
        handle.className = 'resize-handle';
        handle.style.cssText = [
            'position:absolute',
            'left:0',                    // Changed from right:0 to left:0
            'top:0',
            'width:5px',
            'height:100%',
            'cursor:col-resize',
            'user-select:none',
            'z-index:10',
            'background:transparent'
        ].join(';');
        
        th.style.position = 'relative';
        th.appendChild(handle);
        
        var startX, startWidth;
        
        handle.addEventListener('mousedown', function(e) {
            e.preventDefault();
            e.stopPropagation();
            startX = e.pageX;
            
            // Get the column to the LEFT of this one
            var leftTh = ths[index - 1];
            if (!leftTh) return; // No column to the left
            
            startWidth = leftTh.offsetWidth;
            
            function onMouseMove(e) {
                // Calculate new width for the LEFT column
                var delta = e.pageX - startX;
                var newWidth = startWidth + delta;
                
                if (newWidth > 30) { // Minimum width
                    // Resize the LEFT column
                    leftTh.style.width = newWidth + 'px';
                    leftTh.style.minWidth = newWidth + 'px';
                    
                    // Update corresponding col in colgroup (index - 1)
                    var colElements = tableElement.querySelectorAll('colgroup col');
                    if (colElements[index-1]) { // Index matches the left column's col
                        colElements[index-1].style.width = newWidth + 'px';
                    }
                    
                    // Update all cells in the LEFT column
                    var rows = tableElement.querySelectorAll('tbody tr');
                    rows.forEach(function(row) {
                        if (row.cells[index - 1]) {
                            row.cells[index - 1].style.width = newWidth + 'px';
                            row.cells[index - 1].style.minWidth = newWidth + 'px';
                        }
                    });
                }
            }
            
            function onMouseUp() {
                document.removeEventListener('mousemove', onMouseMove);
                document.removeEventListener('mouseup', onMouseUp);
                document.body.style.cursor = '';
            }
            
            document.body.style.cursor = 'col-resize';
            document.addEventListener('mousemove', onMouseMove);
            document.addEventListener('mouseup', onMouseUp);
        });
    });
}


})(window);
