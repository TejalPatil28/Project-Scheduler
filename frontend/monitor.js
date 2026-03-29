// ─────────────────────────────────────────────────────────────
//  monitor.js  —  Tabbed Monitor Screen
//  Called from app.js:  renderMonitor(container, data)
//  Data format:  { cells: { "I10": {v:"..."}, "I12": {v:"..."}, ... }, ... }
//  Row 10  = column headers
//  Row 11  = ignored
//  Row 12+ = project data rows
// ─────────────────────────────────────────────────────────────

(function (global) {
  "use strict";

  // ── Constants ───────────────────────────────────────────────
  var HEADER_ROW = 10;
  var DATA_START  = 12;
  var ID_COL      = "I";   // "Sales OR No" — row identifier

  // ── Tab definitions ─────────────────────────────────────────
  // cols: array of Excel column letters in display order
  var TABS = [
    {
      id:    "overview",
      label: "Project Overview",
      icon:  "🗂",
      cols:  ["I","B","F","J","G","H","K","L","N","AG","AP","AR","BA"],
    },
    {
      id:    "team",
      label: "Team & Roles",
      icon:  "👥",
      cols:  ["I","C","D","Z","AA","AB","AC","AD","AE","AF"],
    },
    {
      id:    "effort",
      label: "Effort & Value",
      icon:  "📊",
      cols:  ["I","M","AH","AI","AJ","AK","AL","AM","AN","AO"],
    },
    {
      id:    "ld",
      label: "Liquidated Damages",
      icon:  "⚖️",
      cols:  ["I","O","P","Q","R","S"],
    },
    {
      id:    "manpower",
      label: "Manpower & Visits",
      icon:  "🔧",
      cols:  ["I","T","U","V","W","X","Y"],
    },
    {
      id:    "milestone",
      label: "Milestone Progress",
      icon:  "🏁",
      scroll: true,
      cols:  [
        "I","BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK",
        "BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV",
        "BW","BX","BY","BZ","CA","CB","CC","CD",
      ],
    },
    {
      id:    "dates",
      label: "Dates & Tracking",
      icon:  "📅",
      cols:  ["I","E","AQ"],
    },
  ];

  // ── Helper: safe cell value ──────────────────────────────────
  function cv(cells, col, row) {
    var key = col + row;
    var info = cells[key];
    if (!info || info.v === null || info.v === undefined) return "";
    return String(info.v);
  }

  // ── Build header map from row 10 ─────────────────────────────
  // Returns { "I": "Sales OR No", "B": "File Name", ... }
  function buildHeaderMap(cells) {
    var map = {};
    // Collect all unique column letters used across all tabs
    var allCols = [];
    TABS.forEach(function (tab) {
      tab.cols.forEach(function (col) {
        if (allCols.indexOf(col) === -1) allCols.push(col);
      });
    });
    allCols.forEach(function (col) {
      var label = cv(cells, col, HEADER_ROW);
      map[col] = label || col; // fallback to column letter
    });
    return map;
  }

  // ── Extract data rows ────────────────────────────────────────
  // SW Head / Admin: all rows
  // SW TL: rows where their initials are in col C (SWH Head) first,
  //        if none found there, fall back to col D (SWE Name)
  function extractRows(cells, maxRow, user) {
    var role     = user && user.role;
    var initials = user && user.short_name ? user.short_name.trim().toUpperCase() : "";

    var isHead = (role === "head" || role === "admin");

    var rows = [];
    for (var r = DATA_START; r <= maxRow; r++) {
      var id = cv(cells, ID_COL, r);
      if (!id || id.trim() === "") continue; // skip blank rows
      rows.push(r);
    }

    // Head / Admin sees everything
    if (isHead) return rows;

    // SW TL: filter by initials — check SWH Head (col C) first
    var swhRows = rows.filter(function(r) {
      var val = cv(cells, "C", r).trim().toUpperCase();
      return val === initials;
    });
    if (swhRows.length > 0) return swhRows;

    // Fallback: check SWE Name (col D)
    return rows.filter(function(r) {
      var val = cv(cells, "D", r).trim().toUpperCase();
      return val === initials;
    });
  }

  // ── Escape HTML ──────────────────────────────────────────────
  function h(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  // ── Render one tab's table ───────────────────────────────────
  function renderTabTable(cells, rows, tab, headerMap) {
    var cols    = tab.cols;
    var scroll  = tab.scroll;

    var wrapStyle = [
      "flex:1",
      "min-height:0",
      "overflow-y:auto",
      scroll ? "overflow-x:auto" : "overflow-x:hidden",
    ].join(";");

    var html = '<div class="monitor-table-wrap" style="' + wrapStyle + '">';
    html += '<table class="monitor-table' + (scroll ? " monitor-table-scroll" : "") + '">';

    // ── thead ──
    html += "<thead><tr>";
    html += '<th class="monitor-th monitor-th-num">#</th>';
    cols.forEach(function (col) {
      html += '<th class="monitor-th">' + h(headerMap[col] || col) + "</th>";
    });
    html += "</tr></thead>";

    // ── tbody ──
    html += "<tbody>";
    if (rows.length === 0) {
      html += '<tr><td colspan="' + (cols.length + 1) + '" class="monitor-empty-row">No data found</td></tr>';
    } else {
      rows.forEach(function (r, idx) {
        var rowClass = idx % 2 === 0 ? "monitor-tr" : "monitor-tr monitor-tr-alt";
        html += '<tr class="' + rowClass + '">';
        html += '<td class="monitor-td monitor-td-num">' + (idx + 1) + "</td>";
        cols.forEach(function (col) {
          var val = cv(cells, col, r);
          html += '<td class="monitor-td" title="' + h(val) + '">' + h(val) + "</td>";
        });
        html += "</tr>";
      });
    }
    html += "</tbody></table></div>";
    return html;
  }

  // ── Main render function ─────────────────────────────────────
  function renderMonitor(container, data, user) {
    var cells   = data.cells  || {};
    var maxRow  = data.max_row || 0;

    var headerMap = buildHeaderMap(cells);
    var rows      = extractRows(cells, maxRow, user);

    var activeTab = 0; // default: Project Overview

    // ── Outer shell ──
    container.innerHTML = "";
    container.className = "monitor-shell";

    // ── Tab bar ──
    var tabBar = document.createElement("div");
    tabBar.className = "monitor-tabbar";

    var tabPanels = [];

    TABS.forEach(function (tab, idx) {
      // Tab button
      var btn = document.createElement("button");
      btn.className = "monitor-tab-btn" + (idx === activeTab ? " active" : "");
      btn.setAttribute("data-tab", idx);
      btn.innerHTML =
        '<span class="monitor-tab-icon">' + tab.icon + "</span>" +
        '<span class="monitor-tab-label">' + h(tab.label) + "</span>";
      btn.onclick = function () {
        switchTab(idx);
      };
      tabBar.appendChild(btn);

      // Panel (pre-rendered)
      var panel = document.createElement("div");
      panel.className = "monitor-panel" + (idx === activeTab ? " active" : "");
      panel.setAttribute("data-panel", idx);
      panel.innerHTML = renderTabTable(cells, rows, tab, headerMap);
      tabPanels.push(panel);
    });

    // ── Panels container ──
    var panelsWrap = document.createElement("div");
    panelsWrap.className = "monitor-panels";
    tabPanels.forEach(function (p) { panelsWrap.appendChild(p); });

    // ── Row count badge ──
    var badge = document.createElement("div");
    badge.className = "monitor-badge";
    badge.textContent = rows.length + " project" + (rows.length !== 1 ? "s" : "");

    // ── Assemble ──
    var topRow = document.createElement("div");
    topRow.className = "monitor-toprow";
    topRow.appendChild(tabBar);
    topRow.appendChild(badge);

    container.appendChild(topRow);
    container.appendChild(panelsWrap);

    // ── Tab switch ──
    function switchTab(idx) {
      activeTab = idx;
      tabBar.querySelectorAll(".monitor-tab-btn").forEach(function (b, i) {
        b.classList.toggle("active", i === idx);
      });
      panelsWrap.querySelectorAll(".monitor-panel").forEach(function (p, i) {
        p.classList.toggle("active", i === idx);
      });
    }
  }

  // ── Expose globally ──────────────────────────────────────────
  global.renderMonitor = renderMonitor;

})(window);
