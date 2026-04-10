// =============================================================
//  monitor.js  —  Config-driven Tabbed Monitor Screen
//
//  Called from app.js:
//    renderMonitor(container, data, user, overdueMap, config)
//
//  config  = one entry from MONITOR_CONFIGS in monitor_config.js
//            e.g. MONITOR_CONFIGS.SW  or  getMonitorConfig(user.role)
//
//  Data format:
//    { cells: { "B12": {v:"..."}, ... }, max_row: N, ... }
//
//  Save pattern:
//    window.__monitorSaveCell(col, row, value) — defined in app.js
//    monitor.js stages changes locally and fires on Save button.
// =============================================================

(function (global) {
  "use strict";

  // ── Helpers ──────────────────────────────────────────────────

  // Read a cell value as string (empty string if missing)
  function cv(cells, col, row) {
    var info = cells[col + row];
    if (!info || info.v === null || info.v === undefined) return "";
    return String(info.v);
  }

  // HTML-escape a value
  function esc(s) {
    return String(s)
      .replace(/&/g,  "&amp;")
      .replace(/</g,  "&lt;")
      .replace(/>/g,  "&gt;")
      .replace(/"/g,  "&quot;");
  }

  // ── Build header map from headerRow ─────────────────────────
  // Returns { "B": "OR No.", "C": "SW Head", ... }
  function buildHeaderMap(cells, config) {
    var map     = {};
    var allCols = [];
    config.tabs.forEach(function (tab) {
      tab.cols.forEach(function (col) {
        if (allCols.indexOf(col) === -1) allCols.push(col);
      });
    });
    allCols.forEach(function (col) {
      map[col] = cv(cells, col, config.headerRow) || col;
    });
    return map;
  }

  // ── Collect TL values from tlCol (all data rows) ─────────────
  // Used to populate the head/TL assignment dropdowns.
  function collectTLValues(cells, allRows, config) {
    var seen = {};
    var vals = [];
    allRows.forEach(function (r) {
      var v = cv(cells, config.tlCol, r).trim();
      if (v && !seen[v]) { seen[v] = true; vals.push(v); }
    });
    return vals.sort();
  }

  // ── Extract visible data rows (role-filtered) ─────────────────
  function extractRows(cells, maxRow, user, config) {
    var role     = (user && user.role) || "";
    var initials = (user && user.short_name) ? user.short_name.trim().toUpperCase() : "";
    var isHead   = (role === "head" || role === "admin" || role.endsWith("_head"));

    var rows = [];
    for (var r = config.dataStart; r <= maxRow; r++) {
      var id = cv(cells, config.idCol, r);
      if (!id || id.trim() === "") continue;
      rows.push(r);
    }

    // Head / admin see everything
    if (isHead) return rows;

    // TL: first try rows where headCol matches their initials,
    // fall back to rows where tlCol matches their initials.
    var headRows = rows.filter(function (r) {
      return cv(cells, config.headCol, r).trim().toUpperCase() === initials;
    });
    if (headRows.length > 0) return headRows;

    return rows.filter(function (r) {
      return cv(cells, config.tlCol, r).trim().toUpperCase() === initials;
    });
  }

  // ── Calculate column sums ─────────────────────────────────────
  function calculateColumnSums(cells, rows, sumCols) {
    var sums = {};
    sumCols.forEach(function (col) { sums[col] = 0; });

    rows.forEach(function (row) {
      sumCols.forEach(function (col) {
        var num = parseFloat(cv(cells, col, row));
        if (!isNaN(num)) sums[col] += num;
      });
    });
    return sums;
  }

  // ── Apply format rules to a raw cell value ────────────────────
  function applyFormat(col, val, formatRules) {
    if (!formatRules) return val;
    for (var i = 0; i < formatRules.length; i++) {
      var rule = formatRules[i];
      if (rule.cols.indexOf(col) === -1) continue;
      var num = parseFloat(val);
      if (isNaN(num)) return val;
      if (rule.format === "decimal2")          return num.toFixed(2);
      if (rule.format === "pct_from_decimal")  return Math.round(num * 100) + "%";
      return val;
    }
    return val;
  }

  // ── col → groupIndex lookup ───────────────────────────────────
  function buildColToGroup(colGroups) {
    var map = {};
    if (!colGroups) return map;
    colGroups.forEach(function (grp, gi) {
      grp.cols.forEach(function (col) { map[col] = gi; });
    });
    return map;
  }

  // ── Render one tab's full table HTML ──────────────────────────
  function renderTabHTML(cells, rows, allRows, tab, headerMap, collapseState, user, overdueMap, config) {
    var cols       = tab.cols;
    var colGroups  = tab.colGroups || null;
    var colToGroup = buildColToGroup(colGroups);
    var role       = (user && user.role) || "";
    var fluid      = tab.fluid;
    var isOverview  = (tab.id === "overview");
    var isMilestone = (tab.id === "milestone");

    // TL values for assignment dropdowns
    var tlValues = collectTLValues(cells, allRows, config);

    // ── Width helpers ─────────────────────────────────────────
    var widthMap = isOverview
      ? (config.colWidths["overview_pct"] || {})
      : (config.colWidths[tab.id] || {});

    function colWidth(col) {
      return widthMap[col] || widthMap["_default"] || (fluid ? "auto" : "40px");
    }

    function computeOverviewTableWidth() {
      var total = 0;
      cols.forEach(function (col) {
        if (hidden(col)) return;
        var w   = colWidth(col);
        var val = parseFloat(w);
        if (!isNaN(val) && w.indexOf("%") !== -1) total += val;
      });
      total += 2; // fixed 36px # col ≈ 2%
      return total <= 100 ? "100%" : total + "%";
    }

    // ── Collapse helper ───────────────────────────────────────
    function hidden(col) {
      if (!colGroups) return false;
      var gi = colToGroup[col];
      if (gi === undefined) return false;
      if (!collapseState[gi]) return false;
      var grpCols = colGroups[gi].cols;
      return col !== grpCols[grpCols.length - 1];
    }

    // ── Color helpers (config-driven) ─────────────────────────
    var isDark     = document.documentElement.getAttribute("data-theme") !== "light";
    var themeKey   = isDark ? "dark" : "light";
    var bgRules    = (config.colorRules     && config.colorRules[themeKey])     || [];
    var fgRules    = (config.textColorRules && config.textColorRules[themeKey]) || [];

    function getCellBg(col) {
      for (var i = 0; i < bgRules.length; i++) {
        if (bgRules[i].cols.indexOf(col) !== -1) return bgRules[i].bg;
      }
      return null;
    }
    function getCellFg(col) {
      for (var i = 0; i < fgRules.length; i++) {
        if (fgRules[i].cols.indexOf(col) !== -1) return fgRules[i].color;
      }
      return null;
    }

    // ── Shared style strings ──────────────────────────────────
    var btnStyle = [
      "cursor:pointer", "font-size:11px", "font-weight:700",
      "padding:1px 7px", "border-radius:3px",
      "border:1px solid " + (isDark ? "#555555" : "#aaaaaa"),
      "background:"      + (isDark ? "#2a2a2a" : "#e8e8e8"),
      "color:"           + (isDark ? "#dddddd" : "#333333"),
      "line-height:1.6",
    ].join(";");

    var inputBase = [
      "width:100%", "box-sizing:border-box",
      "border:none", "outline:none", "background:transparent",
      "font-family:var(--font-body)", "font-size:12px",
      "color:inherit", "padding:0",
    ].join(";");

    // ── Table wrapper ─────────────────────────────────────────
    var wrapStyle = [
      "flex:1", "min-height:0", "overflow-y:auto",
      (isOverview || isMilestone || (tab.scroll && !fluid)) ? "overflow-x:auto" : "overflow-x:hidden",
    ].join(";");

    var overviewTableWidth = isOverview ? computeOverviewTableWidth() : null;

    // Milestone: fixed px width so 46+ cols don't get crushed into 100%
    function computeMilestoneTableWidth() {
      var total = 36; // # col
      cols.forEach(function (col) {
        var w = colWidth(col);
        total += parseFloat(w) || 60;
      });
      return total + "px";
    }
    var milestoneTableWidth = isMilestone ? computeMilestoneTableWidth() : null;

    var tableClass = "monitor-table"
      + ((!isOverview && tab.scroll && !fluid) ? " monitor-table-scroll" : "")
      + (isMilestone ? " monitor-table-milestone" : "")
      + (isOverview  ? " monitor-table-pct"        : "");
    var tableStyle = isOverview
      ? "table-layout:fixed;width:" + overviewTableWidth + ";"
      : isMilestone
        ? "table-layout:fixed;width:" + milestoneTableWidth + ";"
        : (fluid ? "table-layout:fixed;width:100%;" : "");

    var html = '<div class="monitor-table-wrap" style="' + wrapStyle + '">';
    html += '<table class="' + tableClass + '"' + (tableStyle ? ' style="' + tableStyle + '"' : "") + '>';

    // ── <colgroup> ────────────────────────────────────────────
    html += "<colgroup>";
    html += '<col style="width:36px;">';
    cols.forEach(function (col) {
      if (hidden(col)) return;
      html += '<col style="width:' + colWidth(col) + ';">';
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
          var onclickFn = "window.__monitorToggleGroup_" + esc(tab.id) + "(" + gi + ")";

          if (!collapsed) {
            html += '<th colspan="' + grp.cols.length + '" style="position:sticky;top:0;z-index:5;background:rgba(0,0,0,0.3);border:none;text-align:left;padding:2px 0 0 4px;box-sizing:border-box;border-radius:4px 4px 0 0;">'
                  + '<button style="' + btnStyle + '" onclick="' + onclickFn + '">\u2013</button></th>';
            ci += grp.cols.length;
          } else {
            html += '<th style="position:sticky;top:0;z-index:5;background:transparent;border:none;text-align:left;padding:2px 0 0 4px;">'
                  + '<button style="' + btnStyle + '" onclick="' + onclickFn + '">+</button></th>';
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

    // ── ROW 1.5: Sum row (overview tab only) ──────────────────
    if (isOverview && config.sumCols && config.sumCols.length) {
      var sums = calculateColumnSums(cells, rows, config.sumCols);
      html += '<tr style="background:var(--bg4);">';
      cols.forEach(function (col) {
        if (hidden(col)) return;
        if (config.sumCols.indexOf(col) !== -1) {
          var bg  = isDark ? "#4a4a4a" : "#818181";
          var bdr = isDark ? "#666666" : "#555555";
          html += '<th class="monitor-th" style="background:' + bg + ';color:#ffffff;font-weight:700;text-align:center;border-bottom:2px solid ' + bdr + ';">'
               + (sums[col] || 0).toFixed(2) + '</th>';
        } else {
          html += '<th class="monitor-th" style="background:var(--bg4);"></th>';
        }
      });
      html += '</tr>';
    }

    // ── ROW 2: Column header row ──────────────────────────────
    html += '<tr>';
    if (!colGroups || !colGroups.length) {
      html += '<th class="monitor-th monitor-th-num">#</th>';
    }
    cols.forEach(function (col) {
      if (hidden(col)) return;
      var gi      = colToGroup[col];
      var isFirst = (gi !== undefined && colGroups && colGroups[gi].cols[0] === col);
      var w       = colWidth(col);
      var extra   = (isFirst ? "border-left:2px solid var(--accent);" : "")
                  + "width:" + w + " !important;max-width:" + w + " !important;";
      var hdrTxt  = headerMap[col] || col;
      var wrapSt  = (isMilestone || tab.wrapHeaders) ? "white-space:normal;word-break:break-word;line-height:1.2;padding:4px 2px;text-align:center;font-size:9px;" : "";
      html += '<th class="monitor-th" style="' + extra + wrapSt + '" title="' + esc(hdrTxt) + '">' + esc(hdrTxt) + '</th>';
    });
    html += '</tr></thead><tbody>';

    // ── TBODY ─────────────────────────────────────────────────
    var visibleCols = cols.filter(function (c) { return !hidden(c); });

    if (rows.length === 0) {
      html += '<tr><td colspan="' + (visibleCols.length + 1) + '" class="monitor-empty-row">No data found</td></tr>';
    } else {
      rows.forEach(function (r, idx) {
        html += '<tr class="monitor-tr' + (idx % 2 !== 0 ? " monitor-tr-alt" : "") + '">';
        html += '<td class="monitor-td monitor-td-num">' + (idx + 1) + '</td>';

        visibleCols.forEach(function (col) {
          var val       = cv(cells, col, r);
          var bgColor   = getCellBg(col);
          var textColor = getCellFg(col);
          var style     = "";
          if (bgColor)   style += "background-color:" + bgColor + ";";
          if (textColor) style += "color:" + textColor + ";";

          var gi2 = colToGroup[col];
          if (gi2 !== undefined && colGroups && colGroups[gi2].cols[0] === col) {
            style += "border-left:2px solid var(--accent);";
          }

          var styleAttr  = style ? ' style="' + style + '"' : "";
          var cellContent = "";

          // ── MILESTONE tab: progress bar display ──────────────
          if (isMilestone && config.milestoneCols.indexOf(col) !== -1) {
            var raw = parseFloat(val);
            var pct = isNaN(raw) ? 0 : (raw <= 1 && raw > 0 ? Math.round(raw * 100) : Math.round(raw));
            pct = Math.max(0, Math.min(100, pct));

            var taskName  = (headerMap[col] || "").trim();
            var projectId = cv(cells, config.idCol, r);
            var mFg = "#ffffff";
            var mBg = "#000000";

            if (overdueMap && projectId && overdueMap[projectId]) {
              var pTasks = overdueMap[projectId];
              if (pTasks && pTasks[taskName]) {
                var ts = pTasks[taskName];
                var uRole = (user && user.role) || "";
                if      (uRole === config.tlRole   && ts.PLRedActivity  === 1) mFg = "#ff4444";
                else if (uRole === config.headRole  && ts.PMRedActivity  === 1) mFg = "#ff4444";
                else if (ts.AlertDtYellow === 1)                                mFg = "#ffcc00";

                var pFlag = ts.ProgressFlag || 0;
                var pColors = { 1:"#ffb3b3", 2:"#5f933c", 3:"#0096cc", 4:"#005fa3", 5:"#2f491e", 6:"#00d9d9", 7:"#6d006d" };
                if (pColors[pFlag]) mBg = pColors[pFlag];
              }
            }

            var mStyle = "background-color:" + mBg + ";color:" + mFg + ";font-weight:bold;";
            html += '<td class="monitor-td" style="' + mStyle + '" title="' + esc(val) + '">' + pct + '%</td>';
            return; // skip generic append below
          }

          // ── NON-MILESTONE tabs: editable + formatted cells ───
          var editCfg = config.editableCols && config.editableCols[col];
          var canEdit = editCfg && (editCfg.role === null || editCfg.role === role);

          if (canEdit) {
            var oc = "window.__monitorSaveCellEvt('" + col + "'," + r + ",this.value)";

            if (editCfg.type === "tl_dropdown") {
              // Populated dynamically from all values in config.tlCol
              cellContent = '<select style="' + inputBase + ';cursor:pointer;" onchange="' + oc + '" onclick="event.stopPropagation()">';
              cellContent += '<option value="">\u2014</option>';
              if (val && tlValues.indexOf(val) === -1) {
                cellContent += '<option value="' + esc(val) + '" selected>' + esc(val) + '</option>';
              }
              tlValues.forEach(function (v) {
                cellContent += '<option value="' + esc(v) + '"' + (v === val ? " selected" : "") + '>' + esc(v) + '</option>';
              });
              cellContent += '</select>';

            } else if (editCfg.type === "dropdown") {
              var opts = editCfg.options || [];
              cellContent = '<select style="' + inputBase + ';cursor:pointer;" onchange="' + oc + '" onclick="event.stopPropagation()">';
              if (val && opts.indexOf(val) === -1) {
                cellContent += '<option value="' + esc(val) + '" selected>' + esc(val) + '</option>';
              }
              opts.forEach(function (opt) {
                cellContent += '<option value="' + esc(opt) + '"' + (opt === val ? " selected" : "") + '>' + esc(opt || "\u2014") + '</option>';
              });
              cellContent += '</select>';

            } else if (editCfg.type === "date") {
              var dateInputVal = "";
              if (val) {
                try {
                  var d = new Date(val);
                  if (!isNaN(d)) dateInputVal = d.toISOString().split("T")[0];
                } catch (e) {}
              }
              cellContent = '<input type="date" style="' + inputBase + ';cursor:text;" value="' + dateInputVal + '" onchange="' + oc + '" onclick="event.stopPropagation()" />';

            } else if (editCfg.type === "text") {
              cellContent = '<input type="text" value="' + esc(val) + '" style="' + inputBase + ';cursor:text;" onchange="' + oc + '" onclick="event.stopPropagation()" placeholder="\u2014" />';
            }

          } else {
            // Read-only: apply format rules then escape
            cellContent = esc(applyFormat(col, val, config.formatRules));
          }

          html += '<td class="monitor-td"' + styleAttr + ' title="' + esc(val) + '">' + cellContent + '</td>';
        });

        html += "</tr>";
      });
    }

    html += "</tbody></table></div>";
    return html;
  }


  // ── Main render function ──────────────────────────────────────
  function renderMonitor(container, data, user, overdueMap, config) {
    // Fall back to SW config if none provided
    if (!config) {
      config = (typeof MONITOR_CONFIGS !== "undefined" && MONITOR_CONFIGS.SW)
             ? MONITOR_CONFIGS.SW
             : null;
    }
    if (!config) {
      container.innerHTML = '<div class="alert alert-error">No monitor config found.</div>';
      return;
    }

    var cells  = data.cells   || {};
    var maxRow = data.max_row || 0;

    window.__overdueMap = overdueMap || {};

    var headerMap = buildHeaderMap(cells, config);
    var rows      = extractRows(cells, maxRow, user, config);
    var activeTab = 0;

    // allRows = unfiltered — needed to build full TL dropdown list
    var allRows = [];
    for (var r = config.dataStart; r <= maxRow; r++) {
      if (cv(cells, config.idCol, r).trim() !== "") allRows.push(r);
    }

    // Clean up any previous save bar
    var existingBar = document.getElementById("monitor-save-bar");
    if (existingBar) existingBar.remove();
    container.innerHTML = "";
    container.className = "monitor-shell";

    // ── Collapse states (all groups start collapsed) ──────────
    var collapseStates = {};
    config.tabs.forEach(function (tab) {
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

    // ── Global cell-edit handler ──────────────────────────────
    // Inline onchange attributes call window.__monitorSaveCellEvt.
    // Stages the change locally; does NOT call the API yet.
    window.__monitorSaveCellEvt = function (col, row, value) {
      var key = col + row;
      if (!cells[key]) cells[key] = {};
      if (cells[key]._orig === undefined) cells[key]._orig = cells[key].v;
      cells[key].v = value;

      pendingChanges[key] = { col: col, row: row, value: value };
      updateMonitorSaveBar();

      // Re-render active panel so dependents update immediately
      var activePanel  = panelEls[activeTab];
      var activeTabDef = config.tabs[activeTab];
      if (activePanel && activeTabDef) {
        activePanel.innerHTML = renderTabHTML(
          cells, rows, allRows, activeTabDef,
          headerMap, collapseStates[activeTabDef.id] || [], user, overdueMap, config
        );
      }
    };

    // ── Save all pending changes ──────────────────────────────
    window.saveMonitorChanges = function () {
      var btn     = document.getElementById("monitor-save-btn");
      var entries = Object.values(pendingChanges);
      if (!entries.length) return;

      btn.disabled = true;
      btn.innerHTML = '<span class="spinner" style="width:13px;height:13px;border-width:2px"></span>';

      // Fire sequentially — prevents race on the same file
      var chain = Promise.resolve();
      entries.forEach(function (e) {
        chain = chain.then(function () {
          return fetch("/api/monitor/cell", {
            method:      "POST",
            headers:     { "Content-Type": "application/json" },
            credentials: "include",
            body:        JSON.stringify({ col: e.col, row: e.row, value: e.value }),
          }).then(function (r) { return r.json(); });
        });
      });

      chain.then(function () {
        pendingChanges = {};
        updateMonitorSaveBar();
        if (typeof window.toast === "function") window.toast("Monitor saved \u2713");
        btn.disabled = false;
        btn.textContent = "Save Changes";
      }).catch(function (err) {
        if (typeof window.toast === "function") window.toast("Save failed", "error");
        console.error("[Monitor] Save failed:", err);
        btn.disabled = false;
        btn.textContent = "Save Changes";
      });
    };

    // ── Discard all pending changes ───────────────────────────
    window.discardMonitorChanges = function () {
      Object.keys(pendingChanges).forEach(function (key) {
        var e = pendingChanges[key];
        var origKey = e.col + e.row;
        if (cells[origKey] && cells[origKey]._orig !== undefined) {
          cells[origKey].v = cells[origKey]._orig;
          delete cells[origKey]._orig;
        }
      });
      pendingChanges = {};
      updateMonitorSaveBar();

      var activePanel  = panelEls[activeTab];
      var activeTabDef = config.tabs[activeTab];
      if (activePanel && activeTabDef) {
        activePanel.innerHTML = renderTabHTML(
          cells, rows, allRows, activeTabDef,
          headerMap, collapseStates[activeTabDef.id] || [], user, overdueMap, config
        );
      }
    };

    // ── Build shell DOM ───────────────────────────────────────
    var tabBar     = document.createElement("div");
    tabBar.className = "monitor-tabbar";

    var panelsWrap = document.createElement("div");
    panelsWrap.className = "monitor-panels";

    var badge = document.createElement("div");
    badge.className   = "monitor-badge";
    badge.textContent = rows.length + " project" + (rows.length !== 1 ? "s" : "");

    var topRow = document.createElement("div");
    topRow.className = "monitor-toprow";
    topRow.appendChild(tabBar);
    // Add "Add Project" button for PM Head only
    var role = (user && user.role) || "";
    if (role === "pm_head") {
        var addBtn = document.createElement("button");
        addBtn.className = "btn btn-primary btn-sm";
        addBtn.style.cssText = "margin-left:auto;margin-right:8px;padding:4px 12px;font-size:12px;";
        addBtn.innerHTML = '<span style="margin-right:4px;">+</span> Add Project';
        addBtn.onclick = function() { showAddProjectModal(cells, allRows, config, user, function() { renderMonitor(container, data, user, overdueMap, config); }, container); };
        topRow.appendChild(addBtn);
    }
    topRow.appendChild(badge);

    var saveBar = document.createElement("div");
    saveBar.id        = "monitor-save-bar";
    saveBar.className = "save-bar";
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
    document.body.appendChild(saveBar);

    // ── Create tab panels ─────────────────────────────────────
    var panelEls = [];

    config.tabs.forEach(function (tab, idx) {
      // Tab button
      var tbtn = document.createElement("button");
      tbtn.className = "monitor-tab-btn" + (idx === activeTab ? " active" : "");
      tbtn.setAttribute("data-tab", idx);
      tbtn.innerHTML =
        '<span class="monitor-tab-icon">'  + tab.icon  + "</span>" +
        '<span class="monitor-tab-label">' + esc(tab.label) + "</span>";
      tbtn.onclick = function () { switchTab(idx); };
      tabBar.appendChild(tbtn);

      // Panel div
      var panel = document.createElement("div");
      panel.className = "monitor-panel" + (idx === activeTab ? " active" : "");
      panel.setAttribute("data-panel", idx);
      panelEls.push(panel);
      panelsWrap.appendChild(panel);

      // Toggle handler for collapsible groups
      if (tab.colGroups && tab.colGroups.length) {
        (function (t, p) {
          window["__monitorToggleGroup_" + t.id] = function (gi) {
            collapseStates[t.id][gi] = !collapseStates[t.id][gi];
            p.innerHTML = renderTabHTML(
              cells, rows, allRows, t,
              headerMap, collapseStates[t.id], user, overdueMap, config
            );
            var tbl = p.querySelector(".monitor-table");
            if (tbl) setTimeout(function () { makeColumnsResizable(tbl, t.id); }, 50);
          };
        })(tab, panel);
      }

      // Initial render
      panel.innerHTML = renderTabHTML(
        cells, rows, allRows, tab,
        headerMap, collapseStates[tab.id] || [], user, overdueMap, config
      );

      var tbl = panel.querySelector(".monitor-table");
      if (tbl) setTimeout(function () { makeColumnsResizable(tbl, tab.id); }, 50);
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

  // ── Refresh helper (called from app.js if needed) ─────────────
  function refreshMonitorData(container, currentUser, config) {
    fetch("/api/monitor/sheet?t=" + Date.now(), { credentials: "include" })
      .then(function (r)  { return r.json(); })
      .then(function (result) {
        if (container && typeof renderMonitor === "function") {
          renderMonitor(container, result.sheet, currentUser, result.overdue || {}, config);
        }
      })
      .catch(function (err) { console.error("[Monitor] Refresh failed:", err); });
  }

  // ── Column resize handles ─────────────────────────────────────
  function makeColumnsResizable(tableElement, tabId) {
    if (!tableElement) return;
    var ths = tableElement.querySelectorAll(".monitor-th");
    ths.forEach(function (th, index) {
      if (index === 0 && th.classList.contains("monitor-th-num")) return;

      var existingHandle = th.querySelector(".resize-handle");
      if (existingHandle) existingHandle.remove();

      var handle = document.createElement("div");
      handle.className = "resize-handle";
      handle.style.cssText = "position:absolute;left:0;top:0;width:5px;height:100%;cursor:col-resize;user-select:none;z-index:10;background:transparent;";
      th.style.position = "relative";
      th.appendChild(handle);

      var startX, startWidth;

      handle.addEventListener("mousedown", function (e) {
        e.preventDefault();
        e.stopPropagation();
        startX = e.pageX;
        var leftTh = ths[index - 1];
        if (!leftTh) return;
        startWidth = leftTh.offsetWidth;

        function onMouseMove(e) {
          var newWidth = startWidth + (e.pageX - startX);
          if (newWidth > 30) {
            leftTh.style.width    = newWidth + "px";
            leftTh.style.minWidth = newWidth + "px";
            var cols = tableElement.querySelectorAll("colgroup col");
            if (cols[index - 1]) cols[index - 1].style.width = newWidth + "px";
            tableElement.querySelectorAll("tbody tr").forEach(function (row) {
              if (row.cells[index - 1]) {
                row.cells[index - 1].style.width    = newWidth + "px";
                row.cells[index - 1].style.minWidth = newWidth + "px";
              }
            });
          }
        }

        function onMouseUp() {
          document.removeEventListener("mousemove", onMouseMove);
          document.removeEventListener("mouseup",   onMouseUp);
          document.body.style.cursor = "";
        }

        document.body.style.cursor = "col-resize";
        document.addEventListener("mousemove", onMouseMove);
        document.addEventListener("mouseup",   onMouseUp);
      });
    });
  }

// ── Add Project Modal for PM Head ──────────────────────────────
function showAddProjectModal(cells, allRows, config, user, refreshCallback, container) {
    // Remove existing modal if any
    var existing = document.getElementById("monitor-add-project-modal");
    if (existing) existing.remove();
    
    var modal = document.createElement("div");
    modal.id = "monitor-add-project-modal";
    modal.className = "modal-overlay";
    modal.style.cssText = "position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:2000;display:flex;align-items:center;justify-content:center;";
    
    modal.innerHTML = `
        <div class="modal" style="max-width:500px;width:90%;background:var(--bg1);border-radius:12px;box-shadow:0 20px 35px rgba(0,0,0,0.3);">
            <div class="modal-header" style="display:flex;justify-content:space-between;align-items:center;padding:16px 20px;border-bottom:1px solid var(--border);">
                <div class="modal-title" style="font-size:18px;font-weight:700;">Add New Project</div>
                <button class="btn btn-ghost btn-sm" onclick="closeAddProjectModal()" style="background:none;border:none;font-size:20px;cursor:pointer;">&times;</button>
            </div>
            <div class="modal-body" style="padding:20px;">
                <div id="map-err" class="alert alert-error hidden" style="margin-bottom:15px;"></div>
                
                <div class="form-group" style="margin-bottom:15px;">
                    <label style="display:block;margin-bottom:5px;font-weight:600;">OR Number *</label>
                    <input type="text" id="map_or_number" class="form-input" style="width:100%;padding:8px;border-radius:6px;border:1px solid var(--border);background:var(--bg2);" placeholder="e.g. FSL/2425/PNQ/019">
                </div>
                
                <div class="form-group" style="margin-bottom:15px;">
                    <label style="display:block;margin-bottom:5px;font-weight:600;">Section *</label>
                    <input type="text" id="map_section" class="form-input" style="width:100%;padding:8px;border-radius:6px;border:1px solid var(--border);background:var(--bg2);" placeholder="Enter section">
                </div>
                
                <div class="form-group" style="margin-bottom:15px;">
                    <label style="display:block;margin-bottom:5px;font-weight:600;">OV in lakhs *</label>
                    <input type="number" step="0.01" id="map_ov_value" class="form-input" style="width:100%;padding:8px;border-radius:6px;border:1px solid var(--border);background:var(--bg2);" placeholder="0.00">
                </div>
                
                <div class="form-group" style="margin-bottom:15px;">
                    <label style="display:block;margin-bottom:5px;font-weight:600;">Assign to (PM Name) *</label>
                    <select id="map_assign_to" class="form-input" style="width:100%;padding:8px;border-radius:6px;border:1px solid var(--border);background:var(--bg2);">
                        <option value="">Select PM</option>
                    </select>
                </div>
            </div>
            <div class="modal-footer" style="display:flex;justify-content:flex-end;gap:10px;padding:16px 20px;border-top:1px solid var(--border);">
                <button class="btn btn-secondary" onclick="closeAddProjectModal()">Cancel</button>
                <button class="btn btn-primary" id="map-create-btn">Create Project</button>
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
    
    // Populate PM dropdown from existing tlCol values
    var pmSelect = document.getElementById("map_assign_to");
    if (pmSelect && allRows) {
        var pmValues = {};
        allRows.forEach(function(row) {
            var pm = cv(cells, config.tlCol, row).trim();
            if (pm && !pmValues[pm]) {
                pmValues[pm] = true;
                var option = document.createElement("option");
                option.value = pm;
                option.textContent = pm;
                pmSelect.appendChild(option);
            }
        });
    }
    
    // Bind create button event
    document.getElementById("map-create-btn").onclick = function() {
        createProjectFromMonitor(config, refreshCallback);
    };
}

// Close modal function
window.closeAddProjectModal = function() {
    var modal = document.getElementById("monitor-add-project-modal");
    if (modal) modal.remove();
};

// Create project from monitor
async function createProjectFromMonitor(config, refreshCallback) {
    var btn = document.getElementById("map-create-btn");
    var errEl = document.getElementById("map-err");
    var orNumber = document.getElementById("map_or_number").value.trim();
    var section = document.getElementById("map_section").value;
    var ovValue = document.getElementById("map_ov_value").value;
    var assignTo = document.getElementById("map_assign_to").value;
    
    // Validation
    if (!orNumber) {
        errEl.textContent = "OR Number is required";
        errEl.classList.remove("hidden");
        return;
    }
    if (!section) {
        errEl.textContent = "Section is required";
        errEl.classList.remove("hidden");
        return;
    }
    if (!ovValue) {
        errEl.textContent = "OV value is required";
        errEl.classList.remove("hidden");
        return;
    }
    if (!assignTo) {
        errEl.textContent = "Please select a PM to assign";
        errEl.classList.remove("hidden");
        return;
    }
    
    btn.disabled = true;
    btn.innerHTML = '<span class="spinner" style="width:13px;height:13px;border-width:2px"></span> Creating...';
    errEl.classList.add("hidden");
    
    try {
        var result = await fetch("/api/monitor/create-project", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            credentials: "include",
            body: JSON.stringify({
                or_number: orNumber,
                section: section,
                ov_value: ovValue,
                assign_to: assignTo
            })
        });
        
        var data = await result.json();
        
        if (!result.ok) {
            throw new Error(data.error || "Creation failed");
        }
        
        closeAddProjectModal();
        
        if (typeof window.toast === "function") {
            window.toast("Project created successfully!");
        }
        
        // Refresh the monitor view
        if (refreshCallback) refreshCallback();
        
    } catch(err) {
        errEl.textContent = err.message;
        errEl.classList.remove("hidden");
        btn.disabled = false;
        btn.innerHTML = "Create Project";
    }
}
  
  // ── Exports ───────────────────────────────────────────────────
  global.renderMonitor      = renderMonitor;
  global.refreshMonitorData = refreshMonitorData;

})(window);
