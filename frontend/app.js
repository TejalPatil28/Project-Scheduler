// ── State ──────────────────────────────────────────────────────
var state = {
  user: null,
  project: null,
  tasks: [],
  pendingChanges: {},
  activePhase: "ALL",
  theme: localStorage.getItem("theme") || "dark",
  sheetCache: {},  // project_id -> sheet data, client-side cache
};

// ── getSheetColumnWidth: reads from active schedule config ────
// schedule_config.js must be loaded before app.js.
function getSheetColumnWidth(col) {
  var cfg = (state && state.user) ? getScheduleConfig(state.user.role) : SCHEDULE_CONFIGS["SW"];
  var widths = cfg ? cfg.colWidths : {};
  return widths[col] || widths._default || "5%";
}

function applyTheme(t) {
  state.theme = t;
  document.documentElement.setAttribute("data-theme", t);
  localStorage.setItem("theme", t);
}
applyTheme(state.theme);

var app = document.getElementById("app");

// ── Toast ──────────────────────────────────────────────────────
function toast(msg, type) {
  type = type || "success";
  var c = document.getElementById("toast-container");
  var t = document.createElement("div");
  t.className = "toast toast-" + type;
  t.textContent = msg;
  c.appendChild(t);
  setTimeout(function() { t.remove(); }, 3200);
}

// ── Helpers ────────────────────────────────────────────────────
function roleLabel(role) {
  return userRoleLabel(role);  // delegated to schedule_config.js ROLE_LABELS
}

function phaseColor(ph) {
  var map = {P1:"#3b82f6",P2:"#f59e0b",P3:"#22c55e",P4:"#a855f7",P5:"#f43f5e"};
  return map[ph] || "#7c849e";
}

function ownerBadgeClass(owner) {
  var map = {HW:"badge-blue",SW:"badge-purple",MFG:"badge-green",PM:"badge-teal",EC:"badge-amber",BYR:"badge-gray"};
  return map[owner] || "badge-gray";
}

function ringColor(pct) {
  if (pct >= 100) return "#22c55e";
  if (pct >= 60)  return "#f59e0b";
  if (pct >= 20)  return "#00d4aa";
  return "#f43f5e";
}

function calcOverall(tasks) {
  if (!tasks.length) return 0;
  var sum = 0;
  tasks.forEach(function(t) { sum += t.percent_complete; });
  return Math.round(sum / tasks.length);
}

function taskKey(t) {
  return t.task_code + "_" + t.task_name.trim();
}

function isOverdue(d) {
  return d && new Date(d) < new Date();
}

function isSoon(d) {
  if (!d) return false;
  var diff = (new Date(d) - new Date()) / 86400000;
  return diff >= 0 && diff <= 7;
}

function h(str) {
  return String(str || "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

// ── Init ───────────────────────────────────────────────────────
async function init() {
  try {
    var user = await API.me();
    state.user = user;
    renderShell();
    renderDashboard();
  } catch(e) {
    renderLogin();
  }
}

// ── LOGIN ──────────────────────────────────────────────────────
function renderLogin() {
  app.innerHTML = [
    '<div class="login-wrap">',
      '<div class="login-grid"></div>',
      '<div class="login-glow"></div>',
      '<div class="login-card">',
        '<div class="login-logo" style="width:46px;height:46px;background:var(--accent);border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:22px;font-weight:800;color:white;">W</div>',
        '<div class="login-title" style="font-size:24px;font-weight:800;margin-bottom:4px;">Workezz</div>',
'<div class="login-sub" style="font-size:16px;color:var(--text2);margin-bottom:20px;">Project Scheduler</div>',
'<div style="font-size:11px;color:var(--text3);margin-bottom:26px;text-align:center;">Sign in to your workspace</div>',
        '<div id="login-err" class="alert alert-error hidden"></div>',
        '<div class="form-group">',
          '<label class="form-label">Username</label>',
          '<input class="form-input" id="l-user" type="text" placeholder="your.username" autocomplete="username">',
        '</div>',
        '<div class="form-group">',
          '<label class="form-label">Password</label>',
          '<input class="form-input" id="l-pass" type="password" placeholder="&bull;&bull;&bull;&bull;&bull;&bull;&bull;&bull;" autocomplete="current-password">',
        '</div>',
        '<button class="btn btn-primary btn-full" id="l-btn" style="margin-top:4px" onclick="doLogin()">Sign In &rarr;</button>',
      '</div>',
    '</div>'
  ].join("");

  document.getElementById("l-pass").addEventListener("keydown", function(e) { if (e.key==="Enter") doLogin(); });
  document.getElementById("l-user").addEventListener("keydown", function(e) { if (e.key==="Enter") document.getElementById("l-pass").focus(); });
  setTimeout(function() { var el = document.getElementById("l-user"); if(el) el.focus(); }, 100);
}

async function doLogin() {
  var btn   = document.getElementById("l-btn");
  var errEl = document.getElementById("login-err");
  var username = document.getElementById("l-user").value.trim();
  var password = document.getElementById("l-pass").value;
  errEl.classList.add("hidden");
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner" style="width:15px;height:15px;border-width:2px"></span>';
  try {
    var user = await API.login(username, password);
    state.user = user;
    renderShell();
    renderDashboard();
  } catch(err) {
    errEl.textContent = err.message || "Login failed";
    errEl.classList.remove("hidden");
    btn.disabled = false;
    btn.textContent = "Sign In \u2192";
  }
}

// ── SHELL ──────────────────────────────────────────────────────
function buildPopoverHTML() {
  var u = state.user;
  var isDark = state.theme === "dark";
  var adminItems = "";
  if (u.role === "admin") {
    adminItems = [
      '<div class="popover-sep"></div>',
      '<button class="popover-item" onclick="switchToUsers(); closeSettingsPopover();">\u{1F465} Manage Users</button>',
      '<button class="popover-item" onclick="doRebuildIndex(); closeSettingsPopover();">\u{1F504} Rebuild Index</button>',
      '<button class="popover-item" onclick="refreshMasterList(); closeSettingsPopover();">\u{1F504} Refresh Projects</button>',
    ].join("");
  }
  return [
    '<div class="popover hidden" id="user-popover">',
      '<div class="popover-header">',
        '<div class="avatar avatar-md">' + h(u.name[0].toUpperCase()) + '</div>',
        '<div>',
          '<div class="popover-name">' + h(u.name) + '</div>',
          '<div class="popover-role">' + userRoleLabel(u.role) + '</div>',
        '</div>',
      '</div>',
      '<div class="popover-section">',
        '<div class="popover-item" onclick="handleThemeToggle(event)">',
          '<div class="popover-item-left">',
            '<span id="theme-icon">' + (isDark ? "\u{1F319}" : "\u2600\uFE0F") + '</span>',
            '<span id="theme-label">' + (isDark ? "Dark Mode" : "Light Mode") + '</span>',
          '</div>',
          '<div class="theme-toggle ' + (isDark ? "" : "on") + '" id="theme-toggle"><div class="theme-thumb"></div></div>',
        '</div>',
        '<div class="popover-sep"></div>',
        '<button class="popover-item" onclick="switchToDashboard(); closeSettingsPopover();">\u{1F4CA} Dashboard</button>',
        '<button class="popover-item" onclick="switchToMonitor(); closeSettingsPopover();">\u{1F50D} Monitor</button>',
        '<button class="popover-item" onclick="switchToProjects(); closeSettingsPopover();">\u{1F4C1} Schedules</button>',
        adminItems,
        '<div class="popover-sep"></div>',
        '<button class="popover-item danger" onclick="doLogout()">\u23FB Sign Out</button>',
      '</div>',
    '</div>'
  ].join("");
}

function renderShell() {
  var u = state.user;

  app.innerHTML = [
    '<div class="shell">',
      // Full-width topbar at shell level
      '<div class="topbar">',
        '<div class="topbar-left">',
          '<div class="topbar-brand">',
            '<div class="brand-icon" style="width:28px;height:28px;font-size:16px;font-weight:800;color:white;background:var(--accent);border-radius:8px;display:flex;align-items:center;justify-content:center;">W</div>',
            '<div>',
              '<div class="brand-name">Workezz</div>',
              '<div class="brand-ver">Project Scheduler</div>',
            '</div>',
          '</div>',
          
          
        '</div>',
        '<div class="topbar-right" id="topbar-right">',
          '<div class="topbar-user-chip">',
            '<div class="avatar avatar-sm">' + h(u.name[0].toUpperCase()) + '</div>',
            '<div style="line-height:1.2">',
              '<div style="font-size:12px;font-weight:600;color:var(--text1)">' + h(u.name) + '</div>',
              '<div style="font-size:10px;color:var(--text2)">' + userRoleLabel(u.role) + '</div>',
            '</div>',
          '</div>',
          '<div class="popover-wrap" id="user-popover-wrap" style="position:relative">',
            '<button onclick="openSettingsPopover(event)" style="width:34px;height:34px;border-radius:var(--radius-sm);border:1px solid var(--border);display:flex;align-items:center;justify-content:center;background:var(--bg3);cursor:pointer;font-size:17px;color:var(--text2)">',
              '&#9881;',
            '</button>',
            buildPopoverHTML(),
          '</div>',
        '</div>',
      '</div>',
      // Body row: sidebar + main
      '<div class="shell-body">',
        '<aside class="sidebar" id="master-sidebar">',
          '<div class="master-sidebar-header">',
            '<div class="master-sidebar-title">My Files</div>',
            '<button class="sidebar-toggle-btn" onclick="toggleSidebar()" title="Close sidebar">&#8249;</button>',
          '</div>',
          '<div class="master-pane-list" id="master-list">',
            '<div class="master-empty">Loading...</div>',
          '</div>',
        '</aside>',
        // Mobile topbar (mobile only)
        '<div class="mobile-topbar" id="mobile-topbar">',
          '<div class="brand-icon" style="width:26px;height:26px;font-size:15px;font-weight:800;color:white;background:var(--accent);border-radius:8px;display:flex;align-items:center;justify-content:center;">W</div>',
          '<div id="mobile-title" style="font-family:var(--font-display);font-size:14px;font-weight:700;flex:1">Projects</div>',
        '</div>',
        '<div class="main">',
          '<div class="page" id="page-content"></div>',
        '</div>',
      '</div>',
    '</div>'
  ].join("");

  document.addEventListener("click", function(e) {
    var wrap = document.getElementById("user-popover-wrap");
    var pop  = document.getElementById("user-popover");
    if (wrap && pop && !wrap.contains(e.target) && !pop.contains(e.target)) pop.classList.add("hidden");
  });
}

function openSettingsPopover(e) {
  if (e) e.stopPropagation();
  var pop = document.getElementById("user-popover");
  if (pop) pop.classList.toggle("hidden");
}
function closeSettingsPopover() {
  var pop = document.getElementById("user-popover");
  if (pop) pop.classList.add("hidden");
}

function toggleSidebar() {
  var sidebar = document.getElementById("master-sidebar");
  var btn     = sidebar ? sidebar.querySelector(".sidebar-toggle-btn") : null;
  var saveBar = document.getElementById("save-bar");
  if (!sidebar) return;
  var collapsed = sidebar.classList.toggle("collapsed");
  // flip the arrow direction
  if (btn) btn.innerHTML = collapsed ? "&#8250;" : "&#8249;";
  // save-bar left: rail width when collapsed, full width when open
  if (saveBar) saveBar.style.left = collapsed ? "var(--sidebar-rail)" : "var(--sidebar-w)";
  // Refresh file status when sidebar OPENS (not collapsed)
  if (!collapsed) {
    refreshFileStatusOnOpen();
  }
}

function handleThemeToggle(e) {
  e.stopPropagation();
  var newTheme = state.theme === "dark" ? "light" : "dark";
  applyTheme(newTheme);
  var toggle = document.getElementById("theme-toggle");
  var icon   = document.getElementById("theme-icon");
  var label  = document.getElementById("theme-label");
  if (toggle) toggle.classList.toggle("on", newTheme === "light");
  if (icon)   icon.textContent  = newTheme === "dark" ? "\u{1F319}" : "\u2600\uFE0F";
  if (label)  label.textContent = newTheme === "dark" ? "Dark Mode" : "Light Mode";
  // Re-render sheet if open so colors update immediately
  var xlGrid = document.getElementById("xl-grid");
  if (xlGrid && state.project) {
    renderXLGrid();
  }
}

function setNav(active) {
  document.querySelectorAll(".nav-item").forEach(function(n) { n.classList.remove("active"); });
  var el = document.getElementById("nav-" + active);
  if (el) el.classList.add("active");
}

function setTopbar(title, showBack, backFn) {
  var titleEl     = document.getElementById("topbar-title");
  var backSlot    = document.getElementById("topbar-back-slot");
  var mobileTitle = document.getElementById("mobile-title");
  if (titleEl) {
    if (title) {
      titleEl.textContent = title;
      titleEl.classList.remove("hidden");
    } else {
      titleEl.classList.add("hidden");
    }
  }
  if (mobileTitle) mobileTitle.textContent = title || "";
  if (backSlot) { backSlot.innerHTML = ""; }
}

async function doLogout() {
  try { await API.logout(); } catch(e) {}
  state.user = null;
  app.innerHTML = "";
  renderLogin();
}

async function doRebuildIndex() {
  toast("Rebuilding index...");
  try {
    var res = await API.req("POST", "/admin/rebuild-index");
    toast(res.message);
  } catch(err) { toast(err.message, "error"); }
}

// ── DASHBOARD ─────────────────────────────────────────────────
function getProjectStatus(p) {
  var today = new Date(); today.setHours(0,0,0,0);
  if ((p.overall_percent || 0) >= 100) return "completed";
  if (p.start_date && new Date(p.start_date) > today) return "notstarted";
  return "ongoing";
}

var dashState = { tab: "all", search: "", activeMasterPid: null };

async function renderDashboard() {
  var pg = document.getElementById("page-content");
  if (pg) pg.classList.remove("no-pad");
  // Load data and populate sidebar - used by both views
  var page = document.getElementById("page-content");
  page.innerHTML = '<div class="loading"><span class="spinner"></span> Loading...</div>';

  var projects = [], masterList = [];
  try {
    var results = await Promise.allSettled([API.getProjects(), API.req("GET", "/master/projects")]);
    if (results[0].status === "fulfilled") projects = results[0].value;
    if (results[1].status === "fulfilled") masterList = results[1].value;
  } catch(err) {
    page.innerHTML = '<div class="alert alert-error">' + h(err.message) + '</div>';
    return;
  }

  renderMasterList(masterList, projects);
  window._dashProjects = projects;
  window._masterList   = masterList;

  // Default to Dashboard view
  switchToDashboard();
}

async function refreshMonitorData() {
    var page = document.getElementById("page-content");
    var monitorGrid = document.getElementById("monitor-grid");
    if (!monitorGrid) return;
    
    try {
        var result = await API.req("GET", "/monitor/sheet?t=" + Date.now());
        if (result && !result.error) {
            monitorGrid.innerHTML = "";
            var monitorCfg = getMonitorConfig(state.user.role);
            renderMonitor(monitorGrid, result.sheet, state.user, result.overdue || {}, monitorCfg);
        }
    } catch(err) {
        console.error("Failed to refresh monitor:", err);
    }
}

// Utility functions to switch to another screen

function switchToUsers() {
  setNavBtn("users");
  closeSettingsPopover();
  renderUsers();
}

async function switchToMonitor() {
    setNavBtn("monitor");
    setTopbar(null, false);
    var sidebar = document.getElementById("master-sidebar");
    if (sidebar) sidebar.style.display = "none";
    var pg = document.getElementById("page-content");
    if (pg) pg.classList.add("no-pad");
    var page = document.getElementById("page-content");
    page.innerHTML = '<div class="loading"><span class="spinner"></span> Loading monitor...</div>';
    try {
        var result = await API.req("GET", "/monitor/sheet?t=" + Date.now());
        if (!result || result.error) {
            page.innerHTML = '<div class="empty" style="padding-top:80px"><div class="empty-icon">&#9906;</div><div class="empty-text">No monitoring file found for your account.</div></div>';
            return;
        }
        var wrap = document.createElement("div");
        wrap.id = "monitor-grid";
        wrap.style.flex = "1";
        wrap.style.minHeight = "0";
        wrap.style.overflow = "auto";
        page.innerHTML = "";
        page.appendChild(wrap);
        // Pass sheet data and overdue data separately
        var monitorCfg = getMonitorConfig(state.user.role);
        renderMonitor(wrap, result.sheet, state.user, result.overdue || {}, monitorCfg);
    } catch(err) {
        page.innerHTML = '<div class="alert alert-error">' + h(err.message) + '</div>';
    }
}

function switchToProjects() {
  setNavBtn("projects");
  setTopbar(null, false);
  var sidebar = document.getElementById("master-sidebar");
  if (sidebar) sidebar.style.display = "";
  var page = document.getElementById("page-content");
  page.innerHTML = '<div class="projects-empty-state"><div class="projects-empty-icon">&#9672;</div><div class="projects-empty-text">Select a project from the sidebar</div></div>';
}

function switchToDashboard() {
  setNavBtn("dashboard");
  setTopbar(null, false);
  var sidebar = document.getElementById("master-sidebar");
  if (sidebar) sidebar.style.display = "none";
  var projects = window._dashProjects || [];
  var page = document.getElementById("page-content");
  page.innerHTML = renderDashboardGridHTML(projects);
  bindDashSearch(projects);
  filterAndRender(projects);
}

function setNavBtn(active) {
  ["dashboard", "monitor", "projects"].forEach(function(id) {
    var btn = document.getElementById("nav-btn-" + id);
    if (btn) btn.classList.toggle("active", id === active);
  });
}

function renderDashboardGridHTML(projects) {
  var counts = { all: projects.length, notstarted: 0, ongoing: 0, completed: 0 };
  projects.forEach(function(p) { counts[getProjectStatus(p)]++; });

  function tabBtn(id, label, count) {
    var active = dashState.tab === id ? " dash-tab-active" : "";
    return '<button class="dash-tab' + active + '" onclick="setDashTab(\'' + id + '\')">'
      + label + ' <span class="dash-tab-count">' + count + '</span></button>';
  }

  return [
    '<div class="dash-header">',
      '<div class="dash-title">Projects</div>',
      '<div class="search-wrap">',
        '<span class="search-icon">&#9906;</span>',
        '<input class="search-input" id="search" placeholder="Search projects..." value="' + h(dashState.search) + '">',
      '</div>',
    '</div>',
    '<div class="dash-tabs" id="dash-tabs">',
      tabBtn("all",        "All",          counts.all),
      tabBtn("ongoing",    "Ongoing",      counts.ongoing),
      tabBtn("notstarted", "Not Started",  counts.notstarted),
      tabBtn("completed",  "Completed",    counts.completed),
    '</div>',
    '<div class="proj-grid" id="proj-grid"></div>'
  ].join("");
}

function bindDashSearch(projects) {
  var el = document.getElementById("search");
  if (el) {
    el.addEventListener("input", function(e) {
      dashState.search = e.target.value;
      filterAndRender(projects);
    });
  }
  filterAndRender(projects);
}

function renderMasterList(masterList, projects) {
  var el = document.getElementById("master-list");
  if (!el) return;

  // Store master list globally for filtering
  window._masterListFull = masterList || [];
  window._projectsFull = projects || [];


  // Build search input HTML (NO REFRESH BUTTON HERE - it's already in sidebar header)
  var searchHtml = [
    '<div class="master-search-wrap" style="padding: 8px 10px; border-bottom: 1px solid var(--border);">',
      '<input type="text" id="master-search-input" class="form-input" style="font-size:12px; padding:6px 8px; width: 100%;" placeholder="Search project ID..." autocomplete="off">',
    '</div>'
  ].join("");

  // If no projects found
  if (!masterList || masterList.length === 0) {
    el.innerHTML = [
      searchHtml,
      '<div class="master-empty" style="padding: 20px; text-align: center; color: var(--text3);">No projects found</div>',
      '<div id="master-list-items" class="master-pane-list"></div>'
    ].join("");
    
    // Bind search input event
    var searchInput = document.getElementById("master-search-input");
    if (searchInput) {
      searchInput.addEventListener("input", function(e) {
        filterMasterList(e.target.value);
      });
    }
    
    
    return;
  }

  // If projects exist, show full list with button
  var html = [
    searchHtml,
    '<div id="master-list-items" class="master-pane-list"></div>'
  ].join("");

  el.innerHTML = html;

  // Bind search input event
  var searchInput = document.getElementById("master-search-input");
  if (searchInput) {
    searchInput.addEventListener("input", function(e) {
      filterMasterList(e.target.value);
    });
  }
  

  // Render the actual project items
  filterMasterList("");
}



async function refreshMasterList() {
    toast("Refreshing project list...");
    try {
      // Clear monitor cache on server
        await API.req("POST", "/admin/clear-monitor-cache");
        
        // Clear the client-side cache of master list
        window._masterListFull = null;
        window._masterList = null;
        
        // Fetch fresh master list from server
        var masterList = await API.req("GET", "/master/projects");
        var projects = window._dashProjects || [];
        
        // Update global variables
        window._masterList = masterList;
        window._masterListFull = masterList;
        
        // Re-render the sidebar list
        renderMasterList(masterList, projects);
        
        toast("Project list refreshed");
    } catch(err) {
        toast("Failed to refresh: " + err.message, "error");
    }
}

function filterMasterList(searchTerm) {
  var masterList = window._masterListFull || [];
  var projects = window._projectsFull || [];
  var container = document.getElementById("master-list-items");
  if (!container) return;

  var user = state.user || {};
  var role = user.role || "";
  var initials = (user.short_name || "").trim().toUpperCase();
  var isHead = (role === "head" || role === "admin" || (role && role.endsWith("_head")));

  // Apply role filtering — column keys from schedule config
  var schedCfg  = getScheduleConfig(role);
  var masterCols = (schedCfg && schedCfg.masterListCols) || { head: "swh_head", tl: "swe_name" };
  var roleFiltered;
  if (isHead) {
    roleFiltered = masterList;
  } else {
    var headRows = masterList.filter(function(m) {
      return (m[masterCols.head] || "").trim().toUpperCase() === initials;
    });
    roleFiltered = headRows.length > 0
      ? headRows
      : masterList.filter(function(m) {
          return (m[masterCols.tl] || "").trim().toUpperCase() === initials;
        });
  }

  // Apply search term on top of role filter
  var term = searchTerm.toLowerCase().trim();
  var filtered = roleFiltered.filter(function(m) {
    var pid = m.project_id || "";
    var fileId = m.file_id || "";
    return term === "" || pid.toLowerCase().includes(term) || fileId.toLowerCase().includes(term);
  });

  if (filtered.length === 0) {
    container.innerHTML = '<div class="master-empty" style="padding: 20px; text-align: center; color: var(--text3);">No matching projects</div>';
    return;
  }

  renderMasterItems(filtered, projects, container);
}

function renderMasterItems(masterList, projects, container) {
  if (!masterList.length) {
    container.innerHTML = '<div class="master-empty">No matching projects</div>';
    return;
  }

  var projMap = {};
  (projects || []).forEach(function(p) { projMap[p.id] = p; });

  var html = "";
  masterList.forEach(function(m) {
    var pid      = m.project_id;
    var stale    = m.stale;
    var exists   = m.file_exists;
    var isActive = dashState.activeMasterPid === pid;
    var proj     = projMap[m.file_id || pid];

    var staleIcon = stale ? '<span title="Not updated in last 2 days" style="font-size:11px;flex-shrink:0;margin-left:auto;padding-left:6px;">\u{1F550}</span>' : '';

    var customerName = proj ? h(proj.customer_name || "") : "";
    var pct = proj ? (proj.overall_percent || 0) : null;
    var progressBar = pct !== null
      ? '<div class="master-prog-bar"><div class="master-prog-fill" style="width:' + pct + '%;background:' + (pct >= 80 ? "var(--green)" : pct >= 40 ? "var(--amber)" : "var(--accent)") + '"></div></div>'
      : "";

    // Determine if this is a new project (needs setup)
    var isNewProject = (m.da_status === "NEW");

    var classes = "master-item"
      + (isActive ? " active" : "")
      + (stale    ? " stale"  : "")
      + (exists   ? ""        : " no-file")
      + (isNewProject ? " new-project" : "");

    var itemBg = isNewProject
      ? "background:rgba(59,130,246,0.18);border-color:rgba(59,130,246,0.5);"
      : !exists
        ? "background:rgba(244,63,94,0.28);border-color:rgba(244,63,94,0.5);"
        : "";

    var safeId   = h(pid);
    var safeFile = h(m.file_id || "");
    html += '<div class="' + classes + '" data-pid="' + safeId + '" data-fid="' + safeFile + '" data-exists="' + exists + '" onclick="handleMasterClick(this)" style="' + itemBg + '">'
      + '<div style="flex:1;min-width:0;">'
      +   '<div class="master-item-id">' + safeId + '</div>'
      +   (customerName ? '<div class="master-item-name">' + customerName + '</div>' : '')
      +   progressBar
      + '</div>'
      + staleIcon
      + '</div>';
  });

  container.innerHTML = html;
}

function handleMasterClick(el) {
  var pid    = el.getAttribute("data-pid");
  var fileId = el.getAttribute("data-fid");
  var exists = el.getAttribute("data-exists") === "true";
  onMasterItemClick(pid, exists, fileId);
}

function onMasterItemClick(pid, exists, fileId) {
  dashState.activeMasterPid = pid;
  document.querySelectorAll(".master-item").forEach(function(el) {
    el.classList.remove("active");
  });
  if (event && event.currentTarget) event.currentTarget.classList.add("active");

  if (!exists) {
    var page = document.getElementById("page-content");
    setTopbar(pid, true, renderDashboard);
    page.innerHTML =
      '<div class="empty" style="padding-top:80px">'
      + '<div class="empty-icon" style="font-size:32px;margin-bottom:14px">&#9906;</div>'
      + '<div class="empty-text" style="font-size:14px;color:var(--text2)">No project file found for <span style="color:var(--accent);font-family:var(--font-mono)">' + h(pid) + '</span></div>'
      + '<div style="font-size:12px;color:var(--text3);margin-top:6px">This project is tracked in your master file but has not been added to the database yet.</div>'
      + '</div>';
    return;
  }
  
  // Find the project in masterList to check da_status
  var masterProject = window._masterListFull.find(function(m) { return m.file_id === fileId || m.project_id === pid; });
  
  // If it's a NEW project (da_status === "NEW"), show setup form
  if (masterProject && masterProject.da_status === "NEW") {
    renderSetupForm(fileId || pid);
  } else {
    openProject(fileId || pid);
  }
}

function setDashTab(tab) {
  dashState.tab = tab;
  document.querySelectorAll(".dash-tab").forEach(function(b) {
    b.classList.remove("dash-tab-active");
  });
  var tabs = document.querySelectorAll(".dash-tab");
  var map = { all: 0, ongoing: 1, notstarted: 2, completed: 3 };
  if (tabs[map[tab]]) tabs[map[tab]].classList.add("dash-tab-active");
  filterAndRender(window._dashProjects || []);
}

function filterAndRender(projects) {
  var q = (dashState.search || "").toLowerCase();
  var list = projects.filter(function(p) {
    var matchTab = dashState.tab === "all" || getProjectStatus(p) === dashState.tab;
    var matchSearch = !q
      || (p.or_number||"").toLowerCase().includes(q)
      || (p.customer_name||"").toLowerCase().includes(q)
      || (p.project_desc||"").toLowerCase().includes(q)
      || (p.section||"").toLowerCase().includes(q);
    return matchTab && matchSearch;
  });
  renderProjectCards(list);
}

function renderProjectCards(list) {
  var grid = document.getElementById("proj-grid");
  if (!grid) return;
  if (!list.length) {
    grid.innerHTML = '<div class="empty" style="grid-column:1/-1"><div class="empty-icon">&#9672;</div><div class="empty-text">No results.</div></div>';
    return;
  }
  var html = "";
  list.forEach(function(p) {
    var sec = p.section || "";
    var ph  = sec === "SW" ? "P4" : sec === "HW" ? "P1" : sec === "MFG" ? "P3" : "P1";
    var color = phaseColor(ph);
    var tls = [];
    if (p.pm_name)     tls.push({label:"PM",  name:p.pm_name,     col:"#00d4aa"});
    if (p.hw_tl_name)  tls.push({label:"HW",  name:p.hw_tl_name,  col:"#3b82f6"});
    if (p.sw_tl_name)  tls.push({label:"SW",  name:p.sw_tl_name,  col:"#a855f7"});
    if (p.mfg_tl_name) tls.push({label:"MFG", name:p.mfg_tl_name, col:"#22c55e"});

    var tlsHtml = tls.map(function(t) {
      return '<div class="tl-chip">'
        + '<div class="tl-dot" style="background:' + t.col + '"></div>'
        + '<span style="color:var(--text3)">' + h(t.label) + '</span>'
        + '<span style="color:var(--text1);font-weight:600">' + h(t.name) + '</span>'
        + '</div>';
    }).join("");

    var footer = ""
      + (p.start_date ? '<span class="proj-date">Start: ' + h(p.start_date) + '</span>' : "")
      + (p.ld_date    ? '<span class="proj-date" style="color:var(--amber)">LD: ' + h(p.ld_date) + '</span>' : "")
      + '<div style="flex:1"></div>'
      + '<span style="font-family:var(--font-mono);font-size:10px;color:var(--text3)">' + h(sec||"—") + '</span>';

    var status = getProjectStatus(p);
    var statusMap = {
      completed:  { label: "Completed",   cls: "status-completed" },
      ongoing:    { label: "Ongoing",      cls: "status-ongoing"   },
      notstarted: { label: "Not Started",  cls: "status-notstarted"}
    };
    var st = statusMap[status];
    var statusBadge = '<span class="proj-status-badge ' + st.cls + '">' + st.label + '</span>';

    var pid = h(p.id);
    html += '<div class="proj-card" onclick="openProject(\'' + pid + '\')">'
      + '<div class="proj-card-top-bar" style="background:' + color + '"></div>'
      + '<div class="proj-card-body">'
      + '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px">'
      + '<div class="proj-or">' + h(p.or_number||p.id) + '</div>'
      + statusBadge
      + '</div>'
      + '<div class="proj-name">' + h(p.customer_name||"—") + '</div>'
      + '<div class="proj-desc">' + h(p.project_desc||"") + '</div>'
      + '<div class="proj-tls">' + tlsHtml + '</div>'
      + '</div>'
      + '<div class="proj-card-footer">' + footer + '</div>'
      + '</div>';
  });
  grid.innerHTML = html;
}

// ── PROJECT DETAIL ────────────────────────────────────────────
async function openProject(projectId) {
  var sidebar = document.getElementById("master-sidebar");
  if (sidebar) sidebar.style.display = "";
  var pg = document.getElementById("page-content");
  if (pg) pg.classList.add("no-pad");
  state.pendingChanges = {};
  state.activePhase = "ALL";
  var page = document.getElementById("page-content");
  page.innerHTML = '<div class="loading"><span class="spinner"></span> Loading project...</div>';
  setTopbar("Loading...", true, renderDashboard);

  // Highlight active item in sidebar
  dashState.activeMasterPid = projectId;
  document.querySelectorAll(".master-item").forEach(function(el) {
    var idEl = el.querySelector(".master-item-id");
    // master-item-id shows display pid; compare onclick fileId via data attr instead
    el.classList.toggle("active", el.getAttribute("data-fileid") === projectId);
  });

  var project, tasks;
  try {
    var results = await Promise.all([API.getProject(projectId), API.getTasks(projectId)]);
    project = results[0];
    tasks   = results[1];
  } catch(err) {
    page.innerHTML = '<div class="alert alert-error">' + h(err.message) + '</div>';
    return;
  }

  state.project = project;
  state.tasks   = tasks;
  setNavBtn("projects");
  setTopbar(project.customer_name || projectId, true, function() {
    switchToProjects();
  });
  renderProjectPage();
}

function metaItem(label, value, color) {
  var style = color ? ' style="color:' + color + '"' : '';
  return '<div class="meta-item"><div class="meta-label">' + label + '</div><div class="meta-value"' + style + '>' + h(value) + '</div></div>';
}

function renderProjectPage() {
  var p         = state.project;
  var container = document.getElementById("page-content");

  function cell(val, cls) {
    return '<div class="xi-cell' + (cls ? ' ' + cls : '') + '">'
      + '<span class="xi-val">' + h(val !== null && val !== undefined && val !== "" ? String(val) : "—") + '</span>'
      + '</div>';
  }
  function label(txt, cls) {
    return '<div class="xi-cell xi-lbl' + (cls ? ' ' + cls : '') + '">' + h(txt) + '</div>';
  }
  function hdr(txt, span) {
    return '<div class="xi-cell xi-hdr' + (span ? ' xi-span' + span : '') + '">' + h(txt) + '</div>';
  }
  function mfgDrop() {
    var gon = (p.mfg_loc||"GON") === "GON" ? " selected" : "";
    var dub = (p.mfg_loc||"GON") === "DUB" ? " selected" : "";
    return '<div class="xi-cell xi-editable">'
      + '<select class="xl-info-select" onchange="onMfgLocChange(this.value)">'
      + '<option value="GON"' + gon + '>GON</option>'
      + '<option value="DUB"' + dub + '>DUB</option>'
      + '</select></div>';
  }

  var leftHTML = [
    // ── Header rows 1-4 ──────────────────────────────────────
    '<div class="xi-row">',
      hdr("Guidelines", 2),
      label("*OR No."),   cell(p.or_number, "xi-accent"),
      label("*Quote No."), cell(p.quote_number),
    '</div>',
    '<div class="xi-row">',
      hdr("Cautions!!", 2),
      label("*Master OR"), cell(p.master_or, "xi-accent"),
      label("*Sales Eng."), cell(p.sales_engineer),
    '</div>',
    '<div class="xi-row">',
      hdr("", 2),
      label("Client PO#"),  cell(p.client_po),
      label("*Sales Mgr."), cell(p.sales_manager),
    '</div>',
    '<div class="xi-row xi-doc-row">',
      hdr("IMS Doc", 2),
      cell("FOX/IA/SALES/019", "xi-small xi-span2"),
      label("Ver"), cell("V 2.0"),
    '</div>',

    // ── Section divider ──────────────────────────────────────
    '<div class="xi-row xi-section-row">',
      hdr("*Compulsory field", 2), hdr("PM", 4),
    '</div>',

    // ── Project details rows 9-15 ────────────────────────────
    '<div class="xi-row">',
      label("*PO Value in lacs", 2), cell(p.po_value, "xi-accent xi-span2"),
    '</div>',
    '<div class="xi-row">',
      label("*Customer Name", 2), cell(p.customer_name || p.customer_name2, "xi-span2"),
    '</div>',
    '<div class="xi-row">',
      label("End Customer", 2), cell(p.end_customer, "xi-span2"),
    '</div>',
    '<div class="xi-row">',
      label("Consultant", 2), cell(p.consultant, "xi-span2"),
    '</div>',
    '<div class="xi-row">',
      label("*Project Desc.", 2), cell(p.project_desc || p.project_desc2, "xi-span2 xi-wrap"),
    '</div>',
    '<div class="xi-row">',
      label("*Section", 2), cell(p.section || p.section2, "xi-accent xi-span2"),
    '</div>',
    '<div class="xi-row">',
      label("Mfg Loc", 2), mfgDrop(), cell(""),
    '</div>',

    // ── Efforts rows 16-21 ───────────────────────────────────
    '<div class="xi-row xi-section-row">',
      hdr("Actual Efforts", 2), hdr("Total PO with VO", 2), cell(p.po_value),
    '</div>',
    '<div class="xi-row">',
      cell(""), label("*HW Efforts"), cell(p.hw_efforts, "xi-accent"), cell(""),
    '</div>',
    '<div class="xi-row">',
      cell(""), label("*No of Std Panels"), cell(p.std_panels || p.std_panels2), cell(""),
    '</div>',
    '<div class="xi-row">',
      cell(""), label("*No of Act Panels"), cell(p.act_panels || p.act_panels2), cell(""),
    '</div>',
    '<div class="xi-row">',
      cell(""), label("*SW Efforts"), cell(p.sw_efforts, "xi-accent"), cell(""),
    '</div>',
    '<div class="xi-row">',
      cell(""), label("*Mfg Efforts"), cell(p.mfg_efforts, "xi-accent"), cell(""),
    '</div>',

    // ── Dates header row 22 ──────────────────────────────────
    '<div class="xi-row xi-section-row">',
      cell(""), hdr("Customer Dates"), hdr("PM Dates"), cell(""),
    '</div>',

    // ── Date rows 23-34 ──────────────────────────────────────
    '<div class="xi-row">',
      label("PO Date"),     cell(""), cell(p.po_date), cell(""),
    '</div>',
    '<div class="xi-row">',
      label("OPF Recpt"),   cell(""), cell(p.opf_recpt), cell(""),
    '</div>',
    '<div class="xi-row">',
      label("HW Input"),    cell(""), cell(p.hw_input), cell(""),
    '</div>',
    '<div class="xi-row">',
      label("Dwg. Sub."),   cell(""), cell(p.dwg_sub), cell(""),
    '</div>',
    '<div class="xi-row">',
      label("Dwg Appr"),    cell(""), cell(p.dwg_appr), cell(""),
    '</div>',
    '<div class="xi-row">',
      label("HW FAT"),      cell(""), cell(p.hw_fat), cell(""),
    '</div>',
    '<div class="xi-row">',
      label("Dispatch"),    cell(""), cell(p.dispatch), cell(""),
    '</div>',
    '<div class="xi-row">',
      label("SW Input"),    cell(""), cell(p.sw_input), cell(""),
    '</div>',
    '<div class="xi-row">',
      label("SW FAT"),      cell(""), cell(p.sw_fat), cell(""),
    '</div>',
    '<div class="xi-row">',
      label("Install"),     cell(""), cell(p.install), cell(""),
    '</div>',
    '<div class="xi-row">',
      label("PreComm."),    cell(""), cell(p.precomm), cell(""),
    '</div>',
    '<div class="xi-row">',
      label("Comm."),       cell(""), cell(p.comm), cell(""),
    '</div>',

    // ── Stake holders rows 35-42 ─────────────────────────────
    '<div class="xi-row xi-section-row">',
      hdr("Stake Holders", 4),
    '</div>',
    '<div class="xi-row">',
      label("Sales"), cell(p.sh_sales, "xi-accent"),
      label("MFG"),   cell(p.sh_mfg, "xi-accent"),
    '</div>',
    '<div class="xi-row">',
      label("HW"),  cell(p.sh_hw, "xi-accent"),
      label("E&C"), cell(p.sh_ec, "xi-accent"),
    '</div>',
    '<div class="xi-row">',
      label("SW"),  cell(p.sh_sw, "xi-accent"),
      label("A/C"), cell(p.sh_ac, "xi-accent"),
    '</div>',
    '<div class="xi-row">',
      label("BYR"), cell(p.sh_byr, "xi-accent"), cell(""), cell(""),
    '</div>',

    // ── Actuals rows 43-52 ───────────────────────────────────
    '<div class="xi-row xi-section-row">',
      hdr("Actuals", 4),
    '</div>',
    '<div class="xi-row">',
      label("Balance Panels"), cell(p.balance_panels),
      label("Panel Disp Act"), cell(p.panel_disp_act),
    '</div>',
    '<div class="xi-row">',
      label("Est. VA%"), cell(p.est_va_pct),
      label("Est. VA"),  cell(p.est_va),
    '</div>',
    '<div class="xi-row">',
      label("Est. SM%"), cell(p.est_sm_pct),
      label("Est. SM"),  cell(p.est_sm),
    '</div>',
    '<div class="xi-row">',
      label("Act. VA%"), cell(p.act_va_pct),
      label("Act. VA"),  cell(p.act_va),
    '</div>',
    '<div class="xi-row">',
      label("Act. SM%"), cell(p.act_sm_pct),
      label("Act. SM"),  cell(p.act_sm),
    '</div>',
    '<div class="xi-row">',
      label("Remark", 2), cell(p.reason_remark, "xi-span2 xi-wrap"),
    '</div>',

  ].join("");

  container.innerHTML = [
  '<div class="project-header-bar" style="display:flex;align-items:center;justify-content:space-between;padding:8px 20px;background:var(--bg2);border-bottom:1px solid var(--border);margin-bottom:0;flex-shrink:0;">',
    '<div class="project-title" style="font-family:var(--font-display);font-size:18px;font-weight:700;color:var(--text1);">' + h(state.project.customer_name || state.project.id) + '</div>',
    '<div id="project-timestamp" style="font-size:11px;color:var(--text2);font-family:var(--font-mono);">Last updated: --</div>',
  '</div>',
    '<div class="proj-body-layout">',
      '<div class="proj-body-right" style="flex:1;min-width:0">',
        '<div class="xl-wrap" id="xl-wrap">',
          '<div class="xl-grid" id="xl-grid"></div>',
        '</div>',
      '</div>',
    '</div>',
    '<div class="save-bar" id="save-bar"' + (state.user && (state.user.role === "admin" || state.user.role === "head" || (state.user.role && state.user.role.endsWith("_head"))) ? ' style="display:none"' : '') + '>',
      '<div class="save-bar-left">',
        '<div class="save-count" id="change-count">0</div>',
        '<div class="save-msg">unsaved changes</div>',
      '</div>',
      '<div class="save-actions">',
        '<button class="btn btn-secondary btn-sm" onclick="discardChanges()">Discard</button>',
        '<button class="btn btn-primary btn-sm" id="save-btn" onclick="saveChanges()">Save Changes</button>',
      '</div>',
    '</div>'
  ].join("");

  renderXLGrid();
}

// ── SETUP FORM RENDERER v2 (for NEW projects) ─────────────────
// Clean modern form — layout inspired by Excel PrjSch left panel
// but rendered as a polished UI (not a raw spreadsheet clone).
// No non-editable display fields. No huge section tag columns.
// ──────────────────────────────────────────────────────────────
async function renderSetupForm(projectId) {
  var page = document.getElementById("page-content");
  page.innerHTML = '<div class="loading"><span class="spinner"></span> Loading setup form...</div>';
  setTopbar("Project Setup", true, function () { switchToProjects(); });

  try {
    var result = await API.req("GET", "/projects/" + projectId + "/setup");

    if (result.already_setup) {
      openProject(projectId);
      return;
    }

    var fields    = result.fields    || {};
    var taskNames = result.task_names || {};

    // ── Tiny helpers ──────────────────────────────────────────
    function has(f) { return !!fields[f]; }

    function inp(name, opts) {
      opts = opts || {};
      var cfg  = fields[name];
      if (!cfg) return "";
      var type = opts.type || cfg.type || "text";
      var req  = opts.req  ? ' required' : '';
      var cls  = "sf2-input" + (opts.wide ? " wide" : "");

      if (type === "dropdown") {
        var os = (cfg.options || []).map(function(o) {
          return '<option value="' + o + '">' + o + '</option>';
        }).join("");
        return '<select name="' + name + '" class="' + cls + '"' + req + '>'
          + '<option value="">Select…</option>' + os + '</select>';
      }
      if (type === "date") {
        return '<input type="date" name="' + name + '" class="' + cls + '"' + req + '>';
      }
      if (type === "number") {
        return '<input type="number" step="any" name="' + name
          + '" class="' + cls + '" placeholder="0"' + req + '>';
      }
      return '<input type="text" name="' + name + '" class="' + cls + '"' + req + '>';
    }

    // Field group: label + input in a .sf2-field div
    function field(label, name, opts) {
      if (!has(name)) return "";
      opts = opts || {};
      var reqMark = opts.req ? '<span class="sf2-req-dot">*</span>' : '';
      return '<div class="sf2-field' + (opts.cls ? ' ' + opts.cls : '') + '">'
        + '<label class="sf2-label">' + reqMark + label + '</label>'
        + inp(name, opts)
        + '</div>';
    }

    // Section header — compact pill style
    function sec(icon, title) {
      return '<div class="sf2-section-head">'
        + '<span class="sf2-section-icon">' + icon + '</span>'
        + '<span class="sf2-section-title">' + title + '</span>'
        + '</div>';
    }

    // Row wrapper (2-col or 3-col grid)
    function row2(a, b)    { return '<div class="sf2-row2">' + a + b + '</div>'; }
    function row3(a, b, c) { return '<div class="sf2-row3">' + a + b + c + '</div>'; }
    function row1(a)       { return '<div class="sf2-row1">' + a + '</div>'; }

    // ── CSS ───────────────────────────────────────────────────
    var css = `<style>
/* ── Root wrap ── */
.sf2-wrap {
  display: flex;
  flex-direction: column;
  height: 100%;
  overflow: hidden;
  background: var(--bg1);
  font-family: var(--font-sans, 'Inter', 'Segoe UI', Arial, sans-serif);
}

/* ── Top bar ── */
.sf2-topbar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 12px 24px;
  background: var(--bg2);
  border-bottom: 1px solid var(--border);
  flex-shrink: 0;
}
.sf2-topbar-left { display: flex; flex-direction: column; gap: 2px; }
.sf2-topbar-name { font-size: 16px; font-weight: 700; color: var(--text1); letter-spacing: -0.02em; }
.sf2-topbar-id   { font-size: 11px; color: var(--text3); font-family: var(--font-mono); }
.sf2-topbar-actions { display: flex; gap: 10px; align-items: center; }

/* ── Scrollable body ── */
.sf2-body {
  flex: 1;
  overflow-y: auto;
  padding: 20px 24px 80px;
}

/* ── Section header ── */
.sf2-section-head {
  display: flex;
  align-items: center;
  gap: 8px;
  margin: 24px 0 12px;
  padding-bottom: 7px;
  border-bottom: 1.5px solid var(--border);
}
.sf2-section-head:first-child { margin-top: 0; }
.sf2-section-icon {
  width: 26px; height: 26px;
  border-radius: 6px;
  background: var(--accent);
  color: #fff;
  font-size: 13px;
  display: flex; align-items: center; justify-content: center;
  flex-shrink: 0;
}
.sf2-section-title {
  font-size: 12px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.08em;
  color: var(--text2);
}

/* ── Grid rows ── */
.sf2-row2 {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 12px 20px;
  margin-bottom: 10px;
}
.sf2-row3 {
  display: grid;
  grid-template-columns: 1fr 1fr 1fr;
  gap: 12px 20px;
  margin-bottom: 10px;
}
.sf2-row1 {
  display: grid;
  grid-template-columns: 1fr;
  gap: 12px 20px;
  margin-bottom: 10px;
}

/* ── Individual field ── */
.sf2-field {
  display: flex;
  flex-direction: column;
  gap: 4px;
}
.sf2-field.full { grid-column: 1 / -1; }

/* ── Label ── */
.sf2-label {
  font-size: 11px;
  font-weight: 600;
  color: var(--text2);
  display: flex;
  align-items: center;
  gap: 4px;
  user-select: none;
}
.sf2-req-dot {
  width: 5px; height: 5px;
  border-radius: 50%;
  background: var(--accent);
  display: inline-block;
  flex-shrink: 0;
}

/* ── Inputs ── */
.sf2-input {
  height: 34px;
  padding: 0 10px;
  border: 1px solid var(--border);
  border-radius: 6px;
  background: var(--bg2);
  color: var(--text1);
  font-family: inherit;
  font-size: 13px;
  width: 100%;
  box-sizing: border-box;
  transition: border-color 0.15s, box-shadow 0.15s;
}
.sf2-input:focus {
  outline: none;
  border-color: var(--accent);
  box-shadow: 0 0 0 2px rgba(59,130,246,0.18);
}
select.sf2-input { cursor: pointer; }
textarea.sf2-input {
  height: auto;
  min-height: 60px;
  padding: 8px 10px;
  resize: vertical;
  line-height: 1.4;
}

/* ── Dates subsection: customer | PM layout ── */
.sf2-dates-grid {
  display: grid;
  grid-template-columns: 160px 1fr 1fr;
  gap: 0;
  border: 1px solid var(--border);
  border-radius: 8px;
  overflow: hidden;
  margin-bottom: 12px;
}
.sf2-dg-head {
  font-size: 10px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.07em;
  color: var(--text3);
  padding: 5px 10px;
  background: var(--bg2);
  border-bottom: 1px solid var(--border);
}
.sf2-dg-head.center { text-align: center; }
.sf2-dg-head.accent-col { background: rgba(59,130,246,0.07); }
.sf2-dg-lbl {
  font-size: 12px;
  font-weight: 500;
  color: var(--text2);
  padding: 6px 10px;
  background: var(--bg2);
  border-right: 1px solid var(--border);
  border-top: 1px solid var(--border);
  display: flex;
  align-items: center;
}
.sf2-dg-lbl.req::before {
  content: '';
  width: 5px; height: 5px;
  border-radius: 50%;
  background: var(--accent);
  display: inline-block;
  margin-right: 6px;
  flex-shrink: 0;
}
.sf2-dg-cell {
  padding: 4px 6px;
  border-top: 1px solid var(--border);
  border-right: 1px solid var(--border);
  background: var(--bg1);
  display: flex;
  align-items: center;
}
.sf2-dg-cell:last-child { border-right: none; }
.sf2-dg-cell.auto-cell {
  font-size: 10px;
  color: var(--text3);
  font-style: italic;
  justify-content: center;
  background: var(--bg2);
}
.sf2-dg-cell input {
  width: 100%;
  border: none;
  outline: none;
  background: transparent;
  font-family: inherit;
  font-size: 12px;
  color: var(--text1);
  padding: 2px 2px;
}
.sf2-dg-cell input:focus { background: rgba(59,130,246,0.07); border-radius: 3px; }

/* ── Stakeholders: inline badge-style layout ── */
.sf2-sh-grid {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 10px;
  margin-bottom: 12px;
}

/* ── Scope checkboxes ── */
.sf2-scope-row {
  display: flex;
  gap: 24px;
  flex-wrap: wrap;
  padding: 10px 4px;
  margin-bottom: 12px;
}
.sf2-scope-item {
  display: flex;
  align-items: center;
  gap: 7px;
  font-size: 13px;
  font-weight: 500;
  color: var(--text1);
  cursor: pointer;
}
.sf2-scope-item input[type=checkbox] {
  width: 15px; height: 15px;
  cursor: pointer;
  accent-color: var(--accent);
}

/* ── Task table ── */
.sf2-task-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 12px;
  border: 1px solid var(--border);
  border-radius: 8px;
  overflow: hidden;
  margin-top: 4px;
}
.sf2-task-table th {
  padding: 7px 10px;
  font-size: 10px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.06em;
  color: var(--text2);
  background: var(--bg2);
  border-bottom: 1px solid var(--border);
  text-align: left;
}
.sf2-task-table th:not(:first-child) { text-align: center; }
.sf2-task-table td {
  padding: 4px 8px;
  border-top: 1px solid var(--border);
  color: var(--text1);
}
.sf2-task-table td:first-child {
  font-size: 12px;
  font-weight: 500;
  color: var(--text1);
  background: var(--bg2);
  border-right: 1px solid var(--border);
}
.sf2-task-table td input {
  width: 100%;
  border: none;
  outline: none;
  background: transparent;
  font-family: inherit;
  font-size: 12px;
  color: var(--text1);
  padding: 4px 4px;
  text-align: center;
}
.sf2-task-table td input:focus { background: rgba(59,130,246,0.08); border-radius: 3px; }
.sf2-task-table tr:hover td { background: rgba(59,130,246,0.03); }
.sf2-task-table tr:hover td:first-child { background: var(--bg2); }

/* ── Bottom save bar ── */
.sf2-footer {
  position: sticky;
  bottom: 0;
  display: flex;
  justify-content: flex-end;
  gap: 10px;
  padding: 12px 24px;
  background: var(--bg2);
  border-top: 1px solid var(--border);
  flex-shrink: 0;
}
</style>`;

    // ── Build form HTML ───────────────────────────────────────
    var body = "";

    // ── 1. Header Info ────────────────────────────────────────
    body += sec("📋", "Header Information");
    body += row2(
      field("Master OR",        "master_or",       { req: true }),
      field("Quote Number",     "quote_number",     { req: true })
    );
    body += row2(
      field("Client PO#",       "client_po"),
      field("Sales Engineer",   "sales_engineer",  { req: true })
    );
    body += row1(
      field("Sales Manager",    "sales_manager",   { req: true })
    );

    // ── 2. Project Details ────────────────────────────────────
    body += sec("🏗", "Project Details");
    body += row2(
      field("PO Value (in lacs)", "po_value",      { req: true, type: "number" }),
      field("Section",            "section",        { req: true, type: "dropdown" })
    );
    body += row1(
      field("Customer Name",    "customer_name",   { req: true })
    );
    body += row2(
      field("End Customer",     "end_customer"),
      field("Consultant",       "consultant")
    );
    body += row1(
      field("Project Description", "project_desc", { req: true })
    );
    body += row1(
      field("Manufacturing Location", "mfg_loc",  { type: "dropdown" })
    );

    // ── 3. Efforts ────────────────────────────────────────────
    body += sec("⚙", "Efforts");
    body += row3(
      field("HW Efforts",       "hw_efforts",      { req: true, type: "number" }),
      field("SW Efforts",       "sw_efforts",      { req: true, type: "number" }),
      field("Mfg Efforts",      "mfg_efforts",     { req: true, type: "number" })
    );
    body += row2(
      field("No. of Std Panels", "std_panels",     { req: true, type: "number" }),
      field("No. of Act Panels", "act_panels",     { req: true, type: "number" })
    );

    // ── 4. Dates — 3-column grid (label | cust date | PM auto) ──
    body += sec("📅", "Dates");
    body += '<div class="sf2-dates-grid">';
    // header row
    body += '<div class="sf2-dg-head">Milestone</div>';
    body += '<div class="sf2-dg-head center accent-col">Customer Date</div>';
    body += '<div class="sf2-dg-head center">PM Date</div>';
    // date rows
    var dateList = [
      ["PO Date",    "po_date",   true],
      ["OPF Recpt",  "opf_recpt", false],
      ["HW Input",   "hw_input",  false],
      ["Dwg. Sub.",  "dwg_sub",   false],
      ["Dwg Appr",   "dwg_appr",  false],
      ["HW FAT",     "hw_fat",    false],
      ["Dispatch",   "dispatch",  false],
      ["SW Input",   "sw_input",  false],
      ["SW FAT",     "sw_fat",    false],
      ["Install",    "install",   false],
      ["PreComm.",   "precomm",   false],
      ["Comm.",      "comm",      false],
    ];
    dateList.forEach(function(dr) {
      if (!has(dr[1]) && !dr[2]) return; // skip if not in fields and not required
      body += '<div class="sf2-dg-lbl' + (dr[2] ? ' req' : '') + '">' + dr[0] + '</div>';
      body += '<div class="sf2-dg-cell"><input type="date" name="' + dr[1] + '"></div>';
      body += '<div class="sf2-dg-cell auto-cell">auto</div>';
    });
    body += '</div>';

    // ── 5. Stakeholders ───────────────────────────────────────
    body += sec("👥", "Stakeholders");
    var shList = [
      ["Sales", "sh_sales"], ["MFG",   "sh_mfg"],
      ["HW",    "sh_hw"],    ["E&C",   "sh_ec"],
      ["SW",    "sh_sw"],    ["A/C",   "sh_ac"],
      ["BYR",   "sh_byr"],
    ];
    body += '<div class="sf2-sh-grid">';
    shList.forEach(function(sh) {
      body += field(sh[0], sh[1]);
    });
    body += '</div>';

    // ── 6. Actuals ────────────────────────────────────────────
    body += sec("📊", "Actuals");
    body += row2(
      field("Est. VA%",        "est_va_pct",  { type: "number" }),
      field("Est. VA",         "est_va",      { type: "number" })
    );
    body += row2(
      field("Est. SM%",        "est_sm_pct",  { type: "number" }),
      field("Est. SM",         "est_sm",      { type: "number" })
    );
    body += row2(
      field("Act. VA%",        "act_va_pct",  { type: "number" }),
      field("Act. VA",         "act_va",      { type: "number" })
    );
    body += row2(
      field("Act. SM%",        "act_sm_pct",  { type: "number" }),
      field("Act. SM",         "act_sm",      { type: "number" })
    );
    body += row2(
      field("Balance Panels",  "balance_panels", { type: "number" }),
      field("Panel Disp Act",  "panel_disp_act", { type: "number" })
    );
    body += row1(
      field("Reason / Remark", "reason_remark")
    );

    // ── 7. LD Details ─────────────────────────────────────────
    var hasLD = has("ld_date") || has("ld_maxwk") || has("ld_maxov") || has("ld_remarks");
    if (hasLD) {
      body += sec("⚠", "LD Details");
      body += row2(
        field("LD Date",      "ld_date",    { type: "date" }),
        field("LD Max/Wk %",  "ld_maxwk",  { type: "number" })
      );
      body += row2(
        field("LD Max/OV %",  "ld_maxov",  { type: "number" }),
        field("LD Remarks",   "ld_remarks")
      );
    }

    // ── 8. Warranty ───────────────────────────────────────────
    if (has("warranty")) {
      body += sec("🛡", "Warranty");
      body += row1(field("Warranty Terms", "warranty"));
    }

    // ── 9. Scope ──────────────────────────────────────────────
    var scopeList = ["scope_hw", "scope_sw", "scope_mfg", "scope_inst", "scope_com"];
    var scopeLabels = { scope_hw: "HW", scope_sw: "SW", scope_mfg: "MFG", scope_inst: "Inst", scope_com: "Com" };
    var hasScope = scopeList.some(function(f) { return has(f); });
    if (hasScope) {
      body += sec("🔍", "Scope");
      body += '<div class="sf2-scope-row">';
      scopeList.forEach(function(f) {
        if (!has(f)) return;
        body += '<label class="sf2-scope-item">'
          + '<input type="checkbox" name="' + f + '" value="1"> '
          + scopeLabels[f] + '</label>';
      });
      body += '</div>';
    }

    // ── 10. Lead Times ────────────────────────────────────────
    if (has("critical_lead_time") || has("normal_lead_time")) {
      body += sec("⏱", "Lead Times");
      body += row2(
        field("Critical Lead Time", "critical_lead_time", { type: "number" }),
        field("Normal Lead Time",   "normal_lead_time",   { type: "number" })
      );
    }

    // ── 11. Task-Specific Fields ──────────────────────────────
    var taskRows = {};
    Object.keys(fields).forEach(function(key) {
      var cfg = fields[key];
      if (!cfg || !cfg.task_row) return;
      var r = cfg.task_row;
      if (!taskRows[r]) {
        taskRows[r] = { s: null, u: null, ad: null, name: taskNames[r] || ("Task " + r) };
      }
      // Avoid matching "sh_*", "sw_*", "std_*", "sc_*" as task 's' fields
      if (key.match(/^s\d/) || key === "s")  taskRows[r].s  = key;
      if (key.match(/^u\d/) || key === "u")  taskRows[r].u  = key;
      if (key.match(/^ad/))                  taskRows[r].ad = key;
    });

    var sortedTaskRows = Object.keys(taskRows).sort(function(a, b) {
      return parseInt(a) - parseInt(b);
    });

    if (sortedTaskRows.length > 0) {
      body += sec("📝", "Task-Specific Fields");
      body += '<table class="sf2-task-table">';
      body += '<thead><tr>'
        + '<th style="width:40%">Task Description</th>'
        + '<th style="width:20%">Lead Time (S)</th>'
        + '<th style="width:20%">Effort Days (U)</th>'
        + '<th style="width:20%">Payment % (AD)</th>'
        + '</tr></thead><tbody>';

      sortedTaskRows.forEach(function(rn) {
        var tk = taskRows[rn];
        body += '<tr>'
          + '<td>' + h(tk.name) + '</td>'
          + '<td><input type="number" step="any" name="' + (tk.s  || "") + '" placeholder="—" ' + (tk.s  ? '' : 'disabled') + '></td>'
          + '<td><input type="number" step="any" name="' + (tk.u  || "") + '" placeholder="—" ' + (tk.u  ? '' : 'disabled') + '></td>'
          + '<td><input type="number" step="any" name="' + (tk.ad || "") + '" placeholder="—" ' + (tk.ad ? '' : 'disabled') + '></td>'
          + '</tr>';
      });

      body += '</tbody></table>';
    }

    // ── Assemble ──────────────────────────────────────────────
    var html = css + [
      '<div class="sf2-wrap">',
        '<div class="sf2-topbar">',
          '<div class="sf2-topbar-left">',
            '<div class="sf2-topbar-name">Project Setup</div>',
            '<div class="sf2-topbar-id">' + h(projectId) + '</div>',
          '</div>',
          '<div class="sf2-topbar-actions">',
            '<button type="button" class="btn btn-secondary btn-sm" onclick="switchToProjects()">Cancel</button>',
            '<button type="button" class="btn btn-primary btn-sm" id="sf2-save-btn" onclick="doSetupSave(\'' + h(projectId) + '\')">Save &amp; Continue →</button>',
          '</div>',
        '</div>',
        '<div class="sf2-body">',
          '<form id="project-setup-form">',
            body,
          '</form>',
        '</div>',
      '</div>',
    ].join("");

    page.innerHTML = html;

    // ── Save handler ──────────────────────────────────────────
    window.doSetupSave = async function(pid) {
      var btn = document.getElementById("sf2-save-btn");
      var formData = {};
      var els = document.getElementById("project-setup-form").elements;
      for (var i = 0; i < els.length; i++) {
        var el = els[i];
        if (!el.name || el.disabled) continue;
        if (el.type === "checkbox") {
          formData[el.name] = el.checked ? el.value : "";
        } else {
          formData[el.name] = el.value;
        }
      }
      btn.disabled = true;
      btn.innerHTML = '<span class="spinner" style="width:13px;height:13px;border-width:2px"></span>';
      try {
        var res = await API.req("POST", "/projects/" + pid + "/setup", formData);
        if (res.success) {
          toast("Project setup completed!");

          // Patch in-memory master list directly — don't clear cache
          if (window._masterListFull) {
            window._masterListFull.forEach(function(m) {
              if (m.file_id === pid || m.project_id === pid) {
                m.da_status = "CONFIGURED";
              }
            });
          }
          if (window._masterList) {
            window._masterList.forEach(function(m) {
              if (m.file_id === pid || m.project_id === pid) {
                m.da_status = "CONFIGURED";
              }
            });
          }

          // Re-render sidebar with patched data
          var searchTerm = document.getElementById("master-search-input")?.value || "";
          filterMasterList(searchTerm);

          openProject(pid);
        } else {
          toast(res.message || "Save failed", "error");
          btn.disabled = false;
          btn.innerHTML = "Save &amp; Continue →";
        }
      } catch(err) {
        toast(err.message || "Save failed", "error");
        btn.disabled = false;
        btn.innerHTML = "Save &amp; Continue →";
      }
    };

  } catch(err) {
    console.error("Setup form error:", err);
    page.innerHTML = '<div class="alert alert-error">Failed to load setup form: ' + h(err.message) + '</div>';
  }
}

function updateProjectHeaderTimestamp(timestamp) {
  var timestampEl = document.getElementById("project-timestamp");
  if (timestampEl && timestamp) {
    timestampEl.textContent = "Last updated: " + timestamp;
  }
}

function onMfgLocChange(val) {
  state.project.mfg_loc = val;
}


function onMfgLocChange(val) {
  // Store locally for now - view only except this dropdown
  state.project.mfg_loc = val;
}

function filterPhase(phase) {
  state.activePhase = phase;
  document.querySelectorAll(".phase-btn").forEach(function(btn) {
    var txt = btn.textContent.trim();
    var isThis = phase === "ALL" ? txt === "All Phases" : txt.startsWith(phase);
    btn.classList.toggle("active", isThis);
  });
  renderXLGrid();
}

function renderXLGrid() {
  var wrap = document.getElementById("xl-grid");
  if (!wrap) return;
  wrap.innerHTML = '<div style="padding:20px;color:var(--text2);font-size:12px">Loading sheet...</div>';
  loadSheetView();
}

async function loadSheetView() {
  var wrap = document.getElementById("xl-grid");
  if (!wrap || !state.project) return;
  var pid = state.project.id;
  try {
    // Use client-side cache if available — instant re-open
    if (state.sheetCache[pid]) {
      renderExcelMirror(wrap, state.sheetCache[pid]);

      if (state.sheetCache[pid].last_modified) {
        updateProjectHeaderTimestamp(state.sheetCache[pid].last_modified);
      }

      return;
    }
    var data = await API.req("GET", "/projects/" + pid + "/sheet" );
    state.sheetCache[pid] = data;  // cache in browser
    // Update timestamp
    if (data.last_modified) {
      updateProjectHeaderTimestamp(data.last_modified);
    }

    renderExcelMirror(wrap, data);
  } catch(err) {
    wrap.innerHTML = '<div style="padding:20px;color:var(--text3)">Could not load sheet: ' + h(err.message||"error") + '</div>';
  }
}

function renderExcelMirror(container, data) {
  // Detect dark mode first — used throughout this function
  var isDark = document.documentElement.getAttribute("data-theme") !== "light";

  var cells        = data.cells;
  var colWidths    = data.col_widths;
  var rowHeights   = data.row_heights;
  var maxRow       = data.max_row;
  var cols         = data.cols;
  var schedCfg  = getScheduleConfig((state.user && state.user.role) || "sw_tl");
  var colGroups = schedCfg ? schedCfg.colGroups : [];
  var editableFill = data.editable_fill || null;
  var infoRows     = data.info_rows    || [];
  var headerRows   = data.header_rows  || [];
  var banner       = data.project_banner || {};
  var leftPanel    = data.left_panel   || {};
  var infoRowSet   = {};
  infoRows.forEach(function(r) { infoRowSet[r] = true; });
  var headerRowSet = {};
  headerRows.forEach(function(r) { headerRowSet[r] = true; });

  // ── Pull all rendering config from schedule_config.js ───────
  var TASK_START_ROW = schedCfg ? schedCfg.taskStartRow : 9;
  var TASK_END_ROW   = schedCfg ? schedCfg.taskEndRow   : 55;
  var skipColsSet    = {};
  ((schedCfg && schedCfg.skipCols) || ["AC","AE"]).forEach(function(c){ skipColsSet[c] = true; });

  // Help dropdown — may be null for depts that have no color-coded help column (e.g. PM)
  var helpDropdown  = (schedCfg && schedCfg.helpDropdown) || null;
  var AD_COL        = helpDropdown ? (helpDropdown.col || "AD") : null;
  var AD_OPTIONS    = helpDropdown ? (helpDropdown.options    || []) : [];
  var AD_COLORS     = helpDropdown ? (isDark ? (helpDropdown.colorsDark  || {}) : (helpDropdown.colorsLight || {})) : {};
  var AD_FG_DARK    = helpDropdown ? (helpDropdown.fgDark     || {}) : {};

  // Color rules — map from config format to existing internal format
  var _cfgColorRules = (schedCfg && schedCfg.colorRules)
    ? (isDark ? schedCfg.colorRules.dark : schedCfg.colorRules.light)
    : [];
  var colorRules = _cfgColorRules.map(function(r) {
    return { columns: r.cols, rows: r.rows, color: r.color };
  });

  // Text rules
  var _cfgTextRules = (schedCfg && schedCfg.textRules) || [];
  
  // ── Text rule evaluator (config-driven) ──────────────────────
  // Supported conditions: not_equal | less_than | value_zero | value_equals | value_less_than
  function getTextRuleColor(col, row, cells) {
    for (var i = 0; i < _cfgTextRules.length; i++) {
      var rule = _cfgTextRules[i];
      if (rule.col !== col) continue;
      if (rule.rows === "9+" && row < 9) continue;

      // Cross-column comparison: cell ≠ compareWith col
      if (rule.condition === "not_equal" && rule.compareWith) {
        var curVal  = cells[rule.col         + row] ? String(cells[rule.col         + row].v || "") : "";
        var cmpVal  = cells[rule.compareWith + row] ? String(cells[rule.compareWith + row].v || "") : "";
        var nCur = normalizeValue(curVal);
        var nCmp = normalizeValue(cmpVal);
        if (nCur === "" && typeof nCmp === "number") nCur = 0;
        if (nCmp === "" && typeof nCur === "number") nCmp = 0;
        if (nCur !== nCmp) return rule.color;
      }

      // Cross-column comparison: cell < compareWith col
      if (rule.condition === "less_than" && rule.compareWith) {
        var curVal  = cells[rule.col         + row] ? String(cells[rule.col         + row].v || "") : "";
        var cmpVal  = cells[rule.compareWith + row] ? String(cells[rule.compareWith + row].v || "") : "";
        var numCur  = parseFloat(curVal.replace("%", "").trim());
        var numCmp  = parseFloat(cmpVal.replace("%", "").trim());
        if (!isNaN(numCur) && !isNaN(numCmp) && numCur < numCmp) return rule.color;
      }

      // Self-value: cell is 0, "0", null, or empty
      if (rule.condition === "value_zero") {
        var cellInfo = cells[rule.col + row];
        var rawVal   = cellInfo ? String(cellInfo.v !== null && cellInfo.v !== undefined ? cellInfo.v : "") : "";
        var num      = parseFloat(rawVal.trim());
        if (rawVal.trim() === "" || rawVal.trim() === "0" || (!isNaN(num) && num === 0)) return rule.color;
      }

      // Self-value: cell === rule.value (string match after trim)
      if (rule.condition === "value_equals" && rule.value !== undefined) {
        var cellInfo = cells[rule.col + row];
        var rawVal   = cellInfo ? String(cellInfo.v !== null && cellInfo.v !== undefined ? cellInfo.v : "") : "";
        if (rawVal.trim() === String(rule.value).trim()) return rule.color;
      }

      // Self-value: parseFloat(cell) < rule.value
      if (rule.condition === "value_less_than" && rule.value !== undefined) {
        var cellInfo = cells[rule.col + row];
        var rawVal   = cellInfo ? String(cellInfo.v !== null && cellInfo.v !== undefined ? cellInfo.v : "") : "";
        var num      = parseFloat(rawVal.replace("%", "").trim());
        if (!isNaN(num) && num < rule.value) return rule.color;
      }
    }
    return null;
  }

    // Helper function to normalize values for comparison
    function normalizeValue(val) {
        if (val === null || val === undefined || val === "") return "";
        
        var str = String(val).trim();
        
        // Check if it's a number (with possible leading zeros like "02", "2.0", etc.)
        var num = parseFloat(str);
        if (!isNaN(num) && isFinite(num)) {
            // If it's a whole number, compare as number
            if (num === Math.floor(num)) {
                return num;
            }
            return num;
        }
        
        // Not a number, return as string
        return str;
    }

  // Function to check if a cell should get a custom background color
  function getCustomBackgroundColor(col, row) {
    for (var i = 0; i < colorRules.length; i++) {
      var rule = colorRules[i];
      // Check if column is in rule's columns list
      if (rule.columns.indexOf(col) !== -1) {
        // Check row condition
        if (rule.rows === "all") {
          return rule.color;
        } else if (rule.rows === "9+" && row >= 9) {
          return rule.color;
        }
      }
    }
    return null;
  }

  // Build a pending changes map: coord -> new value (staged, not yet saved)
  var _editPending = {};

  function renderInputCell(info, coord, col, tdStyle, overrideColor) {
    var curVal = _editPending[coord] !== undefined ? _editPending[coord] : (info.v || "");
    var inputTextColor = overrideColor || (isDark ? "#d8d8d8" : "#000000");
    var inputStyle = "width:100%;height:100%;border:none;outline:none;background:transparent;font-family:Calibri,Arial,sans-serif;font-size:11px;padding:0 2px;box-sizing:border-box;color:" + inputTextColor + ";";
    var input = "";

    var _editColCfg = schedCfg && schedCfg.editableCols ? schedCfg.editableCols[col] : null;
    var _colType    = _editColCfg ? _editColCfg.type : null;

    // "readonly" type — show value as plain text, no input widget
    if (_colType === "readonly") {
      return h(String(curVal !== null && curVal !== undefined ? curVal : ""));
    }

    if (_colType === "date") {
      var dateVal = curVal ? curVal : "";
      if (dateVal) {
        try {
          var d = new Date(dateVal);
          if (!isNaN(d)) dateVal = d.toISOString().split("T")[0];
        } catch(e) {}
      }
      input = "<input type=\"date\" style=\"" + inputStyle + "\" value=\"" + dateVal + "\" data-coord=\"" + coord + "\" data-col=\"" + col + "\" onchange=\"__xlEditCell(this)\">";
    } else if (_colType === "number") {
      var _min = (_editColCfg && _editColCfg.min !== undefined) ? _editColCfg.min : 0;
      var _max = (_editColCfg && _editColCfg.max !== undefined) ? _editColCfg.max : 100;
      var numVal = String(curVal || "0").replace("%", "").trim();
      input = "<input type=\"number\" min=\"" + _min + "\" max=\"" + _max + "\" style=\"" + inputStyle + "\" value=\"" + numVal + "\" data-coord=\"" + coord + "\" data-col=\"" + col + "\" onchange=\"__xlEditCell(this)\">";
    } else if (_colType === "dropdown") {
      // dropdown options come from helpDropdown (only valid when helpDropdown is non-null)
      var opts = AD_OPTIONS.map(function(o) {
        return "<option value=\"" + o + "\"" + (o === curVal ? " selected" : "") + ">" + o + "</option>";
      }).join("");
      input = "<select style=\"" + inputStyle + "cursor:pointer;\" data-coord=\"" + coord + "\" data-col=\"" + col + "\" onchange=\"__xlEditCell(this);__applyAdColor(this)\" ><option value=\"\"></option>" + opts + "</select>";
    } else if (_colType === "text") {
      input = "<input type=\"text\" style=\"" + inputStyle + "\" value=\"" + h(curVal || "") + "\" data-coord=\"" + coord + "\" data-col=\"" + col + "\" onchange=\"__xlEditCell(this)\">";
    }
    return input;
  }

  // Track collapse state per group - START COLLAPSED
  var collapseState = colGroups.map(function(g) { return true; });

  // Fast lookup: col_letter -> groupIndex
  var colToGroup = {};
  colGroups.forEach(function(g, gi) {
    g.cols.forEach(function(c) { colToGroup[c] = gi; });
  });

  // Returns true if this col should be hidden due to group collapse
  function isCollapsedCol(col) {
    var gi = colToGroup[col];
    if (gi === undefined || gi < 0) return false;
    if (!collapseState[gi]) return false;
    var lastCol = colGroups[gi].cols[colGroups[gi].cols.length - 1];
    return col !== lastCol;
  }

  // Returns true if this col should be skipped entirely (from config.skipCols)
  function isSkippedCol(col) {
    return !!skipColsSet[col];
  }

  function borderStyle(weight) {
    var w = weight === "medium" ? "2px" : weight === "thick" ? "3px" : "1px";
    var c = weight === "medium" ? "#8c8c8c" : weight === "thick" ? "#595959" : "#BFBFBF";
    return w + " solid " + c;
  }

  var xlBg         = isDark ? "#1a1a1a" : "#ffffff";
  var xlZebraOdd  = isDark ? "#1a1a1a" : "#ffffff";
  var xlZebraEven = isDark ? "#1f1f1f" : "#f7f9fc";

  // ── Render project info banner ────────────────────────────
  var bBg       = isDark ? "#16213e" : "#2e5fa3";
  var bBorder   = isDark ? "#2a3a6a" : "#1e3f7a";
  var bLabel    = isDark ? "#7a9cc8" : "rgba(255,255,255,0.7)";
  var bValue    = isDark ? "#e0eeff" : "#ffffff";
  var bDivider  = isDark ? "rgba(255,255,255,0.1)" : "rgba(255,255,255,0.25)";
  var bFont     = "font-family:Calibri,Arial,sans-serif;";

  function bannerField(label, value) {
    if (!value) return "";
    return '<div style="display:flex;flex-direction:column;justify-content:center;padding:0 20px;border-right:1px solid ' + bDivider + ';min-width:100px;">' +
      '<div style="' + bFont + 'font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:0.08em;color:' + bLabel + ';margin-bottom:3px;">' + label + '</div>' +
      '<div style="' + bFont + 'font-size:13px;font-weight:700;color:' + bValue + ';white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:200px;">' + h(value) + '</div>' +
    '</div>';
  }

  var lpInfo         = leftPanel.project_info  || [];
  var lpStakeholders = leftPanel.stakeholders  || [];
  var lp_project_desc = (lpInfo[4]         || {}).value || "";
  var lp_po_value     = (lpInfo[0]         || {}).value || "";
  var lp_sw_efforts   = (lpInfo[6]         || {}).value || "";
  var lp_pm           = (lpStakeholders[1] || {}).value || "";
  var ldHasData = !!(banner.ld_date_val || banner.ld_maxwk_val || banner.ld_maxov_val || banner.ld_remarks_val);

  var bannerHtml = '<div style="position:sticky;top:0;z-index:10;flex-shrink:0;">';
  bannerHtml +=     '<div style="display:flex;align-items:stretch;height:64px;background:' + bBg + ';border-bottom:3px solid ' + bBorder + ';overflow:visible;position:relative;">';
  bannerHtml +=       '<div style="display:flex;flex-direction:column;justify-content:center;padding:0 16px 0 14px;border-right:1px solid ' + bDivider + ';flex-shrink:0;">';
  bannerHtml +=         '<div style="' + bFont + 'font-size:16px;font-weight:800;color:' + bValue + ';letter-spacing:-0.02em;white-space:nowrap;">' + h(banner.or_number || "") + '</div>';
  bannerHtml +=         '<div style="' + bFont + 'font-size:11px;font-weight:500;color:' + bLabel + ';margin-top:3px;max-width:240px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="' + h(lp_project_desc) + '">' + h(lp_project_desc) + '</div>';
  bannerHtml +=       '</div>';
  bannerHtml +=       '<div style="display:flex;flex-direction:column;justify-content:center;padding:0 20px;border-right:1px solid ' + bDivider + ';flex-shrink:0;gap:6px;">';
  bannerHtml +=         '<div style="display:flex;align-items:center;gap:8px;"><div style="' + bFont + 'font-size:9px;font-weight:700;text-transform:uppercase;color:' + bLabel + ';white-space:nowrap;">Sales Engineer</div><div style="' + bFont + 'font-size:13px;font-weight:700;color:' + bValue + ';white-space:nowrap;">' + h(banner.sales_engineer || "") + '</div></div>';
  bannerHtml +=         '<div style="display:flex;align-items:center;gap:8px;"><div style="' + bFont + 'font-size:9px;font-weight:700;text-transform:uppercase;color:' + bLabel + ';white-space:nowrap;">PM</div><div style="' + bFont + 'font-size:13px;font-weight:700;color:' + bValue + ';white-space:nowrap;">' + h(lp_pm) + '</div></div>';
  bannerHtml +=       '</div>';
  bannerHtml +=       '<div style="display:flex;flex-direction:column;justify-content:center;padding:0 20px;border-right:1px solid ' + bDivider + ';flex-shrink:0;gap:6px;">';
  bannerHtml +=         '<div style="display:flex;align-items:center;gap:8px;"><div style="' + bFont + 'font-size:9px;font-weight:700;text-transform:uppercase;color:' + bLabel + ';white-space:nowrap;">PO Value</div><div style="' + bFont + 'font-size:13px;font-weight:700;color:' + bValue + ';white-space:nowrap;">' + h(lp_po_value) + '</div></div>';
  bannerHtml +=         '<div style="display:flex;align-items:center;gap:8px;"><div style="' + bFont + 'font-size:9px;font-weight:700;text-transform:uppercase;color:' + bLabel + ';white-space:nowrap;">SW Efforts</div><div style="' + bFont + 'font-size:13px;font-weight:700;color:' + bValue + ';white-space:nowrap;">' + h(lp_sw_efforts) + '</div></div>';
  bannerHtml +=       '</div>';
   
    

    // LD Section
  var ldChecked = !!(banner.ld_date_val);  // Check if LD date exists
  
  if (ldChecked) {
    // Container for LD SECTION (left side)
    bannerHtml += '<div style="display:flex;flex-direction:column;justify-content:center;padding:0 16px;flex-shrink:0;border-left:1px solid ' + bDivider + ';">';
    
    // Row 1: Checkbox only
    bannerHtml += '<div style="display:flex;align-items:center;margin-bottom:6px;">';
    bannerHtml += '<div style="display:flex;align-items:center;gap:6px;">';
    bannerHtml += '<input type="checkbox" id="ld-checkbox" checked disabled style="width:14px;height:14px;cursor:default;opacity:0.7;">';
    bannerHtml += '<span style="' + bFont + 'font-size:11px;font-weight:600;color:' + bLabel + ';">LD Applicable</span>';
    bannerHtml += '</div>';
    bannerHtml += '</div>';
    
    // LD Date (formatted) + Button in the same row
    var ldDateFormatted = '';
    if (banner.ld_date_val) {
        var dateStr = banner.ld_date_val;
        var dateParts = dateStr.split(' ')[0].split('-');
        if (dateParts.length === 3) {
            var year = dateParts[0];
            var month = dateParts[1];
            var day = dateParts[2];
            var monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
            var monthName = monthNames[parseInt(month) - 1];
            ldDateFormatted = parseInt(day) + ' ' + monthName + ' ' + year;
        } else {
            ldDateFormatted = dateStr;
        }
    }
    
    bannerHtml += '<div style="display:flex;align-items:center;gap:10px;">';
    bannerHtml += '<div style="' + bFont + 'font-size:12px;font-weight:500;color:' + bValue + ';">' + h(ldDateFormatted) + '</div>';
    bannerHtml += '<button id="ld-toggle-btn" style="background:transparent;border:none;cursor:pointer;font-size:14px;color:' + bLabel + ';padding:2px 6px;">→</button>';
    bannerHtml += '</div>';
    
    bannerHtml += '</div>'; // End LD SECTION container
    
    // DETAILS SECTION container (initially hidden)
    bannerHtml += '<div id="ld-additional-details" style="display:none;align-items:flex-start;padding:0 16px;flex-shrink:0;">';
    bannerHtml += '<div style="display:flex;gap:24px;">';
    
        // Left column: Max/Wk and Max/OV (stacked vertically) - convert to percentages
    function formatPercentage(val) {
        if (val === undefined || val === null) return '';
        var num = parseFloat(val);
        if (isNaN(num)) return val;
        // Convert decimal to percentage (0.005 -> 0.5%)
        if (num >= 0 && num <= 1) {
            return (num * 100) + '%';
        }
        // Already a number like 5 -> 5%
        return num + '%';
    }
    
    var maxwkDisplay = formatPercentage(banner.ld_maxwk_val);
    var maxovDisplay = formatPercentage(banner.ld_maxov_val);
    
    bannerHtml += '<div style="display:flex;flex-direction:column;gap:6px;margin-top:8px;">';
    if (banner.ld_maxwk_val)   bannerHtml += '<div style="display:flex;align-items:center;gap:6px;"><div style="' + bFont + 'font-size:9px;font-weight:700;text-transform:uppercase;color:' + bLabel + ';">' + (banner.ld_maxwk_lbl||"LD Max/Wk") + '</div><div style="' + bFont + 'font-size:12px;font-weight:500;color:' + bValue + ';">' + h(maxwkDisplay) + '</div></div>';
    if (banner.ld_maxov_val)   bannerHtml += '<div style="display:flex;align-items:center;gap:6px;"><div style="' + bFont + 'font-size:9px;font-weight:700;text-transform:uppercase;color:' + bLabel + ';">' + (banner.ld_maxov_lbl||"LD Max/OV") + '</div><div style="' + bFont + 'font-size:12px;font-weight:500;color:' + bValue + ';">' + h(maxovDisplay) + '</div></div>';
    bannerHtml += '</div>';
    
    // Right column: Remarks (with word wrapping)
    if (banner.ld_remarks_val) {
      bannerHtml += '<div style="display:flex;flex-direction:column;gap:6px;max-width:250px;margin-top:9px;">';
            // Convert remarks value to percentage if it's a number
      var remarksValue = banner.ld_remarks_val;
      var displayValue = remarksValue;
      
      // Check if it's a decimal number between 0 and 1 (like 0.005 = 0.5%)
      if (typeof remarksValue === 'number' || !isNaN(parseFloat(remarksValue))) {
        var num = parseFloat(remarksValue);
        if (num >= 0 && num <= 1) {
          // Convert decimal to percentage (0.005 -> 0.5%)
          displayValue = (num * 100) + '%';
        } else if (num > 1 && num <= 100) {
          // Already a percentage number (5 -> 5%)
          displayValue = num + '%';
        }
      }

      bannerHtml += '<div style="display:flex;align-items:flex-start;gap:6px;">';
      bannerHtml += '<div style="' + bFont + 'font-size:9px;font-weight:700;text-transform:uppercase;color:' + bLabel + ';white-space:nowrap;">' + (banner.ld_remarks_lbl||"LD Remarks") + '</div>';
      bannerHtml += '<div style="' + bFont + 'font-size:12px;font-weight:500;color:' + bValue + ';word-wrap:break-word;white-space:normal;line-height:1.4;max-width:500px;display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden;">' + h(displayValue) + '</div>';
      bannerHtml += '</div>';
      bannerHtml += '</div>';
    }
    
    bannerHtml += '</div>'; // End inner flex
    bannerHtml += '</div>'; // End DETAILS SECTION container
  }
  
  bannerHtml +=   '</div>';
  bannerHtml += '</div>';
  var xlCellBg   = isDark ? "#1e1e1e" : "#ffffff";
  var xlHeaderBg = isDark ? "#2a2a2a" : "#f2f2f2";
  var xlCornerBg = isDark ? "#242424" : "#e8e8e8";
  var xlBorder   = isDark ? "#3a3a3a" : "#3a3a3a";
  var xlCellBorder = isDark ? "#2e2e2e" : "#2e2e2e";
  var xlText     = isDark ? "#1a1a1a" : "#1a1a1a";
  var xlRowNumColor = isDark ? "#666666" : "#999999";
  var xlGroupBarBg = isDark ? "#1a2620" : "#e8f5ee";
  var xlGroupBarBorder = isDark ? "#2a4a38" : "#b8d8c8";

  function buildTable() {
   // Calculate total percentage width of ALL columns (including hidden/collapsed ones)
    // This ensures the table expands beyond 100% when groups are expanded
    // Sum ONLY visible (non-collapsed) columns — mirrors monitor.js computeOverviewTableWidth()
// This makes the table grow when groups expand, so existing columns never shift
    var totalWidthPct = 0;
    for (var ci = 0; ci < cols.length; ci++) {
      var colCheck = cols[ci];
      if (isSkippedCol(colCheck)) continue;
      if (isCollapsedCol(colCheck)) continue;  // ← skip hidden cols
      var w = getSheetColumnWidth(colCheck);
      var pctVal = parseFloat(w);
      if (!isNaN(pctVal)) totalWidthPct += pctVal;
    }
    totalWidthPct += 3; // row number column buffer
    // If visible columns fit within 100%, keep table at 100% (no scroll)
    // If they exceed it, let the table grow so new cols push right
    var tableWidthVal = totalWidthPct <= 100 ? "100%" : totalWidthPct + "%";

    var tableStyle = [
      "border-collapse:collapse",
      "table-layout:fixed",
      "font-family:Calibri,Arial,sans-serif",
      "font-size:11px",
      "background:" + xlBg,
      "color:" + xlText,
      "width:" + tableWidthVal + "%",  // ✅ Table width expands based on ALL columns
    ].join(";");

    // ── Build left panel HTML ────────────────────────────────
    var lpBg      = isDark ? "#161616" : "#f0f4f8";
    var lpBorder  = isDark ? "#2a2a2a" : "#d0d8e4";
    var lpLabel   = isDark ? "#7a7777" : "#888888";
    var lpValue   = isDark ? "#c2bfbf" : "#333333";
    var lpSection = isDark ? "#333333" : "#d8e2ee";
    var lpFont    = "font-family:Calibri,Arial,sans-serif;";

    function lpRow(label, value, wrap) {
      if (!label) return "";
      var valueStyle = wrap
        ? lpFont + 'font-size:11px;font-weight:400;color:' + lpValue + ';white-space:normal;word-break:break-word;'
        : lpFont + 'font-size:11px;font-weight:400;color:' + lpValue + ';text-align:right;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:100px;cursor:default;';
      return '<div style="display:flex;justify-content:space-between;align-items:baseline;padding:2px 10px;gap:8px;">' +
        '<span style="' + lpFont + 'font-size:11px;font-weight:600;color:' + lpLabel + ';white-space:nowrap;">' + h(label) + '</span>' +
        '<span title="' + h(value || "") + '" style="' + valueStyle + '">' + h(value || "—") + '</span>' +
      '</div>';
    }

    function lpSectionHeader(title) {
      return '<div style="padding:4px 10px 2px;background:' + lpSection + ';font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;color:' + lpLabel + ';border-top:1px solid ' + lpBorder + ';border-bottom:1px solid ' + lpBorder + ';">' + title + '</div>';
    }

    var lpId = "lp-" + Date.now();
    var leftPanelHtml = '<div id="' + lpId + '" style="width:200px;min-width:200px;flex-shrink:0;background:' + lpBg + ';border-right:1px solid ' + lpBorder + ';display:flex;flex-direction:column;overflow:hidden;transition:width 0.2s,min-width 0.2s;position:relative;">';
    leftPanelHtml += '<button id="' + lpId + '-btn" title="Toggle panel" onclick="(function(){'
      + 'var p=document.getElementById(\'' + lpId + '\');'
      + 'var inner=document.getElementById(\'' + lpId + '-inner\');'
      + 'var btn=document.getElementById(\'' + lpId + '-btn\');'
      + 'var isCollapsed=p.getAttribute(\'data-collapsed\')==\'1\';'
      + 'if(isCollapsed){'
      +   'p.style.width=\'200px\';p.style.minWidth=\'200px\';'
      +   'inner.style.display=\'block\';'
      +   'btn.innerHTML=\'&#8249;\';'
      +   'btn.style.cssText=\'position:absolute;top:4px;right:4px;z-index:6;width:18px;height:18px;border-radius:3px;background:' + lpSection + ';border:1px solid ' + lpBorder + ';cursor:pointer;color:' + lpLabel + ';font-size:13px;font-weight:700;display:flex;align-items:center;justify-content:center;\';'
      +   'p.setAttribute(\'data-collapsed\',\'0\');'
      + '}else{'
      +   'p.style.width=\'24px\';p.style.minWidth=\'24px\';'
      +   'inner.style.display=\'none\';'
      +   'btn.innerHTML=\'&#8250;\';'
      +   'btn.style.cssText=\'position:absolute;top:4px;left:3px;z-index:6;width:18px;height:18px;border-radius:3px;background:' + lpSection + ';border:1px solid ' + lpBorder + ';cursor:pointer;color:' + lpLabel + ';font-size:13px;font-weight:700;display:flex;align-items:center;justify-content:center;\';'
      +   'p.setAttribute(\'data-collapsed\',\'1\');'
      + '}'
      + '})()" style="position:absolute;top:4px;right:4px;z-index:6;width:18px;height:18px;border-radius:3px;background:' + lpSection + ';border:1px solid ' + lpBorder + ';cursor:pointer;color:' + lpLabel + ';font-size:13px;font-weight:700;display:flex;align-items:center;justify-content:center;">&#8249;</button>';
    leftPanelHtml += '<div id="' + lpId + '-inner" style="overflow-y:auto;flex:1;padding-top:26px;">';
    // Project info — hide 0 (PO Value), 1 (Customer), 4 (Project Desc), 6 (SW Efforts) — shown in banner
    leftPanelHtml += lpSectionHeader("Project Info");
    var removedPI = new Set([0, 1, 4, 6]);
    (leftPanel.project_info || []).forEach(function(item, idx) {
      if (removedPI.has(idx)) return;
      var wrap = (idx === 2 || idx === 3); // End Customer, Consultant
      leftPanelHtml += lpRow(item.label, item.value, wrap);
    });
    // Dates — two column table (PM | SWE) with Excel cell colors
    leftPanelHtml += lpSectionHeader("Dates");
    leftPanelHtml += '<table style="width:100%;border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:10px;">'
      + '<thead><tr>'
      + '<th style="padding:2px 6px;color:' + lpLabel + ';font-weight:700;text-align:left;border-bottom:1px solid ' + lpBorder + ';"></th>'
      + '<th style="padding:2px 4px;color:' + lpLabel + ';font-weight:700;text-align:center;border-bottom:1px solid ' + lpBorder + ';font-size:9px;">PM</th>'
      + '<th style="padding:2px 4px;color:' + lpLabel + ';font-weight:700;text-align:center;border-bottom:1px solid ' + lpBorder + ';font-size:9px;">SWE</th>'
      + '</tr></thead><tbody>';
    (leftPanel.dates || []).forEach(function(item) {
      var pmBg  = (item.pm_fill  && item.pm_fill  !== "#FFFFFF") ? item.pm_fill  : "transparent";
      var swBg  = (item.swe_fill && item.swe_fill !== "#FFFFFF") ? item.swe_fill : "transparent";
      
      // In dark mode, force light text color; in light mode, use Excel font or default
      var pmFg = isDark ? "#c2bfbf" : (item.pm_font || lpValue);
      var swFg = isDark ? "#c2bfbf" : (item.swe_font || lpValue);
      var pmVal  = item.pm  || item.value || "—";
      var sweVal = item.swe || "—";
      leftPanelHtml += '<tr style="border-bottom:1px solid ' + lpBorder + ';">'
        + '<td style="padding:2px 6px;color:' + lpLabel + ';font-weight:600;white-space:nowrap;">' + h(item.label) + '</td>'
        + '<td style="padding:2px 3px;text-align:center;background:' + pmBg + ';color:' + pmFg + ';font-weight:700;white-space:nowrap;" title="' + h(pmVal) + '">' + h(pmVal) + '</td>'
        + '<td style="padding:2px 3px;text-align:center;background:' + swBg + ';color:' + swFg + ';font-weight:700;white-space:nowrap;" title="' + h(sweVal) + '">' + h(sweVal) + '</td>'
        + '</tr>';
    });
    leftPanelHtml += '</tbody></table>';
    // Stakeholders — hide Sales (idx 0), PM (idx 1), and logged-in user's own row
    var userShort = ((state.user && state.user.short_name) || "").toUpperCase();
    leftPanelHtml += lpSectionHeader("Stakeholders");
    (leftPanel.stakeholders || []).forEach(function(item, idx) {
      if (idx === 0 || idx === 1) return;
      if (userShort && (item.value || "").toUpperCase() === userShort) return;
      leftPanelHtml += lpRow(item.label, item.value);
    });
    // Warranty
    if ((leftPanel.warranty || []).some(function(w){ return w.value; })) {
      leftPanelHtml += lpSectionHeader("Warranty");
      (leftPanel.warranty || []).forEach(function(item) {
        leftPanelHtml += lpRow(item.label, item.value, true);
      });
    }
    // Extra sections — config-driven (e.g. PM's VA/SM actuals block)
    // Server sends leftPanel.extra_sections[] matching schedCfg.leftPanel.extraSections
    var _extraSections = leftPanel.extra_sections || [];
    _extraSections.forEach(function(section) {
      if (!section || !section.rows || !section.rows.length) return;
      var hasAnyValue = section.rows.some(function(row) { return row.value || row.actual; });
      if (!hasAnyValue) return;
      leftPanelHtml += lpSectionHeader(section.title || "");
      section.rows.forEach(function(row) {
        if (row.actual !== undefined && row.actual !== null && row.actual !== "") {
          // Two-value row (estimated vs actual) — show as "est → act"
          leftPanelHtml += lpRow(row.label, (row.value || "—") + " → " + (row.actual || "—"));
        } else {
          leftPanelHtml += lpRow(row.label, row.value, row.wrap);
        }
      });
    });
    leftPanelHtml += '</div>'; // end inner
    leftPanelHtml += '</div>'; // end panel

    var html = '<div style="display:flex;flex-direction:column;flex: 1;min-height: 0;position:relative;background:' + xlBg + ';overflow:hidden;">';
    html += '<div style="flex-shrink:0;overflow:hidden;">' + bannerHtml + '</div>';
    html += '<div style="display:flex;flex:1;min-height:0;overflow:hidden;">';
    html += leftPanelHtml;
    html += '<div style="overflow-x:auto;overflow-y:auto;flex:1;position:relative;">';
    html += '<table style="' + tableStyle + '">';
    html += '<thead>';

    if (colGroups.length > 0) {
      // ── GROUP BAR ROW (buttons with borders) ──
      html += '<tr style="height:20px;">';

      var ci = 0;
      while (ci < cols.length) {
        var col = cols[ci];
        if (isSkippedCol(col)) { ci++; continue; }
        var gi  = colToGroup[col];

        if (gi !== undefined && gi >= 0 && cols[ci] === colGroups[gi].cols[0]) {
          var group     = colGroups[gi];
          var collapsed = collapseState[gi];
          var btnLabel  = collapsed ? "+" : "–";
          var btnStyle  = "cursor:pointer;font-size:11px;font-weight:700;padding:0px 6px;border-radius:3px;border:1px solid " + (isDark ? "#555" : "#aaa") + ";background:" + (isDark ? "#2a2a2a" : "#e8e8e8") + ";color:" + (isDark ? "#ddd" : "#333") + ";line-height:1.4;";
          
          if (!collapsed) {
            // Expanded: show button above the first column of the group
            var firstW = getSheetColumnWidth(group.cols[0]);  // ✅ CHANGED: use percentage
            html += '<th style="position:sticky;top:0;z-index:4;background:transparent;border:none;text-align:left;padding:0 0 0 6px;width:' + firstW + ';">';  // ✅ CHANGED: removed "px"
            html += '<button style="' + btnStyle + '" onclick="__xlToggleGroup(' + gi + ')">' + btnLabel + '</button>';
            html += '</th>';
            // Empty cells for the remaining columns in the group
            for (var gci = 1; gci < group.cols.length; gci++) {
              var gcw = getSheetColumnWidth(group.cols[gci]);  // ✅ CHANGED: use percentage
              html += '<th style="position:sticky;top:0;z-index:4;background:transparent;border:none;width:' + gcw + ';"></th>';  // ✅ CHANGED: removed "px"
            }
          } else {
            // Collapsed: show button above the last column of the group (the only visible one)
            var lastCol = group.cols[group.cols.length - 1];
            var lastW   = getSheetColumnWidth(lastCol);  // ✅ CHANGED: use percentage
            html += '<th style="position:sticky;top:0;z-index:4;background:transparent;border:none;text-align:left;padding:0 0 0 6px;width:' + lastW + ';">';  // ✅ CHANGED: removed "px"
            html += '<button style="' + btnStyle + '" onclick="__xlToggleGroup(' + gi + ')">' + btnLabel + '</button>';
            html += '</th>';
          }
          ci += group.cols.length;
        } else {
          var cw = getSheetColumnWidth(col);  // ✅ CHANGED: use percentage
          html += '<th style="position:sticky;top:0;z-index:4;background:transparent;border:none;width:' + cw + ';"></th>';  // ✅ CHANGED: removed "px"
          ci++;
        }
      }
      html += '<\/tr>';
    }

    html += '</thead>';

    // ── TBODY ──
    html += '</thead>';

    // ── Custom header row — driven by schedule_config.js ─────
    var chDefs = schedCfg ? (isDark ? schedCfg.customHeaders.dark : schedCfg.customHeaders.light) : {};
    var chRowBg    = isDark ? "#141414" : "#e8edf5";
    var chBorder   = isDark ? "#333333" : "#c0c7d8";
    var chFallbackBg = isDark ? "#1a1a2e" : "#e8edf5";
    var chFallbackFg = isDark ? "#888888" : "#555555";

    html += '<thead id="xl-custom-header">';
    html += '<tr style="height:26px;">';
    for (var chi = 0; chi < cols.length; chi++) {
      var chCol = cols[chi];
      if (isSkippedCol(chCol)) continue;
      if (isCollapsedCol(chCol)) continue;
      var chCw = getSheetColumnWidth(chCol);
      var chDef = chDefs[chCol] || {};
      var chBg = chDef.bg || chFallbackBg;
      var chFg = chDef.fg || chFallbackFg;
      var chLabel = chDef.label || chCol;
      var chStyle = [
        "width:" + chCw,
        "min-width:" + chCw,
        "background:" + chBg,
        "color:" + chFg,
        "font-family:Calibri,Arial,sans-serif",
        "font-size:10px",
        "font-weight:700",
        "text-align:center",
        "padding:2px 3px",
        "border:1px solid " + chBorder,
        "white-space:nowrap",
        "overflow:hidden",
        "text-overflow:ellipsis",
        "letter-spacing:0.02em",
        "position:sticky",
        "top:0",
        "z-index:4",
      ].join(";");
      html += '<th style="' + chStyle + '">' + chLabel + '</th>';
    }
    html += '</tr>';
    html += '</thead>';

    html += '<tbody>';
    // Green color for Z cell when 100% complete (from Excel CF rule $Z9>=100%)
    
    // Find the last row that has task data (task_code and task_name not empty)
    // Find the last row that has task data (task_code and task_name not empty)
    var maxTaskRow = maxRow;
    for (var r = maxRow; r >= TASK_START_ROW; r--) {
        var taskCodeCell = cells["H" + r];
        var taskNameCell = cells["I" + r];
        var taskCode = taskCodeCell ? taskCodeCell.v : null;
        var taskName = taskNameCell ? taskNameCell.v : null;
        if (taskCode && taskName && taskCode.toString().trim() !== "" && taskName.toString().trim() !== "") {
            maxTaskRow = r;
            break;
        }
    }
    for (var r = 1; r <= maxTaskRow; r++) {
      // Skip info rows (1-5) — shown in banner instead
      if (infoRowSet[r]) continue;
      // Skip Excel header rows 6, 7, 8 — replaced by custom header above
      if (r === 6 || r === 7 || r === 8) continue;

      var rh = rowHeights[String(r)] || 20;
      var isComplete = false;
      var zCoord = "Z" + r;
      var zInfo  = cells[zCoord];
      if (zInfo) {
        var zVal = String(zInfo.v || "");
        isComplete = (zVal === "100%" || zVal === "100");
      }

      var dataRowIndex = r - infoRows.length;
      var isEvenDataRow = (dataRowIndex % 2 === 0);

      var rowBg = isEvenDataRow ? xlZebraEven : xlZebraOdd;
      html += '<tr style="height:' + rh + 'px;background:' + rowBg + '">';

      for (var ci3 = 0; ci3 < cols.length; ci3++) {
        var col3  = cols[ci3];

        if (isSkippedCol(col3)) continue;

        var coord = col3 + r;
        var info  = cells[coord];

        // skip all cols in a collapsed group
        if (isCollapsedCol(col3)) continue;

        // merged slave — skip
        if (info && info.skip) continue;

        var cw3 = getSheetColumnWidth(col3);
        var tdStyle = [
          "width:" + cw3,
          "min-width:" + cw3,
          "overflow:hidden",
          "padding:1px 3px",
          "border:1px solid " + xlCellBorder,
          "vertical-align:bottom",
          "white-space:nowrap",
        ];
        // ── EDITABILITY — computed once, used by both bg and input rendering ──
        var isReadOnly = state.user && (state.user.role === "admin" || state.user.role === "head" || (state.user.role && state.user.role.endsWith("_head")));
        // isADColumn: only fires when dept has a helpDropdown with a designated col
        var isADColumn = (AD_COL !== null && col3 === AD_COL && r >= TASK_START_ROW);
        // A col marked "readonly" in editableCols is never editable regardless of server flag
        var _editColCfgRow = schedCfg && schedCfg.editableCols ? schedCfg.editableCols[col3] : null;
        var _isConfigReadonly = _editColCfgRow && _editColCfgRow.type === "readonly";
        var isEditable = false;
        if (_isConfigReadonly) {
            isEditable = false;
        } else if (isADColumn) {
            // Help-dropdown editability: editable only when % complete < 100
            var _zInfoAD = cells["Z" + r];
            var _zValAD  = _zInfoAD ? String(_zInfoAD.v || "") : "";
            var _zDoneAD = (_zValAD === "100%" || _zValAD === "100");
            isEditable = !isReadOnly && !_zDoneAD;
        } else {
            isEditable = !isReadOnly && !!(info && info.editable);
        }

        // ── BACKGROUND COLOR LOGIC ──────────────────────────────
        // Shared Z-complete flag (used for Z green and AB yellow)
        var _zInfoRow = cells["Z" + r];
        var _zValRow  = _zInfoRow ? String(_zInfoRow.v || "") : "";
        var isZComplete = (col3 === "Z" && r >= 9) && (_zValRow === "100%" || _zValRow === "100");

        // AB yellow: Z % complete < AA expected % complete
        var isABBehind = false;
        if (col3 === "AB" && r >= 9) {
            var _aaInfo = cells["AA" + r];
            var _zNum   = parseFloat(_zValRow.replace("%", "").trim());
            var _aaNum  = parseFloat(_aaInfo ? String(_aaInfo.v || "").replace("%", "").trim() : "");
            isABBehind = !isNaN(_zNum) && !isNaN(_aaNum) && _zNum < _aaNum;
        }

        var customBg = getCustomBackgroundColor(col3, r);
        var bgColor  = null;

        // --- HELP DROPDOWN COLUMN (e.g. AD in SW) ---
        if (isADColumn) {
            var _adInfo       = cells[AD_COL + r];
            var _adValue      = _adInfo ? String(_adInfo.v || "").trim() : "";
            var _adValueLower = _adValue.toLowerCase();
            var _adColor      = AD_COLORS[_adValueLower];

            if (_adColor) {
                bgColor = _adColor;
                // Dark mode: use per-dept fg from config; light mode: white
                tdStyle.push("color:" + (isDark ? (AD_FG_DARK[_adValueLower] || "#ffffff") : "#ffffff"));
            } else if (isEditable) {
                bgColor = isDark ? "#1a3a5a" : "#b8d8ff";
                tdStyle.push("color:" + (isDark ? "#88ccff" : "#0066cc"));
            } else {
                bgColor = rowBg;
            }
        }
        // --- NON-AD COLUMNS ---
        else {
            if (isZComplete) {
                bgColor = isDark ? "#1a3a1a" : "#92d050";
                if (isDark) tdStyle.push("color:#4ade80");
            } else if (isABBehind) {
                // AB column behind schedule: yellow background
                bgColor = isDark ? "#3a2e00" : "#fff176";
            } else if (customBg) {
                bgColor = customBg;
                if (isDark) tdStyle.push("color:#ffffff");
            } else {
                bgColor = rowBg;
            }
        }

        // Apply background color (single push — duplicate removed)
        if (bgColor) {
            tdStyle.push("background-color:" + bgColor);
        }

        var content = "";
        var textRuleColor = getTextRuleColor(col3, r, cells);

        if (info) {
          if (info.font) {
            var f = info.font;
            if (f.bold)      tdStyle.push("font-weight:bold");
            if (f.italic)    tdStyle.push("font-style:italic");
            if (f.underline) tdStyle.push("text-decoration:underline");
            if (f.size)      tdStyle.push("font-size:" + f.size + "px");
            if (f.color)     tdStyle.push("color:" + f.color);
          }
          if (textRuleColor) tdStyle.push("color:" + textRuleColor);
          if (info.align) {
            var a = info.align;
            if (a.h === "center" || a.h === "centerContinuous") tdStyle.push("text-align:center");
            else if (a.h === "right") tdStyle.push("text-align:right");
            else if (a.h === "left")  tdStyle.push("text-align:left");
            if (a.v === "center")     tdStyle.push("vertical-align:middle");
            else if (a.v === "top")   tdStyle.push("vertical-align:top");
            if (a.wrap)               tdStyle.push("white-space:pre-wrap");
          }
          if (info.border) {
            var b = info.border;
            if (b.left)   tdStyle.push("border-left:"   + borderStyle(b.left));
            if (b.right)  tdStyle.push("border-right:"  + borderStyle(b.right));
            if (b.top)    tdStyle.push("border-top:"    + borderStyle(b.top));
            if (b.bottom) tdStyle.push("border-bottom:" + borderStyle(b.bottom));
          }

          if (isEditable) {
            tdStyle.push("padding:0");
            content = renderInputCell(info, coord, col3, tdStyle, textRuleColor);
          } else if (info.v !== null && info.v !== undefined) {
            content = h(String(info.v));
          }
        } else {
          // No cell info
          content = "";
        }

        var attrs = ' style="' + tdStyle.join(";") + '"';
        if (info && info.rowspan && info.rowspan > 1) attrs += ' rowspan="' + info.rowspan + '"';
        if (info && info.colspan && info.colspan > 1) attrs += ' colspan="' + info.colspan + '"';
        html += '<td' + attrs + '>' + content + '</td>';
      }
      html += '</tr>';
    }
    html += '</tbody></table></div></div></div>';

      // Add click handler for LD toggle button (will run after HTML is rendered)
    setTimeout(function() {
      var toggleBtn = document.getElementById("ld-toggle-btn");
      var detailsDiv = document.getElementById("ld-additional-details");
      if (toggleBtn && detailsDiv) {
        toggleBtn.addEventListener("click", function() {
          if (detailsDiv.style.display === "none") {
            detailsDiv.style.display = "block";
            toggleBtn.textContent = "←";
          } else {
            detailsDiv.style.display = "none";
            toggleBtn.textContent = "→";
          }
        });
      }
    }, 100);
    container.innerHTML = html;
  }

  // ── Edit handler — fires when user changes an input cell ──
  window.__xlEditCell = function(el) {
    var coord = el.getAttribute("data-coord");
    var col   = el.getAttribute("data-col");
    var val   = el.value;
    _editPending[coord] = val;

    // Extract row number from coord (e.g. "X12" -> 12)
    var row = parseInt(coord.replace(/[A-Z]+/, ""), 10);

    // Build field map from editableCols config — maps col letter to task field name.
    // Any col not in this map is silently ignored (won't produce a pending change).
    // "readonly" cols are excluded — they can never produce a change.
    var fieldMap = {};
    if (schedCfg && schedCfg.editableCols) {
      // Static known mappings — col → server field name
      var _knownFields = {
        "X":  "actual_start",
        "Y":  "actual_end",
        "Z":  "percent_complete",
        "AF": "remark",
      };
      // Add help-dropdown col if present
      if (AD_COL) _knownFields[AD_COL] = "help_required";

      Object.keys(schedCfg.editableCols).forEach(function(c) {
        var cfg = schedCfg.editableCols[c];
        if (cfg.type !== "readonly" && _knownFields[c]) {
          fieldMap[c] = _knownFields[c];
        }
      });
    }

    var field = fieldMap[col];
    if (!field) return;

    var taskKey = "__row_" + row;
    if (!state.pendingChanges[taskKey]) {
      state.pendingChanges[taskKey] = { _row: row };
    }
    state.pendingChanges[taskKey][field] = val;
    updateSaveBar();
  };

  // Assign a stable id to container so toggle can always find it
  if (!container.id) container.id = "xlm-grid-" + Date.now();
  var containerId = container.id;

  // Toggle handler
  window.__xlToggleGroup = function(gi) {
    collapseState[gi] = !collapseState[gi];
    var c = document.getElementById(containerId);
    if (c) { container = c; buildTable(); }
  };

  buildTable();
}

function onXLPct(input) {
  var val    = Math.max(0, Math.min(100, parseInt(input.value) || 0));
  var key    = input.getAttribute("data-key");
  var owner  = input.getAttribute("data-owner");
  var orig = state.tasks.find(function(t) { return taskKey(t) === key; });
  var pending = state.pendingChanges[key] || (orig ? {
    percent_complete: orig.percent_complete,
    actual_start: orig.actual_start,
    actual_end:   orig.actual_end,
    owner:        owner,
    task_key:     key
  } : null);
  if (!pending) return;
  pending.percent_complete = val;
  pending.task_key = key;
  pending.owner    = owner;
  state.pendingChanges[key] = pending;
  _checkClean(key, orig);
  updateSaveBar();
}

function onXLDate(input) {
  var key    = input.getAttribute("data-key");
  var field  = input.getAttribute("data-field");
  var owner  = input.getAttribute("data-owner");
  var val    = input.value || null;
  var orig   = state.tasks.find(function(t) { return taskKey(t) === key; });
  var pending = state.pendingChanges[key] || (orig ? {
    percent_complete: orig.percent_complete,
    actual_start: orig.actual_start,
    actual_end:   orig.actual_end,
    owner:        owner,
    task_key:     key
  } : null);
  if (!pending) return;
  pending[field]   = val;
  pending.task_key = key;
  pending.owner    = owner;
  state.pendingChanges[key] = pending;
  _checkClean(key, orig);
  updateSaveBar();
}

function _checkClean(key, orig) {
  if (!orig) return;
  var p = state.pendingChanges[key];
  if (p &&
      p.percent_complete === orig.percent_complete &&
      p.actual_start     === (orig.actual_start||null) &&
      p.actual_end       === (orig.actual_end||null)) {
    delete state.pendingChanges[key];
  }
}

function onSlider(input, safeKey, key, owner) {
  var val = parseInt(input.value);
  var color = ringColor(val);
  var C = 2 * Math.PI * 14;
  var offset = C - (val / 100) * C;
  var ring  = document.getElementById("ring-" + safeKey);
  var pctEl = document.getElementById("rpct-" + safeKey);
  var svEl  = document.getElementById("sv-" + safeKey);
  if (ring)  { ring.setAttribute("stroke-dashoffset", offset.toFixed(2)); ring.setAttribute("stroke", color); }
  if (pctEl) pctEl.textContent = val + "%";
  if (svEl)  svEl.textContent  = val + "%";
  var row = input.closest("tr");
  var remarkEl = row ? row.querySelector(".remark-input") : null;
  var remark = remarkEl ? remarkEl.value : (state.pendingChanges[key] ? state.pendingChanges[key].remark : "");
  markChange(key, val, remark, owner);
}

function onRemark(textarea, key, owner) {
  var found = state.tasks.find(function(t) { return taskKey(t) === key; });
  var pct = state.pendingChanges[key] ? state.pendingChanges[key].percent_complete : (found ? found.percent_complete : 0);
  markChange(key, pct, textarea.value, owner);
}

function markChange(key, pct, remark, owner) {
  var orig = state.tasks.find(function(t) { return taskKey(t) === key; });
  var origRemark = orig ? (orig.remark||"") : "";
  if (orig && orig.percent_complete === pct && origRemark === remark) {
    delete state.pendingChanges[key];
  } else {
    state.pendingChanges[key] = {percent_complete:pct, remark:remark, owner:owner, task_key:key};
  }
  updateSaveBar();
}

function updateSaveBar() {
  var count = Object.keys(state.pendingChanges).length;
  var bar   = document.getElementById("save-bar");
  var countEl = document.getElementById("change-count");
  if (bar)     bar.classList.toggle("visible", count > 0);
  if (countEl) countEl.textContent = count;
}

function discardChanges() {
  state.pendingChanges = {};
  updateSaveBar();
  renderXLGrid();
}

async function saveChanges() {
  var btn = document.getElementById("save-btn");
  var allPending = Object.values(state.pendingChanges);
  if (!allPending.length) return;

  // Separate row-based edits (from sheet inputs) from task-key edits
  var taskUpdates = allPending.filter(function(u) { return !u._row; });
  var rowUpdates  = allPending.filter(function(u) { return !!u._row; });

  btn.disabled = true;
  btn.innerHTML = '<span class="spinner" style="width:13px;height:13px;border-width:2px"></span>';
  try {
    var saveRes = null;
    // Send task updates (existing system)
    if (taskUpdates.length) {
      saveRes = await API.saveTasks(state.project.id, taskUpdates);
      taskUpdates.forEach(function(u) {
        var task = state.tasks.find(function(t) { return taskKey(t) === u.task_key; });
        if (task) { task.percent_complete = u.percent_complete; task.remark = u.remark; }
      });
    }
    // Send row-based edits (sheet input cells)
    if (rowUpdates.length) {
      saveRes = await API.saveTasks(state.project.id, rowUpdates);
    }
    state.pendingChanges = {};
    updateSaveBar();
    var total = taskUpdates.length + rowUpdates.length;
    toast("Saved " + total + " change" + (total > 1 ? "s" : "") + " \u2713");

    // Use server-returned timestamp so it stays consistent after renderXLGrid re-fetch.
    // The server stamps last_modified into its cache at save time, so this value
    // will also be returned by the subsequent sheet fetch inside renderXLGrid.
    var serverTs = saveRes && saveRes.last_modified;
    if (serverTs) {
      updateProjectHeaderTimestamp(serverTs);
      // Patch client-side sheet cache too so loadSheetView cache-hit also shows it
      if (state.project && state.sheetCache[state.project.id]) {
        state.sheetCache[state.project.id].last_modified = serverTs;
      }
      // Fire-and-forget: write timestamp to monitor BA column
      API.updateMonitorTimestamp(state.project.id, serverTs).catch(function() {});
      // ADD THIS LINE: Refresh monitor data after save
            refreshMonitorData();
    }

    // Force reload from updated JSON cache
    if (state.project) {
        delete state.sheetCache[state.project.id];
        await loadSheetView();  // This reloads UI from the updated cache
    }
  } catch(err) {
    toast(err.message || "Save failed", "error");
  } finally {
    btn.disabled = false;
    btn.textContent = "Save Changes";
  }
}


async function refreshFileStatusOnOpen() {
  try {
    var result = await API.req("POST", "/refresh-file-status");
    var updatedProjects = result.projects;
    
    if (updatedProjects && window._masterListFull) {
      var updatedStatusMap = {};
      updatedProjects.forEach(function(p) {
        updatedStatusMap[p.project_id] = p.file_exists;
      });
      
      // Update global master list
      window._masterListFull.forEach(function(project) {
        if (updatedStatusMap.hasOwnProperty(project.project_id)) {
          project.file_exists = updatedStatusMap[project.project_id];
        }
      });
      
      if (window._masterList) {
        window._masterList.forEach(function(project) {
          if (updatedStatusMap.hasOwnProperty(project.project_id)) {
            project.file_exists = updatedStatusMap[project.project_id];
          }
        });
      }
      
      // Re-render the sidebar
      var searchTerm = document.getElementById("master-search-input")?.value || "";
      filterMasterList(searchTerm);
    }
  } catch(err) {
    console.error("Failed to refresh file status:", err);
  }
}


// ── Boot ───────────────────────────────────────────────────────
window.addEventListener("DOMContentLoaded", init);

// __Rendering the monitoring grid______________________________

function renderMonitorGrid(container, data) {
    var cells      = data.cells;
    var rowHeights = data.row_heights;
    var maxRow     = data.max_row;
    var cols       = data.cols;

    var isDark = document.documentElement.getAttribute("data-theme") !== "light";

    var borderColor = isDark ? "#3a3a3a" : "#d0d0d0";
    var numBg       = isDark ? "#242424" : "#e8e8e8";
    var headerBg    = isDark ? "#2a2a2a" : "#f2f2f2";
    var evenRowBg   = isDark ? "#242424" : "#f7f9fc";
    var oddRowBg    = isDark ? "#1e1e1e" : "#ffffff";

    // ── Compute total % width so table can expand beyond 100%
    //    when many columns are present (mirrors monitor.js logic)
    var totalPct = 2; // ~36px # col ≈ 2%
    for (var ci = 0; ci < cols.length; ci++) {
        var w = getSheetColumnWidth(cols[ci]);
        if (w.indexOf("%") !== -1) totalPct += parseFloat(w) || 0;
    }
    var tableWidth = totalPct <= 100 ? "100%" : totalPct + "%";

    var tableStyle = [
        "border-collapse:collapse",
        "table-layout:fixed",          // fixed layout — widths come from <colgroup>
        "width:" + tableWidth,         // lets % cols distribute properly
        "font-family:Calibri,Arial,sans-serif",
        "font-size:11px",
        "background:" + oddRowBg,
        "color:" + (isDark ? "#e0e0e0" : "#000000"),
    ].join(";");

html += '<table style="' + tableStyle + '">';

// ── <colgroup> — this is what drives % column widths ─────────
html += '<colgroup>';
html += '<col style="width:36px;">';  // # column
for (var ci0 = 0; ci0 < cols.length; ci0++) {
  var col0 = cols[ci0];  // ✅ ci0 matches
  if (isSkippedCol(col0)) continue;
  if (isCollapsedCol(col0)) continue;
  html += '<col style="width:' + getSheetColumnWidth(col0) + ';">';
}
html += '</colgroup>';

    // ── Header row ───────────────────────────────────────────────
    html += '<thead><tr style="background:' + headerBg + ';">';
    html += '<th style="'
        + 'position:sticky;left:0;z-index:2;'
        + 'background:' + numBg + ';'
        + 'border:1px solid ' + borderColor + ';'
        + 'width:36px;min-width:36px;'
        + 'padding:4px;text-align:center;font-weight:600;">#</th>';

    for (var ci = 0; ci < cols.length; ci++) {
        var col = cols[ci];
        // No inline width needed — <colgroup> drives it in table-layout:fixed
        html += '<th style="'
            + 'border:1px solid ' + borderColor + ';'
            + 'padding:4px 6px;text-align:center;font-weight:600;'
            + 'white-space:normal;word-break:break-word;line-height:1.3;'
            + 'vertical-align:middle;overflow:hidden;">'
            + h(col) + '</th>';
    }
    html += '</tr></thead>';

    // ── Body rows ────────────────────────────────────────────────
    html += '<tbody>';
    for (var r = 1; r <= maxRow; r++) {
        var rh     = rowHeights[String(r)] || 20;
        var rowBg  = (r % 2 === 0) ? evenRowBg : oddRowBg;

        html += '<tr style="height:' + rh + 'px;background:' + rowBg + ';">';
        html += '<td style="'
            + 'position:sticky;left:0;z-index:1;'
            + 'background:' + numBg + ';'
            + 'border:1px solid ' + borderColor + ';'
            + 'text-align:center;font-size:11px;color:' + (isDark ? "#888" : "#999") + ';">'
            + r + '</td>';

        for (var ci = 0; ci < cols.length; ci++) {
            var col   = cols[ci];
            var coord = col + r;
            var info  = cells[coord];
            var value = (info && info.v !== undefined && info.v !== null) ? h(String(info.v)) : "";

            html += '<td style="'
                + 'border:1px solid ' + borderColor + ';'
                + 'padding:2px 4px;'
                + 'overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">'
                + value + '</td>';
        }
        html += '</tr>';
    }
    html += '</tbody></table>';

    container.innerHTML = html;
}