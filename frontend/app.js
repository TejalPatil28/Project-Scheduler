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
  var map = {admin:"Admin",head:"Head",pm:"Project Manager",hw_tl:"HW Team Lead",sw_tl:"SW Team Lead",mfg_tl:"MFG Team Lead"};
  return map[role] || role;
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
        '<div class="login-logo">&#x2B21;</div>',
        '<div class="login-title">Workezz Project Scheduler</div>',
        '<div class="login-sub">Sign in to your workspace</div>',
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
      '<button class="popover-item" onclick="showUserModal(); closeSettingsPopover();">\u2795 Add User</button>',
      '<button class="popover-item" onclick="renderUsers(); closeSettingsPopover();">\u{1F465} Manage Users</button>',
      '<button class="popover-item" onclick="doRebuildIndex(); closeSettingsPopover();">\u{1F504} Rebuild Index</button>',
    ].join("");
  }
  return [
    '<div class="popover hidden" id="user-popover">',
      '<div class="popover-header">',
        '<div class="avatar avatar-md">' + h(u.name[0].toUpperCase()) + '</div>',
        '<div>',
          '<div class="popover-name">' + h(u.name) + '</div>',
          '<div class="popover-role">' + roleLabel(u.role) + '</div>',
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
        '<button class="popover-item" onclick="switchToProjects(); closeSettingsPopover();">\u{1F4C1} Projects</button>',
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
              '<div style="font-size:10px;color:var(--text2)">' + roleLabel(u.role) + '</div>',
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
    var data = await API.req("GET", "/monitor/sheet");
    if (!data || data.error) {
      page.innerHTML = '<div class="empty" style="padding-top:80px"><div class="empty-icon">&#9906;</div><div class="empty-text">No monitoring file found for your account.</div></div>';
      return;
    }
    var wrap = document.createElement("div");
    wrap.id = "xl-grid";
    wrap.style.flex = "1";
    wrap.style.minHeight = "0";
    page.innerHTML = "";
    page.appendChild(wrap);
    renderExcelMirror(wrap, data);
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

  if (!masterList.length) {
    el.innerHTML = '<div class="master-empty">No master file found</div>';
    return;
  }

  // Store master list globally for filtering
  window._masterListFull = masterList;
  window._projectsFull = projects;

  // Build search input + list container
  var html = [
    '<div class="master-search-wrap" style="padding: 8px 10px; border-bottom: 1px solid var(--border);">',
      '<input type="text" id="master-search-input" class="form-input" style="font-size:12px; padding:6px 8px;" placeholder="Search project ID..." autocomplete="off">',
    '</div>',
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

  // Initial render
  filterMasterList("");
}

function filterMasterList(searchTerm) {
  var masterList = window._masterListFull || [];
  var projects = window._projectsFull || [];
  var container = document.getElementById("master-list-items");
  if (!container) return;

  var term = searchTerm.toLowerCase().trim();
  
  var filtered = masterList.filter(function(m) {
    var pid = m.project_id || "";
    var fileId = m.file_id || "";
    return term === "" || pid.toLowerCase().includes(term) || fileId.toLowerCase().includes(term);
  });

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

    var classes = "master-item"
      + (isActive ? " active" : "")
      + (stale    ? " stale"  : "")
      + (exists   ? ""        : " no-file");

    var itemBg = !exists ? "background:rgba(244,63,94,0.28);border-color:rgba(244,63,94,0.5);" : "";

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
  openProject(fileId || pid);
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
    '<div id="project-timestamp" class="project-timestamp" style="font-size:11px;color:var(--text2);font-family:var(--font-mono);"></div>',
  '</div>',
    '<div class="proj-body-layout">',
      '<div class="proj-body-right" style="flex:1;min-width:0">',
        '<div class="xl-wrap" id="xl-wrap">',
          '<div class="xl-grid" id="xl-grid"></div>',
        '</div>',
      '</div>',
    '</div>',
    '<div class="save-bar" id="save-bar">',
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
      return;
    }
    var data = await API.req("GET", "/projects/" + pid + "/sheet");
    state.sheetCache[pid] = data;  // cache in browser
    renderExcelMirror(wrap, data);
  } catch(err) {
    wrap.innerHTML = '<div style="padding:20px;color:var(--text3)">Could not load sheet: ' + h(err.message||"error") + '</div>';
  }
}

function renderExcelMirror(container, data) {
  var cells        = data.cells;
  var colWidths    = data.col_widths;
  var rowHeights   = data.row_heights;
  var maxRow       = data.max_row;
  var cols         = data.cols;
   console.log("All columns in grid:", cols);
  console.log("Does column R exist in cols?", cols.indexOf("R") !== -1);
  var colGroups    = data.col_groups || [];
  var editableFill = data.editable_fill || null;
  var infoRows     = data.info_rows    || [];
  var headerRows   = data.header_rows  || [];
  var banner       = data.project_banner || {};
  var leftPanel    = data.left_panel   || {};
  var infoRowSet   = {};
  infoRows.forEach(function(r) { infoRowSet[r] = true; });
  var headerRowSet = {};
  headerRows.forEach(function(r) { headerRowSet[r] = true; });

  // Add these constants
  var TASK_START_ROW = 9;
  var TASK_END_ROW = 55;

  var AD_OPTIONS = ["Engineering","Purchase","Software","Project Management","Manufacturing","Sales","Client"];

    // ── COLOR CONFIGURATION ─────────────────────────────────
  // Define all custom background colors for columns/ranges
  // Add new rules here as needed
    var colorRules = isDark ? [
      { name: "task columns E-W", columns: ["E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","AA","AB"], rows: "9+", color: "#232a2e" },
      { name: "actual dates X,Y,Z", columns: ["X","Y","Z"], rows: "9+", color: "#1a2a3a" },
    ] : [
      { name: "task columns E-W", columns: ["E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","AA","AB"], rows: "9+", color: "#d9d9d9" },
      { name: "actual dates X,Y,Z", columns: ["X","Y","Z"], rows: "9+", color: "#90b4df" },
    ];
    // AD Column color mapping based on dropdown value
    var AD_COLORS = {
        "engineering":        "#ffb3b3",
        "purchase":           "#5f933c",
        "software":           "#0096cc",
        "project management": "#005fa3",
        "manufacturing":      "#2f491e",
        "sales":              "#00d9d9",
        "client":             "#6d006d"
    };
  
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

  function renderInputCell(info, coord, col, tdStyle) {
    var curVal = _editPending[coord] !== undefined ? _editPending[coord] : (info.v || "");
    var inputTextColor = isDark ? "#d8d8d8" : "#000000";  // contrast-aware
    var inputStyle = "width:100%;height:100%;border:none;outline:none;background:transparent;font-family:Calibri,Arial,sans-serif;font-size:11px;padding:0 2px;box-sizing:border-box;color:" + inputTextColor + ";";
    var input = "";

    if (col === "X" || col === "Y") {
      // Date picker
      var dateVal = curVal ? curVal : "";
      // Convert dd-Mon-yy to yyyy-mm-dd for input[type=date]
      if (dateVal) {
        try {
          var d = new Date(dateVal);
          if (!isNaN(d)) dateVal = d.toISOString().split("T")[0];
        } catch(e) {}
      }
      input = "<input type=\"date\" style=\"" + inputStyle + "\" value=\"" + dateVal + "\" data-coord=\"" + coord + "\" data-col=\"" + col + "\" onchange=\"__xlEditCell(this)\">";
    } else if (col === "Z") {
      // Number 0-100 — value from Excel is e.g. "100%" so strip the %
      var numVal = String(curVal || "0").replace("%", "").trim();
      input = "<input type=\"number\" min=\"0\" max=\"100\" style=\"" + inputStyle + "\" value=\"" + numVal + "\" data-coord=\"" + coord + "\" data-col=\"" + col + "\" onchange=\"__xlEditCell(this)\">";
    } else if (col === "AD") {
      // Dropdown
      var opts = AD_OPTIONS.map(function(o) {
        return "<option value=\"" + o + "\"" + (o === curVal ? " selected" : "") + ">" + o + "</option>";
      }).join("");
      input = "<select style=\"" + inputStyle + "cursor:pointer;\" data-coord=\"" + coord + "\" data-col=\"" + col + "\" onchange=\"__xlEditCell(this);__applyAdColor(this)\" ><option value=\"\"></option>" + opts + "</select>";
    } else if (col === "AF") {
      // Remarks - free text input
      input = "<input type=\"text\" style=\"" + inputStyle + "\" value=\"" + h(curVal || "") + "\" data-coord=\"" + coord + "\" data-col=\"" + col + "\" onchange=\"__xlEditCell(this)\">";
    }
    return input;
  }

  // Track collapse state per group
  var collapseState = colGroups.map(function(g) { return g.collapsed === true; });

  // Fast lookup: col_letter -> groupIndex
  var colToGroup = {};
  colGroups.forEach(function(g, gi) {
    g.cols.forEach(function(c) { colToGroup[c] = gi; });
  });

  // Returns true if this col should be hidden (all cols in group hide when collapsed)
  function isCollapsedCol(col) {
    var gi = colToGroup[col];
    if (gi === undefined || gi < 0) return false;
    if (!collapseState[gi]) return false;
    // When collapsed, hide all cols EXCEPT the last one (which shows as placeholder)
    var lastCol = colGroups[gi].cols[colGroups[gi].cols.length - 1];
    return col !== lastCol;
  }

  function borderStyle(weight) {
    var w = weight === "medium" ? "2px" : weight === "thick" ? "3px" : "1px";
    var c = weight === "medium" ? "#8c8c8c" : weight === "thick" ? "#595959" : "#BFBFBF";
    return w + " solid " + c;
  }

  // Detect dark mode
  var isDark = document.documentElement.getAttribute("data-theme") !== "light";
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
  if (ldHasData) {
    bannerHtml +=     '<div style="display:flex;flex-direction:column;justify-content:center;padding:0 20px;flex-shrink:0;gap:6px;margin-left:auto;border-left:1px solid ' + bDivider + ';">';
    if (banner.ld_date_val)    bannerHtml += '<div style="display:flex;align-items:center;gap:8px;"><div style="' + bFont + 'font-size:9px;font-weight:700;text-transform:uppercase;color:' + bLabel + ';white-space:nowrap;">' + (banner.ld_date_lbl||"LD Date") + '</div><div style="' + bFont + 'font-size:13px;font-weight:700;color:' + bValue + ';white-space:nowrap;">' + h(banner.ld_date_val) + '</div></div>';
    if (banner.ld_maxwk_val)   bannerHtml += '<div style="display:flex;align-items:center;gap:8px;"><div style="' + bFont + 'font-size:9px;font-weight:700;text-transform:uppercase;color:' + bLabel + ';white-space:nowrap;">' + (banner.ld_maxwk_lbl||"LD Max/Wk") + '</div><div style="' + bFont + 'font-size:13px;font-weight:700;color:' + bValue + ';white-space:nowrap;">' + h(banner.ld_maxwk_val) + '</div></div>';
    if (banner.ld_maxov_val)   bannerHtml += '<div style="display:flex;align-items:center;gap:8px;"><div style="' + bFont + 'font-size:9px;font-weight:700;text-transform:uppercase;color:' + bLabel + ';white-space:nowrap;">' + (banner.ld_maxov_lbl||"LD Max/OV") + '</div><div style="' + bFont + 'font-size:13px;font-weight:700;color:' + bValue + ';white-space:nowrap;">' + h(banner.ld_maxov_val) + '</div></div>';
    if (banner.ld_remarks_val) bannerHtml += '<div style="display:flex;align-items:center;gap:8px;"><div style="' + bFont + 'font-size:9px;font-weight:700;text-transform:uppercase;color:' + bLabel + ';white-space:nowrap;">' + (banner.ld_remarks_lbl||"LD Remarks") + '</div><div style="' + bFont + 'font-size:13px;font-weight:700;color:' + bValue + ';white-space:nowrap;">' + h(banner.ld_remarks_val) + '</div></div>';
    bannerHtml +=     '</div>';
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
    var tableStyle = [
      "border-collapse:collapse",
      "table-layout:fixed",
      "font-family:Calibri,Arial,sans-serif",
      "font-size:11px",
      "background:" + xlBg,
      "color:" + xlText,
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
      var pmFg  = item.pm_font  || lpValue;
      var swFg  = item.swe_font || lpValue;
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
    leftPanelHtml += '</div>'; // end inner
    leftPanelHtml += '</div>'; // end panel

    var html = '<div style="display:flex;flex-direction:column;height:calc(115vh - var(--topbar-h) - 36px);position:relative;background:' + xlBg + ';overflow:hidden;">';
    html += '<div style="flex-shrink:0;overflow:hidden;">' + bannerHtml + '</div>';
    html += '<div style="display:flex;flex:1;min-height:0;overflow:hidden;">';
    html += leftPanelHtml;
    html += '<div style="overflow:auto;flex:1;position:relative;">';
    html += '<table style="' + tableStyle + '">';
    html += '<thead>';

          if (colGroups.length > 0) {
      // ── GROUP BAR ROW (buttons with borders) ──
      html += '<tr style="height:20px;">';

      var ci = 0;
      while (ci < cols.length) {
        var col = cols[ci];
        var gi  = colToGroup[col];

        if (gi !== undefined && gi >= 0 && cols[ci] === colGroups[gi].cols[0]) {
          var group     = colGroups[gi];
          var collapsed = collapseState[gi];
          var btnLabel  = collapsed ? "+" : "–";
          var btnStyle  = "cursor:pointer;font-size:11px;font-weight:700;padding:0px 6px;border-radius:3px;border:1px solid " + (isDark ? "#555" : "#aaa") + ";background:" + (isDark ? "#2a2a2a" : "#e8e8e8") + ";color:" + (isDark ? "#ddd" : "#333") + ";line-height:1.4;";
          
          if (!collapsed) {
            // Expanded: show button above the first column of the group
            var firstW = colWidths[group.cols[0]] || 64;
            html += '<th style="position:sticky;top:0;z-index:4;background:transparent;border:none;text-align:left;padding:0 0 0 6px;width:' + firstW + 'px;">';
            html += '<button style="' + btnStyle + '" onclick="__xlToggleGroup(' + gi + ')">' + btnLabel + '</button>';
            html += '</th>';
            // Empty cells for the remaining columns in the group
            for (var gci = 1; gci < group.cols.length; gci++) {
              var gcw = colWidths[group.cols[gci]] || 64;
              html += '<th style="position:sticky;top:0;z-index:4;background:transparent;border:none;width:' + gcw + 'px;"></th>';
            }
          } else {
            // Collapsed: show button above the last column of the group (the only visible one)
            var lastCol = group.cols[group.cols.length - 1];
            var lastW   = colWidths[lastCol] || 64;
            html += '<th style="position:sticky;top:0;z-index:4;background:transparent;border:none;text-align:left;padding:0 0 0 6px;width:' + lastW + 'px;">';
            html += '<button style="' + btnStyle + '" onclick="__xlToggleGroup(' + gi + ')">' + btnLabel + '</button>';
            html += '</th>';
          }
          ci += group.cols.length;
        } else {
          var cw = colWidths[col] || 64;
          html += '<th style="position:sticky;top:0;z-index:4;background:transparent;border:none;width:' + cw + 'px;"></th>';
          ci++;
        }
      }
      html += '<\/tr>';
    }

    html += '</thead>';

    // ── TBODY ──
    html += '</thead>';

    // ── CUSTOM HEADER ROW (replaces hidden Excel rows 6-8) ───────
    // Defines semantic column headers for the task grid
    var customHeaderDefs = {
      // col letter -> label, and optional bg/text overrides
      "E":  { label: "Buffer Time",         bg: "#fcd5b4", fg: "#000000" },
      "F":  { label: "PH",           bg: "#fcd5b4", fg: "#000000" },
      "G":  { label: "TFO",       bg: "#fcd5b4", fg: "#000000" },
      "H":  { label: "Phase ID",     bg: "#fcd5b4", fg: "#000000" },
      "I":  { label: "Task Description",     bg: "#fcd5b4", fg: "#000000" },
      "J":  { label: "P1 Start Date",         bg: "#fcd5b4", fg: "#000000" },
      "K":  { label: "P1 End Date",          bg: "#fcd5b4", fg: "#000000" },
      "L":  { label: "P2 Start Date",   bg: "#fcd5b4", fg: "#000000" },
      "M":  { label: "P2 End Date",           bg: "#fcd5b4", fg: "#000000" },
      "N":  { label: "P3 Start Date",     bg: "#fcd5b4", fg: "#000000" },
      "O":  { label: "P3 End Date",       bg: "#fcd5b4", fg: "#000000" },
      "P":  { label: "Org Plan Date",       bg: "#fcd5b4", fg: "#000000" },
      "Q":  { label: "Org End Date",        bg: "#fcd5b4", fg: "#000000" },
      "R":  { label: "Ref. Lead Time",       bg: "#fcd5b4", fg: "#000000" },
      "S":  { label: "Lead Time",       bg: "#fcd5b4", fg: "#000000" },
      "T":  { label: "Intlk",         bg: "#fcd5b4", fg: "#000000" },
      "U":  { label: "Effort Days",         bg: "#fcd5b4", fg: "#000000" },
      "V":  { label: "Cur. Start Date",        bg: "#fcd5b4", fg: "#000000" },
      "W":  { label: "Cur. End Date",     bg: "#fcd5b4", fg: "#000000" },
      "X":  { label: "Act. Start Date",    bg: "#fcd5b4", fg: "#000000" },
      "Y":  { label: "Act. End Date",      bg: "#fcd5b4", fg: "#000000" },
      "Z":  { label: "% Complete",        bg: "#fcd5b4", fg: "#000000" },
      "AA": { label: "Exptd % Completion",      bg: "#fcd5b4", fg: "#000000" },
      "AB": { label: "Alert Date for 80%",        bg: "#fcd5b4", fg: "#000000" },
      "AD": { label: "Help Req. from",     bg: "#fcd5b4", fg: "#000000" },
      "AF": { label: "Remark",        bg: "#fcd5b4", fg: "#000000" },
    };
    // Light mode overrides
    var customHeaderDefsLight = {
      "E":  { label: "Buffer Time",         bg: "#fcd5b4", fg: "#000000" },
      "F":  { label: "PH",           bg: "#fcd5b4", fg: "#000000" },
      "G":  { label: "TFO",       bg: "#fcd5b4", fg: "#000000" },
      "H":  { label: "Phase ID",     bg: "#fcd5b4", fg: "#000000" },
      "I":  { label: "Task Description",     bg: "#fcd5b4", fg: "#000000" },
      "J":  { label: "P1 Start Date",         bg: "#fcd5b4", fg: "#000000" },
      "K":  { label: "P1 End Date",          bg: "#fcd5b4", fg: "#000000" },
      "L":  { label: "P2 Start Date",   bg: "#fcd5b4", fg: "#000000" },
      "M":  { label: "P2 End Date",           bg: "#fcd5b4", fg: "#000000" },
      "N":  { label: "P3 Start Date",     bg: "#fcd5b4", fg: "#000000" },
      "O":  { label: "P3 End Date",       bg: "#fcd5b4", fg: "#000000" },
      "P":  { label: "Org Plan Date",       bg: "#fcd5b4", fg: "#000000" },
      "Q":  { label: "Org End Date",        bg: "#fcd5b4", fg: "#000000" },
      "R":  { label: "Ref. Lead Time",       bg: "#fcd5b4", fg: "#000000" },
      "S":  { label: "Lead Time",       bg: "#fcd5b4", fg: "#000000" },
      "T":  { label: "Intlk",         bg: "#fcd5b4", fg: "#000000" },
      "U":  { label: "Effort Days",         bg: "#fcd5b4", fg: "#000000" },
      "V":  { label: "Cur. Start Date",        bg: "#fcd5b4", fg: "#000000" },
      "W":  { label: "Cur. End Date",     bg: "#fcd5b4", fg: "#000000" },
      "X":  { label: "Act. Start Date",    bg: "#fcd5b4", fg: "#000000" },
      "Y":  { label: "Act. End Date",      bg: "#fcd5b4", fg: "#000000" },
      "Z":  { label: "% Complete",        bg: "#fcd5b4", fg: "#000000" },
      "AA": { label: "Exptd % Completion",      bg: "#fcd5b4", fg: "#000000" },
      "AB": { label: "Alert Date for 80%",        bg: "#fcd5b4", fg: "#000000" },
      "AD": { label: "Help Req. from",     bg: "#fcd5b4", fg: "#000000" },
      "AF": { label: "Remark",        bg: "#fcd5b4", fg: "#000000" },
    };
    var chDefs = isDark ? customHeaderDefs : customHeaderDefsLight;
    var chRowBg    = isDark ? "#141414" : "#e8edf5";
    var chBorder   = isDark ? "#333333" : "#c0c7d8";
    var chFallbackBg = isDark ? "#1a1a2e" : "#e8edf5";
    var chFallbackFg = isDark ? "#888888" : "#555555";

    html += '<thead id="xl-custom-header">';
    html += '<tr style="height:26px;">';
    for (var chi = 0; chi < cols.length; chi++) {
      var chCol = cols[chi];
      if (isCollapsedCol(chCol)) continue;
      var chCw = colWidths[chCol] || 64;
      var chDef = chDefs[chCol] || {};
      var chBg = chDef.bg || chFallbackBg;
      var chFg = chDef.fg || chFallbackFg;
      var chLabel = chDef.label || chCol;
      var chStyle = [
        "width:" + chCw + "px",
        "min-width:" + chCw + "px",
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
        var coord = col3 + r;
        var info  = cells[coord];

            // DEBUG: Check column R
      if (col3 === "R") {
          console.log("===== COLUMN R DEBUG =====");
          console.log("Row:", r);
          console.log("coord:", coord);
          console.log("info object:", info);
          console.log("info.v:", info ? info.v : "info is null/undefined");
          console.log("info.skip:", info ? info.skip : "N/A");
          console.log("=========================");
      }

        // skip all cols in a collapsed group
        if (isCollapsedCol(col3)) continue;

        // merged slave — skip
        if (info && info.skip) continue;

        var cw3 = colWidths[col3] || 64;
        var tdStyle = [
          "width:" + cw3 + "px",
          "min-width:" + cw3 + "px",
          "overflow:hidden",
          "padding:1px 3px",
          "border:1px solid " + xlCellBorder,
          "vertical-align:bottom",
          "white-space:nowrap",
        ];
        var content = "";

                // ── BACKGROUND COLOR LOGIC ──────────────────────────────
        // Define AD column colors
        var AD_COLORS = {
            "engineering":        "#ffb3b3",
            "purchase":           "#5f933c",
            "software":           "#0096cc",
            "project management": "#005fa3",
            "manufacturing":      "#2f491e",
            "sales":              "#00d9d9",
            "client":             "#6d006d"
        };
        
        // Check if AD column
        var isADColumn = (col3 === "AD" && r >= 9);
        
        // Check if Z column is 100% complete
        var isZComplete = false;
        if (col3 === "Z" && r >= 9) {
            var zCoord = "Z" + r;
            var zInfo = cells[zCoord];
            var zVal = zInfo ? String(zInfo.v || "") : "";
            isZComplete = (zVal === "100%" || zVal === "100");
        }
        
        // Get custom background color from rules (for non-AD columns)
        var customBg = getCustomBackgroundColor(col3, r);
        
        // Determine which background color to apply
        var bgColor = null;
        
        if (isADColumn) {
            // AD column: use color based on dropdown value
            var adCoord = "AD" + r;
            var adInfo = cells[adCoord];
            var adValue = adInfo ? String(adInfo.v || "").toLowerCase().trim() : "";
            var adColor = AD_COLORS[adValue];
            if (adColor) {
                bgColor = adColor;
                // All AD colors have dark backgrounds — always white text
                tdStyle.push("color:#ffffff");
            }
        } else if (isZComplete) {
            // Z column at 100%: green
            bgColor = "#92d050";
        } else if (customBg) {
            // Custom color from rules (E-W, X,Y,Z)
            bgColor = customBg;
        } else {
            // Default row background
            bgColor = rowBg;
        }
        
        // Apply the background color if we have one
        if (bgColor) {
            tdStyle.push("background-color:" + bgColor);
        }
        
        var content = "";

        if (info) {
          if (info.font) {
            var f = info.font;
            if (f.bold)      tdStyle.push("font-weight:bold");
            if (f.italic)    tdStyle.push("font-style:italic");
            if (f.underline) tdStyle.push("text-decoration:underline");
            if (f.size)      tdStyle.push("font-size:" + f.size + "px");
            if (f.color)     tdStyle.push("color:" + f.color);
          }
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
              // For AD column, determine editability based on Z value (ignore backend editable flag)
    var isADColumnForEdit = (col3 === "AD" && r >= 9);
    var canEditAD = true;

    if (isADColumnForEdit && r >= 9) {
        var zCoord = "Z" + r;
        var zInfo = cells[zCoord];
        var zVal = zInfo ? String(zInfo.v || "") : "";
        var zIsComplete = (zVal === "100%" || zVal === "100");
        canEditAD = !zIsComplete;  // Only editable if NOT 100%
        
        // DEBUG: Log AD column info
        console.log("========== AD COLUMN DEBUG ==========");
        console.log("Row:", r);
        console.log("Z Value from cells:", zVal);
        console.log("Z Is Complete (100%):", zIsComplete);
        console.log("Can Edit AD:", canEditAD);
        console.log("=====================================");
    }

    // Use frontend-calculated value for AD column, otherwise use backend editable flag
    var isEditable = isADColumnForEdit ? canEditAD : (info && info.editable);

    // DEBUG: Log the final decision
    if (isADColumnForEdit) {
        console.log("FINAL - AD Column Row", r, "isEditable:", isEditable);
    }

    if (isEditable) {
        console.log("Rendering INPUT for:", coord, "col:", col3);
        tdStyle.push("padding:0");
        content = renderInputCell(info, coord, col3, tdStyle);
    } else if (info && info.v !== null && info.v !== undefined) {
        console.log("Rendering TEXT for:", coord, "value:", info.v);
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

    // Map col to task field
    var fieldMap = { "X": "actual_start", "Y": "actual_end", "Z": "percent_complete", "AD": "help_required", "AF": "remark" };
    var field = fieldMap[col];
    if (!field) return;

    // Find existing pending change for this row or create one
    // We key task changes by row number
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
    console.log("[XL] toggle clicked, gi=", gi, "current state=", collapseState[gi]);
    collapseState[gi] = !collapseState[gi];
    console.log("[XL] new state=", collapseState[gi]);
    var c = document.getElementById(containerId);
    console.log("[XL] container found?", !!c, "id=", containerId);
    if (c) { container = c; buildTable(); console.log("[XL] buildTable done"); }
  };

  buildTable();
  console.log("[XL] initial buildTable done. colGroups=", colGroups.length, "containerId=", containerId);
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
    // Send task updates (existing system)
    if (taskUpdates.length) {
      await API.saveTasks(state.project.id, taskUpdates);
      taskUpdates.forEach(function(u) {
        var task = state.tasks.find(function(t) { return taskKey(t) === u.task_key; });
        if (task) { task.percent_complete = u.percent_complete; task.remark = u.remark; }
      });
    }
    // Send row-based edits (sheet input cells)
    if (rowUpdates.length) {
      await API.saveTasks(state.project.id, rowUpdates);
    }
    state.pendingChanges = {};
    updateSaveBar();
    var total = taskUpdates.length + rowUpdates.length;
    toast("Saved " + total + " change" + (total > 1 ? "s" : "") + " \u2713");
    // Invalidate client cache so re-render fetches fresh data
    if (state.project) delete state.sheetCache[state.project.id];
    renderXLGrid();
  } catch(err) {
    toast(err.message || "Save failed", "error");
  } finally {
    btn.disabled = false;
    btn.textContent = "Save Changes";
  }
}

// ── USERS ─────────────────────────────────────────────────────
async function renderUsers() {
  setNav("users");
  setTopbar("User Management", false);
  var page = document.getElementById("page-content");
  page.innerHTML = '<div class="loading"><span class="spinner"></span> Loading users...</div>';

  var users = [];
  try { users = await API.getUsers(); }
  catch(err) { page.innerHTML = '<div class="alert alert-error">' + h(err.message) + '</div>'; return; }

  var ROLE_BADGE = {admin:"badge-red",head:"badge-purple",pm:"badge-blue",hw_tl:"badge-amber",sw_tl:"badge-green",mfg_tl:"badge-gray"};

  var rows = users.map(function(u) {
    var uJson = h(JSON.stringify(u));
    return '<div style="background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);padding:14px 16px;display:flex;align-items:center;gap:14px;animation:fadeUp 0.3s ease">'
      + '<div class="avatar avatar-md">' + h(u.name[0].toUpperCase()) + '</div>'
      + '<div style="flex:1;min-width:0">'
      + '<div style="font-weight:600;font-size:14px;margin-bottom:2px">' + h(u.name) + '</div>'
      + '<div style="font-size:11px;color:var(--text2);font-family:var(--font-mono)">@' + h(u.username) + ' &middot; ' + h(u.short_name) + '</div>'
      + '</div>'
      + '<span class="badge ' + (ROLE_BADGE[u.role]||"badge-gray") + '">' + roleLabel(u.role) + '</span>'
      + '<div style="display:flex;gap:6px;flex-shrink:0">'
      + '<button class="btn btn-ghost btn-sm" onclick=\'showUserModal(' + uJson + ')\'>Edit</button>'
      + '<button class="btn btn-danger btn-sm" onclick="doDeleteUser(\'' + h(u.username) + '\',\'' + h(u.name) + '\')">Del</button>'
      + '</div>'
      + '</div>';
  }).join("");

  page.innerHTML = '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px">'
    + '<div class="dash-title">Users <span style="color:var(--text3);font-size:14px;font-weight:400">(' + users.length + ')</span></div>'
    + '<button class="btn btn-primary btn-sm" onclick="showUserModal()">+ New User</button>'
    + '</div>'
    + '<div style="display:flex;flex-direction:column;gap:8px">' + rows + '</div>';
}

function showUserModal(user) {
  var isEdit = !!user;
  var u = user || {};
  var ROLES = ["admin","head","pm","hw_tl","sw_tl","mfg_tl"];
  var existing = document.getElementById("user-modal-overlay");
  if (existing) existing.remove();

  var roleOptions = ROLES.map(function(r) {
    return '<option value="' + r + '"' + (r===u.role?" selected":"") + '>' + roleLabel(r) + '</option>';
  }).join("");

  var overlay = document.createElement("div");
  overlay.id = "user-modal-overlay";
  overlay.className = "modal-overlay";
  overlay.innerHTML = '<div class="modal">'
    + '<div class="modal-header">'
    + '<div class="modal-title">' + (isEdit ? "Edit User" : "New User") + '</div>'
    + '<button class="btn btn-ghost btn-sm" onclick="document.getElementById(\'user-modal-overlay\').remove()">&times;</button>'
    + '</div>'
    + '<div class="modal-body">'
    + '<div id="um-err" class="alert alert-error hidden"></div>'
    + '<div class="form-group"><label class="form-label">Full Name</label>'
    + '<input class="form-input" id="um-name" type="text" value="' + h(u.name||"") + '" placeholder="Full Name"></div>'
    + '<div class="form-group"><label class="form-label">Short Name / Initials <span style="color:var(--text3)">(must match Excel)</span></label>'
    + '<input class="form-input" id="um-short" type="text" value="' + h(u.short_name||"") + '" placeholder="KDM" style="text-transform:uppercase;font-family:var(--font-mono)"></div>'
    + '<div class="form-group"><label class="form-label">Username</label>'
    + '<input class="form-input" id="um-username" type="text" value="' + h(u.username||"") + '" placeholder="username"' + (isEdit?" readonly style='opacity:0.5'":"") + '></div>'
    + '<div class="form-group"><label class="form-label">Password ' + (isEdit?"(blank = keep current)":"") + '</label>'
    + '<input class="form-input" id="um-pass" type="password" placeholder="&bull;&bull;&bull;&bull;&bull;&bull;&bull;&bull;"></div>'
    + '<div class="form-group"><label class="form-label">Role</label>'
    + '<select class="form-select" id="um-role">' + roleOptions + '</select></div>'
    + '</div>'
    + '<div class="modal-footer">'
    + '<button class="btn btn-secondary" onclick="document.getElementById(\'user-modal-overlay\').remove()">Cancel</button>'
    + '<button class="btn btn-primary" id="um-save-btn" onclick="doSaveUser(\'' + h(u.username||"") + '\')">'
    + (isEdit ? "Save Changes" : "Create User") + '</button>'
    + '</div></div>';

  overlay.addEventListener("click", function(e) { if (e.target === overlay) overlay.remove(); });
  document.body.appendChild(overlay);
}

async function doSaveUser(existingUsername) {
  var isEdit = !!existingUsername;
  var btn    = document.getElementById("um-save-btn");
  var errEl  = document.getElementById("um-err");
  var name     = document.getElementById("um-name").value.trim();
  var short    = document.getElementById("um-short").value.trim().toUpperCase();
  var username = isEdit ? existingUsername : document.getElementById("um-username").value.trim();
  var password = document.getElementById("um-pass").value;
  var role     = document.getElementById("um-role").value;
  errEl.classList.add("hidden");
  if (!name || !short || !username || (!isEdit && !password)) {
    errEl.textContent = "Please fill all required fields";
    errEl.classList.remove("hidden");
    return;
  }
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner" style="width:13px;height:13px;border-width:2px"></span>';
  try {
    if (isEdit) {
      var updates = {name:name, short_name:short, role:role};
      if (password) updates.password = password;
      await API.updateUser(existingUsername, updates);
      toast("User updated \u2713");
    } else {
      await API.createUser({name:name, short_name:short, username:username, password:password, role:role});
      toast("User created \u2713");
    }
    document.getElementById("user-modal-overlay").remove();
    var nav = document.getElementById("nav-users");
    if (nav && nav.classList.contains("active")) renderUsers();
  } catch(err) {
    errEl.textContent = err.message || "Failed";
    errEl.classList.remove("hidden");
    btn.disabled = false;
    btn.textContent = isEdit ? "Save Changes" : "Create User";
  }
}

async function doDeleteUser(username, name) {
  if (!confirm('Delete "' + name + '"? This cannot be undone.')) return;
  try { await API.deleteUser(username); toast("User deleted"); renderUsers(); }
  catch(err) { toast(err.message, "error"); }
}

// ── Boot ───────────────────────────────────────────────────────
window.addEventListener("DOMContentLoaded", init);

