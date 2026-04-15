// =============================================================
//  users.js  —  User Management Page
//
//  Single source of truth for all role definitions.
//  All tabs, badges, labels, dropdowns are derived from ROLES.
//
//  To add a new role:       add one entry to ROLES
//  To add a new department: add one entry to ROLES + one to TABS
//  Nothing else changes.
//
//  Called via switchToUsers() in app.js.
// =============================================================

// ── Department definitions ─────────────────────────────────────
// Single source of truth for department colors.
// Badge colors, avatar colors, tab accents all derive from here.
// To add a department: add one entry here + roles in ROLES + tab in TABS.
var DEPARTMENTS = {
  PM:  { label: "PM",  color: "#0d9488", light: "#f0fdfa", text: "#0f766e" },
  SW:  { label: "SW",  color: "#7c3aed", light: "#f5f3ff", text: "#6d28d9" },
  HW:  { label: "HW",  color: "#2563eb", light: "#eff6ff", text: "#1d4ed8" },
  MFG: { label: "MFG", color: "#16a34a", light: "#f0fdf4", text: "#15803d" },
};

// ── Role definitions ───────────────────────────────────────────
// dept: null  → admin-only, visible only in "All" tab
// rank: display order within tab (lower = first)
var ROLES = [
  { id: "admin",    label: "Admin",           dept: null,  rank: 0 },
  { id: "pm_head",  label: "PM Head",         dept: "PM",  rank: 1 },
  { id: "pm",       label: "Project Manager", dept: "PM",  rank: 2 },
  { id: "sw_head",  label: "SW Head",         dept: "SW",  rank: 1 },
  { id: "sw_tl",    label: "SW Team Lead",    dept: "SW",  rank: 2 },
  { id: "hw_head",  label: "HW Head",         dept: "HW",  rank: 1 },
  { id: "hw_tl",    label: "HW Team Lead",    dept: "HW",  rank: 2 },
  { id: "mfg_head", label: "MFG Head",        dept: "MFG", rank: 1 },
  { id: "mfg_tl",   label: "MFG Team Lead",   dept: "MFG", rank: 2 },
];

// ── Tab definitions ────────────────────────────────────────────
// depts: null  → show all users including admin
// To add a new department tab: one entry here + entries in ROLES above
var TABS = [
  { id: "all",  label: "All",  icon: "⊞",  depts: null    },
  { id: "pm",   label: "PM",   icon: "📋", depts: ["PM"]  },
  { id: "sw",   label: "SW",   icon: "💻", depts: ["SW"]  },
  { id: "hw",   label: "HW",   icon: "🔧", depts: ["HW"]  },
  { id: "mfg",  label: "MFG",  icon: "🏭", depts: ["MFG"] },
];

// ── Derived lookups — auto-built from ROLES ────────────────────
// DO NOT edit manually. Edit ROLES above instead.
var VALID_ROLES = ROLES.map(function(r) { return r.id; });
var ROLE_LABELS = {};
var ROLE_DEPT   = {};
ROLES.forEach(function(r) {
  ROLE_LABELS[r.id] = r.label;
  ROLE_DEPT[r.id]   = r.dept;
});

// ── Active page state ──────────────────────────────────────────
var _usersState = {
  activeTab:   "all",
  searchQuery: "",
  allUsers:    [],
};

// ── Public helper (used by app.js / schedule_config.js) ────────
function userRoleLabel(role) {
  return ROLE_LABELS[role] || role;
}

// ── Get department config for a role ──────────────────────────
function _deptCfg(role) {
  var dk = ROLE_DEPT[role];
  return dk ? (DEPARTMENTS[dk] || null) : null;
}

// ── Filter helpers ─────────────────────────────────────────────
function _filterForTab(users, tabId) {
  var tab = TABS.find(function(t) { return t.id === tabId; });
  if (!tab || tab.depts === null) return users;
  return users.filter(function(u) {
    return tab.depts.indexOf(ROLE_DEPT[u.role]) !== -1;
  });
}

function _filterBySearch(users, q) {
  if (!q) return users;
  var lq = q.toLowerCase();
  return users.filter(function(u) {
    return (u.name       || "").toLowerCase().indexOf(lq) !== -1 ||
           (u.username   || "").toLowerCase().indexOf(lq) !== -1 ||
           (u.short_name || "").toLowerCase().indexOf(lq) !== -1 ||
           userRoleLabel(u.role).toLowerCase().indexOf(lq) !== -1;
  });
}

function _sortUsers(users) {
  return users.slice().sort(function(a, b) {
    var ra = ROLES.find(function(r) { return r.id === a.role; });
    var rb = ROLES.find(function(r) { return r.id === b.role; });
    var rd = ((ra && ra.rank) || 99) - ((rb && rb.rank) || 99);
    return rd !== 0 ? rd : (a.name || "").localeCompare(b.name || "");
  });
}

function _tabCount(users, tab) {
  if (tab.depts === null) return users.length;
  return users.filter(function(u) {
    return tab.depts.indexOf(ROLE_DEPT[u.role]) !== -1;
  }).length;
}

function _initials(name) {
  return (name || "?").split(" ").filter(Boolean)
    .map(function(w) { return w[0]; }).slice(0, 2).join("").toUpperCase();
}

// ── Render one user card ───────────────────────────────────────
function _renderUserCard(u) {
  var uJson = h(JSON.stringify(u));
  var dept  = _deptCfg(u.role);

  var avBg    = dept ? dept.light : "#fef2f2";
  var avFg    = dept ? dept.text  : "#b91c1c";
  var avBdr   = dept ? dept.color : "#ef4444";
  var barClr  = dept ? dept.color : "#ef4444";
  var bdgBg   = dept ? dept.light : "#fef2f2";
  var bdgFg   = dept ? dept.text  : "#b91c1c";
  var bdgBdr  = dept ? dept.color : "#ef4444";

  return [
    '<div class="uc">',

      // Left color bar (department indicator)
      '<div class="uc-bar" style="background:' + barClr + '"></div>',

      // Avatar
      '<div class="uc-av" style="background:' + avBg + ';color:' + avFg + ';border:1.5px solid ' + avBdr + '33">',
        _initials(u.name),
      '</div>',

      // Info
      '<div class="uc-info">',
        '<div class="uc-name">' + h(u.name) + '</div>',
        '<div class="uc-meta">',
          '<span class="uc-uname">@' + h(u.username) + '</span>',
          '<span class="uc-dot">·</span>',
          '<span class="uc-short">' + h(u.short_name) + '</span>',
        '</div>',
      '</div>',

      // Role badge — per-department color
      '<span class="uc-badge" style="background:' + bdgBg + ';color:' + bdgFg + ';border:1px solid ' + bdgBdr + '44">',
        userRoleLabel(u.role),
      '</span>',

      // Actions — appear on card hover
      '<div class="uc-actions">',
        '<button class="uc-btn uc-edit" onclick=\'showUserModal(' + uJson + ')\'>',
          '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>',
          'Edit',
        '</button>',
        '<button class="uc-btn uc-del" onclick="doDeleteUser(\'' + h(u.username) + '\',\'' + h(u.name) + '\')">',
          '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6M14 11v6"/></svg>',
          'Delete',
        '</button>',
      '</div>',

    '</div>',
  ].join("");
}

// ── Render tabs + stats + cards ────────────────────────────────
function _renderUserList() {
  var users     = _usersState.allUsers;
  var activeTab = _usersState.activeTab;
  var query     = _usersState.searchQuery;

  // Tab bar
  var tabsHtml = TABS.map(function(tab) {
    var count   = _tabCount(users, tab);
    var active  = tab.id === activeTab;
    var dk      = tab.depts ? tab.depts[0] : null;
    var dc      = dk ? DEPARTMENTS[dk] : null;
    var aColor  = dc ? dc.color : "var(--accent)";

    return [
      '<button class="ut' + (active ? " ut-on" : "") + '"',
        active && dc ? ' style="border-bottom-color:' + aColor + ';color:' + aColor + '"' : "",
        ' onclick="_setUsersTab(\'' + tab.id + '\')">',
        '<span class="ut-ico">' + tab.icon + '</span>',
        tab.label,
        '<span class="ut-ct' + (active ? " ut-ct-on" : "") + '"',
          active && dc ? ' style="background:' + aColor + '"' : "",
        '>',
          count,
        '</span>',
      '</button>',
    ].join("");
  }).join("");

  // Filter + sort
  var filtered = _filterForTab(users, activeTab);
  filtered     = _filterBySearch(filtered, query);
  filtered     = _sortUsers(filtered);

  // Cards or empty state
  var listHtml = filtered.length === 0
    ? [
        '<div class="u-empty">',
          '<svg width="44" height="44" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1" opacity="0.2"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>',
          '<div class="u-empty-title">No users found</div>',
          '<div class="u-empty-sub">' + (query ? 'No results for "' + h(query) + '"' : 'Add a user to get started') + '</div>',
          query ? '<button class="u-clear" onclick="document.getElementById(\'u-srch\').value=\'\';_usersState.searchQuery=\'\';_renderUserList()">Clear search</button>' : '',
        '</div>',
      ].join("")
    : filtered.map(_renderUserCard).join("");

  var el = document.getElementById("u-list-area");
  if (!el) return;

  el.innerHTML = [
    '<div class="u-tabs">' + tabsHtml + '</div>',
    '<div class="u-stats">',
      '<span>' + filtered.length + ' of ' + users.length + ' users</span>',
    '</div>',
    '<div class="u-cards">' + listHtml + '</div>',
  ].join("");
}

// ── Set active tab ─────────────────────────────────────────────
window._setUsersTab = function(tabId) {
  _usersState.activeTab = tabId;
  _renderUserList();
};

// ── Main render ────────────────────────────────────────────────
async function renderUsers() {
  var pg = document.getElementById("page-content");
  if (pg) pg.classList.remove("no-pad");

  var sidebar = document.getElementById("master-sidebar");
  if (sidebar) sidebar.style.display = "none";

  var page = document.getElementById("page-content");
  page.innerHTML = '<div class="loading"><span class="spinner"></span> Loading users...</div>';

  var users = [];
  try {
    users = await API.getUsers();
  } catch (err) {
    page.innerHTML = '<div class="alert alert-error">' + h(err.message) + '</div>';
    return;
  }

  _usersState.allUsers    = users;
  _usersState.activeTab   = "all";
  _usersState.searchQuery = "";

  // Inject scoped styles once
  var old = document.getElementById("u-styles");
  if (old) old.remove();
  var sEl = document.createElement("style");
  sEl.id  = "u-styles";
  sEl.textContent = _usersCSS();
  document.head.appendChild(sEl);

  page.innerHTML = [
    '<div class="u-page">',

      // Header
      '<div class="u-hdr">',
        '<div>',
          '<h1 class="u-title">User Management</h1>',
          '<p class="u-sub">Manage team members and their access roles</p>',
        '</div>',
        '<button class="u-new" onclick="showUserModal()">',
          '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>',
          'New User',
        '</button>',
      '</div>',

      // Search
      '<div class="u-srch-wrap">',
        '<svg class="u-srch-ico" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>',
        '<input class="u-srch-inp" id="u-srch" type="text" placeholder="Search by name, username or role..." autocomplete="off">',
      '</div>',

      // Dynamic area
      '<div id="u-list-area"></div>',

    '</div>',
  ].join("");

  document.getElementById("u-srch").addEventListener("input", function(e) {
    _usersState.searchQuery = e.target.value;
    _renderUserList();
  });

  _renderUserList();
}

// ── CSS ────────────────────────────────────────────────────────
function _usersCSS() {
  return `
    .u-page {
      padding: 32px 36px;
      max-width: 880px;
      margin: 0 auto;
    }

    /* Header */
    .u-hdr {
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      margin-bottom: 24px;
    }
    .u-title {
      font-size: 22px;
      font-weight: 700;
      color: var(--text1);
      letter-spacing: -0.4px;
      margin: 0 0 3px;
    }
    .u-sub {
      font-size: 13px;
      color: var(--text3);
      margin: 0;
    }

    /* New User button */
    .u-new {
      display: inline-flex;
      align-items: center;
      gap: 7px;
      padding: 9px 18px;
      background: var(--accent);
      color: #fff;
      border: none;
      border-radius: 8px;
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
      transition: filter 0.15s, transform 0.1s;
      flex-shrink: 0;
    }
    .u-new:hover  { filter: brightness(1.08); transform: translateY(-1px); }
    .u-new:active { filter: brightness(0.96); transform: translateY(0);    }

    /* Search */
    .u-srch-wrap { position: relative; margin-bottom: 20px; }
    .u-srch-ico  {
      position: absolute; left: 13px; top: 50%;
      transform: translateY(-50%);
      color: var(--text3); pointer-events: none;
    }
    .u-srch-inp {
      width: 100%;
      padding: 10px 14px 10px 38px;
      border: 1.5px solid var(--border);
      border-radius: 9px;
      background: var(--bg1);
      color: var(--text1);
      font-size: 13px;
      box-sizing: border-box;
      transition: border-color 0.15s, box-shadow 0.15s;
    }
    .u-srch-inp:focus {
      outline: none;
      border-color: var(--accent);
      box-shadow: 0 0 0 3px color-mix(in srgb, var(--accent) 10%, transparent);
    }

    /* Tabs */
    .u-tabs {
      display: flex;
      gap: 2px;
      border-bottom: 1.5px solid var(--border);
      margin-bottom: 14px;
      overflow-x: auto;
      scrollbar-width: none;
    }
    .u-tabs::-webkit-scrollbar { display: none; }

    .ut {
      display: inline-flex;
      align-items: center;
      gap: 7px;
      padding: 9px 16px;
      border: none;
      background: none;
      cursor: pointer;
      color: var(--text3);
      font-size: 13px;
      font-weight: 500;
      border-bottom: 2.5px solid transparent;
      margin-bottom: -1.5px;
      white-space: nowrap;
      border-radius: 0;
      transition: color 0.15s;
    }
    .ut:hover  { color: var(--text1); }
    .ut-on     { color: var(--accent); border-bottom-color: var(--accent); font-weight: 600; }
    .ut-ico    { font-size: 13px; line-height: 1; }

    .ut-ct {
      font-size: 10px;
      font-weight: 700;
      background: var(--bg3);
      color: var(--text3);
      padding: 2px 7px;
      border-radius: 20px;
      min-width: 20px;
      text-align: center;
      transition: background 0.15s, color 0.15s;
    }
    .ut-ct-on { background: var(--accent); color: #fff; }

    /* Stats row */
    .u-stats {
      font-size: 11px;
      font-weight: 600;
      color: var(--text3);
      text-transform: uppercase;
      letter-spacing: 0.4px;
      margin-bottom: 10px;
    }

    /* Cards */
    .u-cards { display: flex; flex-direction: column; gap: 6px; }

    /* User card */
    .uc {
      display: flex;
      align-items: center;
      gap: 14px;
      padding: 12px 16px 12px 0;
      background: var(--bg1);
      border: 1.5px solid var(--border);
      border-radius: 10px;
      overflow: hidden;
      transition: border-color 0.15s, box-shadow 0.15s, transform 0.12s;
      animation: fadeUp 0.15s ease both;
    }
    .uc:hover {
      border-color: color-mix(in srgb, var(--accent) 35%, var(--border));
      box-shadow: 0 2px 10px rgba(0,0,0,0.055);
      transform: translateY(-1px);
    }

    /* Left department bar */
    .uc-bar {
      width: 4px;
      align-self: stretch;
      flex-shrink: 0;
      border-radius: 0 3px 3px 0;
    }

    /* Avatar */
    .uc-av {
      width: 40px;
      height: 40px;
      border-radius: 10px;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 13px;
      font-weight: 700;
      flex-shrink: 0;
      letter-spacing: 0.5px;
    }

    /* Info block */
    .uc-info    { flex: 1; min-width: 0; }
    .uc-name    {
      font-size: 14px;
      font-weight: 600;
      color: var(--text1);
      margin-bottom: 3px;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    .uc-meta  { display: flex; align-items: center; gap: 5px; font-size: 11px; color: var(--text3); }
    .uc-uname { font-family: var(--font-mono); }
    .uc-short { font-family: var(--font-mono); font-weight: 700; color: var(--text2); }
    .uc-dot   { opacity: 0.3; }

    /* Role badge */
    .uc-badge {
      font-size: 11px;
      font-weight: 600;
      padding: 3px 10px;
      border-radius: 20px;
      flex-shrink: 0;
      white-space: nowrap;
    }

    /* Action buttons — fade in on hover */
    .uc-actions {
      display: flex;
      gap: 6px;
      flex-shrink: 0;
      opacity: 0;
      transition: opacity 0.15s;
    }
    .uc:hover .uc-actions { opacity: 1; }

    .uc-btn {
      display: inline-flex;
      align-items: center;
      gap: 5px;
      padding: 5px 11px;
      border-radius: 6px;
      font-size: 12px;
      font-weight: 500;
      cursor: pointer;
      border: 1.5px solid transparent;
      transition: background 0.12s, border-color 0.12s;
    }
    .uc-edit {
      background: var(--bg2);
      color: var(--text2);
      border-color: var(--border);
    }
    .uc-edit:hover {
      background: var(--bg3);
      color: var(--text1);
      border-color: var(--accent);
    }
    .uc-del {
      background: #fff1f2;
      color: #be123c;
      border-color: #fecdd3;
    }
    .uc-del:hover {
      background: #ffe4e6;
      border-color: #fb7185;
    }

    /* Dark mode card + delete button overrides */
    [data-theme="dark"] .uc          { background: var(--bg2); }
    [data-theme="dark"] .uc-del      { background: #2d1215; color: #fb7185; border-color: #4c1d24; }
    [data-theme="dark"] .uc-del:hover{ background: #3d1a1f; border-color: #f43f5e; }

    /* Empty state */
    .u-empty {
      display: flex;
      flex-direction: column;
      align-items: center;
      padding: 64px 20px;
      gap: 8px;
      text-align: center;
    }
    .u-empty-title { font-size: 15px; font-weight: 600; color: var(--text2); }
    .u-empty-sub   { font-size: 12px; color: var(--text3); }
    .u-clear {
      margin-top: 4px;
      background: none;
      border: none;
      color: var(--accent);
      font-size: 12px;
      font-weight: 600;
      cursor: pointer;
      text-decoration: underline;
    }
  `;
}

// ── Add / Edit modal ───────────────────────────────────────────
function showUserModal(user) {
  var isEdit = !!user;
  var u      = user || {};

  var existing = document.getElementById("user-modal-overlay");
  if (existing) existing.remove();

  var roleOptions = ROLES.map(function(r) {
    return '<option value="' + r.id + '"' + (r.id === u.role ? " selected" : "") + '>' + r.label + '</option>';
  }).join("");

  var overlay = document.createElement("div");
  overlay.id        = "user-modal-overlay";
  overlay.className = "modal-overlay";
  overlay.innerHTML = [
    '<div class="modal" style="max-width:440px;width:90%">',
      '<div class="modal-header">',
        '<div class="modal-title">' + (isEdit ? "Edit User" : "New User") + '</div>',
        '<button class="btn btn-ghost btn-sm" onclick="document.getElementById(\'user-modal-overlay\').remove()">&times;</button>',
      '</div>',
      '<div class="modal-body" style="display:flex;flex-direction:column;gap:14px">',
        '<div id="um-err" class="alert alert-error hidden"></div>',
        '<div class="form-group">',
          '<label class="form-label">Full Name</label>',
          '<input class="form-input" id="um-name" type="text" value="' + h(u.name || "") + '" placeholder="Full Name" autocomplete="off">',
        '</div>',
        '<div class="form-group">',
          '<label class="form-label">Short Name / Initials <span style="color:var(--text3);font-weight:400">(must match Excel)</span></label>',
          '<input class="form-input" id="um-short" type="text" value="' + h(u.short_name || "") + '" placeholder="e.g. KDM" style="text-transform:uppercase;font-family:var(--font-mono);letter-spacing:1px" autocomplete="off">',
        '</div>',
        '<div class="form-group">',
          '<label class="form-label">Username</label>',
          '<input class="form-input" id="um-username" type="text" value="' + h(u.username || "") + '" placeholder="username"' +
            (isEdit ? " readonly style='opacity:0.5;cursor:not-allowed'" : "") + ' autocomplete="off">',
        '</div>',
        '<div class="form-group">',
          '<label class="form-label">Password ' +
            (isEdit ? '<span style="color:var(--text3);font-weight:400">(blank = keep current)</span>' : '') +
          '</label>',
          '<input class="form-input" id="um-pass" type="password" placeholder="••••••••" autocomplete="new-password">',
        '</div>',
        '<div class="form-group">',
          '<label class="form-label">Role</label>',
          '<select class="form-select" id="um-role">' + roleOptions + '</select>',
        '</div>',
      '</div>',
      '<div class="modal-footer">',
        '<button class="btn btn-secondary" onclick="document.getElementById(\'user-modal-overlay\').remove()">Cancel</button>',
        '<button class="btn btn-primary" id="um-save-btn" onclick="doSaveUser(\'' + h(u.username || "") + '\')">',
          isEdit ? "Save Changes" : "Create User",
        '</button>',
      '</div>',
    '</div>',
  ].join("");

  overlay.addEventListener("click", function(e) {
    if (e.target === overlay) overlay.remove();
  });
  document.body.appendChild(overlay);

  setTimeout(function() {
    var f = document.getElementById("um-name");
    if (f && !f.value) f.focus();
  }, 50);
}

// ── Save user ──────────────────────────────────────────────────
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
    errEl.textContent = "Please fill all required fields.";
    errEl.classList.remove("hidden");
    return;
  }
  if (!isEdit && password.length < 6) {
    errEl.textContent = "Password must be at least 6 characters.";
    errEl.classList.remove("hidden");
    return;
  }

  btn.disabled  = true;
  btn.innerHTML = '<span class="spinner" style="width:13px;height:13px;border-width:2px"></span> Saving...';

  try {
    if (isEdit) {
      var updates = { name: name, short_name: short, role: role };
      if (password) updates.password = password;
      await API.updateUser(existingUsername, updates);
      toast("User updated ✓");
    } else {
      await API.createUser({ name: name, short_name: short, username: username, password: password, role: role });
      toast("User created ✓");
    }
    document.getElementById("user-modal-overlay").remove();
    var savedTab = _usersState.activeTab;
    await renderUsers();
    _usersState.activeTab = savedTab;
    _renderUserList();
  } catch (err) {
    errEl.textContent = err.message || "Failed to save. Please try again.";
    errEl.classList.remove("hidden");
    btn.disabled  = false;
    btn.innerHTML = isEdit ? "Save Changes" : "Create User";
  }
}

// ── Delete user — proper modal ─────────────────────────────────
async function doDeleteUser(username, name) {
  var existing = document.getElementById("del-overlay");
  if (existing) existing.remove();

  var overlay = document.createElement("div");
  overlay.id        = "del-overlay";
  overlay.className = "modal-overlay";
  overlay.innerHTML = [
    '<div class="modal" style="max-width:380px;width:90%">',
      '<div class="modal-header" style="border-bottom-color:#fee2e2">',
        '<div class="modal-title" style="color:#dc2626">Delete User</div>',
      '</div>',
      '<div class="modal-body">',
        '<p style="margin:0;font-size:14px;color:var(--text1);line-height:1.6">',
          'Delete <strong>' + h(name) + '</strong>?<br>',
          '<span style="font-size:12px;color:var(--text3)">This action cannot be undone.</span>',
        '</p>',
      '</div>',
      '<div class="modal-footer">',
        '<button class="btn btn-secondary" onclick="document.getElementById(\'del-overlay\').remove()">Cancel</button>',
        '<button class="btn btn-danger" id="del-ok-btn">Delete</button>',
      '</div>',
    '</div>',
  ].join("");

  overlay.addEventListener("click", function(e) {
    if (e.target === overlay) overlay.remove();
  });
  document.body.appendChild(overlay);

  document.getElementById("del-ok-btn").onclick = async function() {
    var btn   = document.getElementById("del-ok-btn");
    btn.disabled  = true;
    btn.innerHTML = '<span class="spinner" style="width:13px;height:13px;border-width:2px"></span>';
    try {
      await API.deleteUser(username);
      overlay.remove();
      toast("User deleted");
      var savedTab = _usersState.activeTab;
      await renderUsers();
      _usersState.activeTab = savedTab;
      _renderUserList();
    } catch (err) {
      overlay.remove();
      toast(err.message || "Failed to delete", "error");
    }
  };
}
