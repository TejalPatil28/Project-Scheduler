// ── Users Page ─────────────────────────────────────────────────
// Extracted from app.js — full page route for admin user management.
// Called via switchToUsers() in app.js.

var VALID_ROLES = ["admin", "sw_head", "hw_head", "mfg_head", "pm_head", "pm", "hw_tl", "sw_tl", "mfg_tl"];

var ROLE_LABELS = {
  admin:    "Admin",
  sw_head:  "SW Head",
  hw_head:  "HW Head",
  mfg_head: "MFG Head",
  pm_head:  "PM Head",
  pm:       "Project Manager",
  hw_tl:    "HW Team Lead",
  sw_tl:    "SW Team Lead",
  mfg_tl:   "MFG Team Lead",
};

var ROLE_BADGE = {
  admin:    "badge-red",
  sw_head:  "badge-purple",
  hw_head:  "badge-blue",
  mfg_head: "badge-green",
  pm_head:  "badge-teal",
  pm:       "badge-teal",
  hw_tl:    "badge-amber",
  sw_tl:    "badge-purple",
  mfg_tl:   "badge-gray",
};

function userRoleLabel(role) {
  return ROLE_LABELS[role] || role;
}

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

  var rows = users.map(function (u) {
    var uJson = h(JSON.stringify(u));
    return [
      '<div style="background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);padding:14px 16px;display:flex;align-items:center;gap:14px;animation:fadeUp 0.3s ease">',
        '<div class="avatar avatar-md">' + h(u.name[0].toUpperCase()) + '</div>',
        '<div style="flex:1;min-width:0">',
          '<div style="font-weight:600;font-size:14px;margin-bottom:2px">' + h(u.name) + '</div>',
          '<div style="font-size:11px;color:var(--text2);font-family:var(--font-mono)">@' + h(u.username) + ' &middot; ' + h(u.short_name) + '</div>',
        '</div>',
        '<span class="badge ' + (ROLE_BADGE[u.role] || 'badge-gray') + '">' + userRoleLabel(u.role) + '</span>',
        '<div style="display:flex;gap:6px;flex-shrink:0">',
          '<button class="btn btn-ghost btn-sm" onclick=\'showUserModal(' + uJson + ')\'>Edit</button>',
          '<button class="btn btn-danger btn-sm" onclick="doDeleteUser(\'' + h(u.username) + '\',\'' + h(u.name) + '\')">Delete</button>',
        '</div>',
      '</div>',
    ].join("");
  }).join("");

  page.innerHTML = [
    '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px">',
      '<div class="dash-title">Users <span style="color:var(--text3);font-size:14px;font-weight:400">(' + users.length + ')</span></div>',
      '<button class="btn btn-primary btn-sm" onclick="showUserModal()">+ New User</button>',
    '</div>',
    '<div style="display:flex;flex-direction:column;gap:8px">' + rows + '</div>',
  ].join("");
}

function showUserModal(user) {
  var isEdit = !!user;
  var u = user || {};

  var existing = document.getElementById("user-modal-overlay");
  if (existing) existing.remove();

  var roleOptions = VALID_ROLES.map(function (r) {
    return '<option value="' + r + '"' + (r === u.role ? " selected" : "") + '>' + userRoleLabel(r) + '</option>';
  }).join("");

  var overlay = document.createElement("div");
  overlay.id = "user-modal-overlay";
  overlay.className = "modal-overlay";
  overlay.innerHTML = [
    '<div class="modal">',
      '<div class="modal-header">',
        '<div class="modal-title">' + (isEdit ? "Edit User" : "New User") + '</div>',
        '<button class="btn btn-ghost btn-sm" onclick="document.getElementById(\'user-modal-overlay\').remove()">&times;</button>',
      '</div>',
      '<div class="modal-body">',
        '<div id="um-err" class="alert alert-error hidden"></div>',
        '<div class="form-group">',
          '<label class="form-label">Full Name</label>',
          '<input class="form-input" id="um-name" type="text" value="' + h(u.name || "") + '" placeholder="Full Name">',
        '</div>',
        '<div class="form-group">',
          '<label class="form-label">Short Name / Initials <span style="color:var(--text3)">(must match Excel)</span></label>',
          '<input class="form-input" id="um-short" type="text" value="' + h(u.short_name || "") + '" placeholder="KDM" style="text-transform:uppercase;font-family:var(--font-mono)">',
        '</div>',
        '<div class="form-group">',
          '<label class="form-label">Username</label>',
          '<input class="form-input" id="um-username" type="text" value="' + h(u.username || "") + '" placeholder="username"' + (isEdit ? " readonly style='opacity:0.5'" : "") + '>',
        '</div>',
        '<div class="form-group">',
          '<label class="form-label">Password ' + (isEdit ? "(blank = keep current)" : "") + '</label>',
          '<input class="form-input" id="um-pass" type="password" placeholder="&bull;&bull;&bull;&bull;&bull;&bull;&bull;&bull;">',
        '</div>',
        '<div class="form-group">',
          '<label class="form-label">Role</label>',
          '<select class="form-select" id="um-role">' + roleOptions + '</select>',
        '</div>',
      '</div>',
      '<div class="modal-footer">',
        '<button class="btn btn-secondary" onclick="document.getElementById(\'user-modal-overlay\').remove()">Cancel</button>',
        '<button class="btn btn-primary" id="um-save-btn" onclick="doSaveUser(\'' + h(u.username || "") + '\')">' + (isEdit ? "Save Changes" : "Create User") + '</button>',
      '</div>',
    '</div>',
  ].join("");

  overlay.addEventListener("click", function (e) {
    if (e.target === overlay) overlay.remove();
  });
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
      var updates = { name: name, short_name: short, role: role };
      if (password) updates.password = password;
      await API.updateUser(existingUsername, updates);
      toast("User updated \u2713");
    } else {
      await API.createUser({ name: name, short_name: short, username: username, password: password, role: role });
      toast("User created \u2713");
    }
    document.getElementById("user-modal-overlay").remove();
    renderUsers();
  } catch (err) {
    errEl.textContent = err.message || "Failed";
    errEl.classList.remove("hidden");
    btn.disabled = false;
    btn.textContent = isEdit ? "Save Changes" : "Create User";
  }
}

async function doDeleteUser(username, name) {
  if (!confirm('Delete "' + name + '"? This cannot be undone.')) return;
  try {
    await API.deleteUser(username);
    toast("User deleted");
    renderUsers();
  } catch (err) {
    toast(err.message, "error");
  }
}
