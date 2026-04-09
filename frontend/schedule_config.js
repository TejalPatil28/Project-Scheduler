// =============================================================
//  schedule_config.js  —  Department Schedule Sheet Configuration
//
//  Single source of truth for ALL department project schedule
//  sheet rendering in renderExcelMirror (app.js).
//
//  To add a new department: copy the SW block, change the key
//  and fill in the correct values.
//
//  Consumed by:
//    • app.js → renderExcelMirror uses getScheduleConfig(role)
// =============================================================

var SCHEDULE_CONFIGS = {

  // ═══════════════════════════════════════════════════════════
  //  SW  —  Software Engineering Department
  // ═══════════════════════════════════════════════════════════
  SW: {

    // ── Task row range ────────────────────────────────────────
    taskStartRow: 9,
    taskEndRow:   55,

    // ── Columns to skip entirely (hidden in Excel) ────────────
    skipCols: ["AC", "AE"],

    // ── Editable columns (task rows only) ────────────────────
    // "date"   → date picker  (X, Y)
    // "number" → 0-100 input  (Z)
    // "dropdown" → options list (AD)
    // "text"   → free text    (AF)
    editableCols: {
      "X":  { type: "date"     },
      "Y":  { type: "date"     },
      "Z":  { type: "number",  min: 0, max: 100 },
      "AD": { type: "dropdown" },   // options come from helpDropdown
      "AF": { type: "text"     },
    },

    // ── Help-required dropdown (the AD column) ────────────────
    helpDropdown: {
      col:     "AD",
      options: ["Engineering","Purchase","Software","Project Management","Manufacturing","Sales","Client"],
      colorsDark: {
        "engineering":        "#3d1a1a",
        "purchase":           "#1a2e14",
        "software":           "#0e2233",
        "project management": "#0a1a30",
        "manufacturing":      "#1a2a14",
        "sales":              "#003333",
        "client":             "#2a0a2a",
      },
      colorsLight: {
        "engineering":        "#ffb3b3",
        "purchase":           "#5f933c",
        "software":           "#0096cc",
        "project management": "#005fa3",
        "manufacturing":      "#2f491e",
        "sales":              "#00d9d9",
        "client":             "#6d006d",
      },
      fgDark: {
        "engineering":        "#f87171",
        "purchase":           "#86efac",
        "software":           "#60a5fa",
        "project management": "#93c5fd",
        "manufacturing":      "#a3e635",
        "sales":              "#2dd4bf",
        "client":             "#d8b4fe",
      },
    },

    // ── Column widths (percentage-based) ─────────────────────
    colWidths: {
      _default: "5%",
      "E":  "2%",
      "F":  "2%",
      "G":  "2%",
      "H":  "2.5%",
      "I":  "6%",
      "J":  "5%",
      "K":  "5%",
      "L":  "5%",
      "M":  "5%",
      "N":  "5%",
      "O":  "5%",
      "P":  "4%",
      "Q":  "4%",
      "R":  "4%",
      "S":  "2.5%",
      "T":  "2%",
      "U":  "2.5%",
      "V":  "4%",
      "W":  "4%",
      "X":  "4.5%",
      "Y":  "4.5%",
      "Z":  "3%",
      "AA": "3%",
      "AB": "4%",
      "AD": "5%",
      "AF": "7%",
    },

    // ── Collapsible column groups ─────────────────────────────
    colGroups: [
      { key: "planning_dates", label: "Planning Dates", cols: ["J","K","L","M","N","O","P","Q","R"] },
    ],

    // ── Custom header row definitions ─────────────────────────
    // dark/light keyed separately — bg/fg per column
    customHeaders: {
      dark: {
        "E":  { label: "Buffer Time",          bg: "#2a2010", fg: "#d4a96a" },
        "F":  { label: "PH",                   bg: "#2a2010", fg: "#d4a96a" },
        "G":  { label: "TFO",                  bg: "#2a2010", fg: "#d4a96a" },
        "H":  { label: "Phase ID",             bg: "#2a2010", fg: "#d4a96a" },
        "I":  { label: "Task Description",     bg: "#2a2010", fg: "#d4a96a" },
        "J":  { label: "P1 Start Date",        bg: "#2a2010", fg: "#d4a96a" },
        "K":  { label: "P1 End Date",          bg: "#2a2010", fg: "#d4a96a" },
        "L":  { label: "P2 Start Date",        bg: "#2a2010", fg: "#d4a96a" },
        "M":  { label: "P2 End Date",          bg: "#2a2010", fg: "#d4a96a" },
        "N":  { label: "P3 Start Date",        bg: "#2a2010", fg: "#d4a96a" },
        "O":  { label: "P3 End Date",          bg: "#2a2010", fg: "#d4a96a" },
        "P":  { label: "Org Plan Date",        bg: "#2a2010", fg: "#d4a96a" },
        "Q":  { label: "Org End Date",         bg: "#2a2010", fg: "#d4a96a" },
        "R":  { label: "Ref. Lead Time",       bg: "#2a2010", fg: "#d4a96a" },
        "S":  { label: "Lead Time",            bg: "#2a2010", fg: "#d4a96a" },
        "T":  { label: "Intlk",               bg: "#2a2010", fg: "#d4a96a" },
        "U":  { label: "Effort Days",          bg: "#2a2010", fg: "#d4a96a" },
        "V":  { label: "Cur. Start Date",      bg: "#2a2010", fg: "#d4a96a" },
        "W":  { label: "Cur. End Date",        bg: "#2a2010", fg: "#d4a96a" },
        "X":  { label: "Act. Start Date",      bg: "#2a2010", fg: "#d4a96a" },
        "Y":  { label: "Act. End Date",        bg: "#2a2010", fg: "#d4a96a" },
        "Z":  { label: "% Complete",           bg: "#2a2010", fg: "#d4a96a" },
        "AA": { label: "Exptd % Completion",   bg: "#2a2010", fg: "#d4a96a" },
        "AB": { label: "Alert Date for 80%",   bg: "#2a2010", fg: "#d4a96a" },
        "AD": { label: "Help Req. from",       bg: "#2a2010", fg: "#d4a96a" },
        "AF": { label: "Remark",               bg: "#2a2010", fg: "#d4a96a" },
      },
      light: {
        "E":  { label: "Buffer Time",          bg: "#fcd5b4", fg: "#000000" },
        "F":  { label: "PH",                   bg: "#fcd5b4", fg: "#000000" },
        "G":  { label: "TFO",                  bg: "#fcd5b4", fg: "#000000" },
        "H":  { label: "Phase ID",             bg: "#fcd5b4", fg: "#000000" },
        "I":  { label: "Task Description",     bg: "#fcd5b4", fg: "#000000" },
        "J":  { label: "P1 Start Date",        bg: "#fcd5b4", fg: "#000000" },
        "K":  { label: "P1 End Date",          bg: "#fcd5b4", fg: "#000000" },
        "L":  { label: "P2 Start Date",        bg: "#fcd5b4", fg: "#000000" },
        "M":  { label: "P2 End Date",          bg: "#fcd5b4", fg: "#000000" },
        "N":  { label: "P3 Start Date",        bg: "#fcd5b4", fg: "#000000" },
        "O":  { label: "P3 End Date",          bg: "#fcd5b4", fg: "#000000" },
        "P":  { label: "Org Plan Date",        bg: "#fcd5b4", fg: "#000000" },
        "Q":  { label: "Org End Date",         bg: "#fcd5b4", fg: "#000000" },
        "R":  { label: "Ref. Lead Time",       bg: "#fcd5b4", fg: "#000000" },
        "S":  { label: "Lead Time",            bg: "#fcd5b4", fg: "#000000" },
        "T":  { label: "Intlk",               bg: "#fcd5b4", fg: "#000000" },
        "U":  { label: "Effort Days",          bg: "#fcd5b4", fg: "#000000" },
        "V":  { label: "Cur. Start Date",      bg: "#fcd5b4", fg: "#000000" },
        "W":  { label: "Cur. End Date",        bg: "#fcd5b4", fg: "#000000" },
        "X":  { label: "Act. Start Date",      bg: "#fcd5b4", fg: "#000000" },
        "Y":  { label: "Act. End Date",        bg: "#fcd5b4", fg: "#000000" },
        "Z":  { label: "% Complete",           bg: "#fcd5b4", fg: "#000000" },
        "AA": { label: "Exptd % Completion",   bg: "#fcd5b4", fg: "#000000" },
        "AB": { label: "Alert Date for 80%",   bg: "#fcd5b4", fg: "#000000" },
        "AD": { label: "Help Req. from",       bg: "#fcd5b4", fg: "#000000" },
        "AF": { label: "Remark",               bg: "#fcd5b4", fg: "#000000" },
      },
    },

    // ── Cell background color rules ───────────────────────────
    // Applied to task rows (9+). First match wins.
    colorRules: {
      dark: [
        { cols: ["E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","AA","AB"], rows: "9+", color: "#2a2a2a" },
        { cols: ["X","Y","Z"], rows: "9+", color: "#152030" },
      ],
      light: [
        { cols: ["E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","AA","AB"], rows: "9+", color: "#d9d9d9" },
        { cols: ["X","Y","Z"], rows: "9+", color: "#90b4df" },
      ],
    },

    // ── Cell text color rules (conditional) ───────────────────
    // condition: "not_equal" | "less_than"
    // compare_with: column to compare against (same row)
    textRules: [
      {
        name:         "Lead Time mismatch",
        col:          "S",
        condition:    "not_equal",
        compareWith:  "R",
        rows:         "9+",
        color:        "var(--red)",
      },
      {
        name:         "Progress behind plan",
        col:          "Z",
        condition:    "less_than",
        compareWith:  "AA",
        rows:         "9+",
        color:        "var(--red)",
      },
    ],

    // ── Master list column keys (for filterMasterList) ────────
    // Which keys from the monitor JSON to match against user initials
    masterListCols: {
      head: "swh_head",   // head/admin filter column
      tl:   "swe_name",   // TL filter column
    },

  }, // end SW


  // ═══════════════════════════════════════════════════════════
  //  HW  —  Hardware Engineering Department  (placeholder)
  // ═══════════════════════════════════════════════════════════
  HW: null,   // TODO: fill in when HW schedule config is ready


  // ═══════════════════════════════════════════════════════════
  //  MFG  —  Manufacturing Department  (placeholder)
  // ═══════════════════════════════════════════════════════════
  MFG: null,  // TODO: fill in when MFG schedule config is ready


  // ═══════════════════════════════════════════════════════════
  //  PM  —  Project Management Department  (placeholder)
  // ═══════════════════════════════════════════════════════════
  PM: null,   // TODO: fill in when PM schedule config is ready

};


// ── Role → department lookup ──────────────────────────────────
var ROLE_TO_SCHEDULE_DEPT = {
  "sw_tl":    "SW",
  "sw_head":  "SW",
  "hw_tl":    "HW",
  "hw_head":  "HW",
  "mfg_tl":   "MFG",
  "mfg_head": "MFG",
  "pm":       "PM",
  "pm_head":  "PM",
  "admin":    "SW",   // admin defaults to SW
  "head":     "SW",
};

// ── Role → display label ──────────────────────────────────────
// Single source of truth for all role display names.
// Used by renderShell, popover, topbar chip.
var ROLE_LABELS = {
  "admin":    "Admin",
  "head":     "Head",
  "sw_tl":    "SW Team Lead",
  "sw_head":  "SW Head",
  "hw_tl":    "HW Team Lead",
  "hw_head":  "HW Head",
  "mfg_tl":   "MFG Team Lead",
  "mfg_head": "MFG Head",
  "pm":       "Project Manager",
  "pm_head":  "PM Head",
};

// ── Public helpers ────────────────────────────────────────────

// Returns the schedule config for a given role.
// Falls back to SW if dept config is null (not yet defined).
function getScheduleConfig(role) {
  var dept = ROLE_TO_SCHEDULE_DEPT[role] || "SW";
  return SCHEDULE_CONFIGS[dept] || SCHEDULE_CONFIGS["SW"];
}

// Returns the display label for a role.
function userRoleLabel(role) {
  return ROLE_LABELS[role] || role;
}
