// =============================================================
//  monitor_config.js  —  Department Monitor Configuration
//
//  Single source of truth for ALL department monitor screens.
//  To add a new department: copy the SW block, change the key
//  (e.g. "PM", "HW", "MFG") and fill in the correct values.
//
//  Consumed by:
//    • monitor.js  → renderMonitor(container, data, user, overdueMap, config)
//    • app.js      → passes MONITOR_CONFIGS[dept] to renderMonitor
// =============================================================

var MONITOR_CONFIGS = {

  // ═══════════════════════════════════════════════════════════
  //  SW  —  Software Engineering Department
  // ═══════════════════════════════════════════════════════════
  SW: {

    // ── File / sheet structure ────────────────────────────────
    // These match excel_db.py DEPT_CONFIG["SW"]
    headerRow:   10,   // Row that holds column label headers
    dataStart:   12,   // First row with project data
    idCol:       "B",  // Column shown as the project ID in UI (OR number)
    headCol:     "C",  // Column for department head assignment
    tlCol:       "D",  // Column for team lead assignment (also used for row filtering)

    // ── Role identifiers ──────────────────────────────────────
    // Used to check user.role for editability and row filtering
    headRole:    "sw_head",
    tlRole:      "sw_tl",

    // ── Tab definitions ───────────────────────────────────────
    // Each tab describes: which columns to show, layout mode,
    // and optional collapsible column groups.
    tabs: [
      {
        id:     "overview",
        label:  "Project Overview",
        icon:   "\uD83D\uDDC2",   // 🗂
        scroll: true,
        wrapHeaders: true,
        cols: [
          "C","D","E","F","G",
          "H","I","J","K","L","M",
          "N","O",
          "P","Q","R","S","T","U","V",
          "W","X","Y",
          "AG","AH","AI",
          "AJ","AK","AL","AM",
          "AN","AO","AP",
          "AQ","AR",
          "BA"
        ],
        colGroups: [
          { key: "info1",       label: "Project Info",      cols: ["C","D"] },
          { key: "info",        label: "Project Info",      cols: ["H","I","J","K","L","M"] },
          { key: "ld",          label: "LD Details",        cols: ["P","Q","R","S","T"] },
          { key: "ld1",         label: "LD Details",        cols: ["U","V"] },
          { key: "ld2",         label: "LD Details",        cols: ["W","X","Y"] },
          { key: "additional",  label: "Additional Fields", cols: ["AJ","AK","AL","AM"] },
          { key: "additional1", label: "Additional Fields", cols: ["AN","AO","AP"] },
          { key: "additional2", label: "Additional Fields", cols: ["AQ","AR"] },
        ],
      },
      {
        id:    "team",
        label: "Team & Roles",
        icon:  "\uD83D\uDC65",   // 👥
        fluid: true,
        cols:  ["F","C","D","Z","AA","AB","AC","AD","AE","AF"],
      },
      {
        id:    "milestone",
        label: "Milestone Progress",
        icon:  "\uD83C\uDFC1",   // 🏁
        fluid: true,
        cols:  [
          "F",
          "BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK",
          "BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV","BW"
        ],
      },
    ],

    // ── Column widths ─────────────────────────────────────────
    // Keys must match tab ids above.
    // "overview_pct" is the % layout used for the overview tab.
    // "_default" is the fallback for any column not listed.
    colWidths: {
      overview: {
        _default: "90px",
        "C": "80px",  "D": "80px",  "E": "80px",
        "F": "40px",  "G": "80px",  "H": "50px",
        "I": "90px",  "J": "90px",  "K": "90px",  "L": "90px",
        "M": "80px",  "N": "90px",  "O": "50px",
        "P": "90px",  "Q": "90px",  "R": "90px",  "S": "90px",
        "T": "70px",  "U": "100px", "V": "70px",  "W": "100px",
        "X": "100px", "Y": "70px",
        "AG": "90px", "AH": "50px", "AI": "40px",
        "AJ": "90px", "AK": "90px", "AL": "90px", "AM": "90px",
        "AN": "90px", "AO": "90px", "AP": "90px",
        "AQ": "100px","AR": "120px","BA": "90px",
      },
      overview_pct: {
        _default: "5%",
        "F": "15%",
        "C": "4%",  "D": "4%",  "E": "6%",  "G": "5%",
        "H": "4%",  "I": "8%",  "J": "5%",  "K": "10%", "L": "10%",
        "M": "5%",  "N": "6%",  "O": "5%",
        "P": "5%",  "Q": "5%",  "R": "5%",  "S": "10%", "T": "5%",
        "U": "5%",  "V": "5%",  "W": "5%",  "X": "5%",  "Y": "5%",
        "AG": "5%", "AH": "4%", "AI": "4%",
        "AJ": "5%", "AK": "5%", "AL": "5%", "AM": "6%",
        "AN": "5%", "AO": "5%", "AP": "6%",
        "AQ": "5%", "AR": "6%",
        "BA": "7%",
      },
      team: {
        _default: "9.44%",
        "F": "15%",
        "C": "9.44%", "D": "9.44%",
        "Z": "9.44%", "AA": "9.44%", "AB": "9.44%", "AC": "9.44%",
        "AD": "9.44%","AE": "9.44%", "AF": "9.44%",
      },
      milestone: {
        _default: "3.86%",
        "F": "15%",
      },
    },

    // ── Editable columns ──────────────────────────────────────
    // Each key is a column letter.
    // type:
    //   "tl_dropdown"  → populated dynamically from all values in config.tlCol
    //   "dropdown"     → static options list provided here
    //   "date"         → HTML date input
    //   "text"         → free-text input
    // role:
    //   The user.role that is allowed to edit this column.
    //   null means any authenticated user can edit.
    editableCols: {
      "C":  { type: "tl_dropdown", role: "sw_head" },
      "D":  { type: "tl_dropdown", role: "sw_tl"  },
      "E":  { type: "date",        role: "sw_tl"  },
      "N":  { type: "dropdown",    role: "sw_tl",  options: ["", "SYSTEM", "SERVICE", "E & C"] },
      "AP": { type: "dropdown",    role: "sw_tl",  options: ["", "RUN", "HOLD", "CLOSED"] },
      "AR": { type: "text",        role: "sw_tl"  },
    },

    // ── Milestone columns ─────────────────────────────────────
    // These columns render as progress bars instead of plain text.
    // Must match the milestone tab's cols (excluding the idCol).
    milestoneCols: [
      "BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK",
      "BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV","BW"
    ],

    // ── Sum row columns ───────────────────────────────────────
    // These columns show a running total in the header sum row
    // on the overview tab.
    sumCols: ["M", "Y", "AH", "AI", "AJ", "R"],

    // ── Cell format rules ─────────────────────────────────────
    // Applied in order. First match wins.
    // format types:
    //   "decimal2"        → toFixed(2)
    //   "pct_from_decimal"→ multiply by 100, show as "xx%"
    formatRules: [
      { cols: ["M", "R", "AJ"],              format: "decimal2"         },
      { cols: ["P", "Q", "AK", "AL", "AO"], format: "pct_from_decimal" },
    ],

    // ── Cell background color rules ───────────────────────────
    // dark / light keyed separately.
    // Applied in order — first match wins.
    colorRules: {
      dark: [
        { cols: ["H","I","J","K","L"],  bg: "#2a2a2a" },
        { cols: ["P","Q","R","S"],      bg: "#152030" },
        { cols: ["M","N","O"],          bg: "#1f2a2a" },
      ],
      light: [
        { cols: ["M","G","K","L","F","O","P","Q","R","S","T","U","V","W",
                 "X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI",
                 "AJ","AK","AL","AM","AN","AO","AQ","BA"],  bg: "#747070" },
        { cols: ["E","N","I","J","C","D","AR","AP"],        bg: "#90b4df" },
        { cols: ["H"],                                       bg: "#191616" },
      ],
    },

    // ── Cell text color rules ─────────────────────────────────
    textColorRules: {
      dark: [
        { cols: ["S"],  color: "#ff9999" },
        { cols: ["R"],  color: "#66ccff" },
        { cols: ["M"],  color: "#ffcc66" },
      ],
      light: [
        { cols: ["F","H","O","P","Q","R","S","T","U","V","W","X","Y",
                 "Z","AA","AB","AC","AD","AE","AF","AG","AH","AI",
                 "AJ","AK","AL","AM","AN","AO","AQ"],  color: "#ffba26" },
        { cols: ["M","G"],                              color: "#ffffff" },
      ],
    },

  }, // end SW


  // ═══════════════════════════════════════════════════════════
  //  HW  —  Hardware Engineering Department  (placeholder)
  // ═══════════════════════════════════════════════════════════
  HW: null,   // TODO: fill in when HW monitor is ready

  // ═══════════════════════════════════════════════════════════
  //  MFG  —  Manufacturing Department  (placeholder)
  // ═══════════════════════════════════════════════════════════
  MFG: null,  // TODO: fill in when MFG monitor is ready

  // ═══════════════════════════════════════════════════════════
  //  PM  —  Project Management Department
  //
  //  Monitor file prefix : PL_Monitor
  //  idCol               : B  (Sch File Name)
  //  headCol             : C  (PM Head)
  //  tlCol               : D  (PM Name)
  //  Milestone cols      : BB – CU  (46 tasks)
  // ═══════════════════════════════════════════════════════════
  PM: {

    // ── File / sheet structure ────────────────────────────────
    headerRow:   10,   // Same row layout as SW
    dataStart:   12,   // Same as SW
    idCol:       "B",  // Sch File Name
    headCol:     "C",  // PM Head
    tlCol:       "D",  // PM Name

    // ── Role identifiers ──────────────────────────────────────
    headRole:    "pm_head",
    tlRole:      "pm",

    // ── Tab definitions ───────────────────────────────────────
    tabs: [
      {
        id:     "overview",
        label:  "Project Overview",
        icon:   "\uD83D\uDDC2",   // 🗂
        scroll: true,
        wrapHeaders: true,
        cols: [
          "C","D","E","F","G","H","I",
          "J","K","L","M",
          "N","O","P","Q","R",
          "S","T",
          "U","V","W","X","Y",
          "AG","AH","AI","AJ",
          "AK","AL","AM",
          "AN","AO","AP",
          "AQ","AR","AS",
          "AZ"
        ],
        colGroups: [
          { key: "identity",   label: "Project Identity",  cols: ["C","D"] },
          { key: "sales",      label: "Sales & Customer",  cols: ["G","H","I","J","K"] },
          { key: "financial",  label: "Financial",         cols: ["L","M","N"] },
          { key: "ld",         label: "LD Details",        cols: ["O","P","Q","R","S"] },
          { key: "dispatch",   label: "Dispatch",          cols: ["T","U","V","W","X"] },
          { key: "sw_engg",    label: "SW Engineering",    cols: ["AH","AI","AJ"] },
          { key: "hw_engg",    label: "HW Engineering",    cols: ["AK","AL","AM"] },
          { key: "shop",       label: "Shop & Mfg",        cols: ["AN","AO","AP","AQ"] },
          { key: "status",     label: "Project Status",    cols: ["AR","AS"] },
        ],
      },
      {
        id:    "team",
        label: "Team & Roles",
        icon:  "\uD83D\uDC65",   // 👥
        fluid: true,
        cols:  ["E","C","D","Z","AA","AB","AC","AD","AE","AF"],
      },
      {
        id:    "milestone",
        label: "Milestone Progress",
        icon:  "\uD83C\uDFC1",   // 🏁
        fluid: true,
        cols:  [
          "E",
          "BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK",
          "BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU",
          "BV","BW","BX","BY","BZ","CA","CB","CC","CD","CE",
          "CF","CG","CH","CI","CJ","CK","CL","CM","CN","CO",
          "CP","CQ","CR","CS","CT","CU"
        ],
      },
    ],

    // ── Column widths ─────────────────────────────────────────
    colWidths: {
      overview: {
        _default: "90px",
        "C": "80px",  "D": "80px",  "E": "80px",  "F": "50px",
        "G": "80px",  "H": "90px",  "I": "100px",
        "J": "70px",  "K": "80px",  "L": "90px",  "M": "90px",
        "N": "90px",  "O": "60px",  "P": "70px",  "Q": "90px",  "R": "70px",
        "S": "90px",  "T": "90px",
        "U": "70px",  "V": "70px",  "W": "70px",  "X": "70px",  "Y": "60px",
        "AH": "80px", "AI": "80px", "AJ": "100px",
        "AK": "80px", "AL": "80px", "AM": "110px",
        "AN": "70px", "AO": "70px", "AP": "90px",
        "AQ": "90px", "AR": "90px", "AS": "90px",
        "AZ": "90px",
      },
      overview_pct: {
        _default: "4%",
        "C": "5%",   "D": "5%",   "E": "20%",   "F": "4%",
        "G": "5%",   "H": "6%",   "I": "7%",
        "J": "4%",   "K": "6%",   "L": "5%",   "M": "5%",
        "N": "7%",   "O": "4%",   "P": "4%",   "Q": "5%",   "R": "4%",
        "S": "8%",   "T": "5%",
        "U": "4%",   "V": "4%",   "W": "4%",   "X": "4%",   "Y": "3%",
        "AG": "8%","AH": "5%",  "AI": "5%",  "AJ": "7%",
        "AK": "5%",  "AL": "5%",  "AM": "7%",
        "AN": "4%",  "AO": "4%",  "AP": "5%",
        "AQ": "5%",  "AR": "5%", "AS": "8%",
        "AZ": "6%",
      },
      team: {
        _default: "9.44%",
        "E": "15%",
        "C": "9.44%", "D": "9.44%",
        "Z": "9.44%", "AA": "9.44%", "AB": "9.44%", "AC": "9.44%",
        "AD": "9.44%","AE": "9.44%", "AF": "9.44%",
      },
      milestone: {
        _default: "80px",   // 46 milestone cols × 80px = 3680px → horizontal scroll
        "E": "220px",       // project-id col gets a bit more room
      },
    },

    // ── Editable columns ──────────────────────────────────────
    // Mirrors the same pattern as SW: head assigns head/tl,
    // tl (pm) can edit their own operational fields.
    editableCols: {
      "C":  { type: "tl_dropdown", role: "pm_head" },   // PM Head assignment
      "D":  { type: "tl_dropdown", role: "pm"      },   // PM Name assignment
      "S":  { type: "date",        role: "pm"       },   // Planned Dispatch Date
      "AQ": { type: "dropdown",    role: "pm",  options: ["", "RUN", "HOLD", "CLOSED"] },
      "AR": { type: "date",        role: "pm"       },   // Hold Review Date
      "AS": { type: "text",        role: "pm"       },   // Project Remarks
    },

    // ── Milestone columns ─────────────────────────────────────
    // BB to CU — 46 task milestone columns
    milestoneCols: [
      "BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK",
      "BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU",
      "BV","BW","BX","BY","BZ","CA","CB","CC","CD","CE",
      "CF","CG","CH","CI","CJ","CK","CL","CM","CN","CO",
      "CP","CQ","CR","CS","CT","CU"
    ],

    // ── Sum row columns ───────────────────────────────────────
    // Financial / effort totals shown in overview header sum row
    sumCols: ["J", "K", "Q", "U", "V", "W", "X", "AH", "AK"],

    // ── Cell format rules ─────────────────────────────────────
    formatRules: [
      { cols: ["K", "AH", "AK", "AN", "AO"], format: "decimal2"         },
      { cols: ["O", "P", "AI", "AL", "AP"],  format: "pct_from_decimal" },
    ],

    // ── Cell background color rules ───────────────────────────
    colorRules: {
      dark: [
        { cols: ["G","H","I","J","K"],              bg: "#2a2a2a" },
        { cols: ["N","O","P","Q"],                  bg: "#152030" },
        { cols: ["M","L","R"],                      bg: "#1f2a2a" },
      ],
      light: [
        { cols: ["M","K","U","V","W","X","Y","F","R","O","P","Q","S","T",
                 "AH","AI","AJ","AK","AL","AM","AN","AO","AP",
                 "Z","AA","AB","AC","AD","AE","AF","AG","AZ"],  bg: "#747070" },
        { cols: ["E","N","I","J","C","D","AS","AQ","AR"],       bg: "#90b4df" },
        { cols: ["G"],                                           bg: "#191616" },
      ],
    },

    // ── Cell text color rules ─────────────────────────────────
    textColorRules: {
      dark: [
        { cols: ["Q"],  color: "#ff9999" },
        { cols: ["P"],  color: "#66ccff" },
        { cols: ["K"],  color: "#ffcc66" },
      ],
      light: [
        { cols: ["F","G","R","O","P","Q","S","T","U","V","W","X","Y",
                 "Z","AA","AB","AC","AD","AE","AF","AG",
                 "AH","AI","AJ","AK","AL","AM","AN","AO","AP","AZ"],  color: "#ffba26" },
        { cols: ["K","I"],                                              color: "#ffffff" },
      ],
    },

  }, // end PM

};


// ── Role → department lookup ──────────────────────────────────
// Used by app.js to pick the right config before calling renderMonitor.
var ROLE_TO_MONITOR_DEPT = {
  "sw_tl":    "SW",
  "sw_head":  "SW",
  "hw_tl":    "HW",
  "hw_head":  "HW",
  "mfg_tl":   "MFG",
  "mfg_head": "MFG",
  "pm":       "PM",
  "pm_head":  "PM",
  "admin":    "SW",   // admin defaults to SW; change if needed
  "head":     "SW",
};

// Helper called by app.js:
//   var cfg = getMonitorConfig(state.user.role);
//   renderMonitor(container, data, user, overdueMap, cfg);
function getMonitorConfig(role) {
  var dept = ROLE_TO_MONITOR_DEPT[role] || "SW";
  return MONITOR_CONFIGS[dept] || MONITOR_CONFIGS["SW"];
}
