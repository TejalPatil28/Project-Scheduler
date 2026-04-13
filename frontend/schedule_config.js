// =============================================================
//  schedule_config.js  —  Department Schedule Sheet Configuration
//
//  Single source of truth for ALL department project schedule
//  sheet rendering in renderExcelMirror (app.js).
//
//  To add a new department: copy an existing block, change the
//  key, and fill in the correct values.
//
//  Consumed by:
//    • app.js → renderExcelMirror uses getScheduleConfig(role)
//
// =============================================================
//
//  FRAMEWORK FIELD REFERENCE
//  ─────────────────────────
//  taskStartRow   {number}  First task data row (inclusive)
//  taskEndRow     {number}  Last task data row (inclusive)
//
//  skipCols       {string[]}  Columns hidden in Excel — omit entirely from render
//
//  editableCols   {object}  col → edit descriptor for task rows
//    type options:
//      "date"      → date-picker input
//      "number"    → numeric input  (use min/max to constrain)
//      "dropdown"  → static options list (options[] required)
//      "text"      → free-text input
//      "readonly"  → show value but render no input widget
//                    (use for formula/computed columns you want
//                     visible but never editable, e.g. AC/AD/AE
//                     in PM, or AL % Compl)
//    If a column is absent from editableCols AND not in skipCols,
//    it renders as plain read-only text (same as "readonly").
//    The "readonly" type is explicit — prefer it for columns that
//    carry meaningful data but must never be touched.
//
//  helpDropdown   {object|null}
//    Exactly one column per dept may be a color-coded help
//    dropdown.  Set to null if the dept has no such column.
//    Fields: col, options[], colorsDark{}, colorsLight{}, fgDark{}
//
//  colWidths      {object}  col → CSS width string (%, px, etc.)
//    _default is the fallback for any col not explicitly listed.
//
//  colGroups      {array}  Collapsible column groups
//    Each: { key, label, cols[] }
//    The last col in each group is the "summary" col shown when
//    the group is collapsed.
//
//  customHeaders  {object}  dark/light → col → { label, bg, fg }
//    Override the cell text shown in the header row per column.
//    If a col is absent, the raw Excel header value is used.
//
//  colorRules     {object}  dark/light → array of bg-color rules
//    Each: { cols[], rows, color }
//    rows: "9+" means all task rows; "all" means every row.
//    First match wins.
//
//  textRules      {array}  Conditional text-color rules
//    Each: { name, col, condition, compareWith?, value?, rows, color }
//    condition options:
//      "not_equal"       → color if cell ≠ compareWith col (same row)
//      "less_than"       → color if cell < compareWith col (same row)
//      "value_equals"    → color if cell === value (string match)
//      "value_less_than" → color if parseFloat(cell) < value
//      "value_zero"      → color if cell is 0, "0", or empty
//    First matching rule wins per cell.
//
//  masterListCols {object}  { head, tl }
//    Key names in the monitor JSON used to filter the sidebar
//    project list by user initials.  Must match the keys your
//    backend serialises from the monitor file.
//
//  leftPanel      {object}  Config for the left info-panel
//    dateRows     {array}   Each: { label, custCol, pmCol? }
//      custCol = Excel col letter for Customer date (col C in file)
//      pmCol   = Excel col letter for PM date (col D in file)
//                omit pmCol if the dept only has one date column
//    extraSections {array}  Additional labelled data blocks below
//                            stake holders
//      Each: { title, rows[] }  rows: { label, col, format? }
//        format: "percent" → multiply by 100, append "%"
//                "currency"→ format as number (future use)
//                omit for raw string display
//
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
    editableCols: {
      "X":  { type: "date"                         },
      "Y":  { type: "date"                         },
      "Z":  { type: "number",  min: 0, max: 100    },
      "AD": { type: "dropdown"                     },  // options from helpDropdown
      "AF": { type: "text"                         },
    },

    // ── Help-required dropdown (AD column) ───────────────────
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
      { key: "planning_dates", label: "Planning Dates", cols: ["J","K","L","M","N","O","P","Q","R","S","T"] },
    ],

    // ── Custom header row definitions ─────────────────────────
    customHeaders: {
      dark: {
        "E":  { label: "Buffer Time",        bg: "#2a2010", fg: "#d4a96a" },
        "F":  { label: "PH",                 bg: "#2a2010", fg: "#d4a96a" },
        "G":  { label: "TFO",                bg: "#2a2010", fg: "#d4a96a" },
        "H":  { label: "Phase ID",           bg: "#2a2010", fg: "#d4a96a" },
        "I":  { label: "Task Description",   bg: "#2a2010", fg: "#d4a96a" },
        "J":  { label: "P1 Start Date",      bg: "#2a2010", fg: "#d4a96a" },
        "K":  { label: "P1 End Date",        bg: "#2a2010", fg: "#d4a96a" },
        "L":  { label: "P2 Start Date",      bg: "#2a2010", fg: "#d4a96a" },
        "M":  { label: "P2 End Date",        bg: "#2a2010", fg: "#d4a96a" },
        "N":  { label: "P3 Start Date",      bg: "#2a2010", fg: "#d4a96a" },
        "O":  { label: "P3 End Date",        bg: "#2a2010", fg: "#d4a96a" },
        "P":  { label: "Org Plan Date",      bg: "#2a2010", fg: "#d4a96a" },
        "Q":  { label: "Org End Date",       bg: "#2a2010", fg: "#d4a96a" },
        "R":  { label: "Ref. Lead Time",     bg: "#2a2010", fg: "#d4a96a" },
        "S":  { label: "Lead Time",          bg: "#2a2010", fg: "#d4a96a" },
        "T":  { label: "Intlk",              bg: "#2a2010", fg: "#d4a96a" },
        "U":  { label: "Effort Days",        bg: "#2a2010", fg: "#d4a96a" },
        "V":  { label: "Cur. Start Date",    bg: "#2a2010", fg: "#d4a96a" },
        "W":  { label: "Cur. End Date",      bg: "#2a2010", fg: "#d4a96a" },
        "X":  { label: "Act. Start Date",    bg: "#2a2010", fg: "#d4a96a" },
        "Y":  { label: "Act. End Date",      bg: "#2a2010", fg: "#d4a96a" },
        "Z":  { label: "% Complete",         bg: "#2a2010", fg: "#d4a96a" },
        "AA": { label: "Exptd % Completion", bg: "#2a2010", fg: "#d4a96a" },
        "AB": { label: "Alert Date 80%",     bg: "#2a2010", fg: "#d4a96a" },
        "AD": { label: "Help Req. from",     bg: "#2a2010", fg: "#d4a96a" },
        "AF": { label: "Remark",             bg: "#2a2010", fg: "#d4a96a" },
      },
      light: {
        "E":  { label: "Buffer Time",        bg: "#fcd5b4", fg: "#000000" },
        "F":  { label: "PH",                 bg: "#fcd5b4", fg: "#000000" },
        "G":  { label: "TFO",                bg: "#fcd5b4", fg: "#000000" },
        "H":  { label: "Phase ID",           bg: "#fcd5b4", fg: "#000000" },
        "I":  { label: "Task Description",   bg: "#fcd5b4", fg: "#000000" },
        "J":  { label: "P1 Start Date",      bg: "#fcd5b4", fg: "#000000" },
        "K":  { label: "P1 End Date",        bg: "#fcd5b4", fg: "#000000" },
        "L":  { label: "P2 Start Date",      bg: "#fcd5b4", fg: "#000000" },
        "M":  { label: "P2 End Date",        bg: "#fcd5b4", fg: "#000000" },
        "N":  { label: "P3 Start Date",      bg: "#fcd5b4", fg: "#000000" },
        "O":  { label: "P3 End Date",        bg: "#fcd5b4", fg: "#000000" },
        "P":  { label: "Org Plan Date",      bg: "#fcd5b4", fg: "#000000" },
        "Q":  { label: "Org End Date",       bg: "#fcd5b4", fg: "#000000" },
        "R":  { label: "Ref. Lead Time",     bg: "#fcd5b4", fg: "#000000" },
        "S":  { label: "Lead Time",          bg: "#fcd5b4", fg: "#000000" },
        "T":  { label: "Intlk",              bg: "#fcd5b4", fg: "#000000" },
        "U":  { label: "Effort Days",        bg: "#fcd5b4", fg: "#000000" },
        "V":  { label: "Cur. Start Date",    bg: "#fcd5b4", fg: "#000000" },
        "W":  { label: "Cur. End Date",      bg: "#fcd5b4", fg: "#000000" },
        "X":  { label: "Act. Start Date",    bg: "#fcd5b4", fg: "#000000" },
        "Y":  { label: "Act. End Date",      bg: "#fcd5b4", fg: "#000000" },
        "Z":  { label: "% Complete",         bg: "#fcd5b4", fg: "#000000" },
        "AA": { label: "Exptd % Completion", bg: "#fcd5b4", fg: "#000000" },
        "AB": { label: "Alert Date 80%",     bg: "#fcd5b4", fg: "#000000" },
        "AD": { label: "Help Req. from",     bg: "#fcd5b4", fg: "#000000" },
        "AF": { label: "Remark",             bg: "#fcd5b4", fg: "#000000" },
      },
    },

    // ── Cell background color rules ───────────────────────────
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
    textRules: [
      {
        name:        "Lead Time mismatch",
        col:         "S",
        condition:   "not_equal",
        compareWith: "R",
        rows:        "9+",
        color:       "var(--red)",
      },
      {
        name:        "Progress behind plan",
        col:         "Z",
        condition:   "less_than",
        compareWith: "AA",
        rows:        "9+",
        color:       "var(--red)",
      },
    ],

    // ── Master list column keys (for filterMasterList) ────────
    masterListCols: {
      head: "C",  // head/admin filter key in monitor JSON
      tl:   "D",  // TL filter key in monitor JSON
    },

    // ── Left panel configuration ──────────────────────────────
    // SW has fewer date rows (no HW/Dispatch dates) and no
    // Customer vs PM date split (single date column C only).
    leftPanel: {
      // Each entry: label shown, custCol = Excel col for the date.
      // pmCol is omitted — SW only has one date column (C).
      dateRows: [
        { label: "Internal KOM", custCol: "C" },
        { label: "SW Input",     custCol: "C" },
        { label: "SW FAT",       custCol: "C" },
        { label: "Install",      custCol: "C" },
        { label: "PreComm.",     custCol: "C" },
        { label: "Comm.",        custCol: "C" },
      ],
      // SW has no extra sections below Stake Holders
      extraSections: [],
    },

  }, // end SW


  // ═══════════════════════════════════════════════════════════
  //  PM  —  Project Management Department
  //
  //  Key differences vs SW:
  //    • AC = Payment Value     → readonly (numeric, formula)
  //    • AD = Payment %         → readonly (numeric, formula)
  //    • AE = Recd Payment      → readonly (numeric, formula)
  //    • AF = Remark            → text input  (same as SW)
  //    • AL = % Compl           → readonly (formula, visible)
  //    • AC/AE are NOT skipped  (they carry real data in PM)
  //    • No helpDropdown        (PM has no color-coded help col)
  //    • Left panel has 12 date rows with Customer + PM columns
  //    • Left panel has extra data section below Stake Holders
  //      (VA/SM estimates + actuals + reason remark)
  // ═══════════════════════════════════════════════════════════
  PM: {

    // ── Task row range ────────────────────────────────────────
    taskStartRow: 9,
    taskEndRow:   55,

    // ── Columns to skip entirely (hidden in Excel) ────────────
    // PM does NOT skip AC or AE — they carry payment data.
    // AG–AK are scope flags only meaningful on row 9 header;
    // they are skipped in the task grid view.
    skipCols: ["AG", "AH", "AI", "AJ", "AK"],

    // ── Editable columns (task rows only) ────────────────────
    // AC, AD, AE are formula-driven read-only payment columns.
    // AL is a formula-driven % completion — also read-only.
    // Only X, Y (actual dates), Z (% complete), AF (remark)
    // are user-editable on PM task rows.
    editableCols: {
      "X":  { type: "date"                      },
      "Y":  { type: "date"                      },
      "Z":  { type: "number", min: 0, max: 100  },
      "AC": { type: "readonly"                  },  // Payment Value
      "AD": { type: "readonly"                  },  // Payment %
      "AE": { type: "readonly"                  },  // Recd Payment
      "AF": { type: "text"                      },  // Remark
      "AL": { type: "readonly"                  },  // % Compl (formula)
    },

    // ── Help-required dropdown ────────────────────────────────
    // PM has no color-coded help dropdown column.
    helpDropdown: null,

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
      "AC": "4.5%",  // Payment Value
      "AD": "3.5%",  // Payment %
      "AE": "3.5%",  // Recd Payment
      "AF": "7%",    // Remark
      "AL": "3%",    // % Compl
    },

    // ── Collapsible column groups ─────────────────────────────
    colGroups: [
      { key: "planning_dates", label: "Planning Dates",  cols: ["J","K","L","M","N","O","P","Q","R"] },
      { key: "payment",        label: "Payment",         cols: ["AC","AD","AE"] },
    ],

    // ── Custom header row definitions ─────────────────────────
    customHeaders: {
      dark: {
        "E":  { label: "Buffer Time",        bg: "#1a2020", fg: "#7ecfc0" },
        "F":  { label: "PH",                 bg: "#1a2020", fg: "#7ecfc0" },
        "G":  { label: "TFO",                bg: "#1a2020", fg: "#7ecfc0" },
        "H":  { label: "Phase ID",           bg: "#1a2020", fg: "#7ecfc0" },
        "I":  { label: "Task Description",   bg: "#1a2020", fg: "#7ecfc0" },
        "J":  { label: "P1 Start Date",      bg: "#1a2020", fg: "#7ecfc0" },
        "K":  { label: "P1 End Date",        bg: "#1a2020", fg: "#7ecfc0" },
        "L":  { label: "P2 Start Date",      bg: "#1a2020", fg: "#7ecfc0" },
        "M":  { label: "P2 End Date",        bg: "#1a2020", fg: "#7ecfc0" },
        "N":  { label: "P3 Start Date",      bg: "#1a2020", fg: "#7ecfc0" },
        "O":  { label: "P3 End Date",        bg: "#1a2020", fg: "#7ecfc0" },
        "P":  { label: "Org Plan Date",      bg: "#1a2020", fg: "#7ecfc0" },
        "Q":  { label: "Org End Date",       bg: "#1a2020", fg: "#7ecfc0" },
        "R":  { label: "Ref. Lead Time",     bg: "#1a2020", fg: "#7ecfc0" },
        "S":  { label: "Lead Time",          bg: "#1a2020", fg: "#7ecfc0" },
        "T":  { label: "Intlk",              bg: "#1a2020", fg: "#7ecfc0" },
        "U":  { label: "Effort Days",        bg: "#1a2020", fg: "#7ecfc0" },
        "V":  { label: "Cur. Start Date",    bg: "#1a2020", fg: "#7ecfc0" },
        "W":  { label: "Cur. End Date",      bg: "#1a2020", fg: "#7ecfc0" },
        "X":  { label: "Act. Start Date",    bg: "#1a2020", fg: "#7ecfc0" },
        "Y":  { label: "Act. End Date",      bg: "#1a2020", fg: "#7ecfc0" },
        "Z":  { label: "% Complete",         bg: "#1a2020", fg: "#7ecfc0" },
        "AA": { label: "Exptd % Completion", bg: "#1a2020", fg: "#7ecfc0" },
        "AB": { label: "Alert Date 80%",     bg: "#1a2020", fg: "#7ecfc0" },
        "AC": { label: "Pymt Value",         bg: "#0e2030", fg: "#60c8f0" },
        "AD": { label: "Pymt %",             bg: "#0e2030", fg: "#60c8f0" },
        "AE": { label: "Recd Pymt",          bg: "#0e2030", fg: "#60c8f0" },
        "AF": { label: "Remark",             bg: "#1a2020", fg: "#7ecfc0" },
        "AL": { label: "% Compl",            bg: "#1a2030", fg: "#7ecfc0" },
      },
      light: {
        "E":  { label: "Buffer Time",        bg: "#fcd5b4", fg: "#000000" },
        "F":  { label: "PH",                 bg: "#fcd5b4", fg: "#000000" },
        "G":  { label: "TFO",                bg: "#fcd5b4", fg: "#000000" },
        "H":  { label: "Phase ID",           bg: "#fcd5b4", fg: "#000000" },
        "I":  { label: "Task Description",   bg: "#fcd5b4", fg: "#000000" },
        "J":  { label: "P1 Start Date",      bg: "#fcd5b4", fg: "#000000" },
        "K":  { label: "P1 End Date",        bg: "#fcd5b4", fg: "#000000" },
        "L":  { label: "P2 Start Date",      bg: "#fcd5b4", fg: "#000000" },
        "M":  { label: "P2 End Date",        bg: "#fcd5b4", fg: "#000000" },
        "N":  { label: "P3 Start Date",      bg: "#fcd5b4", fg: "#000000" },
        "O":  { label: "P3 End Date",        bg: "#fcd5b4", fg: "#000000" },
        "P":  { label: "Org Plan Date",      bg: "#fcd5b4", fg: "#000000" },
        "Q":  { label: "Org End Date",       bg: "#fcd5b4", fg: "#000000" },
        "R":  { label: "Ref. Lead Time",     bg: "#fcd5b4", fg: "#000000" },
        "S":  { label: "Lead Time",          bg: "#fcd5b4", fg: "#000000" },
        "T":  { label: "Intlk",              bg: "#fcd5b4", fg: "#000000" },
        "U":  { label: "Effort Days",        bg: "#fcd5b4", fg: "#000000" },
        "V":  { label: "Cur. Start Date",    bg: "#fcd5b4", fg: "#000000" },
        "W":  { label: "Cur. End Date",      bg: "#fcd5b4", fg: "#000000" },
        "X":  { label: "Act. Start Date",    bg: "#fcd5b4", fg: "#000000" },
        "Y":  { label: "Act. End Date",      bg: "#fcd5b4", fg: "#000000" },
        "Z":  { label: "% Complete",         bg: "#fcd5b4", fg: "#000000" },
        "AA": { label: "Exptd % Completion", bg: "#fcd5b4", fg: "#000000" },
        "AB": { label: "Alert Date 80%",     bg: "#fcd5b4", fg: "#000000" },
        "AC": { label: "Payment Value",         bg: "#fcd5b4", fg: "#000000" },
        "AD": { label: "Payment %",             bg: "#fcd5b4", fg: "#000000" },
        "AE": { label: "Recd Payment",          bg: "#fcd5b4", fg: "#000000" },
        "AF": { label: "Remark",             bg: "#fcd5b4", fg: "#000000" },
        "AL": { label: "% Compl",            bg: "#fcd5b4", fg: "#000000" },
      },
    },

    // ── Cell background color rules ───────────────────────────
    colorRules: {
      dark: [
        { cols: ["E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","AA","AB"], rows: "9+", color: "#2a2a2a" },
        { cols: ["X","Y","Z"], rows: "9+", color: "#152030" },
        { cols: ["AC","AD","AE"],  rows: "9+", color: "#0d1e2e" },  // payment cols — distinct bg
      ],
      light: [
        { cols: ["E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","AA","AB","AC","AE"], rows: "9+", color: "#d9d9d9" },
        { cols: ["X","Y","Z"], rows: "9+", color: "#90b4df" },
        { cols: ["AD"],  rows: "9+", color: "#cce5f5" },  // payment cols — distinct bg
      ],
    },

    // ── Cell text color rules (conditional) ───────────────────
    textRules: [
      {
        name:        "Lead Time mismatch",
        col:         "S",
        condition:   "not_equal",
        compareWith: "R",
        rows:        "9+",
        color:       "var(--red)",
      },
      {
        name:        "Progress behind plan",
        col:         "Z",
        condition:   "less_than",
        compareWith: "AA",
        rows:        "9+",
        color:       "var(--red)",
      },
      {
        // Highlight unpaid / zero payment rows in the Payment % col
        name:        "Payment not yet received",
        col:         "AE",
        condition:   "value_zero",
        rows:        "9+",
        color:       "var(--text3)",  // dim — not an error, just unpaid
      },
    ],

    // ── Master list column keys (for filterMasterList) ────────
    // These must match the keys your backend serialises from the
    // PM monitor file (PLMon sheet → DEPT_CONFIG["PM"]).
    masterListCols: {
      head: "C",  // PM Head initials column in monitor JSON
      tl:   "D",   // PM Name (TL) column in monitor JSON
    },

    // ── Left panel configuration ──────────────────────────────
    // PM has the full 12-row date grid with both Customer Dates
    // (col C) and PM Dates (col D).  It also has an extra data
    // section below Stake Holders for VA/SM actuals.
    leftPanel: {
      // custCol = Excel column for Customer Dates (col C in file)
      // pmCol   = Excel column for PM Dates       (col D in file)
      dateRows: [
        { label: "PO Date",    custCol: "C", pmCol: "D" },
        { label: "OPF Recpt",  custCol: "C", pmCol: "D" },
        { label: "HW Input",   custCol: "C", pmCol: "D" },
        { label: "Dwg. Sub.",  custCol: "C", pmCol: "D" },
        { label: "Dwg Appr",   custCol: "C", pmCol: "D" },
        { label: "HW FAT",     custCol: "C", pmCol: "D" },
        { label: "Dispatch",   custCol: "C", pmCol: "D" },
        { label: "SW Input",   custCol: "C", pmCol: "D" },
        { label: "SW FAT",     custCol: "C", pmCol: "D" },
        { label: "Install",    custCol: "C", pmCol: "D" },
        { label: "PreComm.",   custCol: "C", pmCol: "D" },
        { label: "Comm.",      custCol: "C", pmCol: "D" },
      ],
      // Extra sections below Stake Holders (rows 43–53 in Excel)
      // col references are Excel column letters in the left panel.
      // format: "percent" → multiply raw decimal by 100, append %
      "extra_sections": [
          {
              "title": "Panels & Value",
              "rows": [
                  {"label": "Balance Panels",  "col": "C", "row": 43},
                  {"label": "Panel Disp Act",  "col": "C", "row": 44},
                  {"label": "Estimated VA%",   "col": "C", "format": "percent", "row": 45},
                  {"label": "Estimated VA",    "col": "C", "row": 46},
                  {"label": "Estimated SM%",   "col": "C", "format": "percent", "row": 47},
                  {"label": "Estimated SM",    "col": "C", "row": 48},
                  {"label": "Actual VA%",      "col": "C", "format": "percent", "row": 49},
                  {"label": "Actual VA",       "col": "C", "row": 50},
                  {"label": "Actual SM%",      "col": "C", "format": "percent", "row": 51},
                  {"label": "Actual SM",       "col": "C", "row": 52},
                  {"label": "Reason / Remark", "col": "A", "row": 53},
              ],
          },
      ],
    },

  }, // end PM


  // ═══════════════════════════════════════════════════════════
  //  HW  —  Hardware Engineering Department  (placeholder)
  // ═══════════════════════════════════════════════════════════
  HW: null,   // TODO: fill in when HW schedule Excel is available


  // ═══════════════════════════════════════════════════════════
  //  MFG  —  Manufacturing Department  (placeholder)
  // ═══════════════════════════════════════════════════════════
  MFG: null,  // TODO: fill in when MFG schedule Excel is available

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
  "admin":    "SW",   // admin defaults to SW; change if needed
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
// Falls back to SW if the dept config is null (not yet defined).
function getScheduleConfig(role) {
  var dept = ROLE_TO_SCHEDULE_DEPT[role] || "SW";
  return SCHEDULE_CONFIGS[dept] || SCHEDULE_CONFIGS["SW"];
}

// Returns the display label for a role.
function userRoleLabel(role) {
  return ROLE_LABELS[role] || role;
}
