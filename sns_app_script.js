/**
 * SNS's Study Plan — Google Apps Script
 * ──────────────────────────────────────
 * After importing the .xlsx into Google Sheets:
 *   Extensions → Apps Script → paste this entire file → Save → Run setup()
 */

// ── Sheet name constants ──────────────────────────────────────────────────────
const SH = {
    CONFIG: "⚙️ Config",
    LP: "🛤️ Learning Paths",
    SL: "📚 Study Log",
    TC: "📋 Tasks by Category",
    SCH: "📅 Task Scheduler",
    PR: "📊 Progress & Charts",
    RL: "🔗 Resource Library",
    GUIDE: "❓ How To Use",
};

// ── Row/col constants for Study Log ──────────────────────────────────────────
const SL_HDR = 3;   // header row
const SL_START = 4;   // first data row
const SL_END = 203;
const SL_COL = {
    NUM: 1, PATH: 2, TOPIC: 3, TAG: 4, START: 5, INIT: 6,
    COMP: 7, STATUS: 8, RES: 9, RTYPE: 10,
    D4: 11, D4DONE: 12, D7: 13, D7DONE: 14, CONF: 15, REMARKS: 16
};

// ── Row/col constants for Learning Paths ─────────────────────────────────────
const LP_HDR = 3;
const LP_START = 4;
const LP_END = 103;
const LP_COL = {
    NUM: 1, NAME: 2, DESC: 3, PLAT: 4, TOTAL: 5, COMP: 6,
    PCT: 7, LEFTOFF: 8, LASTDATE: 9, STATUS: 10, NOTES: 11
};
const LP_ARCHIVE_START = 108; // first row available for archive blocks

// ── Color palette (hex without #) ────────────────────────────────────────────
const CLR = {
    HDR1: "#0D47A1", HDR2: "#1B5E20", HDR3: "#4A148C",
    HDR4: "#006064", HDR5: "#37474F",
    WHITE: "#FFFFFF", LGREY: "#F5F5F5", DGREY: "#455A64",
    LT_BLU: "#E3F2FD", LT_GRN: "#C8E6C9", LT_RED: "#FFCDD2",
    LT_ORG: "#FFE0B2", MED_RED: "#EF9A9A",
    INIT: "#BBDEFB", PROG: "#FFFFF9C4".slice(2), // strip FF alpha
    COMP: "#E8F5E9", DONE: "#A5D6A7",
    ARCH: "#ECEFF1", ARCH_HDR: "#546E7A",
    YELLOW: "#FFFFF9C4".slice(2),
};

// ═════════════════════════════════════════════════════════════════════════════
// SETUP — run once after import
// ═════════════════════════════════════════════════════════════════════════════
function setup() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Remove old triggers to avoid duplicates
    ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

    // Install onEdit trigger
    ScriptApp.newTrigger("onEditDispatcher")
        .forSpreadsheet(ss)
        .onEdit()
        .create();

    // Create named range: LP_Names (Learning Path names list)
    _refreshLPNamedRange(ss);

    // Add custom menu
    _addMenu();

    SpreadsheetApp.getUi().alert(
        "✅ Setup complete!\n\n" +
        "• onEdit trigger installed\n" +
        "• Named range LP_Names created\n" +
        "• Custom menu added (SNS Study Plan)\n\n" +
        "You're ready to go!"
    );
}

function _addMenu() {
    SpreadsheetApp.getUi()
        .createMenu("📚 SNS Study Plan")
        .addItem("🔄 Refresh Tasks by Category", "rebuildCategoryView")
        .addItem("📊 Rebuild Progress Charts", "rebuildProgressCharts")
        .addItem("🔃 Sort Study Log", "sortStudyLog")
        .addItem("📦 Archive Completed Paths", "archiveAllCompletedPaths")
        .addSeparator()
        .addItem("🔁 Refresh LP Named Range", "_refreshLPNamedRangeMenu")
        .addItem("⚙️ Re-run Setup", "setup")
        .addToUi();
}

function _refreshLPNamedRangeMenu() {
    _refreshLPNamedRange(SpreadsheetApp.getActiveSpreadsheet());
    SpreadsheetApp.getUi().alert("✅ LP_Names named range refreshed.");
}

function _refreshLPNamedRange(ss) {
    const lpSheet = ss.getSheetByName(SH.LP);
    const range = lpSheet.getRange(LP_START, LP_COL.NAME, LP_END - LP_START + 1, 1);
    // Remove existing named range if any
    const existing = ss.getNamedRanges().find(nr => nr.getName() === "LP_Names");
    if (existing) existing.remove();
    ss.setNamedRange("LP_Names", range);
}

// ═════════════════════════════════════════════════════════════════════════════
// ON-EDIT DISPATCHER
// ═════════════════════════════════════════════════════════════════════════════
function onEditDispatcher(e) {
    if (!e) return;
    const sheet = e.range.getSheet();
    const name = sheet.getName();
    const row = e.range.getRow();
    const col = e.range.getColumn();

    if (name === SH.SL) {
        _handleStudyLogEdit(sheet, row, col, e.value);
    } else if (name === SH.LP) {
        _handleLPEdit(sheet, row, col, e.value);
    } else if (name === SH.SCH) {
        _handleSchedulerEdit(sheet, row, col);
    }
}

// ═════════════════════════════════════════════════════════════════════════════
// STUDY LOG HANDLERS
// ═════════════════════════════════════════════════════════════════════════════
function _handleStudyLogEdit(sheet, row, col, value) {
    if (row < SL_START || row > SL_END) return;

    // Date columns — show picker dialog
    if ([SL_COL.START, SL_COL.INIT, SL_COL.COMP].includes(col)) {
        _showDatePickerDialog(sheet, row, col);
    }

    // Auto-set status when dates change
    if ([SL_COL.START, SL_COL.INIT, SL_COL.COMP].includes(col)) {
        _autoSetStatus(sheet, row);
    }

    // Update LP stats when task status or path changes
    if ([SL_COL.STATUS, SL_COL.PATH, SL_COL.COMP].includes(col)) {
        _updateAllLPStats();
    }

    // Sort after status or path change (debounced via lock)
    if ([SL_COL.STATUS, SL_COL.PATH].includes(col)) {
        _debouncedSort();
    }

    // Auto-resize row
    _autoResizeRow(sheet, row);

    // Auto-number
    _autoNumberSheet(sheet, SL_START, SL_END, SL_COL.NUM, SL_COL.TOPIC);
}

// ── Date picker dialog ────────────────────────────────────────────────────────
function _showDatePickerDialog(sheet, row, col) {
    const labels = {
        [SL_COL.START]: "Start Date",
        [SL_COL.INIT]: "Initiated Date",
        [SL_COL.COMP]: "Completed Date",
    };
    const existing = sheet.getRange(row, col).getValue();
    const defVal = existing instanceof Date
        ? Utilities.formatDate(existing, Session.getScriptTimeZone(), "dd/MM/yyyy")
        : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

    const html = HtmlService.createHtmlOutput(`
    <html><body style="font-family:Arial;padding:16px;background:#F5F5F5;">
    <h3 style="color:#0D47A1;margin-top:0">📅 ${labels[col]}</h3>
    <input type="date" id="dp" style="font-size:15px;padding:6px;border-radius:4px;
      border:1px solid #90A4AE;width:100%"
      value="${_toInputDate(defVal)}">
    <br><br>
    <button onclick="submit()" style="background:#0D47A1;color:#fff;border:none;
      padding:8px 20px;border-radius:4px;font-size:14px;cursor:pointer">
      ✅ Set Date
    </button>
    &nbsp;
    <button onclick="google.script.host.close()" style="background:#90A4AE;
      color:#fff;border:none;padding:8px 14px;border-radius:4px;font-size:14px;
      cursor:pointer">Cancel</button>
    <script>
      function submit(){
        var v=document.getElementById('dp').value;
        if(v) google.script.run.withSuccessHandler(()=>google.script.host.close())
          .setDateValue('${sheet.getName()}',${row},${col},v);
      }
    </script>
    </body></html>`)
        .setWidth(300).setHeight(160);
    SpreadsheetApp.getUi().showModalDialog(html, "Select Date");
}

function setDateValue(sheetName, row, col, isoDate) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    const parts = isoDate.split("-");
    const d = new Date(+parts[0], +parts[1] - 1, +parts[2]);
    const cell = sheet.getRange(row, col);
    cell.setValue(d);
    cell.setNumberFormat("DD-MMM-YY");
    // Re-trigger logic
    _autoSetStatus(sheet, row);
    _updateAllLPStats();
    _debouncedSort();
    _autoResizeRow(sheet, row);
}

function _toInputDate(ddmmyyyy) {
    const [d, m, y] = ddmmyyyy.split("/");
    return `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
}

// ── Auto-set status ───────────────────────────────────────────────────────────
function _autoSetStatus(sheet, row) {
    const startVal = sheet.getRange(row, SL_COL.START).getValue();
    const initVal = sheet.getRange(row, SL_COL.INIT).getValue();
    const compVal = sheet.getRange(row, SL_COL.COMP).getValue();
    const cell = sheet.getRange(row, SL_COL.STATUS);

    // Don't override manual On Hold
    if (cell.getValue() === "On Hold") return;

    if (compVal instanceof Date && compVal != "") {
        cell.setValue("Completed");
    } else if (initVal instanceof Date && initVal != "") {
        cell.setValue("In Progress");
    } else if (startVal instanceof Date && startVal != "") {
        cell.setValue("Initiated");
    } else {
        cell.setValue("Not Started");
    }
}

// ── Auto-resize row ───────────────────────────────────────────────────────────
function _autoResizeRow(sheet, row) {
    sheet.setRowHeight(row, 10);  // collapse first so it snaps to content
    sheet.autoResizeRow(row);
    if (sheet.getRowHeight(row) < 24) sheet.setRowHeight(row, 24);
}

// ── Debounced sort ────────────────────────────────────────────────────────────
function _debouncedSort() {
    const cache = CacheService.getScriptCache();
    if (cache.get("sortPending")) return;
    cache.put("sortPending", "1", 5); // 5-second debounce
    sortStudyLog();
}

// ═════════════════════════════════════════════════════════════════════════════
// SORT STUDY LOG
// Primary:   Status != Completed → top | Completed → bottom
// Secondary: Learning Path A→Z
// Tertiary:  Start Date oldest→newest
// ═════════════════════════════════════════════════════════════════════════════
function sortStudyLog() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SH.SL);
    const lastR = _lastDataRow(sheet, SL_START, SL_COL.TOPIC);
    if (lastR < SL_START) return;

    const range = sheet.getRange(SL_START, 1, lastR - SL_START + 1, 16);

    // Add a helper sort-key in a temp col (col 17) — not displayed
    const vals = range.getValues();
    vals.forEach(row => {
        const status = row[SL_COL.STATUS - 1];
        row.push(status === "Completed" ? 1 : 0); // 0 = active first
    });

    // Sort in JS: by completed-flag → path → start date
    vals.sort((a, b) => {
        if (a[16] !== b[16]) return a[16] - b[16];
        const pathA = (a[SL_COL.PATH - 1] || "").toString().toLowerCase();
        const pathB = (b[SL_COL.PATH - 1] || "").toString().toLowerCase();
        if (pathA !== pathB) return pathA < pathB ? -1 : 1;
        const dA = a[SL_COL.START - 1] instanceof Date ? a[SL_COL.START - 1] : new Date(0);
        const dB = b[SL_COL.START - 1] instanceof Date ? b[SL_COL.START - 1] : new Date(0);
        return dA - dB;
    });

    // Write back without the helper col
    const cleaned = vals.map(r => r.slice(0, 16));
    range.setValues(cleaned);

    // Re-number
    _autoNumberSheet(sheet, SL_START, lastR, SL_COL.NUM, SL_COL.TOPIC);
}

// ═════════════════════════════════════════════════════════════════════════════
// LEARNING PATHS HANDLERS
// ═════════════════════════════════════════════════════════════════════════════
function _handleLPEdit(sheet, row, col, value) {
    if (row < LP_START || row > LP_END) return;

    // Path name changed → refresh named range
    if (col === LP_COL.NAME) {
        _refreshLPNamedRange(SpreadsheetApp.getActiveSpreadsheet());
        _autoResizeRow(sheet, row);
    }

    // Status set to Completed → archive tasks
    if (col === LP_COL.STATUS && value === "Completed") {
        const pathName = sheet.getRange(row, LP_COL.NAME).getValue();
        if (pathName) {
            const ui = SpreadsheetApp.getUi();
            const resp = ui.alert(
                "Archive Learning Path?",
                `Archive all tasks for "${pathName}" from Study Log into this sheet?`,
                ui.ButtonSet.YES_NO
            );
            if (resp === ui.Button.YES) archivePathTasks(pathName);
        }
    }
}

// ── Update LP stats (Completed count + Last Studied date) ────────────────────
function _updateAllLPStats() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const lpSheet = ss.getSheetByName(SH.LP);
    const slSheet = ss.getSheetByName(SH.SL);

    const lastLP = _lastDataRow(lpSheet, LP_START, LP_COL.NAME);
    if (lastLP < LP_START) return;

    const slData = slSheet.getRange(SL_START, 1, SL_END - SL_START + 1, 16).getValues();
    const lpNames = lpSheet.getRange(LP_START, LP_COL.NAME, lastLP - LP_START + 1, 1).getValues();

    lpNames.forEach(([name], i) => {
        if (!name) return;
        const lpRow = LP_START + i;

        // Count completed tasks for this path
        const completed = slData.filter(r =>
            r[SL_COL.PATH - 1] === name && r[SL_COL.STATUS - 1] === "Completed"
        ).length;
        lpSheet.getRange(lpRow, LP_COL.COMP).setValue(completed || "");

        // Max completed date for this path
        const dates = slData
            .filter(r => r[SL_COL.PATH - 1] === name && r[SL_COL.COMP - 1] instanceof Date)
            .map(r => r[SL_COL.COMP - 1]);
        if (dates.length) {
            const maxDate = new Date(Math.max(...dates.map(d => d.getTime())));
            const cell = lpSheet.getRange(lpRow, LP_COL.LASTDATE);
            cell.setValue(maxDate);
            cell.setNumberFormat("DD-MMM-YY");
        }
    });
}

// ═════════════════════════════════════════════════════════════════════════════
// ARCHIVE PATH TASKS
// Moves tasks from Study Log → collapsible group in Learning Paths sheet
// ═════════════════════════════════════════════════════════════════════════════
function archivePathTasks(pathName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const slSheet = ss.getSheetByName(SH.SL);
    const lpSheet = ss.getSheetByName(SH.LP);

    // Gather matching rows from Study Log
    const lastSL = _lastDataRow(slSheet, SL_START, SL_COL.TOPIC);
    if (lastSL < SL_START) return;
    const slData = slSheet.getRange(SL_START, 1, lastSL - SL_START + 1, 16).getValues();
    const taskRows = [];
    const delRows = [];

    slData.forEach((row, i) => {
        if (row[SL_COL.PATH - 1] === pathName) {
            taskRows.push(row);
            delRows.push(SL_START + i);
        }
    });

    if (!taskRows.length) {
        SpreadsheetApp.getUi().alert(`No tasks found for path "${pathName}".`);
        return;
    }

    // Find next archive insert position in LP sheet
    const archStart = _nextArchiveRow(lpSheet);

    // ── Write archive block ───────────────────────────────────────────────────
    // Header row for the archive block
    lpSheet.insertRowsBefore(archStart, taskRows.length + 3);

    const archHdrRow = archStart;
    const archHdrRange = lpSheet.getRange(archHdrRow, 1, 1, 11);
    archHdrRange.merge();
    const archHdrCell = lpSheet.getRange(archHdrRow, 1);
    archHdrCell.setValue(
        `📦  ARCHIVED: ${pathName}  —  ${taskRows.length} tasks  `
        + `(Completed ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "DD-MMM-yyyy")})`
    );
    archHdrCell.setBackground("#546E7A");
    archHdrCell.setFontColor("#FFFFFF");
    archHdrCell.setFontWeight("bold");
    archHdrCell.setFontSize(10);
    archHdrCell.setFontFamily("Arial");
    lpSheet.setRowHeight(archHdrRow, 26);

    // Column sub-header for archive
    const archColHdrs = ["#", "Topic / Task", "Tag", "Start Date", "Initiated",
        "Completed", "Status", "Resource", "Res. Type",
        "Day-4 ✓", "Day-7 ✓", "Confidence", "Remarks"];
    const archSubHdr = archHdrRow + 1;
    archColHdrs.forEach((h, c) => {
        const cell = lpSheet.getRange(archSubHdr, c + 1);
        cell.setValue(h);
        cell.setBackground("#78909C");
        cell.setFontColor("#FFFFFF");
        cell.setFontWeight("bold");
        cell.setFontSize(9);
        cell.setFontFamily("Arial");
        cell.setHorizontalAlignment("center");
    });
    lpSheet.setRowHeight(archSubHdr, 22);

    // Write task rows
    taskRows.forEach((row, i) => {
        const wr = archSubHdr + 1 + i;
        const bg = i % 2 === 0 ? "#ECEFF1" : "#F5F5F5";
        const cols = [
            i + 1,
            row[SL_COL.TOPIC - 1], row[SL_COL.TAG - 1],
            row[SL_COL.START - 1], row[SL_COL.INIT - 1],
            row[SL_COL.COMP - 1], row[SL_COL.STATUS - 1],
            row[SL_COL.RES - 1], row[SL_COL.RTYPE - 1],
            row[SL_COL.D4DONE - 1], row[SL_COL.D7DONE - 1],
            row[SL_COL.CONF - 1], row[SL_COL.REMARKS - 1],
        ];
        cols.forEach((val, c) => {
            const cell = lpSheet.getRange(wr, c + 1);
            cell.setValue(val);
            cell.setBackground(bg);
            cell.setFontSize(9);
            cell.setFontFamily("Arial");
            if (val instanceof Date) cell.setNumberFormat("DD-MMM-YY");
            if (c === 1) cell.setHorizontalAlignment("left");
        });
        lpSheet.setRowHeight(wr, 20);
    });

    // Spacer row after block
    const spacerRow = archSubHdr + taskRows.length + 1;
    lpSheet.setRowHeight(spacerRow, 8);

    // Group the archive rows so they can be collapsed
    lpSheet.rowGroups.shift; // clear if needed
    try {
        lpSheet.getRange(archHdrRow, 1, taskRows.length + 2, 1)
            .shiftRowGroupDepth(1);
    } catch (e) { /* row grouping may not work via script on all versions */ }

    // ── Delete tasks from Study Log (reverse order) ───────────────────────────
    delRows.reverse().forEach(r => slSheet.deleteRow(r));

    // Re-number and re-sort Study Log
    const newLast = _lastDataRow(slSheet, SL_START, SL_COL.TOPIC);
    if (newLast >= SL_START) {
        _autoNumberSheet(slSheet, SL_START, newLast, SL_COL.NUM, SL_COL.TOPIC);
        sortStudyLog();
    }

    _updateAllLPStats();

    SpreadsheetApp.getUi().alert(
        `✅ Archived ${taskRows.length} tasks from "${pathName}" into Learning Paths sheet.`
    );
}

function archiveAllCompletedPaths() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const lpSheet = ss.getSheetByName(SH.LP);
    const lastLP = _lastDataRow(lpSheet, LP_START, LP_COL.NAME);
    if (lastLP < LP_START) return;

    const data = lpSheet.getRange(LP_START, LP_COL.NAME,
        lastLP - LP_START + 1, 10).getValues();

    data.forEach(row => {
        if (row[LP_COL.STATUS - LP_COL.NAME] === "Completed" || row[9] === "Completed") {
            const name = row[0];
            if (name) archivePathTasks(name);
        }
    });
}

function _nextArchiveRow(lpSheet) {
    const lastRow = lpSheet.getLastRow();
    return Math.max(lastRow + 2, LP_ARCHIVE_START);
}

// ═════════════════════════════════════════════════════════════════════════════
// SCHEDULER HANDLER
// ═════════════════════════════════════════════════════════════════════════════
function _handleSchedulerEdit(sheet, row, col) {
    // Col C = Task dropdown in week panel (rows 23+)
    // Auto-fill Learning Path in col D
    if (col === 3 && row >= 23) {
        const task = sheet.getRange(row, 3).getValue();
        if (task) {
            const path = _getPathForTask(task);
            if (path) sheet.getRange(row, 4).setValue(path);
        }
    }
}

function _getPathForTask(taskName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = ss.getSheetByName(SH.SL)
        .getRange(SL_START, SL_COL.PATH, SL_END - SL_START + 1, 2).getValues();
    const match = data.find(r => r[1] === taskName);
    return match ? match[0] : "";
}

// ── Refresh today panel suggestions ──────────────────────────────────────────
function refreshTodayPanel() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sch = ss.getSheetByName(SH.SCH);
    const sl = ss.getSheetByName(SH.SL);
    const today = new Date(); today.setHours(0, 0, 0, 0);

    const slData = sl.getRange(SL_START, 1, SL_END - SL_START + 1, 16).getValues();
    const due = [];

    slData.forEach(row => {
        const topic = row[SL_COL.TOPIC - 1];
        if (!topic) return;
        const path = row[SL_COL.PATH - 1];
        const d4 = row[SL_COL.D4 - 1];
        const d4done = row[SL_COL.D4DONE - 1];
        const d7 = row[SL_COL.D7 - 1];
        const d7done = row[SL_COL.D7DONE - 1];
        const _sameDay = d => d instanceof Date &&
            new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime() === today.getTime();

        if (_sameDay(d4) && d4done !== "✅ Done")
            due.push([_fmt(today), "Today", topic, path, "Day-4 Revision", "", "", "⏳ Pending", ""]);
        if (_sameDay(d7) && d7done !== "✅ Done")
            due.push([_fmt(today), "Today", topic, path, "Day-7 Revision", "", "", "⏳ Pending", ""]);
    });

    // Clear existing suggestions (rows 8-12)
    sch.getRange(8, 1, 5, 9).clearContent();
    due.slice(0, 5).forEach((row, i) => {
        sch.getRange(8 + i, 1, 1, 9).setValues([row]);
    });
    if (!due.length) {
        sch.getRange(8, 1).setValue("🎉  No revisions due today!");
    }
}

function _fmt(d) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "DD-MMM-YY");
}

// ═════════════════════════════════════════════════════════════════════════════
// REBUILD TASKS BY CATEGORY
// ═════════════════════════════════════════════════════════════════════════════
function rebuildCategoryView() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tcSheet = ss.getSheetByName(SH.TC);
    const slSheet = ss.getSheetByName(SH.SL);
    const lpSheet = ss.getSheetByName(SH.LP);
    const rlSheet = ss.getSheetByName(SH.RL);

    // Clear below row 4
    const lastRow = tcSheet.getLastRow();
    if (lastRow >= 5) tcSheet.getRange(5, 1, lastRow - 4, 9).clearContent()
        .setBackground(null).setBorder(null, null, null, null, null, null);

    const slData = slSheet.getRange(SL_START, 1, SL_END - SL_START + 1, 16).getValues();
    const rlData = rlSheet.getRange(4, 1, 200, 8).getValues();
    const lastLP = _lastDataRow(lpSheet, LP_START, LP_COL.NAME);
    if (lastLP < LP_START) return;

    const lpData = lpSheet.getRange(LP_START, 1, lastLP - LP_START + 1, 11).getValues();

    let writeRow = 5;

    lpData.forEach(lpRow => {
        const pathName = lpRow[LP_COL.NAME - 1];
        const pathStatus = lpRow[LP_COL.STATUS - 1];
        if (!pathName) return;

        const tasks = slData.filter(r => r[SL_COL.PATH - 1] === pathName && r[SL_COL.TOPIC - 1]);
        const total = tasks.length;
        const completed = tasks.filter(r => r[SL_COL.STATUS - 1] === "Completed").length;
        const pct = total ? Math.round(completed / total * 100) : 0;
        const resources = rlData.filter(r => r[4] === pathName && r[1]);

        // ── Path header ─────────────────────────────────────────────────────────
        const pathHdrRange = tcSheet.getRange(writeRow, 1, 1, 9);
        pathHdrRange.merge();
        const pathHdrCell = tcSheet.getRange(writeRow, 1);
        const statusEmoji = {
            "Completed": "✅", "In Progress": "⏳",
            "On Hold": "⏸️", "Not Started": "🔲"
        };
        pathHdrCell.setValue(
            `🛤️  ${pathName}   ${completed}/${total} completed   ${pct}%   `
            + `${statusEmoji[pathStatus] || ""} ${pathStatus || ""}`
        );
        pathHdrCell.setBackground("#1B5E20");
        pathHdrCell.setFontColor("#FFFFFF");
        pathHdrCell.setFontWeight("bold");
        pathHdrCell.setFontSize(10);
        pathHdrCell.setFontFamily("Arial");
        tcSheet.setRowHeight(writeRow, 26);
        writeRow++;

        // ── Column headers ───────────────────────────────────────────────────────
        const colHdrs = ["Task", "Tag", "Start Date", "Status", "D4 ✓", "D7 ✓", "⭐", "Confidence", "Resource Link"];
        colHdrs.forEach((h, c) => {
            const cell = tcSheet.getRange(writeRow, c + 1);
            cell.setValue(h);
            cell.setBackground("#2E7D32");
            cell.setFontColor("#FFFFFF");
            cell.setFontWeight("bold");
            cell.setFontSize(9);
            cell.setFontFamily("Arial");
            cell.setHorizontalAlignment("center");
        });
        tcSheet.setRowHeight(writeRow, 20);
        writeRow++;

        // ── Task rows ────────────────────────────────────────────────────────────
        if (!tasks.length) {
            tcSheet.getRange(writeRow, 1, 1, 9).merge();
            tcSheet.getRange(writeRow, 1).setValue("  — No tasks yet —");
            tcSheet.getRange(writeRow, 1).setFontStyle("italic").setFontColor("#9E9E9E");
            tcSheet.setRowHeight(writeRow, 20);
            writeRow++;
        } else {
            tasks.forEach((t, i) => {
                const bg = t[SL_COL.STATUS - 1] === "Completed" ? "#E8F5E9"
                    : i % 2 === 0 ? "#FFFFFF" : "#F1F8E9";
                const rowData = [
                    t[SL_COL.TOPIC - 1],
                    t[SL_COL.TAG - 1],
                    t[SL_COL.START - 1] instanceof Date
                        ? Utilities.formatDate(t[SL_COL.START - 1], Session.getScriptTimeZone(), "DD-MMM-yy")
                        : "",
                    t[SL_COL.STATUS - 1],
                    t[SL_COL.D4DONE - 1],
                    t[SL_COL.D7DONE - 1],
                    t[SL_COL.CONF - 1],
                    "",   // spare
                    t[SL_COL.RES - 1],
                ];
                rowData.forEach((val, c) => {
                    const cell = tcSheet.getRange(writeRow, c + 1);
                    cell.setValue(val);
                    cell.setBackground(bg);
                    cell.setFontSize(9);
                    cell.setFontFamily("Arial");
                    cell.setHorizontalAlignment(c === 0 || c === 8 ? "left" : "center");
                });
                tcSheet.setRowHeight(writeRow, 20);
                writeRow++;
            });
        }

        // ── Resources for this path ───────────────────────────────────────────────
        if (resources.length) {
            tcSheet.getRange(writeRow, 1, 1, 9).merge();
            tcSheet.getRange(writeRow, 1).setValue(
                `📎  Resources (${resources.length}):  ` +
                resources.map(r => `${r[1]} [${r[2]}]`).join("   |   ")
            );
            tcSheet.getRange(writeRow, 1)
                .setBackground("#F3E5F5").setFontSize(8)
                .setFontColor("#4A148C").setFontFamily("Arial");
            tcSheet.setRowHeight(writeRow, 18);
            writeRow++;
        }

        // Spacer
        tcSheet.setRowHeight(writeRow, 8);
        writeRow++;
    });

    SpreadsheetApp.getActiveSpreadsheet().toast(
        "Tasks by Category rebuilt.", "✅ Done", 3);
}

// ═════════════════════════════════════════════════════════════════════════════
// REBUILD PROGRESS CHARTS
// ═════════════════════════════════════════════════════════════════════════════
function rebuildProgressCharts() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const prSheet = ss.getSheetByName(SH.PR);
    const slSheet = ss.getSheetByName(SH.SL);
    const lpSheet = ss.getSheetByName(SH.LP);

    // Remove all existing charts on this sheet
    prSheet.getCharts().forEach(c => prSheet.removeChart(c));

    const lastLP = _lastDataRow(lpSheet, LP_START, LP_COL.NAME);
    if (lastLP < LP_START) return;

    const slData = slSheet.getRange(SL_START, 1, SL_END - SL_START + 1, 16).getValues();
    const lpData = lpSheet.getRange(LP_START, 1, lastLP - LP_START + 1, 11).getValues();

    // Clear old cards below row 10
    const lastPR = prSheet.getLastRow();
    if (lastPR > 10) prSheet.getRange(11, 1, lastPR - 10, 8).clearContent()
        .setBackground(null).setBorder(null, null, null, null, null, null);

    let writeRow = 11;
    let chartCount = 0;

    lpData.forEach(lpRow => {
        const pathName = lpRow[LP_COL.NAME - 1];
        if (!pathName) return;

        const tasks = slData.filter(r => r[SL_COL.PATH - 1] === pathName && r[SL_COL.TOPIC - 1]);
        const total = tasks.length;
        const completed = tasks.filter(r => r[SL_COL.STATUS - 1] === "Completed").length;
        const inProg = tasks.filter(r => r[SL_COL.STATUS - 1] === "In Progress").length;
        const remaining = total - completed;
        const pct = total ? Math.round(completed / total * 100) : 0;
        const bar = _progressBar(pct);
        const lastDate = lpRow[LP_COL.LASTDATE - 1];
        const leftOff = lpRow[LP_COL.LEFTOFF - 1];
        const status = lpRow[LP_COL.STATUS - 1];

        // ── Card header ──────────────────────────────────────────────────────────
        const cardBg = status === "Completed" ? "#A5D6A7"
            : status === "In Progress" ? "#FFF9C4"
                : status === "On Hold" ? "#FFE0B2" : "#E3F2FD";

        _writeCardRow(prSheet, writeRow, 8,
            `🛤️  ${pathName}`, "#1B5E20", "#FFFFFF", true, 11, 28);
        writeRow++;

        _writeCardRow(prSheet, writeRow, 4,
            `Platform: ${lpRow[LP_COL.PLAT - 1] || "—"}`, cardBg, "#37474F", false, 10, 20);
        _writeCardRow(prSheet, writeRow, 4,
            `Status: ${status || "—"}`, cardBg, "#37474F", true, 10, 20, 5);
        writeRow++;

        _writeCardRow(prSheet, writeRow, 8,
            `Progress: ${bar}  ${completed} / ${total}  (${pct}%)`,
            cardBg, "#1B5E20", true, 10, 22);
        writeRow++;

        _writeCardRow(prSheet, writeRow, 4,
            `Last Studied: ${lastDate instanceof Date ? Utilities.formatDate(lastDate, Session.getScriptTimeZone(), "DD-MMM-yyyy") : "—"}`,
            cardBg, "#37474F", false, 9, 18);
        _writeCardRow(prSheet, writeRow, 4,
            `In Progress: ${inProg}`, cardBg, "#37474F", false, 9, 18, 5);
        writeRow++;

        if (leftOff) {
            _writeCardRow(prSheet, writeRow, 8,
                `📍 Left off at: ${leftOff}`, cardBg, "#0D47A1", false, 9, 18);
            writeRow++;
        }

        // ── Chart data range (write temp data for doughnut) ──────────────────────
        // We write chart source data in cols 7-8 (invisible area)
        const dataRow1 = writeRow;
        prSheet.getRange(dataRow1, 7).setValue("Completed");
        prSheet.getRange(dataRow1, 8).setValue(completed);
        prSheet.getRange(dataRow1 + 1, 7).setValue("Remaining");
        prSheet.getRange(dataRow1 + 1, 8).setValue(Math.max(remaining, 0));
        prSheet.getRange(dataRow1, 7, 2, 2).setFontSize(8).setFontColor("#BDBDBD");

        // Build doughnut chart
        if (total > 0) {
            const chartRange = prSheet.getRange(dataRow1, 7, 2, 2);
            const chart = prSheet.newChart()
                .setChartType(Charts.ChartType.PIE)
                .addRange(chartRange)
                .setOption("title", pathName)
                .setOption("pieHole", 0.5)
                .setOption("colors", ["#2E7D32", "#E0E0E0"])
                .setOption("legend", { position: "bottom" })
                .setOption("chartArea", { width: "80%", height: "80%" })
                .setOption("width", 280)
                .setOption("height", 200)
                .setPosition(writeRow - 4, 1, 0, 0)
                .build();
            prSheet.insertChart(chart);
            chartCount++;
        }

        // Spacer
        prSheet.setRowHeight(writeRow + 2, 10);
        writeRow += 12; // leave room for chart
    });

    ss.toast(`Built ${chartCount} doughnut chart(s).`, "📊 Done", 4);
}

function _writeCardRow(sheet, row, span, text, bg, color, bold, size, height, startCol = 1) {
    const endCol = startCol + span - 1;
    try { sheet.getRange(row, startCol, 1, span).merge(); } catch (e) { }
    const cell = sheet.getRange(row, startCol);
    cell.setValue(text);
    cell.setBackground(bg);
    cell.setFontColor(color);
    cell.setFontWeight(bold ? "bold" : "normal");
    cell.setFontSize(size);
    cell.setFontFamily("Arial");
    cell.setHorizontalAlignment("left");
    cell.setVerticalAlignment("middle");
    cell.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    sheet.setRowHeight(row, height);
}

function _progressBar(pct) {
    const filled = Math.round(pct / 5);
    return "█".repeat(filled) + "░".repeat(20 - filled);
}

// ═════════════════════════════════════════════════════════════════════════════
// UTILITIES
// ═════════════════════════════════════════════════════════════════════════════
function _lastDataRow(sheet, startRow, checkCol) {
    const vals = sheet.getRange(startRow, checkCol,
        sheet.getLastRow() - startRow + 1, 1).getValues();
    let last = startRow - 1;
    vals.forEach((r, i) => { if (r[0] !== "") last = startRow + i; });
    return last;
}

function _autoNumberSheet(sheet, startRow, endRow, numCol, checkCol) {
    const vals = sheet.getRange(startRow, checkCol,
        endRow - startRow + 1, 1).getValues();
    let num = 1;
    vals.forEach((r, i) => {
        if (r[0] !== "") {
            sheet.getRange(startRow + i, numCol).setValue(num++);
        }
    });
}