// =============================================================================
// daily.gs
// Creates and structures the daily pacing tab.
//
// buildDailySheet_() lays out the skeleton: headers, checkpoint columns,
// progress columns, rep rows, and dropdown validations.
// Actual metric data is written separately by publish.gs.
// =============================================================================

// ── Entry point ───────────────────────────────────────────────────────────────

/**
 * Creates (or recreates) today's daily tab.
 * Deletes any existing tab with the same name first to ensure a clean rebuild.
 */
function createTodayTab() {
  const ss   = SpreadsheetApp.getActive();
  const name = formatDailySheetName_(new Date());
  const old  = ss.getSheetByName(name);

  if (old) ss.deleteSheet(old);

  const sheet = ss.insertSheet(name);
  colorDailyTab_(sheet);
  buildDailySheet_(sheet, new Date());
  SpreadsheetApp.setActiveSheet(sheet);
}

// ── Sheet builder ─────────────────────────────────────────────────────────────

/**
 * Builds the full layout for a daily pacing sheet.
 * Structure:
 *   Row 1: Date label | checkpoint section headers (merged) | "Progress Tracking"
 *   Row 2: "Rep Name" | "Manager" | metric labels per checkpoint | progress labels
 *   Row 3+: one row per rep
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Date} dateObj
 */
function buildDailySheet_(sheet, dateObj) {
  const layout     = getLayout_();
  const reps       = getDisplayRoster_();
  const managerMap = getManagerMapFromSchedule_();

  // Reset the entire used area before rebuilding
  const maxRows = Math.max(sheet.getMaxRows(),    reps.length + 30);
  const maxCols = Math.max(sheet.getMaxColumns(), layout.lastCol + 5);
  const resetRange = sheet.getRange(1, 1, maxRows, maxCols);

  try { resetRange.breakApart(); } catch (e) {}
  resetRange.clearContent();
  resetRange.clearFormat();
  resetRange.clearDataValidations();

  // ── Row 1: date + section headers ───────────────────────────────────────
  sheet.getRange(1, 1).setValue('Date').setFontWeight('bold');
  sheet.getRange(1, 2).setValue(formatDailySheetName_(dateObj)).setFontWeight('bold');

  layout.sections.forEach(section => {
    sheet.getRange(1, section.startCol, 1, 4).merge();
    sheet.getRange(1, section.startCol)
      .setValue(section.label)
      .setHorizontalAlignment('center')
      .setFontWeight('bold');
  });

  sheet.getRange(1, layout.progressStartCol, 1, CFG.daily.progressLabels.length).merge();
  sheet.getRange(1, layout.progressStartCol)
    .setValue('Progress Tracking')
    .setHorizontalAlignment('center')
    .setFontWeight('bold');

  // ── Goal adjustments block ───────────────────────────────────────────────
  sheet.getRange(1, layout.reviewFlagCol, 1, 3).merge();
  sheet.getRange(1, layout.reviewFlagCol)
    .setValue('Goal Adjustments')
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setBackground('#6d9eeb')
    .setFontColor('#ffffff');

  sheet.getRange(2, layout.reviewFlagCol, 1, 3)
    .setValues([['Status', 'Reason', 'Hours Removed']])
    .setFontWeight('bold')
    .setBackground('#a4c2f4');

  // ── Row 2: column headers ────────────────────────────────────────────────
  sheet.getRange(2, 1).setValue('Rep Name').setFontWeight('bold').setBackground('#00ff66');
  sheet.getRange(2, 2).setValue('Manager').setFontWeight('bold').setBackground('#00ff66');

  layout.sections.forEach(section => {
    sheet.getRange(2, section.startCol, 1, 4)
      .setValues([CFG.daily.metricLabels])
      .setFontWeight('bold')
      .setBackground('#fff176');
  });

  sheet.getRange(2, layout.progressStartCol, 1, CFG.daily.progressLabels.length)
    .setValues([CFG.daily.progressLabels])
    .setFontWeight('bold')
    .setBackground('#00ff66');

  // ── Background shading for data rows ────────────────────────────────────
  const dataRowCount = Math.max(reps.length, 20);

  layout.sections.forEach(section => {
    sheet.getRange(3, section.startCol, dataRowCount, 4).setBackground('#fff9c4');
  });

  sheet.getRange(3, layout.progressStartCol, dataRowCount, CFG.daily.progressLabels.length)
    .setBackground('#ccffcc');

  // ── Rep names and managers ───────────────────────────────────────────────
  if (reps.length) {
    const rows = reps.map(r => [r.repName, getManagerForRep_(managerMap, r.repName)]);
    sheet.getRange(CFG.daily.firstDataRow, 1, rows.length, 2).setValues(rows);
  }

  // ── Dropdowns on progress columns ───────────────────────────────────────
  applyProgressValidation_(sheet, dataRowCount, layout.progressStartCol);
  applyReviewValidation_(sheet, dataRowCount, layout);

  // ── Column widths ────────────────────────────────────────────────────────
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(1);
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 140);

  for (let col = 3; col <= layout.lastCol; col++) {
    sheet.setColumnWidth(col, 95);
  }
  // Notes column gets extra width
  sheet.setColumnWidth(layout.progressStartCol + 4, 240);

  // "On a Project" column is unused — hide it while preserving column indices
  sheet.hideColumns(layout.progressStartCol + 1);

  SpreadsheetApp.flush();
}

/**
 * Builds a normalized rep-name → sheet-row map from a daily tab.
 * This lets publish/update flows keep working even after supervisors
 * sort or filter the sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Object<string, number>}
 */
function getDailySheetRowMap_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CFG.daily.firstDataRow) return {};

  const rowCount = lastRow - CFG.daily.firstDataRow + 1;
  const nameValues = sheet.getRange(CFG.daily.firstDataRow, 1, rowCount, 1).getDisplayValues();
  const rowMap = {};

  nameValues.forEach((row, index) => {
    const repName = String(row[0] || '').trim();
    if (!repName) return;
    rowMap[normalizeName_(repName)] = CFG.daily.firstDataRow + index;
  });

  return rowMap;
}

// ── Layout calculator ─────────────────────────────────────────────────────────

/**
 * Computes the column positions for every checkpoint section and the
 * progress tracking block.
 *
 * Each checkpoint section occupies 4 columns:
 *   closed | replied | messages | csat
 *
 * Columns 1–2 are reserved for Rep Name and Manager.
 *
 * @returns {{
 *   sections: Array<{
 *     key, label, percent,
 *     startCol, closedCol, repliedCol, messagesCol, csatCol
 *   }>,
 *   progressStartCol: number,
 *   lastCol: number
 * }}
 */
function getLayout_() {
  const sections  = [];
  let   startCol  = 3; // columns 1-2 = name + manager

  CFG.checkpoints.forEach(cp => {
    sections.push({
      key:        cp.key,
      label:      cp.label,
      percent:    cp.percent,
      startCol:   startCol,
      closedCol:  startCol,
      repliedCol: startCol + 1,
      messagesCol:startCol + 2,
      csatCol:    startCol + 3
    });
    startCol += 4;
  });

  const progressStartCol = startCol;
  const notesCol         = progressStartCol + CFG.daily.progressLabels.length - 1;
  const reviewFlagCol    = notesCol + 1;
  const reviewReasonCol  = notesCol + 2;
  const reviewAdjustCol  = notesCol + 3;

  return {
    sections,
    progressStartCol,
    lastCol:         reviewAdjustCol,
    notesCol,
    reviewFlagCol,   // "Needs Review" — auto-set at EOD
    reviewReasonCol, // Supervisor dropdown: reason for underperformance
    reviewAdjustCol  // Supervisor dropdown: goal adjustment %
  };
}

// ── Progress column dropdowns ─────────────────────────────────────────────────

/**
 * Applies dropdown data validation to the five progress tracking columns.
 *
 * Col +0  On Track:     Yes / No / Exempt
 * Col +1  On a Project: Yes / No / In Office
 * Col +2  Actions Taken: Coaching / Slack Check-In / 1:1 / Escalation / …
 * Col +3  EOD Goal Met: Yes / No / Exempt
 * Col +4  Notes:        free text (no validation)
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} rowCount          - Number of data rows to apply validation to.
 * @param {number} progressStartCol  - Column index of the first progress column.
 */
function applyProgressValidation_(sheet, rowCount, progressStartCol) {
  const yesNoExempt = SpreadsheetApp.newDataValidation()
    .requireValueInList(['', 'Yes', 'No', 'Exempt'], true)
    .build();

  const project = SpreadsheetApp.newDataValidation()
    .requireValueInList(['', 'Yes', 'No', 'In Office'], true)
    .build();

  const actions = SpreadsheetApp.newDataValidation()
    .requireValueInList(['', 'Coaching', 'Slack Check-In', '1:1', 'Escalation', 'Working Lunch', 'CTO', 'VTO', 'Off'], true)
    .build();

  sheet.getRange(CFG.daily.firstDataRow, progressStartCol,     rowCount, 1).setDataValidation(yesNoExempt);
  sheet.getRange(CFG.daily.firstDataRow, progressStartCol + 2, rowCount, 1).setDataValidation(actions);
  sheet.getRange(CFG.daily.firstDataRow, progressStartCol + 3, rowCount, 1).setDataValidation(yesNoExempt);
}

// ── Tab color ─────────────────────────────────────────────────────────────────

/**
 * Applies validation to the goal adjustment columns.
 * Supervisors can adjust any agent's daily goal by entering hours to remove.
 *
 * Col reviewFlagCol:   written by applyGoalAdjustments() — no validation needed
 * Col reviewReasonCol: CTO / VTO / Unexcused Absence / Project / Performance
 * Col reviewAdjustCol: positive number — hours to subtract from the agent's worked hours
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} rowCount
 * @param {{ reviewReasonCol: number, reviewAdjustCol: number }} layout
 */
function applyReviewValidation_(sheet, rowCount, layout) {
  const reasonValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['', 'CTO', 'VTO', 'Unexcused Absence', 'Project', 'Performance'], true)
    .build();

  const hoursValidation = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 24)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(CFG.daily.firstDataRow, layout.reviewReasonCol, rowCount, 1)
    .setDataValidation(reasonValidation);
  sheet.getRange(CFG.daily.firstDataRow, layout.reviewAdjustCol, rowCount, 1)
    .setDataValidation(hoursValidation);
}

/** Sets the daily tab color to green. */
function colorDailyTab_(sheet) {
  sheet.setTabColor('#93c47d');
}

/**
 * Sorts the active daily tab by Manager, then Rep Name.
 * Safe to run before future publishes because publish now resolves rows
 * by rep name from the sheet itself.
 */
function sortActiveDailySheetByManager() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const dateObj = parseDailySheetName_(sheet.getName());
  if (!dateObj) {
    SpreadsheetApp.getUi().alert('Please open a daily pacing tab first.');
    return;
  }

  const layout = getLayout_();
  const lastRow = sheet.getLastRow();
  if (lastRow < CFG.daily.firstDataRow) return;

  sheet.getRange(
    CFG.daily.firstDataRow,
    1,
    lastRow - CFG.daily.firstDataRow + 1,
    layout.lastCol
  ).sort([
    { column: 2, ascending: true },
    { column: 1, ascending: true }
  ]);

  SpreadsheetApp.getUi().alert('The active daily tab was sorted by manager, then rep name.');
}

/**
 * Prompts for a manager name and filters the active daily tab so only that
 * manager's reps remain visible.
 */
function filterActiveDailySheetByManager() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const dateObj = parseDailySheetName_(sheet.getName());
  if (!dateObj) {
    SpreadsheetApp.getUi().alert('Please open a daily pacing tab first.');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Filter Daily Tab by Manager',
    'Enter the manager name exactly as shown in column B.',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const managerName = String(response.getResponseText() || '').trim();
  if (!managerName) {
    ui.alert('Enter a manager name, or use "Show All Managers" to clear the filter.');
    return;
  }

  const appliedManager = applyManagerFilterToDailySheet_(sheet, managerName);
  if (!appliedManager) {
    ui.alert('No rows matched manager: ' + managerName);
    return;
  }

  ui.alert('The active daily tab is now filtered to manager: ' + appliedManager);
}

/**
 * Clears any manager filter from the active daily tab.
 */
function clearManagerFilterOnActiveDailySheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const dateObj = parseDailySheetName_(sheet.getName());
  if (!dateObj) {
    SpreadsheetApp.getUi().alert('Please open a daily pacing tab first.');
    return;
  }

  const filter = sheet.getFilter();
  if (filter) {
    filter.removeColumnFilterCriteria(2);
  }

  SpreadsheetApp.getUi().alert('Manager filtering was cleared on the active daily tab.');
}

/**
 * Applies a manager filter to a daily sheet using the built-in sheet filter.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} managerName
 * @returns {string} The matched manager name shown on the sheet, or '' if not found.
 */
function applyManagerFilterToDailySheet_(sheet, managerName) {
  const layout = getLayout_();
  const lastRow = Math.max(sheet.getLastRow(), CFG.daily.firstDataRow);
  const rowCount = lastRow - 1; // include header row 2 through the last data row

  let filter = sheet.getFilter();
  if (!filter) {
    sheet.getRange(2, 1, rowCount, layout.lastCol).createFilter();
    filter = sheet.getFilter();
  }

  const managerValues = sheet.getRange(CFG.daily.firstDataRow, 2, Math.max(lastRow - CFG.daily.firstDataRow + 1, 0), 1)
    .getDisplayValues()
    .map(row => String(row[0] || '').trim())
    .filter(Boolean);

  const uniqueManagers = [...new Set(managerValues)];
  const normalizedTarget = normalizeName_(managerName);
  const matchedManager = uniqueManagers.find(name => normalizeName_(name) === normalizedTarget);
  if (!matchedManager) return '';

  const hiddenManagers = uniqueManagers.filter(name => name !== matchedManager);

  const criteria = SpreadsheetApp.newFilterCriteria()
    .setHiddenValues(hiddenManagers)
    .build();

  filter.setColumnFilterCriteria(2, criteria);
  return matchedManager;
}
