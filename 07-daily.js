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

  // ── Column widths ────────────────────────────────────────────────────────
  sheet.setFrozenRows(2);
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 140);

  for (let col = 3; col <= layout.lastCol; col++) {
    sheet.setColumnWidth(col, 95);
  }
  // Notes column gets extra width
  sheet.setColumnWidth(layout.progressStartCol + 4, 240);

  SpreadsheetApp.flush();
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

  return {
    sections,
    progressStartCol: startCol,
    lastCol:          startCol + CFG.daily.progressLabels.length - 1
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
  sheet.getRange(CFG.daily.firstDataRow, progressStartCol + 1, rowCount, 1).setDataValidation(project);
  sheet.getRange(CFG.daily.firstDataRow, progressStartCol + 2, rowCount, 1).setDataValidation(actions);
  sheet.getRange(CFG.daily.firstDataRow, progressStartCol + 3, rowCount, 1).setDataValidation(yesNoExempt);
}

// ── Tab color ─────────────────────────────────────────────────────────────────

/** Sets the daily tab color to green. */
function colorDailyTab_(sheet) {
  sheet.setTabColor('#93c47d');
}