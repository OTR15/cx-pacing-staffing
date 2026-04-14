// ============================================================
// 13-kpi-adjustment.gs
// Handles supervisor goal adjustments in the pacing sheet.
//
// Supervisors fill in col AG (Reason) and col AH (Goal
// Adjustment) then run "Apply Goal Adjustment" from the menu.
// The script updates col AE (Notes) with adjusted targets and
// saves the originals to col AI for reference.
// ============================================================

const KPI_NOTES_COL      = 31; // Column AE (1-based)
const KPI_ADJUSTMENT_COL = 34; // Column AH (1-based)
const KPI_ORIGINAL_COL   = 35; // Column AI (1-based)
const KPI_HEADER_ROW     = 2;
const KPI_DATA_START_ROW = 3;

// ─────────────────────────────────────────────────────────────
// Add "Apply Goal Adjustment" to the existing menu.
// Call addKpiAdjustmentMenu(menu) from your existing onOpen()
// before menu.addToUi() — pass in your menu object.
// ─────────────────────────────────────────────────────────────
function addKpiAdjustmentMenu(menu) {
  menu.addSeparator();
  menu.addItem('⚡ Apply Goal Adjustment', 'applyGoalAdjustment');
}

// ─────────────────────────────────────────────────────────────
// Main function — applies goal adjustment to the selected row
// ─────────────────────────────────────────────────────────────
function applyGoalAdjustment() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const row   = sheet.getActiveRange().getRow();
  const ui    = SpreadsheetApp.getUi();

  // Must be on a daily tab
  if (!_kpiIsDailyTab(sheet.getName())) {
    ui.alert('Please select a row on a daily pacing tab first.');
    return;
  }

  // Must be in a data row
  if (row <= KPI_HEADER_ROW) {
    ui.alert('Please select an agent row, not the header.');
    return;
  }

  const adjustmentRaw = String(sheet.getRange(row, KPI_ADJUSTMENT_COL).getValue()).trim();
  const notesCell     = sheet.getRange(row, KPI_NOTES_COL);
  const notesStr      = String(notesCell.getValue()).trim();
  const agentName     = String(sheet.getRange(row, 1).getValue()).trim();

  // Validate there's something to adjust
  if (!adjustmentRaw) {
    ui.alert('No Goal Adjustment value found in col AH for this row.\nPlease enter a value first.');
    return;
  }
  if (!notesStr) {
    ui.alert('No Notes string found in col AE for this row.');
    return;
  }

  // ── Exempt ────────────────────────────────────────────────
  if (adjustmentRaw.toLowerCase() === 'exempt') {
    _kpiSaveOriginalIfNeeded(sheet, row, notesStr);
    // Replace status and append exempt note
    let updated = notesStr.replace(/Status:\s*\w+/i, 'Status: Exempt');
    if (!updated.includes('Exempt by supervisor')) {
      updated += ' | Note: Exempt by supervisor adjustment';
    }
    notesCell.setValue(updated);
    sheet.getRange(row, KPI_ADJUSTMENT_COL)
      .setNote('Exempt applied by supervisor.\nOriginal targets saved in col AI.');
    ui.alert('✅ ' + agentName + ' marked Exempt for ' + sheet.getName() + '.');
    return;
  }

  // ── Numeric percentage ────────────────────────────────────
  const pct = parseFloat(adjustmentRaw);
  if (isNaN(pct) || pct < 0 || pct > 100) {
    ui.alert(
      'Invalid Goal Adjustment: "' + adjustmentRaw + '".\n' +
      'Please enter a number between 0 and 100, or "Exempt".'
    );
    return;
  }

  // Save originals before any modification
  _kpiSaveOriginalIfNeeded(sheet, row, notesStr);

  // Always calculate from original targets, not current (possibly already adjusted) Notes
  const originalStr = String(sheet.getRange(row, KPI_ORIGINAL_COL).getValue()).trim();
  const sourceStr   = originalStr || notesStr;
  const targets     = _kpiParseTargets(sourceStr);

  if (targets.c == null && targets.r == null) {
    ui.alert('No C: or R: targets found in Notes for this row. No adjustment made.');
    return;
  }

  // Calculate adjusted targets
  const adjC = targets.c != null ? Math.round(targets.c * pct / 100) : null;
  const adjR = targets.r != null ? Math.round(targets.r * pct / 100) : null;

  // Update Notes string
  const updated = _kpiUpdateTargets(notesStr, adjC, adjR);
  notesCell.setValue(updated);

  // Confirm on adjustment cell
  const noteText =
    'Adjusted to ' + pct + '%' +
    (adjC != null ? ' | C: ' + adjC : '') +
    (adjR != null ? ' | R: ' + adjR : '') +
    '\nOriginal saved in col AI.';
  sheet.getRange(row, KPI_ADJUSTMENT_COL).setNote(noteText);

  ui.alert(
    '✅ Goal adjustment applied for ' + agentName + ' on ' + sheet.getName() + '.\n\n' +
    'Adjusted to ' + pct + '%' +
    (adjC != null ? '\nClosed target: ' + adjC : '') +
    (adjR != null ? '\nReplied target: ' + adjR : '') +
    '\n\nOriginal targets saved in col AI.'
  );
}

// ─────────────────────────────────────────────────────────────
// Saves original C: and R: targets to col AI if not already
// saved. Never overwrites once set.
// ─────────────────────────────────────────────────────────────
function _kpiSaveOriginalIfNeeded(sheet, row, notesStr) {
  const originalCell = sheet.getRange(row, KPI_ORIGINAL_COL);
  if (String(originalCell.getValue()).trim()) return; // already saved

  const targets = _kpiParseTargets(notesStr);
  if (targets.c == null && targets.r == null) return;

  const originalStr = 'Original:' +
    (targets.c != null ? ' C:' + targets.c : '') +
    (targets.r != null ? ' R:' + targets.r : '');

  originalCell.setValue(originalStr)
    .setFontColor('#5f6368')
    .setFontStyle('italic');
}

// ─────────────────────────────────────────────────────────────
// Parses C: and R: values from a Notes or Original string
// ─────────────────────────────────────────────────────────────
function _kpiParseTargets(str) {
  const cMatch = str.match(/C:([\d.]+)/);
  const rMatch = str.match(/R:([\d.]+)/);
  return {
    c: cMatch ? parseFloat(cMatch[1]) : null,
    r: rMatch ? parseFloat(rMatch[1]) : null,
  };
}

// ─────────────────────────────────────────────────────────────
// Replaces C: and R: values in a Notes string
// ─────────────────────────────────────────────────────────────
function _kpiUpdateTargets(notesStr, newC, newR) {
  let updated = notesStr;
  if (newC != null) updated = updated.replace(/C:[\d.]+/, 'C:' + newC);
  if (newR != null) updated = updated.replace(/R:[\d.]+/, 'R:' + newR);
  return updated;
}

// ─────────────────────────────────────────────────────────────
// Checks if the sheet name matches the daily tab format M/d/yy
// ─────────────────────────────────────────────────────────────
function _kpiIsDailyTab(sheetName) {
  return /^\d{1,2}\/\d{1,2}\/\d{2}$/.test(sheetName);
}