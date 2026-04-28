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

// =============================================================
// Weekly snapshot re-scorer
// Supervisor workflow:
//   1. Select the agent's row in the KPI Admin Snapshot table
//      on a weekly report tab (e.g. "Week 4/21 - 4/27").
//   2. Enter a Reason from the dropdown in the Reason column.
//   3. Enter a Goal Adj value (optional):
//        "75"  → 75 % of the import goal (reduce by 25 %)
//        "-10" → subtract 10 tickets from both Replied & Closed goals
//        blank → re-score with no goal change (useful for Exempt).
//   4. Run Pacing Report → Re-score Weekly Row.
// =============================================================

function applyWeeklyGoalAdjustment() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui    = SpreadsheetApp.getUi();
  const row   = sheet.getActiveRange().getRow();

  if (!parseWeeklyTabDate_(sheet.getName())) {
    ui.alert('Please select a row on a weekly report tab (e.g. "Week 4/21 - 4/27").');
    return;
  }

  const tableStartRow = CFG.weekly.kpiSnapshot.tableStartRow;
  const dataStartRow  = tableStartRow + WKPI_DATA_OFFSET;

  if (row < dataStartRow) {
    ui.alert('Please select an agent row inside the KPI Admin Snapshot table.');
    return;
  }

  const rowData    = sheet.getRange(row, 1, 1, WKPI_TOTAL_COLS).getValues()[0];
  const agentName  = String(rowData[WKPI_COL_AGENT       - 1] || '').trim();
  const qaScore    = rowData[WKPI_COL_QA_SCORE  - 1];
  const repliedAct = rowData[WKPI_COL_REPLIED   - 1];
  const closedAct  = rowData[WKPI_COL_CLOSED    - 1];
  const csatAct    = rowData[WKPI_COL_CSAT      - 1];
  const reason     = String(rowData[WKPI_COL_REASON    - 1] || '').trim();
  const adjRaw     = String(rowData[WKPI_COL_GOAL_ADJ  - 1] || '').trim();

  if (!agentName) {
    ui.alert('No agent name found in this row — make sure you selected a data row.');
    return;
  }
  if (!reason) {
    ui.alert('Please choose a Reason from the dropdown in col L before re-scoring.');
    return;
  }

  // ── Exempt shortcut ──────────────────────────────────────────
  if (reason === 'Exempt') {
    sheet.getRange(row, WKPI_COL_OVERALL).setValue('');
    sheet.getRange(row, WKPI_COL_STATUS).setValue('Exempt');
    sheet.getRange(row, WKPI_COL_NOTE).setValue('Exempt by supervisor');
    sheet.getRange(row, 1, 1, WKPI_TOTAL_COLS).setBackground('#f1f3f4');
    ui.alert('✅ ' + agentName + ' marked Exempt for this week.');
    return;
  }

  // ── Resolve base goals from cell notes (set by import) ───────
  const repliedGoalCell = sheet.getRange(row, WKPI_COL_REPLIED_GOAL);
  const closedGoalCell  = sheet.getRange(row, WKPI_COL_CLOSED_GOAL);
  const baseReplied = _kpiParseOriginalNote_(repliedGoalCell.getNote())
    ?? Number(rowData[WKPI_COL_REPLIED_GOAL - 1]);
  const baseClosed  = _kpiParseOriginalNote_(closedGoalCell.getNote())
    ?? Number(rowData[WKPI_COL_CLOSED_GOAL  - 1]);

  // ── Apply supervisor adjustment ───────────────────────────────
  let adjReplied = baseReplied;
  let adjClosed  = baseClosed;

  if (adjRaw) {
    const parsed = _kpiParseGoalAdj_(adjRaw, baseReplied, baseClosed);
    if (!parsed) {
      ui.alert(
        'Invalid Goal Adj: "' + adjRaw + '".\n\n' +
        'Enter a percentage (e.g. "75" = 75% of goal)\n' +
        'or a raw reduction (e.g. "-10" = subtract 10 tickets).'
      );
      return;
    }
    adjReplied = parsed.replied;
    adjClosed  = parsed.closed;
  }

  // ── Re-fetch KPI config for weights / thresholds ─────────────
  const cfg           = getWeeklyKpiConfig_();
  const goalQa        = parseFloat(cfg.GOAL_QA)                  || 90;
  const goalCsat      = parseFloat(cfg.GOAL_CSAT)                || 4.9;
  const globalGoalReplied = parseFloat(cfg.GOAL_TICKETS_REPLIED) || 70;
  const wQa    = parseFloat(cfg.WEIGHT_QA)      || 40;
  const wTix   = parseFloat(cfg.WEIGHT_TICKETS) || 20;
  const wClose = parseFloat(cfg.WEIGHT_CLOSED)  || 20;
  const wCsat  = parseFloat(cfg.WEIGHT_CSAT)    || 20;
  const afQa   = parseFloat(cfg.AUTOFAIL_QA_THRESHOLD)       || 74;
  const afTixGlobal = parseFloat(cfg.AUTOFAIL_TICKETS_THRESHOLD) || 40;
  const afTix  = globalGoalReplied > 0
    ? Math.round(afTixGlobal * (adjReplied / globalGoalReplied))
    : afTixGlobal;

  // ── Score calculation ─────────────────────────────────────────
  const toNum = v => (v === '' || v == null) ? null : Number(v);
  const cap   = (v, goal) => v == null || isNaN(v) ? null : Math.min(v / goal, 1.10);

  const qa      = toNum(qaScore);
  const replied = toNum(repliedAct);
  const closed  = toNum(closedAct);
  const csat    = toNum(csatAct);

  const qaRatio      = cap(qa,      goalQa);
  const repliedRatio = cap(replied, adjReplied);
  const closedRatio  = cap(closed,  adjClosed);
  const csatRatio    = cap(csat,    goalCsat);

  let totalWeight = 0, weightedSum = 0;
  [[qaRatio, wQa], [repliedRatio, wTix], [closedRatio, wClose], [csatRatio, wCsat]]
    .forEach(([ratio, weight]) => {
      if (ratio != null) { weightedSum += ratio * weight; totalWeight += weight; }
    });

  let overallPct = totalWeight > 0 ? (weightedSum / totalWeight) * 100 : null;
  const autoFail = (qa != null && qa <= afQa) || (replied != null && replied < afTix);
  if (autoFail) overallPct = 0;

  // ── Status + note ─────────────────────────────────────────────
  let status, note;
  if (autoFail) {
    status = 'AUTO-FAIL';
    const failures = [];
    if (qa != null && qa <= afQa)          failures.push('QA');
    if (replied != null && replied < afTix) failures.push('Tickets Replied');
    note = failures.join(' + ');
  } else if (overallPct == null) {
    status = 'No data'; note = '';
  } else if (overallPct >= 106) {
    status = 'Exceeding'; note = '';
  } else if (overallPct >= 100) {
    status = 'Meeting'; note = '';
  } else if (overallPct >= 90) {
    status = 'Close'; note = '';
  } else {
    status = 'Not Meeting';
    const opps = [
      { label: 'QA',              ratio: qaRatio,      gap: qa      != null ? goalQa      - qa      : null },
      { label: 'Tickets Replied', ratio: repliedRatio, gap: replied != null ? adjReplied  - replied : null },
      { label: 'Closed Tickets',  ratio: closedRatio,  gap: closed  != null ? adjClosed   - closed  : null },
      { label: 'CSAT',            ratio: csatRatio,    gap: csat    != null ? goalCsat    - csat    : null }
    ].filter(o => o.ratio != null);

    if (opps.length) {
      opps.sort((a, b) => a.ratio - b.ratio);
      const top = opps[0];
      const gap = top.label === 'CSAT'
        ? (top.gap != null ? top.gap.toFixed(2) : '')
        : (top.gap != null ? Math.round(top.gap).toString() : '');
      note = 'Focus: ' + top.label + (top.gap != null && top.gap > 0 ? ' (-' + gap + ')' : '');
    } else {
      note = '';
    }
  }

  // ── Row background ────────────────────────────────────────────
  const bgMap = {
    'Exceeding':   '#e6f4ea',
    'Meeting':     '#e8f0fe',
    'Close':       '#fef7e0',
    'Not Meeting': '#fce8e6',
    'AUTO-FAIL':   '#fce8e6',
    'No data':     '#ffffff'
  };
  const bg = bgMap[status] || '#ffffff';

  // ── Write back ────────────────────────────────────────────────
  repliedGoalCell.setValue(adjReplied);
  closedGoalCell.setValue(adjClosed);
  sheet.getRange(row, WKPI_COL_OVERALL).setValue(overallPct);
  sheet.getRange(row, WKPI_COL_STATUS).setValue(status);
  sheet.getRange(row, WKPI_COL_NOTE).setValue(note);
  // Keep Reason + Goal Adj columns their input background.
  sheet.getRange(row, 1, 1, WKPI_COL_NOTE).setBackground(bg);

  ui.alert(
    '✅ Re-scored: ' + agentName + '\n\n' +
    'Reason: ' + reason + '\n' +
    'Replied Goal: ' + adjReplied + '  |  Closed Goal: ' + adjClosed + '\n' +
    'Overall: ' + (overallPct != null ? overallPct.toFixed(1) + '%' : 'N/A') + '\n' +
    'Status: ' + status
  );
}

// ─────────────────────────────────────────────────────────────
// Reads the original goal value stored in a cell note by the
// weekly import (format: "Original: 67").
// Returns null if the note is absent or unparseable.
// ─────────────────────────────────────────────────────────────
function _kpiParseOriginalNote_(note) {
  const m = String(note || '').match(/Original:\s*([\d.]+)/);
  return m ? parseFloat(m[1]) : null;
}

// ─────────────────────────────────────────────────────────────
// Parses a supervisor's Goal Adj entry and returns adjusted
// Replied and Closed goals.
//
//   "75"  → 75 % of base goals
//   "-10" → subtract 10 from both goals (floor 0)
//
// Returns null for anything that doesn't match either format.
// ─────────────────────────────────────────────────────────────
function _kpiParseGoalAdj_(adjRaw, baseReplied, baseClosed) {
  // Raw reduction: leading minus sign
  const negMatch = adjRaw.match(/^-(\d+(?:\.\d+)?)$/);
  if (negMatch) {
    const n = parseFloat(negMatch[1]);
    return {
      replied: Math.max(0, Math.round(baseReplied - n)),
      closed:  Math.max(0, Math.round(baseClosed  - n))
    };
  }
  // Percentage: plain number 0–100
  const pct = parseFloat(adjRaw);
  if (!isNaN(pct) && pct >= 0 && pct <= 100) {
    return {
      replied: Math.round(baseReplied * pct / 100),
      closed:  Math.round(baseClosed  * pct / 100)
    };
  }
  return null;
}