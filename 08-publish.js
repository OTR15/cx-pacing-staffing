// =============================================================================
// publish.gs
// Checkpoint publishing — pulls live Gorgias data and writes metrics to the
// daily tab for each rep.
//
// Main entry points:
//   publishCurrentCheckpoint()       — detects current hour and publishes it
//   publishCheckpointForDate_()      — publishes a specific checkpoint on any date
//   publishAllCheckpointsToday()     — publishes all checkpoints (use with care)
// =============================================================================

// ── Public entry points ───────────────────────────────────────────────────────

/**
 * Detects the current checkpoint based on the current hour and publishes it.
 * Throws if no checkpoint matches the current hour.
 */
function publishCurrentCheckpoint() {
  const now        = new Date();
  const checkpoint = getCurrentCheckpoint_(now);
  if (!checkpoint) throw new Error('No matching checkpoint for current hour.');
  publishCheckpointForDate_(now, checkpoint);
}

/**
 * Publishes all checkpoints for today in sequence.
 * Warning: this is API-heavy — use for testing only, not production.
 */
function publishAllCheckpointsToday() {
  const today = new Date();
  CFG.checkpoints.forEach(cp => publishCheckpointForDate_(today, cp));
}

// ── Core publish ──────────────────────────────────────────────────────────────

/**
 * Fetches live Gorgias stats and writes them to the daily tab for every
 * active rep at a given checkpoint.
 *
 * For each rep:
 *   - If off / CTO / VTO: marks the row as Exempt and clears metric cells
 *   - If working: fetches actuals, computes targets, colors cells, writes progress
 *
 * A 700ms sleep between reps keeps Gorgias API calls within rate limits.
 *
 * @param {Date}   dateObj    - The date to publish for.
 * @param {{ key, hour, minute, percent, label }} checkpoint
 */
function publishCheckpointForDate_(dateObj, checkpoint) {
  const ss        = SpreadsheetApp.getActive();
  const sheetName = formatDailySheetName_(dateObj);
  let sheet       = ss.getSheetByName(sheetName);
  if (!sheet) sheet = getOrCreateDailySheetForDate_(dateObj);

  const layout      = getLayout_();
  const section     = layout.sections.find(s => s.key === checkpoint.key);
  if (!section) throw new Error('Section not found for checkpoint ' + checkpoint.key);

  const fullRoster  = getDisplayRoster_();
  const rowMap      = getDailySheetRowMap_(sheet);
  const scheduleMap = getScheduleMapForDate_(dateObj);
  const goalsMap    = getGoalsMap_();
  const range       = getCheckpointIsoRange_(dateObj, checkpoint);

  // Fetch closed-tickets-per-agent once for the whole team (one API call)
  const closedMap = getClosedPerAgentMap_(range.startIso, range.endIso);

  for (let i = 0; i < fullRoster.length; i++) {
    const rep      = fullRoster[i];
    const row      = rowMap[normalizeName_(rep.repName)];
    if (!row) continue;
    const schedule = getScheduleForRep_(scheduleMap, rep.repName);

    // ── Not yet started: rep is scheduled but their shift hasn't begun ────
    // If they have a future start time, show black until they clock in.
    const isNotYetStarted = schedule.hours > 0 &&
      schedule.startText &&
      parseTimeToMinutes_(schedule.startText) > (checkpoint.hour * 60 + (checkpoint.minute || 0));

    // ── Unscheduled: rep exists in roster but has no schedule entry at all ─
    const isUnscheduled = !scheduleMap[normalizeName_(rep.repName)] &&
                          !scheduleMap[normalizeFirstName_(rep.repName)];

    if (isUnscheduled || isNotYetStarted) {
      const startsLabel = isNotYetStarted
        ? 'Not Yet Started | Starts: ' + schedule.startText
        : 'Unscheduled';
      applyStatusBlock_(sheet, row, section, layout, 'unscheduled', startsLabel);
      retroactivelyUpdateEarlierCheckpoints_(sheet, layout, row, checkpoint, 'unscheduled');
      continue;
    }

    // ── Exempt: confirmed Off / full-day CTO / full-day VTO ───────────────
    // Also catches reps in the schedule map with hours: 0 and no startText
    // (blank schedule cell that normalized to Off).
    if (schedule.hours <= 0 || ['CTO', 'VTO', 'Off'].includes(schedule.status)) {
      const colorStatus = schedule.status === 'CTO' ? 'cto'
                        : schedule.status === 'VTO' ? 'vto'
                        : 'exempt';
      applyStatusBlock_(sheet, row, section, layout, colorStatus, schedule.status || 'Off');
      retroactivelyUpdateEarlierCheckpoints_(sheet, layout, row, checkpoint, colorStatus);
      continue;
    }

    // ── Partial CTO/VTO: rep worked part of the day ───────────────────────
    // Determine whether this checkpoint falls after, during, or before
    // the rep's working hours to decide how to color and whether to show metrics.
    const isPartialCto = schedule.status === 'Partial CTO';
    const isPartialVto = schedule.status === 'Partial VTO';
    const isPartial    = isPartialCto || isPartialVto;
    const partialColor = isPartialCto ? 'cto' : isPartialVto ? 'vto' : null;

    if (isPartial) {
      const shiftEndMins = parseTimeToMinutes_(schedule.endText);
      const cpPosition   = getCheckpointPosition_(checkpoint, shiftEndMins, layout);
      // 'after'  = rep already left before this checkpoint window started → Exempt
      // 'during' = rep left mid-window → show metrics in CTO/VTO color
      // 'before' = rep is still working this whole window → normal colors below
      if (cpPosition === 'after') {
        applyStatusBlock_(sheet, row, section, layout, partialColor, schedule.status);
        continue;
      }
      if (cpPosition === 'during') {
        // Fall through to metric fetch below — color override applied after writing
      }
    }

    // ── Active rep: compute goals and actuals ─────────────────────────────
    const goals = getEffectiveGoals_(goalsMap, rep.repName);

    const useShiftGoals =
      String(getConfigValue_('USE_SHIFT_BASED_GOALS', true)).toLowerCase() !== 'false' &&
      goals.useShift !== false;

    const effectiveHours = getEffectiveWorkedHours_(schedule);

    // If a supervisor has entered hours to remove, subtract them before scaling goals
    const hoursRemovedRaw = sheet.getRange(row, layout.reviewAdjustCol).getValue();
    const hoursRemoved    = (hoursRemovedRaw !== '' && hoursRemovedRaw !== null && !isNaN(Number(hoursRemovedRaw)))
      ? Number(hoursRemovedRaw) : 0;
    const billedHours = Math.max(0, effectiveHours - hoursRemoved);

    const adjustedGoals = {
      closedTickets:  useShiftGoals ? adjustGoalByShift_(goals.closedTickets,  billedHours) : goals.closedTickets,
      ticketsReplied: useShiftGoals ? adjustGoalByShift_(goals.ticketsReplied, billedHours) : goals.ticketsReplied,
      messagesSent:   useShiftGoals ? adjustGoalByShift_(goals.messagesSent,   billedHours) : goals.messagesSent,
      csat:           goals.csat
    };

    const csatMetrics = getCsatMetrics_(range.startIso, range.endIso, { agents: [rep.agentId] });

    // Closed tickets come from the team-wide per-agent map (one call above)
    // Other metrics require individual agent calls
    const actual = {
      closedTickets:  closedMap[normalizeName_(rep.repName)] || closedMap[normalizeFirstName_(rep.repName)] || 0,
      ticketsReplied: getStatNumber_('total-tickets-replied', range.startIso, range.endIso, { agents: [rep.agentId] }),
      messagesSent:   getStatNumber_('total-messages-sent',   range.startIso, range.endIso, { agents: [rep.agentId] }),
      csat:           csatMetrics.averageRating
    };

    const targets = {
      closedTickets:  getCheckpointTarget_(adjustedGoals.closedTickets,  checkpoint.percent),
      ticketsReplied: getCheckpointTarget_(adjustedGoals.ticketsReplied, checkpoint.percent),
      messagesSent:   getCheckpointTarget_(adjustedGoals.messagesSent,   checkpoint.percent),
      csat:           adjustedGoals.csat
    };

    // ── Write metrics and colors ──────────────────────────────────────────
    const metricStatuses = getMetricStatuses_(actual, targets);

    // If rep left mid-window (during) or is partial, override block color
    const isDuringPartial = isPartial &&
      getCheckpointPosition_(checkpoint, parseTimeToMinutes_(schedule.endText), layout) === 'during';

    const pacingStatus = isDuringPartial          ? partialColor
                       : getWorstPacingStatus_(actual, targets);
    const blockColor   = isDuringPartial          ? partialColor : null;

    const onTrack      = computeOnTrack_(actual, targets);
    const eodMet  = checkpoint.key === 'EOD' ? computeEodGoalMet_(actual, adjustedGoals) : '';

    const onProject      = schedule.inOffice    ? 'In Office'      : '';
    const actions        = schedule.workingLunch ? 'Working Lunch' : '';

    writeMetrics_(sheet, row, section, actual);

    if (blockColor) {
      // Mid-window departure: paint all 4 metric cells in the partial color
      applyPacingColor_(sheet.getRange(row, section.startCol, 1, 4), blockColor);
    } else {
      applyPacingColor_(sheet.getRange(row, section.closedCol),   metricStatuses.closedTickets);
      applyPacingColor_(sheet.getRange(row, section.repliedCol),  metricStatuses.ticketsReplied);
      applyPacingColor_(sheet.getRange(row, section.messagesCol), metricStatuses.messagesSent);
      applyPacingColor_(sheet.getRange(row, section.csatCol),     metricStatuses.csat);
    }

    sheet.getRange(row, layout.progressStartCol).setValue(onTrack);
    applyPacingColor_(sheet.getRange(row, layout.progressStartCol), pacingStatus);

    sheet.getRange(row, layout.progressStartCol + 1).setValue(onProject);
    sheet.getRange(row, layout.progressStartCol + 2).setValue(actions);
    sheet.getRange(row, layout.progressStartCol + 3).setValue(eodMet);
    sheet.getRange(row, layout.progressStartCol + 4).setValue(
      'Status: '   + schedule.status +
      ' | Scheduled: ' + round2_(schedule.hours) +
      ' | Effective: ' + round2_(effectiveHours) +
      ' | Targets ' + checkpoint.key +
      ' C:'    + round1_(targets.closedTickets) +
      ' R:'    + round1_(targets.ticketsReplied) +
      ' M:'    + round1_(targets.messagesSent) +
      ' CSAT:' + adjustedGoals.csat
    );

    // Brief pause to avoid hitting Gorgias rate limits
    Utilities.sleep(700);
  }

  SpreadsheetApp.flush();
}

// ── Status block helpers ──────────────────────────────────────────────────────

/**
 * Determines where a rep's shift end falls relative to a checkpoint window.
 *
 * Returns:
 *   'before'  — rep's shift ends after this checkpoint hour (still working)
 *   'during'  — rep's shift ends within this checkpoint's window
 *                (between the previous checkpoint hour and this one)
 *   'after'   — rep's shift ended before this checkpoint window started
 *
 * @param {{ key: string, hour: number }} checkpoint
 * @param {number} shiftEndMins - Shift end in minutes from midnight
 * @param {{ sections: Array }} layout
 * @returns {'before'|'during'|'after'}
 */
function getCheckpointPosition_(checkpoint, shiftEndMins, layout) {
  const cpMins = checkpoint.hour * 60 + (checkpoint.minute || 0);

  // Find the previous checkpoint's hour to define the window start
  const cpIndex   = layout.sections.findIndex(s => s.key === checkpoint.key);
  const prevCpMins = cpIndex > 0
    ? CFG.checkpoints[cpIndex - 1].hour * 60 + (CFG.checkpoints[cpIndex - 1].minute || 0)
    : 0;

  if (shiftEndMins >= cpMins)     return 'before';  // still working at checkpoint time
  if (shiftEndMins > prevCpMins)  return 'during';  // left mid-window
  return 'after';                                    // already gone before window opened
}

/**
 * Clears a checkpoint's metric cells and paints the entire block
 * (metrics + progress columns) with the given status color.
 * Used for Off, CTO, VTO, and Unscheduled reps.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row
 * @param {{ startCol: number }} section
 * @param {{ progressStartCol: number }} layout
 * @param {string} colorStatus   - e.g. 'cto', 'vto', 'unscheduled', 'exempt'
 * @param {string} statusLabel   - Text written to the Actions Taken cell
 */
function applyStatusBlock_(sheet, row, section, layout, colorStatus, statusLabel) {
  // Clear and color the 4 metric cells
  const metricRange = sheet.getRange(row, section.startCol, 1, 4);
  metricRange.clearContent();
  applyPacingColor_(metricRange, colorStatus);

  // For not-yet-started reps: don't mark Exempt — they will work later.
  // On Track = No (they have 0s now), EOD Goal Met = blank (TBD at EOD).
  // For all other statuses (CTO, VTO, Off): mark fully Exempt.
  const isNotYetStarted = colorStatus === 'unscheduled';

  applyPacingColor_(sheet.getRange(row, layout.progressStartCol), colorStatus);
  sheet.getRange(row, layout.progressStartCol).setValue(isNotYetStarted ? 'No' : 'Exempt');
  sheet.getRange(row, layout.progressStartCol + 1).setValue('');
  sheet.getRange(row, layout.progressStartCol + 2).setValue(statusLabel);
  sheet.getRange(row, layout.progressStartCol + 3).setValue(isNotYetStarted ? '' : 'Exempt');
  sheet.getRange(row, layout.progressStartCol + 4).setValue('Status: ' + statusLabel);
}

/**
 * Retroactively repaints all earlier checkpoint columns for a rep row
 * when their status changes (e.g. CTO entered late in the day).
 *
 * For full Off/CTO/VTO/Unscheduled: all earlier checkpoints are repainted.
 * Partial CTO/VTO are intentionally left alone — earlier checkpoints
 * where the rep was actively working should keep their real metric colors.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {{ sections: Array }} layout
 * @param {number} row
 * @param {{ key: string }} currentCheckpoint
 * @param {string} colorStatus
 */
function retroactivelyUpdateEarlierCheckpoints_(sheet, layout, row, currentCheckpoint, colorStatus) {
  // Partial statuses: don't retroactively overwrite real metric data
  if (colorStatus === 'exempt') return;

  const currentIdx = layout.sections.findIndex(s => s.key === currentCheckpoint.key);
  if (currentIdx <= 0) return; // nothing before this one

  for (let s = 0; s < currentIdx; s++) {
    const sec = layout.sections[s];
    const range = sheet.getRange(row, sec.startCol, 1, 4);
    // Only repaint if the cell has no real numeric data (i.e. was already exempt/blank)
    const vals = range.getValues()[0];
    const hasData = vals.some(v => v !== '' && v !== 0 && !isNaN(Number(v)));
    if (!hasData) {
      applyPacingColor_(range, colorStatus);
    }
  }
}

// ── Metric writing / clearing ─────────────────────────────────────────────────

/**
 * Writes the four metric values into their checkpoint columns for a row.
 * CSAT is written as '' when no surveys exist (keeps the cell clean).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row
 * @param {{ closedCol, repliedCol, messagesCol, csatCol }} section
 * @param {{ closedTickets, ticketsReplied, messagesSent, csat }} actual
 */
function writeMetrics_(sheet, row, section, actual) {
  sheet.getRange(row, section.closedCol).setValue(actual.closedTickets);
  sheet.getRange(row, section.repliedCol).setValue(actual.ticketsReplied);
  sheet.getRange(row, section.messagesCol).setValue(actual.messagesSent);
  sheet.getRange(row, section.csatCol).setValue(actual.csat === '' ? '' : actual.csat);
}

/**
 * Clears metric cells for a row and resets their background to the
 * default light-yellow data zone color.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row
 * @param {{ startCol: number }} section
 */
function clearMetricBlock_(sheet, row, section) {
  sheet.getRange(row, section.startCol, 1, 4).clearContent().setBackground('#fff9c4');
}

// ── Sheet creation helpers ────────────────────────────────────────────────────

/**
 * Gets the daily sheet for a given date, creating and building it
 * from scratch if it doesn't already exist.
 *
 * @param {Date} dateObj
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateDailySheetForDate_(dateObj) {
  const ss   = SpreadsheetApp.getActive();
  const name = formatDailySheetName_(dateObj);
  let sh     = ss.getSheetByName(name);
  if (sh) return sh;

  sh = ss.insertSheet(name);
  buildDailySheet_(sh, dateObj);
  return sh;
}

// ── Checkpoint detection ──────────────────────────────────────────────────────

/**
 * Returns the checkpoint whose hour matches the current hour, or null.
 * Used by publishCurrentCheckpoint() to auto-detect which block to fill.
 *
 * @param {Date} dateObj
 * @returns {{ key, hour, minute, percent, label }|null}
 */
function getCurrentCheckpoint_(dateObj) {
  const hour = Number(Utilities.formatDate(dateObj, CFG.timezone, 'H'));
  return CFG.checkpoints.find(cp => cp.hour === hour) || null;
}

// ── Ad-hoc rerun for a past date ──────────────────────────────────────────────

/**
 * Re-runs the EOD publish for an arbitrary past date.
 *
 * @param {number} year  - 4-digit year (e.g. 2026)
 * @param {number} month - 1-based month (e.g. 3 for March)
 * @param {number} day   - Day of month
 */
function rerunEODForDate(year, month, day) {
  const date = new Date(year, month - 1, day);
  const eod  = CFG.checkpoints.find(c => c.key === 'EOD');
  if (!eod) throw new Error('EOD checkpoint not found');
  publishCheckpointForDate_(date, eod);
}

// ── Underperformance flagging & adjustment ────────────────────────────────────

/** Threshold below which a rep is flagged for underperformance review. */
const UNDERPERFORMANCE_THRESHOLD = 0.65;

/**
 * Returns true if any non-CSAT metric is below UNDERPERFORMANCE_THRESHOLD
 * of its adjusted goal.
 *
 * @param {{ closedTickets, ticketsReplied, messagesSent, csat }} actual
 * @param {{ closedTickets, ticketsReplied, messagesSent }} goals
 * @returns {boolean}
 */
function isUnderperforming_(actual, goals) {
  const metrics = [
    { actual: actual.closedTickets,  goal: goals.closedTickets  },
    { actual: actual.ticketsReplied, goal: goals.ticketsReplied },
    { actual: actual.messagesSent,   goal: goals.messagesSent   }
  ];

  return metrics.some(m =>
    m.goal > 0 && (m.actual / m.goal) < UNDERPERFORMANCE_THRESHOLD
  );
}

/**
 * Reads the Reason and Hours Removed columns on the active daily tab and
 * applies the supervisor's hour-based adjustments to all checkpoint colors
 * and EOD Goal Met.
 *
 * Adjustment logic:
 *   Hours Removed (number) → subtracts that many hours from the rep's effective
 *     worked hours, recomputes shift-scaled goals, and repaints all checkpoint
 *     metric cells that already have data.
 *
 * Any agent can be adjusted — no underperformance flag required.
 * Run from the menu after supervisors have filled in the adjustment columns.
 */
function applyGoalAdjustments() {
  const ss        = SpreadsheetApp.getActive();
  const sheet     = ss.getActiveSheet();
  const sheetName = sheet.getName();

  if (!parseDailySheetName_(sheetName)) {
    SpreadsheetApp.getUi().alert('Please navigate to a daily tab before running Apply Goal Adjustments.');
    return;
  }

  const dateObj = parseDailySheetName_(sheetName);

  const layout     = getLayout_();
  const eodSection = layout.sections.find(s => s.key === 'EOD');
  const fullRoster = getDisplayRoster_();
  const goalsMap   = getGoalsMap_();
  const rowMap     = getDailySheetRowMap_(sheet);

  let adjustedCount = 0;

  for (let i = 0; i < fullRoster.length; i++) {
    const rep = fullRoster[i];
    const row = rowMap[normalizeName_(rep.repName)];
    if (!row) continue;

    const reason      = String(sheet.getRange(row, layout.reviewReasonCol).getValue() || '').trim();
    const adjustRaw   = sheet.getRange(row, layout.reviewAdjustCol).getValue();
    const hoursRemoved = Number(adjustRaw);

    if (adjustRaw === '' || adjustRaw === null || isNaN(hoursRemoved)) continue;

    adjustedCount++;

    const goals         = getEffectiveGoals_(goalsMap, rep.repName);
    const useShiftGoals =
      String(getConfigValue_('USE_SHIFT_BASED_GOALS', true)).toLowerCase() !== 'false' &&
      goals.useShift !== false;

    // Read effective hours from the Notes column — written at publish time and
    // survives schedule changes, so adjustments work on any past tab
    const noteText       = String(sheet.getRange(row, layout.notesCol).getValue() || '');
    const effMatch       = noteText.match(/Effective:\s*([\d.]+)/);
    const effectiveHours = effMatch
      ? Number(effMatch[1])
      : Number(getConfigValue_('STANDARD_SHIFT_HOURS', CFG.standardShiftHours));

    // Reduce worked hours by the supervisor-entered amount (floor at 0)
    const adjustedHours = Math.max(0, effectiveHours - hoursRemoved);

    const adjustedGoals = {
      closedTickets:  useShiftGoals ? adjustGoalByShift_(goals.closedTickets,  adjustedHours) : goals.closedTickets,
      ticketsReplied: useShiftGoals ? adjustGoalByShift_(goals.ticketsReplied, adjustedHours) : goals.ticketsReplied,
      messagesSent:   useShiftGoals ? adjustGoalByShift_(goals.messagesSent,   adjustedHours) : goals.messagesSent,
      csat:           goals.csat
    };

    // Repaint every checkpoint that already has data written
    layout.sections.forEach(section => {
      const vals = sheet.getRange(row, section.startCol, 1, 4).getValues()[0];
      const hasData = vals.some(v => v !== '' && v !== null);
      if (!hasData) return;

      const actual = {
        closedTickets:  Number(vals[0]) || 0,
        ticketsReplied: Number(vals[1]) || 0,
        messagesSent:   Number(vals[2]) || 0,
        csat:           vals[3] === '' ? '' : Number(vals[3]) || 0
      };

      const targets = {
        closedTickets:  adjustedGoals.closedTickets  * section.percent,
        ticketsReplied: adjustedGoals.ticketsReplied * section.percent,
        messagesSent:   adjustedGoals.messagesSent   * section.percent,
        csat:           adjustedGoals.csat
      };

      const statuses = getMetricStatuses_(actual, targets);
      applyPacingColor_(sheet.getRange(row, section.closedCol),   statuses.closedTickets);
      applyPacingColor_(sheet.getRange(row, section.repliedCol),  statuses.ticketsReplied);
      applyPacingColor_(sheet.getRange(row, section.messagesCol), statuses.messagesSent);
      applyPacingColor_(sheet.getRange(row, section.csatCol),     statuses.csat);
    });

    // Recompute EOD Goal Met only if EOD data has already been published
    const eodVals    = sheet.getRange(row, eodSection.startCol, 1, 4).getValues()[0];
    const eodHasData = eodVals.some(v => v !== '' && v !== null);
    let   eodMet     = null;
    if (eodHasData) {
      const eodActual = {
        closedTickets:  Number(eodVals[0]) || 0,
        ticketsReplied: Number(eodVals[1]) || 0,
        messagesSent:   Number(eodVals[2]) || 0,
        csat:           eodVals[3] === '' ? '' : Number(eodVals[3]) || 0
      };
      eodMet = computeEodGoalMet_(eodActual, adjustedGoals);
      sheet.getRange(row, layout.progressStartCol + 3).setValue(eodMet);
    }

    // Mark the status column so supervisors know the adjustment was applied
    const hoursLabel = '-' + hoursRemoved + 'h' + (reason ? ' / ' + reason : '');
    const metLabel   = eodMet === 'Yes' ? '✓ Applied: ' : 'Applied: ';
    sheet.getRange(row, layout.reviewFlagCol)
      .setValue(metLabel + hoursLabel)
      .setBackground('#b7e1cd')
      .setFontColor('#000000')
      .setFontWeight('normal');
  }

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    'Goal adjustments applied to ' + adjustedCount + ' rep(s).\n\n' +
    'All checkpoint colors and EOD Goal Met have been updated.'
  );
}

/**
 * were not yet started at those checkpoints but showed red zeros instead.
 * Safe to run multiple times — only overwrites cells with no real data.
 * Delete after use.
 */
function fixNotYetStartedColorsToday() {
  const ss        = SpreadsheetApp.getActive();
  const dateObj   = new Date();
  const sheetName = formatDailySheetName_(dateObj);
  const sheet     = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Today tab not found: ' + sheetName);

  const layout      = getLayout_();
  const fullRoster  = getDisplayRoster_();
  const rowMap      = getDailySheetRowMap_(sheet);
  const scheduleMap = getScheduleMapForDate_(dateObj);

  for (let i = 0; i < fullRoster.length; i++) {
    const rep      = fullRoster[i];
    const row      = rowMap[normalizeName_(rep.repName)];
    if (!row) continue;
    const schedule = getScheduleForRep_(scheduleMap, rep.repName);

    if (!schedule.startText || schedule.hours <= 0) continue;

    const startMins = parseTimeToMinutes_(schedule.startText);

    // For each checkpoint that fired before this rep's shift started,
    // repaint the block black if it has no real data
    layout.sections.forEach(sec => {
      const cpMins = CFG.checkpoints.find(c => c.key === sec.key).hour * 60;
      if (startMins <= cpMins) return; // rep was working by this checkpoint — leave it

      const range = sheet.getRange(row, sec.startCol, 1, 4);
      const vals  = range.getValues()[0];
      const hasData = vals.some(v => v !== '' && v !== 0 && !isNaN(Number(v)));
      if (!hasData) {
        applyPacingColor_(range, 'unscheduled');
      }
    });
  }

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Not-yet-started colors fixed for today.');
}
 /* Run once manually from the Apps Script editor, then safe to delete.
 */
function backfillMarch31() {
  const date = new Date(2026, 2, 31); // March 31, 2026
  const keys = ['11AM', '2PM', '6PM', 'EOD'];

  keys.forEach((key, i) => {
    const cp = CFG.checkpoints.find(c => c.key === key);
    if (!cp) throw new Error('Checkpoint not found: ' + key);
    publishCheckpointForDate_(date, cp);
    if (i < keys.length - 1) Utilities.sleep(3000);
  });

  SpreadsheetApp.getUi().alert('March 31 backfill complete: 11AM, 2PM, 6PM, EOD.');
}

/**
 * Migrates a single sheet's review block from the old "Underperformance Review"
 * layout (percentage dropdowns) to the new "Goal Adjustments" layout
 * (hours-removed numeric input). Used by both single and bulk migration.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function migrateSheetToGoalAdjustments_(sheet) {
  const layout   = getLayout_();
  const lastRow  = sheet.getLastRow();
  const rowCount = Math.max(lastRow - CFG.daily.firstDataRow + 1, 1);

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

  const hoursValidation = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 24)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(CFG.daily.firstDataRow, layout.reviewAdjustCol, rowCount, 1)
    .setDataValidation(hoursValidation);
}

/**
 * Migrates the active daily tab to the Goal Adjustments layout.
 * Run from the menu: Pacing Report → Migrate Tab to Goal Adjustments
 */
function migrateTabToGoalAdjustments() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (!parseDailySheetName_(sheet.getName())) {
    SpreadsheetApp.getUi().alert('Please navigate to a daily tab first.');
    return;
  }
  migrateSheetToGoalAdjustments_(sheet);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Tab migrated to Goal Adjustments layout.');
}

/**
 * Migrates every daily tab in the spreadsheet to the Goal Adjustments layout.
 * Safe to run multiple times — only reformats headers and validation.
 * Run from the menu: Pacing Report → Migrate All Tabs to Goal Adjustments
 */
function migrateAllTabsToGoalAdjustments() {
  const ss      = SpreadsheetApp.getActive();
  const sheets  = ss.getSheets().filter(sh => parseDailySheetName_(sh.getName()));

  if (!sheets.length) {
    SpreadsheetApp.getUi().alert('No daily tabs found.');
    return;
  }

  sheets.forEach(sh => migrateSheetToGoalAdjustments_(sh));
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Migrated ' + sheets.length + ' daily tab(s) to Goal Adjustments layout.');
}

function rebuildAndRepublishToday() {
  const now  = new Date();
  const hour = Number(Utilities.formatDate(now, CFG.timezone, 'H'));

  const fired = CFG.checkpoints.filter(cp => cp.hour <= hour);
  if (!fired.length) {
    SpreadsheetApp.getUi().alert('No checkpoints have fired yet today.');
    return;
  }

  createTodayTab();

  fired.forEach((cp, i) => {
    publishCheckpointForDate_(now, cp);
    if (i < fired.length - 1) Utilities.sleep(3000);
  });

  SpreadsheetApp.getUi().alert(
    'Today tab rebuilt and republished through ' +
    fired[fired.length - 1].label + '.'
  );
}
