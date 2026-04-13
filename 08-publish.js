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
  const scheduleMap = getScheduleMapForDate_(dateObj);
  const goalsMap    = getGoalsMap_();
  const range       = getCheckpointIsoRange_(dateObj, checkpoint);

  // Fetch closed-tickets-per-agent once for the whole team (one API call)
  const closedMap = getClosedPerAgentMap_(range.startIso, range.endIso);

  for (let i = 0; i < fullRoster.length; i++) {
    const rep      = fullRoster[i];
    const row      = CFG.daily.firstDataRow + i;
    const schedule = getScheduleForRep_(scheduleMap, rep.repName);

    // ── Exempt: off / CTO / VTO ───────────────────────────────────────────
    if (schedule.hours <= 0 || ['CTO', 'VTO', 'Off'].includes(schedule.status)) {
      clearMetricBlock_(sheet, row, section);
      sheet.getRange(row, layout.progressStartCol).setValue('Exempt').setBackground('#ffffff');
      sheet.getRange(row, layout.progressStartCol + 1).setValue('');
      sheet.getRange(row, layout.progressStartCol + 2).setValue(schedule.status || 'Off');
      sheet.getRange(row, layout.progressStartCol + 3).setValue('Exempt');
      sheet.getRange(row, layout.progressStartCol + 4).setValue('Status: ' + (schedule.status || 'Off'));
      continue;
    }

    // ── Active rep: compute goals and actuals ─────────────────────────────
    const goals = getEffectiveGoals_(goalsMap, rep.repName);

    const useShiftGoals =
      String(getConfigValue_('USE_SHIFT_BASED_GOALS', true)).toLowerCase() !== 'false' &&
      goals.useShift !== false;

    const effectiveHours = getEffectiveWorkedHours_(schedule);

    const adjustedGoals = {
      closedTickets:  useShiftGoals ? adjustGoalByShift_(goals.closedTickets,  effectiveHours) : goals.closedTickets,
      ticketsReplied: useShiftGoals ? adjustGoalByShift_(goals.ticketsReplied, effectiveHours) : goals.ticketsReplied,
      messagesSent:   useShiftGoals ? adjustGoalByShift_(goals.messagesSent,   effectiveHours) : goals.messagesSent,
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
    const pacingStatus   = getWorstPacingStatus_(actual, targets);
    const onTrack        = computeOnTrack_(actual, targets);
    const eodMet         = checkpoint.key === 'EOD' ? computeEodGoalMet_(actual, adjustedGoals) : '';
    const onProject      = schedule.inOffice    ? 'In Office'      : '';
    const actions        = schedule.workingLunch ? 'Working Lunch' : '';

    writeMetrics_(sheet, row, section, actual);
    applyPacingColor_(sheet.getRange(row, section.closedCol),   metricStatuses.closedTickets);
    applyPacingColor_(sheet.getRange(row, section.repliedCol),  metricStatuses.ticketsReplied);
    applyPacingColor_(sheet.getRange(row, section.messagesCol), metricStatuses.messagesSent);
    applyPacingColor_(sheet.getRange(row, section.csatCol),     metricStatuses.csat);

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

/**
 * Backfills missed checkpoints for March 31 starting from 11AM.
 * Run once manually from the Apps Script editor, then safe to delete.
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

function fixSheetNow() {
  normalizeCurrentWeekSchedule();
  createTodayTab();
  publishCurrentCheckpoint();

  SpreadsheetApp.getUi().alert(
    'Sheet refreshed.\n\n' +
    'Schedule normalized, today tab checked, and current checkpoint republished.'
  );
}

function fixYesterdayAndToday() {
  normalizeCurrentWeekSchedule();
  rerunYesterday();
  publishAllCheckpointsToday();

  SpreadsheetApp.getUi().alert(
    'Yesterday and today have been rebuilt using the current schedule.'
  );
}

function rerunAllCheckpointsForDate(year, month, day) {
  const date = new Date(year, month - 1, day);

  CFG.checkpoints.forEach(cp => {
    publishCheckpointForDate_(date, cp);
  });
}

function rerunYesterday() {
  const d = new Date();
  d.setDate(d.getDate() - 1);

  rerunAllCheckpointsForDate(
    d.getFullYear(),
    d.getMonth() + 1,
    d.getDate()
  );
}

function refreshTodayOnly() {
  normalizeCurrentWeekSchedule();
  publishAllCheckpointsToday();

  SpreadsheetApp.getUi().alert('Today has been refreshed.');
}

