// =============================================================================
// goals.gs
// Goal loading, shift-based adjustment, checkpoint targets,
// on-track computation, and pacing color logic.
// =============================================================================

// ── Goals map ─────────────────────────────────────────────────────────────────

/**
 * Reads the Goals sheet and returns a map of rep name → goal object.
 * The special key '_default' holds team-wide goals used when no
 * rep-specific override is found.
 *
 * Goals sheet layout (1-indexed rows):
 *   Row 3:  team defaults [closed, replied, messages, csat, useShiftGoals]
 *   Row 15+: per-rep overrides [name, closed, replied, messages, csat]
 *
 * @returns {Object} { _default: {...}, 'rep name': {...}, ... }
 */
function getGoalsMap_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.goalsSheetName);
  if (!sh) throw new Error('Goals sheet missing.');

  const values   = sh.getDataRange().getValues();
  const defaults = values[2]; // Row 3 (0-indexed row 2)

  const map = {
    _default: {
      closedTickets:  defaults[0] !== '' ? Number(defaults[0]) : CFG.baselineGoals.closedTickets,
      ticketsReplied: defaults[1] !== '' ? Number(defaults[1]) : CFG.baselineGoals.ticketsReplied,
      messagesSent:   defaults[2] !== '' ? Number(defaults[2]) : CFG.baselineGoals.messagesSent,
      csat:           defaults[3] !== '' ? Number(defaults[3]) : CFG.baselineGoals.csat,
      useShift:       String(defaults[4]).toLowerCase() !== 'false'
    }
  };

  // Per-rep overrides start at row 15 (0-indexed row 14)
  for (let r = 14; r < values.length; r++) {
    const row  = values[r] || [];
    const name = String(row[0] || '').trim();
    if (!name) continue;

    map[normalizeName_(name)] = {
      closedTickets:  row[1] !== '' ? Number(row[1]) : map._default.closedTickets,
      ticketsReplied: row[2] !== '' ? Number(row[2]) : map._default.ticketsReplied,
      messagesSent:   row[3] !== '' ? Number(row[3]) : map._default.messagesSent,
      csat:           row[4] !== '' ? Number(row[4]) : map._default.csat,
      useShift:       map._default.useShift
    };
  }

  return map;
}

/**
 * Returns the effective goals for a rep, falling back to team defaults.
 * @param {Object} goalsMap - From getGoalsMap_()
 * @param {string} repName
 * @returns {{ closedTickets, ticketsReplied, messagesSent, csat, useShift }}
 */
function getEffectiveGoals_(goalsMap, repName) {
  return goalsMap[normalizeName_(repName)] || goalsMap._default;
}

// ── Shift-based goal adjustment ───────────────────────────────────────────────

/**
 * Scales a daily goal proportionally to actual hours worked vs. standard shift.
 * e.g. A 4-hour shift on a standard 8-hour day → 50% of the daily goal.
 *
 * @param {number} baseGoal     - Full-day goal for a standard shift.
 * @param {number} hoursWorked  - Effective hours worked today (after lunch deduction).
 * @returns {number} Adjusted goal, or 0 if hoursWorked is 0.
 */
function adjustGoalByShift_(baseGoal, hoursWorked) {
  const standardHours = Number(getConfigValue_('STANDARD_SHIFT_HOURS', CFG.standardShiftHours));
  if (!hoursWorked || hoursWorked <= 0) return 0;
  return baseGoal * (hoursWorked / standardHours);
}

/**
 * Calculates effective hours worked after applying the auto-lunch deduction.
 * If a shift is >= threshold hours, one lunch hour is subtracted.
 *
 * Config keys:
 *   AUTO_LUNCH_THRESHOLD_HOURS  (default: 9)
 *   AUTO_LUNCH_DEDUCTION_HOURS  (default: 1)
 *
 * Examples at defaults: 8h → 8h, 9h → 8h, 10h → 9h
 *
 * @param {{ hours: number }} schedule - Schedule entry for the rep.
 * @returns {number}
 */
function getEffectiveWorkedHours_(schedule) {
  const threshold = Number(getConfigValue_('AUTO_LUNCH_THRESHOLD_HOURS', 9));
  const deduction = Number(getConfigValue_('AUTO_LUNCH_DEDUCTION_HOURS', 1));
  let hours = Number(schedule.hours || 0);
  if (hours >= threshold) hours -= deduction;
  return Math.max(0, hours);
}

// ── Checkpoint targets ────────────────────────────────────────────────────────

/**
 * Returns the expected value at a checkpoint given the full-day goal
 * and the checkpoint's pacing percentage.
 *
 * @param {number} goal             - Full-day (shift-adjusted) goal.
 * @param {number} checkpointPercent - e.g. 0.25 for 9 AM
 * @returns {number}
 */
function getCheckpointTarget_(goal, checkpointPercent) {
  return goal * checkpointPercent;
}

// ── On-track / EOD computation ────────────────────────────────────────────────

/**
 * Returns 'Yes' if all metrics meet or exceed their checkpoint targets,
 * 'No' otherwise. CSAT is skipped if no surveys have been received yet.
 *
 * @param {{ closedTickets, ticketsReplied, messagesSent, csat }} actual
 * @param {{ closedTickets, ticketsReplied, messagesSent, csat }} targets
 * @returns {'Yes'|'No'}
 */
function computeOnTrack_(actual, targets) {
  const checks = [
    actual.closedTickets  >= targets.closedTickets,
    actual.ticketsReplied >= targets.ticketsReplied,
    actual.messagesSent   >= targets.messagesSent,
    actual.csat === '' || actual.csat >= targets.csat
  ];
  return checks.every(Boolean) ? 'Yes' : 'No';
}

/**
 * Same as computeOnTrack_ but compares against full-day goals (used at EOD).
 *
 * @param {{ closedTickets, ticketsReplied, messagesSent, csat }} actual
 * @param {{ closedTickets, ticketsReplied, messagesSent, csat }} goals
 * @returns {'Yes'|'No'}
 */
function computeEodGoalMet_(actual, goals) {
  const checks = [
    actual.closedTickets  >= goals.closedTickets,
    actual.ticketsReplied >= goals.ticketsReplied,
    actual.messagesSent   >= goals.messagesSent,
    actual.csat === '' || actual.csat >= goals.csat
  ];
  return checks.every(Boolean) ? 'Yes' : 'No';
}

// ── Pacing status & coloring ──────────────────────────────────────────────────

/**
 * Returns a pacing status string for a single metric.
 *
 * Two modes (controlled by PACING_MODE in Config):
 *   'percent' (default): ratio = actual / target
 *     green  = ratio >= PACING_GREEN_MIN  (default 1.0)
 *     yellow = ratio >= PACING_YELLOW_MIN (default 0.9)
 *     red    = below yellow
 *
 *   'raw': shortfall = target - actual
 *     green  = shortfall <= PACING_GREEN_MAX_SHORTFALL  (default 0)
 *     yellow = shortfall <= PACING_YELLOW_MAX_SHORTFALL (default 5)
 *     red    = above yellow
 *
 * Returns 'exempt' when actual or target is missing / zero.
 *
 * @param {number|string} actual
 * @param {number|string} target
 * @returns {'green'|'yellow'|'red'|'exempt'}
 */
function getPacingStatus_(actual, target) {
  if (actual === '' || target === '' || target === null || target === undefined || Number(target) <= 0) {
    return 'exempt';
  }

  const mode = String(getConfigValue_('PACING_MODE', 'percent')).toLowerCase();

  if (mode === 'raw') {
    const greenMax  = Number(getConfigValue_('PACING_GREEN_MAX_SHORTFALL',  0));
    const yellowMax = Number(getConfigValue_('PACING_YELLOW_MAX_SHORTFALL', 5));
    const shortfall = Number(target) - Number(actual);

    if (shortfall <= greenMax)  return 'green';
    if (shortfall <= yellowMax) return 'yellow';
    return 'red';
  }

  // Default: percent mode
  const greenMin  = Number(getConfigValue_('PACING_GREEN_MIN',  1.0));
  const yellowMin = Number(getConfigValue_('PACING_YELLOW_MIN', 0.9));
  const ratio     = Number(actual) / Number(target);

  if (ratio >= greenMin)  return 'green';
  if (ratio >= yellowMin) return 'yellow';
  return 'red';
}

/**
 * Returns pacing status for each metric individually.
 * CSAT is 'exempt' if no surveys were received (actual === '').
 *
 * @param {{ closedTickets, ticketsReplied, messagesSent, csat }} actual
 * @param {{ closedTickets, ticketsReplied, messagesSent, csat }} targets
 * @returns {{ closedTickets, ticketsReplied, messagesSent, csat }}
 */
function getMetricStatuses_(actual, targets) {
  return {
    closedTickets:  getPacingStatus_(actual.closedTickets,  targets.closedTickets),
    ticketsReplied: getPacingStatus_(actual.ticketsReplied, targets.ticketsReplied),
    messagesSent:   getPacingStatus_(actual.messagesSent,   targets.messagesSent),
    csat:           actual.csat === '' ? 'exempt' : getPacingStatus_(actual.csat, targets.csat)
  };
}

/**
 * Returns the single worst pacing status across all metrics.
 * Priority: red > yellow > green > exempt.
 *
 * @param {{ closedTickets, ticketsReplied, messagesSent, csat }} actual
 * @param {{ closedTickets, ticketsReplied, messagesSent, csat }} targets
 * @returns {'red'|'yellow'|'green'|'exempt'}
 */
function getWorstPacingStatus_(actual, targets) {
  const statuses = Object.values(getMetricStatuses_(actual, targets));
  if (statuses.includes('red'))    return 'red';
  if (statuses.includes('yellow')) return 'yellow';
  if (statuses.includes('green'))  return 'green';
  return 'exempt';
}

/**
 * Applies the appropriate background color to a cell based on pacing status.
 *
 * Status values and their colors:
 *   green       - on target           (#b6d7a8 soft green)
 *   yellow      - near target         (#ffe599 soft yellow)
 *   red         - behind target       (#f4cccc soft red)
 *   cto         - full/partial CTO    (#d9b8f5 muted purple)
 *   vto         - full/partial VTO    (#fce5cd soft orange)
 *   unscheduled - not on schedule     (#000000 black)
 *   exempt      - off / no data       (#ffffff white)
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 * @param {'green'|'yellow'|'red'|'cto'|'vto'|'unscheduled'|'exempt'} status
 */
function applyPacingColor_(range, status) {
  const colors = {
    green:        '#b6d7a8',
    yellow:       '#ffe599',
    red:          '#f4cccc',
    cto:          '#d9b8f5',
    vto:          '#fce5cd',
    unscheduled:  '#000000'
  };
  const bg   = colors[status] || '#ffffff';
  const font = status === 'unscheduled' ? '#ffffff' : '#000000';
  range.setBackground(bg).setFontColor(font);
}