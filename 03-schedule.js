// =============================================================================
// schedule.gs
// Schedule normalization, parsing, and lookup.
//
// Flow:
//   normalizeCurrentWeekSchedule()
//     → reads the raw Schedule sheet
//     → parses each cell with parseScheduleCell_()
//     → writes clean rows to Schedule_Normalized
//
//   getScheduleMapForDate_(dateObj)
//     → primary: reads Schedule_Normalized (fast, already parsed)
//     → fallback: parses the raw Schedule sheet directly
// =============================================================================

// ── Normalization (run once per week) ─────────────────────────────────────────

/**
 * Reads the active Schedule sheet, parses every agent/date cell,
 * and writes the results to the Schedule_Normalized sheet.
 * Run this each week after the schedule is updated.
 */
function normalizeCurrentWeekSchedule() {
  const ss = SpreadsheetApp.getActive();
  const scheduleName = getConfigValue_('SCHEDULE_SHEET_NAME', CFG.scheduleSheetName);
  const scheduleSheet = ss.getSheetByName(scheduleName);

  if (!scheduleSheet) {
    throw new Error('Schedule sheet not found. Configured name: ' + scheduleName);
  }

  const targetSheet = getOrCreateSheet_(ss, CFG.normalizedScheduleSheetName);
  targetSheet.clear();
  targetSheet.getRange(1, 1, 1, 10).setValues([[
    'Date', 'Agent Name', 'Manager', 'Start', 'End', 'Hours',
    'Status', 'In Office', 'Working Lunch', 'Raw Value'
  ]]);

  const range  = scheduleSheet.getDataRange();
  const values = range.getDisplayValues();
  const bgs    = range.getBackgrounds();
  const out    = [];

  // Collect date columns from header row 2
  const dateCols = [];
  for (let c = CFG.schedule.firstDateCol - 1; c < values[CFG.schedule.headerRow2 - 1].length; c++) {
    const dateText = String(values[CFG.schedule.headerRow2 - 1][c] || '').trim();
    if (dateText) {
      dateCols.push({ colIndex0: c, dateText: normalizeDateText_(dateText) });
    }
  }

  // Walk agent rows
  for (let r = CFG.schedule.firstDataRow - 1; r < values.length; r++) {
    const agentName = String(values[r][CFG.schedule.nameCol - 1] || '').trim();
    if (!agentName) continue;

    const manager = String(values[r][CFG.schedule.managerCol - 1] || '').trim();

    dateCols.forEach(dc => {
      const rawValue = values[r][dc.colIndex0];
      const bg       = bgs[r][dc.colIndex0];
      const parsed   = parseScheduleCell_(rawValue, bg);

      out.push([
        dc.dateText,
        agentName,
        manager,
        parsed.startText,
        parsed.endText,
        parsed.hours,
        parsed.status,
        parsed.inOffice,
        parsed.workingLunch,
        String(rawValue || '')
      ]);
    });
  }

  if (out.length) {
    targetSheet.getRange(2, 1, out.length, out[0].length).setValues(out);
  }

  targetSheet.autoResizeColumns(1, 10);
}

// ── Schedule map for a single date ────────────────────────────────────────────

/**
 * Returns a map of { normalizedName → schedule entry } for a given date.
 * Each name is stored under both full-name and first-name keys for
 * flexible matching during publish.
 *
 * Reads from Schedule_Normalized first; falls back to raw sheet parsing
 * if the normalized sheet is empty.
 *
 * @param {Date} dateObj
 * @returns {Object}
 */
function getScheduleMapForDate_(dateObj) {
  const ss         = SpreadsheetApp.getActive();
  const normalized = ss.getSheetByName(CFG.normalizedScheduleSheetName);
  const targetDate = formatDailySheetName_(dateObj);
  const map        = {};

  if (normalized && normalized.getLastRow() > 1) {
    const values = normalized.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      const row      = values[i];
      const dateText = normalizeDateText_(String(row[0] || ''));
      if (dateText !== targetDate) continue;

      const agent = String(row[1] || '').trim();
      if (!agent) continue;

      const entry = {
        startText:   String(row[3] || ''),
        endText:     String(row[4] || ''),
        hours:       Number(row[5] || 0),
        status:      String(row[6] || 'Off'),
        inOffice:    row[7] === true || String(row[7]).toLowerCase() === 'true',
        workingLunch: row[8] === true || String(row[8]).toLowerCase() === 'true'
      };

      map[normalizeName_(agent)]     = entry;
      map[normalizeFirstName_(agent)] = entry;
    }

    if (Object.keys(map).length) return map;
  }

  // Fallback: parse the raw schedule sheet directly
  return parseScheduleSheetForDate_(dateObj);
}

/**
 * Parses the raw Schedule sheet for a single date column.
 * Used as a fallback when Schedule_Normalized is not populated.
 *
 * @param {Date} dateObj
 * @returns {Object}
 */
function parseScheduleSheetForDate_(dateObj) {
  const ss           = SpreadsheetApp.getActive();
  const scheduleName = getConfigValue_('SCHEDULE_SHEET_NAME', CFG.scheduleSheetName);
  const scheduleSheet = ss.getSheetByName(scheduleName);

  if (!scheduleSheet) throw new Error('Schedule sheet not found.');

  const values     = scheduleSheet.getDataRange().getDisplayValues();
  const bgs        = scheduleSheet.getDataRange().getBackgrounds();
  const targetDate = formatDailySheetName_(dateObj);
  let   targetCol0 = -1;

  for (let c = CFG.schedule.firstDateCol - 1; c < values[CFG.schedule.headerRow2 - 1].length; c++) {
    const dateText = normalizeDateText_(String(values[CFG.schedule.headerRow2 - 1][c] || ''));
    if (dateText === targetDate) {
      targetCol0 = c;
      break;
    }
  }

  if (targetCol0 === -1) {
    throw new Error(
      'Could not find date ' + targetDate +
      ' in row ' + CFG.schedule.headerRow2 + ' of schedule sheet.'
    );
  }

  const map = {};
  for (let r = CFG.schedule.firstDataRow - 1; r < values.length; r++) {
    const agent = String(values[r][CFG.schedule.nameCol - 1] || '').trim();
    if (!agent) continue;

    const entry = parseScheduleCell_(values[r][targetCol0], bgs[r][targetCol0]);
    map[normalizeName_(agent)]     = entry;
    map[normalizeFirstName_(agent)] = entry;
  }

  return map;
}

// ── Cell parser ───────────────────────────────────────────────────────────────

/**
 * Parses a single schedule cell value and its background color into a
 * structured schedule entry.
 *
 * Status values:
 *   'Active'       - working a normal shift
 *   'OT'           - overtime shift
 *   'Partial VTO'  - working part of day, rest is VTO
 *   'Partial CTO'  - working part of day, rest is CTO
 *   'CTO'          - full day CTO (no hours)
 *   'VTO'          - full day VTO (no hours)
 *   'Off'          - no shift / empty cell
 *
 * inOffice is derived from a yellow cell background.
 * workingLunch is derived from "(WL)", "WL", or "Working Lunch" in the cell text.
 *
 * @param {string} rawValue - Display value from the schedule cell.
 * @param {string} bg       - Background color hex string.
 * @returns {{ status, hours, startText, endText, inOffice, workingLunch }}
 */
function parseScheduleCell_(rawValue, bg) {
  const raw        = String(rawValue || '').trim();
  const text       = raw.toLowerCase();
  const inOffice   = isYellow_(bg);
  const workingLunch = /\(wl\)|\bwl\b|working lunch/i.test(raw);

  if (!raw) {
    return { status: 'Off', hours: 0, startText: '', endText: '', inOffice, workingLunch };
  }

  // Full-day CTO (no time range present)
  if (/^cto$/i.test(raw) || (text.includes('cto') && !hasTimeRange_(text))) {
    return { status: 'CTO', hours: 0, startText: '', endText: '', inOffice, workingLunch };
  }

  // Full-day VTO (no time range present)
  if (/^vto$/i.test(raw) || (text.includes('vto') && !hasTimeRange_(text))) {
    return { status: 'VTO', hours: 0, startText: '', endText: '', inOffice, workingLunch };
  }

  const ranges = extractEffectiveTimeRanges_(raw);
  const hours  = ranges.reduce((sum, r) => sum + getHoursBetween_(r.start, r.end), 0);

  let status = 'Active';
  if (text.includes('vto') && hours > 0) status = 'Partial VTO';
  if (text.includes('cto') && hours > 0) status = 'Partial CTO';
  if (/\bot\b/i.test(raw))              status = 'OT';

  return {
    status:      hours > 0 ? status : 'Off',
    hours:       round2_(hours),
    startText:   ranges.length ? ranges[0].start : '',
    endText:     ranges.length ? ranges[ranges.length - 1].end : '',
    inOffice,
    workingLunch
  };
}

// ── Time range extraction ─────────────────────────────────────────────────────

/**
 * Extracts the "effective" time ranges from a cell value.
 *
 * Special case: if the cell has exactly one main range and one parenthetical
 * range (e.g. "8am-4pm (6am-2pm OT option)"), only the main range is used.
 * This avoids double-counting optional/alternate shifts.
 *
 * @param {string} raw
 * @returns {Array<{ start: string, end: string }>}
 */
function extractEffectiveTimeRanges_(raw) {
  const cleaned = String(raw || '').replace(/[–—]/g, '-').replace(/\n/g, ' ').trim();

  // Parenthetical time ranges, e.g. "(8am-4pm)"
  const parenPattern = /\(\s*\d{1,2}(?::\d{2})?\s*(?:am|pm)\s*-\s*\d{1,2}(?::\d{2})?\s*(?:am|pm)\s*\)/gi;
  const parenMatches = cleaned.match(parenPattern) || [];

  // Main body without parenthetical ranges
  const withoutParens = cleaned.replace(parenPattern, '');
  const mainMatches   = withoutParens.match(
    /\d{1,2}(?::\d{2})?\s*(?:am|pm)\s*-\s*\d{1,2}(?::\d{2})?\s*(?:am|pm)/gi
  ) || [];

  // If one main + one parenthetical, treat the paren as an alternative — ignore it
  if (mainMatches.length === 1 && parenMatches.length >= 1) {
    return [splitRange_(mainMatches[0])];
  }

  // Otherwise use all ranges found in the full string
  const allMatches = cleaned.match(
    /\d{1,2}(?::\d{2})?\s*(?:am|pm)\s*-\s*\d{1,2}(?::\d{2})?\s*(?:am|pm)/gi
  ) || [];

  return allMatches.map(splitRange_);
}

/**
 * Splits a "Xam-Ypm" string into { start, end }.
 * @param {string} match
 * @returns {{ start: string, end: string }}
 */
function splitRange_(match) {
  const parts = match.replace(/[()]/g, '').split('-');
  return { start: parts[0].trim(), end: parts[1].trim() };
}

/**
 * Returns true if the text contains a time range pattern like "8am-4pm".
 * Used to distinguish "CTO 8am-12pm" (partial) from plain "CTO" (full day).
 * @param {string} text
 * @returns {boolean}
 */
function hasTimeRange_(text) {
  const cleaned = String(text || '').replace(/[–—]/g, '-');
  return /\d{1,2}(?::\d{2})?\s*(?:am|pm)\s*-\s*\d{1,2}(?::\d{2})?\s*(?:am|pm)/i.test(cleaned);
}

// ── Time math ─────────────────────────────────────────────────────────────────

/**
 * Returns hours between two time strings (e.g. "8am", "4:30pm").
 * Handles overnight shifts by adding 24h when end < start.
 * @param {string} startText
 * @param {string} endText
 * @returns {number}
 */
function getHoursBetween_(startText, endText) {
  const startMins = parseTimeToMinutes_(startText);
  const endMins   = parseTimeToMinutes_(endText);
  let diff = endMins - startMins;
  if (diff < 0) diff += 24 * 60; // overnight shift
  return diff / 60;
}

/**
 * Converts a time string like "2pm", "10:30am" into total minutes from midnight.
 * @param {string} timeText
 * @returns {number}
 */
function parseTimeToMinutes_(timeText) {
  const m = String(timeText).trim().match(/(\d{1,2})(?::(\d{2}))?\s*(am|pm)/i);
  if (!m) return 0;

  let hour       = Number(m[1]);
  const minute   = Number(m[2] || 0);
  const ampm     = m[3].toLowerCase();

  if (ampm === 'pm' && hour !== 12) hour += 12;
  if (ampm === 'am' && hour === 12) hour = 0;

  return hour * 60 + minute;
}

// ── Color helper ──────────────────────────────────────────────────────────────

/**
 * Returns true if a cell background color indicates "in office" (yellow tones).
 * Add additional hex values here if your schedule uses other yellow shades.
 * @param {string} hex
 * @returns {boolean}
 */
function isYellow_(hex) {
  if (!hex) return false;
  return ['#ffff00', '#fff200', '#fce8b2', '#f4cccc', '#ffd966']
    .indexOf(String(hex).toLowerCase()) !== -1;
}

// ── Default / fallback ────────────────────────────────────────────────────────

/**
 * Returns a safe "not scheduled" entry for reps missing from the schedule.
 * @returns {{ status, hours, startText, endText, inOffice, workingLunch }}
 */
function defaultSchedule_() {
  return { status: 'Off', hours: 0, startText: '', endText: '', inOffice: false, workingLunch: false };
}

// ── Schedule rollover (Saturday EOD) ─────────────────────────────────────────

/**
 * On Saturday at 11pm, advances CURRENT_SCHEDULE_TAB to NEXT_SCHEDULE_TAB
 * in Config and re-normalizes the schedule.
 * Called automatically from publishEOD().
 */
function rolloverScheduleTabIfNeeded_() {
  const now     = new Date();
  const dayName = Utilities.formatDate(now, CFG.timezone, 'EEE');
  const hour    = Number(Utilities.formatDate(now, CFG.timezone, 'H'));

  if (dayName !== 'Sat' || hour < 23) return;

  const currentTab = String(getConfigValue_('CURRENT_SCHEDULE_TAB', '') || '').trim();
  const nextTab    = String(getConfigValue_('NEXT_SCHEDULE_TAB',    '') || '').trim();

  if (!nextTab || nextTab === currentTab) return;

  setConfigValue_('CURRENT_SCHEDULE_TAB', nextTab);
  setConfigValue_('NEXT_SCHEDULE_TAB', '');

  try {
    normalizeCurrentWeekSchedule();
  } catch (err) {
    Logger.log('Schedule rollover completed, but normalize failed: ' + err);
  }
}