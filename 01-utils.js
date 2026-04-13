// =============================================================================
// 01-utils.gs
// Pure utility helpers with no side effects.
// No Sheets API calls, no CFG references (except timezone for date formatting).
// Safe to call from anywhere.
// =============================================================================

// ── Date / sheet name formatting ─────────────────────────────────────────────

/**
 * Returns the daily sheet tab name for a given date, e.g. "3/30/26".
 * @param {Date} dateObj
 * @returns {string}
 */
function formatDailySheetName_(dateObj) {
  return Utilities.formatDate(dateObj, CFG.timezone, 'M/d/yy');
}

/**
 * Parses a daily sheet tab name like "3/30/26" back into a Date object.
 * Returns null if the name doesn't match the expected pattern.
 * @param {string} name
 * @returns {Date|null}
 */
function parseDailySheetName_(name) {
  const m = String(name || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})$/);
  if (!m) return null;
  return new Date(2000 + Number(m[3]), Number(m[1]) - 1, Number(m[2]));
}

/**
 * Normalizes a raw date string from the schedule sheet into the same
 * M/d/yy format used by formatDailySheetName_.
 * Handles both 2-digit and 4-digit years.
 * @param {string} dateText
 * @returns {string}
 */
function normalizeDateText_(dateText) {
  const trimmed = String(dateText || '').trim();
  const direct = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (direct) {
    let year = direct[3];
    if (year.length === 4) year = year.slice(2);
    return Number(direct[1]) + '/' + Number(direct[2]) + '/' + year;
  }
  return trimmed;
}

// ── Name normalization ────────────────────────────────────────────────────────

/**
 * Lowercases and collapses whitespace in a name for fuzzy matching.
 * @param {string} name
 * @returns {string}
 */
function normalizeName_(name) {
  return String(name || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

/**
 * Returns just the first word of a normalized name.
 * Used as a fallback when full-name matching fails.
 * @param {string} name
 * @returns {string}
 */
function normalizeFirstName_(name) {
  const cleaned = normalizeName_(name);
  return cleaned.split(' ')[0] || cleaned;
}

// ── Math helpers ──────────────────────────────────────────────────────────────

/** Rounds to 1 decimal place. */
function round1_(num) {
  return Math.round(Number(num || 0) * 10) / 10;
}

/** Rounds to 2 decimal places. */
function round2_(num) {
  return Math.round(Number(num || 0) * 100) / 100;
}

// ── Sheet helpers ─────────────────────────────────────────────────────────────

/**
 * Returns an existing sheet by name, or creates it if it doesn't exist.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} name
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/**
 * Retrieves a required value from Script Properties.
 * Throws a descriptive error if the property is missing, which surfaces
 * clearly in the Apps Script execution log.
 * @param {string} key
 * @returns {string}
 */
function getRequiredProperty_(key) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) throw new Error('Missing script property: ' + key);
  return value;
}

// ── Date stripping ────────────────────────────────────────────────────────────

/**
 * Strips the time component from a Date, returning midnight local time.
 * Uses the spreadsheet timezone via Utilities.formatDate to avoid DST issues.
 * @param {Date} dateObj
 * @returns {Date}
 */
function stripTime_(dateObj) {
  return new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
}

/**
 * Same as stripTime_ but uses local JS Date arithmetic (safe for weekly math).
 * @param {Date} dateObj
 * @returns {Date}
 */
function stripTimeLocal_(dateObj) {
  return new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
}

/**
 * Adds `days` to a date and returns a new date at midnight.
 * @param {Date}   dateObj
 * @param {number} days
 * @returns {Date}
 */
function addDaysLocal_(dateObj, days) {
  const d = new Date(dateObj);
  d.setDate(d.getDate() + days);
  return stripTimeLocal_(d);
}

// ── WoW delta ─────────────────────────────────────────────────────────────────

/**
 * Computes a week-over-week delta. Returns '' if either value is missing or
 * the previous value is zero (avoids divide-by-zero confusion in display).
 * @param {number|string} currentValue
 * @param {number|string} previousValue
 * @returns {number|string}
 */
function wowDelta_(currentValue, previousValue) {
  if (currentValue === '' || previousValue === '' || previousValue === 0 || previousValue === null) return '';
  return round2_(currentValue - previousValue);
}