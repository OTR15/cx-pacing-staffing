// =============================================================================
// roster.gs
// Agent roster loading and rep-to-manager / rep-to-schedule resolution.
//
// The Roster sheet is the source of truth for agent IDs and display names.
// Name matching uses a two-step strategy: full normalized name first,
// then first-name-only as a fallback.
// =============================================================================

/**
 * Loads all active agents from the Roster sheet, sorted alphabetically.
 * Filters out excluded agents (bots, test accounts) defined in CFG.excludedAgents.
 * Deduplicates by normalized name in case of accidental duplicate rows.
 *
 * @returns {Array<{ agentId: number, repName: string }>}
 */
function getDisplayRoster_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.rosterSheetName);

  if (!sh) throw new Error('Roster sheet missing. Run Setup: Seed Project first.');

  const seen = {};
  const rows = sh.getDataRange().getValues()
    .slice(1) // skip header
    .filter(row => row[0] && row[1])
    .map(row => ({ agentId: Number(row[0]), repName: String(row[1]).trim() }))
    .filter(rep => !CFG.excludedAgents.includes(rep.repName))
    .filter(rep => {
      const key = normalizeName_(rep.repName);
      if (seen[key]) return false;
      seen[key] = true;
      return true;
    });

  return rows.sort((a, b) => a.repName.localeCompare(b.repName));
}

// ── Manager resolution ────────────────────────────────────────────────────────

/**
 * Builds a name → manager map from the Schedule sheet.
 * Each agent is stored under both full normalized name and first name
 * to support the same fuzzy matching used elsewhere.
 *
 * @returns {Object} e.g. { 'adam steen': 'Quint', 'adam': 'Quint', ... }
 */
function getManagerMapFromSchedule_() {
  const ss = SpreadsheetApp.getActive();
  const scheduleSheet = ss.getSheetByName(
    getConfigValue_('SCHEDULE_SHEET_NAME', CFG.scheduleSheetName)
  );
  if (!scheduleSheet) return {};

  const values = scheduleSheet.getDataRange().getDisplayValues();
  const map    = {};

  for (let r = CFG.schedule.firstDataRow - 1; r < values.length; r++) {
    const name    = String(values[r][CFG.schedule.nameCol - 1]    || '').trim();
    const manager = String(values[r][CFG.schedule.managerCol - 1] || '').trim();
    if (!name) continue;

    map[normalizeName_(name)]     = manager;
    map[normalizeFirstName_(name)] = manager;
  }

  return map;
}

/**
 * Looks up a rep's manager from the map built by getManagerMapFromSchedule_().
 * Tries full name first, then first name only.
 *
 * @param {Object} managerMap - From getManagerMapFromSchedule_()
 * @param {string} repName
 * @returns {string} Manager name, or '' if not found.
 */
function getManagerForRep_(managerMap, repName) {
  return managerMap[normalizeName_(repName)] ||
         managerMap[normalizeFirstName_(repName)] ||
         '';
}

// ── Schedule resolution ───────────────────────────────────────────────────────

/**
 * Looks up a rep's schedule entry for today from the schedule map.
 * Tries full name first, then first name only.
 * Falls back to defaultSchedule_() if neither matches.
 *
 * @param {Object} scheduleMap - From getScheduleMapForDate_()
 * @param {string} repName
 * @returns {{ status, hours, startText, endText, inOffice, workingLunch }}
 */
function getScheduleForRep_(scheduleMap, repName) {
  return scheduleMap[normalizeName_(repName)] ||
         scheduleMap[normalizeFirstName_(repName)] ||
         defaultSchedule_();
}