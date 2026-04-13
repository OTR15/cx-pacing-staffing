// =============================================================================
// automation.gs
// Trigger installation, daily tab management, and checkpoint trigger wrappers.
//
// All triggers should be installed by the primary admin account.
// Triggers run under the installing account's credentials.
// =============================================================================

// ── Trigger management ────────────────────────────────────────────────────────

/**
 * Removes all existing pacing triggers and installs a fresh full set.
 * Safe to re-run — old triggers are always cleared first.
 *
 * Schedule:
 *   12:00 AM  createTodayTab               — create new daily tab
 *    1:00 AM  manageDailyTabs              — hide old tabs
 *    6:55 AM  normalizeCurrentWeekSchedule — refresh schedule before first checkpoint
 *    7:00 AM  publish7AM
 *    9:00 AM  publish9AM
 *   11:00 AM  publish11AM
 *    2:00 PM  publish2PM
 *    6:00 PM  publish6PM
 *   11:00 PM  publishEOD (also triggers weekly report + schedule rollover on Sat/Sun)
 */
function installTriggers() {
  removePacingTriggers();

  ScriptApp.newTrigger('createTodayTab')                 .timeBased().atHour(0) .everyDays(1).create();
  ScriptApp.newTrigger('manageDailyTabs')                .timeBased().atHour(1) .everyDays(1).create();
  ScriptApp.newTrigger('normalizeCurrentWeekSchedule')   .timeBased().atHour(6) .nearMinute(55).everyDays(1).create();
  ScriptApp.newTrigger('publish7AM')                     .timeBased().atHour(7) .nearMinute(0) .everyDays(1).create();
  ScriptApp.newTrigger('publish9AM')                     .timeBased().atHour(9) .nearMinute(0) .everyDays(1).create();
  ScriptApp.newTrigger('publish11AM')                    .timeBased().atHour(11).nearMinute(0) .everyDays(1).create();
  ScriptApp.newTrigger('publish2PM')                     .timeBased().atHour(14).nearMinute(0) .everyDays(1).create();
  ScriptApp.newTrigger('publish6PM')                     .timeBased().atHour(18).nearMinute(0) .everyDays(1).create();
  ScriptApp.newTrigger('publishEOD')                     .timeBased().atHour(23).nearMinute(0) .everyDays(1).create();

  SpreadsheetApp.getUi().alert(
    'Daily triggers installed.\n\n' +
    'They will run as the account that installed them.'
  );
}

/**
 * Removes ALL project triggers.
 * Called by installTriggers() before reinstalling, and available as a menu item.
 */
function removePacingTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));
  try {
    SpreadsheetApp.getUi().alert('All pacing triggers were removed.');
  } catch (e) {
    // Silently continue when called from installTriggers (not from menu)
  }
}

// ── Daily tab management ──────────────────────────────────────────────────────

/**
 * Keeps only the most recent N daily tabs visible (N = SHOW_DAILY_TABS_DAYS in Config).
 * Older tabs are hidden (not deleted) to preserve history.
 * Runs automatically at 1 AM each day.
 */
function manageDailyTabs() {
  const ss          = SpreadsheetApp.getActive();
  const visibleDays = Number(getConfigValue_('SHOW_DAILY_TABS_DAYS', 7));

  const dailySheets = ss.getSheets()
    .map(sh => ({ sheet: sh, name: sh.getName(), date: parseDailySheetName_(sh.getName()) }))
    .filter(x => x.date)
    .sort((a, b) => b.date.getTime() - a.date.getTime()); // newest first

  dailySheets.forEach((item, index) => {
    if (index < visibleDays) {
      item.sheet.showSheet();
    } else {
      item.sheet.hideSheet();
    }
  });
}

// ── Checkpoint trigger wrappers ───────────────────────────────────────────────
// These thin wrappers exist because Apps Script time-based triggers must
// point to named top-level functions. They delegate to publishCheckpointForDate_().

function publish7AM()  { publishCheckpointForDate_(new Date(), CFG.checkpoints[0]); }
function publish9AM()  { publishCheckpointForDate_(new Date(), CFG.checkpoints[1]); }
function publish11AM() { publishCheckpointForDate_(new Date(), CFG.checkpoints[2]); }
function publish2PM()  { publishCheckpointForDate_(new Date(), CFG.checkpoints[3]); }
function publish6PM()  { publishCheckpointForDate_(new Date(), CFG.checkpoints[4]); }

/**
 * EOD trigger: publishes EOD checkpoint, then runs end-of-day housekeeping.
 * Housekeeping only executes on the appropriate day of week:
 *   - Saturday: schedule rollover (advance to next week's schedule tab)
 *   - Sunday:   build weekly report
 * manageDailyTabs is also run nightly from the 1 AM trigger, but calling it
 * after EOD ensures the new day's tab is visible immediately if needed.
 */
function publishEOD() {
  publishCheckpointForDate_(new Date(), CFG.checkpoints[5]);
  rolloverScheduleTabIfNeeded_();
  buildWeeklyReportIfSunday_(new Date());
  manageWeeklyTabs();
}

// ── Test / debug helpers ──────────────────────────────────────────────────────

/**
 * Runs a single checkpoint for today without waiting for the scheduled trigger.
 * @param {string} checkpointKey - e.g. '7AM', '2PM', 'EOD'
 */
function testCheckpointToday(checkpointKey) {
  const cp = CFG.checkpoints.find(c => c.key === checkpointKey);
  if (!cp) throw new Error('Invalid checkpoint key: ' + checkpointKey);
  publishCheckpointForDate_(new Date(), cp);
}

// Individual checkpoint test functions (callable from Apps Script editor)
function test7AMToday()  { testCheckpointToday('7AM');  }
function test9AMToday()  { testCheckpointToday('9AM');  }
function test11AMToday() { testCheckpointToday('11AM'); }
function test2PMToday()  { testCheckpointToday('2PM');  }
function test6PMToday()  { testCheckpointToday('6PM');  }
function testEODToday()  { testCheckpointToday('EOD');  }

/**
 * Full-day smoke test: rebuilds today's tab and runs all 6 checkpoints in order.
 * Heavy API usage — for testing only, not production.
 */
function testFullDayTodaySlow() {
  const keys = ['7AM', '9AM', '11AM', '2PM', '6PM', 'EOD'];
  createTodayTab();

  keys.forEach((key, i) => {
    testCheckpointToday(key);
    if (i < keys.length - 1) Utilities.sleep(3000);
  });

  SpreadsheetApp.getUi().alert('Full-day test completed.');
}

/**
 * Logs parsed dates for all daily tabs — useful for debugging
 * manageDailyTabs() if tabs are being incorrectly hidden/shown.
 */
function debugDailyTabs() {
  const ss      = SpreadsheetApp.getActive();
  const results = ss.getSheets().map(sh => ({
    name:   sh.getName(),
    parsed: parseDailySheetName_(sh.getName())
  }));
  Logger.log(JSON.stringify(results, null, 2));
}


function suggestNextScheduleTabName_(currentTabName) {
  const m = String(currentTabName || '').match(/^(\d{1,2})\/(\d{1,2})\s*-\s*(\d{1,2})\/(\d{1,2})$/);
  if (!m) return '';

  const nowYear = new Date().getFullYear();
  const startMonth = Number(m[1]) - 1;
  const startDay = Number(m[2]);

  const currentStart = new Date(nowYear, startMonth, startDay);
  const nextStart = new Date(currentStart);
  nextStart.setDate(currentStart.getDate() + 7);

  const nextEnd = new Date(nextStart);
  nextEnd.setDate(nextStart.getDate() + 6);

  return (nextStart.getMonth() + 1) + '/' + nextStart.getDate() +
    ' - ' +
    (nextEnd.getMonth() + 1) + '/' + nextEnd.getDate();
}

function autofillNextScheduleTab() {
  const currentTab = String(getConfigValue_('CURRENT_SCHEDULE_TAB', '') || '').trim();
  if (!currentTab) {
    SpreadsheetApp.getUi().alert('CURRENT_SCHEDULE_TAB is blank.');
    return;
  }

  const suggestion = suggestNextScheduleTabName_(currentTab);
  if (!suggestion) {
    SpreadsheetApp.getUi().alert('Could not parse CURRENT_SCHEDULE_TAB.');
    return;
  }

  setConfigValue_('NEXT_SCHEDULE_TAB', suggestion);
  SpreadsheetApp.getUi().alert('NEXT_SCHEDULE_TAB set to: ' + suggestion);
}