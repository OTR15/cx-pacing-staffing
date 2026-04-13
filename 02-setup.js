// =============================================================================
// setup.gs
// One-time and admin functions: menu registration, sheet seeding,
// tab visibility, and the Team Guide builder.
// =============================================================================

// ── Menu ──────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Pacing Report')
    .addItem('Fix Sheet Now', 'fixSheetNow')
    .addItem('Rebuild Yesterday + Today', 'fixYesterdayAndToday')
    .addItem('Refresh Today Only', 'refreshTodayOnly')
    .addSeparator()
    .addItem('Suggest Next Schedule Tab', 'autofillNextScheduleTab')
    .addSeparator()
    .addItem('Show Admin Tabs', 'unhideUtilitySheets')
    .addItem('Hide Admin Tabs', 'hideUtilitySheetsMenu_')
    .addItem('Open Team Guide', 'buildTeamGuideTab')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu('Admin Tools')
        .addItem('Normalize Schedule', 'normalizeCurrentWeekSchedule')
        .addItem('Build Today Tab', 'createTodayTab')
        .addItem('Publish Current Checkpoint', 'publishCurrentCheckpoint')
        .addItem('Publish Full Day Test', 'testFullDayTodaySlow')
        .addItem('Install Daily Triggers', 'installTriggers')
        .addItem('Remove Triggers', 'removePacingTriggers')
        .addItem('Seed Project', 'seedPrototypeSetup')
        .addItem('Setup Checklist', 'showSetupChecklist')
        .addItem('Daily Use Guide', 'showDailyUseGuide')
    )
    .addToUi();
}
// ── Seed / first-run setup ────────────────────────────────────────────────────

/**
 * Creates all required admin sheets with default values if they don't exist.
 * Safe to re-run — existing sheets are rebuilt, not duplicated.
 */
function seedPrototypeSetup() {
  const ss = SpreadsheetApp.getActive();
  ensureConfigSheet_(ss);
  ensureRosterSheet_(ss);
  ensureGoalsSheet_(ss);
  ensureNormalizedScheduleSheet_(ss);
  buildTeamGuideTab();
  hideUtilitySheets_(ss);

  SpreadsheetApp.getUi().alert(
    'Pacing project seeded.\n\n' +
    'Next steps:\n' +
    '1) Setup: Show Admin Tabs\n' +
    '2) Confirm Schedule tab name in Config\n' +
    '3) Run: Normalize Schedule\n' +
    '4) Run: Build Today Tab\n' +
    '5) Automation: Install Daily Triggers'
  );
}

// ── Sheet builders ────────────────────────────────────────────────────────────

/** Rebuilds the Config sheet with all default key/value pairs. */
function ensureConfigSheet_(ss) {
  const sh = getOrCreateSheet_(ss, CFG.configSheetName);
  sh.clear();

  const values = [
    ['Key',                          'Value'],
    ['SUBDOMAIN',                    CFG.subdomain],
    ['TIMEZONE',                     CFG.timezone],
    ['STANDARD_SHIFT_HOURS',         CFG.standardShiftHours],
    ['USE_SHIFT_BASED_GOALS',        true],
    ['SCHEDULE_SHEET_NAME',          CFG.scheduleSheetName],
    ['CURRENT_SCHEDULE_TAB',         '3/22 - 3/28'],
    ['NEXT_SCHEDULE_TAB',            ''],
    ['AUTO_LUNCH_DEDUCTION_HOURS',   1],
    ['AUTO_LUNCH_THRESHOLD_HOURS',   9],
    ['SHOW_DAILY_TABS_DAYS',         7],
    ['PACING_MODE',                  'percent'],
    ['PACING_GREEN_MIN',             1.0],
    ['PACING_YELLOW_MIN',            0.9],
    ['PACING_GREEN_MAX_SHORTFALL',   0],
    ['PACING_YELLOW_MAX_SHORTFALL',  5]
  ];

  sh.getRange(1, 1, values.length, 2).setValues(values);
  sh.getRange(1, 1, 1, 2).setFontWeight('bold');
  sh.autoResizeColumns(1, 2);
}

/** Rebuilds the Roster sheet from DEFAULT_ROSTER. */
function ensureRosterSheet_(ss) {
  const sh = getOrCreateSheet_(ss, CFG.rosterSheetName);
  sh.clear();
  sh.getRange(1, 1, 1, 2).setValues([['Agent ID', 'Agent Name']]);
  sh.getRange(2, 1, DEFAULT_ROSTER.length, 2).setValues(DEFAULT_ROSTER);
  sh.getRange(1, 1, 1, 2).setFontWeight('bold');
  sh.autoResizeColumns(1, 2);
}

/** Rebuilds the Goals sheet with default goals, checkpoints, and an empty override table. */
function ensureGoalsSheet_(ss) {
  const sh = getOrCreateSheet_(ss, CFG.goalsSheetName);
  sh.clear();

  // Row 1–3: team default goals
  sh.getRange(1, 1, 3, 5).setValues([
    ['Team Default Goals', '', '', '', ''],
    ['Closed Tickets', 'Tickets Replied', 'Messages Sent', 'CSAT', 'Use Shift Based Goals'],
    [
      CFG.baselineGoals.closedTickets,
      CFG.baselineGoals.ticketsReplied,
      CFG.baselineGoals.messagesSent,
      CFG.baselineGoals.csat,
      true
    ]
  ]);

  // Row 5+: checkpoint definitions
  sh.getRange(5, 1, CFG.checkpoints.length + 1, 6).setValues([
    ['Checkpoint', 'Label', 'Hour', 'Minute', 'Pacing %', 'Track'],
    ...CFG.checkpoints.map(cp => [cp.key, cp.label, cp.hour, cp.minute, cp.percent, true])
  ]);

  // Row 14+: per-rep override table (header + one blank row)
  sh.getRange(14, 1, 2, 5).setValues([
    ['Rep Name', 'Closed Tickets', 'Tickets Replied', 'Messages Sent', 'CSAT'],
    ['', '', '', '', '']
  ]);

  sh.getRange(1, 1, 1, 5).merge();
  sh.getRange(1, 1).setHorizontalAlignment('center').setFontWeight('bold');
  sh.getRange(5, 1, 1, 6).setFontWeight('bold');
  sh.getRange(14, 1, 1, 5).setFontWeight('bold');
  sh.autoResizeColumns(1, 6);
}

/** Rebuilds the Schedule_Normalized sheet with the correct header row. */
function ensureNormalizedScheduleSheet_(ss) {
  const sh = getOrCreateSheet_(ss, CFG.normalizedScheduleSheetName);
  sh.clear();
  sh.getRange(1, 1, 1, 10).setValues([[
    'Date', 'Agent Name', 'Manager', 'Start', 'End', 'Hours',
    'Status', 'In Office', 'Working Lunch', 'Raw Value'
  ]]);
  sh.getRange(1, 1, 1, 10).setFontWeight('bold');
  sh.autoResizeColumns(1, 10);
}

// ── Tab visibility ────────────────────────────────────────────────────────────

/** Menu wrapper — hides admin tabs. */
function hideUtilitySheetsMenu_() {
  hideUtilitySheets_(SpreadsheetApp.getActive());
}

/** Hides all sheets listed in CFG.hiddenSheetNames. */
function hideUtilitySheets_(ss) {
  CFG.hiddenSheetNames.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && !sh.isSheetHidden()) sh.hideSheet();
  });
}

/** Shows all sheets listed in CFG.hiddenSheetNames. */
function unhideUtilitySheets() {
  const ss = SpreadsheetApp.getActive();
  CFG.hiddenSheetNames.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) sh.showSheet();
  });
}

// ── Team Guide tab ────────────────────────────────────────────────────────────

/** Rebuilds the Team Guide tab with usage documentation. */
function buildTeamGuideTab() {
  const ss = SpreadsheetApp.getActive();
  const sh = getOrCreateSheet_(ss, CFG.teamGuideSheetName);
  sh.clear();

  const rows = [
    ['Gorgias Pacing Report — Team Guide'],
    [''],
    ['Purpose'],
    ['This sheet tracks daily pacing for the team using Gorgias stats and the team schedule.'],
    ['It creates a daily tab, adjusts goals by shift length, colors each metric, and marks reps as exempt when appropriate.'],
    [''],
    ['Main Features'],
    ['• Daily tabs created automatically'],
    ['• Checkpoints: 7 AM, 9 AM, 11 AM, 2 PM, 6 PM, EOD'],
    ['• Metrics: Closed Tickets, Tickets Replied, Messages Sent, CSAT'],
    ['• Shift-adjusted goals'],
    ['• Lunch deduction rule'],
    ['• Red / Yellow / Green metric coloring'],
    ['• Off / CTO / VTO reps marked Exempt'],
    ['• Last 7 daily tabs stay visible; older daily tabs auto-hide'],
    [''],
    ['Daily Use'],
    ['1. Open todays tab'],
    ['2. Review metric colors and On Track column'],
    ['3. If needed, use the Pacing Report menu to Normalize Schedule or Publish Current Checkpoint'],
    [''],
    ['Admin / Setup Menu'],
    ['Setup: Seed Project → creates Config, Roster, Goals, Schedule_Normalized, and this guide'],
    ['Setup: Show Admin Tabs → unhides helper tabs'],
    ['Setup: Hide Admin Tabs → hides helper tabs'],
    ['Run: Normalize Schedule → refreshes machine-readable schedule data'],
    ['Run: Build Today Tab → rebuilds todays tab'],
    ['Run: Publish Current Checkpoint → fills the current time block only'],
    ['Run: Publish Full Day Test → testing only, heavier API usage'],
    ['Automation: Install Daily Triggers → installs daily automation'],
    ['Automation: Remove Triggers → removes all pacing triggers'],
    [''],
    ['Onboarding / Offboarding'],
    ['Roster changes do NOT require code changes.'],
    ['To add someone: go to the Roster tab and add Agent ID + Agent Name on a new row.'],
    ['To remove someone: delete their row from the Roster tab.'],
    ['The script reads the Roster tab directly.'],
    ['Important: the name should match how the rep appears in the schedule or be close enough for first-name matching.'],
    [''],
    ['Goals'],
    ['Goals are editable in the Goals tab.'],
    ['You can change team defaults and rep-specific overrides without editing code.'],
    ['Checkpoint pacing percentages are also controlled in the Goals tab.'],
    [''],
    ['Schedule Source'],
    ['The Schedule tab is the source of truth.'],
    ['Recommended layout: names in column A, managers in column C, dates in row 2, daily shift values starting in column D.'],
    ['You can manually update this tab weekly or use IMPORTRANGE from another spreadsheet.'],
    ['IMPORTRANGE is recommended if the source schedule is stable and permissioned correctly.'],
    [''],
    ['Lunch / Effective Hours'],
    ['By default, scheduled shifts of 9+ hours subtract 1 hour for lunch.'],
    ['Examples: 9 scheduled = 8 effective, 11 scheduled = 10 effective.'],
    ['This is controlled in the Config tab.'],
    [''],
    ['Color Logic'],
    ['Each metric cell is colored separately.'],
    ['Green = on target'],
    ['Yellow = near target'],
    ['Red = behind target'],
    ['Default percent thresholds are controlled in Config.'],
    [''],
    ['Tab Visibility / Archiving'],
    ['Daily tabs from the last 7 days remain visible by default.'],
    ['Older daily tabs are automatically hidden each day.'],
    ['This keeps the workbook clean without deleting history.'],
    [''],
    ['Triggers / Automation'],
    ['Triggers run as the account that installs them.'],
    ['The primary admin should install and own the triggers.'],
    ['Once installed, most users do not need Apps Script access.'],
    [''],
    ['Troubleshooting'],
    ['If todays tab is wrong:'],
    ['1. Run Normalize Schedule'],
    ['2. Run Build Today Tab'],
    ['3. Run Publish Current Checkpoint'],
    ['If metrics are blank, confirm script properties and Roster IDs.'],
    ['If reps are missing or Off incorrectly, check the Schedule tab and todays date in row 2.']
  ];

  sh.getRange(1, 1, rows.length, 1).setValues(rows);
  sh.setColumnWidth(1, 700);
  sh.getRange(1, 1).setFontSize(14).setFontWeight('bold');
}

// ── Help dialogs ──────────────────────────────────────────────────────────────

function showSetupChecklist() {
  SpreadsheetApp.getUi().alert(
    'Setup Checklist\n\n' +
    '1. Setup: Seed Project\n' +
    '2. Setup: Show Admin Tabs\n' +
    '3. Confirm Schedule tab name in Config\n' +
    '4. Run: Normalize Schedule\n' +
    '5. Run: Build Today Tab\n' +
    '6. Help: Build Team Guide Tab\n' +
    '7. Automation: Install Daily Triggers'
  );
}

function showDailyUseGuide() {
  SpreadsheetApp.getUi().alert(
    'Daily Use Guide\n\n' +
    'Normalize Schedule = refresh source schedule\n' +
    'Build Today Tab = rebuild today tab\n' +
    'Publish Current Checkpoint = fill the current time block\n' +
    'Publish Full Day Test = testing only, heavier API usage'
  );
}