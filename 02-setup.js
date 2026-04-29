// =============================================================================
// setup.gs
// One-time and admin functions: menu registration, sheet seeding,
// tab visibility, and the Team Guide builder.
// =============================================================================

// ── Menu ──────────────────────────────────────────────────────────────────────

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Pacing Report')
    .addItem('Publish Current Checkpoint',       'publishCurrentCheckpoint')
    .addItem('Rebuild Current Day',              'rebuildAndRepublishToday')
    .addItem('Apply Goal Adjustments',                'applyGoalAdjustments')
    .addItem('Migrate Tab to Goal Adjustments',       'migrateTabToGoalAdjustments')
    .addItem('Migrate All Tabs to Goal Adjustments',  'migrateAllTabsToGoalAdjustments')
    .addToUi();

  ui.createMenu('KPI Supervisor View')
    .addItem('Sort Active Tab by Manager',   'sortActiveDailySheetByManager')
    .addItem('Filter Active Tab to Manager', 'filterActiveDailySheetByManager')
    .addItem('Show All Managers',            'clearManagerFilterOnActiveDailySheet')
    .addToUi();

  ui.createMenu('Mode')
    .addItem('Internal Mode', 'setInternalMode')
    .addItem('External Mode', 'setExternalMode')
    .addToUi();

  ui.createMenu('Admin')
    .addItem('Show Admin Tabs',         'unhideUtilitySheets')
    .addItem('Hide Admin Tabs',         'hideUtilitySheetsMenu_')
    .addSeparator()
    .addItem('Install Daily Triggers',  'installTriggers')
    .addItem('Remove Daily Triggers',   'removePacingTriggers')
    .addSeparator()
    .addItem('Set Next Schedule Tab',   'promptSetNextScheduleTab')
    .addItem('Normalize Schedule',      'normalizeCurrentWeekSchedule')
    .addItem('Rebuild Weekly Tabs',     'rebuildAllWeeklyReports')
    .addSeparator()
    .addItem('Organize Tabs',           'organizeTabs')
    .addToUi();

  ui.createMenu('Help')
    .addItem('Daily Use Guide',           'showDailyUseGuide')
    .addItem('Setup Checklist',           'showSetupChecklist')
    .addItem('Build Team Guide Tab',      'buildTeamGuideTab')
    .addItem('Build Case Use Summary Tab','buildCaseUseSummaryTab')
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
    ['PACING_YELLOW_MAX_SHORTFALL',  5],
    ['STAFFING_SHEET_NAME',                      CFG.staffing.sheetName],
    ['STAFFING_TICKETS_PER_PRODUCTIVE_HOUR',     CFG.staffing.ticketsPerProductiveHour],
    ['STAFFING_AGED_RISK_WEIGHT',                CFG.staffing.agedRiskWeight],
    ['STAFFING_RESERVE_HOURS_BUFFER',            CFG.staffing.reserveHoursBuffer],
    ['STAFFING_MINIMUM_AGENTS_FLOOR',            CFG.staffing.minimumAgentsFloor],
    ['STAFFING_CAUTION_UNASSIGNED_THRESHOLD',    CFG.staffing.cautionUnassignedThreshold],
    ['STAFFING_ESTIMATED_WORKABLE_TICKETS_PER_HOUR', CFG.staffing.estimatedWorkableTicketsPerHour],
    ['STAFFING_END_OF_DAY_HOUR',                 CFG.staffing.endOfDayHour],
    ['STAFFING_PULSE_LOG_SPREADSHEET_ID',        CFG.staffing.pulseLogSpreadsheetId],
    ['STAFFING_WORKABLE_VOLUME_LOG_SHEET_NAME',  CFG.staffing.workableVolumeLogSheetName],
    ['STAFFING_OVERNIGHT_INFLOW_LOG_SHEET_NAME', CFG.staffing.overnightInflowLogSheetName],
    ['STAFFING_OBSERVED_DATA_BLEND_WEIGHT',      CFG.staffing.observedDataBlendWeight],
    ['STAFFING_USE_OBSERVED_DATA',               CFG.staffing.useObservedData],
    ['STAFFING_SHADOW_MODEL_ENABLED',            CFG.staffing.shadowModelEnabled],
    ['STAFFING_MINIMUM_OBSERVED_SAMPLE_DAYS',    CFG.staffing.minimumObservedSampleDays],
    ['STAFFING_WORKABLE_RATE_MULTIPLIER',        CFG.staffing.workableRateMultiplier],
    ['STAFFING_SEND_HOME_BUFFER_MULTIPLIER',     CFG.staffing.sendHomeBufferMultiplier],
    ['QA_LEAD_NAME',                             ''],
    ['QA_LEAD_REPORT_CARD_ID',                   ''],
    ['TEAM_TCPH_GOAL',                           ''],
    ['TEAM_TRPH_GOAL',                           ''],
    ['TEAM_QA_GOAL',                             ''],
    ['TEAM_CSAT_GOAL',                           '']
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
    ['1. Open today tab'],
    ['2. Review metric colors and On Track column'],
    ['3. If needed, use the Pacing Report menu to Normalize Schedule or Publish Current Checkpoint'],
    [''],
    ['Admin / Setup Menu'],
    ['Setup: Seed Project → creates Config, Roster, Goals, Schedule_Normalized, and this guide'],
    ['Setup: Show Admin Tabs → unhides helper tabs'],
    ['Setup: Hide Admin Tabs → hides helper tabs'],
    ['Run: Normalize Schedule → refreshes machine-readable schedule data'],
    ['Run: Build Today Tab → rebuilds today tab'],
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
    ['If today tab is wrong:'],
    ['1. Run Normalize Schedule'],
    ['2. Run Build Today Tab'],
    ['3. Run Publish Current Checkpoint'],
    ['If metrics are blank, confirm script properties and Roster IDs.'],
    ['If reps are missing or Off incorrectly, check the Schedule tab and today date in row 2.']
  ];

  sh.getRange(1, 1, rows.length, 1).setValues(rows);
  sh.setColumnWidth(1, 700);
  sh.getRange(1, 1).setFontSize(14).setFontWeight('bold');
}

// ── Case Use Summary tab ──────────────────────────────────────────────────────

function buildCaseUseSummaryTab() {
  const ss = SpreadsheetApp.getActive();
  const sh = getOrCreateSheet_(ss, CFG.caseUseSummarySheetName);
  sh.clear();
  sh.setTabColor('#1a237e');
  sh.setColumnWidth(1, 210);
  sh.setColumnWidth(2, 510);
  sh.setColumnWidth(3, 250);

  const NAVY        = '#1a237e';
  const NAVY_MID    = '#283593';
  const NAVY_LIGHT  = '#e8eaf6';
  const WHITE       = '#ffffff';
  const INK         = '#212121';
  const GRAY        = '#555555';

  let r = 1;

  // Title
  sh.setRowHeight(r, 46);
  sh.getRange(r, 1, 1, 3).merge()
    .setValue('Pacing Report  —  Data Summary')
    .setBackground(NAVY).setFontColor(WHITE)
    .setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  r++;

  sh.setRowHeight(r, 10); r++;

  // ── Section 1: About This Report ─────────────────────────────────────────
  sh.setRowHeight(r, 28);
  sh.getRange(r, 1, 1, 3).merge()
    .setValue('About This Report')
    .setBackground(NAVY_MID).setFontColor(WHITE)
    .setFontSize(11).setFontWeight('bold').setVerticalAlignment('middle');
  r++;

  const about = [
    'The Pacing Report tracks how the CX team is performing against daily and weekly goals. ' +
    'It pulls ticket data from Gorgias and cross-references the team schedule to account for ' +
    'actual hours worked — so goals scale with each agent\'s shift rather than assuming ' +
    'everyone works the same hours.',

    'Reports update automatically at six checkpoints throughout the day (7AM, 9AM, 11AM, 2PM, ' +
    '6PM, and EOD). Each checkpoint shows how the team is tracking relative to where they should ' +
    'be at that point in the day, with color-coded pacing for each metric.',

    'Weekly summaries roll up each agent\'s full week and surface it alongside quality scores ' +
    'and customer satisfaction averages. The Team Dashboard brings all of it into one view ' +
    'for quick leadership reference.'
  ];

  about.forEach(text => {
    sh.setRowHeight(r, 52);
    sh.getRange(r, 1, 1, 3).merge()
      .setValue(text)
      .setFontSize(10).setFontColor(INK).setBackground(WHITE)
      .setWrap(true).setVerticalAlignment('middle');
    r++;
  });

  sh.setRowHeight(r, 12); r++;

  // ── Section 2: Data Sources ───────────────────────────────────────────────
  sh.setRowHeight(r, 28);
  sh.getRange(r, 1, 1, 3).merge()
    .setValue('Data Sources')
    .setBackground(NAVY_MID).setFontColor(WHITE)
    .setFontSize(11).setFontWeight('bold').setVerticalAlignment('middle');
  r++;

  sh.setRowHeight(r, 24);
  sh.getRange(r, 1, 1, 2)
    .setValues([['Source', 'What It Provides']])
    .setFontWeight('bold').setBackground(NAVY_LIGHT).setFontColor(INK)
    .setVerticalAlignment('middle');
  sh.getRange(r, 3).setBackground(NAVY_LIGHT);
  r++;

  [
    ['Gorgias',
     'Closed tickets, replies, and CSAT scores for each agent — pulled via API at each daily checkpoint.'],
    ['Team Schedule',
     'Shift hours and status (active, CTO, VTO, off). Used to calculate each agent\'s adjusted goal for the day. Agents on shorter shifts get proportionally smaller targets.'],
    ['QA Lead Report Card',
     'Weekly QA form completion for the QA Lead, tracked in a separate spreadsheet and pulled automatically each time the weekly report is built.'],
    ['Pulse Log',
     'End-of-day snapshots of inbox health: total open tickets, unassigned count, and a breakdown of how long tickets have been waiting (aging buckets from under 8 hours to over 24 hours).']
  ].forEach(([source, desc], i) => {
    sh.setRowHeight(r, 46);
    const bg = i % 2 === 0 ? WHITE : NAVY_LIGHT;
    sh.getRange(r, 1).setValue(source)
      .setFontWeight('bold').setFontSize(10).setFontColor(INK)
      .setBackground(bg).setWrap(true).setVerticalAlignment('middle');
    sh.getRange(r, 2).setValue(desc)
      .setFontSize(10).setFontColor(INK)
      .setBackground(bg).setWrap(true).setVerticalAlignment('middle');
    sh.getRange(r, 3).setBackground(bg);
    r++;
  });

  sh.setRowHeight(r, 12); r++;

  // ── Section 3: Metric Definitions ────────────────────────────────────────
  sh.setRowHeight(r, 28);
  sh.getRange(r, 1, 1, 3).merge()
    .setValue('Metric Definitions')
    .setBackground(NAVY_MID).setFontColor(WHITE)
    .setFontSize(11).setFontWeight('bold').setVerticalAlignment('middle');
  r++;

  sh.setRowHeight(r, 24);
  sh.getRange(r, 1, 1, 3)
    .setValues([['Metric', 'What It Measures', 'How the Goal Is Set']])
    .setFontWeight('bold').setBackground(NAVY_LIGHT).setFontColor(INK)
    .setVerticalAlignment('middle');
  r++;

  [
    ['TCPH\n(Tickets Closed Per Hour)',
     'How many tickets an agent closes per hour of effective work time. The primary throughput metric.',
     'Daily closed ticket target ÷ shift hours'],
    ['TRPH\n(Tickets Replied Per Hour)',
     'How many tickets an agent replies to per hour. Measures responsiveness alongside closure rate.',
     'Daily reply target ÷ shift hours'],
    ['QA Score',
     'Average quality review score from ticket audits, expressed as a percentage. Reflects accuracy, tone, and process adherence.',
     'Team default: 90%'],
    ['CSAT',
     'Average customer satisfaction rating from post-ticket surveys, on a 1–5 scale. Direct feedback from customers.',
     'Set per agent in the Goals tab (default: 4.7)'],
    ['Overall %',
     'The average of each agent\'s four metric scores measured against their individual goals. The primary number used for weekly KPI status.',
     '—'],
    ['Pacing Color',
     'Whether an agent is on track at a given checkpoint, based on the share of expected output at that point in the day.',
     'Green = on track  ·  Yellow = slightly behind  ·  Red = behind']
  ].forEach(([metric, def, goal], i) => {
    sh.setRowHeight(r, 54);
    const bg = i % 2 === 0 ? WHITE : NAVY_LIGHT;
    sh.getRange(r, 1).setValue(metric)
      .setFontWeight('bold').setFontSize(10).setFontColor(INK)
      .setBackground(bg).setWrap(true).setVerticalAlignment('middle');
    sh.getRange(r, 2).setValue(def)
      .setFontSize(10).setFontColor(INK)
      .setBackground(bg).setWrap(true).setVerticalAlignment('middle');
    sh.getRange(r, 3).setValue(goal)
      .setFontSize(10).setFontColor(GRAY)
      .setBackground(bg).setWrap(true).setVerticalAlignment('middle');
    r++;
  });

  SpreadsheetApp.flush();
}

// ── New menu functions ────────────────────────────────────────────────────────

/**
 * Prompts for the next week's schedule tab name and saves it to Config.
 * Replaces the need to open the Config sheet manually.
 */
function promptSetNextScheduleTab() {
  const ui       = SpreadsheetApp.getUi();
  const current  = String(getConfigValue_('NEXT_SCHEDULE_TAB', '') || '').trim();
  const response = ui.prompt(
    'Set Next Schedule Tab',
    'Enter the name of next week schedule tab:\n(current value: "' + (current || 'none') + '")',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const value = response.getResponseText().trim();
  if (!value) {
    ui.alert('No value entered — Config was not updated.');
    return;
  }

  setConfigValue_('NEXT_SCHEDULE_TAB', value);
  ui.alert('Next Schedule Tab set to: "' + value + '"');
}

/**
 * Organizes all sheets into the preferred display order:
 *   1. Team Guide
 *   2. Daily tabs (newest first)
 *   3. Weekly tabs (newest first)
 *   4. Staffing tab
 *   5. Admin/hidden tabs (Config, Roster, Goals, Schedule_Normalized, Schedule)
 *
 * Tabs not matching any category are placed before admin tabs.
 */
function organizeTabs() {
  const ss     = SpreadsheetApp.getActive();
  const sheets = ss.getSheets();

  const adminNames   = ['Config', 'Roster', 'Goals', 'Schedule_Normalized', 'Schedule'];
  const staffingName = 'Staffing';
  const visibleDays  = Number(getConfigValue_('SHOW_DAILY_TABS_DAYS', 7));

  const teamGuide = sheets.filter(sh => sh.getName() === CFG.teamGuideSheetName);

  const daily = sheets
    .filter(sh => parseDailySheetName_(sh.getName()))
    .sort((a, b) => parseDailySheetName_(b.getName()) - parseDailySheetName_(a.getName()));

  const weekly = sheets
    .filter(sh => parseWeeklyTabDate_(sh.getName()))
    .sort((a, b) => parseWeeklyTabDate_(b.getName()) - parseWeeklyTabDate_(a.getName()));

  const staffing = sheets.filter(sh => sh.getName() === staffingName);
  const admin    = sheets.filter(sh => adminNames.includes(sh.getName()));
  const other    = sheets.filter(sh =>
    !parseDailySheetName_(sh.getName()) &&
    !parseWeeklyTabDate_(sh.getName()) &&
    sh.getName() !== CFG.teamGuideSheetName &&
    sh.getName() !== staffingName &&
    !adminNames.includes(sh.getName())
  );

  // Apply correct visibility before reordering
  daily.forEach((sh, i)  => i < visibleDays                  ? sh.showSheet() : sh.hideSheet());
  weekly.forEach((sh, i) => i < CFG.weekly.visibleTabCount   ? sh.showSheet() : sh.hideSheet());
  admin.forEach(sh  => sh.hideSheet());
  teamGuide.forEach(sh => sh.showSheet());
  staffing.forEach(sh  => sh.showSheet());

  const ordered = [...teamGuide, ...daily.slice(0, visibleDays), ...weekly.slice(0, CFG.weekly.visibleTabCount), ...staffing, ...other, ...admin];

  ordered.forEach((sh, i) => {
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(i + 1);
  });

  SpreadsheetApp.getUi().alert('Tabs organized successfully.');
}

// ── Tab modes ─────────────────────────────────────────────────────────────────

function setInternalMode() { applyTabMode_('internal'); }
function setExternalMode() { applyTabMode_('external'); }

/**
 * Applies Internal or External mode:
 *
 *   Internal — Team Guide · 7 daily · 4 weekly · Staffing
 *   External — Case Use Summary · Team Dashboard · 3 daily · 1 weekly
 *
 * Sets tab colors, shows/hides sheets, and reorders visible tabs.
 */
function applyTabMode_(mode) {
  const ss     = SpreadsheetApp.getActive();
  const sheets = ss.getSheets();

  const C_DAILY      = '#b6d7a8'; // light green
  const C_WEEKLY     = '#9fc5f8'; // light blue
  const C_HOT_PINK   = '#ff69b4';
  const C_DARK_BLUE  = '#1a237e';
  const C_PURPLE     = '#7b4f9e';

  const INTERNAL_DAYS  = 7;
  const EXTERNAL_DAYS  = 3;
  const INTERNAL_WEEKS = 4;
  const EXTERNAL_WEEKS = 1;

  // Categorise all sheets
  const daily = sheets
    .filter(sh => parseDailySheetName_(sh.getName()))
    .sort((a, b) => parseDailySheetName_(b.getName()) - parseDailySheetName_(a.getName()));

  const weekly = sheets
    .filter(sh => parseWeeklyTabDate_(sh.getName()))
    .sort((a, b) => parseWeeklyTabDate_(b.getName()) - parseWeeklyTabDate_(a.getName()));

  const teamGuide  = ss.getSheetByName(CFG.teamGuideSheetName);
  const caseUse    = ss.getSheetByName(CFG.caseUseSummarySheetName);
  const teamDash   = ss.getSheetByName(DASH_TAB_NAME);
  const staffing   = ss.getSheetByName(CFG.staffing.sheetName);
  const adminSheets = CFG.hiddenSheetNames.map(n => ss.getSheetByName(n)).filter(Boolean);

  const knownNames = new Set([
    CFG.teamGuideSheetName, CFG.caseUseSummarySheetName,
    DASH_TAB_NAME, CFG.staffing.sheetName,
    ...CFG.hiddenSheetNames
  ]);
  const other = sheets.filter(sh =>
    !parseDailySheetName_(sh.getName()) &&
    !parseWeeklyTabDate_(sh.getName()) &&
    !knownNames.has(sh.getName())
  );

  // Apply tab colors
  daily.forEach( sh => sh.setTabColor(C_DAILY));
  weekly.forEach(sh => sh.setTabColor(C_WEEKLY));
  if (teamGuide) teamGuide.setTabColor(C_DARK_BLUE);
  if (caseUse)   caseUse.setTabColor(C_DARK_BLUE);
  if (teamDash)  teamDash.setTabColor(C_PURPLE);
  if (staffing)  staffing.setTabColor(C_HOT_PINK);

  // Show / hide
  const visibleDays  = mode === 'internal' ? INTERNAL_DAYS  : EXTERNAL_DAYS;
  const visibleWeeks = mode === 'internal' ? INTERNAL_WEEKS : EXTERNAL_WEEKS;

  daily.forEach( (sh, i) => i < visibleDays  ? sh.showSheet() : sh.hideSheet());
  weekly.forEach((sh, i) => i < visibleWeeks ? sh.showSheet() : sh.hideSheet());
  adminSheets.forEach(sh => sh.hideSheet());
  other.forEach(      sh => sh.hideSheet());

  if (mode === 'internal') {
    if (teamGuide) teamGuide.showSheet();
    if (caseUse)   caseUse.hideSheet();
    if (teamDash)  teamDash.hideSheet();
    if (staffing)  staffing.showSheet();
  } else {
    if (teamGuide) teamGuide.hideSheet();
    if (caseUse)   caseUse.showSheet();
    if (teamDash)  teamDash.showSheet();
    if (staffing)  staffing.hideSheet();
  }

  // Reorder visible tabs
  const ordered = mode === 'internal'
    ? [teamGuide, ...daily.slice(0, visibleDays), ...weekly.slice(0, visibleWeeks), staffing]
    : [caseUse,   teamDash, ...daily.slice(0, visibleDays), ...weekly.slice(0, visibleWeeks)];

  ordered.filter(Boolean).forEach((sh, i) => {
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(i + 1);
  });

  SpreadsheetApp.getUi().alert(
    (mode === 'internal' ? 'Internal' : 'External') + ' Mode applied.\n\n' +
    (mode === 'internal'
      ? 'Visible: Team Guide  ·  ' + visibleDays + ' daily tabs  ·  ' + visibleWeeks + ' weekly tabs  ·  Staffing'
      : 'Visible: Case Use Summary  ·  Team Dashboard  ·  ' + visibleDays + ' daily tabs  ·  ' + visibleWeeks + ' weekly tab')
  );
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
