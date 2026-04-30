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
  sh.setTabColor('#2e7d32');
  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 530);

  const GREEN_DARK  = '#1b5e20';
  const GREEN_MID   = '#2e7d32';
  const GREEN_LIGHT = '#e8f5e9';
  const WHITE       = '#ffffff';
  const INK         = '#212121';

  let r = 1;

  const title_ = (text) => {
    sh.setRowHeight(r, 46);
    sh.getRange(r, 1, 1, 2).merge()
      .setValue(text)
      .setBackground(GREEN_DARK).setFontColor(WHITE)
      .setFontSize(16).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    r++;
  };

  const spacer_ = (h) => { sh.setRowHeight(r, h || 10); r++; };

  const sectionHeader_ = (text) => {
    sh.setRowHeight(r, 28);
    sh.getRange(r, 1, 1, 2).merge()
      .setValue(text)
      .setBackground(GREEN_MID).setFontColor(WHITE)
      .setFontSize(11).setFontWeight('bold').setVerticalAlignment('middle');
    r++;
  };

  const prose_ = (text, h) => {
    sh.setRowHeight(r, h || 52);
    sh.getRange(r, 1, 1, 2).merge()
      .setValue(text)
      .setFontSize(10).setFontColor(INK).setBackground(WHITE)
      .setWrap(true).setVerticalAlignment('middle');
    r++;
  };

  const qa_ = (question, answer, h, even) => {
    sh.setRowHeight(r, h || 48);
    const bg = even ? GREEN_LIGHT : WHITE;
    sh.getRange(r, 1)
      .setValue(question)
      .setFontWeight('bold').setFontSize(10).setFontColor(INK)
      .setBackground(bg).setWrap(true).setVerticalAlignment('top');
    sh.getRange(r, 2)
      .setValue(answer)
      .setFontSize(10).setFontColor(INK)
      .setBackground(bg).setWrap(true).setVerticalAlignment('top');
    r++;
  };

  const colHeader_ = (a, b) => {
    sh.setRowHeight(r, 24);
    sh.getRange(r, 1, 1, 2).setValues([[a, b]])
      .setFontWeight('bold').setBackground(GREEN_LIGHT).setFontColor(INK)
      .setVerticalAlignment('middle');
    r++;
  };

  // ── Title ────────────────────────────────────────────────────────────────
  title_('Pacing Report: Supervisor Guide');
  spacer_();

  // ── Overview ─────────────────────────────────────────────────────────────
  prose_(
    'This guide covers everything a supervisor needs to use the Pacing Report day-to-day. ' +
    'The report pulls live ticket data from Gorgias, cross-references the team schedule, and ' +
    'produces color-coded pacing for every agent at six checkpoints throughout the day. Goals ' +
    'scale automatically to each agent\'s actual shift length, and supervisors can make same-day ' +
    'adjustments when hours change.',
    58
  );
  spacer_();

  // ── Reading the Daily Tab ─────────────────────────────────────────────────
  sectionHeader_('Reading the Daily Tab');
  colHeader_('Topic', 'Detail');

  qa_('Checkpoints',
    '6 checkpoints fire automatically each day: 7AM, 9AM, 11AM, 2PM, 8PM (EOD). Each one pulls live Gorgias data for the team and fills in that time block.',
    40, false);

  qa_('Metric columns',
    'Each checkpoint block has 4 columns: Closed Tickets, Tickets Replied, Messages Sent, CSAT. Each is colored independently.',
    40, true);

  qa_('Pacing colors',
    'Green = on track for their goal at this point in the day.  Yellow = slightly behind.  Red = behind.  ' +
    'The thresholds are set in the Config tab (default: green at 100%, yellow at 90%).',
    48, false);

  qa_('Exempt / CTO / VTO',
    'Agents who are off, on full CTO, or full VTO have their metric cells cleared and painted a neutral color. ' +
    'They are excluded from pacing calculations for the day.',
    48, true);

  qa_('Partial CTO / VTO',
    'Agents who work part of the day and take CTO or VTO for the rest will show real metric data for the ' +
    'checkpoints they were working, and a CTO/VTO color for checkpoints after they left.',
    52, false);

  qa_('On Track column',
    'Auto-populated at each publish. Shows Yes, No, or Exempt based on whether the agent is meeting their checkpoint target across all metrics.',
    44, true);

  qa_('EOD Goal Met column',
    'Auto-populated at the EOD checkpoint. Shows Yes, No, or Exempt. Reflects whether the agent hit their full-day goal after all adjustments are applied.',
    44, false);

  qa_('Actions Taken column',
    'Auto-populated in certain situations — Working Lunch, CTO, VTO, Off. Supervisors do not need to fill this in manually.',
    44, true);

  qa_('Notes column (Column AE)',
    'Written automatically at each publish. Contains the agent\'s schedule status, scheduled hours, effective hours after lunch deduction, and the target values for the last published checkpoint. ' +
    'Reporting tools read this column — when a goal adjustment is applied, this column updates automatically.',
    60, false);

  spacer_();

  // ── Goal Adjustments ─────────────────────────────────────────────────────
  sectionHeader_('Goal Adjustments');
  colHeader_('Topic', 'Detail');

  qa_('What it does',
    'When a supervisor removes hours from an agent\'s goal (CTO hours, a project block, etc.), the system recomputes that agent\'s targets and repaints all checkpoint colors accordingly. ' +
    'The EOD Goal Met column also recalculates if EOD has already been published.',
    56, false);

  qa_('When to use it',
    'Any time an agent\'s available hours for the day changed after the schedule was set — approved CTO, a mid-day project pull, unexcused absence, or a performance-related adjustment.',
    52, true);

  qa_('How to apply',
    '1. Open the daily tab for the relevant date.\n' +
    '2. Find the agent\'s row.\n' +
    '3. Enter the number of hours to remove in the Hours Removed column.\n' +
    '4. Select a reason from the Reason dropdown.\n' +
    '5. Run Pacing Report → Apply Goal Adjustments from the menu.',
    80, false);

  qa_('Valid reasons',
    'CTO  ·  VTO  ·  Unexcused Absence  ·  Project  ·  Performance',
    36, true);

  qa_('What updates',
    'All checkpoint colors for that agent are repainted. EOD Goal Met recalculates (if EOD is published). ' +
    'The Notes column (Column AE) updates with the new targets. The Status column shows "Applied" with the hours removed.',
    56, false);

  qa_('Can I adjust any agent?',
    'Yes. Any agent on the roster can receive a goal adjustment regardless of their current pacing status.',
    36, true);

  spacer_();

  // ── Weekly KPI ────────────────────────────────────────────────────────────
  sectionHeader_('Weekly KPI');
  colHeader_('Topic', 'Detail');

  qa_('What it shows',
    'Each agent\'s full-week performance across four metrics: QA Score, Tickets Replied, Closed Tickets, and CSAT. ' +
    'An Overall % is calculated as a weighted average of all four. The weekly tab is generated by running Admin → Rebuild Weekly Tabs.',
    64, false);

  qa_('KPI statuses',
    'Exceeding (106%+)  ·  Meeting (100-105%)  ·  Close (90-99%)  ·  Not Meeting (below 90%)  ·  Auto-Fail  ·  Exempt',
    44, true);

  qa_('Auto-Fail',
    'Triggered when an agent\'s QA score is at or below 74%, or their weekly tickets replied falls below roughly 40% of their goal (scales with shift hours). ' +
    'Auto-fail agents\' actual scores are still included in the team-wide Overall Average. Auto-fails are called out by name on the Team Dashboard.',
    64, false);

  qa_('Team Dashboard',
    'A separate tab that shows team-level TCPH, TRPH, QA Avg, and CSAT; the Overall Avg; auto-fail names; ' +
    'week-over-week deltas; QA Lead stats; and the inbox health snapshot from the Pulse Log. Built alongside the weekly tab.',
    60, true);

  spacer_();

  // ── Menus ─────────────────────────────────────────────────────────────────
  sectionHeader_('Menus');
  colHeader_('Menu Item', 'What It Does');

  qa_('Pacing Report → Publish Current Checkpoint',
    'Pulls live Gorgias data and fills in the current time block for all agents. Runs automatically on a trigger — use this only if a checkpoint was missed or needs to be re-run.',
    52, false);

  qa_('Pacing Report → Rebuild Current Day',
    'Deletes and recreates today\'s tab, then republishes all checkpoints that have already fired. Use this if the tab got into a bad state.',
    48, true);

  qa_('Pacing Report → Apply Goal Adjustments',
    'Reads the Hours Removed and Reason columns on the active daily tab and repaints all checkpoint colors for affected agents. Must be run manually after entering adjustments.',
    52, false);

  qa_('KPI Supervisor View → Filter Active Tab to Manager',
    'Hides all rows except agents belonging to the entered manager. Useful for per-manager reviews. Publish still works correctly when a filter is active.',
    52, true);

  qa_('KPI Supervisor View → Show All Managers',
    'Clears the manager filter and makes all agent rows visible again.',
    36, false);

  qa_('KPI Supervisor View → Sort Active Tab by Manager',
    'Sorts the active daily tab by manager, then by rep name within each manager group.',
    40, true);

  qa_('Mode → Internal / External',
    'Internal shows the full tab set (Team Guide, 7 daily, 4 weekly, Staffing). ' +
    'External trims to Case Use Summary, Team Dashboard, 3 daily, 1 weekly — use this when sharing the sheet with leadership.',
    52, false);

  qa_('Admin → Normalize Schedule',
    'Reads the Schedule tab and rebuilds the internal Schedule_Normalized tab. Run this at the start of each week before the first publish.',
    48, true);

  qa_('Admin → Rebuild Weekly Tabs',
    'Regenerates all weekly KPI tabs and the Team Dashboard. Run this at the end of each week.',
    40, false);

  spacer_();

  // ── Roster ────────────────────────────────────────────────────────────────
  sectionHeader_('Adding and Removing Agents');
  colHeader_('Action', 'How');

  qa_('Add an agent',
    'Open the Roster tab (Admin → Show Admin Tabs to make it visible). Add a new row with the agent\'s Gorgias Agent ID and their name exactly as it appears in the schedule.',
    52, false);

  qa_('Remove an agent',
    'Delete their row from the Roster tab. They will no longer appear on new daily tabs. Historical tabs are unaffected.',
    44, true);

  qa_('Name matching',
    'The system matches agents between the schedule and Gorgias by name. The name in the Roster tab should match the schedule closely. First-name matching is used as a fallback.',
    52, false);

  spacer_();

  // ── Troubleshooting ───────────────────────────────────────────────────────
  sectionHeader_('Troubleshooting');
  colHeader_('Issue', 'What to Do');

  qa_('Today\'s tab has wrong data or wrong agents',
    '1. Run Admin → Normalize Schedule.\n2. Run Pacing Report → Rebuild Current Day.\nThis recreates the tab and republishes all checkpoints that have fired so far today.',
    56, false);

  qa_('A checkpoint didn\'t fire',
    'Run Pacing Report → Publish Current Checkpoint while still within that checkpoint\'s hour, or use Rebuild Current Day to catch up all missed checkpoints at once.',
    52, true);

  qa_('Agent shows as Unscheduled but should be working',
    'Check the Schedule tab. Their name may not match the Roster tab closely enough, or their shift cell may be blank. Fix the schedule and run Normalize Schedule, then re-run the checkpoint.',
    56, false);

  qa_('Goal adjustment colors didn\'t update',
    'Make sure you ran Pacing Report → Apply Goal Adjustments after entering the hours. Also confirm you are on the correct daily tab before running it.',
    48, true);

  qa_('Metrics show as blank',
    'Confirm the agent\'s Gorgias Agent ID in the Roster tab is correct. A wrong ID means no data is returned from the API.',
    44, false);

  SpreadsheetApp.flush();
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
    .setValue('Pacing Report: Data Summary')
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
    'It pulls ticket data from Gorgias and cross-references the team schedule along with daily adjustments made by supervisors to account for ' +
    'actual hours worked, so goals are custom to each agent\'s shift rather than assuming ' +
    'everyone works the same hours.',

    'Reports update automatically at six checkpoints throughout the day (7AM, 9AM, 11AM, 2PM, ' +
    '6PM, and 8PM). Each checkpoint shows how the team is tracking relative to where they should ' +
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
     'Closed tickets, replies, and CSAT scores for each agent, pulled via API at each daily checkpoint.'],
    ['Team Schedule',
     'Shift hours and status (active, CTO, VTO, off). Used to calculate each agent\'s adjusted goal for the day. Agents on shorter shifts get proportionally smaller targets.'],
    ['QA Lead Report Card',
     'Weekly QA form completion for the QA Lead, tracked in a separate spreadsheet and pulled automatically each time the weekly report is built.'],
    ['Pulse Log',
     'End-of-day snapshots of inbox health: total open tickets, unassigned count, and a breakdown of how long tickets have been waiting (aging buckets from under 8 hours to over 24 hours).\n\nUse inbox data as context alongside KPIs. Low ticket output with a healthy inbox may indicate a well-managed workflow. High ticket output with an overloaded inbox may point to a staffing need. Low output and an overloaded inbox together is the most urgent signal. It might mean that volume is high and agents are not keeping up.']
  ].forEach(([source, desc], i) => {
    sh.setRowHeight(r, i === 3 ? 110 : 46);
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
     'N/A'],
    ['Auto-Fail',
     'A status applied when an agent\'s QA score falls at or below the auto-fail threshold, or their weekly tickets replied falls below the minimum ticket threshold. The agent\'s actual scores are still calculated and included in the team-wide Overall Avg. Auto-fail is surfaced separately as a count and name list on the Team Dashboard.',
     'QA score at or below 74%  ·  Tickets Replied below ~40% of weekly goal (scales with shift hours)'],
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
