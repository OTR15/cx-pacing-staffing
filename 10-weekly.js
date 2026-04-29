// =============================================================================
// weekly.gs
// Builds weekly leadership summary reports from daily tab data.
//
// Flow:
//   buildWeeklyReportIfSunday_()         — called from publishEOD each night
//     → buildWeeklyReportForWeek_()      — creates/rebuilds the weekly tab
//       → collectWeeklyData_()           — aggregates EOD data from daily tabs
//       → writeWeeklyLeadershipSummary_  — team totals with WoW delta
//       → writeWeeklyRepTable_           — per-rep volume + efficiency
//       → writeWeeklyDailyMovementSection_ — team daily breakdown
//       → writeWeeklyCharts_             — embedded charts
//
// NOTE: getPreviousWeekSummary_() reads from hardcoded cell addresses (B3–B10)
// in the previous week's tab. These will silently return wrong data if the
// weekly report layout ever changes.
// TODO: Store weekly totals in named ranges or a summary sheet instead.
// =============================================================================

// ── Entry points ──────────────────────────────────────────────────────────────

/**
 * Builds the weekly report only if today is Sunday.
 * Called from publishEOD() each night.
 * @param {Date} dateObj
 */
function buildWeeklyReportIfSunday_(dateObj) {
  const dayName = Utilities.formatDate(dateObj, CFG.timezone, 'EEE');
  if (dayName !== 'Sun') return;
  buildWeeklyReportForWeek_(dateObj);
}

/**
 * Builds (or rebuilds) the weekly report tab for the week containing dateObj.
 * Week is defined as Monday–Sunday.
 * @param {Date} dateObj - Any date within the target week.
 */
function buildWeeklyReportForWeek_(dateObj) {
  const ss      = SpreadsheetApp.getActive();
  const week    = getWeekRangeMondaySunday_(dateObj);
  const tabName = getWeeklyTabName_(week.monday, week.sunday);

  let sh = ss.getSheetByName(tabName);


  if (!sh) {
    sh = ss.insertSheet(tabName);
  } else {
    clearWeeklyReportSheet_(sh);
  }

  sh.setTabColor(CFG.weekly.tabColor);

  const weekData = collectWeeklyData_(week.monday, week.sunday);
  const kpiSnapshot = collectWeeklyKpiSnapshot_(week.monday, week.sunday);
  writeWeeklyLeadershipSummary_(sh, weekData, week.monday, week.sunday);
  writeWeeklyRepTable_(sh, weekData);
  writeWeeklyDailyMovementSection_(sh, weekData);
  writeWeeklyKpiSnapshotSection_(sh, kpiSnapshot, week.monday, week.sunday);
  writeWeeklyCharts_(sh, weekData, kpiSnapshot);
  buildTeamDashboard_(weekData, kpiSnapshot, week.monday, week.sunday);

  sh.autoResizeColumns(1, Math.min(21, sh.getMaxColumns()));
}

/**
 * Backfills weekly reports for all Sundays from a start date through today.
 * Useful when the script is first deployed on an existing workbook.
 * Adjust the start date as needed.
 */
function backfillWeeklyReports() {
  const start = new Date(2026, 2, 22); // March 22, 2026 — update as needed
  const today = new Date();

  let d = stripTimeLocal_(start);
  while (d <= today) {
    const dayName = Utilities.formatDate(d, CFG.timezone, 'EEE');
    if (dayName === 'Sun') {
      buildWeeklyReportForWeek_(d);
    }
    d = addDaysLocal_(d, 1);
  }

  manageWeeklyTabs();
}

function rebuildAllWeeklyReports() {
  backfillWeeklyReports();
  SpreadsheetApp.getUi().alert('Weekly tabs rebuilt, including the missing Week 3/30 - 4/5.');
}

// ── Data collection ───────────────────────────────────────────────────────────

/**
 * Aggregates EOD metric data from each daily tab in the week range.
 * Only tabs that exist are included (missing days are skipped gracefully).
 *
 * Per-rep data includes:
 *   - Daily breakdowns (for the movement section)
 *   - Weekly totals
 *   - Per-hour efficiency metrics
 *   - CSAT average (only from days with surveys)
 *   - Days worked count
 *
 * @param {Date} monday - Week start (Monday midnight)
 * @param {Date} sunday - Week end (Sunday midnight)
 * @returns {Object} weekData — see return shape below
 */
function collectWeeklyData_(monday, sunday) {
  const ss             = SpreadsheetApp.getActive();
  const layout         = getLayout_();
  const eodSection     = layout.sections.find(s => s.key === 'EOD');
  const progressStartCol = layout.progressStartCol;

  const repMap         = {};
  const teamDaily      = [];
  let weekClosed       = 0;
  let weekReplied      = 0;
  let weekMessages     = 0;
  let weekEffectiveHours = 0;

  for (let i = 0; i < 7; i++) {
    const day     = addDaysLocal_(monday, i);
    const tabName = formatDailySheetName_(day);
    const sh      = ss.getSheetByName(tabName);
    if (!sh) continue;

    const lastRow = sh.getLastRow();
    if (lastRow < CFG.daily.firstDataRow) continue;

    const rowCount  = lastRow - CFG.daily.firstDataRow + 1;
    const names     = sh.getRange(CFG.daily.firstDataRow, 1, rowCount, 1).getDisplayValues();
    const eodMetrics = sh.getRange(CFG.daily.firstDataRow, eodSection.startCol, rowCount, 4).getValues();
    const notes     = sh.getRange(CFG.daily.firstDataRow, progressStartCol + 4, rowCount, 1).getDisplayValues();

    let teamClosed           = 0;
    let teamReplied          = 0;
    let teamMessages         = 0;
    let teamEffectiveHoursDay = 0;

    for (let r = 0; r < rowCount; r++) {
      const repName = String(names[r][0] || '').trim();
      if (!repName) continue;

      const closed  = Number(eodMetrics[r][0] || 0);
      const replied = Number(eodMetrics[r][1] || 0);
      const messages = Number(eodMetrics[r][2] || 0);
      const csat    = eodMetrics[r][3] === '' ? '' : Number(eodMetrics[r][3] || 0);
      const effectiveHours = extractHoursFromWeeklyNote_(String(notes[r][0] || ''));

      if (!repMap[repName]) {
        repMap[repName] = {
          repName,
          daily:               {},
          weeklyClosed:        0,
          weeklyReplied:       0,
          weeklyMessages:      0,
          weeklyEffectiveHours: 0,
          csatValues:          [],
          daysWorked:          0
        };
      }

      repMap[repName].daily[tabName] = { closed, replied, messages, csat, effectiveHours };
      repMap[repName].weeklyClosed        += closed;
      repMap[repName].weeklyReplied       += replied;
      repMap[repName].weeklyMessages      += messages;
      repMap[repName].weeklyEffectiveHours += effectiveHours;

      if (effectiveHours > 0) repMap[repName].daysWorked += 1;
      if (csat !== '' && !isNaN(csat)) repMap[repName].csatValues.push(csat);

      teamClosed           += closed;
      teamReplied          += replied;
      teamMessages         += messages;
      teamEffectiveHoursDay += effectiveHours;
    }

    teamDaily.push({ dateObj: day, tabName, closed: teamClosed, replied: teamReplied, messages: teamMessages, effectiveHours: teamEffectiveHoursDay });
    weekClosed         += teamClosed;
    weekReplied        += teamReplied;
    weekMessages       += teamMessages;
    weekEffectiveHours += teamEffectiveHoursDay;
  }

  const reps = Object.keys(repMap).sort().map(name => {
    const rep     = repMap[name];
    const avgCsat = rep.csatValues.length
      ? rep.csatValues.reduce((a, b) => a + b, 0) / rep.csatValues.length
      : '';

    return {
      repName:              rep.repName,
      weeklyClosed:         rep.weeklyClosed,
      weeklyReplied:        rep.weeklyReplied,
      weeklyMessages:       rep.weeklyMessages,
      weeklyEffectiveHours: rep.weeklyEffectiveHours,
      closedPerHour:        rep.weeklyEffectiveHours ? rep.weeklyClosed   / rep.weeklyEffectiveHours : 0,
      repliedPerHour:       rep.weeklyEffectiveHours ? rep.weeklyReplied  / rep.weeklyEffectiveHours : 0,
      messagesPerHour:      rep.weeklyEffectiveHours ? rep.weeklyMessages / rep.weeklyEffectiveHours : 0,
      avgCsat,
      daysWorked: rep.daysWorked,
      daily:      rep.daily
    };
  });

  return {
    monday,
    sunday,
    reps,
    teamDaily,
    teamTotals: {
      closed:          weekClosed,
      replied:         weekReplied,
      messages:        weekMessages,
      effectiveHours:  weekEffectiveHours,
      closedPerHour:   weekEffectiveHours ? weekClosed   / weekEffectiveHours : 0,
      repliedPerHour:  weekEffectiveHours ? weekReplied  / weekEffectiveHours : 0,
      messagesPerHour: weekEffectiveHours ? weekMessages / weekEffectiveHours : 0
    },
    previousWeek: getPreviousWeekSummary_(monday)
  };
}

/**
 * Extracts the effective hours value from the Notes cell written by publishCheckpointForDate_.
 * Looks for "Effective: X" or "Hours: X" patterns.
 * Returns 0 if neither pattern is found.
 *
 * @param {string} noteText
 * @returns {number}
 */
function extractHoursFromWeeklyNote_(noteText) {
  const text = String(noteText || '');
  let m = text.match(/Effective:\s*([0-9]+(?:\.[0-9]+)?)/i);
  if (m) return Number(m[1]);
  m = text.match(/Hours:\s*([0-9]+(?:\.[0-9]+)?)/i);
  if (m) return Number(m[1]);
  return 0;
}

// ── Report writers ────────────────────────────────────────────────────────────

/**
 * Writes the leadership summary section (rows 1–10):
 *   Row 1:  title bar
 *   Row 2:  column headers
 *   Rows 3–10: team metrics with this-week, prior-week, and WoW delta
 *
 * NOTE: getPreviousWeekSummary_() reads B3:B10 from the previous tab.
 * These cell references are fragile — see file header TODO.
 */
function writeWeeklyLeadershipSummary_(sh, weekData, monday, sunday) {
  sh.getRange('A1:H1').merge();
  sh.getRange('A1')
    .setValue(
      'Weekly Leadership Report: ' +
      Utilities.formatDate(monday, CFG.timezone, 'M/d/yy') +
      ' - ' +
      Utilities.formatDate(sunday, CFG.timezone, 'M/d/yy')
    )
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#9fc5f8');

  const t = weekData.teamTotals;
  const p = weekData.previousWeek;

  const rows = [
    ['Metric',             'This Week',                  'Previous Week',           'WoW Delta'],
    ['Closed Tickets',     t.closed,                     p.closed,                  wowDelta_(t.closed,          p.closed)],
    ['Tickets Replied',    t.replied,                    p.replied,                 wowDelta_(t.replied,         p.replied)],
    ['Messages Sent',      t.messages,                   p.messages,                wowDelta_(t.messages,        p.messages)],
    ['Effective Hours',    round2_(t.effectiveHours),    p.effectiveHours,          wowDelta_(t.effectiveHours,  p.effectiveHours)],
    ['Closed per Hour',    round2_(t.closedPerHour),     p.closedPerHour   === '' ? '' : round2_(p.closedPerHour),   wowDelta_(t.closedPerHour,   p.closedPerHour)],
    ['Replied per Hour',   round2_(t.repliedPerHour),    p.repliedPerHour  === '' ? '' : round2_(p.repliedPerHour),  wowDelta_(t.repliedPerHour,  p.repliedPerHour)],
    ['Messages per Hour',  round2_(t.messagesPerHour),   p.messagesPerHour === '' ? '' : round2_(p.messagesPerHour), wowDelta_(t.messagesPerHour, p.messagesPerHour)]
  ];

  sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  sh.getRange(2, 1, 1, 4).setFontWeight('bold').setBackground('#cfe2f3');
}

/**
 * Writes the per-rep performance table starting at row 12.
 * Columns: name, volume metrics, efficiency per hour, days worked.
 */
function writeWeeklyRepTable_(sh, weekData) {
  const startRow = 12;
  const headers  = [[
    'Rep Name', 'Weekly Closed', 'Weekly Replied', 'Weekly Messages',
    'Effective Hours', 'Closed/Hr', 'Replied/Hr', 'Messages/Hr', 'Days Worked'
  ]];

  const rows = weekData.reps.map(rep => [
    rep.repName,
    rep.weeklyClosed,
    rep.weeklyReplied,
    rep.weeklyMessages,
    round2_(rep.weeklyEffectiveHours),
    round2_(rep.closedPerHour),
    round2_(rep.repliedPerHour),
    round2_(rep.messagesPerHour),
    rep.daysWorked
  ]);

  sh.getRange(startRow, 1, 1, headers[0].length).setValues(headers).setFontWeight('bold').setBackground('#cfe2f3');
  if (rows.length) {
    sh.getRange(startRow + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
}

/**
 * Writes the daily team movement table starting at row 12, column 13.
 * Shows team totals and efficiency per day with a blank "Movement Note" column.
 */
function writeWeeklyDailyMovementSection_(sh, weekData) {
  const startRow = 12;
  const startCol = 13;

  sh.getRange(startRow, startCol, 1, 9).setValues([[
    'Day', 'Closed', 'Replied', 'Messages', 'Effective Hours',
    'Closed/Hr', 'Replied/Hr', 'Messages/Hr', 'Movement Note'
  ]]).setFontWeight('bold').setBackground('#cfe2f3');

  const rows = weekData.teamDaily.map(day => [
    Utilities.formatDate(day.dateObj, CFG.timezone, 'EEE M/d'),
    day.closed,
    day.replied,
    day.messages,
    round2_(day.effectiveHours),
    day.effectiveHours ? round2_(day.closed   / day.effectiveHours) : 0,
    day.effectiveHours ? round2_(day.replied  / day.effectiveHours) : 0,
    day.effectiveHours ? round2_(day.messages / day.effectiveHours) : 0,
    '' // Movement Note — filled manually by leadership
  ]);

  if (rows.length) {
    sh.getRange(startRow + 1, startCol, rows.length, rows[0].length).setValues(rows);
  }
}

/**
 * Embeds four charts into the weekly report tab:
 *   1. Line chart: team daily volume (closed/replied/messages over the week)
 *   2. Bar chart:  weekly closed tickets by rep
 *   3. Column chart: hourly productivity by rep
 *   4. Bar chart: weekly KPI overall score snapshot by agent
 */
function writeWeeklyCharts_(sh, weekData, kpiSnapshot) {
  const dailyStartRow = 13;
  const dailyStartCol = 13;
  const repTableStartRow = 12;
  const chartStartCol = CFG.weekly.chartStartCol || 24;
  const snapshotStartRow = CFG.weekly.kpiSnapshot.tableStartRow;

  sh.insertChart(
    sh.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(sh.getRange(dailyStartRow, dailyStartCol, weekData.teamDaily.length + 1, 4))
      .setPosition(2, chartStartCol, 0, 0)
      .setOption('title', 'Team Daily Volume')
      .setOption('width', 720)
      .setOption('height', 260)
      .build()
  );

  sh.insertChart(
    sh.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sh.getRange(repTableStartRow, 1, weekData.reps.length + 1, 2))
      .setPosition(20, chartStartCol, 0, 0)
      .setOption('title', 'Weekly Closed Tickets by Rep')
      .setOption('width', 720)
      .setOption('height', 260)
      .setOption('legend', { position: 'none' })
      .build()
  );

  sh.insertChart(
    sh.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sh.getRange(repTableStartRow, 1, weekData.reps.length + 1, 9))
      .setPosition(38, chartStartCol, 0, 0)
      .setOption('title', 'Hourly Productivity by Rep')
      .setOption('width', 720)
      .setOption('height', 260)
      .build()
  );

  if (kpiSnapshot.chartRows.length) {
    // Chart data source starts after the snapshot's Reason + Goal Adj columns (WKPI_TOTAL_COLS = 13).
    const kpiChartDataCol = WKPI_TOTAL_COLS + 2; // col 15
    const chartDataRowCount = kpiSnapshot.chartRows.length + 1;
    sh.getRange(snapshotStartRow + 1, kpiChartDataCol, chartDataRowCount, 2).clearContent();
    sh.getRange(snapshotStartRow + 1, kpiChartDataCol, 1, 2).setValues([['Agent', 'Overall %']]);
    sh.getRange(snapshotStartRow + 2, kpiChartDataCol, kpiSnapshot.chartRows.length, 2).setValues(kpiSnapshot.chartRows);
    sh.getRange(snapshotStartRow + 2, kpiChartDataCol + 1, kpiSnapshot.chartRows.length, 1).setNumberFormat('0.0"%"');

    sh.insertChart(
      sh.newChart()
        .setChartType(Charts.ChartType.BAR)
        .addRange(sh.getRange(snapshotStartRow + 1, kpiChartDataCol, chartDataRowCount, 2))
        .setPosition(CFG.weekly.kpiSnapshot.chartStartRow, chartStartCol, 0, 0)
        .setOption('title', 'Weekly KPI Overall % by Agent')
        .setOption('width', 720)
        .setOption('height', 300)
        .setOption('legend', { position: 'none' })
        .build()
    );
  }
}

function clearWeeklyReportSheet_(sh) {
  sh.getCharts().forEach(chart => sh.removeChart(chart));
  sh.clear();
  // sh.clear() does not clear data validation — remove it explicitly so
  // stale validation rules don't persist across rebuilds.
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
}

// ── Tab management ────────────────────────────────────────────────────────────

/**
 * Keeps only the most recent N weekly tabs visible (N = CFG.weekly.visibleTabCount).
 * Older weekly tabs are hidden (not deleted).
 */
function manageWeeklyTabs() {
  const ss = SpreadsheetApp.getActive();

  const weeklyTabs = ss.getSheets()
    .map(sh => ({ sheet: sh, name: sh.getName(), date: parseWeeklyTabDate_(sh.getName()) }))
    .filter(x => x.date)
    .sort((a, b) => b.date.getTime() - a.date.getTime());

  weeklyTabs.forEach((item, index) => {
    if (index < CFG.weekly.visibleTabCount) {
      item.sheet.showSheet();
    } else {
      item.sheet.hideSheet();
    }
  });
}

// ── Date helpers ──────────────────────────────────────────────────────────────

/**
 * Returns the Monday and Sunday dates bounding the week that contains dateObj.
 * Uses ISO weekday (Mon=1 … Sun=7).
 *
 * @param {Date} dateObj
 * @returns {{ monday: Date, sunday: Date }}
 */
function getWeekRangeMondaySunday_(dateObj) {
  const d     = new Date(dateObj);
  const jsDay = Number(Utilities.formatDate(d, CFG.timezone, 'u')); // Mon=1, Sun=7
  const monday = new Date(d);
  monday.setDate(d.getDate() - (jsDay - 1));
  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);

  return { monday: stripTimeLocal_(monday), sunday: stripTimeLocal_(sunday) };
}

/**
 * Returns the weekly tab name for a given Monday–Sunday range.
 * e.g. "Week 3/24 - 3/30"
 */
function getWeeklyTabName_(monday, sunday) {
  return 'Week ' +
    Utilities.formatDate(monday, CFG.timezone, 'M/d') +
    ' - ' +
    Utilities.formatDate(sunday, CFG.timezone, 'M/d');
}

/**
 * Parses a weekly tab name like "Week 3/24 - 3/30" into the Monday Date.
 * Returns null if the name doesn't match the expected pattern.
 * @param {string} name
 * @returns {Date|null}
 */
function parseWeeklyTabDate_(name) {
  const m = String(name || '').match(/^Week\s+(\d{1,2})\/(\d{1,2})\s+-\s+(\d{1,2})\/(\d{1,2})$/);
  if (!m) return null;
  const currentYear = Number(Utilities.formatDate(new Date(), CFG.timezone, 'yyyy'));
  return new Date(currentYear, Number(m[1]) - 1, Number(m[2]));
}

// ── Previous week summary ─────────────────────────────────────────────────────

/**
 * Reads team totals from the previous week's report tab.
 *
 * FRAGILE: reads from hardcoded cell addresses B3–B10.
 * If the weekly report layout changes, these will silently return wrong data.
 * TODO: Replace with a named range or a dedicated summary row.
 *
 * Returns empty-string values for all fields if the previous week's tab
 * doesn't exist yet.
 *
 * @param {Date} currentMonday - The Monday of the current week.
 * @returns {Object}
 */
function getPreviousWeekSummary_(currentMonday) {
  const ss          = SpreadsheetApp.getActive();
  const prevMonday  = addDaysLocal_(currentMonday, -7);
  const prevSunday  = addDaysLocal_(prevMonday, 6);
  const prevTabName = getWeeklyTabName_(prevMonday, prevSunday);
  const sh          = ss.getSheetByName(prevTabName);

  if (!sh) {
    return { closed: '', replied: '', messages: '', effectiveHours: '', closedPerHour: '', repliedPerHour: '', messagesPerHour: '', avgCsat: '' };
  }

  return {
    closed:          Number(sh.getRange('B3').getValue() || 0),
    replied:         Number(sh.getRange('B4').getValue() || 0),
    messages:        Number(sh.getRange('B5').getValue() || 0),
    avgCsat:         sh.getRange('B6').getValue() === '' ? '' : Number(sh.getRange('B6').getValue() || 0),
    effectiveHours:  Number(sh.getRange('B7').getValue() || 0),
    closedPerHour:   Number(sh.getRange('B8').getValue() || 0),
    repliedPerHour:  Number(sh.getRange('B9').getValue() || 0),
    messagesPerHour: Number(sh.getRange('B10').getValue() || 0)
  };
}
