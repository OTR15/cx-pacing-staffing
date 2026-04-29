// =============================================================================
// 15-team-dashboard.gs
// Builds and maintains the Team Dashboard tab.
// Called from buildWeeklyReportForWeek_() in 10-weekly.js each time a weekly
// report is built. The tab is purple and named 'Team Dashboard'.
//
// Sections (top → bottom):
//   1. Title
//   2. Team Performance  — TCPH, TRPH, QA Avg, CSAT (QA lead excluded)
//   3. KPI Report Card   — Overall % averages, status distribution, auto-fails
//   4. Week-over-Week    — vs last week and 4-week avg for all 5 metrics
//   5. QA Lead           — form completion from the QA tracker spreadsheet
//   6. Trend chart       — % of goal over time, auto-updating line chart
//   7. Trend data table  — one row per week, source for the chart; persists
// =============================================================================

const DASH_TAB_NAME  = 'Team Dashboard';
const DASH_TAB_COLOR = '#7b4f9e';

// ── Fixed section row anchors ─────────────────────────────────────────────────
const DASH_R_TITLE        = 1;
const DASH_R_TEAM_HDR     = 3;
const DASH_R_TEAM_LABELS  = 4;
const DASH_R_TEAM_ACTUAL  = 5;
const DASH_R_TEAM_GOAL    = 6;
const DASH_R_TEAM_PCT     = 7;
const DASH_R_KPI_HDR      = 9;
const DASH_R_KPI_AVGS     = 10;
const DASH_R_KPI_DIST     = 11;
const DASH_R_KPI_FAILS    = 12;
const DASH_R_WOW_HDR      = 14;
const DASH_R_WOW_COLS     = 15;
const DASH_R_WOW_TCPH     = 16;
const DASH_R_WOW_TRPH     = 17;
const DASH_R_WOW_QA       = 18;
const DASH_R_WOW_CSAT     = 19;
const DASH_R_WOW_OVERALL  = 20;
const DASH_R_LEAD_HDR     = 22;
const DASH_R_LEAD_DATA    = 23;
const DASH_R_CHART_START  = 25;
const DASH_R_CHART_END    = 44;
const DASH_R_TREND_HDR    = 46;
const DASH_R_TREND_COLS   = 47;
const DASH_R_TREND_DATA   = 48;

// ── Trend table columns (1-based) ─────────────────────────────────────────────
const TD_WEEK      = 1;
const TD_TCPH      = 2;
const TD_TRPH      = 3;
const TD_QA        = 4;
const TD_CSAT      = 5;
const TD_OVR_ALL   = 6;
const TD_OVR_EXCL  = 7;
const TD_QA_FORMS  = 8;
const TD_QA_TGT    = 9;
const TD_TCPH_GOAL = 10;
const TD_TRPH_GOAL = 11;
const TD_QA_GOAL   = 12;
const TD_CSAT_GOAL = 13;
const TD_TCPH_PCT     = 14; // % of goal — chart series
const TD_TRPH_PCT     = 15;
const TD_QA_PCT       = 16;
const TD_CSAT_PCT     = 17;
const TD_QA_LEAD_PCT      = 18; // QA Lead completed/target %
const TD_MEGHAN_QA_PCT    = 19; // Meghan individual QA % of goal
const TD_TOTAL            = 19;

// ── Pulse Log constants ───────────────────────────────────────────────────────
const DASH_PULSE_COL      = 10; // Column J — first column of pulse table
const DASH_PULSE_NCOLS    = 7;  // Date + 6 metrics
const DASH_PULSE_ROW_HDR  = 3;
const DASH_PULSE_ROW_COLS = 4;
const DASH_PULSE_ROW_DATA = 5;  // rows 5–11 for Mon–Sun

// Pulse log source column offsets (0-based into getValues row)
// Timestamp=0, Total Open=1, Unassigned=3, Over24=5,
// 24hrBkt=7, 22hrBkt=9, 20hrBkt=11, 18hrBkt=13, 16hrBkt=15,
// 14hrBkt=17, 12hrBkt=19, 10hrBkt=21, 8hrBkt=23, Under8=25
const PL_TIMESTAMP  = 0;
const PL_TOTAL_OPEN = 1;
const PL_UNASSIGNED = 3;
const PL_OVER_24    = 5;
const PL_BKT_24     = 7;
const PL_BKT_22     = 9;
const PL_BKT_20     = 11;
const PL_BKT_18     = 13;
const PL_BKT_16     = 15;
const PL_BKT_14     = 17;
const PL_BKT_12     = 19;
const PL_BKT_10     = 21;
const PL_BKT_8      = 23;

// ── Color palette ─────────────────────────────────────────────────────────────
const DASH_C_PURPLE      = '#7b4f9e';
const DASH_C_BLUE        = '#4a86e8';
const DASH_C_BLUE_LIGHT  = '#cfe2f3';
const DASH_C_TEAL        = '#00897b';
const DASH_C_TEAL_LIGHT  = '#b2dfdb';
const DASH_C_GRAY        = '#666666';
const DASH_C_GRAY_LIGHT  = '#f3f3f3';
const DASH_C_GREEN       = '#b6d7a8';
const DASH_C_YELLOW      = '#ffe599';
const DASH_C_RED         = '#f4cccc';
const DASH_C_WHITE       = '#ffffff';

// =============================================================================
// Entry point — called from buildWeeklyReportForWeek_()
// =============================================================================

function buildTeamDashboard_(weekData, kpiSnapshot, monday, sunday) {
  const ss         = SpreadsheetApp.getActive();
  const qaLeadName = String(getConfigValue_('QA_LEAD_NAME', '') || '').trim();

  // Get or create tab
  let sh    = ss.getSheetByName(DASH_TAB_NAME);
  const isNew = !sh;
  if (isNew) sh = ss.insertSheet(DASH_TAB_NAME);
  sh.setTabColor(DASH_TAB_COLOR);

  // Compute all metrics
  const goals         = getTeamDashboardGoals_();
  const teamMetrics   = computeTeamMetrics_(weekData, kpiSnapshot, qaLeadName);
  const kpiStats      = computeKpiReportCardStats_(kpiSnapshot, qaLeadName);
  const meghanQaPct   = computeMeghanQaPct_(kpiSnapshot);
  const qaStats       = collectQALeadWeeklyStats_(monday);
  const pulseData     = collectPulseLogData_(monday);
  const trendData     = readTrendData_(sh);
  const wowMetrics    = computeWoWMetrics_(weekData.previousWeek, trendData, goals);

  // Clear content rows only — trend data (row 48+) and chart (overlay) are preserved
  if (!isNew) {
    sh.getRange(1, 1, DASH_R_CHART_END, sh.getMaxColumns())
      .clearContent()
      .clearFormat();
  }

  // Write all sections
  writeDashTitle_(sh, monday, sunday);
  writeDashTeamPerformance_(sh, teamMetrics, goals);
  writeDashKpiReportCard_(sh, kpiStats);
  writeDashWoW_(sh, teamMetrics, kpiStats, wowMetrics);
  writeDashQALead_(sh, qaLeadName, qaStats);
  writeDashPulseLog_(sh, pulseData);
  writeTrendSectionHeaders_(sh);

  // Upsert current week's row in the trend table
  const weekLabel = dashWeekLabel_(monday, sunday);
  upsertTrendRow_(sh, weekLabel, teamMetrics, kpiStats, qaStats, goals, trendData, meghanQaPct);

  // Build chart on first creation or rebuild if it's missing the Meghan series
  const existingCharts = sh.getCharts();
  if (existingCharts.length === 0 || existingCharts[0].getRanges().length < 7) {
    existingCharts.forEach(c => sh.removeChart(c));
    buildDashboardChart_(sh);
  }

  SpreadsheetApp.flush();
}

// =============================================================================
// Goals
// =============================================================================

function getTeamDashboardGoals_() {
  const goalsMap      = getGoalsMap_();
  const def           = goalsMap._default;
  const standardHours = Number(getConfigValue_('STANDARD_SHIFT_HOURS', CFG.standardShiftHours));

  const derivedTcph = standardHours > 0 ? def.closedTickets  / standardHours : 6.5;
  const derivedTrph = standardHours > 0 ? def.ticketsReplied / standardHours : 8.0;

  return {
    tcph: Number(getConfigValue_('TEAM_TCPH_GOAL', '')) || derivedTcph,
    trph: Number(getConfigValue_('TEAM_TRPH_GOAL', '')) || derivedTrph,
    qa:   Number(getConfigValue_('TEAM_QA_GOAL',   '')) || 90,
    csat: Number(getConfigValue_('TEAM_CSAT_GOAL',  '')) || (def.csat || 4.7)
  };
}

// =============================================================================
// Metric computation
// =============================================================================

function computeTeamMetrics_(weekData, kpiSnapshot, qaLeadName) {
  const qaKey = normalizeName_(qaLeadName);

  // TCPH / TRPH from weekData (hours-based, QA lead excluded)
  const agentReps    = weekData.reps.filter(r => normalizeName_(r.repName) !== qaKey);
  const totalClosed  = agentReps.reduce((s, r) => s + (r.weeklyClosed         || 0), 0);
  const totalReplied = agentReps.reduce((s, r) => s + (r.weeklyReplied        || 0), 0);
  const totalHours   = agentReps.reduce((s, r) => s + (r.weeklyEffectiveHours || 0), 0);
  const tcph         = totalHours > 0 ? totalClosed  / totalHours : 0;
  const trph         = totalHours > 0 ? totalReplied / totalHours : 0;

  // QA avg / CSAT avg from kpiSnapshot (QA lead excluded)
  const agentRows = kpiSnapshot.rows.filter(r =>
    normalizeName_(String(r[WKPI_COL_AGENT - 1] || '')) !== qaKey
  );
  const qaVals   = agentRows.map(r => Number(r[WKPI_COL_QA_SCORE - 1])).filter(v => v > 0);
  const csatVals = agentRows.map(r => Number(r[WKPI_COL_CSAT     - 1])).filter(v => v > 0);

  const qa   = qaVals.length   ? qaVals.reduce((a, b)   => a + b) / qaVals.length   : 0;
  const csat = csatVals.length ? csatVals.reduce((a, b) => a + b) / csatVals.length : 0;

  return { tcph, trph, qa, csat, totalHours, totalClosed, totalReplied };
}

function computeKpiReportCardStats_(kpiSnapshot, qaLeadName) {
  const qaKey    = normalizeName_(qaLeadName);
  const rows     = kpiSnapshot.rows.filter(r =>
    normalizeName_(String(r[WKPI_COL_AGENT - 1] || '')) !== qaKey
  );
  const counts    = { Exceeding: 0, Meeting: 0, Close: 0, 'Not Meeting': 0, 'AUTO-FAIL': 0 };
  const autoFails = [];
  const allPcts   = [];
  const exclPcts  = [];

  rows.forEach(row => {
    const name    = String(row[WKPI_COL_AGENT   - 1] || '').trim();
    const overall = Number(row[WKPI_COL_OVERALL  - 1]);
    const status  = String(row[WKPI_COL_STATUS   - 1] || '').trim();

    if (status in counts) counts[status]++;
    if (!isNaN(overall) && overall > 0) allPcts.push(overall);

    if (status === 'AUTO-FAIL') {
      autoFails.push(name);
    } else if (!isNaN(overall) && overall > 0) {
      exclPcts.push(overall);
    }
  });

  const avg     = allPcts.length  ? allPcts.reduce((a, b)   => a + b) / allPcts.length  : 0;
  const avgExcl = exclPcts.length ? exclPcts.reduce((a, b)  => a + b) / exclPcts.length : 0;

  return { avg, avgExcl, counts, autoFails, totalAgents: rows.length };
}

// =============================================================================
// Meghan individual QA % of goal
// =============================================================================

function computeMeghanQaPct_(kpiSnapshot) {
  const row = kpiSnapshot.rows.find(r =>
    normalizeFirstName_(String(r[WKPI_COL_AGENT - 1] || '')) === 'meghan'
  );
  if (!row) return '';
  const score = Number(row[WKPI_COL_QA_SCORE - 1]);
  const goal  = Number(row[WKPI_COL_QA_GOAL  - 1]);
  return (score > 0 && goal > 0) ? dashRound_(score / goal * 100) : '';
}

// =============================================================================
// QA Lead stats — cross-spreadsheet read
// =============================================================================

/**
 * Reads the QA Lead's weekly stats from the QA Lead Report Card spreadsheet.
 * Checks the Weekly Review sheet (current week, row 2) first, then falls back
 * to the Week Archive sheet (historical rows) when looking at past weeks.
 *
 * Weekly Review / Week Archive column layout (1-based):
 *   1=Week Start, 2=Weekly Target, 3=Weekly Actual, 4=Days Met, 5=Days Missed,
 *   6=Shortfall, 7=Backlog Days
 *
 * Returns: { completed, target, daysMet, daysMissed, shortfall } or nulls if unavailable.
 */
function collectQALeadWeeklyStats_(monday) {
  const spreadsheetId = String(getConfigValue_('QA_LEAD_REPORT_CARD_ID', '') || '').trim();
  if (!spreadsheetId) return qaLeadStatsEmpty_();

  try {
    const ss           = SpreadsheetApp.openById(spreadsheetId);
    const mondayKey    = dashDateKey_(monday);

    // ── Try Weekly Review sheet first (current week) ─────────────────────
    const reviewSh = ss.getSheetByName('Weekly Review');
    if (reviewSh && reviewSh.getLastRow() >= 2) {
      const row = reviewSh.getRange(2, 1, 1, 7).getValues()[0];
      if (dashDateKey_(row[0]) === mondayKey) {
        return qaLeadStatsFromRow_(row);
      }
    }

    // ── Fall back to Week Archive for historical weeks ────────────────────
    // The archive sheet has section headers embedded between data rows, so we
    // scan all rows and only match rows where col A is an actual Date object.
    const archiveSh = ss.getSheetByName('Week Archive');
    if (!archiveSh || archiveSh.getLastRow() < 1) return qaLeadStatsEmpty_();

    const lastRow = archiveSh.getLastRow();
    const allRows = archiveSh.getRange(1, 1, lastRow, 7).getValues();
    // Match only weekly summary rows: col A = Date matching monday, col B = numeric
    // target, col C = numeric actual. This skips daily-pace rows which have a day
    // name string in col B and a Date in col C.
    const match   = allRows.find(r =>
      r[0] instanceof Date && !isNaN(r[0]) &&
      dashDateKey_(r[0]) === mondayKey &&
      typeof r[1] === 'number' &&
      typeof r[2] === 'number'
    );
    return match ? qaLeadStatsFromRow_(match) : qaLeadStatsEmpty_();

  } catch (e) {
    Logger.log('collectQALeadWeeklyStats_ error: ' + e.message);
    return qaLeadStatsEmpty_();
  }
}

function qaLeadStatsFromRow_(row) {
  return {
    completed:  num_(row[2]), // Weekly Actual
    target:     num_(row[1]), // Weekly Target
    daysMet:    num_(row[3]),
    daysMissed: num_(row[4]),
    shortfall:  num_(row[5])
  };
}

function qaLeadStatsEmpty_() {
  return { completed: null, target: null, daysMet: null, daysMissed: null, shortfall: null };
}

function dashDateKey_(dateVal) {
  if (!dateVal) return '';
  try {
    // Add 12 hours to avoid midnight timezone boundary issues when dates are
    // read from external spreadsheets with a different timezone setting.
    const ms = (dateVal instanceof Date ? dateVal.getTime() : new Date(dateVal).getTime());
    return Utilities.formatDate(new Date(ms + 12 * 3600000), CFG.timezone, 'yyyy-MM-dd');
  } catch (e) {
    return String(dateVal);
  }
}

// ISO 8601 week number
function dashIsoWeekNumber_(dateObj) {
  const d = new Date(Date.UTC(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

// =============================================================================
// Week-over-Week
// =============================================================================

function computeWoWMetrics_(previousWeek, trendData, goals) {
  const rows    = trendData.rows;
  const prevRow = rows.length >= 1 ? rows[rows.length - 1] : null;

  const prev = prevRow ? {
    tcph:    num_(prevRow[TD_TCPH     - 1]),
    trph:    num_(prevRow[TD_TRPH     - 1]),
    qa:      num_(prevRow[TD_QA       - 1]),
    csat:    num_(prevRow[TD_CSAT     - 1]),
    overall: num_(prevRow[TD_OVR_EXCL - 1])
  } : {
    // No trend data yet — fall back to previousWeek object for TCPH/TRPH/CSAT
    tcph:    num_(previousWeek.closedPerHour  || 0),
    trph:    num_(previousWeek.repliedPerHour || 0),
    qa:      0,
    csat:    num_(previousWeek.avgCsat        || 0),
    overall: 0
  };

  // 4-week average from last 4 trend rows
  const four = rows.slice(-4);
  const avg4 = col => {
    const vals = four.map(r => num_(r[col - 1])).filter(v => v > 0);
    return vals.length ? vals.reduce((a, b) => a + b) / vals.length : null;
  };

  return {
    prev,
    avg4: {
      tcph:    avg4(TD_TCPH),
      trph:    avg4(TD_TRPH),
      qa:      avg4(TD_QA),
      csat:    avg4(TD_CSAT),
      overall: avg4(TD_OVR_EXCL)
    }
  };
}

// =============================================================================
// Trend data read / upsert
// =============================================================================

function readTrendData_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < DASH_R_TREND_DATA) return { rows: [] };

  const numRows = lastRow - DASH_R_TREND_DATA + 1;
  const rows    = sh.getRange(DASH_R_TREND_DATA, 1, numRows, TD_TOTAL)
    .getValues()
    .filter(r => String(r[0] || '').trim());

  return { rows };
}

function upsertTrendRow_(sh, weekLabel, teamMetrics, kpiStats, qaStats, goals, trendData, meghanQaPct) {
  const { tcph, trph, qa, csat } = teamMetrics;
  const safePct = (actual, goal) => goal > 0 ? dashRound_(actual / goal * 100) : 0;

  const newRow = [
    weekLabel,
    dashRound_(tcph),
    dashRound_(trph),
    dashRound_(qa),
    dashRound_(csat),
    dashRound_(kpiStats.avg),
    dashRound_(kpiStats.avgExcl),
    qaStats.completed !== null ? qaStats.completed : '',
    qaStats.target,
    dashRound_(goals.tcph),
    dashRound_(goals.trph),
    dashRound_(goals.qa),
    dashRound_(goals.csat),
    safePct(tcph, goals.tcph),
    safePct(trph, goals.trph),
    safePct(qa,   goals.qa),
    safePct(csat, goals.csat),
    (qaStats.completed !== null && qaStats.target > 0)
      ? dashRound_(qaStats.completed / qaStats.target * 100) : '',
    meghanQaPct !== undefined ? meghanQaPct : ''
  ];

  const existing = trendData.rows.findIndex(r => String(r[0] || '').trim() === weekLabel);
  if (existing >= 0) {
    sh.getRange(DASH_R_TREND_DATA + existing, 1, 1, TD_TOTAL).setValues([newRow]);
  } else {
    sh.getRange(DASH_R_TREND_DATA + trendData.rows.length, 1, 1, TD_TOTAL).setValues([newRow]);
  }
}

// =============================================================================
// Section writers
// =============================================================================

function writeDashTitle_(sh, monday, sunday) {
  sh.getRange(DASH_R_TITLE, 1, 1, 8).merge()
    .setValue('Team Dashboard — Week of ' + dashWeekLabel_(monday, sunday))
    .setBackground(DASH_C_PURPLE)
    .setFontColor(DASH_C_WHITE)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sh.setRowHeight(DASH_R_TITLE, 36);
}

function writeDashTeamPerformance_(sh, teamMetrics, goals) {
  const { tcph, trph, qa, csat } = teamMetrics;
  const green  = Number(getConfigValue_('PACING_GREEN_MIN',  1.0));
  const yellow = Number(getConfigValue_('PACING_YELLOW_MIN', 0.9));

  // Section header
  sh.getRange(DASH_R_TEAM_HDR, 1, 1, 5).merge()
    .setValue('TEAM PERFORMANCE  (agents only)')
    .setBackground(DASH_C_BLUE).setFontColor(DASH_C_WHITE)
    .setFontWeight('bold').setHorizontalAlignment('left');

  // Column labels
  sh.getRange(DASH_R_TEAM_LABELS, 1, 1, 5)
    .setValues([['', 'TCPH', 'TRPH', 'QA Avg', 'CSAT']])
    .setFontWeight('bold').setBackground(DASH_C_BLUE_LIGHT)
    .setHorizontalAlignment('center');

  // Actual values
  sh.getRange(DASH_R_TEAM_ACTUAL, 1, 1, 5)
    .setValues([['Actual', dashRound_(tcph), dashRound_(trph), dashRound_(qa) + '%', dashRound_(csat)]]);

  // Goal values
  sh.getRange(DASH_R_TEAM_GOAL, 1, 1, 5)
    .setValues([['Goal', dashRound_(goals.tcph), dashRound_(goals.trph), goals.qa + '%', goals.csat]])
    .setFontColor(DASH_C_GRAY);

  // % of goal — colored
  sh.getRange(DASH_R_TEAM_PCT, 1).setValue('% of Goal').setFontWeight('bold');
  [
    { actual: tcph, goal: goals.tcph },
    { actual: trph, goal: goals.trph },
    { actual: qa,   goal: goals.qa   },
    { actual: csat, goal: goals.csat }
  ].forEach((m, i) => {
    const ratio = m.goal > 0 ? m.actual / m.goal : 0;
    const bg    = ratio >= green ? DASH_C_GREEN : ratio >= yellow ? DASH_C_YELLOW : DASH_C_RED;
    sh.getRange(DASH_R_TEAM_PCT, 2 + i)
      .setValue(Math.round(ratio * 100) + '%')
      .setBackground(bg)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  });
}

function writeDashKpiReportCard_(sh, kpiStats) {
  const { avg, avgExcl, counts, autoFails, totalAgents } = kpiStats;
  const green  = Number(getConfigValue_('PACING_GREEN_MIN',  1.0));
  const yellow = Number(getConfigValue_('PACING_YELLOW_MIN', 0.9));
  const scaledGreen  = green  * 85; // treat 85% as "meeting" baseline for overall %
  const scaledYellow = yellow * 85;

  // Section header
  sh.getRange(DASH_R_KPI_HDR, 1, 1, 6).merge()
    .setValue('KPI REPORT CARD  (' + totalAgents + ' agents)')
    .setBackground(DASH_C_BLUE).setFontColor(DASH_C_WHITE)
    .setFontWeight('bold').setHorizontalAlignment('left');

  // Two averages side by side
  const avgBg     = avg     >= scaledGreen ? DASH_C_GREEN : avg     >= scaledYellow ? DASH_C_YELLOW : DASH_C_RED;
  const avgExclBg = avgExcl >= scaledGreen ? DASH_C_GREEN : avgExcl >= scaledYellow ? DASH_C_YELLOW : DASH_C_RED;

  sh.getRange(DASH_R_KPI_AVGS, 1, 1, 6).setValues([[
    'Overall Avg (all)', dashRound_(avg) + '%', '',
    'Overall Avg (excl. auto-fails)', dashRound_(avgExcl) + '%', ''
  ]]);
  sh.getRange(DASH_R_KPI_AVGS, 2).setBackground(avgBg).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(DASH_R_KPI_AVGS, 5).setBackground(avgExclBg).setFontWeight('bold').setHorizontalAlignment('center');

  // Status distribution
  sh.getRange(DASH_R_KPI_DIST, 1, 1, 6).merge()
    .setValue(
      'Exceeding: '   + (counts['Exceeding']   || 0) +
      '  ·  Meeting: '    + (counts['Meeting']    || 0) +
      '  ·  Close: '      + (counts['Close']      || 0) +
      '  ·  Not Meeting: '+ (counts['Not Meeting']|| 0) +
      '  ·  AUTO-FAIL: '  + (counts['AUTO-FAIL']  || 0)
    )
    .setFontColor(DASH_C_GRAY)
    .setHorizontalAlignment('left');

  // Auto-fail row
  if (autoFails.length) {
    sh.getRange(DASH_R_KPI_FAILS, 1, 1, 6).merge()
      .setValue('AUTO-FAIL: ' + autoFails.join(', '))
      .setBackground(DASH_C_RED)
      .setFontColor('#cc0000')
      .setFontWeight('bold');
  } else {
    sh.getRange(DASH_R_KPI_FAILS, 1, 1, 6).merge()
      .setValue('No auto-fails this week ✓')
      .setBackground(DASH_C_GREEN)
      .setFontColor('#274e13');
  }
}

function writeDashWoW_(sh, teamMetrics, kpiStats, wowMetrics) {
  const { prev, avg4 } = wowMetrics;
  const curr = {
    tcph:    teamMetrics.tcph,
    trph:    teamMetrics.trph,
    qa:      teamMetrics.qa,
    csat:    teamMetrics.csat,
    overall: kpiStats.avgExcl
  };

  // Section header
  sh.getRange(DASH_R_WOW_HDR, 1, 1, 6).merge()
    .setValue('WEEK OVER WEEK')
    .setBackground(DASH_C_BLUE).setFontColor(DASH_C_WHITE)
    .setFontWeight('bold').setHorizontalAlignment('left');

  // Column headers
  sh.getRange(DASH_R_WOW_COLS, 1, 1, 6)
    .setValues([['Metric', 'This Week', 'Last Week', 'Δ WoW', '4-Week Avg', 'Δ vs Avg']])
    .setFontWeight('bold').setBackground(DASH_C_BLUE_LIGHT)
    .setHorizontalAlignment('center');

  const rows = [
    { label: 'TCPH',         curr: curr.tcph,    prev: prev.tcph,    avg4: avg4.tcph,    fmt: v => dashRound_(v)        },
    { label: 'TRPH',         curr: curr.trph,    prev: prev.trph,    avg4: avg4.trph,    fmt: v => dashRound_(v)        },
    { label: 'QA Avg',       curr: curr.qa,      prev: prev.qa,      avg4: avg4.qa,      fmt: v => dashRound_(v) + '%'  },
    { label: 'CSAT',         curr: curr.csat,    prev: prev.csat,    avg4: avg4.csat,    fmt: v => dashRound_(v)        },
    { label: 'Overall %',    curr: curr.overall, prev: prev.overall, avg4: avg4.overall, fmt: v => dashRound_(v) + '%'  }
  ];

  rows.forEach((row, i) => {
    const r         = DASH_R_WOW_TCPH + i;
    const delta     = (row.prev && row.prev > 0) ? row.curr - row.prev    : null;
    const deltaAvg  = (row.avg4 && row.avg4 > 0) ? row.curr - row.avg4   : null;
    const fmtDelta  = v => v === null ? '--' : (v >= 0 ? '+' : '') + dashRound_(v);
    const arrow     = v => v === null ? '' : v >  0.001 ? ' ↑' : v < -0.001 ? ' ↓' : ' →';

    sh.getRange(r, 1, 1, 6).setValues([[
      row.label,
      row.fmt(row.curr),
      row.prev  ? row.fmt(row.prev)  : '--',
      fmtDelta(delta)    + arrow(delta),
      row.avg4  ? row.fmt(row.avg4)  : '--',
      fmtDelta(deltaAvg) + arrow(deltaAvg)
    ]]);
    sh.getRange(r, 1).setFontWeight('bold');
    sh.getRange(r, 2).setHorizontalAlignment('center');
    sh.getRange(r, 3).setHorizontalAlignment('center').setFontColor(DASH_C_GRAY);
    sh.getRange(r, 5).setHorizontalAlignment('center').setFontColor(DASH_C_GRAY);

    const deltaBg    = d => d === null ? null : d > 0.001 ? DASH_C_GREEN : d < -0.001 ? DASH_C_RED : DASH_C_WHITE;
    if (delta    !== null) sh.getRange(r, 4).setBackground(deltaBg(delta)).setHorizontalAlignment('center');
    if (deltaAvg !== null) sh.getRange(r, 6).setBackground(deltaBg(deltaAvg)).setHorizontalAlignment('center');
  });
}

function writeDashQALead_(sh, qaLeadName, qaStats) {
  sh.getRange(DASH_R_LEAD_HDR, 1, 1, 6).merge()
    .setValue('QA LEAD' + (qaLeadName ? '  —  ' + qaLeadName : ''))
    .setBackground(DASH_C_TEAL).setFontColor(DASH_C_WHITE)
    .setFontWeight('bold').setHorizontalAlignment('left');

  if (qaStats.completed === null) {
    sh.getRange(DASH_R_LEAD_DATA, 1, 1, 6).merge()
      .setValue('QA Lead Report Card data unavailable — set QA_LEAD_REPORT_CARD_ID in Config.')
      .setFontColor(DASH_C_GRAY)
      .setFontStyle('italic');
    return;
  }

  const pct    = qaStats.target > 0 ? qaStats.completed / qaStats.target : 0;
  const pctLbl = Math.round(pct * 100) + '%';
  const bg     = pct >= 1 ? DASH_C_GREEN : pct >= 0.8 ? DASH_C_YELLOW : DASH_C_RED;
  const status = pct >= 1 ? '✓ Met Goal' : pct >= 0.8 ? 'Near Goal' : 'Below Goal';

  sh.getRange(DASH_R_LEAD_DATA, 1, 1, 6).setValues([[
    'Completed / Target',
    qaStats.completed + ' / ' + qaStats.target,
    pctLbl,
    'Days Met: '    + (qaStats.daysMet    !== null ? qaStats.daysMet    : '--'),
    'Days Missed: ' + (qaStats.daysMissed !== null ? qaStats.daysMissed : '--'),
    qaStats.shortfall ? 'Shortfall: ' + qaStats.shortfall : status
  ]]);

  sh.getRange(DASH_R_LEAD_DATA, 2).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(DASH_R_LEAD_DATA, 3).setBackground(bg).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(DASH_R_LEAD_DATA, 6).setBackground(bg).setHorizontalAlignment('center');
}

function writeTrendSectionHeaders_(sh) {
  sh.getRange(DASH_R_TREND_HDR, 1, 1, TD_TOTAL).merge()
    .setValue('TREND DATA  (chart source — one row per week, auto-updated)')
    .setBackground(DASH_C_GRAY).setFontColor(DASH_C_WHITE)
    .setFontWeight('bold').setHorizontalAlignment('left');

  sh.getRange(DASH_R_TREND_COLS, 1, 1, TD_TOTAL).setValues([[
    'Week', 'TCPH', 'TRPH', 'QA Avg', 'CSAT',
    'Overall % (all)', 'Overall % (excl auto-fails)',
    'QA Forms', 'QA Target',
    'TCPH Goal', 'TRPH Goal', 'QA Goal', 'CSAT Goal',
    'TCPH % Goal', 'TRPH % Goal', 'QA % Goal', 'CSAT % Goal', 'QA Lead % Goal',
    'Meghan QA % Goal'
  ]]).setFontWeight('bold').setBackground(DASH_C_GRAY_LIGHT);
}

// =============================================================================
// Trend chart — created once, auto-updates as trend rows are appended
// =============================================================================

function buildDashboardChart_(sh) {
  const maxRows   = 200; // accommodate future weeks
  const headerRow = DASH_R_TREND_COLS;
  const dataRows  = maxRows + 1; // header + data

  // Chart uses: Week label (col 1) + five % of goal series (cols 14-18) + Meghan QA (col 19)
  const weekRange      = sh.getRange(headerRow, TD_WEEK,           dataRows, 1);
  const tcphRange      = sh.getRange(headerRow, TD_TCPH_PCT,       dataRows, 1);
  const trphRange      = sh.getRange(headerRow, TD_TRPH_PCT,       dataRows, 1);
  const qaRange        = sh.getRange(headerRow, TD_QA_PCT,         dataRows, 1);
  const csatRange      = sh.getRange(headerRow, TD_CSAT_PCT,       dataRows, 1);
  const qaLeadRange    = sh.getRange(headerRow, TD_QA_LEAD_PCT,    dataRows, 1);
  const meghanQaRange  = sh.getRange(headerRow, TD_MEGHAN_QA_PCT,  dataRows, 1);

  const chart = sh.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(weekRange)
    .addRange(tcphRange)
    .addRange(trphRange)
    .addRange(qaRange)
    .addRange(csatRange)
    .addRange(qaLeadRange)
    .addRange(meghanQaRange)
    .setOption('title', 'Team Performance Trend (% of Goal)')
    .setOption('width',  700)
    .setOption('height', 360)
    .setOption('legend', { position: 'bottom' })
    .setOption('series', {
      0: { color: '#4285F4', lineWidth: 2, labelInLegend: 'TCPH'        },
      1: { color: '#EA4335', lineWidth: 2, labelInLegend: 'TRPH'        },
      2: { color: '#34A853', lineWidth: 2, labelInLegend: 'QA Avg'      },
      3: { color: '#FBBC04', lineWidth: 2, labelInLegend: 'CSAT'        },
      4: { color: '#7b4f9e', lineWidth: 2, labelInLegend: 'QA Lead',
           lineDashStyle: [4, 4] },
      5: { color: '#FF6D00', lineWidth: 2, labelInLegend: 'Meghan QA',
           lineDashStyle: [2, 2] }
    })
    .setOption('hAxis', { slantedText: true, slantedTextAngle: 45 })
    .setOption('vAxis', { title: '% of Goal', format: '0"%"', viewWindow: { min: 50 } })
    .setPosition(DASH_R_CHART_START, 1, 0, 0)
    .build();

  sh.insertChart(chart);
}

// =============================================================================
// Pulse Log — EOD inbox health table
// =============================================================================

/**
 * Reads EOD (last entry of each day) pulse log data for the given week.
 * Computes cumulative aging buckets: over18+, over12+, over8+.
 * Source: STAFFING_PULSE_LOG_SPREADSHEET_ID → 'WoW Summary' tab.
 */
function collectPulseLogData_(monday) {
  const spreadsheetId = String(getConfigValue_('STAFFING_PULSE_LOG_SPREADSHEET_ID', '') || '').trim();
  if (!spreadsheetId) return [];

  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sh = ss.getSheetByName('WoW Summary');
    if (!sh || sh.getLastRow() < 2) return [];

    const lastRow = sh.getLastRow();
    const rows    = sh.getRange(2, 1, lastRow - 1, 26).getValues();

    // Group by date in Phoenix timezone — keep the last (EOD) row per date
    const dateMap = {};
    rows.forEach(row => {
      const ts = row[PL_TIMESTAMP];
      if (!(ts instanceof Date) || isNaN(ts)) return;
      const key = dashDateKey_(ts);
      if (!key) return;
      if (!dateMap[key] || ts > dateMap[key][PL_TIMESTAMP]) dateMap[key] = row;
    });

    // Build one entry per day of the week
    const result = [];
    for (let i = 0; i < 7; i++) {
      const day = addDaysLocal_(monday, i);
      const key = dashDateKey_(day);
      const row = dateMap[key];

      if (!row) { result.push({ date: day, hasData: false }); continue; }

      const over24     = num_(row[PL_OVER_24]);
      const over18plus = over24 + num_(row[PL_BKT_24]) + num_(row[PL_BKT_22]) +
                         num_(row[PL_BKT_20]) + num_(row[PL_BKT_18]);
      const over12plus = over18plus + num_(row[PL_BKT_16]) + num_(row[PL_BKT_14]) +
                         num_(row[PL_BKT_12]);
      const over8plus  = over12plus + num_(row[PL_BKT_10]) + num_(row[PL_BKT_8]);

      result.push({
        date: day, hasData: true,
        totalOpen:  num_(row[PL_TOTAL_OPEN]),
        unassigned: num_(row[PL_UNASSIGNED]),
        over24, over18plus, over12plus, over8plus
      });
    }
    return result;

  } catch (e) {
    Logger.log('collectPulseLogData_ error: ' + e.message);
    return [];
  }
}

/**
 * Writes the EOD inbox health table to the right side of the dashboard
 * starting at column J (DASH_PULSE_COL), aligned with the top sections.
 */
function writeDashPulseLog_(sh, pulseData) {
  const c = DASH_PULSE_COL;
  sh.getRange(DASH_PULSE_ROW_HDR, c, 12, DASH_PULSE_NCOLS).clearContent().clearFormat();

  // Title
  sh.getRange(DASH_PULSE_ROW_HDR, c, 1, DASH_PULSE_NCOLS).merge()
    .setValue('INBOX HEALTH  (EOD Snapshot)')
    .setBackground(DASH_C_BLUE).setFontColor(DASH_C_WHITE)
    .setFontWeight('bold').setHorizontalAlignment('left');

  if (!pulseData || pulseData.length === 0) {
    sh.getRange(DASH_PULSE_ROW_COLS, c, 1, DASH_PULSE_NCOLS).merge()
      .setValue('Pulse Log unavailable — set STAFFING_PULSE_LOG_SPREADSHEET_ID in Config.')
      .setFontColor(DASH_C_GRAY).setFontStyle('italic');
    return;
  }

  // Column headers
  sh.getRange(DASH_PULSE_ROW_COLS, c, 1, DASH_PULSE_NCOLS)
    .setValues([['Day', 'Total Open', 'Unassigned', 'Over 24hr', 'Over 18hr+', 'Over 12hr+', 'Over 8hr+']])
    .setFontWeight('bold').setBackground(DASH_C_BLUE_LIGHT)
    .setHorizontalAlignment('center');

  // Color thresholds for aging buckets
  const agingBg = (val, warn, crit) =>
    val === 0 ? DASH_C_GREEN : val <= warn ? DASH_C_YELLOW : DASH_C_RED;

  pulseData.forEach((d, i) => {
    const row      = DASH_PULSE_ROW_DATA + i;
    const dayLabel = Utilities.formatDate(d.date, CFG.timezone, 'EEE M/d');
    const altBg    = i % 2 === 1 ? DASH_C_GRAY_LIGHT : DASH_C_WHITE;

    if (!d.hasData) {
      sh.getRange(row, c, 1, DASH_PULSE_NCOLS)
        .setValues([[dayLabel, '--', '--', '--', '--', '--', '--']])
        .setBackground(altBg);
      sh.getRange(row, c).setFontWeight('bold');
      sh.getRange(row, c + 1, 1, DASH_PULSE_NCOLS - 1).setFontColor(DASH_C_GRAY).setHorizontalAlignment('center');
      return;
    }

    sh.getRange(row, c, 1, DASH_PULSE_NCOLS)
      .setValues([[dayLabel, d.totalOpen, d.unassigned, d.over24, d.over18plus, d.over12plus, d.over8plus]])
      .setBackground(altBg);

    sh.getRange(row, c).setFontWeight('bold');
    sh.getRange(row, c + 1, 1, DASH_PULSE_NCOLS - 1).setHorizontalAlignment('center');

    // Conditional coloring: over24 is strictest (any > 0 is a concern)
    sh.getRange(row, c + 3).setBackground(agingBg(d.over24,     3,  10));
    sh.getRange(row, c + 4).setBackground(agingBg(d.over18plus, 8,  20));
    sh.getRange(row, c + 5).setBackground(agingBg(d.over12plus, 20, 50));
    sh.getRange(row, c + 6).setBackground(agingBg(d.over8plus,  40, 80));
  });
}

// =============================================================================
// Utilities
// =============================================================================

function dashRound_(v) {
  return Math.round(Number(v) * 100) / 100;
}

function dashWeekLabel_(monday, sunday) {
  const fmt = d => Utilities.formatDate(d, CFG.timezone, 'MMM d');
  return fmt(monday) + ' – ' + fmt(sunday);
}

function num_(v) {
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}
