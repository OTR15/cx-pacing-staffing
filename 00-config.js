// =============================================================================
// config.gs
// Central configuration object and default roster data.
// All tunable constants live here — no need to touch other files for most
// day-to-day adjustments (goals, checkpoints, sheet names, etc.).
// =============================================================================

const CFG = {
  // ── Gorgias ────────────────────────────────────────────────────────────────
  subdomain: 'oatsovernight',
  timezone: 'America/Phoenix',

  // ── Sheet names ────────────────────────────────────────────────────────────
  scheduleSheetName: 'Schedule',
  rosterSheetName: 'Roster',
  goalsSheetName: 'Goals',
  configSheetName: 'Config',
  normalizedScheduleSheetName: 'Schedule_Normalized',
  teamGuideSheetName: 'Team Guide',
  caseUseSummarySheetName: 'Case Use Summary',

  // ── Schedule layout (1-based column/row positions) ─────────────────────────
  schedule: {
    headerRow1: 1,
    headerRow2: 2,
    firstDataRow: 3,
    nameCol: 1,
    managerCol: 3,
    firstDateCol: 4
  },

  // ── Weekly report settings ─────────────────────────────────────────────────
  weekly: {
    visibleTabCount: 4,    // How many weekly tabs to keep visible
    tabColor: '#9fc5f8',
    chartStartCol: 24,
    kpiSnapshot: {
      spreadsheetId: '1kNBJtMMsNo3W13fnIbkCBcCbHdzgOCNk24eA1w_zPq4',
      adminSheetName: 'ADMIN_VIEW',
      tableStartRow: 32,
      chartStartRow: 56
    }
  },

  // ── Daily sheet settings ───────────────────────────────────────────────────
  daily: {
    firstDataRow: 3,
    metricLabels: ['Closed Tickets', 'Tickets Replied', 'Messages Sent', 'CSAT'],
    progressLabels: ['On Track', 'On a Project', 'Actions Taken', 'EOD Goal Met', 'Notes']
  },

  // ── Shift / goal calculation ───────────────────────────────────────────────
  standardShiftHours: 8,

  baselineGoals: {
    closedTickets: 55,
    ticketsReplied: 65,
    messagesSent: 80,
    csat: 4.7
  },

  // ── Checkpoint definitions ─────────────────────────────────────────────────
  // percent = fraction of daily goal expected by this checkpoint
  checkpoints: [
    { key: '7AM',  label: '7am Pacing Report',  hour: 7,  minute: 0, percent: 0.10 },
    { key: '9AM',  label: '9am Pacing Report',  hour: 9,  minute: 0, percent: 0.25 },
    { key: '11AM', label: '11am Pacing Report', hour: 11, minute: 0, percent: 0.40 },
    { key: '2PM',  label: '2pm Pacing Report',  hour: 14, minute: 0, percent: 0.60 },
    { key: '6PM',  label: '6pm Pacing Report',  hour: 18, minute: 0, percent: 0.85 },
    { key: 'EOD',  label: 'EOD Report',         hour: 23, minute: 0, percent: 1.00 }
  ],

  // ── Admin / utility sheets to hide from normal users ──────────────────────
  hiddenSheetNames: ['Config', 'Roster', 'Goals', 'Schedule_Normalized', 'Schedule'],

  // ── Agents to exclude from pacing reports (bots, test accounts, etc.) ──────
  excludedAgents: [
    'AI Agent Bot',
    'AI Agent- Bot',
    'ChargeFlow',
    'Digital Genius',
    'Gorgias Bot',
    'Gorgias Contact Form Bot',
    'Gorgias Convert Bot',
    'Gorgias Help Center Bot',
    'Gorgias Help Center- Bot',
    'Gorgias Helpdesk Bot',
    'Gorgias Helpdesk- Bot',
    'Gorgias Mobile Bot',
    'Gorgias Support Agent',
    'Gorgias Workflows Bot',
    'Gorgias Workflows- Bot',
    'Antecedes',
    'Fraser',
    'Isaac',
    'Quint Test'
  ],

  staffing: {
    sheetName: 'Staffing',
    ticketsPerProductiveHour: 8,
    agedRiskWeight: 1.5,
    reserveHoursBuffer: 11,
    minimumAgentsFloor: 4,
    cautionUnassignedThreshold: 12,
    estimatedWorkableTicketsPerHour: 50,
    endOfDayHour: 18,
    pulseLogSpreadsheetId: '1ozcKrCo_wgFLqt1F8oIC7otR16wfij8huZlJ5Yv-L_w',
    workableVolumeLogSheetName: 'Workable Volume Log',
    overnightInflowLogSheetName: 'Overnight Inflow Log',
    observedDataBlendWeight: 0,
    useObservedData: false,
    shadowModelEnabled: true,
    minimumObservedSampleDays: 10,
    workableRateMultiplier: 1,
    sendHomeBufferMultiplier: 1
  }
};

// =============================================================================
// Default roster — used to seed the Roster sheet on first setup.
// After setup, the Roster sheet is the source of truth; this array is only
// used when running "Setup: Seed Project" on a fresh spreadsheet.
// Format: [Gorgias Agent ID, Display Name]
// =============================================================================
const DEFAULT_ROSTER = [
  [628557002, 'adam steen'],
  [834693217, 'AJ Klekas'],
  [814613981, 'Cassidy Klekas'],
  [634702294, 'Claryssa Graeve'],
  [891598544, 'Desmayia Johnson'],
  [841069273, 'Destiny Cook'],
  [628556021, 'Elley Sieck'],
  [628556223, 'Emma Mayberry'],
  [628557297, 'Janiecea Allison'],
  [841069832, 'Kahmeel Ballard'],
  [628556442, 'Kayla Barnett'],
  [631111319, 'KeOsha Williams'],
  [841070393, 'Kim Williams'],
  [637347178, 'Quel Rodriguez'],
  [632667318, 'Rebecca Kolesar'],
  [628556645, 'Santos Porter'],
  [628557574, 'Vio Renteria']
];

// =============================================================================
// Config sheet helpers
// These are placed here (rather than utils.gs) because they depend on CFG
// and are referenced by almost every other file at startup.
// =============================================================================

/**
 * Reads a value from the Config sheet by key.
 * Falls back to `fallback` if the sheet or key is missing.
 *
 * @param {string} key     - The key name in column A of the Config sheet.
 * @param {*}      fallback - Value to return if the key is not found.
 * @returns {*} The stored value, or fallback.
 */
function getConfigValue_(key, fallback) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.configSheetName);
  if (!sh) return fallback;

  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === key) return values[i][1];
  }
  return fallback;
}

/**
 * Writes a value to the Config sheet by key.
 * Appends a new row if the key does not already exist.
 *
 * @param {string} key   - The key name in column A.
 * @param {*}      value - The value to store in column B.
 */
function setConfigValue_(key, value) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.configSheetName);
  if (!sh) throw new Error('Config sheet missing.');

  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === key) {
      sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }

  sh.appendRow([key, value]);
}
