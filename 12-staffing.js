// =============================================================================
// staffing.gs
// Staffing recommendation helpers.
// =============================================================================

/**
 * Returns staffing assumptions, using CFG.staffing defaults with Config sheet
 * overrides when present.
 *
 * @returns {{
 *   ticketsPerProductiveHour: number,
 *   agedRiskWeight: number,
 *   reserveHoursBuffer: number,
 *   minimumAgentsFloor: number,
 *   cautionUnassignedThreshold: number,
 *   estimatedWorkableTicketsPerHour: number,
 *   endOfDayHour: number,
 *   pulseLogSpreadsheetId: string,
 *   workableVolumeLogSheetName: string,
 *   overnightInflowLogSheetName: string,
 *   observedDataBlendWeight: number,
 *   useObservedData: boolean,
 *   shadowModelEnabled: boolean,
 *   minimumObservedSampleDays: number,
 *   workableRateMultiplier: number,
 *   sendHomeBufferMultiplier: number
 * }}
 */
function getStaffingAssumptions_() {
  const defaults = CFG.staffing || {};

  return {
    ticketsPerProductiveHour: Number(getConfigValue_(
      'STAFFING_TICKETS_PER_PRODUCTIVE_HOUR',
      defaults.ticketsPerProductiveHour
    )) || defaults.ticketsPerProductiveHour,

    agedRiskWeight: Number(getConfigValue_(
      'STAFFING_AGED_RISK_WEIGHT',
      defaults.agedRiskWeight
    )) || defaults.agedRiskWeight,

    reserveHoursBuffer: Number(getConfigValue_(
      'STAFFING_RESERVE_HOURS_BUFFER',
      defaults.reserveHoursBuffer
    )) || defaults.reserveHoursBuffer,

    minimumAgentsFloor: Number(getConfigValue_(
      'STAFFING_MINIMUM_AGENTS_FLOOR',
      defaults.minimumAgentsFloor
    )) || defaults.minimumAgentsFloor,

    cautionUnassignedThreshold: Number(getConfigValue_(
      'STAFFING_CAUTION_UNASSIGNED_THRESHOLD',
      defaults.cautionUnassignedThreshold
    )) || defaults.cautionUnassignedThreshold,

    estimatedWorkableTicketsPerHour: Number(getConfigValue_(
      'STAFFING_ESTIMATED_WORKABLE_TICKETS_PER_HOUR',
      defaults.estimatedWorkableTicketsPerHour || 40
    )) || defaults.estimatedWorkableTicketsPerHour || 40,

    endOfDayHour: Number(getConfigValue_(
      'STAFFING_END_OF_DAY_HOUR',
      defaults.endOfDayHour || 18
    )) || defaults.endOfDayHour || 18,

    pulseLogSpreadsheetId: String(getConfigValue_(
      'STAFFING_PULSE_LOG_SPREADSHEET_ID',
      defaults.pulseLogSpreadsheetId || ''
    ) || '').trim(),

    workableVolumeLogSheetName: String(getConfigValue_(
      'STAFFING_WORKABLE_VOLUME_LOG_SHEET_NAME',
      defaults.workableVolumeLogSheetName || 'Workable Volume Log'
    ) || defaults.workableVolumeLogSheetName || 'Workable Volume Log').trim(),

    overnightInflowLogSheetName: String(getConfigValue_(
      'STAFFING_OVERNIGHT_INFLOW_LOG_SHEET_NAME',
      defaults.overnightInflowLogSheetName || 'Overnight Inflow Log'
    ) || defaults.overnightInflowLogSheetName || 'Overnight Inflow Log').trim(),

    observedDataBlendWeight: Number(getConfigValue_(
      'STAFFING_OBSERVED_DATA_BLEND_WEIGHT',
      defaults.observedDataBlendWeight || 0
    )) || defaults.observedDataBlendWeight || 0,

    useObservedData: String(getConfigValue_(
      'STAFFING_USE_OBSERVED_DATA',
      defaults.useObservedData
    )).toLowerCase() === 'true',

    shadowModelEnabled: String(getConfigValue_(
      'STAFFING_SHADOW_MODEL_ENABLED',
      defaults.shadowModelEnabled
    )).toLowerCase() === 'true',

    minimumObservedSampleDays: Number(getConfigValue_(
      'STAFFING_MINIMUM_OBSERVED_SAMPLE_DAYS',
      defaults.minimumObservedSampleDays || 10
    )) || defaults.minimumObservedSampleDays || 10,

    workableRateMultiplier: Number(getConfigValue_(
      'STAFFING_WORKABLE_RATE_MULTIPLIER',
      defaults.workableRateMultiplier || 1
    )) || defaults.workableRateMultiplier || 1,

    sendHomeBufferMultiplier: Number(getConfigValue_(
      'STAFFING_SEND_HOME_BUFFER_MULTIPLIER',
      defaults.sendHomeBufferMultiplier || 1
    )) || defaults.sendHomeBufferMultiplier || 1
  };
}

/**
 * Estimates remaining inflow from a checkpoint to the configured end of day.
 *
 * @param {{ hour: number, minute: number }} checkpoint
 * @returns {number}
 */
function estimateRemainingInflow_(checkpoint) {
  const assumptions = getStaffingAssumptions_();
  const checkpointMinutes =
    (Number(checkpoint.hour || 0) * 60) + Number(checkpoint.minute || 0);
  const endMinutes = Number(assumptions.endOfDayHour || 0) * 60;

  const hoursRemaining = Math.max(0, (endMinutes - checkpointMinutes) / 60);

  return Number(assumptions.estimatedWorkableTicketsPerHour || 0) * hoursRemaining;
}

/**
 * Returns projected work remaining in productive hours for a checkpoint.
 * unassigned is intentionally excluded here to avoid double-counting it on top
 * of totalOpen; it should be handled separately as a caution/block signal.
 *
 * Formula:
 *   (totalOpen + (agedRisk * agedRiskWeight) + estimatedInflow) / ticketsPerProductiveHour
 *
 * @param {{
*   totalOpen: number,
*   agedRisk: number,
*   estimatedInflow: number
* }} pulseInput
* @returns {number}
*/
function computeProjectedWorkRemaining_(pulseInput) {
 const assumptions = getStaffingAssumptions_();
 const totalOpen = Number((pulseInput || {}).totalOpen || 0);
 const agedRisk = Number((pulseInput || {}).agedRisk || 0);
 const estimatedInflow = Number((pulseInput || {}).estimatedInflow || 0);

 const weightedAgedRisk = agedRisk * assumptions.agedRiskWeight;
 const tph = assumptions.ticketsPerProductiveHour;

 if (tph <= 0) return 0;

 return (totalOpen + weightedAgedRisk + estimatedInflow) / tph;
}

/**
 * Returns projected capacity remaining in productive hours.
 * Input is already in hours, so no conversion is applied.
 *
 * @param {number} remainingProductiveHours
 * @returns {number}
 */
function computeProjectedCapacityRemaining_(remainingProductiveHours) {
 const hours = Number(remainingProductiveHours);
 return isNaN(hours) ? 0 : hours;
}

/**
 * Returns excess capacity in productive hours.
 *
 * @param {number} projectedCapacityRemaining
 * @param {number} projectedWorkRemaining
 * @returns {number}
 */
function computeExcessCapacity_(projectedCapacityRemaining, projectedWorkRemaining) {
 const capacity = Number(projectedCapacityRemaining);
 const work = Number(projectedWorkRemaining);

 return (isNaN(capacity) ? 0 : capacity) - (isNaN(work) ? 0 : work);
}

/**
 * Returns the recommended number of agents that can be sent home safely.
 *
 * @param {number} excessCapacity
 * @param {number} activeAgentCount
 * @param {number} remainingProductiveHours
 * @returns {number}
 */
function computeRecommendedSendHomeCount_(excessCapacity, activeAgentCount, remainingProductiveHours) {
 const assumptions = getStaffingAssumptions_();
 const excess = Number(excessCapacity);
 const active = Number(activeAgentCount);
 const hours = Number(remainingProductiveHours);

 const safeExcess = isNaN(excess) ? 0 : excess;
 const safeActive = isNaN(active) ? 0 : active;
 const safeHours = isNaN(hours) ? 0 : hours;
 const minimumAgentsFloor = Number(assumptions.minimumAgentsFloor || 0);
 const reserveHoursBuffer = Number(assumptions.reserveHoursBuffer || 0);

 if (safeActive <= 0) return 0;
 if (safeActive <= minimumAgentsFloor) return 0;

 const avgRemainingHoursPerAgent = safeHours / safeActive;
 if (avgRemainingHoursPerAgent <= 0) return 0;

 const sendableHours = safeExcess - reserveHoursBuffer;
 if (sendableHours <= 0) return 0;

 const rawSendHome = Math.floor(sendableHours / avgRemainingHoursPerAgent);
 const maxAllowed = safeActive - minimumAgentsFloor;

 return Math.max(0, Math.min(rawSendHome, maxAllowed));
}

/**
 * Returns the recommendation status for a checkpoint staffing row.
 *
 * @param {number} recommendedSendHomeCount
 * @param {number} excessCapacity
 * @param {number} unassigned
 * @param {number} activeAgentCount
 * @returns {'BLOCK'|'HOLD'|'CAUTION'|'SEND'}
 */
function getRecommendationStatus_(recommendedSendHomeCount, excessCapacity, unassigned, activeAgentCount) {
 const assumptions = getStaffingAssumptions_();
 const sendHomeCount = Number(recommendedSendHomeCount);
 const excess = Number(excessCapacity);
 const unassignedCount = Number(unassigned);
 const active = Number(activeAgentCount);

 const safeSendHomeCount = isNaN(sendHomeCount) ? 0 : sendHomeCount;
 const safeExcess = isNaN(excess) ? 0 : excess;
 const safeUnassignedCount = isNaN(unassignedCount) ? 0 : unassignedCount;
 const safeActive = isNaN(active) ? 0 : active;
 const minimumAgentsFloor = Number(assumptions.minimumAgentsFloor || 0);
 const cautionUnassignedThreshold = Number(assumptions.cautionUnassignedThreshold || 0);

 if (safeExcess < 0) return 'BLOCK';
 if (safeActive <= minimumAgentsFloor) return 'BLOCK';
 if (safeUnassignedCount > cautionUnassignedThreshold) return 'CAUTION';
 if (safeSendHomeCount >= 1) return 'SEND';
 return 'HOLD';
}

/**
 * Returns a short human-readable explanation for a staffing recommendation.
 *
 * @param {{
 *   projectedCapacityRemaining: number,
 *   projectedWorkRemaining: number,
 *   excessCapacity: number,
 *   unassigned: number,
 *   recommendedSendHomeCount: number
 * }} row
 * @returns {string}
 */
function buildRecommendationExplanation_(row) {
  const capacity = Number((row || {}).projectedCapacityRemaining || 0);
  const work = Number((row || {}).projectedWorkRemaining || 0);
  const excess = Number((row || {}).excessCapacity || 0);
  const unassigned = Number((row || {}).unassigned || 0);
  const sendHomeCount = Number((row || {}).recommendedSendHomeCount || 0);
  const status = String((row || {}).recommendationStatus || '');

  const parts = [
    'Capacity ' + round1_(capacity) + 'h vs work ' + round1_(work) + 'h',
    'excess ' + round1_(excess) + 'h'
  ];

  if (unassigned > 0) {
    parts.push('unassigned=' + unassigned);
  }

  if (status === 'SEND' && sendHomeCount >= 1) {
    parts.push('Send ' + sendHomeCount);
  } else if (status === 'BLOCK') {
    parts.push('Block');
  } else if (status === 'CAUTION') {
    parts.push('Caution');
  } else {
    parts.push('Hold');
  }

  return parts.join(' | ');
}

/**
 * Returns a formatted checkpoint timestamp in the project timezone.
 *
 * @param {Date} dateObj
 * @param {{ hour: number, minute: number }} checkpoint
 * @returns {string}
 */
function getCheckpointTimestampText_(dateObj, checkpoint) {
  const d = new Date(dateObj);

  d.setHours(checkpoint.hour);
  d.setMinutes(checkpoint.minute || 0);
  d.setSeconds(0);
  d.setMilliseconds(0);

  return Utilities.formatDate(d, CFG.timezone, 'M/d/yy h:mm a');
}

/**
 * Returns remaining shift hours from a checkpoint time to shift end.
 * First-pass version handles same-day shifts only.
 *
 * @param {Date} dateObj
 * @param {{ hour: number, minute: number }} checkpoint
 * @param {{ startText: string, endText: string }} schedule
 * @returns {number}
 */
function getRemainingShiftHoursForScheduleAtCheckpoint_(dateObj, checkpoint, schedule) {
  const startText = String((schedule || {}).startText || '').trim();
  const endText = String((schedule || {}).endText || '').trim();

  if (!startText || !endText) return 0;

  const checkpointMinutes = (Number(checkpoint.hour || 0) * 60) + Number(checkpoint.minute || 0);
  const startMinutes = parseTimeToMinutes_(startText);
  const endMinutes = parseTimeToMinutes_(endText);

  if (isNaN(startMinutes) || isNaN(endMinutes)) return 0;
  if (endMinutes <= startMinutes) return 0;
  if (checkpointMinutes >= endMinutes) return 0;

  const effectiveStart = checkpointMinutes <= startMinutes ? startMinutes : checkpointMinutes;
  return (endMinutes - effectiveStart) / 60;
}

/**
 * Returns active staffing coverage at a checkpoint for the roster.
 * First-pass version applies a simple lunch deduction for longer shifts by
 * scaling remaining effective hours down proportionally.
 *
 * @param {Date} dateObj
 * @param {{ hour: number, minute: number }} checkpoint
 * @param {Array<{ agentId: number, repName: string }>} roster
 * @param {Object} scheduleMap
 * @param {Object} assumptions
 * @returns {{
 *   activeAgentCount: number,
 *   activeReps: Array<{
 *     agentId: number,
 *     repName: string,
 *     schedule: Object,
 *     remainingShiftHours: number,
 *     remainingEffectiveHours: number
 *   }>
 * }}
 */
function getCheckpointActiveCoverage_(dateObj, checkpoint, roster, scheduleMap, assumptions) {
  const activeReps = [];

  (roster || []).forEach(rep => {
    const schedule = getScheduleForRep_(scheduleMap, rep.repName);
    if (!schedule) return;
    if (['CTO', 'VTO', 'Off'].includes(schedule.status)) return;

    const remainingShiftHours = getRemainingShiftHoursForScheduleAtCheckpoint_(
      dateObj,
      checkpoint,
      schedule
    );

    if (remainingShiftHours <= 0) return;

    const scheduledHours = Number(schedule.hours || 0);
    const effectiveShiftHours =
      scheduledHours >= 9 ? Math.max(0, scheduledHours - 1) : scheduledHours;

    const remainingEffectiveHours = scheduledHours > 0
      ? Math.max(0, remainingShiftHours * (effectiveShiftHours / scheduledHours))
      : Math.max(0, remainingShiftHours);

    activeReps.push({
      agentId: rep.agentId,
      repName: rep.repName,
      schedule,
      remainingShiftHours,
      remainingEffectiveHours
    });
  });

  return {
    activeAgentCount: activeReps.length,
    activeReps
  };
  
}

function getRemainingProductiveHoursAtCheckpoint_(activeCoverage) {
  const activeReps = ((activeCoverage || {}).activeReps) || [];
  if (!activeReps.length) return 0;

  return activeReps.reduce((sum, rep) => {
    const hours = Number((rep || {}).remainingEffectiveHours);
    return sum + (isNaN(hours) ? 0 : hours);
  }, 0);
}

/**
 * Builds a single checkpoint staffing row.
 *
 * @param {Date} dateObj
 * @param {{ key: string, label: string, hour: number, minute: number }} checkpoint
 * @param {{
 *   pulseInputs: Object,
 *   roster: Array,
 *   scheduleMap: Object,
 *   assumptions: Object
 * }} context
 * @returns {{
 *   checkpointKey: string,
 *   checkpointLabel: string,
 *   checkpointTimestamp: string,
 *   totalOpen: number,
 *   unassigned: number,
 *   agedRisk: number,
 *   estimatedInflow: number,
 *   activeAgentCount: number,
 *   remainingProductiveHours: number,
 *   projectedWorkRemaining: number,
 *   projectedCapacityRemaining: number,
 *   excessCapacity: number,
 *   recommendedSendHomeCount: number,
 *   recommendationStatus: string,
 *   explanation: string
 * }}
 */
function buildCheckpointStaffingRow_(dateObj, checkpoint, context) {
  const pulseInputs = ((context || {}).pulseInputs || {}).byCheckpointKey || {};
  const pulseInput = pulseInputs[checkpoint.key] || {};
  const totalOpen = Number(pulseInput.totalOpen || 0);
  const unassigned = Number(pulseInput.unassigned || 0);
  const agedRisk = Number(pulseInput.agedRisk || 0);
  const manualEstimatedInflow = Number(pulseInput.estimatedInflow || 0);
  const estimatedInflow = manualEstimatedInflow > 0
    ? manualEstimatedInflow
    : estimateRemainingInflow_(checkpoint);

  // Stubbed for first pass; replace with schedule-driven coverage logic later.
  const coverage = getCheckpointActiveCoverage_(
    dateObj,
    checkpoint,
    context.roster,
    context.scheduleMap,
    context.assumptions
  );
  
  const activeAgentCount = coverage.activeAgentCount;
  
  const remainingProductiveHours =
    getRemainingProductiveHoursAtCheckpoint_(coverage);

  const projectedWorkRemaining = computeProjectedWorkRemaining_({
    totalOpen,
    agedRisk,
    estimatedInflow
  });

  const projectedCapacityRemaining =
    computeProjectedCapacityRemaining_(remainingProductiveHours);

  const excessCapacity = computeExcessCapacity_(
    projectedCapacityRemaining,
    projectedWorkRemaining
  );

  const recommendedSendHomeCount = computeRecommendedSendHomeCount_(
    excessCapacity,
    activeAgentCount,
    remainingProductiveHours
  );

  const recommendationStatus = getRecommendationStatus_(
    recommendedSendHomeCount,
    excessCapacity,
    unassigned,
    activeAgentCount
  );

  const row = {
    checkpointKey: checkpoint.key,
    checkpointLabel: checkpoint.label,
    checkpointTimestamp: getCheckpointTimestampText_(dateObj, checkpoint),

    totalOpen,
    unassigned,
    agedRisk,
    estimatedInflow,

    activeAgentCount,
    remainingProductiveHours,

    projectedWorkRemaining,
    projectedCapacityRemaining,
    excessCapacity,

    recommendedSendHomeCount,
    recommendationStatus,
    explanation: ''
  };

  row.explanation = buildRecommendationExplanation_(row);
  return row;
}

/**
 * Builds staffing rows for all configured checkpoints on a date.
 *
 * @param {Date} dateObj
 * @param {Object} inputs
 * @returns {Array<Object>}
 */
function buildCheckpointStaffingRows_(dateObj, inputs) {
  return CFG.checkpoints.map(checkpoint => {
    return buildCheckpointStaffingRow_(dateObj, checkpoint, inputs);
  });
}

/**
 * Builds the staffing model for a given date.
 *
 * @param {Date} dateObj
 * @returns {{
 *   dateObj: Date,
 *   dateLabel: string,
 *   assumptions: Object,
 *   pulseInputs: Object,
 *   observedMetrics: Object,
 *   rows: Array<Object>
 * }}
 */
function buildStaffingModelForDate_(dateObj) {
  const inputs = getStaffingInputsForDate_(dateObj);
  const rows = buildCheckpointStaffingRows_(dateObj, inputs);

  return {
    dateObj,
    dateLabel: inputs.dateLabel,
    assumptions: inputs.assumptions,
    pulseInputs: inputs.pulseInputs,
    observedMetrics: inputs.observedMetrics,
    rows
  };
}

/**
 * Normalizes checkpoint keys from sheet input values.
 *
 * @param {*} value
 * @returns {string}
 */
function normalizeCheckpointKey_(value) {
  const text = String(value || '').trim().toUpperCase();

  const map = {
    '7AM': '7AM',
    '7:00 AM': '7AM',
    '9AM': '9AM',
    '9:00 AM': '9AM',
    '11AM': '11AM',
    '11:00 AM': '11AM',
    '2PM': '2PM',
    '2:00 PM': '2PM',
    '6PM': '6PM',
    '6:00 PM': '6PM',
    'EOD': 'EOD'
  };

  return map[text] || text;
}

/**
 * Returns pulse inputs for a date, keyed by checkpoint.
 * Reads the input block starting at row 5:
 *   Checkpoint | totalOpen | unassigned | agedRisk | estimatedInflow
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Date} dateObj
 * @returns {{ byCheckpointKey: Object<string, {
*   totalOpen: number,
*   unassigned: number,
*   agedRisk: number,
*   estimatedInflow: number
* }>}}
*/
function getPulseInputsForDate_(sheet, dateObj) {
  const byCheckpointKey = {};

  CFG.checkpoints.forEach(checkpoint => {
    byCheckpointKey[checkpoint.key] = {
      totalOpen: 0,
      unassigned: 0,
      agedRisk: 0,
      estimatedInflow: 0
    };
  });

  if (!sheet) return { byCheckpointKey };

  const values = sheet.getRange(5, 1, 6, 5).getDisplayValues();

  values.forEach(row => {
    const checkpointKey = normalizeCheckpointKey_(row[0]);
    if (!checkpointKey || !byCheckpointKey[checkpointKey]) return;

    byCheckpointKey[checkpointKey] = {
      totalOpen: Number(row[1]) || 0,
      unassigned: Number(row[2]) || 0,
      agedRisk: Number(row[3]) || 0,
      estimatedInflow: Number(row[4]) || 0
    };
  });
  return { byCheckpointKey };
}

function getDateTextForTimeZone_(dateObj, timezone, format) {
  return Utilities.formatDate(dateObj, timezone || CFG.timezone, format || 'yyyy-MM-dd');
}

function safeNumberOrBlank_(value) {
  if (value === '' || value === null || value === undefined) return '';
  const num = Number(value);
  return isNaN(num) ? '' : num;
}

function averageNumbers_(values) {
  const nums = (values || [])
    .filter(value => value !== '' && value !== null && value !== undefined)
    .map(value => Number(value))
    .filter(value => !isNaN(value));
  if (!nums.length) return '';
  const total = nums.reduce((sum, value) => sum + value, 0);
  return round2_(total / nums.length);
}

function readSheetRecordsByHeader_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0].map(header => String(header || '').trim());

  return values.slice(1).map(row => {
    const record = {};
    headers.forEach((header, index) => {
      record[header] = row[index];
    });
    return record;
  });
}

function getPulseLogSpreadsheet_(assumptions) {
  const spreadsheetId = String((assumptions || {}).pulseLogSpreadsheetId || '').trim();
  if (!spreadsheetId) return null;

  try {
    return SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    Logger.log('Unable to open Pulse Log spreadsheet: ' + e);
    return null;
  }
}

function getObservedWorkableMetricsForDate_(dateObj, assumptions) {
  const timezone = CFG.timezone;
  const dateText = getDateTextForTimeZone_(dateObj, timezone, 'yyyy-MM-dd');
  const pulseSs = getPulseLogSpreadsheet_(assumptions);

  const fallback = {
    status: pulseSs ? 'missing_log_tabs' : 'missing_pulse_log',
    dateText: dateText,
    latestHourlyTimestampAz: '',
    latestHourlyWorkableOpenInbox: '',
    latestHourlyWorkableClosed: '',
    latestHourlyWorkableTotalVolume: '',
    latestHourlySourceStatus: '',
    overnightBusinessDateAz: '',
    overnightWorkableInflowProxy: '',
    overnightSourceStatus: '',
    hourlySampleCount7d: 0,
    overnightSampleCount7d: 0,
    avgHourlyWorkableOpenInbox7d: '',
    avgOvernightInflowProxy7d: '',
    readiness: 'Waiting for workable data',
    shadowStatus: 'Shadow model is scaffolding only in this release.'
  };

  if (!pulseSs) return fallback;

  const workableSheet = pulseSs.getSheetByName((assumptions || {}).workableVolumeLogSheetName);
  const overnightSheet = pulseSs.getSheetByName((assumptions || {}).overnightInflowLogSheetName);
  if (!workableSheet || !overnightSheet) return fallback;

  const workableRows = readSheetRecordsByHeader_(workableSheet);
  const overnightRows = readSheetRecordsByHeader_(overnightSheet);

  const usableWorkableRows = workableRows.filter(row => {
    return String(row.source_status || '') === 'view_open_inbox_only';
  });
  const usableOvernightRows = overnightRows.filter(row => {
    return String(row.source_status || '') === 'view_open_inbox_only';
  });

  const matchingHourlyRows = usableWorkableRows.filter(row => String(row.date_az || '') === dateText);
  const latestHourly = matchingHourlyRows.length ? matchingHourlyRows[matchingHourlyRows.length - 1] : null;

  const matchingOvernight = usableOvernightRows.filter(row => String(row.business_date_az || '') === dateText);
  const overnightRow = matchingOvernight.length ? matchingOvernight[matchingOvernight.length - 1] : null;

  const trailingStart = new Date(dateObj);
  trailingStart.setDate(trailingStart.getDate() - 6);
  const trailingStartText = getDateTextForTimeZone_(trailingStart, timezone, 'yyyy-MM-dd');

  const recentHourlyRows = usableWorkableRows.filter(row => {
    const rowDate = String(row.date_az || '');
    return rowDate && rowDate >= trailingStartText && rowDate <= dateText;
  });
  const recentOvernightRows = usableOvernightRows.filter(row => {
    const rowDate = String(row.business_date_az || '');
    return rowDate && rowDate >= trailingStartText && rowDate <= dateText;
  });

  const hourlyDays = {};
  recentHourlyRows.forEach(row => {
    hourlyDays[String(row.date_az || '')] = true;
  });
  const overnightDays = {};
  recentOvernightRows.forEach(row => {
    overnightDays[String(row.business_date_az || '')] = true;
  });

  const hourlyDayCount = Object.keys(hourlyDays).filter(Boolean).length;
  const overnightDayCount = Object.keys(overnightDays).filter(Boolean).length;
  const minDays = Number((assumptions || {}).minimumObservedSampleDays || 0);

  return {
    status: 'ready',
    dateText: dateText,
    latestHourlyTimestampAz: latestHourly ? String(latestHourly.timestamp_az || '') : '',
    latestHourlyWorkableOpenInbox: latestHourly ? safeNumberOrBlank_(latestHourly.workable_open_inbox) : '',
    latestHourlyWorkableClosed: latestHourly ? safeNumberOrBlank_(latestHourly.workable_closed_hour) : '',
    latestHourlyWorkableTotalVolume: latestHourly ? safeNumberOrBlank_(latestHourly.workable_total_volume) : '',
    latestHourlySourceStatus: latestHourly ? String(latestHourly.source_status || '') : '',
    overnightBusinessDateAz: overnightRow ? String(overnightRow.business_date_az || '') : '',
    overnightWorkableInflowProxy: overnightRow ? safeNumberOrBlank_(overnightRow.overnight_workable_inflow_proxy) : '',
    overnightSourceStatus: overnightRow ? String(overnightRow.source_status || '') : '',
    hourlySampleCount7d: hourlyDayCount,
    overnightSampleCount7d: overnightDayCount,
    avgHourlyWorkableOpenInbox7d: averageNumbers_(recentHourlyRows.map(row => row.workable_open_inbox)),
    avgOvernightInflowProxy7d: averageNumbers_(recentOvernightRows.map(row => row.overnight_workable_inflow_proxy)),
    readiness: hourlyDayCount >= minDays && overnightDayCount >= minDays
      ? 'Enough sample days for review'
      : 'Collecting baseline (' + hourlyDayCount + '/' + minDays + ' hourly days, ' + overnightDayCount + '/' + minDays + ' overnight days)',
    shadowStatus: 'Observed data is visible here, but official staffing recommendations still use the legacy model.'
  };
}

/**
 * Loads all staffing inputs needed to build the model for a date.
 * First-pass version uses stubbed pulse inputs.
 *
 * @param {Date} dateObj
* @returns {{
*   dateObj: Date,
*   dateLabel: string,
*   roster: Array,
*   scheduleMap: Object,
*   assumptions: Object,
*   pulseInputs: Object,
*   observedMetrics: Object
* }}
*/
function getStaffingInputsForDate_(dateObj) {
 const dateLabel = formatDailySheetName_(dateObj);
 const roster = getDisplayRoster_();
 const scheduleMap = getScheduleMapForDate_(dateObj);
 const assumptions = getStaffingAssumptions_();

 let staffingSheet;
 try {
   staffingSheet = getOrCreateStaffingSheet_();
 } catch (e) {
   staffingSheet = null;
 }

 const pulseInputs = getPulseInputsForDate_(staffingSheet, dateObj);
 const observedMetrics = getObservedWorkableMetricsForDate_(dateObj, assumptions);

 return {
   dateObj,
   dateLabel,
   roster,
   scheduleMap,
   assumptions,
   pulseInputs,
   observedMetrics
 };
}

function testBuildStaffingModel() {
  const dateObj = new Date();
  const model = buildStaffingModelForDate_(dateObj);
  Logger.log(JSON.stringify(model, null, 2));
}

/**
 * Writes the staffing recommendation output table.
 * Uses a fixed output block so a future manual input section can live above it.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Array<Object>} rows
 */
function writeStaffingRecommendationTable_(sheet, rows) {
  const outputStartRow = 24;
  const headers = [[
    'Checkpoint Key',
    'Checkpoint Label',
    'Timestamp',
    'totalOpen',
    'unassigned',
    'agedRisk',
    'estimatedInflow',
    'Active Agents',
    'Remaining Productive Hours',
    'Projected Work Remaining',
    'Projected Capacity Remaining',
    'Excess Capacity',
    'Recommended Send Home',
    'Status',
    'Explanation'
  ]];

  const values = (rows || []).map(row => [
    row.checkpointKey,
    row.checkpointLabel,
    row.checkpointTimestamp,
    row.totalOpen,
    row.unassigned,
    row.agedRisk,
    row.estimatedInflow,
    row.activeAgentCount,
    row.remainingProductiveHours,
    row.projectedWorkRemaining,
    row.projectedCapacityRemaining,
    row.excessCapacity,
    row.recommendedSendHomeCount,
    row.recommendationStatus,
    row.explanation
  ]);

  const totalRows = Math.max(sheet.getMaxRows() - outputStartRow + 1, 1);
  const outputRange = sheet.getRange(outputStartRow, 1, totalRows, headers[0].length);

  outputRange.clearContent();
  outputRange.clearFormat();

  sheet.getRange(outputStartRow, 1, 1, headers[0].length)
    .setValues(headers)
    .setFontWeight('bold');

  if (values.length) {
    sheet.getRange(outputStartRow + 1, 1, values.length, headers[0].length)
      .setValues(values);
  }
}

/**
 * Returns the Staffing sheet, creating it if needed.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateStaffingSheet_() {
  const ss = SpreadsheetApp.getActive();
  const staffingCfg = CFG.staffing || {};
  const sheetName = getConfigValue_(
    'STAFFING_SHEET_NAME',
    staffingCfg.sheetName || 'Staffing'
  );

  return getOrCreateSheet_(ss, sheetName);
}

/**
 * Writes the manual staffing input section in rows 1–10.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {{ dateLabel: string, pulseInputs: Object }} model
 */
function writeStaffingInputSection_(sheet, model) {
  const pulseInputs = ((model || {}).pulseInputs || {}).byCheckpointKey || {};
  const existingValues = sheet.getRange(5, 1, 6, 5).getValues();

  sheet.getRange(1, 1).setValue('Date');
  sheet.getRange(1, 2).setValue((model || {}).dateLabel || '');

  sheet.getRange(3, 1, 1, 5).setValues([[
    'Checkpoint Key',
    'totalOpen',
    'unassigned',
    'agedRisk',
    'estimatedInflow'
  ]]).setFontWeight('bold');

  sheet.getRange(5, 1, 6, 1).setNumberFormat('@');

  CFG.checkpoints.forEach((checkpoint, i) => {
    const rowIndex = i + 5;
    const existingRow = existingValues[i] || [];
    const defaults = pulseInputs[checkpoint.key] || {};

    sheet.getRange(rowIndex, 1).setValue(checkpoint.key);
    sheet.getRange(rowIndex, 2, 1, 4).setValues([[
      isNaN(Number(existingRow[1])) ? Number(defaults.totalOpen || 0) : Number(existingRow[1]),
      isNaN(Number(existingRow[2])) ? Number(defaults.unassigned || 0) : Number(existingRow[2]),
      isNaN(Number(existingRow[3])) ? Number(defaults.agedRisk || 0) : Number(existingRow[3]),
      isNaN(Number(existingRow[4])) ? Number(defaults.estimatedInflow || 0) : Number(existingRow[4])
    ]]);
  });
}

function writeStaffingObservedDataSection_(sheet, model) {
  const observed = (model || {}).observedMetrics || {};

  const rows = [
    ['Observed Data', '', ''],
    ['Pulse Log Date', observed.dateText || '', 'Observed metrics are read from the Pulse Log spreadsheet.'],
    ['Latest Hourly Timestamp', observed.latestHourlyTimestampAz || '', 'Latest workable inbox hourly snapshot for the selected date.'],
    ['Latest Hourly Workable Open', observed.latestHourlyWorkableOpenInbox, 'Directly from the agent-facing workable inbox view.'],
    ['Latest Hourly Workable Closed', observed.latestHourlyWorkableClosed, 'Blank until closed workable logic is cleanly defined.'],
    ['Latest Hourly Workable Total', observed.latestHourlyWorkableTotalVolume, 'Blank until closed workable logic is available.'],
    ['Overnight Inflow Proxy', observed.overnightWorkableInflowProxy, 'Current proxy: max(0, SOD workable open - EOD workable open).'],
    ['Avg Hourly Workable Open 7d', observed.avgHourlyWorkableOpenInbox7d, 'Trailing 7-day average across workable hourly rows.'],
    ['Avg Overnight Inflow Proxy 7d', observed.avgOvernightInflowProxy7d, 'Trailing 7-day average across overnight proxy rows.'],
    ['Observed Readiness', observed.readiness || '', 'Baseline collection status only in this release.'],
    ['Shadow Model Status', observed.shadowStatus || '', 'Official staffing recommendations remain on the legacy model.']
  ];

  const startRow = 12;
  const startCol = 1;
  const width = 3;
  const clearRows = 11;

  sheet.getRange(startRow, startCol, clearRows, width).clearContent().clearFormat();
  sheet.getRange(startRow, startCol, rows.length, width).setValues(rows);
  sheet.getRange(startRow, startCol, 1, width).setFontWeight('bold');
  sheet.getRange(startRow + 1, startCol, rows.length - 1, 1).setFontWeight('bold');
  sheet.autoResizeColumns(startCol, width);
}

/**
 * Writes the staffing sheet for the current model.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {{ dateLabel: string, pulseInputs: Object, rows: Array<Object> }} model
 */
function writeStaffingSheet_(sheet, model) {
  writeStaffingInputSection_(sheet, model);
  writeStaffingObservedDataSection_(sheet, model);
  writeStaffingRecommendationTable_(sheet, (model || {}).rows || []);
}

/**
 * Builds and publishes staffing output for a given date.
 *
 * @param {Date} dateObj
 */
function publishStaffingForDate_(dateObj) {
  const sheet = getOrCreateStaffingSheet_();
  const model = buildStaffingModelForDate_(dateObj);
  writeStaffingSheet_(sheet, model);
}

/**
 * Publishes staffing output for today.
 */
function publishTodayStaffing() {
  publishStaffingForDate_(new Date());
}
