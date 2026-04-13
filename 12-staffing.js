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
 *   cautionUnassignedThreshold: number
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
    )) || defaults.cautionUnassignedThreshold
  };
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

 const parts = [
   'Capacity ' + round1_(capacity) + 'h vs work ' + round1_(work) + 'h',
   'excess ' + round1_(excess) + 'h'
 ];

 if (unassigned > 0) {
   parts.push('unassigned=' + unassigned);
 }

 parts.push(sendHomeCount >= 1 ? 'Send ' + sendHomeCount : 'Hold');

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
  const estimatedInflow = Number(pulseInput.estimatedInflow || 0);

  // Stubbed for first pass; replace with schedule-driven coverage logic later.
  const activeAgentCount = 0;
  const remainingProductiveHours = 0;

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
    rows
  };
}
