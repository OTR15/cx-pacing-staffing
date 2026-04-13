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
