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
    )),
    agedRiskWeight: Number(getConfigValue_(
      'STAFFING_AGED_RISK_WEIGHT',
      defaults.agedRiskWeight
    )),
    reserveHoursBuffer: Number(getConfigValue_(
      'STAFFING_RESERVE_HOURS_BUFFER',
      defaults.reserveHoursBuffer
    )),
    minimumAgentsFloor: Number(getConfigValue_(
      'STAFFING_MINIMUM_AGENTS_FLOOR',
      defaults.minimumAgentsFloor
    )),
    cautionUnassignedThreshold: Number(getConfigValue_(
      'STAFFING_CAUTION_UNASSIGNED_THRESHOLD',
      defaults.cautionUnassignedThreshold
    ))
  };
}
