// =============================================================================
// gorgias.gs
// All Gorgias Stats API calls.
//
// Authentication: Basic auth using script properties
//   GORGIAS_API_USERNAME  (your Gorgias login email)
//   GORGIAS_API_KEY       (your Gorgias API key)
//
// All public functions accept ISO 8601 date strings for time ranges.
// Retry logic is built into postGorgiasStatWithRetry_() — 3 attempts
// with increasing delays to handle transient API errors.
// =============================================================================

// ── Public stat helpers ───────────────────────────────────────────────────────

/**
 * Fetches a single numeric stat from the Gorgias Stats API.
 * Returns 0 if the API returns no value.
 *
 * @param {string} statName     - Gorgias stat endpoint name (e.g. 'total-tickets-replied')
 * @param {string} startIso     - Period start as ISO 8601 string
 * @param {string} endIso       - Period end as ISO 8601 string
 * @param {Object} extraFilters - Additional filter fields merged into the request body
 * @returns {number}
 */
function getStatNumber_(statName, startIso, endIso, extraFilters) {
  const json = postGorgiasStatWithRetry_(statName, startIso, endIso, extraFilters || {});
  return Number((((json || {}).data || {}).data || {}).value || 0);
}

/**
 * Fetches CSAT metrics for a given period and optional agent filter.
 * Returns averageRating as '' when no surveys were sent (avoids showing 0).
 *
 * @param {string} startIso
 * @param {string} endIso
 * @param {Object} extraFilters - e.g. { agents: [agentId] }
 * @returns {{ averageRating: number|'', totalSent: number }}
 */
function getCsatMetrics_(startIso, endIso, extraFilters) {
  const json  = postGorgiasStatWithRetry_('satisfaction-surveys', startIso, endIso, extraFilters || {});
  const items = (((json || {}).data || {}).data) || [];

  let averageRating = '';
  let totalSent     = 0;

  items.forEach(item => {
    if (item.name === 'average_rating') averageRating = Number(item.value || 0);
    if (item.name === 'total_sent')     totalSent     = Number(item.value || 0);
  });

  return {
    averageRating: totalSent > 0 ? averageRating : '',
    totalSent
  };
}

/**
 * Fetches tickets-closed-per-agent and returns a lookup map of
 * { normalizedName → count } for fast per-rep lookup during publish.
 * Each name is stored under both full and first-name keys.
 *
 * @param {string} startIso
 * @param {string} endIso
 * @returns {Object}
 */
function getClosedPerAgentMap_(startIso, endIso) {
  const json  = postGorgiasStatWithRetry_('tickets-closed-per-agent', startIso, endIso, {});
  const lines = ((((json || {}).data || {}).data || {}).lines) || [];
  const map   = {};

  lines.forEach(row => {
    const agentName = row && row[0] && row[0].value ? String(row[0].value) : '';
    const total     = row && row[1] && row[1].value ? Number(row[1].value) : 0;
    if (!agentName) return;

    map[normalizeName_(agentName)]     = total;
    map[normalizeFirstName_(agentName)] = total;
  });

  return map;
}

// ── Date range helper ─────────────────────────────────────────────────────────

/**
 * Returns the ISO date range for a checkpoint: start of day → checkpoint time.
 *
 * NOTE: The UTC offset is hardcoded to -07:00 (MST/PDT).
 * This is a known limitation — it will be off by one hour during DST transitions.
 * TODO: Replace with a dynamic offset derived from CFG.timezone.
 *
 * @param {Date}   dateObj    - The date to generate the range for.
 * @param {{ hour: number, minute: number }} checkpoint
 * @returns {{ startIso: string, endIso: string }}
 */
function getCheckpointIsoRange_(dateObj, checkpoint) {
  const start   = Utilities.formatDate(dateObj, CFG.timezone, 'yyyy-MM-dd') + 'T00:00:00-07:00';
  const endDate = new Date(dateObj);
  endDate.setHours(checkpoint.hour);
  endDate.setMinutes(checkpoint.minute || 0);
  endDate.setSeconds(0);
  const end = Utilities.formatDate(endDate, CFG.timezone, "yyyy-MM-dd'T'HH:mm:ss-07:00");

  return { startIso: start, endIso: end };
}

// ── Retry wrapper ─────────────────────────────────────────────────────────────

/**
 * Calls postGorgiasStat_ with up to 3 attempts and increasing delays.
 * Delays: 0ms, 1200ms, 2500ms between attempts.
 * Throws the last error if all attempts fail.
 *
 * @param {string} statName
 * @param {string} startIso
 * @param {string} endIso
 * @param {Object} extraFilters
 * @returns {Object} Parsed JSON response
 */
function postGorgiasStatWithRetry_(statName, startIso, endIso, extraFilters) {
  const waits   = [0, 1200, 2500];
  let lastErr;

  for (let i = 0; i < waits.length; i++) {
    if (waits[i] > 0) Utilities.sleep(waits[i]);
    try {
      return postGorgiasStat_(statName, startIso, endIso, extraFilters || {});
    } catch (err) {
      lastErr = err;
    }
  }

  throw lastErr;
}

// ── Core API call ─────────────────────────────────────────────────────────────

/**
 * POSTs to the Gorgias Stats API and returns the parsed JSON response.
 * Throws a descriptive error on non-2xx responses.
 *
 * @param {string} statName
 * @param {string} startIso
 * @param {string} endIso
 * @param {Object} extraFilters
 * @returns {Object}
 */
function postGorgiasStat_(statName, startIso, endIso, extraFilters) {
  const username  = getRequiredProperty_('GORGIAS_API_USERNAME');
  const apiKey    = getRequiredProperty_('GORGIAS_API_KEY');
  const subdomain = getConfigValue_('SUBDOMAIN', CFG.subdomain);

  const url     = 'https://' + subdomain + '.gorgias.com/api/stats/' + statName;
  const payload = {
    filters: Object.assign(
      { period: { start_datetime: startIso, end_datetime: endIso } },
      extraFilters || {}
    )
  };

  const response = UrlFetchApp.fetch(url, {
    method:          'post',
    contentType:     'application/json',
    headers: {
      Authorization: 'Basic ' + Utilities.base64Encode(username + ':' + apiKey)
    },
    payload:         JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const body = response.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error('Gorgias stat error for ' + statName + ': ' + code + ' ' + body);
  }

  return JSON.parse(body);
}
