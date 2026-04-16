// =============================================================================
// weekly-kpi.gs
// Builds a static weekly KPI admin snapshot using the same source data as the
// KPI dashboard, without needing to call the other Apps Script project.
// =============================================================================

function weeklyKpiParseLocalDate_(value) {
  if (value instanceof Date) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  const str = String(value || '').trim();
  const m = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }

  const parsed = new Date(value);
  return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

function weeklyKpiOpenSpreadsheet_(id) {
  return SpreadsheetApp.openById(String(id || '').trim());
}

function getWeeklyKpiConfig_() {
  const dashboard = weeklyKpiOpenSpreadsheet_(CFG.weekly.kpiSnapshot.spreadsheetId);
  const sh = dashboard.getSheetByName('CONFIG');
  if (!sh) throw new Error('KPI CONFIG sheet not found.');

  const data = sh.getDataRange().getValues();
  const cfg = {};
  data.forEach(row => {
    const key = String(row[0] || '').trim();
    if (key && !key.startsWith('──') && key !== 'Setting Key') {
      cfg[key] = String(row[1] || '').trim();
    }
  });
  return cfg;
}

function weeklyKpiFetchCSAT_(startDate, endDate, cfg) {
  const ss = weeklyKpiOpenSpreadsheet_(cfg.CSAT_SPREADSHEET_ID);
  const sh = ss.getSheetByName(cfg.CSAT_SHEET_NAME);
  if (!sh) return {};

  const raw = sh.getDataRange().getValues();
  const hdr = raw[0].map(h => String(h).trim());

  const iDate = hdr.indexOf(cfg.CSAT_COL_DATE);
  const iAgent = hdr.indexOf(cfg.CSAT_COL_AGENT_NAME);
  const iScore = hdr.indexOf(cfg.CSAT_COL_SCORE);
  if ([iDate, iAgent, iScore].some(i => i < 0)) return {};

  const sd = weeklyKpiParseLocalDate_(startDate); sd.setHours(0, 0, 0, 0);
  const ed = weeklyKpiParseLocalDate_(endDate); ed.setHours(23, 59, 59, 999);
  const scores = {};

  for (let i = 1; i < raw.length; i++) {
    const row = raw[i];
    const dateRaw = row[iDate];
    if (!dateRaw) continue;

    const d = new Date(dateRaw);
    if (d < sd || d > ed) continue;

    const agent = String(row[iAgent] || '').trim();
    const score = parseFloat(row[iScore]);
    if (!agent || isNaN(score)) continue;

    if (!scores[agent]) scores[agent] = [];
    scores[agent].push(score);
  }

  const result = {};
  Object.keys(scores).forEach(agent => {
    const values = scores[agent];
    result[agent] = values.reduce((sum, value) => sum + value, 0) / values.length;
  });

  return result;
}

function weeklyKpiFetchQA_(startDate, endDate, cfg) {
  const ss = weeklyKpiOpenSpreadsheet_(cfg.QA_SPREADSHEET_ID);
  const sh = ss.getSheetByName(cfg.QA_SHEET_NAME);
  if (!sh) return {};

  const raw = sh.getDataRange().getValues();
  const hdr = raw[0].map(h => String(h).trim());

  const iDate = hdr.indexOf(cfg.QA_COL_DATE);
  const iAgent = hdr.indexOf(cfg.QA_COL_AGENT_NAME);
  const iScore = hdr.indexOf(cfg.QA_COL_SCORE);
  if ([iDate, iAgent, iScore].some(i => i < 0)) return {};

  const isPercent = String(cfg.QA_SCORE_IS_PERCENT || '').toUpperCase() === 'TRUE';
  const sd = weeklyKpiParseLocalDate_(startDate); sd.setHours(0, 0, 0, 0);
  const ed = weeklyKpiParseLocalDate_(endDate); ed.setHours(23, 59, 59, 999);
  const scores = {};

  for (let i = 1; i < raw.length; i++) {
    const row = raw[i];
    const dateRaw = row[iDate];
    if (!dateRaw) continue;

    const d = new Date(dateRaw);
    if (d < sd || d > ed) continue;

    const agent = String(row[iAgent] || '').trim();
    let score = parseFloat(row[iScore]);
    if (!agent || isNaN(score)) continue;

    if (!isPercent) score *= 100;
    if (!scores[agent]) scores[agent] = [];
    scores[agent].push(score);
  }

  const result = {};
  Object.keys(scores).forEach(agent => {
    const values = scores[agent];
    result[agent] = values.reduce((sum, value) => sum + value, 0) / values.length;
  });

  return result;
}

function weeklyKpiParseNotesString_(raw) {
  const str = String(raw || '').trim();
  if (!str) return null;

  const statusMatch = str.match(/Status:\s*([^|]+)/i);
  const status = statusMatch ? statusMatch[1].trim() : '';
  const isActive = status.toLowerCase() === 'active';

  if (!isActive) return { skip: true, goalClosed: null, goalReplied: null };

  const closedMatch = str.match(/C:([\d.]+)/);
  const repliedMatch = str.match(/R:([\d.]+)/);

  return {
    skip: false,
    goalClosed: closedMatch ? parseFloat(closedMatch[1]) : null,
    goalReplied: repliedMatch ? parseFloat(repliedMatch[1]) : null
  };
}

function weeklyKpiFetchPacing_(startDate, endDate, cfg) {
  const ss = weeklyKpiOpenSpreadsheet_(cfg.PACING_SPREADSHEET_ID);
  const eodClosedCol = (parseInt(cfg.PACING_EOD_CLOSED_COL, 10) || 23) - 1;
  const eodRepliedCol = (parseInt(cfg.PACING_EOD_REPLIED_COL, 10) || 24) - 1;
  const dataStartRow = (parseInt(cfg.PACING_DATA_START_ROW, 10) || 3) - 1;
  const agentData = {};

  const sd = weeklyKpiParseLocalDate_(startDate); sd.setHours(0, 0, 0, 0);
  const ed = weeklyKpiParseLocalDate_(endDate); ed.setHours(0, 0, 0, 0);

  for (let d = new Date(sd); d <= ed; d.setDate(d.getDate() + 1)) {
    const tabName = Utilities.formatDate(d, CFG.timezone, 'M/d/yy');
    const sh = ss.getSheetByName(tabName);
    if (!sh) continue;

    const raw = sh.getDataRange().getValues();
    const headerRow = raw[1] || [];
    const notesCol = headerRow.findIndex(h => String(h || '').trim().toLowerCase() === 'notes');
    const notesIdx = notesCol >= 0 ? notesCol : raw[0].length - 1;

    for (let i = dataStartRow; i < raw.length; i++) {
      const row = raw[i];
      const agent = String(row[0] || '').trim();
      if (!agent) continue;

      const notes = weeklyKpiParseNotesString_(row[notesIdx]);
      if (!notes || notes.skip) continue;

      const closed = parseFloat(row[eodClosedCol]);
      const replied = parseFloat(row[eodRepliedCol]);
      if (closed === 0 && replied === 0) continue;

      if (!agentData[agent]) {
        agentData[agent] = { replied: [], closed: [], goalReplied: [], goalClosed: [] };
      }

      if (!isNaN(replied)) agentData[agent].replied.push(replied);
      if (!isNaN(closed)) agentData[agent].closed.push(closed);
      if (notes.goalReplied != null) agentData[agent].goalReplied.push(notes.goalReplied);
      if (notes.goalClosed != null) agentData[agent].goalClosed.push(notes.goalClosed);
    }
  }

  const avg = values => values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : null;
  const result = {};

  Object.keys(agentData).forEach(agent => {
    const data = agentData[agent];
    result[agent] = {
      replied: avg(data.replied),
      closed: avg(data.closed),
      goalReplied: avg(data.goalReplied),
      goalClosed: avg(data.goalClosed),
      activeDays: data.replied.length
    };
  });

  return result;
}

function weeklyKpiFetchPacingBackfill_(startDate, endDate, cfg) {
  const ssId = cfg.PACING_BACKFILL_SPREADSHEET_ID || cfg.PACING_SPREADSHEET_ID;
  const ss = weeklyKpiOpenSpreadsheet_(ssId);
  const sh = ss.getSheetByName(cfg.PACING_BACKFILL_SHEET_NAME);
  if (!sh) return {};

  const raw = sh.getDataRange().getValues();
  const hdr = raw[0].map(h => String(h).trim());

  const iDate = hdr.indexOf('Date');
  const iCheckpoint = hdr.indexOf('Checkpoint');
  const iAgent = hdr.indexOf('Agent Name');
  const iEffHours = hdr.indexOf('Effective Hours');
  const iClosed = hdr.indexOf('Closed Tickets');
  const iReplied = hdr.indexOf('Tickets Replied');
  if ([iDate, iCheckpoint, iAgent, iEffHours, iClosed, iReplied].some(i => i < 0)) return {};

  const fullDayReplied = parseFloat(cfg.GOAL_TICKETS_REPLIED) || 70;
  const fullDayClosed = parseFloat(cfg.GOAL_CLOSED) || 53;
  const sd = weeklyKpiParseLocalDate_(startDate); sd.setHours(0, 0, 0, 0);
  const ed = weeklyKpiParseLocalDate_(endDate); ed.setHours(23, 59, 59, 999);
  const agentDayEod = {};

  for (let i = 1; i < raw.length; i++) {
    const row = raw[i];
    const dateRaw = row[iDate];
    const checkpoint = String(row[iCheckpoint] || '').trim().toUpperCase();
    if (!dateRaw || checkpoint !== 'EOD') continue;

    const d = new Date(dateRaw);
    if (d < sd || d > ed) continue;

    const agent = String(row[iAgent] || '').trim();
    const effHours = parseFloat(row[iEffHours]);
    const closed = parseFloat(row[iClosed]);
    const replied = parseFloat(row[iReplied]);
    if (!agent || isNaN(effHours) || effHours <= 0) continue;

    const hoursRatio = Math.min(effHours / 8, 1);
    const dayStr = Utilities.formatDate(d, CFG.timezone, 'yyyy-MM-dd');

    if (!agentDayEod[agent]) agentDayEod[agent] = {};
    agentDayEod[agent][dayStr] = {
      replied: isNaN(replied) ? null : replied,
      closed: isNaN(closed) ? null : closed,
      goalReplied: Math.round(fullDayReplied * hoursRatio),
      goalClosed: Math.round(fullDayClosed * hoursRatio)
    };
  }

  const avg = values => values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : null;
  const result = {};

  Object.keys(agentDayEod).forEach(agent => {
    const days = Object.values(agentDayEod[agent]);
    result[agent] = {
      replied: avg(days.map(day => day.replied).filter(value => value != null)),
      closed: avg(days.map(day => day.closed).filter(value => value != null)),
      goalReplied: avg(days.map(day => day.goalReplied).filter(value => value != null)),
      goalClosed: avg(days.map(day => day.goalClosed).filter(value => value != null)),
      activeDays: days.length
    };
  });

  return result;
}

function weeklyKpiMergePacingData_(startDate, endDate, cfg) {
  const dailyPacingStart = new Date(2026, 2, 22);
  const sd = weeklyKpiParseLocalDate_(startDate);
  const ed = weeklyKpiParseLocalDate_(endDate);
  const useBackfill = sd < dailyPacingStart;
  const useDaily = ed >= dailyPacingStart;

  let backfillData = {};
  let dailyData = {};

  if (useBackfill) {
    const backfillEnd = useDaily ? new Date(dailyPacingStart.getTime() - 86400000) : ed;
    backfillData = weeklyKpiFetchPacingBackfill_(sd, backfillEnd, cfg);
  }

  if (useDaily) {
    const dailyStart = useBackfill ? dailyPacingStart : sd;
    dailyData = weeklyKpiFetchPacing_(dailyStart, ed, cfg);
  }

  const allAgents = new Set([
    ...Object.keys(backfillData),
    ...Object.keys(dailyData)
  ]);

  const merged = {};
  allAgents.forEach(agent => {
    const backfill = backfillData[agent];
    const daily = dailyData[agent];

    if (backfill && daily) {
      const backfillDays = backfill.activeDays || 1;
      const dailyDays = daily.activeDays || 1;
      const totalDays = backfillDays + dailyDays;
      const weightedAverage = (a, b) => {
        if (a == null && b == null) return null;
        if (a == null) return b;
        if (b == null) return a;
        return (a * backfillDays + b * dailyDays) / totalDays;
      };

      merged[agent] = {
        replied: weightedAverage(backfill.replied, daily.replied),
        closed: weightedAverage(backfill.closed, daily.closed),
        goalReplied: daily.goalReplied != null ? daily.goalReplied : backfill.goalReplied,
        goalClosed: daily.goalClosed != null ? daily.goalClosed : backfill.goalClosed,
        activeDays: totalDays
      };
    } else {
      merged[agent] = backfill || daily;
    }
  });

  return merged;
}

function weeklyKpiGetExcludedAgents_(cfg) {
  return new Set(String(cfg.EXCLUDED_AGENTS || '').split(',').map(name => name.trim()).filter(Boolean));
}

function weeklyKpiGetExemptAgents_(cfg) {
  return new Set(String(cfg.EXEMPT_AGENTS || '').split(',').map(name => name.trim()).filter(Boolean));
}

function weeklyKpiBuildQaNameMap_(qaMap, pacMap, csatMap, cfg) {
  const format = String(cfg.QA_NAME_FORMAT || 'first').toLowerCase().trim();
  const overrides = String(cfg.QA_NAME_OVERRIDES || '');
  const manualMap = {};

  if (overrides.trim()) {
    overrides.split(',').forEach(pair => {
      const parts = pair.split('=');
      if (parts.length === 2) manualMap[parts[0].trim()] = parts[1].trim();
    });
  }

  if (format === 'full') return manualMap;

  const fullNames = new Set([
    ...Object.keys(pacMap),
    ...Object.keys(csatMap)
  ]);
  const autoMap = {};

  Object.keys(qaMap).forEach(qaName => {
    if (manualMap[qaName]) return;
    const qaFirst = qaName.trim().toLowerCase();
    let matched = null;
    let matchCount = 0;

    fullNames.forEach(fullName => {
      const firstName = fullName.trim().split(' ')[0].toLowerCase();
      if (firstName === qaFirst) {
        matched = fullName;
        matchCount++;
      }
    });

    if (matchCount === 1) autoMap[qaName] = matched;
  });

  return Object.assign(autoMap, manualMap);
}

function weeklyKpiResolveQaScore_(fullName, qaMap, nameMap) {
  if (qaMap[fullName] != null) return qaMap[fullName];
  const qaKey = Object.keys(nameMap).find(key => nameMap[key] === fullName);
  return qaKey && qaMap[qaKey] != null ? qaMap[qaKey] : null;
}

function weeklyKpiGetAllAgents_(csatMap, pacMap, exempt) {
  const all = new Set([
    ...Object.keys(csatMap),
    ...Object.keys(pacMap),
    ...exempt
  ]);

  return [...all].sort();
}

function weeklyKpiCalcAgentScore_(agentName, csatMap, qaMap, pacMap, nameMap, cfg, exempt) {
  const isExempt = exempt.has(agentName);
  const goalQa = parseFloat(cfg.GOAL_QA) || 90;
  const goalCsat = parseFloat(cfg.GOAL_CSAT) || 4.9;
  const globalGoalReplied = parseFloat(cfg.GOAL_TICKETS_REPLIED) || 70;
  const globalGoalClosed = parseFloat(cfg.GOAL_CLOSED) || 53;
  const pacAgent = pacMap[agentName];
  const goalReplied = pacAgent && pacAgent.goalReplied != null ? pacAgent.goalReplied : globalGoalReplied;
  const goalClosed = pacAgent && pacAgent.goalClosed != null ? pacAgent.goalClosed : globalGoalClosed;

  const wQa = parseFloat(cfg.WEIGHT_QA) || 40;
  const wTix = parseFloat(cfg.WEIGHT_TICKETS) || 20;
  const wClosed = parseFloat(cfg.WEIGHT_CLOSED) || 20;
  const wCsat = parseFloat(cfg.WEIGHT_CSAT) || 20;

  const afQa = parseFloat(cfg.AUTOFAIL_QA_THRESHOLD) || 74;
  const globalAfTix = parseFloat(cfg.AUTOFAIL_TICKETS_THRESHOLD) || 40;
  const afTix = globalGoalReplied > 0
    ? Math.round(globalAfTix * (goalReplied / globalGoalReplied))
    : globalAfTix;

  const qa = weeklyKpiResolveQaScore_(agentName, qaMap, nameMap);
  const replied = pacAgent ? pacAgent.replied : null;
  const closed = pacAgent ? pacAgent.closed : null;
  const csat = csatMap[agentName] != null ? csatMap[agentName] : null;

  if (isExempt) {
    return {
      qa, replied, closed, csat,
      overallPct: null,
      autoFail: false,
      isExempt: true,
      goals: { qa: goalQa, replied: goalReplied, closed: goalClosed, csat: goalCsat },
      afThresholds: { qa: afQa, tix: afTix },
      qaRatio: null, repliedRatio: null, closedRatio: null, csatRatio: null
    };
  }

  const cap = (value, goal) => value == null ? null : Math.min(value / goal, 1.10);
  const qaRatio = cap(qa, goalQa);
  const repliedRatio = cap(replied, goalReplied);
  const closedRatio = cap(closed, goalClosed);
  const csatRatio = cap(csat, goalCsat);

  let totalWeight = 0;
  let weightedSum = 0;
  [[qaRatio, wQa], [repliedRatio, wTix], [closedRatio, wClosed], [csatRatio, wCsat]].forEach(([ratio, weight]) => {
    if (ratio != null) {
      weightedSum += ratio * weight;
      totalWeight += weight;
    }
  });

  let overallPct = totalWeight > 0 ? (weightedSum / totalWeight) * 100 : null;
  const autoFail =
    (qa != null && qa <= afQa) ||
    (replied != null && replied < afTix);
  if (autoFail) overallPct = 0;

  return {
    qa, replied, closed, csat,
    overallPct,
    autoFail,
    isExempt: false,
    goals: { qa: goalQa, replied: goalReplied, closed: goalClosed, csat: goalCsat },
    afThresholds: { qa: afQa, tix: afTix },
    qaRatio, repliedRatio, closedRatio, csatRatio
  };
}

function collectWeeklyKpiSnapshot_(monday, sunday) {
  const cfg = getWeeklyKpiConfig_();
  const csatMap = weeklyKpiFetchCSAT_(monday, sunday, cfg);
  const qaMap = weeklyKpiFetchQA_(monday, sunday, cfg);
  const pacMap = weeklyKpiMergePacingData_(monday, sunday, cfg);
  const excluded = weeklyKpiGetExcludedAgents_(cfg);
  const exempt = weeklyKpiGetExemptAgents_(cfg);
  const qaNameMap = weeklyKpiBuildQaNameMap_(qaMap, pacMap, csatMap, cfg);

  const allAgents = weeklyKpiGetAllAgents_(csatMap, pacMap, exempt)
    .filter(name => !excluded.has(name));

  const rows = [];
  const backgrounds = [];
  const chartRows = [];

  const colors = {
    exceeding: '#e6f4ea',
    meeting: '#e8f0fe',
    close: '#fef7e0',
    failing: '#fce8e6',
    exempt: '#f1f3f4',
    nodata: '#ffffff'
  };

  allAgents.forEach(agent => {
    const score = weeklyKpiCalcAgentScore_(agent, csatMap, qaMap, pacMap, qaNameMap, cfg, exempt);

    let status;
    let note = '';
    let background = colors.nodata;

    if (score.isExempt) {
      status = 'Exempt';
      note = 'Shown but not scored';
      background = colors.exempt;
    } else if (score.autoFail) {
      status = 'AUTO-FAIL';
      const failures = [];
      if (score.qa != null && score.qa <= score.afThresholds.qa) failures.push('QA');
      if (score.replied != null && score.replied < score.afThresholds.tix) failures.push('Tickets Replied');
      note = failures.join(' + ');
      background = colors.failing;
    } else if (score.overallPct == null) {
      status = 'No data';
      background = colors.nodata;
    } else if (score.overallPct >= 106) {
      status = 'Exceeding';
      background = colors.exceeding;
    } else if (score.overallPct >= 100) {
      status = 'Meeting';
      background = colors.meeting;
    } else if (score.overallPct >= 90) {
      status = 'Close';
      background = colors.close;
    } else {
      status = 'Not Meeting';
      const opportunities = [
        { label: 'QA', ratio: score.qaRatio, gap: score.qa != null ? score.goals.qa - score.qa : null },
        { label: 'Tickets Replied', ratio: score.repliedRatio, gap: score.replied != null ? score.goals.replied - score.replied : null },
        { label: 'Closed Tickets', ratio: score.closedRatio, gap: score.closed != null ? score.goals.closed - score.closed : null },
        { label: 'CSAT', ratio: score.csatRatio, gap: score.csat != null ? score.goals.csat - score.csat : null }
      ].filter(item => item.ratio != null);

      if (opportunities.length) {
        opportunities.sort((a, b) => a.ratio - b.ratio);
        const top = opportunities[0];
        if (top.gap == null || top.gap <= 0) {
          note = 'Focus: ' + top.label;
        } else {
          const gap = top.label === 'CSAT' ? top.gap.toFixed(2) : Math.round(top.gap).toString();
          note = 'Focus: ' + top.label + ' (-' + gap + ')';
        }
      }

      background = colors.failing;
    }

    rows.push([
      agent,
      score.qa,
      score.goals.qa,
      score.replied,
      score.goals.replied,
      score.closed,
      score.goals.closed,
      score.csat,
      score.overallPct,
      status,
      note
    ]);
    backgrounds.push(background);

    if (score.overallPct != null) {
      chartRows.push([agent, score.overallPct]);
    }
  });

  rows.sort((a, b) => {
    const aScore = a[8];
    const bScore = b[8];
    if (aScore == null && bScore == null) return a[0].localeCompare(b[0]);
    if (aScore == null) return 1;
    if (bScore == null) return -1;
    return bScore - aScore;
  });

  chartRows.sort((a, b) => b[1] - a[1]);

  return {
    rows,
    backgrounds: rows.map(row => {
      const matchIndex = allAgents.indexOf(row[0]);
      return matchIndex >= 0 ? backgrounds[matchIndex] : colors.nodata;
    }),
    chartRows
  };
}

function writeWeeklyKpiSnapshotSection_(sh, snapshot, monday, sunday) {
  const startRow = CFG.weekly.kpiSnapshot.tableStartRow;
  const title = 'KPI Admin Snapshot: ' +
    Utilities.formatDate(monday, CFG.timezone, 'M/d/yy') +
    ' - ' +
    Utilities.formatDate(sunday, CFG.timezone, 'M/d/yy');

  sh.getRange(startRow, 1, 1, 11).merge()
    .setValue(title)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#d9ead3');

  const headers = [[
    'Agent', 'QA Score', 'QA Goal', 'Tickets Replied', 'Replied Goal',
    'Closed Tickets', 'Closed Goal', 'CSAT', 'Overall %', 'Status', 'Note'
  ]];

  sh.getRange(startRow + 1, 1, 1, headers[0].length)
    .setValues(headers)
    .setFontWeight('bold')
    .setBackground('#cfe2f3');

  if (!snapshot.rows.length) return;

  sh.getRange(startRow + 2, 1, snapshot.rows.length, headers[0].length).setValues(snapshot.rows);
  snapshot.backgrounds.forEach((background, index) => {
    sh.getRange(startRow + 2 + index, 1, 1, headers[0].length).setBackground(background);
  });

  sh.getRange(startRow + 2, 2, snapshot.rows.length, 2).setNumberFormat('0.0"%"');
  sh.getRange(startRow + 2, 4, snapshot.rows.length, 4).setNumberFormat('0.0');
  sh.getRange(startRow + 2, 8, snapshot.rows.length, 2).setNumberFormat('0.0');
  sh.getRange(startRow + 2, 9, snapshot.rows.length, 1).setNumberFormat('0.0"%"');
}
