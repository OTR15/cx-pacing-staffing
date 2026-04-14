function testStaffingAssumptions_() {
    const a = getStaffingAssumptions_();
    Logger.log(JSON.stringify(a, null, 2));
  }

  function testCapacity_() {
    Logger.log(computeProjectedCapacityRemaining_(5));      // 5
    Logger.log(computeProjectedCapacityRemaining_(0));      // 0
    Logger.log(computeProjectedCapacityRemaining_(null));   // 0
    Logger.log(computeProjectedCapacityRemaining_('bad'));  // 0
  }
  function testcapacity() {testCapacity_}

  function debugScheduleForToday() {
  const scheduleMap = getScheduleMapForDate_(new Date());
  const roster = getDisplayRoster_();
  const layout = getLayout_();
  const checkpoint = CFG.checkpoints[2]; // 11AM

  roster.forEach(rep => {
    const schedule = getScheduleForRep_(scheduleMap, rep.repName);
    const inMap = !!scheduleMap[normalizeName_(rep.repName)] || 
                  !!scheduleMap[normalizeFirstName_(rep.repName)];
    const startMins = parseTimeToMinutes_(schedule.startText);
    const cpMins = checkpoint.hour * 60;
    const isNotYetStarted = schedule.hours > 0 && 
                            schedule.startText && 
                            startMins > cpMins;

    Logger.log('parseTimeToMinutes_ test: 12pm = ' + parseTimeToMinutes_('12pm'));

    Logger.log(
      rep.repName + 
      ' | inMap: ' + inMap +
      ' | status: ' + schedule.status +
      ' | hours: ' + schedule.hours +
      ' | startText: "' + schedule.startText + '"' +
      ' | startMins: ' + startMins +
      ' | cpMins: ' + cpMins +
      ' | isNotYetStarted: ' + isNotYetStarted
    );
  });
}

function debugGoalAdjustments() {
  const ss        = SpreadsheetApp.getActive();
  const sheet     = ss.getActiveSheet();
  const sheetName = sheet.getName();

  if (!parseDailySheetName_(sheetName)) {
    SpreadsheetApp.getUi().alert('Please navigate to a daily tab before running this.');
    return;
  }

  const layout  = getLayout_();
  const lastRow = sheet.getLastRow();

  Logger.log('reviewFlagCol: '   + layout.reviewFlagCol);
  Logger.log('reviewReasonCol: ' + layout.reviewReasonCol);
  Logger.log('reviewAdjustCol: ' + layout.reviewAdjustCol);

  const nameValues = sheet.getRange(
    CFG.daily.firstDataRow, 1,
    lastRow - CFG.daily.firstDataRow + 1, 1
  ).getValues();

  nameValues.forEach((r, i) => {
    const name   = String(r[0] || '').trim();
    if (!name) return;
    const row    = CFG.daily.firstDataRow + i;
    const flag   = String(sheet.getRange(row, layout.reviewFlagCol).getValue()   || '').trim();
    const reason = String(sheet.getRange(row, layout.reviewReasonCol).getValue() || '').trim();
    const adjust = String(sheet.getRange(row, layout.reviewAdjustCol).getValue() || '').trim();

    Logger.log(name + ' | row: ' + row + ' | flag: "' + flag + '" | reason: "' + reason + '" | adjust: "' + adjust + '"');
  });
}
