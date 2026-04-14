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
