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