// =========================================================================
// 💰 CRESCENTIA — SHIFT REGISTERS & RECONCILIATION (Phase 3, refined)
// Shift_Registers cols: 1 Shift_ID,2 Date,3 Shift_User,4 Counter_Name,
//   5 Opening_Cash,6 Cash_Collected,7 Petty_Expenses,8 Refunds_Issued,
//   9 Expected_Closing_Cash,10 Actual_Physical_Cash,11 Variance_Amount,
//   12 Status,13 Timestamp(OPEN time),14 Closed_At(CLOSE time)
// Mismatch is tallied by an optional adjustment (CASH_SHORT / CASH_OVER)
// posted to Finance_Master_Ledger + a mandatory reason on the audit trail.
// =========================================================================

function acc_fmtDay_(v) { var d = (v instanceof Date) ? v : new Date(v); return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, ACC_CFG.TZ, "yyyy-MM-dd"); }
function acc_fmtTime_(v) { var d = (v instanceof Date) ? v : new Date(v); return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, ACC_CFG.TZ, "HH:mm"); }

// cash movement (CASH income − cash expenses) since a moment — the shift's auto-calc.
function acc_cashSince_(fromDate) {
  var from = (fromDate instanceof Date && !isNaN(fromDate.getTime())) ? fromDate.getTime() : 0;
  var cashIn = 0, cashOut = 0;
  acc_pharmaRows_().concat(acc_labRows_()).forEach(function (r) {
    if (!r.realized || r.net <= 0) return;
    if (acc_str_(r.mode).toLowerCase() !== 'cash') return;
    var d = r.realizedDate || r.billDate;
    if (d && d.getTime() >= from) cashIn += r.net;
  });
  acc_readObjects_(ACC_CFG.LEDGER).forEach(function (r) {
    var out = acc_money_(r['Amount_Out']);
    if (out <= 0) return;
    if (acc_str_(r['Payment_Mode']).toLowerCase() !== 'cash') return;
    var d = acc_toDate_(r['Timestamp']);
    if (d && d.getTime() >= from) cashOut += out;
  });
  return { cashIn: acc_money_(cashIn), cashOut: acc_money_(cashOut) };
}

function getShiftState(counterName) {
  try {
    var rows = acc_readObjects_(ACC_CFG.SHIFTS);
    var open = null, history = [];
    
    rows.forEach(function (r) {
      var rec = {
        shiftId: acc_str_(r['Shift_ID']), 
        date: acc_fmtDay_(r['Date']) || acc_str_(r['Date']),
        user: acc_str_(r['Shift_User']), 
        incharge: acc_str_(r['Shift_User']), // Mapped to new Incharge column
        counter: acc_str_(r['Counter_Name']),
        opening: acc_money_(r['Opening_Cash']), 
        collected: acc_money_(r['Cash_Collected']),
        petty: acc_money_(r['Petty_Expenses']), 
        refunds: acc_money_(r['Refunds_Issued']),
        expected: acc_money_(r['Expected_Closing_Cash']), 
        actual: acc_money_(r['Actual_Physical_Cash']),
        variance: acc_money_(r['Variance_Amount']), 
        status: acc_str_(r['Status']),
        openTime: acc_fmtTime_(r['Timestamp']),
        closeTime: r['Closed_At'] ? acc_fmtTime_(r['Closed_At']) : '',
        openedAt: acc_toDate_(r['Timestamp'])
      };
      if (rec.status.toUpperCase() === 'OPEN' && (!counterName || rec.counter === counterName)) open = rec;
      history.push(rec);
    });
    history.reverse();

    var suggestion = null;
    if (open) {
      var c = acc_cashSince_(open.openedAt);
      suggestion = { cashIn: c.cashIn, cashOut: c.cashOut };
    }
    
    // Strip native Dates before transit
    if (open) delete open.openedAt;
    history.forEach(function(h) { delete h.openedAt; });

    return { success: true, open: open, suggestion: suggestion, history: history.slice(0, 30) };
  } catch (e) { return { success: false, message: e.message }; }
}

function openShift(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var counter = acc_str_(payload.counterName).trim();
    if (!counter) return { success: false, message: "Counter name required." };
    var opening = acc_money_(payload.openingCash);
    if (opening < 0) return { success: false, message: "Opening cash cannot be negative." };

    var sh = acc_sheet_(ACC_CFG.SHIFTS);
    var data = sh.getDataRange().getValues(); 
    for (var i = 1; i < data.length; i++) {
      if (acc_str_(data[i][3]).trim() === counter && acc_str_(data[i][11]).toUpperCase() === 'OPEN')
        return { success: false, message: "Counter '" + counter + "' already has an open shift. Close it first." };
    }

    var now = new Date();
    var shiftId = "SH-" + Utilities.formatDate(now, ACC_CFG.TZ, "yyyyMMdd") + "-" + Date.now().toString().slice(-5);
    sh.appendRow([
      shiftId, Utilities.formatDate(now, ACC_CFG.TZ, "yyyy-MM-dd"),
      acc_str_(payload.user) || 'UNKNOWN', counter, opening,
      "", "", "", "", "", "", "OPEN", now, ""
    ]);
    acc_audit_(payload.user, 'OPEN_SHIFT', 'Shift_Register', shiftId, '', 'Opening: ' + opening, counter);
    SpreadsheetApp.flush();
    return { success: true, message: "Shift opened for " + counter + ".", shiftId: shiftId };
  } catch (e) { return { success: false, message: "Error: " + e.message }; }
  finally { lock.releaseLock(); }
}

function closeShift(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var target = acc_str_(payload.shiftId).trim();
    if (!target) return { success: false, message: "Missing shift reference." };

    var collected = acc_money_(payload.cashCollected);
    var petty = acc_money_(payload.pettyExpenses);
    var refunds = acc_money_(payload.refunds);
    var actual = acc_money_(payload.actualCash);
    var reason = acc_str_(payload.reason).trim();

    var sh = acc_sheet_(ACC_CFG.SHIFTS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (acc_str_(data[i][0]).trim() !== target) continue;
      if (acc_str_(data[i][11]).toUpperCase() !== 'OPEN')
        return { success: false, message: "Shift " + target + " is not open." };

      var counter = acc_str_(data[i][3]);
      var opening = acc_money_(data[i][4]);
      var expected = acc_money_(opening + collected - petty - refunds);
      var variance = acc_money_(actual - expected);
      var mismatch = Math.abs(variance) > 0;
      var status = !mismatch ? 'RECONCILED' : (Math.abs(variance) > ACC_CFG.VARIANCE_FLAG ? 'MISMATCH' : 'CLOSED');
      var row = i + 1;
      var now = new Date();

      sh.getRange(row, 6).setValue(collected);
      sh.getRange(row, 7).setValue(petty);
      sh.getRange(row, 8).setValue(refunds);
      sh.getRange(row, 9).setValue(expected);
      sh.getRange(row, 10).setValue(actual);
      sh.getRange(row, 11).setValue(variance);
      sh.getRange(row, 12).setValue(status);
      sh.getRange(row, 14).setValue(now);

      // Post tallying adjustment to Ledger
      var adj = '';
      if (payload.postAdjustment && mismatch) {
        var led = acc_sheet_(ACC_CFG.LEDGER), aid = "ADJ-" + Date.now().toString().slice(-9);
        if (variance > 0)  // Surplus
          led.appendRow([aid, now, 'Adjustment', 'CASH_OVER', counter, target, 'Drawer Surplus', 'Cash', Math.abs(variance), 0, acc_str_(payload.user) || 'UNKNOWN', '', "FALSE", reason]);
        else               // Shortage
          led.appendRow([aid, now, 'Adjustment', 'CASH_SHORT', counter, target, 'Drawer Shortage', 'Cash', 0, Math.abs(variance), acc_str_(payload.user) || 'UNKNOWN', '', "FALSE", reason]);
        adj = ' Adjustment ' + aid + ' posted.';
      }

      acc_audit_(payload.user, 'CLOSE_SHIFT', 'Shift_Register', target, 'Expected: ' + expected,
        'Actual: ' + actual, 'Variance: ' + variance + (status === 'MISMATCH' ? ' [MISMATCH]' : '') + (reason ? ' | ' + reason : ''));
      SpreadsheetApp.flush();
      return { success: true, message: "Shift closed. Variance ₹" + variance + "." + adj, variance: variance, status: status };
    }
    return { success: false, message: "Shift " + target + " not found." };
  } catch (e) { return { success: false, message: "Error: " + e.message }; }
  finally { lock.releaseLock(); }
}