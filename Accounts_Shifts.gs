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

// ===========================================================================
// WHICH MONEY BELONGS TO WHICH DRAWER
// ---------------------------------------------------------------------------
// The Shift_Registers sheet has always had a Counter_Name column and the
// drawer screen has always offered three counters — but acc_cashSince_()
// ignored the counter entirely. It added up EVERY cash pharmacy bill and
// EVERY cash lab bill and offered that same figure as the expected cash in
// all three drawers.
//
// So the pharmacist counting her till at six o'clock was shown a figure that
// included the lab's takings and the front desk's, the lab technician was
// shown the same number, and whichever of them closed first "reconciled"
// against two other people's money. The variance was meaningless, and
// because it was meaningless the mismatch flag on the Finance Hub was
// meaningless too.
//
// A drawer holds what was taken AT THAT COUNTER. These are the sources:
//
//   PHARMACY      pharmacy invoices           (acc_pharmaRows_)
//   LAB           lab bills                   (acc_labRows_)
//   RECEPTION     OP consultations and        (acc_opRows_ +
//                 hospital bills               acc_hospitalRowsSafe_)
//
// and the expenses that come out of a drawer are the ones booked to that
// department. An expense with no department is a clinic-level payment and
// belongs to no single till, so it is left out of all three rather than
// charged to whichever drawer happens to close first.
//
// The three sum to the clinic's cash position, which is what the Finance Hub
// now shows — see getAllDrawerStates().
// ===========================================================================

/** Counter name -> the billing sources whose cash lands in that drawer. */
var ACC_COUNTER_SOURCES = {
  'PHARMACY':  ['Pharmacy'],
  'LAB':       ['Lab'],
  'RECEPTION': ['OP_Consultation', 'Hospital'],
  // Spellings a clinic may have typed into Counter_Name over time.
  'FRONT DESK':  ['OP_Consultation', 'Hospital'],
  'FRONT OFFICE':['OP_Consultation', 'Hospital'],
  'BILLING':     ['OP_Consultation', 'Hospital'],
  'OP':          ['OP_Consultation', 'Hospital']
};

/** The canonical counter key, or '' for one this file does not know. */
function acc_counterKey_(counterName) {
  var k = acc_str_(counterName).toUpperCase().replace(/\s+/g, ' ').trim();
  return ACC_COUNTER_SOURCES[k] ? k : '';
}

/**
 * Does this ledger expense come out of this counter's drawer?
 *
 * Matched on Department, and on the counter name appearing in the entity or
 * the category — the shift's own CASH_SHORT / CASH_OVER adjustments are
 * written with the counter in the Department column, so they reconcile
 * against the right till.
 */
function acc_expenseIsCounters_(row, counterName) {
  var counter = acc_str_(counterName).toUpperCase();
  if (!counter) return false;
  var dept = acc_str_(row['Department']).toUpperCase();
  if (!dept) return false;                 // clinic-level: no single drawer
  return dept === counter || dept.indexOf(counter) !== -1 || counter.indexOf(dept) !== -1;
}

/**
 * Cash movement at ONE counter since a moment — the shift's auto-calc.
 *
 * @param {Date} fromDate
 * @param {string} [counterName]  omit for the whole clinic (the old behaviour,
 *                                kept for the Finance Hub's total)
 */
function acc_cashSince_(fromDate, counterName) {
  var from = (fromDate instanceof Date && !isNaN(fromDate.getTime())) ? fromDate.getTime() : 0;
  var cashIn = 0, cashOut = 0;

  var key = acc_counterKey_(counterName);
  var wanted = key ? ACC_COUNTER_SOURCES[key] : null;   // null = every source

  var rows = acc_pharmaRows_()
    .concat(acc_labRows_())
    .concat(acc_opRows_())
    .concat(acc_hospitalRowsSafe_());

  rows.forEach(function (r) {
    if (!r.realized || r.net <= 0) return;
    if (acc_str_(r.mode).toLowerCase() !== 'cash') return;
    if (wanted && wanted.indexOf(acc_str_(r.source)) === -1) return;
    var d = r.realizedDate || r.billDate;
    if (d && d.getTime() >= from) {
      cashIn += (r.realizedAmount === undefined) ? r.net : r.realizedAmount;
    }
  });

  acc_readObjects_(ACC_CFG.LEDGER).forEach(function (r) {
    var out = acc_money_(r['Amount_Out']);
    if (out <= 0) return;
    if (acc_str_(r['Payment_Mode']).toLowerCase() !== 'cash') return;
    if (counterName && !acc_expenseIsCounters_(r, counterName)) return;
    var d = acc_toDate_(r['Timestamp']);
    if (d && d.getTime() >= from) cashOut += out;
  });

  return { cashIn: acc_money_(cashIn), cashOut: acc_money_(cashOut),
           counter: acc_str_(counterName), scoped: !!key };
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
      // Scoped to THIS counter. It used to be the whole clinic's cash, in
      // every drawer — see the note above acc_cashSince_.
      var c = acc_cashSince_(open.openedAt, open.counter);
      suggestion = { cashIn: c.cashIn, cashOut: c.cashOut,
                     scoped: c.scoped, counter: open.counter };
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
    var shiftId = acc_newId_("SH");
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
        var led = acc_sheet_(ACC_CFG.LEDGER), aid = acc_newId_("ADJ");
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
// ===========================================================================
// THE ROLL-UP: three drawers, one cash position
// ---------------------------------------------------------------------------
// The Finance Hub showed a single "Cash in drawer" figure computed inside
// getAccountsDashboard() by adding every open shift's float and then every
// cash transaction in the clinic since the EARLIEST open shift — which
// counted the pharmacy's takings against the lab's float and back again, and
// produced a number that matched no till in the building.
//
// This is the honest version: each counter reconciled against its own
// sources, and the three added up. A drawer that is closed contributes
// nothing and says so, because "₹0" and "nobody has opened this till today"
// are different facts and only one of them is a problem.
// ===========================================================================

/** The counters the clinic runs, in the order the drawer screen shows them. */
var ACC_COUNTERS = ['Reception', 'Pharmacy', 'Lab'];

/**
 * FRONTEND ENTRY. Every counter's live drawer, and their total.
 *
 * @param {string} [sessionToken]  when given, the caller must hold
 *                                 accounts.read; omitted, it degrades to the
 *                                 same figures the hub already shows.
 * @return {{success, drawers:Array, totals:Object, message:string}}
 */
function getAllDrawerStates(sessionToken) {
  try {
    if (sessionToken && typeof crescRequire_ === 'function') {
      crescRequire_(sessionToken, 'accounts.read');
    }

    var rows = acc_readObjects_(ACC_CFG.SHIFTS);
    var openByCounter = {};
    rows.forEach(function (r) {
      if (acc_str_(r['Status']).toUpperCase() !== 'OPEN') return;
      openByCounter[acc_str_(r['Counter_Name']).toUpperCase()] = r;
    });

    var drawers = [];
    var totals = { opening: 0, cashIn: 0, cashOut: 0, expected: 0, openCount: 0 };

    ACC_COUNTERS.forEach(function (counter) {
      var r = openByCounter[counter.toUpperCase()];
      if (!r) {
        drawers.push({
          counter: counter, open: false, shiftId: '', user: '',
          opening: 0, cashIn: 0, cashOut: 0, expected: 0,
          openedAt: '', note: 'No shift open at this counter.'
        });
        return;
      }

      var opening = acc_money_(r['Opening_Cash']);
      var openedAt = acc_toDate_(r['Timestamp']);
      var c = acc_cashSince_(openedAt, counter);
      var expected = acc_money_(opening + c.cashIn - c.cashOut);

      totals.opening += opening;
      totals.cashIn += c.cashIn;
      totals.cashOut += c.cashOut;
      totals.expected += expected;
      totals.openCount++;

      drawers.push({
        counter: counter,
        open: true,
        shiftId: acc_str_(r['Shift_ID']),
        user: acc_str_(r['Shift_User']),
        opening: opening,
        cashIn: c.cashIn,
        cashOut: c.cashOut,
        expected: expected,
        // A drawer whose timestamp cannot be read would otherwise reconcile
        // against every transaction ever recorded.
        openedAt: openedAt ? acc_fmtTs_(openedAt) : '',
        note: openedAt ? '' : 'This shift has no readable open time, so its ' +
                              'expected cash cannot be trusted.'
      });
    });

    Object.keys(totals).forEach(function (k) {
      if (k !== 'openCount') totals[k] = acc_money_(totals[k]);
    });

    return { success: true, drawers: drawers, totals: totals, message: '' };
  } catch (err) {
    var m = String((err && err.message) || err);
    return { success: false, drawers: [], totals: null,
             message: m.replace('FORBIDDEN: ', '') };
  }
}
