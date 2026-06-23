// =========================================================================
// 💰 CRESCENTIA — FINANCE DASHBOARD (scoped, carry-forward, IP-aware)
// Revenue (cash-basis): counter-paid clinical income (Pay_Status PAID) +
//   IP discharge settlements (Recognized_Realized). IP credit/tab rows are
//   excluded at source. Running Cash/Bank carry forward to period end and
//   include IP_ADVANCE / IP_SETTLEMENT / IP_REFUND / CASH_OVER / CASH_SHORT
//   as pure cash movements (never revenue or expense).
// KPIs: Gross Revenue, Total Expenses, Net Flow, Running Cash · Drawer,
//   Running Bank, Advances Held (liability), Pending Collections (Ins + IP).
// =========================================================================

function accd_isCash_(m) { return String(m || '').toLowerCase().indexOf('cash') !== -1; }
function accd_day_(d) { var x = acc_toDate_(d); return x ? Utilities.formatDate(x, ACC_CFG.TZ, "yyyy-MM-dd") : ''; }
function accd_month_(d) { var x = acc_toDate_(d); return x ? Utilities.formatDate(x, ACC_CFG.TZ, "yyyy-MM") : ''; }
function accd_income_() {
  var rows = acc_pharmaRows_().concat(acc_labRows_());
  if (typeof acc_opRows_ === 'function') { try { rows = rows.concat(acc_opRows_()); } catch (e) {} }
  return rows;
}
function accd_ipSettlements_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName('IP_Settlements');
  if (!sh) return [];
  var d = sh.getDataRange().getValues(); if (d.length < 2) return [];
  var h = d[0].map(function (x) { return acc_str_(x).trim(); }), out = [];
  for (var i = 1; i < d.length; i++) { var o = {}; for (var c = 0; c < h.length; c++) o[h[c]] = d[i][c]; out.push(o); }
  return out;
}

function getFinanceDashboard(scope, periodKey) {
  try {
    scope = (scope || 'OVERALL').toUpperCase();
    var tz = ACC_CFG.TZ;
    var nowMonth = Utilities.formatDate(new Date(), tz, "yyyy-MM");
    var nowDay = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    var month = (scope === 'MONTH') ? (periodKey || nowMonth) : null;
    var day = (scope === 'DAY') ? (periodKey || nowDay) : null;

    function inWin(d) {
      if (scope === 'OVERALL') return true;
      if (scope === 'MONTH') return accd_month_(d) === month;
      return accd_day_(d) === day;
    }
    var cutoff = (scope === 'OVERALL') ? '9999-12-31' : (scope === 'MONTH') ? (month + '-31') : day;
    function upto(d) { var ds = accd_day_(d); return ds && ds <= cutoff; }

    // --- counter-paid clinical income (source sheets) ---
    var collected = 0, cashRunIn = 0, bankRunIn = 0;
    accd_income_().forEach(function (r) {
      if (!r.realized || r.net <= 0) return;     // realized == Pay_Status PAID only
      var d = r.realizedDate || r.billDate, cash = accd_isCash_(r.mode);
      if (inWin(d)) collected += r.net;
      if (upto(d)) { cash ? cashRunIn += r.net : bankRunIn += r.net; }
    });

    // --- IP discharge settlements: recognized revenue (cash-basis portion) ---
    var ipRevenue = 0, insReceipts = 0;
    accd_ipSettlements_().forEach(function (s) {
      var d = acc_toDate_(s['Timestamp']);
      if (inWin(d)) ipRevenue += acc_money_(s['Recognized_Realized']);
    });

    // --- ledger: expenses, transfers, IP & shift cash adjustments ---
    var paidOut = 0, cashRunOut = 0, bankRunOut = 0, depRun = 0, wdRun = 0;
    acc_readObjects_(ACC_CFG.LEDGER).forEach(function (r) {
      var cat = acc_str_(r['Category']).toUpperCase();
      var out = acc_money_(r['Amount_Out']), inn = acc_money_(r['Amount_In']);
      var d = acc_toDate_(r['Timestamp']), cash = accd_isCash_(r['Payment_Mode']);
      // pure cash-position categories (not revenue, not expense)
      if (cat === 'CASH_DEPOSIT') { if (upto(d)) depRun += out; return; }
      if (cat === 'CASH_WITHDRAWAL') { if (upto(d)) wdRun += inn; return; }
      if (cat === 'CASH_OVER' || cat === 'IP_ADVANCE' || cat === 'IP_SETTLEMENT') { if (upto(d)) { cash ? cashRunIn += inn : bankRunIn += inn; } return; }
      // insurance settlement received: realized revenue + bank/cash in
      if (cat === 'INSURANCE_RECEIPT') { if (upto(d)) { cash ? cashRunIn += inn : bankRunIn += inn; } if (inWin(d)) insReceipts += inn; return; }
      if (cat === 'CASH_SHORT' || cat === 'IP_REFUND') { if (upto(d)) { cash ? cashRunOut += out : bankRunOut += out; } return; }
      // genuine expenses
      if (out > 0) {
        if (inWin(d)) paidOut += out;
        if (upto(d)) { cash ? cashRunOut += out : bankRunOut += out; }
      }
    });

    var runningCash = acc_money_(cashRunIn - cashRunOut - depRun + wdRun);
    var runningBank = acc_money_(bankRunIn - bankRunOut + depRun - wdRun);
    var grossRevenue = acc_money_(collected + ipRevenue + insReceipts);
    var net = acc_money_(grossRevenue - paidOut);

    // --- advances held (liability) + open IP tab outstanding ---
    var advHeld = 0, openTab = 0;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var advSh = ss.getSheetByName('IP_Advances');
    if (advSh) { var ad = advSh.getDataRange().getValues(), ah = ad[0].map(function (x) { return acc_str_(x).trim(); }); var aS = ah.indexOf('Status'), aA = ah.indexOf('Amount'); for (var i = 1; i < ad.length; i++) if (acc_str_(ad[i][aS]).toUpperCase() === 'HELD') advHeld += acc_money_(ad[i][aA]); }
    var chSh = ss.getSheetByName('IP_Charges');
    if (chSh) { var cd = chSh.getDataRange().getValues(), ch = cd[0].map(function (x) { return acc_str_(x).trim(); }); var cS = ch.indexOf('Status'), cA = ch.indexOf('Amount'); for (var j = 1; j < cd.length; j++) if (acc_str_(cd[j][cS]).toUpperCase() === 'ON_TAB') openTab += acc_money_(cd[j][cA]); }

    // saved discharge drafts: ward charges entered but not yet posted (active admissions only)
    var activeIP = {}, admSh = ss.getSheetByName('IP_Admissions');
    if (admSh) { var add = admSh.getDataRange().getValues(), ahh = add[0].map(function (x) { return acc_str_(x).trim(); }); var aiIP = ahh.indexOf('IP Number'), aiSt = ahh.indexOf('Status'); for (var k = 1; k < add.length; k++) if (acc_str_(add[k][aiSt]).toUpperCase() === 'ACTIVE') activeIP[acc_str_(add[k][aiIP]).trim()] = true; }
    var drSh = ss.getSheetByName('IP_Discharge_Drafts');
    if (drSh) {
      var dd2 = drSh.getDataRange().getValues(), dh2 = dd2[0].map(function (x) { return acc_str_(x).trim(); }), diIP = dh2.indexOf('IP_Number'), diW = dh2.indexOf('Ward_JSON');
      for (var dr = 1; dr < dd2.length; dr++) {
        if (!activeIP[acc_str_(dd2[dr][diIP]).trim()]) continue;
        try { JSON.parse(acc_str_(dd2[dr][diW]) || '[]').forEach(function (w) { openTab += (w.amount !== undefined && w.amount !== '') ? acc_money_(w.amount) : acc_money_(w.multiplier || 1) * acc_money_(w.rate); }); } catch (e) {}
      }
    }
    var pendingIP = Math.max(0, acc_money_(openTab - advHeld));
    // --- pending insurance: outstanding on open claims (authoritative ledger) ---
    var pendingInsurance = 0, OPEN_INS = ['PENDING', 'SUBMITTED', 'QUERIED', 'APPROVED'];
    var clSh = ss.getSheetByName('Insurance_Claims_Ledger');
    if (clSh) {
      var ld = clSh.getDataRange().getValues(), lh = ld[0].map(function (x) { return acc_str_(x).trim(); });
      var qS = lh.indexOf('Settlement_Status'), qC = lh.indexOf('Claimed_Amount'), qP = lh.indexOf('Settled_Amount');
      for (var m = 1; m < ld.length; m++) if (OPEN_INS.indexOf(acc_str_(ld[m][qS]).toUpperCase()) !== -1) pendingInsurance += Math.max(0, acc_money_(ld[m][qC]) - (qP >= 0 ? acc_money_(ld[m][qP]) : 0));
    }
    var pendingCollections = acc_money_(pendingIP + pendingInsurance);

    // --- open shifts only ---
    var openShifts = [];
    acc_readObjects_(ACC_CFG.SHIFTS).forEach(function (s) {
      if (acc_str_(s['Status']).toUpperCase() !== 'OPEN') return;
      var opening = acc_money_(s['Opening_Cash']);
      var c = (typeof acc_cashSince_ === 'function') ? acc_cashSince_(acc_toDate_(s['Timestamp'])) : { cashIn: 0, cashOut: 0 };
      openShifts.push({ shiftId: acc_str_(s['Shift_ID']), counter: acc_str_(s['Counter_Name']), user: acc_str_(s['Shift_User']), opening: opening, currentExpected: acc_money_(opening + c.cashIn - c.cashOut) });
    });

    var kpis = [
      { label: 'Gross Revenue', value: grossRevenue, color: 'green' },
      { label: 'Total Expenses', value: acc_money_(paidOut), color: 'red' },
      { label: 'Net Flow', value: net, color: 'amber' },
      { label: 'Running Cash · Drawer', value: runningCash, color: 'blue' },
      { label: 'Running Bank (UPI/Card)', value: runningBank, color: 'teal' },
      { label: 'Advances Held (liability)', value: advHeld, color: 'purple' },
      { label: 'Pending Collections (Insurance + IP)', value: pendingCollections, color: 'amber' }
    ];

    return { success: true, scope: scope, periodKey: (month || day || ''), label: (scope === 'OVERALL' ? 'All time' : (month || day)), kpis: kpis, openShifts: openShifts };
  } catch (e) { return { success: false, message: e.message }; }
}

// WRITE: cash <-> bank transfer (unchanged contract).
function recordBankTransfer(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!payload) return { success: false, message: "No data received." };
    var dir = acc_str_(payload.direction).toUpperCase(), amt = acc_money_(payload.amount);
    if (amt <= 0) return { success: false, message: "Amount must be greater than 0." };
    if (dir !== 'DEPOSIT' && dir !== 'WITHDRAW') return { success: false, message: "Invalid transfer direction." };
    if (typeof acc_isLocked_ === 'function' && acc_isLocked_(acc_period_(new Date()))) return { success: false, message: "Current period is locked." };
    var sh = acc_sheet_(ACC_CFG.LEDGER), txnId = "TRF-" + Date.now().toString().slice(-9), ts = new Date(), user = acc_str_(payload.loggedBy) || 'UNKNOWN';
    if (dir === 'DEPOSIT') sh.appendRow([txnId, ts, 'Transfer', 'CASH_DEPOSIT', '', acc_str_(payload.refId), 'Bank Deposit', 'Cash', 0, amt, user, '', "FALSE", acc_str_(payload.notes)]);
    else sh.appendRow([txnId, ts, 'Transfer', 'CASH_WITHDRAWAL', '', acc_str_(payload.refId), 'Bank Withdrawal', 'Cash', amt, 0, user, '', "FALSE", acc_str_(payload.notes)]);
    acc_audit_(user, dir === 'DEPOSIT' ? 'CASH_DEPOSIT' : 'CASH_WITHDRAWAL', 'Finance_Master_Ledger', txnId, '', amt, acc_str_(payload.notes));
    SpreadsheetApp.flush();
    return { success: true, message: (dir === 'DEPOSIT' ? 'Deposit' : 'Withdrawal') + " recorded · ₹" + amt + " moved.", txnId: txnId };
  } catch (e) { return { success: false, message: "Error: " + e.message }; }
  finally { lock.releaseLock(); }
}