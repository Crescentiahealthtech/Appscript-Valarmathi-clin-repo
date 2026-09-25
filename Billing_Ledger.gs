// ============================================================================
// Billing_Ledger.gs — Crescentia HealthTech / CresRx
// The pharmacy and lab billing ledgers: every bill in a period, settled and
// unsettled, in the same shape the hospital billing ledger already had.
// ----------------------------------------------------------------------------
// WHY. Hospital billing had a ledger — a date range, a status filter, a
// search, totals, and reprint and collect on every row. The pharmacy had a
// "Dashboard & Search" and a credit panel; the lab had tabs for today,
// yesterday and "unsettled", and no way at all to see last month's bills or
// reprint an unsettled one. So "what did the lab bill in August, and what is
// still owed on it" had no screen.
//
// Both desks now return the same row shape, which Billing_Ledger_Kit.html
// (in the shell) draws for either:
//
//   { id, date, time, day, sortMs, ageDays, patientId, patientName, mobile,
//     detail, doctor, type, mode, net, paid, balance, refund, status,
//     settledAt, by, admissionId }
//
//   status is one of:
//     PAID        settled at the counter
//     IP_SETTLED  settled on the patient's discharge bill
//     PART_PAID   some paid, some owed
//     UNPAID      nothing paid yet (an OP credit bill)
//     ON_ACCOUNT  on an admission's running bill, settled at discharge
//     CANCELLED   voided — kept on record, counted nowhere
//
// Filters: from/to (yyyy-MM-dd, either may be blank for "any date"), status
// (ALL, SETTLED, UNSETTLED, or one of the above), search, limit.
// ============================================================================

var CRESC_LEDGER_SETTLED   = ['PAID', 'IP_SETTLED'];
var CRESC_LEDGER_UNSETTLED = ['PART_PAID', 'UNPAID', 'ON_ACCOUNT'];

function cresc_ledgerMoney_(v) { var n = parseFloat(v); return isNaN(n) ? 0 : Math.round(n * 100) / 100; }

/** The filter, cleaned. */
function cresc_ledgerFilter_(f) {
  f = f || {};
  var day = function (v) { v = String(v || '').trim(); return /^\d{4}-\d{2}-\d{2}$/.test(v) ? v : ''; };
  return {
    from: day(f.from), to: day(f.to),
    status: String(f.status || 'ALL').trim().toUpperCase(),
    q: String(f.search || '').trim().toLowerCase(),
    limit: Math.min(Math.max(parseInt(f.limit, 10) || 500, 1), 3000)
  };
}

/** Whether a row passes the filter. */
function cresc_ledgerKeep_(r, f) {
  if (f.from && (!r.day || r.day < f.from)) return false;
  if (f.to && (!r.day || r.day > f.to)) return false;
  var st = r.status;
  if (f.status === 'SETTLED' && CRESC_LEDGER_SETTLED.indexOf(st) === -1) return false;
  if (f.status === 'UNSETTLED' && CRESC_LEDGER_UNSETTLED.indexOf(st) === -1) return false;
  if (f.status !== 'ALL' && f.status !== 'SETTLED' && f.status !== 'UNSETTLED' && st !== f.status) return false;
  if (f.q) {
    var hay = [r.id, r.patientId, r.patientName, r.mobile, r.detail, r.doctor, r.admissionId]
      .join(' ').toLowerCase();
    if (hay.indexOf(f.q) === -1) return false;
  }
  return true;
}

/**
 * Totals over the rows the filter kept — cancelled bills counted separately
 * and nowhere else — with the money taken split by mode and what is still
 * owed split by age.
 */
function cresc_ledgerTotals_(rows) {
  var t = { count: 0, net: 0, paid: 0, balance: 0, refund: 0,
            cancelled: 0, cancelledNet: 0, unsettled: 0,
            byMode: {}, aging: { d0_7: 0, d8_30: 0, d31_90: 0, d90p: 0 } };
  rows.forEach(function (r) {
    if (r.status === 'CANCELLED') { t.cancelled++; t.cancelledNet = cresc_ledgerMoney_(t.cancelledNet + r.net); return; }
    t.count++;
    t.net = cresc_ledgerMoney_(t.net + r.net);
    t.paid = cresc_ledgerMoney_(t.paid + r.paid);
    t.balance = cresc_ledgerMoney_(t.balance + r.balance);
    t.refund = cresc_ledgerMoney_(t.refund + (r.refund || 0));
    if (r.paid > 0 && r.status !== 'IP_SETTLED') {
      var m = String(r.mode || 'CASH').toUpperCase();
      t.byMode[m] = cresc_ledgerMoney_((t.byMode[m] || 0) + r.paid);
    }
    if (CRESC_LEDGER_UNSETTLED.indexOf(r.status) !== -1 && r.balance > 0) {
      t.unsettled++;
      var a = r.ageDays || 0;
      var k = a <= 7 ? 'd0_7' : a <= 30 ? 'd8_30' : a <= 90 ? 'd31_90' : 'd90p';
      t.aging[k] = cresc_ledgerMoney_(t.aging[k] + r.balance);
    }
  });
  return t;
}

/** Sort newest first (oldest first for an unsettled list: the oldest debt is the first call). */
function cresc_ledgerFinish_(rows, f) {
  var kept = rows.filter(function (r) { return cresc_ledgerKeep_(r, f); });
  if (f.status === 'UNSETTLED') kept.sort(function (a, b) { return a.sortMs - b.sortMs; });
  else kept.sort(function (a, b) { return b.sortMs - a.sortMs; });
  var totals = cresc_ledgerTotals_(kept);
  return { success: true, message: '', count: kept.length, totals: totals,
           rows: kept.slice(0, f.limit), truncated: kept.length > f.limit };
}

/** Common date fields for a row. */
function cresc_ledgerWhen_(v, tz) {
  var d = cresc_parseDate_(v);
  if (!d || cresc_isSheetEpoch_(d)) return { date: '', time: '', day: '', sortMs: 0, ageDays: 0 };
  return {
    date: Utilities.formatDate(d, tz, 'dd-MMM-yyyy'),
    time: Utilities.formatDate(d, tz, 'hh:mm a'),
    day: Utilities.formatDate(d, tz, 'yyyy-MM-dd'),
    sortMs: d.getTime(),
    ageDays: Math.max(0, Math.floor((Date.now() - d.getTime()) / 86400000))
  };
}

// ---------------------------------------------------------------------------
// PHARMACY
// ---------------------------------------------------------------------------

/**
 * FRONTEND ENTRY. The pharmacy billing ledger.
 * @param {Object} filter  see the top of this file
 */
function getPharmacyLedger(filter, sessionToken) {
  try {
    var actor = crescRequire_(sessionToken, ['pharmacy.bill', 'accounts.read']);
    var f = cresc_ledgerFilter_(filter);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = Session.getScriptTimeZone();
    var sh = ss.getSheetByName(PH_SHEETS.INVOICES);
    if (!sh || sh.getLastRow() < 2) return cresc_ledgerFinish_([], f);

    var d = sh.getDataRange().getValues();
    var h = d[0].map(function (x) { return String(x || '').trim(); });
    var c = function (n, fb) { var i = h.indexOf(n); return i === -1 ? fb : i; };
    var C = { no: c('Invoice_No', 0), ts: c('Timestamp', 1), type: c('Bill_Type', 2), pid: c('Patient_ID', 3),
              name: c('Patient_Name', 4), mob: c('Mobile', 5), doc: c('Doctor', 9), net: c('Net', 13),
              mode: c('Pay_Mode', 14), pay: c('Pay_Status', 16), st: c('Status', 17), n: c('Item_Count', 18),
              by: c('Created_By', 19), at: c('Settled_At', 21), sby: c('Settled_By', 22) };

    var refunds = (typeof _refundInfoByInvoice_ === 'function') ? _refundInfoByInvoice_(ss) : {};
    var onTab = cresc_ledgerOnTab_('PHARMACY');

    var rows = [];
    for (var i = 1; i < d.length; i++) {
      var no = String(d[i][C.no] || '').trim();
      if (!no) continue;
      var w = cresc_ledgerWhen_(d[i][C.ts], tz);
      var net = cresc_ledgerMoney_(d[i][C.net]);
      var pay = String(d[i][C.pay] || '').trim().toUpperCase();
      var st = String(d[i][C.st] || 'ACTIVE').trim().toUpperCase();
      var refund = cresc_ledgerMoney_((refunds[no.toUpperCase()] || {}).refund || 0);
      var mode = String(d[i][C.mode] || 'CASH').trim().toUpperCase();

      var status, paid = 0, balance = 0;
      if (st === 'CANCELLED') { status = 'CANCELLED'; }
      else if (pay === 'IP_SETTLED') { status = 'IP_SETTLED'; paid = net; }
      else if (pay === 'PAID') { status = 'PAID'; paid = cresc_ledgerMoney_(Math.max(0, net - refund)); }
      else {
        balance = cresc_ledgerMoney_(Math.max(0, net - refund));
        status = onTab[no.toUpperCase()] ? 'ON_ACCOUNT' : 'UNPAID';
        if (!balance) status = 'PAID';
      }
      var count = parseInt(d[i][C.n], 10) || 0;
      rows.push({
        id: no, date: w.date, time: w.time, day: w.day, sortMs: w.sortMs, ageDays: w.ageDays,
        patientId: String(d[i][C.pid] || ''), patientName: String(d[i][C.name] || '') || 'Walk-in',
        mobile: String(d[i][C.mob] || ''),
        detail: count + ' item' + (count === 1 ? '' : 's') + (st === 'RETURNED' ? ' · returned' :
                st === 'PARTIAL-RETURN' ? ' · part returned' : ''),
        doctor: String(d[i][C.doc] || ''), type: String(d[i][C.type] || ''),
        mode: mode, net: net, paid: paid, balance: balance, refund: refund, status: status,
        settledAt: d[i][C.at] ? cresc_formatDate_(d[i][C.at], 'dd-MMM-yyyy hh:mm a') : '',
        by: String(d[i][C.by] || ''), admissionId: ''
      });
    }
    var out = cresc_ledgerFinish_(rows, f);
    out.canCollect = actor.permissions.indexOf('pharmacy.bill') !== -1 ||
                     actor.permissions.indexOf('accounts.settle') !== -1;
    return out;
  } catch (err) {
    return { success: false, rows: [], message: cresc_reason_(err) };
  }
}

/** { SOURCE_REF: true } for every pharmacy or lab bill still ON_TAB on an admission. */
function cresc_ledgerOnTab_(source) {
  var out = {};
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('IP_Charges');
    if (!sh || sh.getLastRow() < 2) return out;
    var d = sh.getDataRange().getValues();
    var h = d[0].map(function (x) { return String(x || '').trim(); });
    var cS = h.indexOf('Source'), cR = h.indexOf('Source_Ref'), cT = h.indexOf('Status');
    if (cS < 0 || cR < 0 || cT < 0) return out;
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][cS]).toUpperCase() !== source) continue;
      if (String(d[i][cT]).toUpperCase() !== 'ON_TAB') continue;
      out[String(d[i][cR] || '').trim().toUpperCase()] = true;
    }
  } catch (e) {}
  return out;
}

// ---------------------------------------------------------------------------
// LAB
// ---------------------------------------------------------------------------

/**
 * FRONTEND ENTRY. The lab billing ledger.
 * @param {Object} filter  see the top of this file
 */
function getLabLedger(filter, sessionToken) {
  try {
    var actor = crescRequire_(sessionToken, ['lab.bill', 'accounts.read']);
    var f = cresc_ledgerFilter_(filter);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = Session.getScriptTimeZone();
    var sh = ss.getSheetByName('LAB_BILLING');
    if (!sh || sh.getLastRow() < 2) return cresc_ledgerFinish_([], f);

    var d = sh.getDataRange().getValues();
    var h = d[0].map(function (x) { return String(x || '').trim(); });
    var c = function (n) { return h.indexOf(n); };
    var C = { id: c('BillID'), order: c('OrderID'), pid: c('PatientID'), name: c('PatientName'),
              cat: c('BillingCategory'), adm: c('AdmissionID'), tests: c('TestsJSON'),
              net: c('NetAmount'), mode: c('PaymentMode'), paid: c('PaidAmount'), bal: c('BalanceAmount'),
              st: c('PaymentStatus'), rcpt: c('ReceiptNumber'), at: c('BilledAt'), by: c('BilledBy') };
    var get = function (row, k) { return C[k] === -1 ? '' : row[C[k]]; };

    // How many collections each bill has had since it was raised.
    var collected = {};
    try {
      labb_settlements_().forEach(function (s) {
        var k = String(s.BillID || '').trim();
        if (!k) return;
        collected[k] = collected[k] || { n: 0, amount: 0 };
        collected[k].n++;
        collected[k].amount = cresc_ledgerMoney_(collected[k].amount + (Number(s.Amount) || 0));
      });
    } catch (e) {}

    var rows = [];
    for (var i = 1; i < d.length; i++) {
      var id = String(get(d[i], 'id') || '').trim();
      if (!id) continue;
      var w = cresc_ledgerWhen_(get(d[i], 'at'), tz);
      var net = cresc_ledgerMoney_(get(d[i], 'net'));
      var paid = cresc_ledgerMoney_(get(d[i], 'paid'));
      var raw = String(get(d[i], 'st') || '').trim().toUpperCase();
      var mode = String(get(d[i], 'mode') || 'CASH').trim().toUpperCase();
      var adm = String(get(d[i], 'adm') || '').trim();
      var isIp = mode === 'IP_ACCOUNT' || String(get(d[i], 'cat') || '').toUpperCase() === 'IP_ACCOUNT';
      var bal = cresc_ledgerMoney_(get(d[i], 'bal'));
      if (!bal && raw === 'ON_ACCOUNT') bal = net;

      var status;
      if (raw === 'CANCELLED') { status = 'CANCELLED'; paid = 0; bal = 0; }
      else if (raw === 'IP_SETTLED') { status = 'IP_SETTLED'; paid = net; bal = 0; }
      else if (raw === 'ON_ACCOUNT' || (isIp && raw !== 'PAID')) { status = 'ON_ACCOUNT'; }
      else if (bal > 0) { status = paid > 0 ? 'PART_PAID' : 'UNPAID'; }
      else { status = 'PAID'; if (!paid) paid = net; }

      var tests = '';
      try { tests = parseTestsForUI_(get(d[i], 'tests')); } catch (e) { tests = ''; }
      var col = collected[id];
      rows.push({
        id: id, receipt: String(get(d[i], 'rcpt') || id), orderId: String(get(d[i], 'order') || ''),
        date: w.date, time: w.time, day: w.day, sortMs: w.sortMs, ageDays: w.ageDays,
        patientId: String(get(d[i], 'pid') || ''), patientName: String(get(d[i], 'name') || '') || 'Unknown',
        mobile: '', detail: (tests || 'Lab tests') + (col ? ' · ' + col.n + ' later payment' + (col.n === 1 ? '' : 's') : ''),
        doctor: '', type: isIp ? 'IP' : 'OP', mode: mode,
        net: net, paid: paid, balance: bal, refund: 0, status: status,
        settledAt: '', by: String(get(d[i], 'by') || ''), admissionId: adm
      });
    }
    var out = cresc_ledgerFinish_(rows, f);
    out.canCollect = actor.permissions.indexOf('lab.bill') !== -1 ||
                     actor.permissions.indexOf('accounts.settle') !== -1;
    return out;
  } catch (err) {
    return { success: false, rows: [], message: cresc_reason_(err) };
  }
}
