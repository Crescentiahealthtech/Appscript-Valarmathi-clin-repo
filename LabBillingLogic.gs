// ==========================================
// 💰 LAB BILLING DESK BACKEND ENGINE
// ==========================================

function getLabBillingWorkspace(sessionToken) {
  try { crescRequire_(sessionToken, ['lab.bill', 'accounts.read']); }
  catch (err) { return { success: false, message: cresc_reason_(err) }; }
  return lab_billingWorkspace_();
}

/**
 * THE SAME READ, WITHOUT THE PERMISSION CHECK, for getLabDailyCollection and
 * the admin dashboard, both of which validate the caller first and both of
 * which used to get "success: false" from passing no token — which is why
 * the dashboard's lab collection read as zero on days money had been taken.
 */
function lab_billingWorkspace_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const billingSheet = ss.getSheetByName("LAB_BILLING");
    const ordersSheet = ss.getSheetByName("LAB_ORDERS");

    if (!billingSheet || !ordersSheet) {
      return { success: false, message: "Database Sheets missing." };
    }

    // 1. Establish strict Midnight-to-Midnight Boundaries for "Today" (IST Timezone)
    const tz = Session.getScriptTimeZone();
    const now = new Date();
    const todayStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");
    let d = new Date(now); d.setDate(d.getDate() - 1);
    const yesterdayStr = Utilities.formatDate(d, tz, "yyyy-MM-dd");
    // Cancellations stay on the desk for a week, then leave the working view.
    // They are never deleted: "Find older" reads them back (labListCancelled).
    const cancelCutoff = now.getTime() - LABB_CANCELLED_DAYS * 86400000;
    const dayOf = function (v) {
      var dt = (typeof cresc_parseDate_ === 'function') ? cresc_parseDate_(v) : new Date(v);
      return (dt && !isNaN(dt.getTime())) ? dt : null;
    };
    const fmt = function (dt) { return dt ? dt.toLocaleString('en-IN') : ''; };

    // Today's settlements of older balances count toward today's collection.
    const settledToday = labb_settlementsOn_(todayStr, tz);

    // Admissions already settled: an on-account lab bill against one of these
    // missed the discharge bill and is flagged, not quietly carried.
    const settledIPs = {};
    try {
      const ss2 = ss.getSheetByName('IP_Settlements');
      if (ss2 && ss2.getLastRow() > 1) {
        ss2.getRange(2, 2, ss2.getLastRow() - 1, 1).getValues()
          .forEach(function (r) { if (r[0]) settledIPs[String(r[0]).trim()] = true; });
      }
    } catch (e) {}

    // 2. Fetch & Parse Billing Data
    const bData = billingSheet.getDataRange().getValues();
    const bHeaders = bData.length > 0 ? bData.shift() : [];

    let billedOrderIds = new Set();
    let billsList = [];
    let stats = { pendingCount: 0, todayOP: settledToday.total, todayIP: 0, todayCount: 0,
                  unsettledCount: 0, unsettledTotal: 0, unsettledOP: 0, unsettledIP: 0,
                  olderCancelled: 0 };

    bData.forEach(row => {
      let b = {};
      bHeaders.forEach((h, i) => b[h] = row[i]);
      if (!b.BillID) return;

      const status = String(b.PaymentStatus || '').trim().toUpperCase();
      // A CANCELLED bill is not a bill. Counting its order as "already
      // billed" is what would strand that order: it would never come back to
      // the pending queue and could never be billed correctly, which is the
      // usual reason for voiding one in the first place.
      const isCancelled = status === 'CANCELLED';
      if (b.OrderID && !isCancelled) billedOrderIds.add(b.OrderID);

      const billDate = dayOf(b.BilledAt);
      const billDay = billDate ? Utilities.formatDate(billDate, tz, "yyyy-MM-dd") : '';
      const isStrictlyToday = billDay === todayStr;
      const isIP = (b.PaymentMode === 'IP_ACCOUNT' || b.BillingCategory === 'IP_ACCOUNT' || !!b.AdmissionID);
      const net = Number(b.NetAmount) || 0;
      const paid = Number(b.PaidAmount) || 0;
      // An on-account bill carries its whole net as the balance even where
      // the column was left blank — the same rule the dashboard uses.
      const balance = Number(b.BalanceAmount) || (status === 'ON_ACCOUNT' ? net : 0);

      const card = function (tabType) {
        return {
          tabType: tabType,
          status: status,
          cancelReason: b.CancelReason || '',
          cancelledBy: b.CancelledBy || '',
          orderId: b.OrderID || '',
          billId: b.BillID,
          paymentMode: b.PaymentMode || 'CASH',
          patientName: b.PatientName || 'Unknown',
          patientId: b.PatientID || '',
          admissionId: b.AdmissionID || '',
          testNames: parseTestsForUI_(b.TestsJSON) || "Lab Tests",
          receiptNumber: b.ReceiptNumber || b.BillID,
          billedAt: fmt(billDate),
          net: net, paid: paid, balance: balance,
          discount: Number(b.DiscountAmount) || 0,
          isIP: isIP,
          ipSettledWithout: isIP && status === 'ON_ACCOUNT' && !!settledIPs[String(b.AdmissionID || '').trim()],
          isStrictlyToday: isStrictlyToday
        };
      };

      if (isCancelled) {
        const when = dayOf(b.CancelledAt) || billDate;
        if (when && when.getTime() >= cancelCutoff) billsList.push(card('CANCELLED'));
        else stats.olderCancelled++;
        return;
      }

      // UNSETTLED: every bill with money still owed, WHENEVER it was raised.
      // This is the lab's share of the dashboard's Pending Credit, and it had
      // no screen: part-paid and unpaid OP bills were filed under "Paid", and
      // only today's and yesterday's were listed at all.
      const open = balance > 0 && (status === 'PENDING' || status === 'PARTIAL' ||
                                   status === 'ON_ACCOUNT' || status === 'CREDIT');
      if (open) {
        billsList.push(card('UNSETTLED'));
        stats.unsettledCount++;
        stats.unsettledTotal += balance;
        if (isIP) stats.unsettledIP += balance; else stats.unsettledOP += balance;
      }

      if (isStrictlyToday || billDay === yesterdayStr) {
        if (isIP) billsList.push(card('IP'));
        else if (status === 'PAID' || paid > 0) billsList.push(card('PAID'));
      }

      if (isStrictlyToday) {
        stats.todayCount++;
        // Money TAKEN today, not billed today: an unpaid OP bill used to add
        // its whole net to "Today Collection".
        if (isIP) stats.todayIP += net;
        else stats.todayOP += Math.min(paid, net) - (settledToday.byBill[b.BillID] || 0);
      }
    });

    // 3. Fetch PENDING Orders
    const oData = ordersSheet.getDataRange().getValues();
    const oHeaders = oData.length > 0 ? oData.shift() : [];

    oData.forEach(row => {
      let o = {};
      oHeaders.forEach((h, i) => o[h] = row[i]);
      if (!o.OrderID) return;

      var oStatus = String(o.OrderStatus || '').trim().toUpperCase();

      // Cancelled orders stay visible for a week, then leave the working
      // view like cancelled bills do.
      if (oStatus === 'CANCELLED') {
        const when = dayOf(o.CancelledAt) || dayOf(o.CreatedAt);
        if (!when || when.getTime() < cancelCutoff) { stats.olderCancelled++; return; }
        billsList.push({
          tabType: 'CANCELLED',
          orderId: o.OrderID,
          billId: '',
          paymentMode: '',
          patientName: o.PatientName || 'Unknown',
          patientId: o.PatientID || '',
          testNames: o.TestNames || 'Lab Tests',
          receiptNumber: '',
          billedAt: fmt(when),
          net: 0, paid: 0, balance: 0,
          discount: 0,
          cancelReason: o.CancelReason || '',
          cancelledBy: o.CancelledBy || '',
          isStrictlyToday: false
        });
        return;
      }

      if (!billedOrderIds.has(o.OrderID) && oStatus !== 'DELETE') {
        stats.pendingCount++;
        billsList.push({
          tabType: 'PENDING',
          orderId: o.OrderID,
          billId: '',
          paymentMode: '',
          patientName: o.PatientName || 'Unknown',
          patientId: o.PatientID || '',
          testNames: o.TestNames || "Pending Tests",
          receiptNumber: '',
          billedAt: fmt(dayOf(o.CreatedAt)),
          net: 0, paid: 0, balance: 0,
          discount: 0,
          isStrictlyToday: false
        });
      }
    });

    ['todayOP', 'todayIP', 'unsettledTotal', 'unsettledOP', 'unsettledIP'].forEach(function (k) {
      stats[k] = Math.round(stats[k] * 100) / 100;
    });
    stats.cancelledDays = LABB_CANCELLED_DAYS;
    return { success: true, bills: billsList, stats: stats, settlementsToday: settledToday.rows };

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ---------------------------------------------------------------------------
// SETTLING A BALANCE
//
// A part-paid or unpaid OP lab bill used to be settleable only from the
// Finance Hub's receivables tab — which did not list OP lab bills at all
// (acc_labRows_ counted ON_ACCOUNT only). The lab desk now collects it, and
// every collection is a row on LAB_SETTLEMENTS with its own time and mode, so
// the lab till counts the cash on the day it was actually taken.
// ---------------------------------------------------------------------------

var LABB_CANCELLED_DAYS = 7;
var LABB_SETTLEMENTS = 'LAB_SETTLEMENTS';
var LABB_SETTLEMENT_HEADERS = ['SettlementID', 'BillID', 'PatientID', 'Amount', 'PaymentMode',
                               'SettledAt', 'SettledBy', 'Note'];

function labb_settlementSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(LABB_SETTLEMENTS);
  if (!sh) {
    sh = ss.insertSheet(LABB_SETTLEMENTS);
    sh.getRange(1, 1, 1, LABB_SETTLEMENT_HEADERS.length).setValues([LABB_SETTLEMENT_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Every settlement row, as objects. [] when the sheet does not exist yet. */
function labb_settlements_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LABB_SETTLEMENTS);
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getDataRange().getValues();
  var h = data.shift().map(function (x) { return String(x).trim(); });
  return data.filter(function (r) { return r[0]; }).map(function (r) {
    var o = {}; h.forEach(function (k, i) { o[k] = r[i]; }); return o;
  });
}

/** Settlements taken on one day: { total, byBill: {billId: amount}, rows } */
function labb_settlementsOn_(dayStr, tz) {
  var out = { total: 0, byBill: {}, rows: [] };
  labb_settlements_().forEach(function (r) {
    var dt = (typeof cresc_parseDate_ === 'function') ? cresc_parseDate_(r.SettledAt) : new Date(r.SettledAt);
    if (!dt || isNaN(dt.getTime())) return;
    if (Utilities.formatDate(dt, tz, 'yyyy-MM-dd') !== dayStr) return;
    var amt = Number(r.Amount) || 0;
    out.total += amt;
    out.byBill[r.BillID] = (out.byBill[r.BillID] || 0) + amt;
    out.rows.push({ billId: String(r.BillID), amount: amt, mode: String(r.PaymentMode || ''),
                    at: dt.toLocaleString('en-IN'), by: String(r.SettledBy || '') });
  });
  return out;
}

/**
 * The write. Adds `amount` to the bill's paid figure, reduces its balance,
 * and logs the collection. The caller has checked permission and the lock.
 * @return {{success, message, balance?}}
 */
function labb_settle_(billId, amount, mode, who, note) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LAB_BILLING');
  if (!sh) return { success: false, message: 'LAB_BILLING is missing.' };
  var values = sh.getDataRange().getValues();
  var col = {};
  values[0].forEach(function (x, i) { col[String(x).trim()] = i; });
  ['BillID', 'NetAmount', 'PaidAmount', 'BalanceAmount', 'PaymentStatus'].forEach(function (k) {
    if (col[k] === undefined) throw new Error('LAB_BILLING has no ' + k + ' column.');
  });
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][col.BillID]).trim() !== String(billId).trim()) continue;
    var status = String(values[i][col.PaymentStatus] || '').toUpperCase();
    if (status === 'CANCELLED') return { success: false, message: 'Bill ' + billId + ' was cancelled.' };
    if (status === 'PAID') return { success: false, message: 'Bill ' + billId + ' is already settled.' };
    var net = Number(values[i][col.NetAmount]) || 0;
    var paid = Number(values[i][col.PaidAmount]) || 0;
    var bal = Number(values[i][col.BalanceAmount]) || (status === 'ON_ACCOUNT' ? net : Math.max(0, net - paid));
    var amt = Math.round((Number(amount) || 0) * 100) / 100;
    if (!(amt > 0)) return { success: false, message: 'Enter the amount received.' };
    if (amt > bal + 0.009) {
      return { success: false, message: 'That is more than the ₹' + bal.toFixed(2) + ' still owed on this bill.' };
    }
    var newPaid = Math.round((paid + amt) * 100) / 100;
    var newBal = Math.round((bal - amt) * 100) / 100;
    var row = i + 1;
    sh.getRange(row, col.PaidAmount + 1).setValue(newPaid);
    sh.getRange(row, col.BalanceAmount + 1).setValue(newBal);
    sh.getRange(row, col.PaymentStatus + 1).setValue(newBal <= 0 ? 'PAID' : 'PARTIAL');
    var now = new Date();
    labb_settlementSheet_().appendRow([
      'LSET-' + Utilities.formatDate(now, 'Asia/Kolkata', 'yyMMdd-HHmmss') + '-' +
        Utilities.getUuid().substring(0, 4).toUpperCase(),
      String(billId), String(values[i][col.PatientID !== undefined ? col.PatientID : 0] || ''),
      amt, String(mode || 'CASH').toUpperCase(), now, String(who || ''), String(note || '')
    ]);
    SpreadsheetApp.flush();
    try { labAudit_('BILL_SETTLED', 'BILL', billId, null, { amount: amt, mode: mode, by: who, balance: newBal }); }
    catch (e) {}
    return { success: true, balance: newBal,
             message: newBal <= 0 ? 'Bill ' + billId + ' is now fully paid.'
                                  : '₹' + amt.toFixed(2) + ' received. ₹' + newBal.toFixed(2) + ' still owed.' };
  }
  return { success: false, message: 'Bill ' + billId + ' was not found.' };
}

/**
 * FRONTEND ENTRY (lab billing desk). Collects all or part of what is owed on
 * an OP lab bill. IP on-account bills are settled at discharge, not here.
 * @param {{billId, amount, payMode, note?}} payload
 */
function settleLabBillBalance(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var actor = crescRequire_(sessionToken, ['lab.bill', 'accounts.settle']);
    payload = payload || {};
    var mode = String(payload.payMode || 'CASH').toUpperCase();
    if (['CASH', 'UPI', 'CARD', 'BANK'].indexOf(mode) === -1) {
      return { success: false, message: 'Choose how it was paid: cash, UPI, card or bank.' };
    }
    if (typeof acc_isLocked_ === 'function' && typeof acc_period_ === 'function' &&
        acc_isLocked_(acc_period_(new Date()))) {
      return { success: false, message: 'This month is locked in the Finance Hub; settlements are frozen.' };
    }
    return labb_settle_(payload.billId, payload.amount, mode,
                        actor.displayName || actor.username, payload.note);
  } catch (err) {
    return { success: false, message: cresc_reason_(err) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function parseTestsForUI_(jsonStr) {
  try {
    let arr = JSON.parse(jsonStr);
    return arr.map(t => t.testName).join(", ");
  } catch (e) {
    return "Lab Tests";
  }
}

function getLabDailyCollection(sessionToken) {
  try { crescRequire_(sessionToken, ['lab.bill', 'accounts.read']); }
  catch (err) { return { success: false, message: cresc_reason_(err) }; }
  return lab_dailyCollection_();
}

/** THE SAME FIGURES, WITHOUT THE PERMISSION CHECK, for the admin dashboard. */
function lab_dailyCollection_() {
  try {
    const ws = lab_billingWorkspace_();
    if (!ws.success) throw new Error(ws.message);

    let dayBills = ws.bills.filter(b => (b.tabType === 'PAID' || b.tabType === 'IP') && b.isStrictlyToday === true);
    dayBills.sort((a, b) => new Date(a.billedAt) - new Date(b.billedAt));

    // settlements: balances of earlier bills collected today — part of the
    // day's takings, with their own mode.
    return { success: true, bills: dayBills, totalNet: ws.stats.todayOP,
             settlements: ws.settlementsToday || [] };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Generates the Physical HTML Cashier Receipt
 */
// ==========================================
// 🖨️ LAB RECEIPT PRINT ENGINE (BACKEND)
// ==========================================

/**
 * Generates the physical HTML for the Lab Receipt Pop-up
 * Matches the premium UI/UX of the Lab Integration Engine
 */
function getLabReceiptHtml(billId, sessionToken) {
  try {
    crescRequire_(sessionToken, ['lab.bill', 'accounts.read']);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("LAB_BILLING");
    if (!sheet) throw new Error("LAB_BILLING sheet not found.");

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    let b = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][headers.indexOf("BillID")] === billId) {
        b = {
          billId: data[i][headers.indexOf("BillID")] || '',
          orderId: data[i][headers.indexOf("OrderID")] || '',
          patientId: data[i][headers.indexOf("PatientID")] || '',
          patientName: data[i][headers.indexOf("PatientName")] || 'Unknown Patient',
          category: data[i][headers.indexOf("BillingCategory")] || '',
          itemsJSON: data[i][headers.indexOf("TestsJSON")] || '[]',
          gross: Number(data[i][headers.indexOf("GrossAmount")]) || 0,
          discPct: Number(data[i][headers.indexOf("DiscountPercent")]) || 0,
          discAmt: Number(data[i][headers.indexOf("DiscountAmount")]) || 0,
          net: Number(data[i][headers.indexOf("NetAmount")]) || 0,
          payMode: data[i][headers.indexOf("PaymentMode")] || '',
          paid: Number(data[i][headers.indexOf("PaidAmount")]) || 0,
          balance: Number(data[i][headers.indexOf("BalanceAmount")]) || 0,
          payStatus: data[i][headers.indexOf("PaymentStatus")] || '',
          receipt: data[i][headers.indexOf("ReceiptNumber")] || billId,
          billedAtRaw: data[i][headers.indexOf("BilledAt")],
          billedAt: cresc_formatDate_(data[i][headers.indexOf("BilledAt")], 'dd-MMM-yyyy hh:mm a')
        };
        break;
      }
    }
    
    if (!b) throw new Error("Bill not found in database.");

    // THE PAYMENTS SO FAR. A part-paid or unpaid bill used to reprint as the
    // original invoice and nothing else — "Paid ₹200" with no word of the
    // ₹300 collected last week — so the reprint could not serve as the
    // receipt the patient was asking for. Every collection is a
    // LAB_SETTLEMENTS row; what was paid at billing is the rest of PaidAmount.
    const settlements = labb_settlements_().filter(function (r) {
      return String(r.BillID || '').trim() === String(billId).trim();
    }).map(function (r) {
      return { at: cresc_formatDate_(r.SettledAt, 'dd-MMM-yyyy hh:mm a'),
               amount: Number(r.Amount) || 0, mode: String(r.PaymentMode || ''),
               by: String(r.SettledBy || ''), note: String(r.Note || '') };
    });
    const later = settlements.reduce(function (t, r) { return t + r.amount; }, 0);
    const atBilling = Math.max(0, Math.round((b.paid - later) * 100) / 100);
    const cancelled = String(b.payStatus).toUpperCase() === 'CANCELLED';
    const paymentRows = []
      .concat(atBilling > 0 ? [{ at: b.billedAt, amount: atBilling, mode: b.payMode, by: '', note: 'At billing' }] : [])
      .concat(settlements);

    // Parse itemized tests
    let items = [];
    try { if (b.itemsJSON) items = JSON.parse(b.itemsJSON); } catch (e) {}
    
    let itemRows = items.map((it, idx) => {
      return `
        <tr>
          <td style="padding:8px 10px; border-bottom:1px solid #e5e7eb;">${idx + 1}</td>
          <td style="padding:8px 10px; border-bottom:1px solid #e5e7eb;">
            <strong>${_esc_(it.testName)}</strong>
            ${it.testId ? `<br><span style="font-size:10px;color:#6b7280;">${_esc_(it.testId)}</span>` : ''}
          </td>
          <td style="padding:8px 10px; border-bottom:1px solid #e5e7eb; text-align:right;">₹${Number(it.price).toFixed(2)}</td>
        </tr>
      `;
    }).join('');

    // Branding Properties (Fallback if not set)
    // One letterhead for every document the clinic prints — see
    // Clinic_Profile.gs. This file used to default to "Crescentia
    // HealthTech" while LabIntegrationEngine.gs defaulted to "Crescentia
    // Clinic" and the pharmacy invoice had the name typed into its markup,
    // so a patient holding all three receipts was holding paper from what
    // looked like different organisations.
    const clinic        = cresc_clinic_();
    const clinicName    = clinic.name;
    const clinicAddress = clinic.address;
    const clinicPhone   = clinic.phone;
    const gstNumber     = clinic.gstin;

    const isIp = (b.category === 'IP_ACCOUNT');
    
    const payLine = isIp
      ? `<div style="font-weight:700;color:#0369a1;">Posted to IP Account &bull; Settled at discharge</div>`
      : `<div>Payment: <strong>${_esc_(b.payMode)}</strong> &bull; Status: <strong style="${b.payStatus === 'PAID' ? 'color:#10b981;' : 'color:#dc2626;'}">${_esc_(b.payStatus)}</strong></div>`;

    // Construct Clean, Premium HTML
    const html = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <title>Lab Invoice - ${b.receipt}</title>
        <style>
          *{box-sizing:border-box;margin:0;padding:0;font-family:'Helvetica Neue',Arial,sans-serif;}
          body{background:#fff;color:#111827;}
          @media print{@page{margin:1cm;} .no-print{display:none!important;} body{background:#fff;} }
        </style>
      </head>
      <body>
        <div style="max-width:760px;margin:18px auto;border:1px solid #e5e7eb;border-radius:10px;overflow:hidden;">
          
          <div style="padding:18px 24px;border-bottom:2px solid #0369a1;display:flex;justify-content:space-between;align-items:flex-start;">
            <div>
              <div style="font-size:20px;font-weight:800;color:#0369a1;text-transform:uppercase;">${_esc_(clinicName)}</div>
              ${clinicAddress ? `<div style="font-size:12px;color:#6b7280;margin-top:4px;">${_esc_(clinicAddress)}</div>` : ''}
              ${clinicPhone ? `<div style="font-size:12px;color:#6b7280;">Phone: ${_esc_(clinicPhone)}</div>` : ''}
              ${gstNumber ? `<div style="font-size:11px;color:#6b7280;margin-top:2px;">GSTIN: ${_esc_(gstNumber)}</div>` : ''}
            </div>
            <div style="text-align:right;">
              <div style="font-size:16px;font-weight:800;letter-spacing:1px;color:#111827;">LAB INVOICE</div>
              <div style="font-size:12px;color:#4b5563;margin-top:4px;">Inv: <strong>${_esc_(b.receipt)}</strong></div>
              <div style="font-size:11px;color:#6b7280;">Date: ${_esc_(b.billedAt)}</div>
            </div>
          </div>

          <div style="padding:12px 24px;background:#f9fafb;border-bottom:1px solid #e5e7eb;display:flex;justify-content:space-between;font-size:13px;">
            <div><span style="color:#6b7280;">Patient:</span> <strong>${_esc_(b.patientName)}</strong> &bull; ${_esc_(b.patientId)}</div>
            <div><span style="color:#6b7280;">Order ID:</span> ${_esc_(b.orderId)}</div>
          </div>

          <div style="padding:8px 24px;">
            <table style="width:100%;border-collapse:collapse;font-size:13px;">
              <thead>
                <tr style="background:#f3f4f6;">
                  <th style="padding:8px 10px;text-align:left;color:#374151;font-size:11px;text-transform:uppercase;">#</th>
                  <th style="padding:8px 10px;text-align:left;color:#374151;font-size:11px;text-transform:uppercase;">Investigation</th>
                  <th style="padding:8px 10px;text-align:right;color:#374151;font-size:11px;text-transform:uppercase;">Amount</th>
                </tr>
              </thead>
              <tbody>
                ${itemRows}
              </tbody>
            </table>
          </div>

          <div style="padding:8px 24px 16px;display:flex;justify-content:flex-end;">
            <table style="font-size:13px;min-width:260px;">
              <tr><td style="padding:4px 10px;color:#6b7280;">Gross Amount</td><td style="padding:4px 10px;text-align:right;">₹${b.gross.toFixed(2)}</td></tr>
              ${b.discAmt > 0 ? `<tr><td style="padding:4px 10px;color:#10b981;">Discount (${b.discPct}%)</td><td style="padding:4px 10px;text-align:right;color:#10b981;">- ₹${b.discAmt.toFixed(2)}</td></tr>` : ''}
              <tr style="border-top:2px solid #111827;">
                <td style="padding:6px 10px;font-weight:800;font-size:15px;">Net Payable</td>
                <td style="padding:6px 10px;text-align:right;font-weight:800;font-size:16px;color:#0369a1;">₹${b.net.toFixed(2)}</td>
              </tr>
              ${!isIp ? `<tr><td style="padding:4px 10px;color:#6b7280;">Paid</td><td style="padding:4px 10px;text-align:right;">₹${b.paid.toFixed(2)}</td></tr>` : ''}
              ${!isIp && b.balance > 0 ? `<tr><td style="padding:4px 10px;color:#dc2626;">Balance</td><td style="padding:4px 10px;text-align:right;color:#dc2626;font-weight:700;">₹${b.balance.toFixed(2)}</td></tr>` : ''}
            </table>
          </div>

          ${(!isIp && paymentRows.length) ? `
          <div style="padding:0 24px 14px;">
            <div style="font-size:11px;color:#6b7280;font-weight:700;text-transform:uppercase;margin-bottom:6px;">Payments received</div>
            <table style="width:100%;border-collapse:collapse;font-size:12px;">
              <thead><tr style="background:#f3f4f6;">
                <th style="padding:5px 8px;text-align:left;">Date</th><th style="padding:5px 8px;text-align:left;">Mode</th>
                <th style="padding:5px 8px;text-align:left;">Note</th><th style="padding:5px 8px;text-align:right;">Amount</th>
              </tr></thead>
              <tbody>${paymentRows.map(function (r) {
                return '<tr><td style="padding:5px 8px;border-bottom:1px solid #e5e7eb;">' + _esc_(r.at) + '</td>' +
                       '<td style="padding:5px 8px;border-bottom:1px solid #e5e7eb;">' + _esc_(r.mode) + '</td>' +
                       '<td style="padding:5px 8px;border-bottom:1px solid #e5e7eb;">' + _esc_(r.note || (r.by ? 'Collected by ' + r.by : '')) + '</td>' +
                       '<td style="padding:5px 8px;border-bottom:1px solid #e5e7eb;text-align:right;">₹' + Number(r.amount).toFixed(2) + '</td></tr>';
              }).join('')}</tbody>
            </table>
          </div>` : ''}

          ${cancelled ? `<div style="padding:0 24px 14px;"><span style="border:2px solid #b91c1c;color:#b91c1c;display:inline-block;padding:4px 12px;font-weight:800;letter-spacing:.05em;border-radius:4px;font-size:12px;">CANCELLED — NOT A RECEIPT</span></div>` : ''}

          <div style="padding:12px 24px;background:#f9fafb;border-top:1px solid #e5e7eb;font-size:12px;color:#374151;display:flex;justify-content:space-between;align-items:center;">
            ${payLine}
            <div style="color:#9ca3af;">This is a computer-generated invoice.</div>
          </div>

        </div>

        <div class="no-print" style="text-align:center;padding:14px;margin-bottom:20px;">
          <button onclick="window.print();" style="background:#0369a1;color:#fff;border:none;padding:12px 28px;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer;box-shadow:0 4px 6px rgba(3,105,161,0.2);">
            🖨️ Print / Save PDF
          </button>
        </div>
      </body>
      </html>
    `;

    return { success: true, html: html, balance: b.balance, status: b.payStatus };
  } catch (error) {
    return { success: false, message: cresc_reason_(error) };
  }
}

function _esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"'`=\/]/g, function (s) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;', '/': '&#x2F;', '`': '&#x60;', '=': '&#x3D;' }[s];
  });
}

// ==========================================
// 🚀 LAB INVOICE COMMUNICATION ENGINE (WHATSAPP & EMAIL)
// ==========================================

/**
 * Server-Side function: Generates Invoice PDF, saves to Drive, and returns public link for WhatsApp.
 */
function generateAndStoreLabInvoicePDF(billId, sessionToken) {
  try {
    // 1. Generate HTML using existing engine
    crescRequire_(sessionToken, 'lab.bill');

    // ── DPDP s.6 / s.5: the patient's COMMUNICATION consent, checked here ──
    // The register has carried this purpose since it was built and nothing
    // read it. This dispatch now refuses unless the patient has agreed to be
    // sent documents this way, and the refusal says how to ask them. See
    // DPDP_Dispatch.gs.
    var __pid = dpdp_resolvePatientFor_('LAB_INVOICE', billId);
    var __gate = dpdpRequireDispatchConsent_(__pid, 'WHATSAPP',
                   crescActor_(sessionToken), 'LAB_INVOICE');
    if (!__gate.ok) return { success: false, code: 'CONSENT_REQUIRED',
                             consentState: __gate.state, patientId: __pid,
                             notice: __gate.notice, message: __gate.message };
    const reportResponse = getLabReceiptHtml(billId); 
    if (!reportResponse.success) throw new Error("HTML Generation Failed: " + reportResponse.message);

    // 2. Convert to PDF Blob
    const htmlBlob = Utilities.newBlob(reportResponse.html, 'text/html', 'invoice.html');
    const pdfBlob = htmlBlob.getAs('application/pdf');
    pdfBlob.setName("Crescentia_Lab_Invoice_" + billId + ".pdf");

    // 3. Drive Folder Architecture specifically for Invoices
    const rootFolderName = "Crescentia_Lab_Invoices";
    let rootFolder;
    const rootFolders = DriveApp.getFoldersByName(rootFolderName);
    if (rootFolders.hasNext()) {
      rootFolder = rootFolders.next();
    } else {
      rootFolder = DriveApp.createFolder(rootFolderName);
    }

    const now = new Date();
    const yearStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy");
    const monthStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMMM");

    let yearFolder;
    const yearFolders = rootFolder.getFoldersByName(yearStr);
    if (yearFolders.hasNext()) yearFolder = yearFolders.next();
    else yearFolder = rootFolder.createFolder(yearStr);

    let monthFolder;
    const monthFolders = yearFolder.getFoldersByName(monthStr);
    if (monthFolders.hasNext()) monthFolder = monthFolders.next();
    else monthFolder = yearFolder.createFolder(monthStr);

    // 4. Save the file
    const file = monthFolder.createFile(pdfBlob);

    // Registered, not just shared. dpdpIssueDocumentLink_ publishes the
    // file to Drive so the patient gets a link WhatsApp can preview and
    // their phone can open, and writes it into Document_Grants with an
    // expiry — after which the daily sweep makes the file private again
    // and every forwarded copy of the link stops working. A bare
    // setSharing(ANYONE_WITH_LINK) here would publish this invoice
    // for ever with nothing able to take it back. See DPDP_Documents.gs.
    // __pid, not billId: the register's Patient_ID column was being given a
    // bill number, so no lab invoice could be listed back to its patient.
    var grant = dpdpIssueDocumentLink_(file, 'LAB_INVOICE', __pid || billId,
                                       (crescActor_(sessionToken) || {}).username || '');
    if (!grant.success) return { success: false, message: grant.message };

    // 5. Return the private, expiring link
    return { success: true, link: grant.url, expiresAt: grant.expiresAt,
             grantId: grant.grantId, message: grant.message };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Server-Side function: Generates Invoice PDF and emails it directly via GMAIL API.
 */
function emailLabInvoicePDF(billId, patientEmail, sessionToken) {
  try {
    crescRequire_(sessionToken, 'lab.bill');

    // ── DPDP s.6 / s.5: the patient's COMMUNICATION consent, checked here ──
    // The register has carried this purpose since it was built and nothing
    // read it. This dispatch now refuses unless the patient has agreed to be
    // sent documents this way, and the refusal says how to ask them. See
    // DPDP_Dispatch.gs.
    var __pid = dpdp_resolvePatientFor_('LAB_INVOICE', billId);
    var __gate = dpdpRequireDispatchConsent_(__pid, 'EMAIL',
                   crescActor_(sessionToken), 'LAB_INVOICE');
    if (!__gate.ok) return { success: false, code: 'CONSENT_REQUIRED',
                             consentState: __gate.state, patientId: __pid,
                             notice: __gate.notice, message: __gate.message };
    const reportResponse = getLabReceiptHtml(billId); 
    if (!reportResponse.success) throw new Error("HTML Generation Failed");

    const htmlBlob = Utilities.newBlob(reportResponse.html, 'text/html', 'invoice.html');
    const pdfBlob = htmlBlob.getAs('application/pdf');
    pdfBlob.setName("Crescentia_Lab_Invoice_" + billId + ".pdf");

    const subject = "Your Payment Receipt - Crescentia Clinic & Diagnostics";
    const plainBody = "Dear Patient, please find your lab invoice attached. Regards, Crescentia Clinic.";

    // Premium HTML Email Layout for GmailApp
    const richHtmlBody = `
      <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden;">
        <div style="background-color: #0369a1; color: white; padding: 20px; text-align: center;">
          <h2 style="margin: 0; letter-spacing: 1px;">CRESCENTIA CLINIC & DIAGNOSTICS</h2>
        </div>
        <div style="padding: 30px;">
          <p style="font-size: 16px;">Dear Patient,</p>
          <p style="font-size: 15px; line-height: 1.5;">Thank you for choosing Crescentia HealthTech. We have received your payment for the recent laboratory investigations.</p>
          
          <div style="background-color: #f0fdf4; border-left: 4px solid #10b981; padding: 15px; margin: 25px 0;">
            <p style="margin: 0; color: #065f46;"><strong>Secure PDF Attached:</strong> Please find your official payment receipt / invoice attached to this email.</p>
          </div>
          
          <p style="font-size: 14px; color: #4b5563;">Wishing you the best of health,<br><br><strong>The Billing Team</strong><br>Crescentia Clinic & Diagnostics</p>
          <hr style="border: 0; border-top: 1px solid #e5e7eb; margin: 30px 0 15px 0;">
          <p style="font-size: 11px; color: #9ca3af; text-align: center; margin: 0;">This is an automatically generated dispatch. Please do not reply to this email.</p>
        </div>
      </div>
    `;

    GmailApp.sendEmail(patientEmail, subject, plainBody, {
      htmlBody: richHtmlBody,
      attachments: [pdfBlob],
      name: "Crescentia Billing Desk"
    });

    return { success: true, message: "Invoice emailed successfully." };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
