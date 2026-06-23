// =========================================================================
// 💸 CRESCENTIA — PAYABLES ("Where the money goes")
// Strictly separates RECEIVING a bill (a liability accrual, no cash moves)
// from PAYING it (cash leaves -> Finance_Master_Ledger expense).
// Entity_Type: VENDOR | DOCTOR | PAYROLL.
// Extends Accounts_Payable with Due_Date, GST_Input, Paid_At, Logged_By, Notes.
// Depends on Accounts.js: ACC_CFG, acc_sheet_, acc_money_, acc_str_, acc_toDate_, acc_audit_.
// =========================================================================

var PAY_CFG = { SHEET: 'Accounts_Payable', TYPES: ['VENDOR', 'DOCTOR', 'PAYROLL', 'PETTY', 'REFUND', 'EXPENSE'] };
var PAY_HEADERS = ['Payable_ID', 'Date', 'Entity_Type', 'Vendor_Doctor_Name', 'Document_Type', 'Invoice_Ref', 'Total_Amount', 'Amount_Paid', 'Pending_Due', 'Status', 'Due_Date', 'GST_Input', 'Paid_At', 'Logged_By', 'Notes', 'Drawer'];

function pay_sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(PAY_CFG.SHEET);
  if (!sh) { sh = ss.insertSheet(PAY_CFG.SHEET); sh.appendRow(PAY_HEADERS); sh.setFrozenRows(1); return sh; }
  // ensure the extended columns exist (older sheet had 10)
  var hdr = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), PAY_HEADERS.length)).getValues()[0];
  for (var c = 0; c < PAY_HEADERS.length; c++) if (acc_str_(hdr[c]).trim() !== PAY_HEADERS[c]) sh.getRange(1, c + 1).setValue(PAY_HEADERS[c]);
  return sh;
}
function pay_objs_() {
  var sh = pay_sheet_(), d = sh.getDataRange().getValues(); if (d.length < 2) return [];
  var h = d[0].map(function (x) { return acc_str_(x).trim(); }), out = [];
  for (var i = 1; i < d.length; i++) { var o = { _row: i + 1 }; for (var c = 0; c < h.length; c++) o[h[c]] = d[i][c]; out.push(o); }
  return out;
}
function pay_col_(sh, name) { return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (x) { return acc_str_(x).trim(); }).indexOf(name) + 1; }
function pay_id_() { return 'PAY-' + Date.now().toString().slice(-9) + Math.floor(Math.random() * 90 + 10); }
function pay_days_(due) { var d = acc_toDate_(due); if (!d) return null; return Math.floor(((new Date()).setHours(0, 0, 0, 0) - d.setHours(0, 0, 0, 0)) / 86400000); }

// RECEIVE a bill (no cash). payload = {entityType, name, documentType, invoiceRef,
//   amount, gstInput, dueDate, notes, user, payNow, payMode}
function recordPayable(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var entity = acc_str_(payload.entityType).toUpperCase();
    if (PAY_CFG.TYPES.indexOf(entity) === -1) entity = 'VENDOR';
    var name = acc_str_(payload.name).trim();
    if (!name) return { success: false, message: "Vendor / payee name required." };
    var amt = acc_money_(payload.amount);
    if (amt <= 0) return { success: false, message: "Amount must be greater than 0." };

    var sh = pay_sheet_(), id = pay_id_(), now = new Date(), user = acc_str_(payload.user) || 'UNKNOWN';
    var doc = acc_str_(payload.documentType) || ({ PAYROLL: 'SALARY', DOCTOR: 'PAYOUT', PETTY: 'PETTY_CASH', REFUND: 'REFUND', EXPENSE: 'EXPENSE' }[entity] || 'INVOICE');
    var drawer = acc_str_(payload.drawer);
    // petty / refund / expense / payroll / doctor payouts are typically paid immediately
    var payNow = payload.payNow === true;
    var paid = payNow ? amt : 0, pending = amt - paid, status = payNow ? 'PAID' : 'PENDING';
    sh.appendRow([id, now, entity, name, doc, acc_str_(payload.invoiceRef), amt, paid, pending, status,
      acc_str_(payload.dueDate), acc_money_(payload.gstInput), payNow ? now : '', user, acc_str_(payload.notes), drawer]);
    acc_audit_(user, 'PAYABLE_RECORD', PAY_CFG.SHEET, id, '', amt, entity + ' / ' + name + (drawer ? ' / ' + drawer : ''));

    if (payNow) {
      var mode = acc_str_(payload.payMode) || 'Bank';
      var ledEntity = (entity === 'PETTY' && drawer) ? drawer : name;
      acc_sheet_(ACC_CFG.LEDGER).appendRow([id + '-P', now, 'Payable', entity, ledEntity, id, doc + ' · ' + name, mode, 0, amt, user, '', 'FALSE', acc_str_(payload.notes)]);
      acc_audit_(user, 'PAYABLE_PAY', PAY_CFG.SHEET, id, '', amt, mode + ' (pay-now)');
    }
    SpreadsheetApp.flush();
    return { success: true, message: (payNow ? "Recorded & paid ₹" : "Liability recorded ₹") + amt + " · " + name + ".", payableId: id };
  } catch (e) { return { success: false, message: "Error: " + e.message }; }
  finally { lock.releaseLock(); }
}

// PAY (full or partial) -> cash leaves to the ledger as an expense.
function payPayable(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var target = acc_str_(payload.payableId).trim();
    var amt = acc_money_(payload.amount);
    if (!target) return { success: false, message: "Missing payable reference." };
    if (amt <= 0) return { success: false, message: "Payment must be greater than 0." };

    var sh = pay_sheet_(), d = sh.getDataRange().getValues(), h = d[0].map(function (x) { return acc_str_(x).trim(); });
    var cId = h.indexOf('Payable_ID'), cTot = h.indexOf('Total_Amount'), cPaid = h.indexOf('Amount_Paid'), cPend = h.indexOf('Pending_Due'), cStat = h.indexOf('Status'), cPaidAt = h.indexOf('Paid_At'), cEnt = h.indexOf('Entity_Type'), cName = h.indexOf('Vendor_Doctor_Name'), cDoc = h.indexOf('Document_Type');
    for (var i = 1; i < d.length; i++) {
      if (acc_str_(d[i][cId]).trim() !== target) continue;
      var total = acc_money_(d[i][cTot]), paid = acc_money_(d[i][cPaid]), pending = acc_money_(total - paid);
      if (pending <= 0) return { success: false, message: "Already fully paid." };
      if (amt > pending + 0.01) return { success: false, message: "Payment exceeds pending ₹" + pending + "." };
      var newPaid = acc_money_(paid + amt), newPending = acc_money_(total - newPaid);
      var status = newPending <= 0.01 ? 'PAID' : 'PARTIAL';
      var now = new Date(), row = i + 1, user = acc_str_(payload.user) || 'UNKNOWN';
      var mode = acc_str_(payload.payMode) || 'Bank', entity = acc_str_(d[i][cEnt]), name = acc_str_(d[i][cName]), doc = acc_str_(d[i][cDoc]);

      sh.getRange(row, cPaid + 1).setValue(newPaid);
      sh.getRange(row, cPend + 1).setValue(newPending);
      sh.getRange(row, cStat + 1).setValue(status);
      if (cPaidAt >= 0) sh.getRange(row, cPaidAt + 1).setValue(now);

      acc_sheet_(ACC_CFG.LEDGER).appendRow([target + '-P' + Date.now().toString().slice(-5), now, 'Payable', entity, name, target, doc + ' · ' + name, mode, 0, amt, user, '', 'FALSE', acc_str_(payload.notes)]);
      acc_audit_(user, 'PAYABLE_PAY', PAY_CFG.SHEET, target, 'Pending: ' + pending, 'Paid: ' + amt, mode + ' | now ' + status);
      SpreadsheetApp.flush();
      return { success: true, message: "Paid ₹" + amt + " · " + name + " (" + status + ").", status: status, pending: newPending };
    }
    return { success: false, message: "Payable " + target + " not found." };
  } catch (e) { return { success: false, message: "Error: " + e.message }; }
  finally { lock.releaseLock(); }
}

// Dashboard for the Payables screen: open queue + aging + recent paid.
function getPayables(filter) {
  try {
    var f = acc_str_(filter).toUpperCase();
    var open = [], paidRecent = [], totalPending = 0, overdue = 0, dueWeek = 0, byType = { VENDOR: 0, DOCTOR: 0, PAYROLL: 0 };
    pay_objs_().forEach(function (r) {
      var entity = acc_str_(r['Entity_Type']).toUpperCase(), status = acc_str_(r['Status']).toUpperCase();
      if (f && f !== 'ALL' && entity !== f) return;
      var pending = acc_money_(r['Pending_Due']), days = pay_days_(r['Due_Date']);
      var rec = {
        id: acc_str_(r['Payable_ID']), entity: entity, name: acc_str_(r['Vendor_Doctor_Name']),
        doc: acc_str_(r['Document_Type']), ref: acc_str_(r['Invoice_Ref']),
        total: acc_money_(r['Total_Amount']), paid: acc_money_(r['Amount_Paid']), pending: pending,
        status: status, due: acc_str_(r['Due_Date']) ? Utilities.formatDate(acc_toDate_(r['Due_Date']), ACC_CFG.TZ, 'dd-MMM') : '', drawer: acc_str_(r['Drawer']),
        overdueDays: (days !== null && days > 0 && pending > 0) ? days : 0
      };
      if (status === 'PAID') { paidRecent.push(rec); return; }
      open.push(rec); totalPending += pending; byType[entity] = (byType[entity] || 0) + pending;
      if (rec.overdueDays > 0) overdue += pending; else if (days !== null && days >= -7) dueWeek += pending;
    });
    open.sort(function (a, b) { return b.overdueDays - a.overdueDays; });
    paidRecent.reverse();
    return {
      success: true, open: open, paidRecent: paidRecent.slice(0, 15),
      summary: { totalPending: acc_money_(totalPending), overdue: acc_money_(overdue), dueWeek: acc_money_(dueWeek), byType: byType }
    };
  } catch (e) { return { success: false, message: e.message }; }
}