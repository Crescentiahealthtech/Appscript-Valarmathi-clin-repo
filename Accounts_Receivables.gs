// =========================================================================
// 💰 CRESCENTIA HEALTHTECH — ACCOUNTS RECEIVABLES (Phase 2)
// Unified running-tab across Pharmacy (PENDING credit) + Lab (ON_ACCOUNT).
// Grouped by AdmissionID (IP) else Patient_ID (OP / staff credit).
// Settlement flips the SOURCE row to PAID + audits. No ledger income posting.
// Depends on Accounts.js (acc_* helpers, normalizers).
// =========================================================================

// READ: open dues grouped into running tabs.
function getReceivables() {
  try {
    // Hospital invoices (OP consultations, procedures, packages) join pharmacy
    // and lab here: a part-paid or credit bill raised at the billing desk is a
    // receivable like any other, and used to be invisible to this screen.
    var open = acc_pharmaRows_()
      .concat(acc_labRows_())
      .concat(acc_hospitalRowsSafe_())
      .filter(function (r) { return r.open && r.balance > 0; });
    var groups = {}, today = Date.now();

    open.forEach(function (r) {
      var key = r.admissionId ? ('IP:' + r.admissionId) : ('PT:' + (r.patientId || r.name));
      if (!groups[key]) {
        groups[key] = {
          key: key, type: r.admissionId ? 'IP' : 'OP/Credit',
          patientId: r.patientId, name: r.name, admissionId: r.admissionId,
          totalDue: 0, oldest: null, agingDays: 0, bills: []
        };
      }
      var g = groups[key];
      g.totalDue = acc_money_(g.totalDue + r.balance);
      if (!g.name && r.name) g.name = r.name;
      var d = r.billDate;
      if (d && (!g.oldest || d.getTime() < g.oldest)) g.oldest = d.getTime();
      g.bills.push({
        source: r.source, billId: r.billId, date: acc_fmtTs_(r.billDate),
        amount: r.balance, mode: r.mode
      });
    });

    var list = Object.keys(groups).map(function (k) {
      var g = groups[k];
      g.agingDays = g.oldest ? Math.floor((today - g.oldest) / 86400000) : 0;
      g.oldestDate = g.oldest ? acc_dayStr_(new Date(g.oldest)) : '';
      delete g.oldest;
      return g;
    }).sort(function (a, b) { return b.agingDays - a.agingDays; });

    var totalOpen = list.reduce(function (s, g) { return s + g.totalDue; }, 0);
    return { success: true, groups: list, totalOpen: acc_money_(totalOpen), count: list.length };
  } catch (e) { return { success: false, message: e.message }; }
}

// WRITE: settle a single source bill. payload = {source, billId, payMode, loggedBy}
function settleReceivableBill(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!payload || !payload.billId) return { success: false, message: "Missing bill reference." };
    var source = acc_str_(payload.source).toUpperCase();
    var mode = acc_str_(payload.payMode) || 'Cash';

    if (acc_isLocked_(acc_period_(new Date())))
      return { success: false, message: "Current period is locked; settlements frozen." };

    if (source === 'PHARMACY') {
      // Reuse the existing pharmacy settler (flips Pay_Status -> PAID, stamps Settled_At/By).
      if (typeof settleCreditBill !== 'function')
        return { success: false, message: "Pharmacy settlement function unavailable." };
      var res = settleCreditBill({ invoiceNo: payload.billId, payMode: mode, txnId: acc_str_(payload.txnId) });
      if (res && res.success)
        acc_audit_(payload.loggedBy, 'SETTLE_RECEIVABLE', 'Pharmacy', payload.billId, 'PENDING', 'PAID', 'Mode: ' + mode);
      return res;
    }

    if (source === 'LAB') {
      var sh = acc_sheet_(ACC_CFG.LAB_BILLING);
      var values = sh.getDataRange().getValues();
      var head = values[0].map(function (h) { return acc_str_(h).trim(); });
      var col = {}; head.forEach(function (h, i) { col[h] = i; });
      var cBill = col['BillID'], cNet = col['NetAmount'], cPaid = col['PaidAmount'],
          cBal = col['BalanceAmount'], cStat = col['PaymentStatus'], cMode = col['PaymentMode'];
      if (cBill === undefined || cStat === undefined)
        return { success: false, message: "LAB_BILLING schema unexpected." };

      for (var i = 1; i < values.length; i++) {
        if (acc_str_(values[i][cBill]).trim() !== acc_str_(payload.billId).trim()) continue;
        if (acc_str_(values[i][cStat]).toUpperCase() === 'PAID')
          return { success: false, message: "Bill " + payload.billId + " already settled." };
        var net = acc_money_(values[i][cNet]);
        var row = i + 1;
        sh.getRange(row, cStat + 1).setValue('PAID');
        if (cMode !== undefined) sh.getRange(row, cMode + 1).setValue(mode);
        if (cPaid !== undefined) sh.getRange(row, cPaid + 1).setValue(net);
        if (cBal !== undefined) sh.getRange(row, cBal + 1).setValue(0);
        acc_audit_(payload.loggedBy, 'SETTLE_RECEIVABLE', 'Lab', payload.billId, 'ON_ACCOUNT', 'PAID', 'Mode: ' + mode);
        SpreadsheetApp.flush();
        return { success: true, message: "Lab bill " + payload.billId + " settled." };
      }
      return { success: false, message: "Lab bill " + payload.billId + " not found." };
    }

    if (source === 'HOSPITAL') {
      if (typeof hb_recordPayment !== 'function')
        return { success: false, message: "Hospital billing module unavailable." };
      // hb_recordPayment does its own session and role check, writes the
      // payment against the invoice and audits it, so settling from here and
      // settling from the billing desk cannot drift apart.
      var hres = hb_recordPayment(acc_str_(payload.sessionToken), payload.billId,
                                  acc_money_(payload.amount) || acc_money_(payload.balance),
                                  mode, acc_str_(payload.txnId));
      return hres;
    }

    return { success: false, message: "Unknown source: " + source };
  } catch (e) { return { success: false, message: "Error: " + e.message }; }
  finally { lock.releaseLock(); }
}