// =========================================================================
// 🛡️ CRESCENTIA — INSURANCE / TPA DESK
// Each insurer on a discharge becomes a claim here (linked by Settlement_ID).
// Lifecycle: PENDING -> SUBMITTED -> QUERIED -> APPROVED -> SETTLED /
//   SHORT_SETTLED / REJECTED. Settling posts cash-in (INSURANCE_RECEIPT) and
//   recognizes revenue. Network vs Non-Network filterable.
// Extends Insurance_Claims_Ledger. Depends on Accounts.js helpers.
// =========================================================================

var INS_CFG = { SHEET: 'Insurance_Claims_Ledger' };
var INS_HEADERS = ['Claim_ID', 'Date', 'Patient_ID', 'Patient_Name', 'TPA_Company', 'Policy_No', 'PreAuth_Status', 'Total_Billed', 'Claimed_Amount', 'Approved_Amount', 'Deduction_Amount', 'Patient_Liability_Log', 'Settlement_Status', 'IP_Number', 'Settlement_ID', 'Network', 'Settled_Amount', 'Settled_At', 'TPA_Ref', 'Updated_By', 'Notes'];
var INS_OPEN = ['PENDING', 'SUBMITTED', 'QUERIED', 'APPROVED'];

function ins_sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(INS_CFG.SHEET);
  if (!sh) { sh = ss.insertSheet(INS_CFG.SHEET); sh.appendRow(INS_HEADERS); sh.setFrozenRows(1); return sh; }
  var hdr = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), INS_HEADERS.length)).getValues()[0];
  for (var c = 0; c < INS_HEADERS.length; c++) if (acc_str_(hdr[c]).trim() !== INS_HEADERS[c]) sh.getRange(1, c + 1).setValue(INS_HEADERS[c]);
  return sh;
}
function ins_objs_() {
  var sh = ins_sheet_(), d = sh.getDataRange().getValues(); if (d.length < 2) return [];
  var h = d[0].map(function (x) { return acc_str_(x).trim(); }), out = [];
  for (var i = 1; i < d.length; i++) { var o = { _row: i + 1 }; for (var c = 0; c < h.length; c++) o[h[c]] = d[i][c]; out.push(o); }
  return out;
}
function ins_col_(sh, n) { return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (x) { return acc_str_(x).trim(); }).indexOf(n) + 1; }
function ins_days_(v) { var d = acc_toDate_(v); if (!d) return 0; return Math.floor(((new Date()).setHours(0, 0, 0, 0) - d.setHours(0, 0, 0, 0)) / 86400000); }

// called by settleDischarge — one claim per insurer
function ins_createClaimsForSettlement_(setId, ip, adm, billed, insurers, user) {
  var sh = ins_sheet_(), now = new Date();
  (insurers || []).forEach(function (x) {
    var amt = acc_money_(x.amount); if (amt <= 0) return;
    var row = {
      Claim_ID: 'CLM-' + Date.now().toString().slice(-9) + Math.floor(Math.random() * 90 + 10), Date: now,
      Patient_ID: adm.patientId, Patient_Name: adm.patientName, TPA_Company: acc_str_(x.name), Policy_No: '',
      PreAuth_Status: 'APPROVED', Total_Billed: acc_money_(billed), Claimed_Amount: amt, Approved_Amount: amt,
      Deduction_Amount: 0, Patient_Liability_Log: '', Settlement_Status: 'PENDING',
      IP_Number: ip, Settlement_ID: setId, Network: x.network ? 'TRUE' : 'FALSE', Settled_Amount: 0,
      Settled_At: '', TPA_Ref: '', Updated_By: user, Notes: ''
    };
    sh.appendRow(INS_HEADERS.map(function (h) { return row[h] === undefined ? '' : row[h]; }));
  });
  SpreadsheetApp.flush();
}

function getInsuranceClaims(payload) {
  try {
    payload = payload || {};
    var fStatus = acc_str_(payload.status).toUpperCase(), fNet = acc_str_(payload.network).toUpperCase();
    var claims = [], totalPending = 0, netPending = 0, nonNetPending = 0, byStatus = {};
    ins_objs_().forEach(function (r) {
      var status = acc_str_(r['Settlement_Status']).toUpperCase() || 'PENDING';
      var net = acc_str_(r['Network']).toUpperCase() === 'TRUE';
      var claimed = acc_money_(r['Claimed_Amount']), settled = acc_money_(r['Settled_Amount']);
      var outstanding = Math.max(0, claimed - settled), open = INS_OPEN.indexOf(status) !== -1;
      byStatus[status] = (byStatus[status] || 0) + 1;
      if (open) { totalPending += outstanding; if (net) netPending += outstanding; else nonNetPending += outstanding; }
      if (fStatus && fStatus !== 'ALL' && status !== fStatus) return;
      if (fNet === 'NETWORK' && !net) return;
      if (fNet === 'NON' && net) return;
      var d = acc_toDate_(r['Date']);
      claims.push({
        id: acc_str_(r['Claim_ID']), date: d ? Utilities.formatDate(d, ACC_CFG.TZ, 'dd-MMM') : '',
        ip: acc_str_(r['IP_Number']), patient: acc_str_(r['Patient_Name']), insurer: acc_str_(r['TPA_Company']),
        network: net, policy: acc_str_(r['Policy_No']), preAuth: acc_str_(r['PreAuth_Status']),
        billed: acc_money_(r['Total_Billed']), claimed: claimed, approved: acc_money_(r['Approved_Amount']),
        settled: settled, outstanding: outstanding, deduction: acc_money_(r['Deduction_Amount']),
        status: status, tpaRef: acc_str_(r['TPA_Ref']), ageDays: open ? ins_days_(r['Date']) : 0, notes: acc_str_(r['Notes'])
      });
    });
    claims.sort(function (a, b) { return b.ageDays - a.ageDays; });
    return { success: true, claims: claims, summary: { totalPending: acc_money_(totalPending), netPending: acc_money_(netPending), nonNetPending: acc_money_(nonNetPending), byStatus: byStatus } };
  } catch (e) { return { success: false, message: e.message }; }
}

// raise a claim by hand (OP cashless, standalone, or any non-IP case)
function createManualClaim(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    payload = payload || {};
    var insurer = acc_str_(payload.insurer).trim();
    if (!insurer) return { success: false, message: "Insurer / TPA required." };
    var patient = acc_str_(payload.patientName).trim();
    if (!patient) return { success: false, message: "Patient name required." };
    var claimed = acc_money_(payload.claimedAmount);
    if (claimed <= 0) return { success: false, message: "Claimed amount must be greater than 0." };
    var billed = acc_money_(payload.billedAmount) || claimed;
    var source = (acc_str_(payload.source).toUpperCase() || 'OP');
    var caseRef = acc_str_(payload.caseRef).trim() || source;
    var user = acc_str_(payload.user) || 'UNKNOWN', now = new Date();
    var id = 'CLM-' + Date.now().toString().slice(-9) + Math.floor(Math.random() * 90 + 10);
    var row = {
      Claim_ID: id, Date: now, Patient_ID: acc_str_(payload.patientId), Patient_Name: patient,
      TPA_Company: insurer, Policy_No: acc_str_(payload.policyNo), PreAuth_Status: acc_str_(payload.preAuth).toUpperCase() || 'REQUESTED',
      Total_Billed: billed, Claimed_Amount: claimed, Approved_Amount: claimed, Deduction_Amount: 0,
      Patient_Liability_Log: '', Settlement_Status: 'PENDING', IP_Number: caseRef, Settlement_ID: '',
      Network: payload.network ? 'TRUE' : 'FALSE', Settled_Amount: 0, Settled_At: '', TPA_Ref: '',
      Updated_By: user, Notes: '[' + source + '] ' + acc_str_(payload.notes)
    };
    ins_sheet_().appendRow(INS_HEADERS.map(function (h) { return row[h] === undefined ? '' : row[h]; }));
    acc_audit_(user, 'CLAIM_RAISE', INS_CFG.SHEET, id, '', claimed, source + ' / ' + insurer + ' / ' + patient);
    SpreadsheetApp.flush();
    return { success: true, message: "Claim raised · " + insurer + " · ₹" + claimed + ".", claimId: id };
  } catch (e) { return { success: false, message: "Error: " + e.message }; }
  finally { lock.releaseLock(); }
}

// update lifecycle / pre-auth / approved amount / refs (no cash)
function updateClaim(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var id = acc_str_(payload.claimId).trim(); if (!id) return { success: false, message: "Missing claim." };
    var sh = ins_sheet_(), d = sh.getDataRange().getValues(), h = d[0].map(function (x) { return acc_str_(x).trim(); });
    var cId = h.indexOf('Claim_ID');
    for (var i = 1; i < d.length; i++) {
      if (acc_str_(d[i][cId]).trim() !== id) continue;
      var row = i + 1, claimed = acc_money_(d[i][h.indexOf('Claimed_Amount')]);
      function set(name, v) { var c = h.indexOf(name); if (c >= 0) sh.getRange(row, c + 1).setValue(v); }
      var oldStatus = acc_str_(d[i][h.indexOf('Settlement_Status')]);
      if (payload.status) set('Settlement_Status', acc_str_(payload.status).toUpperCase());
      if (payload.preAuthStatus) set('PreAuth_Status', acc_str_(payload.preAuthStatus).toUpperCase());
      if (payload.policyNo !== undefined) set('Policy_No', acc_str_(payload.policyNo));
      if (payload.tpaRef !== undefined) set('TPA_Ref', acc_str_(payload.tpaRef));
      if (payload.notes !== undefined) set('Notes', acc_str_(payload.notes));
      if (payload.approvedAmount !== undefined && payload.approvedAmount !== '') {
        var app = acc_money_(payload.approvedAmount); set('Approved_Amount', app); set('Deduction_Amount', acc_money_(claimed - app));
      }
      set('Updated_By', acc_str_(payload.user) || 'UNKNOWN');
      acc_audit_(payload.user, 'CLAIM_UPDATE', INS_CFG.SHEET, id, oldStatus, acc_str_(payload.status) || oldStatus, acc_str_(payload.notes));
      SpreadsheetApp.flush();
      return { success: true, message: "Claim updated." };
    }
    return { success: false, message: "Claim " + id + " not found." };
  } catch (e) { return { success: false, message: "Error: " + e.message }; }
  finally { lock.releaseLock(); }
}

// record bank settlement -> cash-in + revenue recognized; close the claim
function settleClaim(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var id = acc_str_(payload.claimId).trim(); if (!id) return { success: false, message: "Missing claim." };
    var settled = acc_money_(payload.settledAmount); if (settled < 0) return { success: false, message: "Invalid amount." };
    var sh = ins_sheet_(), d = sh.getDataRange().getValues(), h = d[0].map(function (x) { return acc_str_(x).trim(); });
    var cId = h.indexOf('Claim_ID');
    for (var i = 1; i < d.length; i++) {
      if (acc_str_(d[i][cId]).trim() !== id) continue;
      var row = i + 1, now = new Date(), user = acc_str_(payload.user) || 'UNKNOWN';
      var claimed = acc_money_(d[i][h.indexOf('Claimed_Amount')]);
      var insurer = acc_str_(d[i][h.indexOf('TPA_Company')]), ip = acc_str_(d[i][h.indexOf('IP_Number')]), setId = acc_str_(d[i][h.indexOf('Settlement_ID')]);
      var shortfall = acc_money_(claimed - settled);
      var status = (payload.reject === true) ? 'REJECTED' : (shortfall <= 0.01 ? 'SETTLED' : 'SHORT_SETTLED');
      var mode = acc_str_(payload.payMode) || 'Bank';
      function set(name, v) { var c = h.indexOf(name); if (c >= 0) sh.getRange(row, c + 1).setValue(v); }
      set('Settled_Amount', settled); set('Settled_At', now); set('Settlement_Status', status);
      set('Deduction_Amount', Math.max(0, shortfall)); set('TPA_Ref', acc_str_(payload.tpaRef)); set('Updated_By', user);
      if (payload.shortfallToPatient === true && shortfall > 0) set('Patient_Liability_Log', 'Shortfall ₹' + shortfall + ' to patient · ' + Utilities.formatDate(now, ACC_CFG.TZ, 'dd-MMM-yyyy'));

      if (settled > 0)
        acc_sheet_(ACC_CFG.LEDGER).appendRow([id + '-S', now, 'Insurance', 'INSURANCE_RECEIPT', insurer, id, 'Claim settled · ' + insurer + (ip ? ' · ' + ip : ''), mode, settled, 0, user, '', 'FALSE', acc_str_(payload.notes)]);
      acc_audit_(user, 'CLAIM_SETTLE', INS_CFG.SHEET, id, 'Claimed: ' + claimed, 'Settled: ' + settled, status + (shortfall > 0 ? ' | shortfall ' + shortfall : ''));

      // mark the IP settlement insurance done when no open claims remain for it
      try {
        if (setId) {
          var anyOpen = false;
          ins_objs_().forEach(function (c) { if (acc_str_(c['Settlement_ID']).trim() === setId && INS_OPEN.indexOf(acc_str_(c['Settlement_Status']).toUpperCase()) !== -1 && acc_str_(c['Claim_ID']).trim() !== id) anyOpen = true; });
          if (!anyOpen) {
            var stSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('IP_Settlements');
            if (stSh) { var sd = stSh.getDataRange().getValues(), sh2 = sd[0].map(function (x) { return acc_str_(x).trim(); }), si = sh2.indexOf('Settlement_ID'), ss = sh2.indexOf('Insurance_Status'); for (var k = 1; k < sd.length; k++) if (acc_str_(sd[k][si]).trim() === setId) { stSh.getRange(k + 1, ss + 1).setValue('SETTLED'); break; } }
          }
        }
      } catch (e) {}
      SpreadsheetApp.flush();
      return { success: true, message: "Claim " + status + " · ₹" + settled + " received." + (shortfall > 0 ? " Shortfall ₹" + shortfall + "." : ""), status: status };
    }
    return { success: false, message: "Claim " + id + " not found." };
  } catch (e) { return { success: false, message: "Error: " + e.message }; }
  finally { lock.releaseLock(); }
}