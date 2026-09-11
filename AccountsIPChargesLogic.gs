// =========================================================================
// 🏥 CRESCENTIA — IP BILLING ENGINE ("The Running Tab")
// Single authoritative charge ledger per admission. Pharmacy/Lab "Bill to IP"
// mirror one line here (idempotent by Source+Source_Ref). Counter payments
// flip a line to PAID_COUNTER so it leaves the tab. Discharge freezes the
// tab, applies advances, splits payers, recognizes revenue ONCE.
//
// Depends on Accounts.js helpers: ACC_CFG, acc_money_, acc_str_, acc_toDate_,
//   acc_audit_, acc_now_(optional). Uses Finance_Master_Ledger for cash.
//
// Sheets (auto-created):
//   IP_Charges:    Charge_ID, IP_Number, Timestamp, Category, Description,
//                  Source, Source_Ref, Amount, GST, Is_NonPayable, Status,
//                  Settlement_ID, Created_By, Notes
//   IP_Advances:   Advance_ID, IP_Number, Timestamp, Amount, Pay_Mode,
//                  Txn_ID, Status, Settlement_ID, Collected_By
//   IP_Settlements: Settlement_ID, IP_Number, Timestamp, Gross_Tab,
//                  Package_Cap, Package_Adjustment, Capped_Gross,
//                  NonPayable_Total, Advance_Applied, Insurance_Approved,
//                  Insurance_Status, Patient_Liability, Patient_Paid,
//                  Patient_Pay_Mode, Refund_Due, Recognized_Realized,
//                  Settled_By, Notes
// =========================================================================

var IPC_CFG = {
  CHARGES: 'IP_Charges', ADVANCES: 'IP_Advances', SETTLEMENTS: 'IP_Settlements',
  ADMISSIONS: 'IP_Admissions', BEDS: 'Master_Beds',
  CHARGE_CATS: ['ROOM', 'NURSING', 'RMO', 'SURGEON', 'OT', 'PHARMACY', 'LAB', 'CONSUMABLE', 'OTHER']
};

// ---- infra ----------------------------------------------------------------
function ipc_ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function ipc_ensure_(name, headers) {
  var ss = ipc_ss_(), sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(headers); sh.setFrozenRows(1); }
  return sh;
}
function ipc_charges_() { return ipc_ensure_(IPC_CFG.CHARGES, ['Charge_ID', 'IP_Number', 'Timestamp', 'Category', 'Description', 'Source', 'Source_Ref', 'Amount', 'GST', 'Is_NonPayable', 'Status', 'Settlement_ID', 'Created_By', 'Notes']); }
function ipc_advances_() { return ipc_ensure_(IPC_CFG.ADVANCES, ['Advance_ID', 'IP_Number', 'Timestamp', 'Amount', 'Pay_Mode', 'Txn_ID', 'Status', 'Settlement_ID', 'Collected_By']); }
function ipc_settlements_() { return ipc_ensure_(IPC_CFG.SETTLEMENTS, ['Settlement_ID', 'IP_Number', 'Timestamp', 'Gross_Tab', 'Package_Code', 'Package_Cap', 'Package_Adjustment', 'Capped_Gross', 'Discount_Percent', 'Discount_Amount', 'NonPayable_Total', 'Advance_Applied', 'Insurance_Approved', 'Insurer_JSON', 'Insurance_Status', 'Patient_Liability', 'Patient_Paid', 'Patient_Pay_Mode', 'Refund_Due', 'Recognized_Realized', 'Settled_By', 'Notes']); }
function ipc_drafts_() { return ipc_ensure_('IP_Discharge_Drafts', ['IP_Number', 'Updated_At', 'Discount_Percent', 'Package_Code', 'Package_Cap', 'Insurance_Approved', 'Insurers_JSON', 'Ward_JSON', 'Remarks', 'Updated_By']); }
function ipc_los_(doa) { var d = acc_toDate_(doa); if (!d) return 1; var ms = (new Date()).getTime() - d.getTime(); return Math.max(1, Math.floor(ms / 86400000) + 1); }
function ipc_objs_(sh) {
  var d = sh.getDataRange().getValues(); if (d.length < 2) return [];
  var h = d[0].map(function (x) { return acc_str_(x).trim(); }), out = [];
  for (var i = 1; i < d.length; i++) { var o = { _row: i + 1 }; for (var c = 0; c < h.length; c++) o[h[c]] = d[i][c]; out.push(o); }
  return out;
}
function ipc_isTrue_(v) { var s = acc_str_(v).trim().toUpperCase(); return s === 'TRUE' || s === 'YES' || s === '1'; }
function ipc_cash_(m) { return acc_str_(m).toLowerCase().indexOf('cash') !== -1; }
function ipc_id_(p) { return p + '-' + Date.now().toString().slice(-9) + Math.floor(Math.random() * 90 + 10); }

// resolve a patient's single ACTIVE admission (used to auto-route IP credit bills)
function ipc_activeAdmissionByPatient_(patientId) {
  var sh = ipc_ss_().getSheetByName(IPC_CFG.ADMISSIONS); if (!sh) return null;
  var d = sh.getDataRange().getValues();
  for (var i = 1; i < d.length; i++)
    if (acc_str_(d[i][1]).trim().toUpperCase() === acc_str_(patientId).trim().toUpperCase() && acc_str_(d[i][11]).toUpperCase() === 'ACTIVE')
      return acc_str_(d[i][0]);
  return null;
}

// at discharge, close the originating Pharmacy/Lab bill so it leaves the
// PENDING pool and is excluded from source income (recognized via settlement).
function ipc_closeSourceBill_(source, ref, setId) {
  var name = (source === 'LAB') ? 'LAB_BILLING' : 'Pharmacy_Invoices';
  var sh = ipc_ss_().getSheetByName(name); if (!sh) return;
  var d = sh.getDataRange().getValues(); if (d.length < 2) return;
  var h = d[0].map(function (x) { return acc_str_(x).trim(); });
  var idCol = (source === 'LAB') ? h.indexOf('BillID') : h.indexOf('Invoice_No');
  var stCol = (source === 'LAB') ? h.indexOf('PaymentStatus') : h.indexOf('Pay_Status');
  if (idCol < 0 || stCol < 0) return;
  for (var i = 1; i < d.length; i++) {
    if (acc_str_(d[i][idCol]).trim() !== acc_str_(ref).trim()) continue;
    sh.getRange(i + 1, stCol + 1).setValue('IP_SETTLED');
    if (source !== 'LAB') { var sb = h.indexOf('Settled_At'), sy = h.indexOf('Settled_By'); if (sb >= 0) sh.getRange(i + 1, sb + 1).setValue(new Date()); if (sy >= 0) sh.getRange(i + 1, sy + 1).setValue('IP:' + setId); }
    break;
  }
}

function ipc_admission_(ipNumber) {
  var sh = ipc_ss_().getSheetByName(IPC_CFG.ADMISSIONS); if (!sh) return null;
  var d = sh.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) if (acc_str_(d[i][0]).trim() === acc_str_(ipNumber).trim())
    return { row: i + 1, ipNumber: acc_str_(d[i][0]), patientId: acc_str_(d[i][1]), patientName: acc_str_(d[i][2]), ageSex: acc_str_(d[i][3]), doa: acc_str_(d[i][4]), consultant: acc_str_(d[i][9]), status: acc_str_(d[i][11]).toUpperCase(), bed: acc_str_(d[i][8]) };
  return null;
}

// ---- charge entry ---------------------------------------------------------
// Manual clinical charge (room/nursing/surgeon/OT/etc.).
function addIpCharge(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ip = acc_str_(payload.ipNumber).trim();
    var adm = ipc_admission_(ip);
    if (!adm) return { success: false, message: "Admission " + ip + " not found." };
    if (adm.status !== 'ACTIVE') return { success: false, message: "Admission is " + adm.status + " — cannot add charges." };
    var amt = acc_money_(payload.amount);
    if (amt <= 0) return { success: false, message: "Amount must be greater than 0." };
    var cat = acc_str_(payload.category).toUpperCase();
    if (IPC_CFG.CHARGE_CATS.indexOf(cat) === -1) cat = 'OTHER';

    var id = ipc_id_('CHG');
    ipc_charges_().appendRow([id, ip, new Date(), cat, acc_str_(payload.description), 'MANUAL', '',
      amt, acc_money_(payload.gst), ipc_isTrue_(payload.isNonPayable) ? 'TRUE' : 'FALSE', 'ON_TAB', '', acc_str_(payload.user) || 'UNKNOWN', acc_str_(payload.notes)]);
    acc_audit_(payload.user, 'IP_CHARGE_ADD', IPC_CFG.CHARGES, id, '', amt, ip + ' / ' + cat);
    SpreadsheetApp.flush();
    return { success: true, message: "Charge added · ₹" + amt, chargeId: id };
  } catch (e) { return { success: false, message: "Error: " + e.message }; }
  finally { lock.releaseLock(); }
}

// Hook for Lab/Pharmacy "Bill to IP". Idempotent on (source, sourceRef).
function billChargeToIp(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ip = acc_str_(payload.ipNumber).trim();
    var source = acc_str_(payload.source).toUpperCase();   // 'PHARMACY' | 'LAB'
    var ref = acc_str_(payload.sourceRef).trim();
    if (!ip || !ref) return { success: false, message: "IP number and source reference required." };
    var adm = ipc_admission_(ip);
    if (!adm) return { success: false, message: "Admission " + ip + " not found." };

    var sh = ipc_charges_(), existing = ipc_objs_(sh);
    for (var i = 0; i < existing.length; i++)
      if (acc_str_(existing[i]['Source']).toUpperCase() === source && acc_str_(existing[i]['Source_Ref']).trim() === ref)
        return { success: true, message: "Already on tab.", chargeId: acc_str_(existing[i]['Charge_ID']), duplicate: true };

    var id = ipc_id_('CHG'), cat = (source === 'LAB') ? 'LAB' : 'PHARMACY';
    sh.appendRow([id, ip, new Date(), cat, acc_str_(payload.description) || (cat + ' charge'), source, ref,
      acc_money_(payload.amount), acc_money_(payload.gst), 'FALSE', 'ON_TAB', '', acc_str_(payload.user) || 'SYSTEM', '']);
    acc_audit_(payload.user, 'IP_BILL_TO_TAB', source, ref, '', acc_money_(payload.amount), ip);
    SpreadsheetApp.flush();
    return { success: true, message: "Posted to IP tab.", chargeId: id };
  } catch (e) { return { success: false, message: "Error: " + e.message }; }
  finally { lock.releaseLock(); }
}

// When an IP patient pays a pharmacy/lab bill at the counter instead of on tab.
function markIpChargePaidAtCounter(source, sourceRef, payMode, user) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sh = ipc_charges_(), d = sh.getDataRange().getValues(), h = d[0].map(function (x) { return acc_str_(x).trim(); });
    var cSrc = h.indexOf('Source'), cRef = h.indexOf('Source_Ref'), cStat = h.indexOf('Status');
    for (var i = 1; i < d.length; i++) {
      if (acc_str_(d[i][cSrc]).toUpperCase() === acc_str_(source).toUpperCase() && acc_str_(d[i][cRef]).trim() === acc_str_(sourceRef).trim() && acc_str_(d[i][cStat]).toUpperCase() === 'ON_TAB') {
        sh.getRange(i + 1, cStat + 1).setValue('PAID_COUNTER');
        acc_audit_(user, 'IP_PAID_AT_COUNTER', source, sourceRef, 'ON_TAB', 'PAID_COUNTER', acc_str_(payMode));
        SpreadsheetApp.flush();
        return { success: true, message: "Charge marked paid at counter — removed from tab." };
      }
    }
    return { success: false, message: "No open tab charge found for that bill." };
  } catch (e) { return { success: false, message: "Error: " + e.message }; }
  finally { lock.releaseLock(); }
}

// ---- advances -------------------------------------------------------------
function collectIpAdvance(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ip = acc_str_(payload.ipNumber).trim();
    var adm = ipc_admission_(ip);
    if (!adm) return { success: false, message: "Admission " + ip + " not found." };
    if (adm.status !== 'ACTIVE') return { success: false, message: "Admission is " + adm.status + "." };
    var amt = acc_money_(payload.amount);
    if (amt <= 0) return { success: false, message: "Advance must be greater than 0." };
    var mode = acc_str_(payload.payMode) || 'Cash', now = new Date();

    var advId = ipc_id_('ADV');
    ipc_advances_().appendRow([advId, ip, now, amt, mode, acc_str_(payload.txnId), 'HELD', '', acc_str_(payload.user) || 'UNKNOWN']);
    // Real cash hits the drawer/bank now (liability, NOT revenue).
    acc_sheet_(ACC_CFG.LEDGER).appendRow([advId, now, 'IP_RECEIPT', 'IP_ADVANCE', ip, advId,
      'IP Advance · ' + ip, mode, amt, 0, acc_str_(payload.user) || 'UNKNOWN', '', 'FALSE', 'Security/Advance deposit']);
    acc_audit_(payload.user, 'IP_ADVANCE', IPC_CFG.ADVANCES, advId, '', amt, ip + ' / ' + mode);
    SpreadsheetApp.flush();
    return { success: true, message: "Advance ₹" + amt + " held for " + ip + ".", advanceId: advId };
  } catch (e) { return { success: false, message: "Error: " + e.message }; }
  finally { lock.releaseLock(); }
}

// ---- running tab (feeds the discharge simulator) --------------------------
function getRunningTab(ipNumber) {
  try {
    var ip = acc_str_(ipNumber).trim();
    var adm = ipc_admission_(ip);
    if (!adm) return { success: false, message: "Admission " + ip + " not found." };
    adm.doaFmt = (function () { var d = acc_toDate_(adm.doa); return d ? Utilities.formatDate(d, ACC_CFG.TZ, 'dd-MMM-yyyy') : acc_str_(adm.doa); })();
    adm.lengthOfStay = ipc_los_(adm.doa);

    var byCat = {}, grossTab = 0, nonPayable = 0, paidCounter = 0, lines = [], groups = {};
    ipc_objs_(ipc_charges_()).forEach(function (c) {
      if (acc_str_(c['IP_Number']).trim() !== ip) return;
      var st = acc_str_(c['Status']).toUpperCase(), amt = acc_money_(c['Amount']);
      if (st === 'VOID') return;
      var d = acc_toDate_(c['Timestamp']);
      var rec = { id: acc_str_(c['Charge_ID']), category: acc_str_(c['Category']), description: acc_str_(c['Description']), source: acc_str_(c['Source']), ref: acc_str_(c['Source_Ref']), amount: amt, gst: acc_money_(c['GST']), nonPayable: ipc_isTrue_(c['Is_NonPayable']), status: st, date: d ? Utilities.formatDate(d, ACC_CFG.TZ, 'dd-MMM') : '' };
      lines.push(rec);
      if (st === 'PAID_COUNTER') { paidCounter += amt; return; }
      if (st === 'ON_TAB') {
        grossTab += amt; byCat[rec.category] = (byCat[rec.category] || 0) + amt; if (rec.nonPayable) nonPayable += amt;
        var src = rec.source.toUpperCase();
        if (src === 'PHARMACY' || src === 'LAB') {
          if (!groups[src]) groups[src] = { dept: src, count: 0, total: 0, items: [] };
          groups[src].count++; groups[src].total += amt;
          groups[src].items.push({ ref: rec.ref, amount: amt, date: rec.date });
        }
      }
    });
    var autoGroups = Object.keys(groups).map(function (k) { var g = groups[k]; g.total = acc_money_(g.total); return g; });

    var advHeld = 0;
    ipc_objs_(ipc_advances_()).forEach(function (a) {
      if (acc_str_(a['IP_Number']).trim() === ip && acc_str_(a['Status']).toUpperCase() === 'HELD') advHeld += acc_money_(a['Amount']);
    });

    // existing draft (ward counts/remarks survive reloads)
    var draft = null;
    ipc_objs_(ipc_drafts_()).forEach(function (r) {
      if (acc_str_(r['IP_Number']).trim() !== ip) return;
      var ward = []; try { ward = JSON.parse(acc_str_(r['Ward_JSON']) || '[]'); } catch (e) { ward = []; }
      var insurers = []; try { insurers = JSON.parse(acc_str_(r['Insurers_JSON']) || '[]'); } catch (e) { insurers = []; }
      draft = { discountPercent: acc_money_(r['Discount_Percent']), packageCode: acc_str_(r['Package_Code']), packageCap: acc_money_(r['Package_Cap']), insuranceApproved: acc_money_(r['Insurance_Approved']), insurers: insurers, ward: ward, remarks: acc_str_(r['Remarks']) };
    });

    return {
      success: true, admission: adm,
      grossTab: acc_money_(grossTab), nonPayable: acc_money_(nonPayable),
      paidAtCounter: acc_money_(paidCounter), advancesHeld: acc_money_(advHeld),
      byCategory: byCat, autoGroups: autoGroups, lines: lines, draft: draft,
      outstanding: acc_money_(grossTab - advHeld)
    };
  } catch (e) { return { success: false, message: e.message }; }
}

// save (upsert) discharge draft — ward counts, rates, remarks; no charges posted yet
function saveDischargeDraft(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ip = acc_str_(payload.ipNumber).trim();
    if (!ip) return { success: false, message: "IP number required." };
    var sh = ipc_drafts_(), d = sh.getDataRange().getValues();
    var row = [ip, new Date(), acc_money_(payload.discountPercent), acc_str_(payload.packageCode), acc_money_(payload.packageCap), acc_money_(payload.insuranceApproved),
      JSON.stringify(payload.insurers || []), JSON.stringify(payload.ward || []), acc_str_(payload.remarks), acc_str_(payload.user) || 'UNKNOWN'];
    for (var i = 1; i < d.length; i++) if (acc_str_(d[i][0]).trim() === ip) { sh.getRange(i + 1, 1, 1, row.length).setValues([row]); SpreadsheetApp.flush(); return { success: true, message: "Draft saved." }; }
    sh.appendRow(row); SpreadsheetApp.flush();
    return { success: true, message: "Draft saved." };
  } catch (e) { return { success: false, message: "Error: " + e.message }; }
  finally { lock.releaseLock(); }
}

// clinical visit log (view + record) — feeds consultant/surgeon counts
function ipc_visits_() { return ipc_ensure_('IP_Visits', ['Visit_ID', 'IP_Number', 'Date', 'Doctor', 'Type', 'Count', 'Remark', 'By']); }
function getIpVisits(ipNumber) {
  try {
    var ip = acc_str_(ipNumber).trim(), out = [];
    ipc_objs_(ipc_visits_()).forEach(function (r) {
      if (acc_str_(r['IP_Number']).trim() !== ip) return;
      var dd = acc_toDate_(r['Date']);
      out.push({ date: dd ? Utilities.formatDate(dd, ACC_CFG.TZ, 'dd-MMM') : acc_str_(r['Date']), doctor: acc_str_(r['Doctor']), type: acc_str_(r['Type']), count: acc_money_(r['Count']), remark: acc_str_(r['Remark']) });
    });
    return { success: true, visits: out };
  } catch (e) { return { success: false, message: e.message }; }
}
function addIpVisit(p) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ip = acc_str_(p.ipNumber).trim(); if (!ip) return { success: false, message: 'IP number required.' };
    ipc_visits_().appendRow([ipc_id_('VIS'), ip, acc_str_(p.date) || Utilities.formatDate(new Date(), ACC_CFG.TZ, 'yyyy-MM-dd'), acc_str_(p.doctor), acc_str_(p.type), acc_money_(p.count) || 1, acc_str_(p.remark), acc_str_(p.user) || 'UNKNOWN']);
    SpreadsheetApp.flush();
    return { success: true, message: 'Visit recorded.' };
  } catch (e) { return { success: false, message: e.message }; }
  finally { lock.releaseLock(); }
}

// list of currently admitted patients (for the discharge picker)
function getOpenAdmissions() {
  try {
    var sh = ipc_ss_().getSheetByName(IPC_CFG.ADMISSIONS);
    if (!sh) return { success: true, admissions: [] };
    var d = sh.getDataRange().getValues(), out = [];
    for (var i = 1; i < d.length; i++)
      if (acc_str_(d[i][11]).toUpperCase() === 'ACTIVE')
        out.push({ ipNumber: acc_str_(d[i][0]), patientId: acc_str_(d[i][1]), patientName: acc_str_(d[i][2]), bed: acc_str_(d[i][8]), doa: acc_str_(d[i][4]), consultant: acc_str_(d[i][9]) });
    return { success: true, admissions: out };
  } catch (e) { return { success: false, message: e.message }; }
}

// ---- discharge: the grand settlement -------------------------------------
// payload = {ipNumber, packageCap, insuranceApproved, patientPayMode, reason, user}
// payload = {ipNumber, wardCharges:[{category,multiplier,rate,remark}], discountPercent,
//            packageCap, insuranceApproved, patientPayMode, reason, user}
// payload also accepts dsOverrideReason — supplied by the client after the
// discharge-summary gate returns DS_NOT_SIGNED. With DS_BILLING_GATE = OFF the
// gate is a no-op, so this settlement behaves exactly as it did before.
function settleDischarge(payload) {
  var lock = LockService.getScriptLock();
  try {
    var ip0 = acc_str_(payload.ipNumber).trim();

    // --- discharge summary gate (outside the lock: it only reads) ----------
    var dsGate = null;
    if (ip0 && typeof dsx_gateCheck_ === 'function') {
      dsGate = dsx_gateCheck_(ip0, payload.dsOverrideReason);
      if (!dsGate.allow) {
        return { success: false, code: dsGate.code, message: dsGate.message,
                 dsSummaryId: dsGate.summaryId, dsSummaryStatus: dsGate.status };
      }
    }

    lock.waitLock(10000);
    var ip = ip0;
    var adm = ipc_admission_(ip);
    if (!adm) return { success: false, message: "Admission " + ip + " not found." };
    if (adm.status !== 'ACTIVE') return { success: false, message: "Admission already " + adm.status + "." };

    var sh = ipc_charges_(), now = new Date(), user = acc_str_(payload.user) || 'UNKNOWN';

    // pre-existing ON_TAB (auto-fetched + prior manual) = always inside package scope
    var preGross = 0;
    ipc_objs_(sh).forEach(function (c) { if (acc_str_(c['IP_Number']).trim() === ip && acc_str_(c['Status']).toUpperCase() === 'ON_TAB') preGross += acc_money_(c['Amount']); });

    // 1. post ward charges (multiplier x rate) as MANUAL lines; track included vs exclusion
    var wardIncluded = 0, wardExclude = 0;
    (payload.wardCharges || []).forEach(function (w) {
      var mult = acc_money_(w.multiplier), rate = acc_money_(w.rate), amt = acc_money_(mult * rate);
      if (amt <= 0) return;
      var cat = acc_str_(w.category).toUpperCase(); if (IPC_CFG.CHARGE_CATS.indexOf(cat) === -1) cat = 'OTHER';
      var excl = (w.exclude === true);
      if (excl) wardExclude += amt; else wardIncluded += amt;
      sh.appendRow([ipc_id_('CHG'), ip, now, cat, acc_str_(w.category) + ' × ' + mult + ' @ ' + rate + (excl ? ' [OUTSIDE PKG]' : ''), 'MANUAL', '', amt, 0, 'FALSE', 'ON_TAB', '', user, acc_str_(w.remark)]);
    });
    SpreadsheetApp.flush();

    // 2. freeze: gather ALL ON_TAB rows
    var d = sh.getDataRange().getValues(), h = d[0].map(function (x) { return acc_str_(x).trim(); });
    var cIp = h.indexOf('IP_Number'), cAmt = h.indexOf('Amount'), cNP = h.indexOf('Is_NonPayable'), cStat = h.indexOf('Status'), cSet = h.indexOf('Settlement_ID'), cSrc = h.indexOf('Source'), cRef = h.indexOf('Source_Ref');
    var grossTab = 0, nonPayable = 0, rows = [], srcRefs = [];
    for (var i = 1; i < d.length; i++) {
      if (acc_str_(d[i][cIp]).trim() !== ip || acc_str_(d[i][cStat]).toUpperCase() !== 'ON_TAB') continue;
      var a = acc_money_(d[i][cAmt]); grossTab += a;
      if (ipc_isTrue_(d[i][cNP])) nonPayable += a;
      rows.push(i + 1);
      var sr = acc_str_(d[i][cRef]).trim(); if (sr) srcRefs.push({ source: acc_str_(d[i][cSrc]).toUpperCase(), ref: sr });
    }

    // 3. package cap applies to the INCLUDED scope; exclusions bill on top
    var cap = acc_money_(payload.packageCap);
    var includedGross = acc_money_(preGross + wardIncluded);
    var cappedIncluded = (cap > 0) ? Math.min(includedGross, cap) : includedGross;
    var pkgAdj = acc_money_(includedGross - cappedIncluded);
    var afterCap = acc_money_(cappedIncluded + wardExclude);
    var discPct = Math.max(0, Math.min(100, acc_money_(payload.discountPercent)));
    var discAmt = acc_money_(afterCap * discPct / 100);
    var billable = acc_money_(afterCap - discAmt);

    // advances HELD
    var ash = ipc_advances_(), ad = ash.getDataRange().getValues(), ah = ad[0].map(function (x) { return acc_str_(x).trim(); });
    var aIp = ah.indexOf('IP_Number'), aAmt = ah.indexOf('Amount'), aStat = ah.indexOf('Status'), aSet = ah.indexOf('Settlement_ID');
    var advTotal = 0, advRows = [];
    for (var j = 1; j < ad.length; j++) if (acc_str_(ad[j][aIp]).trim() === ip && acc_str_(ad[j][aStat]).toUpperCase() === 'HELD') { advTotal += acc_money_(ad[j][aAmt]); advRows.push(j + 1); }

    // 4. multi-insurer split (clamped to insurable portion)
    var insurers = (payload.insurers || []).filter(function (x) { return x && acc_money_(x.amount) > 0; })
      .map(function (x) { return { name: acc_str_(x.name), network: x.network === true, amount: acc_money_(x.amount) }; });
    var insurable = Math.max(0, billable - nonPayable);
    var insTotal = 0; insurers.forEach(function (x) { insTotal += x.amount; });
    if (insTotal > insurable) {  // clamp the last-mile so insurance never exceeds insurable
      var over = insTotal - insurable;
      for (var z = insurers.length - 1; z >= 0 && over > 0; z--) { var cut = Math.min(insurers[z].amount, over); insurers[z].amount = acc_money_(insurers[z].amount - cut); over -= cut; }
      insTotal = insurable;
    }
    var advApplied = Math.min(advTotal, billable);
    var raw = acc_money_(billable - advApplied - insTotal);
    var patientLiability = Math.max(0, raw), refundDue = Math.max(0, -raw), patientPaid = patientLiability;
    var payMode = acc_str_(payload.patientPayMode) || 'Cash';
    var recognizedRealized = acc_money_(advApplied + patientPaid);

    var setId = ipc_id_('STL');

    // 5. commit
    rows.forEach(function (r) { sh.getRange(r, cStat + 1).setValue('SETTLED'); sh.getRange(r, cSet + 1).setValue(setId); });
    srcRefs.forEach(function (s) { try { ipc_closeSourceBill_(s.source, s.ref, setId); } catch (e) {} });
    advRows.forEach(function (r) { ash.getRange(r, aStat + 1).setValue('APPLIED'); ash.getRange(r, aSet + 1).setValue(setId); });

    ipc_settlements_().appendRow([setId, ip, now, grossTab, acc_str_(payload.packageCode), cap, pkgAdj, cappedIncluded, discPct, discAmt, nonPayable, advApplied,
      insTotal, JSON.stringify(insurers), insTotal > 0 ? 'PENDING' : 'NONE', patientLiability, patientPaid, payMode, refundDue, recognizedRealized, user, acc_str_(payload.reason)]);

    // each insurer becomes a trackable claim on the Insurance Desk
    if (insurers.length && typeof ins_createClaimsForSettlement_ === 'function') { try { ins_createClaimsForSettlement_(setId, ip, adm, billable, insurers, user); } catch (e) {} }

    if (patientPaid > 0)
      acc_sheet_(ACC_CFG.LEDGER).appendRow([setId + '-C', now, 'IP_RECEIPT', 'IP_SETTLEMENT', ip, setId, 'IP Discharge · ' + ip, payMode, patientPaid, 0, user, '', 'FALSE', 'Patient net payable']);
    if (refundDue > 0)
      acc_sheet_(ACC_CFG.LEDGER).appendRow([setId + '-R', now, 'IP_RECEIPT', 'IP_REFUND', ip, setId, 'Advance refund · ' + ip, payMode, 0, refundDue, user, '', 'FALSE', 'Excess advance returned']);

    // 6. discharge admission + bed -> Cleaning, clear draft
    var adsh = ipc_ss_().getSheetByName(IPC_CFG.ADMISSIONS);
    adsh.getRange(adm.row, 12).setValue('DISCHARGED');
    adsh.getRange(adm.row, 13).setValue(Utilities.formatDate(now, ACC_CFG.TZ, 'yyyy-MM-dd'));
    var bsh = ipc_ss_().getSheetByName(IPC_CFG.BEDS);
    if (bsh) { var bd = bsh.getDataRange().getValues(); for (var b = 1; b < bd.length; b++) if (acc_str_(bd[b][6]).trim() === ip) { bsh.getRange(b + 1, 3).setValue('Cleaning'); bsh.getRange(b + 1, 4, 1, 4).setValues([['', '', '', '']]); break; } }
    var dsh = ipc_drafts_(), dd = dsh.getDataRange().getValues();
    for (var k = dd.length - 1; k >= 1; k--) if (acc_str_(dd[k][0]).trim() === ip) dsh.deleteRow(k + 1);

    acc_audit_(user, 'IP_DISCHARGE_SETTLE', IPC_CFG.SETTLEMENTS, setId, 'Gross: ' + grossTab, 'Net paid: ' + patientPaid,
      'Pkg: ' + acc_str_(payload.packageCode) + ' | Disc%: ' + discPct + ' | Ins: ' + insTotal + ' | Adv: ' + advApplied + (payload.reason ? ' | ' + payload.reason : ''));
    SpreadsheetApp.flush();

    if (dsGate) {
      if (dsGate.overridden && typeof dsx_logGateOverride_ === 'function') {
        dsx_logGateOverride_(ip, dsGate, user, 'settleDischarge');
      }
      if (typeof dsx_logDischargeCompleted_ === 'function') {
        dsx_logDischargeCompleted_(ip, user, 'settleDischarge');
      }
    }

    return {
      success: true, settlementId: setId,
      message: "Discharged. Patient paid ₹" + patientPaid + (insTotal > 0 ? ", insurance ₹" + insTotal + " pending" : "") + (refundDue > 0 ? ", refund ₹" + refundDue : "") + ".",
      summary: { grossTab: grossTab, includedGross: includedGross, exclusions: wardExclude, packageAdjustment: pkgAdj, discountAmount: discAmt, billable: billable, nonPayable: nonPayable, advanceApplied: advApplied, insuranceTotal: insTotal, insurers: insurers, patientPaid: patientPaid, refundDue: refundDue }
    };
  } catch (e) { return { success: false, message: "Error: " + e.message }; }
  finally { lock.releaseLock(); }
}