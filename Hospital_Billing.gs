// ============================================================================
// Hospital_Billing.gs  —  Crescentia HealthTech / CresRx
// OP / miscellaneous / package billing, wired into the rest of the hospital.
// ----------------------------------------------------------------------------
// WHAT THIS REPLACES
//   The old Hospital Billing screen was a stand-alone form. Its particulars
//   were eight hard-coded <option> tags whose *value* was the rate, it never
//   read a patient's other bills, it had no concept of a package, and its
//   Save button was:
//
//       function saveInvoiceToDB() { alert("...Ready to route ... in Phase 2."); }
//
//   So nothing it produced existed anywhere after the page was closed, and
//   Accounts could not see a rupee of it.
//
// HOW IT CONNECTS NOW
//   Services       Service_Master, a sheet the clinic edits — no code change
//                  to reprice a consultation.
//   Packages       Package_Master, the same sheet the IP discharge settlement
//                  already uses (AccountsIPPackages.gs). One package master,
//                  two desks.
//   Patients       the Patients sheet, through the same profile reader the
//                  rest of the app uses.
//   Appointments   an appointment's Fee column is the consultation TARIFF; the
//                  bill pre-fills from it and records which appointment it
//                  settled, so the two can never disagree again.
//   Admissions     an inpatient's IP number rides on the invoice, so an OP-side
//                  charge raised during a stay lands on the right admission.
//   Pharmacy / Lab their open balances for this patient are shown while the
//                  bill is being made, so the cashier sees the whole due.
//   Accounts       acc_hospitalRows_() in Accounts.gs merges these invoices
//                  into the virtual ledger and the receivables tab. Nothing is
//                  posted to Finance_Master_Ledger: per the model at the top of
//                  Accounts.gs, that sheet holds expenses and manual entries
//                  only, and income is read live from its source sheet. Posting
//                  here as well would double-count every rupee.
//
// CONVENTIONS
//   - Header-map reads and writes only. No positional access.
//   - Every entry point takes a session token and returns
//     {success, message, data} — never a bare value.
//   - Money is rounded once, at the boundary, by hb_money_().
// ============================================================================

var HB_CFG = {
  INVOICES: 'Hospital_Invoices',
  ITEMS:    'Hospital_Invoice_Items',
  SERVICES: 'Service_Master',
  PACKAGES: 'Package_Master',
  TZ:       'Asia/Kolkata',
  LOCK_MS:  15000,
  // Who may raise a bill, and who may only look at one.
  WRITERS:  ['admin', 'accounts', 'accountant', 'receptionist', 'reception'],
  READERS:  ['admin', 'accounts', 'accountant', 'receptionist', 'reception', 'doctor']
};

var HB_H_INVOICES = [
  'Invoice_No', 'Tenant_ID', 'Timestamp', 'Invoice_Date',
  'Patient_ID', 'Patient_Name', 'Age_Sex', 'Mobile', 'Address',
  'Appt_ID', 'IP_Number', 'Doctor_ID', 'Doctor_Name',
  'Bill_Type', 'Package_Code', 'Package_Name',
  'Gross', 'Discount', 'Tax', 'Net', 'Paid', 'Balance',
  'Payment_Mode', 'Payment_Ref', 'Payment_Status',
  'Created_By', 'Notes', 'Status', 'Cancel_Reason'
];

var HB_H_ITEMS = [
  'Item_Key', 'Invoice_No', 'Line_No', 'Service_Code', 'Particulars',
  'Category', 'Qty', 'Rate', 'Discount', 'Tax_Pct', 'Amount', 'Source'
];

var HB_H_SERVICES = ['Service_Code', 'Name', 'Category', 'Rate', 'Tax_Pct', 'Active'];

/** The starting tariff. Seeded once; the clinic owns the sheet afterwards. */
var HB_SEED_SERVICES = [
  ['CONS-GEN',  'General Consultation',            'CONSULTATION', 300,  0, 'YES'],
  ['CONS-SPEC', 'Specialist Consultation',         'CONSULTATION', 500,  0, 'YES'],
  ['CONS-REV',  'Review / Follow-up Consultation', 'CONSULTATION', 150,  0, 'YES'],
  ['CONS-EMER', 'Emergency Consultation',          'CONSULTATION', 800,  0, 'YES'],
  ['REG-NEW',   'New Registration',                'REGISTRATION',  50,  0, 'YES'],
  ['PROC-DRESS','Dressing',                        'PROCEDURE',    200,  0, 'YES'],
  ['PROC-INJ',  'Injection / IV administration',   'PROCEDURE',    100,  0, 'YES'],
  ['PROC-SUT',  'Suturing (minor)',                'PROCEDURE',    800,  0, 'YES'],
  ['PROC-NEB',  'Nebulisation',                    'PROCEDURE',    150,  0, 'YES'],
  ['PROC-CATH', 'Catheterisation',                 'PROCEDURE',    500,  0, 'YES'],
  ['PROC-ECG',  'ECG',                             'DIAGNOSTIC',   200,  0, 'YES'],
  ['IP-ADM',    'IP Admission Charges',            'INPATIENT',    500,  0, 'YES'],
  ['IP-ROOM',   'Room Rent (per day)',             'INPATIENT',   1500,  0, 'YES'],
  ['IP-ICU',    'ICU Monitoring (per day)',        'INPATIENT',   3000,  0, 'YES'],
  ['IP-NURS',   'Nursing Charges (per day)',       'INPATIENT',    400,  0, 'YES'],
  ['IP-RMO',    'RMO Charges (per day)',           'INPATIENT',    300,  0, 'YES'],
  ['IP-VISIT',  'Consultant Visit',                'INPATIENT',    500,  0, 'YES'],
  ['IP-SPVISIT','Specialist Visit',                'INPATIENT',    600,  0, 'YES'],
  ['OT-MINOR',  'Minor OT Charges',                'THEATRE',     3000,  0, 'YES'],
  ['OT-MAJOR',  'Major OT Charges',                'THEATRE',    10000,  0, 'YES'],
  ['AMB-LOCAL', 'Ambulance (local)',               'OTHER',       1000,  0, 'YES'],
  ['MISC-CERT', 'Medical Certificate',             'OTHER',        200,  0, 'YES'],
  ['MISC-REC',  'Records / Photocopy',             'OTHER',        100,  0, 'YES']
];

// ---------------------------------------------------------------------------
// SECTION A — scalars
// ---------------------------------------------------------------------------

function hb_str_(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }
function hb_upper_(v) { return hb_str_(v).toUpperCase(); }
function hb_money_(v) { var n = parseFloat(v); return isNaN(n) ? 0 : Math.round(n * 100) / 100; }
function hb_int_(v) { var n = parseInt(v, 10); return isNaN(n) ? 0 : n; }

function hb_tz_() {
  try { return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || HB_CFG.TZ; }
  catch (e) { return HB_CFG.TZ; }
}

function hb_tenant_() {
  try { return getTenantId_(); } catch (e) { return 'VALARMATHI'; }
}

/** Shared parser — see Date_Utils.gs for why `new Date(text)` was not enough. */
function hb_toDate_(v) {
  return cresc_toDate_(v);
}

function hb_fmt_(v, pattern) {
  var d = hb_toDate_(v);
  if (!d) return '';
  try { return Utilities.formatDate(d, hb_tz_(), pattern || 'dd-MMM-yyyy'); } catch (e) { return ''; }
}

function hb_dayKey_(v) { return hb_fmt_(v, 'yyyy-MM-dd'); }

function hb_ok_(message, data) {
  return { success: true, message: hb_str_(message), data: data === undefined ? null : data };
}

function hb_err_(message) { return { success: false, message: hb_str_(message) }; }

// ---------------------------------------------------------------------------
// SECTION B — access
// ---------------------------------------------------------------------------

/**
 * Validates the session and the role.
 *
 * @param {string} token
 * @param {boolean} needWrite
 * @return {{username:string, role:string}}
 * @throws when the session is dead or the role may not do this
 */
function hb_actor_(token, needWrite) {
  var sess = null;
  try { sess = dc_validateSession_(token); } catch (e) { sess = null; }
  if (!sess) throw new Error('Your session has expired. Please sign in again.');

  var role = hb_str_(sess.role).toLowerCase();
  var allowed = needWrite ? HB_CFG.WRITERS : HB_CFG.READERS;
  if (allowed.indexOf(role) === -1) {
    throw new Error(needWrite
      ? 'Your role (' + role + ') cannot raise or settle a hospital bill.'
      : 'Your role (' + role + ') cannot view hospital billing.');
  }
  return { username: hb_str_(sess.username), role: role };
}

// ---------------------------------------------------------------------------
// SECTION C — sheets
// ---------------------------------------------------------------------------

function hb_ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

/** Creates the sheet with its headers when absent; extends it when a column is new. */
function hb_sheet_(name, headers) {
  var ss = hb_ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8eaf6');
    return sh;
  }
  var width = Math.max(1, sh.getLastColumn());
  var head = sh.getRange(1, 1, 1, width).getValues()[0].map(hb_str_);
  var missing = headers.filter(function (h) { return head.indexOf(h) === -1; });
  if (missing.length) {
    sh.getRange(1, width + 1, 1, missing.length).setValues([missing]);
    sh.getRange(1, width + 1, 1, missing.length).setFontWeight('bold').setBackground('#e8eaf6');
  }
  return sh;
}

function hb_headerMap_(sheet) {
  var width = Math.max(1, sheet.getLastColumn());
  var head = sheet.getRange(1, 1, 1, width).getValues()[0];
  var map = {};
  for (var i = 0; i < head.length; i++) {
    var h = hb_str_(head[i]);
    if (h && map[h] === undefined) map[h] = i;
  }
  return map;
}

/** Every data row as an object keyed by header name, plus _row. */
function hb_readAll_(sheetName) {
  var sh = hb_ss_().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  var width = Math.max(1, sh.getLastColumn());
  var values = sh.getRange(1, 1, sh.getLastRow(), width).getValues();
  var head = values[0].map(hb_str_);
  var out = [];
  for (var i = 1; i < values.length; i++) {
    if (!hb_str_(values[i][0])) continue;
    var o = { _row: i + 1 };
    for (var j = 0; j < head.length; j++) if (head[j]) o[head[j]] = values[i][j];
    out.push(o);
  }
  return out;
}

/** Appends in header order. Unknown keys are ignored, missing ones blank. */
function hb_appendRow_(sheet, obj) {
  var map = hb_headerMap_(sheet);
  var width = Math.max(1, sheet.getLastColumn());
  var row = new Array(width);
  for (var i = 0; i < width; i++) row[i] = '';
  Object.keys(obj).forEach(function (h) {
    if (h.charAt(0) === '_') return;
    var col = map[h];
    if (col === undefined) return;
    row[col] = obj[h];
  });
  sheet.appendRow(row);
  return sheet.getLastRow();
}

// ---------------------------------------------------------------------------
// SECTION D — setup
// ---------------------------------------------------------------------------

/**
 * ADMIN, run once per deployment. Safe to run again: existing rows are never
 * touched, and Service_Master is only seeded when it is empty, so a clinic's
 * edited rates survive an upgrade.
 */
function setupHospitalBilling() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    hb_sheet_(HB_CFG.INVOICES, HB_H_INVOICES);
    hb_sheet_(HB_CFG.ITEMS, HB_H_ITEMS);

    var svc = hb_sheet_(HB_CFG.SERVICES, HB_H_SERVICES);
    var seeded = 0;
    if (svc.getLastRow() < 2) {
      svc.getRange(2, 1, HB_SEED_SERVICES.length, HB_SEED_SERVICES[0].length)
         .setValues(HB_SEED_SERVICES);
      seeded = HB_SEED_SERVICES.length;
    }

    // The package master belongs to AccountsIPPackages.gs; seed it here only
    // when it does not exist at all, so one desk is never waiting on the other.
    var pkgSeeded = 0;
    if (!hb_ss_().getSheetByName(HB_CFG.PACKAGES) && typeof setupPackages === 'function') {
      try { setupPackages(); pkgSeeded = 1; } catch (e) { /* advisory */ }
    }

    SpreadsheetApp.flush();
    return hb_ok_('Hospital billing is ready. ' +
                  (seeded ? seeded + ' services seeded. ' : 'Service_Master left as it is. ') +
                  (pkgSeeded ? 'Package_Master seeded.' : ''),
                  { servicesSeeded: seeded });
  } catch (e) {
    return hb_err_('Setup failed: ' + e.message);
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

// ---------------------------------------------------------------------------
// SECTION E — reads
// ---------------------------------------------------------------------------

/** The tariff, the packages, the doctors and the payment modes, in one call. */
function hb_getBootstrap(token) {
  try {
    var actor = hb_actor_(token, false);

    var services = hb_readAll_(HB_CFG.SERVICES)
      .filter(function (r) {
        var active = hb_upper_(r['Active']);
        return active !== 'NO' && active !== 'FALSE' && active !== '0';
      })
      .map(function (r) {
        return {
          code: hb_str_(r['Service_Code']),
          name: hb_str_(r['Name']),
          category: hb_upper_(r['Category']) || 'OTHER',
          rate: hb_money_(r['Rate']),
          taxPct: hb_money_(r['Tax_Pct'])
        };
      })
      .filter(function (r) { return r.code && r.name; });

    var packages = [];
    try {
      if (typeof getPackages === 'function') {
        var pk = getPackages();
        if (pk && pk.success) packages = pk.packages || [];
      }
    } catch (e) { packages = []; }

    // The same picker feed the booking modal uses, so the two screens never
    // disagree about who is taking clinics. Each doctor's default consult fee
    // rides along, because it is what an OP bill starts from.
    var doctors = [];
    try {
      if (typeof getActiveDoctors === 'function') {
        doctors = (getActiveDoctors() || []).map(function (d) {
          var full = (typeof dc_getDoctorById_ === 'function') ? dc_getDoctorById_(d.doctorId) : null;
          return {
            doctorId: hb_str_(d.doctorId),
            name: hb_str_(d.name),
            specialty: hb_str_(d.specialty),
            consultFee: full ? hb_money_(full.consultFee) : 0
          };
        });
      }
    } catch (e) { doctors = []; }

    return hb_ok_('', {
      services: services,
      packages: packages,
      doctors: doctors,
      paymentModes: ['Cash', 'UPI', 'Card', 'Bank Transfer', 'Credit'],
      canWrite: HB_CFG.WRITERS.indexOf(actor.role) !== -1,
      servicesConfigured: services.length > 0
    });
  } catch (e) {
    return hb_err_(e.message);
  }
}

/**
 * Everything the desk needs to know about one patient before it prices
 * anything: who they are, what they are here for, and what they already owe
 * elsewhere in the hospital.
 *
 * @param {string} query  a patient ID or a mobile number
 */
function hb_getPatientContext(token, query) {
  try {
    hb_actor_(token, false);
    var q = hb_upper_(query);
    if (!q) return hb_err_('Enter a patient ID or mobile number.');

    var patient = hb_findPatient_(q);
    if (!patient) return hb_err_('No patient found for "' + hb_str_(query) + '".');

    return hb_ok_('', {
      patient: patient,
      appointments: hb_recentAppointments_(patient.id),
      admission: hb_liveAdmission_(patient.id),
      outstanding: hb_outstanding_(patient.id),
      invoices: hb_invoicesFor_(patient.id, 10)
    });
  } catch (e) {
    return hb_err_(e.message);
  }
}

/** By patient ID first, then by mobile. */
function hb_findPatient_(q) {
  var sh = hb_ss_().getSheetByName('Patients');
  if (!sh || sh.getLastRow() < 2) return null;
  var width = Math.max(1, sh.getLastColumn());
  var values = sh.getRange(1, 1, sh.getLastRow(), width).getValues();
  var head = values[0].map(hb_str_);
  var idx = {};
  head.forEach(function (h, i) { if (h) idx[h] = i; });

  var col = function (row, name, fallback) {
    var i = (idx[name] === undefined) ? fallback : idx[name];
    return (i === undefined || i < 0) ? '' : hb_str_(row[i]);
  };

  for (var pass = 0; pass < 2; pass++) {
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var id = hb_upper_(col(row, 'Patient_ID', 0));
      if (!id) continue;
      var mobile = col(row, 'Mobile', 6).replace(/\D/g, '');
      var hit = (pass === 0) ? (id === q) : (mobile && mobile === q.replace(/\D/g, ''));
      if (!hit) continue;
      return {
        id: id,
        name: col(row, 'Name', 2),
        age: col(row, 'Age', 3),
        gender: col(row, 'Gender', 4),
        mobile: col(row, 'Mobile', 6),
        address: col(row, 'Address', 8)
      };
    }
  }
  return null;
}

/**
 * This patient's recent appointments, newest first, with the consultation
 * tariff the booking snapshotted and whether a bill already settled it.
 */
function hb_recentAppointments_(patientId) {
  var sh = hb_ss_().getSheetByName('Appointments');
  if (!sh || sh.getLastRow() < 2) return [];
  var width = Math.max(1, sh.getLastColumn());
  var values = sh.getRange(1, 1, sh.getLastRow(), width).getValues();
  var head = values[0].map(hb_str_);
  var idx = {};
  head.forEach(function (h, i) { if (h) idx[h] = i; });
  var docIdx = (idx['Doctor_ID'] === undefined) ? -1 : idx['Doctor_ID'];
  var snapIdx = (idx['Doctor_Name_Snapshot'] === undefined) ? -1 : idx['Doctor_Name_Snapshot'];

  var billed = hb_billedApptIds_();
  var want = hb_upper_(patientId);
  var out = [];

  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (hb_upper_(r[1]) !== want) continue;
    var status = hb_str_(r[6]);
    if (status === 'Blocked' || status === 'DELETE' || status === 'Cancelled') continue;
    var when = hb_toDate_(r[3]);
    out.push({
      apptId: hb_str_(r[0]),
      date: hb_dayKey_(r[3]),
      dateText: hb_fmt_(r[3], 'dd-MMM-yyyy'),
      sortMs: when ? when.getTime() : 0,
      time: hb_str_(r[4]),
      purpose: hb_str_(r[5]),
      status: status,
      tariff: hb_money_(r[7]),
      doctorId: (docIdx === -1) ? '' : hb_str_(r[docIdx]),
      doctorName: (snapIdx === -1) ? '' : hb_str_(r[snapIdx]),
      billed: !!billed[hb_upper_(r[0])]
    });
  }
  out.sort(function (a, b) { return b.sortMs - a.sortMs; });
  return out.slice(0, 10);
}

/** {APPT_ID: true} for every appointment a live invoice already covers. */
function hb_billedApptIds_() {
  var out = {};
  hb_readAll_(HB_CFG.INVOICES).forEach(function (r) {
    if (hb_upper_(r['Status']) === 'CANCELLED') return;
    var id = hb_upper_(r['Appt_ID']);
    if (id) out[id] = true;
  });
  return out;
}

function hb_liveAdmission_(patientId) {
  try {
    if (typeof IPA_CFG === 'undefined') return null;
    var sh = hb_ss_().getSheetByName(IPA_CFG.SHEET);
    if (!sh || sh.getLastRow() < 2) return null;
    var width = Math.max(1, sh.getLastColumn());
    var values = sh.getRange(1, 1, sh.getLastRow(), width).getValues();
    var head = values[0].map(hb_str_);
    var idx = {};
    head.forEach(function (h, i) { if (h) idx[h] = i; });
    var want = hb_upper_(patientId);

    for (var i = values.length - 1; i >= 1; i--) {
      var r = values[i];
      if (hb_upper_(r[idx['Patient_ID']]) !== want) continue;
      var status = hb_upper_(r[idx['Status']]);
      if (status !== 'ACTIVE' && status !== 'ADMITTED') continue;
      return {
        ipNumber: hb_str_(r[idx['IP_Number']]),
        ward: hb_str_(r[idx['Ward_Bed']]) || hb_str_(r[idx['Bed']]),
        consultant: hb_str_(r[idx['Consultant']]),
        doa: hb_fmt_(r[idx['DOA']], 'dd-MMM-yyyy')
      };
    }
  } catch (e) { /* the ward module is optional to this desk */ }
  return null;
}

/**
 * What this patient owes elsewhere. Read through the Accounts normalizers so
 * the figures on this screen are the same figures the receivables tab shows —
 * one definition of "open", not two.
 */
function hb_outstanding_(patientId) {
  var out = { pharmacy: 0, lab: 0, hospital: 0, total: 0, bills: [] };
  var want = hb_upper_(patientId);

  var collect = function (rows, label) {
    (rows || []).forEach(function (r) {
      if (hb_upper_(r.patientId) !== want) return;
      if (!r.open || !(r.balance > 0)) return;
      out[label] = hb_money_(out[label] + r.balance);
      out.bills.push({ source: r.source, billId: r.billId, amount: hb_money_(r.balance),
                       date: hb_fmt_(r.billDate, 'dd-MMM-yyyy') });
    });
  };

  try { if (typeof acc_pharmaRows_ === 'function') collect(acc_pharmaRows_(), 'pharmacy'); } catch (e) {}
  try { if (typeof acc_labRows_ === 'function') collect(acc_labRows_(), 'lab'); } catch (e) {}
  try { collect(acc_hospitalRows_(), 'hospital'); } catch (e) {}

  out.total = hb_money_(out.pharmacy + out.lab + out.hospital);
  return out;
}

function hb_invoicesFor_(patientId, limit) {
  var want = hb_upper_(patientId);
  return hb_readAll_(HB_CFG.INVOICES)
    .filter(function (r) { return hb_upper_(r['Patient_ID']) === want; })
    .map(hb_invoiceSummary_)
    .sort(function (a, b) { return b.sortMs - a.sortMs; })
    .slice(0, limit || 10);
}

function hb_invoiceSummary_(r) {
  var ts = hb_toDate_(r['Timestamp']);
  var stamped = r['Invoice_Date'] || r['Timestamp'];
  return {
    invoiceNo: hb_str_(r['Invoice_No']),
    date: hb_fmt_(stamped, 'dd-MMM-yyyy'),
    // Display text is no use for a range filter, so the sortable key travels
    // with it rather than being re-parsed out of "12-Sep-2026" later.
    day: hb_dayKey_(stamped),
    sortMs: ts ? ts.getTime() : 0,
    patientId: hb_upper_(r['Patient_ID']),
    patientName: hb_str_(r['Patient_Name']),
    apptId: hb_str_(r['Appt_ID']),
    ipNumber: hb_str_(r['IP_Number']),
    doctorName: hb_str_(r['Doctor_Name']),
    billType: hb_upper_(r['Bill_Type']) || 'OP',
    packageName: hb_str_(r['Package_Name']),
    gross: hb_money_(r['Gross']),
    discount: hb_money_(r['Discount']),
    tax: hb_money_(r['Tax']),
    net: hb_money_(r['Net']),
    paid: hb_money_(r['Paid']),
    balance: hb_money_(r['Balance']),
    mode: hb_str_(r['Payment_Mode']),
    payStatus: hb_upper_(r['Payment_Status']) || 'UNPAID',
    status: hb_upper_(r['Status']) || 'ACTIVE',
    createdBy: hb_str_(r['Created_By'])
  };
}

/**
 * The billing ledger.
 * @param {Object} filter {from, to, payStatus, search, limit}
 */
function hb_getInvoices(token, filter) {
  try {
    hb_actor_(token, false);
    filter = filter || {};

    var from = hb_str_(filter.from);
    var to = hb_str_(filter.to);
    var status = hb_upper_(filter.payStatus);
    var q = hb_upper_(filter.search);

    var rows = hb_readAll_(HB_CFG.INVOICES).map(hb_invoiceSummary_);

    var out = rows.filter(function (r) {
      var day = r.day || '';
      if (from && day && day < from) return false;
      if (to && day && day > to) return false;
      if (status && status !== 'ALL' && r.payStatus !== status) return false;
      if (q && (r.invoiceNo + ' ' + r.patientId + ' ' + r.patientName + ' ' + r.ipNumber)
                 .toUpperCase().indexOf(q) === -1) return false;
      return true;
    });

    out.sort(function (a, b) { return b.sortMs - a.sortMs; });
    var totals = out.reduce(function (t, r) {
      if (r.status === 'CANCELLED') return t;
      t.net = hb_money_(t.net + r.net);
      t.paid = hb_money_(t.paid + r.paid);
      t.balance = hb_money_(t.balance + r.balance);
      return t;
    }, { net: 0, paid: 0, balance: 0 });

    var limit = hb_int_(filter.limit) || 200;
    return hb_ok_('', { rows: out.slice(0, limit), totals: totals, count: out.length });
  } catch (e) {
    return hb_err_(e.message);
  }
}

/** One invoice with its lines, for reprinting. */
function hb_getInvoice(token, invoiceNo) {
  try {
    hb_actor_(token, false);
    var want = hb_upper_(invoiceNo);
    var hit = hb_readAll_(HB_CFG.INVOICES).filter(function (r) {
      return hb_upper_(r['Invoice_No']) === want;
    })[0];
    if (!hit) return hb_err_('Invoice ' + hb_str_(invoiceNo) + ' was not found.');

    var items = hb_readAll_(HB_CFG.ITEMS)
      .filter(function (r) { return hb_upper_(r['Invoice_No']) === want; })
      .map(function (r) {
        return {
          lineNo: hb_int_(r['Line_No']),
          serviceCode: hb_str_(r['Service_Code']),
          particulars: hb_str_(r['Particulars']),
          category: hb_upper_(r['Category']),
          qty: hb_money_(r['Qty']),
          rate: hb_money_(r['Rate']),
          discount: hb_money_(r['Discount']),
          taxPct: hb_money_(r['Tax_Pct']),
          amount: hb_money_(r['Amount']),
          source: hb_upper_(r['Source'])
        };
      })
      .sort(function (a, b) { return a.lineNo - b.lineNo; });

    var summary = hb_invoiceSummary_(hit);
    summary.ageSex = hb_str_(hit['Age_Sex']);
    summary.mobile = hb_str_(hit['Mobile']);
    summary.address = hb_str_(hit['Address']);
    summary.notes = hb_str_(hit['Notes']);
    summary.paymentRef = hb_str_(hit['Payment_Ref']);

    return hb_ok_('', { invoice: summary, items: items });
  } catch (e) {
    return hb_err_(e.message);
  }
}

// ---------------------------------------------------------------------------
// SECTION F — writes
// ---------------------------------------------------------------------------

/** HB-yyMM-#### — monthly sequence, derived from the sheet under the lock. */
function hb_nextInvoiceNo_() {
  var prefix = 'HB-' + Utilities.formatDate(new Date(), hb_tz_(), 'yyMM') + '-';
  var max = 0;
  hb_readAll_(HB_CFG.INVOICES).forEach(function (r) {
    var id = hb_upper_(r['Invoice_No']);
    if (id.indexOf(prefix) !== 0) return;
    var n = hb_int_(id.substring(prefix.length));
    if (n > max) max = n;
  });
  return prefix + String(max + 1).padStart(4, '0');
}

/**
 * Totals an invoice, server-side.
 *
 * The client computes the same figures to draw the screen, but the numbers
 * that are stored are these ones: a bill whose total is whatever the browser
 * said it was is not a bill.
 */
function hb_total_(items, billDiscount) {
  var gross = 0, lineDiscount = 0, tax = 0;
  var lines = (items || []).map(function (it, i) {
    var qty = hb_money_(it.qty) || 1;
    var rate = hb_money_(it.rate);
    var disc = hb_money_(it.discount);
    var taxPct = hb_money_(it.taxPct);
    var base = hb_money_(qty * rate);
    var afterDisc = Math.max(0, hb_money_(base - disc));
    var lineTax = hb_money_(afterDisc * taxPct / 100);

    gross = hb_money_(gross + base);
    lineDiscount = hb_money_(lineDiscount + Math.min(disc, base));
    tax = hb_money_(tax + lineTax);

    return {
      lineNo: i + 1,
      serviceCode: hb_str_(it.serviceCode),
      particulars: hb_str_(it.particulars),
      category: hb_upper_(it.category) || 'OTHER',
      qty: qty, rate: rate, discount: Math.min(disc, base), taxPct: taxPct,
      amount: hb_money_(afterDisc + lineTax),
      source: hb_upper_(it.source) || 'MANUAL'
    };
  });

  var billDisc = Math.max(0, hb_money_(billDiscount));
  var subtotal = hb_money_(gross - lineDiscount);
  if (billDisc > subtotal) billDisc = subtotal;

  var discount = hb_money_(lineDiscount + billDisc);
  var net = hb_money_(Math.max(0, gross - discount + tax));

  return { lines: lines, gross: gross, discount: discount, tax: tax, net: net };
}

/**
 * Raises an invoice.
 *
 * @param {Object} payload {
 *     patientId, patientName, ageSex, mobile, address,
 *     apptId, ipNumber, doctorId, doctorName,
 *     billType 'OP'|'PACKAGE'|'IP'|'MISC',
 *     packageCode, packageName,
 *     items [{serviceCode, particulars, category, qty, rate, discount, taxPct, source}],
 *     billDiscount, paid, paymentMode, paymentRef, notes }
 */
function hb_saveInvoice(token, payload) {
  var lock = LockService.getScriptLock();
  try {
    var actor = hb_actor_(token, true);
    payload = payload || {};

    var patientId = hb_upper_(payload.patientId);
    var patientName = hb_str_(payload.patientName);
    if (!patientName) return hb_err_('A patient name is required.');

    var items = payload.items || [];
    if (!items.length) return hb_err_('Add at least one particular before saving.');

    var totals = hb_total_(items, payload.billDiscount);
    if (totals.net <= 0) return hb_err_('The net amount is zero. Price at least one line.');

    var mode = hb_str_(payload.paymentMode) || 'Cash';
    var paid = hb_money_(payload.paid);
    if (paid < 0) return hb_err_('Amount paid cannot be negative.');
    if (paid > totals.net) paid = totals.net;
    var balance = hb_money_(totals.net - paid);

    // A credit bill is one that is not fully paid; naming it that way here
    // means the receivables tab picks it up without a second flag to keep
    // in step.
    var payStatus = balance <= 0 ? 'PAID' : (paid > 0 ? 'PART_PAID' : 'UNPAID');

    lock.waitLock(HB_CFG.LOCK_MS);

    var invSheet = hb_sheet_(HB_CFG.INVOICES, HB_H_INVOICES);
    var itemSheet = hb_sheet_(HB_CFG.ITEMS, HB_H_ITEMS);

    // One invoice per appointment. Without this, "Save" pressed twice — or a
    // slow first click — bills the same consultation to the same patient
    // again, and the second bill is indistinguishable from a real one.
    var apptId = hb_upper_(payload.apptId);
    if (apptId && hb_billedApptIds_()[apptId]) {
      return hb_err_('Appointment ' + apptId + ' has already been billed. ' +
                     'Open that invoice from the ledger instead of raising a second one.');
    }

    var now = new Date();
    var invoiceNo = hb_nextInvoiceNo_();

    hb_appendRow_(invSheet, {
      Invoice_No: invoiceNo,
      Tenant_ID: hb_tenant_(),
      Timestamp: now,
      Invoice_Date: hb_dayKey_(now),
      Patient_ID: patientId,
      Patient_Name: patientName,
      Age_Sex: hb_str_(payload.ageSex),
      Mobile: hb_str_(payload.mobile),
      Address: hb_str_(payload.address),
      Appt_ID: apptId,
      IP_Number: hb_upper_(payload.ipNumber),
      Doctor_ID: hb_str_(payload.doctorId),
      Doctor_Name: hb_str_(payload.doctorName),
      Bill_Type: hb_upper_(payload.billType) || 'OP',
      Package_Code: hb_upper_(payload.packageCode),
      Package_Name: hb_str_(payload.packageName),
      Gross: totals.gross,
      Discount: totals.discount,
      Tax: totals.tax,
      Net: totals.net,
      Paid: paid,
      Balance: balance,
      Payment_Mode: mode,
      Payment_Ref: hb_str_(payload.paymentRef),
      Payment_Status: payStatus,
      Created_By: actor.username,
      Notes: hb_str_(payload.notes),
      Status: 'ACTIVE',
      Cancel_Reason: ''
    });

    totals.lines.forEach(function (l) {
      hb_appendRow_(itemSheet, {
        Item_Key: invoiceNo + ':' + l.lineNo,
        Invoice_No: invoiceNo,
        Line_No: l.lineNo,
        Service_Code: l.serviceCode,
        Particulars: l.particulars,
        Category: l.category,
        Qty: l.qty,
        Rate: l.rate,
        Discount: l.discount,
        Tax_Pct: l.taxPct,
        Amount: l.amount,
        Source: l.source
      });
    });

    hb_audit_(actor, 'RAISE_INVOICE', invoiceNo,
              '', totals.net, 'Paid ' + paid + ' by ' + mode + '; balance ' + balance);
    hb_invalidateDashboard_();
    SpreadsheetApp.flush();

    return hb_ok_('Invoice ' + invoiceNo + ' saved.', {
      invoiceNo: invoiceNo,
      gross: totals.gross, discount: totals.discount, tax: totals.tax,
      net: totals.net, paid: paid, balance: balance, payStatus: payStatus
    });
  } catch (e) {
    return hb_err_(e.message);
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/** Takes money against an outstanding invoice. */
function hb_recordPayment(token, invoiceNo, amount, mode, ref) {
  var lock = LockService.getScriptLock();
  try {
    var actor = hb_actor_(token, true);
    var want = hb_upper_(invoiceNo);
    if (!want) return hb_err_('No invoice was named.');
    var pay = hb_money_(amount);
    if (pay <= 0) return hb_err_('Enter an amount greater than zero.');

    lock.waitLock(HB_CFG.LOCK_MS);

    var sh = hb_sheet_(HB_CFG.INVOICES, HB_H_INVOICES);
    var map = hb_headerMap_(sh);
    var rows = hb_readAll_(HB_CFG.INVOICES);
    var hit = rows.filter(function (r) { return hb_upper_(r['Invoice_No']) === want; })[0];
    if (!hit) return hb_err_('Invoice ' + want + ' was not found.');
    if (hb_upper_(hit['Status']) === 'CANCELLED') return hb_err_('Invoice ' + want + ' is cancelled.');

    var net = hb_money_(hit['Net']);
    var paid = hb_money_(hit['Paid']);
    var balance = hb_money_(net - paid);
    if (balance <= 0) return hb_err_('Invoice ' + want + ' is already settled in full.');
    if (pay > balance) pay = balance;

    var newPaid = hb_money_(paid + pay);
    var newBalance = hb_money_(net - newPaid);
    var payStatus = newBalance <= 0 ? 'PAID' : 'PART_PAID';

    var set = function (header, value) {
      if (map[header] === undefined) return;
      sh.getRange(hit._row, map[header] + 1).setValue(value);
    };
    set('Paid', newPaid);
    set('Balance', newBalance);
    set('Payment_Status', payStatus);
    set('Payment_Mode', hb_str_(mode) || hb_str_(hit['Payment_Mode']) || 'Cash');
    if (hb_str_(ref)) set('Payment_Ref', hb_str_(ref));

    hb_audit_(actor, 'RECORD_PAYMENT', want, balance, newBalance,
              'Received ' + pay + ' by ' + (hb_str_(mode) || 'Cash'));
    hb_invalidateDashboard_();
    SpreadsheetApp.flush();

    return hb_ok_('Received ' + pay + ' against ' + want + '.',
                  { invoiceNo: want, paid: newPaid, balance: newBalance, payStatus: payStatus });
  } catch (e) {
    return hb_err_(e.message);
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/**
 * Cancels an invoice. The row stays — a bill that was printed and handed over
 * cannot be made never to have existed — but it stops counting anywhere.
 */
function hb_cancelInvoice(token, invoiceNo, reason) {
  var lock = LockService.getScriptLock();
  try {
    var actor = hb_actor_(token, true);
    if (actor.role !== 'admin' && actor.role !== 'accounts' && actor.role !== 'accountant') {
      return hb_err_('Only an administrator or accounts may cancel an invoice.');
    }
    var why = hb_str_(reason);
    if (!why) return hb_err_('A reason is required to cancel an invoice.');

    lock.waitLock(HB_CFG.LOCK_MS);
    var sh = hb_sheet_(HB_CFG.INVOICES, HB_H_INVOICES);
    var map = hb_headerMap_(sh);
    var want = hb_upper_(invoiceNo);
    var hit = hb_readAll_(HB_CFG.INVOICES).filter(function (r) {
      return hb_upper_(r['Invoice_No']) === want;
    })[0];
    if (!hit) return hb_err_('Invoice ' + want + ' was not found.');
    if (hb_upper_(hit['Status']) === 'CANCELLED') return hb_err_('Invoice ' + want + ' is already cancelled.');

    if (map['Status'] !== undefined) sh.getRange(hit._row, map['Status'] + 1).setValue('CANCELLED');
    if (map['Cancel_Reason'] !== undefined) sh.getRange(hit._row, map['Cancel_Reason'] + 1).setValue(why);

    hb_audit_(actor, 'CANCEL_INVOICE', want, hb_money_(hit['Net']), 0, why);
    hb_invalidateDashboard_();
    SpreadsheetApp.flush();
    return hb_ok_('Invoice ' + want + ' cancelled.', { invoiceNo: want });
  } catch (e) {
    return hb_err_(e.message);
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

// ---------------------------------------------------------------------------
// SECTION G — plumbing
// ---------------------------------------------------------------------------

/** Into the accounts audit trail, because this is money. Never throws upward. */
function hb_audit_(actor, action, refId, oldVal, newVal, reason) {
  try {
    if (typeof acc_audit_ === 'function') {
      acc_audit_(actor.username, action, 'Hospital_Billing', refId, oldVal, newVal, reason);
    }
  } catch (e) { /* an audit line must not lose a bill */ }
}

function hb_invalidateDashboard_() {
  try { if (typeof invalidateDashboardCache === 'function') invalidateDashboardCache(); } catch (e) {}
}

/**
 * The Accounts normalizer for hospital invoices.
 *
 * Lives here rather than in Accounts.gs so the schema and its reader sit
 * together, and is called from acc_* exactly like the pharmacy and lab ones.
 */
function acc_hospitalRows_() {
  return hb_readAll_(HB_CFG.INVOICES).map(function (r) {
    var cancelled = hb_upper_(r['Status']) === 'CANCELLED';
    var net = hb_money_(r['Net']);
    var paid = hb_money_(r['Paid']);
    var balance = hb_money_(r['Balance']);
    var d = hb_toDate_(r['Timestamp']) || hb_toDate_(r['Invoice_Date']) || new Date();
    return {
      source: 'Hospital',
      billId: hb_str_(r['Invoice_No']),
      patientId: hb_upper_(r['Patient_ID']),
      name: hb_str_(r['Patient_Name']),
      admissionId: hb_str_(r['IP_Number']),
      net: cancelled ? 0 : net,
      // Only money actually taken is income; the rest is a receivable.
      realizedAmount: cancelled ? 0 : paid,
      balance: cancelled ? 0 : balance,
      mode: hb_str_(r['Payment_Mode']) || 'Cash',
      billDate: d,
      realizedDate: d,
      realized: !cancelled && paid > 0,
      open: !cancelled && balance > 0,
      _row: r._row
    };
  });
}