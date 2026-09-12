// ============================================================================
// Barcode_Engine.gs  —  Crescentia HealthTech / CresRx
// BARCODE PHASE 1 : collision-safe IDs, scan resolution, scan-to-check-in
// ----------------------------------------------------------------------------
// ADD AS A NEW FILE. Nothing here overwrites an existing function.
//
// DEPENDS ON (must already be deployed):
//   Doctor_Core.gs          dc_str_, dc_upper_, dc_dateKey_, dc_to12_, dc_headerMap_,
//                           dc_inScope_, resolveScope_, DC_DEFAULT_DOCTOR,
//                           dc_getDoctorById_
//   Doctor_Session_Store.gs dc_validateSession_
//   Doctors_Engine.gs       getTenantId_, logAudit_
//   IP_Admissions_Logic.gs  IPA_CFG, IPA_COL, ipa_isLive_, ipa_resolveWardBed_
//
// SECURITY MODEL
//   * Every entry point validates the session token and fails CLOSED.
//   * resolveScan() is READ-ONLY and takes no script lock (lock contention was
//     the measured cause of rising latency on rapid scans).
//   * The actions a role may take are decided HERE, not in the browser.
//   * Every scan resolved here is audited BY THE SERVER, inside resolveScan().
//     The browser is never trusted to report its own scans.
//
// ID SEQUENCES
//   bc_nextPatientId_ / bc_nextDailyId_ MUST be called while the caller holds
//   LockService.getScriptLock(). All current callers (registerPatient,
//   createLabRequest, collectLabSample) already do.
// ============================================================================

var BC_CFG = {
  LOCK_MS: 10000,
  PROP_PATIENT_SEQ: 'BC_SEQ_PATIENT',
  PROP_PATIENT_PREFIX: 'PATIENT_ID_PREFIX',     // optional ScriptProperty
  DEFAULT_PATIENT_PREFIX: 'LMTVS',
  PATIENT_PAD: 4,
  MAX_PAYLOAD: 64,
  PID_ROW_CACHE_S: 21600,
  SCAN_MAX_GAP_MS: 60,                          // keyboard-wedge detector
  SCAN_MIN_LEN: 6,
  // Server-side action matrix. Mirrors applyRBAC() nav visibility so a role
  // is never offered a screen its sidebar hides.
  // IP_CASESHEET and IP_NOTE are offered only when the scan resolves to a LIVE
  // admission — see bc_scanActions_(). Scanning the wristband at the bedside
  // is the fastest route into that patient's chart, and it was the one thing
  // the scanner could not do: it could open an OP consult for an inpatient,
  // but not their case sheet.
  // NEW_APPOINTMENT is offered to everyone who may write the appointment
  // ledger. The front desk's commonest job is "scan the card, book the next
  // visit", and until now the scan panel could check a patient in but not
  // book them — the card was scanned, then the ID was typed again into the
  // booking modal.
  ROLE_ACTIONS: {
    'admin':          ['CHECK_IN', 'NEW_APPOINTMENT', 'START_CONSULT', 'TIMELINE', 'PHARMACY_BILL', 'LAB_ORDERS', 'LAB_WALKIN', 'PRINT_CARD', 'IP_CASESHEET', 'IP_NOTE'],
    'doctor':         ['CHECK_IN', 'NEW_APPOINTMENT', 'START_CONSULT', 'TIMELINE', 'PHARMACY_BILL', 'LAB_ORDERS', 'LAB_WALKIN', 'PRINT_CARD', 'IP_CASESHEET', 'IP_NOTE'],
    'receptionist':   ['CHECK_IN', 'NEW_APPOINTMENT', 'PRINT_CARD'],
    'reception':      ['CHECK_IN', 'NEW_APPOINTMENT', 'PRINT_CARD'],
    'nurse':          ['TIMELINE', 'IP_NOTE'],
    'pharmacy':       ['PHARMACY_BILL'],
    'pharmacist':     ['PHARMACY_BILL'],
    'lab':            ['LAB_ORDERS', 'LAB_WALKIN', 'PRINT_CARD'],
    'lab technician': ['LAB_ORDERS', 'LAB_WALKIN', 'PRINT_CARD'],
    'accounts':       [],
    'accountant':     []
  },
  STATUS_WRITERS: ['admin', 'doctor', 'receptionist', 'reception'],
  TRANSITIONS: { 'Booked': ['Arrived', 'In-Progress'], 'Arrived': ['In-Progress'] }
};

// ============================================================================
// SECTION A — ID SEQUENCES (call only while holding the script lock)
// ============================================================================

function bc_patientPrefix_() {
  var p = '';
  try { p = PropertiesService.getScriptProperties().getProperty(BC_CFG.PROP_PATIENT_PREFIX) || ''; } catch (e) {}
  p = dc_upper_(p).replace(/[^A-Z]/g, '');
  return p || BC_CFG.DEFAULT_PATIENT_PREFIX;
}

function bc_escapeRegex_(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

/**
 * Next patient ID. Never reuses an ID, even after rows are deleted:
 *   n = max(highest ID ever issued [ScriptProperty], highest ID on the sheet) + 1
 * The sheet scan covers IDs issued before this engine existed; the property
 * covers IDs whose rows were later deleted.
 */
function bc_nextPatientId_(sheet) {
  var props = PropertiesService.getScriptProperties();
  var prefix = bc_patientPrefix_();
  var re = new RegExp('^' + bc_escapeRegex_(prefix) + '(\\d+)$');

  var seen = {};
  var maxOnSheet = 0;
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      var id = dc_upper_(ids[i][0]);
      if (!id) continue;
      seen[id] = true;
      var m = id.match(re);
      if (m) { var n0 = parseInt(m[1], 10); if (n0 > maxOnSheet) maxOnSheet = n0; }
    }
  }

  var stored = parseInt(props.getProperty(BC_CFG.PROP_PATIENT_SEQ) || '0', 10);
  if (isNaN(stored)) stored = 0;

  var n = Math.max(stored, maxOnSheet) + 1;
  var candidate = prefix + String(n).padStart(BC_CFG.PATIENT_PAD, '0');
  while (seen[candidate]) {
    n++;
    candidate = prefix + String(n).padStart(BC_CFG.PATIENT_PAD, '0');
  }
  props.setProperty(BC_CFG.PROP_PATIENT_SEQ, String(n));
  return String(candidate);
}

/**
 * Next per-day sequential ID, e.g. LAB-ORD-20260911-0007.
 * @param {string} series    property namespace, e.g. 'LAB_ORDER'
 * @param {string} prefix    literal prefix, e.g. 'LAB-ORD-'
 * @param {string} sep       separator between date and number ('-' or '')
 * @param {number} width     zero-pad width of the number
 * @param {Sheet=} sheet     optional: verify uniqueness against this sheet
 * @param {number=} col1     1-based column to verify against
 */
function bc_nextDailyId_(series, prefix, sep, width, sheet, col1) {
  var props = PropertiesService.getScriptProperties();
  var key = 'BC_SEQ_' + series;
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');

  var raw = String(props.getProperty(key) || '');
  var parts = raw.split('|');
  var n = (parts[0] === today) ? (parseInt(parts[1], 10) || 0) : 0;

  var candidate = '';
  for (var guard = 0; guard < 1000; guard++) {
    n++;
    candidate = prefix + today + sep + String(n).padStart(width, '0');
    if (!sheet || !col1 || !bc_existsInColumn_(sheet, col1, candidate)) break;
  }
  props.setProperty(key, today + '|' + n);
  return String(candidate);
}

function bc_existsInColumn_(sheet, col1, value) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  return sheet.getRange(2, col1, lastRow - 1, 1)
    .createTextFinder(String(value)).matchEntireCell(true).findNext() !== null;
}

/**
 * ADMIN-RUN, READ-ONLY. Run once from the editor BEFORE printing any patient
 * barcode. Lists every Patient ID that appears on more than one row.
 */
function auditDuplicatePatientIds() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Patients');
  if (!sheet || sheet.getLastRow() < 2) return 'Patients sheet is empty.';
  var vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  var rows = {};
  vals.forEach(function (r, i) {
    var id = dc_upper_(r[0]);
    if (!id) return;
    (rows[id] = rows[id] || []).push('row ' + (i + 2) + ' (' + dc_str_(r[2]) + ')');
  });
  var dups = Object.keys(rows).filter(function (k) { return rows[k].length > 1; });
  var report = dups.length
    ? 'DUPLICATE PATIENT IDS (' + dups.length + '):\n' +
      dups.map(function (k) { return k + ' -> ' + rows[k].join(', '); }).join('\n')
    : 'No duplicate Patient IDs. Safe to print patient barcodes.';
  Logger.log(report);
  return report;
}

// ============================================================================
// SECTION B — CLASSIFICATION (server is authoritative)
// ============================================================================

function bc_classify_(payload) {
  var s = dc_upper_(payload).replace(/\s+/g, '');
  if (!s || s.length > BC_CFG.MAX_PAYLOAD) return { type: 'UNKNOWN', value: '' };

  var prefix = bc_patientPrefix_();
  var pRe = new RegExp('^(?:P-)?(' + bc_escapeRegex_(prefix) + '\\d{4,})$');
  var m = s.match(pRe);
  if (m) return { type: 'PATIENT', value: m[1] };

  if (/^BC\d{8}[0-9A-F]{5}$/.test(s))            return { type: 'LAB_SAMPLE', value: s };
  if (/^LAB-ORD-\d{8}-[0-9A-F]{4}$/.test(s))     return { type: 'LAB_ORDER', value: s };
  return { type: 'UNKNOWN', value: '' };
}

// ============================================================================
// SECTION C — FRONTEND ENTRY POINTS
// ============================================================================

/** Router bootstrap: detector thresholds, ID pattern, role actions. */
function getScanRouterConfig(sessionToken) {
  try {
    var sess = dc_validateSession_(sessionToken);
    if (!sess) return { success: false, message: 'Your session has expired. Please sign in again.' };
    var role = dc_str_(sess.role).toLowerCase();
    if (role === 'patient') return { success: false, message: 'Not authorised.' };

    return {
      success: true,
      message: 'Scanner ready.',
      config: {
        patientPrefix: bc_patientPrefix_(),
        maxGapMs: BC_CFG.SCAN_MAX_GAP_MS,
        minLen: BC_CFG.SCAN_MIN_LEN,
        role: role,
        actions: (BC_CFG.ROLE_ACTIONS[role] || []).slice()
      }
    };
  } catch (e) {
    return { success: false, message: 'Scanner setup failed: ' + e.message };
  }
}

/**
 * READ-ONLY. Resolves a scanned payload for the caller's role and scope.
 * @param {string} payload       raw scanned text
 * @param {string} context       active screen id (for the audit trail only)
 * @param {string} sessionToken
 */
function resolveScan(payload, context, sessionToken) {
  try {
    // resolveScope_ validates the session itself; calling dc_validateSession_
    // as well would touch the Sessions sheet twice per scan.
    var scope = resolveScope_(sessionToken, null);
    if (!scope.ok) return { success: false, message: scope.message || 'Not authorised.' };
    var role = scope.role;

    var cls = bc_classify_(payload);
    if (cls.type === 'UNKNOWN') {
      return { success: false, message: 'This code is not a CresRx ID.', type: 'UNKNOWN' };
    }
    if (cls.type !== 'PATIENT') {
      bc_auditScan_(scope, cls, context, false);
      return { success: true, type: cls.type, id: cls.value, found: false,
               message: 'Lab barcode recognised. Open its order in the Lab workspace to collect or receive it.' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var patient = bc_findPatient_(ss, cls.value);
    if (!patient) {
      bc_auditScan_(scope, cls, context, false);
      return { success: true, type: 'PATIENT', id: cls.value, found: false,
               message: 'No patient registered with ID ' + cls.value + '.' };
    }

    var appointments = bc_todaysAppointments_(ss, cls.value, scope);
    var admission = bc_liveAdmission_(ss, cls.value);

    bc_auditScan_(scope, cls, context, true);

    return {
      success: true,
      type: 'PATIENT',
      id: cls.value,
      found: true,
      patient: patient,
      appointments: appointments,
      admission: admission,
      actions: bc_scanActions_(role, admission),
      message: ''
    };
  } catch (e) {
    return { success: false, message: 'Scan lookup failed: ' + e.message };
  }
}

/**
 * The actions this scan can actually perform.
 *
 * The role matrix says what someone MAY do; this says what is available for
 * the patient in front of them. Ward actions need a live admission — offering
 * "New IP note" for a patient who was discharged last month produces a note
 * with nowhere to go — and the OP actions that assume a clinic queue are
 * demoted at the bedside, where the wristband is being scanned to reach a
 * chart, not to start a consultation.
 */
function bc_scanActions_(role, admission) {
  var all = (BC_CFG.ROLE_ACTIONS[role] || []).slice();
  var IP_ONLY = ['IP_CASESHEET', 'IP_NOTE'];

  if (!admission || !dc_str_(admission.ipNumber)) {
    return all.filter(function (a) { return IP_ONLY.indexOf(a) === -1; });
  }

  // Admitted: put the ward actions first, so the bedside scan lands on them.
  var ip = all.filter(function (a) { return IP_ONLY.indexOf(a) !== -1; });
  var rest = all.filter(function (a) { return IP_ONLY.indexOf(a) === -1; });
  return ip.concat(rest);
}

/**
 * WRITE. Advances today's appointment from a scan: Booked -> Arrived,
 * Booked/Arrived -> In-Progress. Scope-checked and audited.
 */
function barcodeUpdateAppointment(apptId, targetStatus, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(BC_CFG.LOCK_MS);

    var sess = dc_validateSession_(sessionToken);
    if (!sess) return { success: false, message: 'Your session has expired. Please sign in again.' };
    var role = dc_str_(sess.role).toLowerCase();
    if (BC_CFG.STATUS_WRITERS.indexOf(role) === -1) {
      return { success: false, message: 'Your role cannot change appointment status.' };
    }

    var id = dc_str_(apptId);
    var target = dc_str_(targetStatus);
    if (!id) return { success: false, message: 'Appointment ID is missing.' };

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Appointments');
    if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'Appointments sheet is missing.' };

    var cell = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
      .createTextFinder(id).matchEntireCell(true).findNext();
    if (!cell) return { success: false, message: 'Appointment ' + id + ' was not found.' };

    var rowNum = cell.getRow();
    var lastCol = sheet.getLastColumn();
    var row = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
    var m = dc_headerMap_(sheet);
    var statusIdx = (m['Status'] === undefined) ? 6 : m['Status'];
    var docIdx = (m['Doctor_ID'] === undefined) ? -1 : m['Doctor_ID'];

    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (dc_dateKey_(row[3]) !== today) {
      return { success: false, message: 'Only today\'s appointments can be updated by scan.' };
    }

    var scope = resolveScope_(sessionToken, null);
    var rowDoc = (docIdx === -1) ? DC_DEFAULT_DOCTOR : (dc_str_(row[docIdx]) || DC_DEFAULT_DOCTOR);
    if (!dc_inScope_(scope, rowDoc)) {
      return { success: false, message: 'This appointment belongs to another doctor.' };
    }

    var current = dc_str_(row[statusIdx]);
    if (current === target) {
      return { success: true, status: current, message: 'Already ' + target + '.' };
    }
    var allowed = BC_CFG.TRANSITIONS[current] || [];
    if (allowed.indexOf(target) === -1) {
      return { success: false, status: current,
               message: 'Cannot change status from ' + (current || 'blank') + ' to ' + target + '.' };
    }

    sheet.getRange(rowNum, statusIdx + 1).setValue(String(target));
    SpreadsheetApp.flush();

    logAudit_(sess, 'APPOINTMENT_STATUS_SCAN', 'Appointment', id,
              { from: current, to: target, patientId: dc_upper_(row[1]) });

    return { success: true, status: target, message: 'Marked ' + target + '.' };
  } catch (e) {
    return { success: false, message: 'Status update failed: ' + e.message };
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/**
 * Server-side scan audit. Written by resolveScan() itself, so the ledger does
 * not depend on the browser choosing to report. Never throws: a failed audit
 * must not fail the scan the clinician is waiting on.
 */
function bc_auditScan_(scope, cls, context, found) {
  try {
    logAudit_({ username: scope.username, role: scope.role, doctorId: scope.selfDoctorId },
              'BARCODE_SCAN', dc_str_(cls.type), dc_str_(cls.value),
              { context: dc_str_(context).substring(0, 60), found: !!found });
  } catch (e) { /* audit is best-effort, the scan is not */ }
}

// ============================================================================
// SECTION D — READ HELPERS (bounded reads, TextFinder, no getDataRange)
// ============================================================================

function bc_findPatient_(ss, pid) {
  var sheet = ss.getSheetByName('Patients');
  if (!sheet || sheet.getLastRow() < 2) return null;
  var want = dc_upper_(pid);
  var cache = CacheService.getScriptCache();
  var cacheKey = 'BC_PIDROW_' + want;
  var rowNum = 0;

  // Fast path: cached row number, re-verified because rows can shift.
  var cached = parseInt(cache.get(cacheKey) || '0', 10);
  if (cached >= 2 && cached <= sheet.getLastRow()) {
    if (dc_upper_(sheet.getRange(cached, 1).getValue()) === want) rowNum = cached;
  }
  if (!rowNum) {
    var cell = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
      .createTextFinder(want).matchEntireCell(true).matchCase(false).findNext();
    if (!cell) return null;
    rowNum = cell.getRow();
    try { cache.put(cacheKey, String(rowNum), BC_CFG.PID_ROW_CACHE_S); } catch (e) {}
  }

  // Columns C:E only — Name, Age, Gender. Column B (password) is never read.
  var v = sheet.getRange(rowNum, 3, 1, 3).getValues()[0];
  return { id: want, name: dc_str_(v[0]), age: dc_str_(v[1]), gender: dc_str_(v[2]) };
}

function bc_todaysAppointments_(ss, pid, scope) {
  var sheet = ss.getSheetByName('Appointments');
  if (!sheet || sheet.getLastRow() < 2) return [];
  var want = dc_upper_(pid);
  var cells = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1)
    .createTextFinder(want).matchEntireCell(true).matchCase(false).findAll();
  if (!cells.length) return [];

  var m = dc_headerMap_(sheet);
  var lastCol = sheet.getLastColumn();
  var docIdx = (m['Doctor_ID'] === undefined) ? -1 : m['Doctor_ID'];
  var snapIdx = (m['Doctor_Name_Snapshot'] === undefined) ? -1 : m['Doctor_Name_Snapshot'];
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var out = [];

  cells.forEach(function (c) {
    var r = sheet.getRange(c.getRow(), 1, 1, lastCol).getValues()[0];
    if (dc_dateKey_(r[3]) !== today) return;
    var status = dc_str_(r[6]);
    if (['Blocked', 'DELETE', 'Cancelled'].indexOf(status) !== -1) return;
    var rowDoc = (docIdx === -1) ? DC_DEFAULT_DOCTOR : (dc_str_(r[docIdx]) || DC_DEFAULT_DOCTOR);
    if (!dc_inScope_(scope, rowDoc)) return;
    var docName = (snapIdx !== -1) ? dc_str_(r[snapIdx]) : '';
    if (!docName) { var d = dc_getDoctorById_(rowDoc); docName = d ? d.name : rowDoc; }
    out.push({
      apptId: dc_str_(r[0]),
      time12: dc_to12_(r[4]),
      purpose: dc_str_(r[5]),
      status: status,
      doctorId: rowDoc,
      doctorName: docName
    });
  });
  return out;
}

function bc_liveAdmission_(ss, pid) {
  var sheet = ss.getSheetByName(IPA_CFG.SHEET);
  if (!sheet || sheet.getLastRow() < 2) return null;
  var cells = sheet.getRange(2, IPA_COL.PATIENT_ID + 1, sheet.getLastRow() - 1, 1)
    .createTextFinder(dc_upper_(pid)).matchEntireCell(true).matchCase(false).findAll();

  for (var i = cells.length - 1; i >= 0; i--) {
    var r = sheet.getRange(cells[i].getRow(), 1, 1, IPA_CFG.HEADERS.length).getValues()[0];
    if (!ipa_isLive_(r[IPA_COL.STATUS])) continue;
    var wb = ipa_resolveWardBed_(r[IPA_COL.WARD], r[IPA_COL.BED]);
    return {
      ipNumber: dc_str_(r[IPA_COL.IP]),
      ward: wb.ward,
      bed: wb.bed,
      consultant: dc_str_(r[IPA_COL.CONSULTANT])
    };
  }
  return null;
}


// ============================================================================
// SECTION E — PHASE 2 : LABELS AND SPECIMEN SAFETY
// ============================================================================

/** Display name for printed labels. Set CLINIC_NAME in Script Properties. */
function bc_clinicName_() {
  var n = '';
  try { n = PropertiesService.getScriptProperties().getProperty('CLINIC_NAME') || ''; } catch (e) {}
  return dc_str_(n) || 'CRESCENTIA HEALTHTECH';
}

/**
 * Data for a printed patient card or chart sticker. SESSION REQUIRED, staff
 * only. The card shows a name; the QR on it carries only the ID.
 */
function getPatientLabelData(patientId, sessionToken) {
  try {
    var scope = resolveScope_(sessionToken, null);
    if (!scope.ok) return { success: false, message: scope.message || 'Not authorised.' };

    var cls = bc_classify_(patientId);
    var want = (cls.type === 'PATIENT') ? cls.value : dc_upper_(patientId);
    if (!want) return { success: false, message: 'Patient ID is missing.' };

    var p = bc_findPatient_(SpreadsheetApp.getActiveSpreadsheet(), want);
    if (!p) return { success: false, message: 'No patient registered with ID ' + want + '.' };

    // Blood group and registration date sit outside the C:E window bc_findPatient_
    // reads, so they are fetched only when a label is actually being printed.
    var extra = { bloodGroup: '', regDate: '' };
    try {
      var full = pt_readProfile_(want);
      if (full) {
        extra.bloodGroup = dc_str_(full.bloodGroup);
        extra.regDate = dc_str_(full.regDate).substring(0, 10);
      }
    } catch (e) { /* label still prints without them */ }

    logAudit_({ username: scope.username, role: scope.role, doctorId: scope.selfDoctorId },
              'LABEL_PRINTED', 'Patient', want, { kind: 'PATIENT_CARD' });

    return {
      success: true,
      patient: {
        id: p.id, name: p.name, age: p.age, gender: p.gender,
        bloodGroup: extra.bloodGroup, regDate: extra.regDate,
        clinic: bc_clinicName_()
      }
    };
  } catch (e) {
    return { success: false, message: 'Label lookup failed: ' + e.message };
  }
}

/**
 * Tube labels for one lab order. SESSION REQUIRED.
 * Returns one entry per collected, non-rejected tube.
 */
function getSampleLabelData(orderId, sessionToken) {
  try {
    var scope = resolveScope_(sessionToken, null);
    if (!scope.ok) return { success: false, message: scope.message || 'Not authorised.' };

    var id = dc_str_(orderId);
    if (!id) return { success: false, message: 'Order ID is missing.' };

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LAB.SAMPLES);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'No samples found.' };

    var map = labHeaderMap(sheet);
    var cells = sheet.getRange(2, map['OrderID'] + 1, sheet.getLastRow() - 1, 1)
      .createTextFinder(id).matchEntireCell(true).findAll();
    if (!cells.length) return { success: false, message: 'No tubes recorded for ' + id + '.' };

    var lastCol = sheet.getLastColumn();
    var patient = null;
    var tubes = [];
    cells.forEach(function (c) {
      var r = sheet.getRange(c.getRow(), 1, 1, lastCol).getValues()[0];
      if (dc_upper_(r[map['CollectionStatus']]) === 'REJECTED') return;
      if (!patient) patient = bc_findPatient_(SpreadsheetApp.getActiveSpreadsheet(), dc_upper_(r[map['PatientID']]));
      tubes.push({
        barcode: dc_str_(r[map['BarcodeID']]),
        sampleId: dc_str_(r[map['SampleID']]),
        sampleType: dc_str_(r[map['SampleType']]).replace(/_/g, ' '),
        patientId: dc_upper_(r[map['PatientID']]),
        patientName: dc_str_(r[map['PatientName']]),
        age: patient ? patient.age : '',
        gender: patient ? patient.gender : '',
        collectedAt: dc_str_(r[map['CollectedAt']]).substring(0, 16)
      });
    });

    if (!tubes.length) return { success: false, message: 'All tubes for this order were rejected.' };

    logAudit_({ username: scope.username, role: scope.role, doctorId: scope.selfDoctorId },
              'LABEL_PRINTED', 'LabOrder', id, { kind: 'TUBE_LABELS', count: tubes.length });

    return { success: true, orderId: id, tubes: tubes, clinic: bc_clinicName_() };
  } catch (e) {
    return { success: false, message: 'Tube label lookup failed: ' + e.message };
  }
}

/**
 * TWO-SCAN IDENTITY MATCH, performed before a tube is drawn.
 *
 * The phlebotomist scans the patient's card, then works the order on screen.
 * If the scanned patient is not the patient the order belongs to, collection
 * is refused. This is the barcode control against wrong-blood-in-tube — the
 * error class that a verbal "state your name and date of birth" check is
 * meant to catch and routinely does not, because the tube is often labelled
 * away from the bedside.
 *
 * A mismatch is audited as a patient-safety event, not just refused: a near
 * miss that nobody can see later is a near miss that repeats.
 */
function verifyCollectionIdentity(orderId, scannedCode, sessionToken) {
  try {
    var scope = resolveScope_(sessionToken, null);
    if (!scope.ok) return { success: false, message: scope.message || 'Not authorised.' };

    var id = dc_str_(orderId);
    if (!id) return { success: false, message: 'Order ID is missing.' };

    var cls = bc_classify_(scannedCode);
    if (cls.type !== 'PATIENT') {
      return { success: true, match: false, reason: 'NOT_A_PATIENT_CODE',
               message: 'That is not a patient code. Scan the patient\'s card or wristband.' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var oSheet = ss.getSheetByName(LAB.ORDERS);
    if (!oSheet || oSheet.getLastRow() < 2) return { success: false, message: 'Lab orders sheet is missing.' };

    var oMap = labHeaderMap(oSheet);
    var cell = oSheet.getRange(2, oMap['OrderID'] + 1, oSheet.getLastRow() - 1, 1)
      .createTextFinder(id).matchEntireCell(true).findNext();
    if (!cell) return { success: false, message: 'Order ' + id + ' was not found.' };

    var row = oSheet.getRange(cell.getRow(), 1, 1, oSheet.getLastColumn()).getValues()[0];
    var expectedId = dc_upper_(row[oMap['PatientID']]);
    var expectedName = dc_str_(row[oMap['PatientName']]);
    var match = (expectedId === cls.value);

    var actor = { username: scope.username, role: scope.role, doctorId: scope.selfDoctorId };
    if (!match) {
      var scanned = bc_findPatient_(ss, cls.value);
      logAudit_(actor, 'SPECIMEN_ID_MISMATCH', 'LabOrder', id, {
        expectedPatient: expectedId, scannedPatient: cls.value,
        scannedName: scanned ? scanned.name : 'UNKNOWN_ID'
      });
      return {
        success: true, match: false, reason: 'WRONG_PATIENT',
        expected: { id: expectedId, name: expectedName },
        scanned: { id: cls.value, name: scanned ? scanned.name : '' },
        message: 'STOP — wrong patient. This order belongs to ' + expectedName +
                 ' (' + expectedId + '), but ' + (scanned ? scanned.name : 'ID ' + cls.value) +
                 ' was scanned.'
      };
    }

    logAudit_(actor, 'SPECIMEN_ID_VERIFIED', 'LabOrder', id, { patientId: expectedId });
    return {
      success: true, match: true,
      expected: { id: expectedId, name: expectedName },
      message: 'Identity confirmed by barcode — ' + expectedName + '.'
    };
  } catch (e) {
    return { success: false, message: 'Identity check failed: ' + e.message };
  }
}

/**
 * PER-TUBE receiving at the bench. SESSION REQUIRED.
 *
 * receiveLabSample(orderId) receives an entire order in one go, so a tube that
 * never reached the lab is silently marked received and its TAT clock starts
 * against a specimen nobody has. Scanning each tube receives exactly that tube
 * and reports what is still outstanding, so a missing tube is visible at the
 * bench instead of at result-entry.
 */
function receiveLabSampleByBarcode(barcodeId, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(BC_CFG.LOCK_MS);

    var sess = dc_validateSession_(sessionToken);
    if (!sess) return { success: false, message: 'Your session has expired. Please sign in again.' };
    var role = dc_str_(sess.role).toLowerCase();
    if (['admin', 'lab', 'lab technician', 'nurse'].indexOf(role) === -1) {
      return { success: false, message: 'Your role cannot receive lab samples.' };
    }

    var code = dc_upper_(barcodeId);
    if (!code) return { success: false, message: 'Tube barcode is missing.' };

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LAB.SAMPLES);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'No samples found.' };

    var map = labHeaderMap(sheet);
    var cell = sheet.getRange(2, map['BarcodeID'] + 1, sheet.getLastRow() - 1, 1)
      .createTextFinder(code).matchEntireCell(true).matchCase(false).findNext();
    if (!cell) {
      return { success: false, message: 'Tube ' + code + ' is not in this system. Do not process it.' };
    }

    var rowNum = cell.getRow();
    var lastCol = sheet.getLastColumn();
    var row = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
    var orderId = dc_str_(row[map['OrderID']]);

    if (dc_upper_(row[map['CollectionStatus']]) === 'REJECTED') {
      return { success: false, orderId: orderId,
               message: 'Tube ' + code + ' was rejected (' + dc_str_(row[map['RejectionReason']]) + '). Do not process it.' };
    }

    var already = dc_str_(row[map['ReceivedAtLabAt']]);
    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    if (!already) {
      sheet.getRange(rowNum, map['ReceivedAtLabAt'] + 1).setValue(nowStr);
      sheet.getRange(rowNum, map['ReceivedBy'] + 1).setValue(dc_str_(sess.username) || 'SYSTEM');
    }

    // What is still missing from this order?
    var outstanding = [];
    var receivedCount = 0, totalCount = 0;
    var cells = sheet.getRange(2, map['OrderID'] + 1, sheet.getLastRow() - 1, 1)
      .createTextFinder(orderId).matchEntireCell(true).findAll();
    cells.forEach(function (c) {
      var r = sheet.getRange(c.getRow(), 1, 1, lastCol).getValues()[0];
      if (dc_upper_(r[map['CollectionStatus']]) === 'REJECTED') return;
      totalCount++;
      if (dc_str_(r[map['ReceivedAtLabAt']])) { receivedCount++; return; }
      outstanding.push({
        barcode: dc_str_(r[map['BarcodeID']]),
        sampleType: dc_str_(r[map['SampleType']]).replace(/_/g, ' ')
      });
    });

    var complete = (outstanding.length === 0);
    if (complete) advanceOrderStatus(orderId, 'IN_PROCESS');

    logAudit_(sess, 'SAMPLE_RECEIVED_SCAN', 'LabSample', code,
              { orderId: orderId, alreadyReceived: !!already, outstanding: outstanding.length });
    SpreadsheetApp.flush();

    return {
      success: true,
      orderId: orderId,
      barcode: code,
      duplicate: !!already,
      sampleType: dc_str_(row[map['SampleType']]).replace(/_/g, ' '),
      patientName: dc_str_(row[map['PatientName']]),
      patientId: dc_upper_(row[map['PatientID']]),
      received: receivedCount,
      total: totalCount,
      outstanding: outstanding,
      complete: complete,
      message: already
        ? 'Tube ' + code + ' was already received.'
        : (complete
            ? 'Tube received. All ' + totalCount + ' tube(s) are in — processing started.'
            : 'Tube received. ' + outstanding.length + ' tube(s) still outstanding.')
    };
  } catch (e) {
    return { success: false, message: 'Receive failed: ' + e.message };
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}