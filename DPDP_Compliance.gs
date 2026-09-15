// ============================================================================
// DPDP_Compliance.gs  —  Crescentia HealthTech
// The Digital Personal Data Protection Act, 2023: the parts a system has to
// provide rather than a policy document.
// ----------------------------------------------------------------------------
// WHAT THE ACT REQUIRES OF THIS APPLICATION
//
// A clinic processing patient records is a DATA FIDUCIARY and the patient is
// a DATA PRINCIPAL. Four obligations cannot be met by writing a policy — the
// software has to do them:
//
//   s.5   NOTICE      — tell the patient, at or before collection, what is
//                       collected, why, how to withdraw and how to complain.
//   s.6   CONSENT     — free, specific, informed, unambiguous, with a clear
//                       affirmative action, and AS EASY TO WITHDRAW AS TO
//                       GIVE. A consent nobody recorded cannot be proved.
//   s.8(7) RETENTION  — erase when the purpose is served, unless retention is
//                       required by law. For clinical records that law is the
//                       NMC Ethics Regulations (3 years) and the Clinical
//                       Establishments rules; for a marketing mobile number
//                       there is no such law.
//   s.11-13 RIGHTS    — access, correction, erasure, grievance redressal, and
//                       nomination. Each has a statutory clock.
//
// WHAT THIS FILE IS AND IS NOT
//
// It provides the registers and the machinery: consent recorded per purpose
// with its notice version, withdrawal, a data-principal request queue with
// its deadlines, a retention sweep that reports before it deletes, and a
// shared-link register so a document put on Drive can be un-shared.
//
// It does NOT make the clinic compliant. Compliance is a Consent Notice
// written for patients, a named grievance officer, a retention schedule the
// clinic has decided, and the security fixes listed in
// docs/DPDP_READINESS.md — several of which are in other files and some of
// which are deployment settings, not code.
//
// SETUP
//   Run dpdpSetup() once. Then dpdpReadinessCheck() to see where you stand.
// ============================================================================

var DPDP_CFG = {
  CONSENT:   'Consent_Register',
  REQUESTS:  'DPDP_Requests',
  SHARED:    'Shared_Documents',
  NOTICE_PROP: 'DPDP_NOTICE_VERSION',
  OFFICER_PROP: 'DPDP_GRIEVANCE_OFFICER',
  TZ: 'Asia/Kolkata',

  // Statutory-ish clocks. The Act sets the obligation; the Rules set the
  // period. These are the clinic's committed response times, shown to the
  // patient and used to flag an overdue request.
  RESPONSE_DAYS: { ACCESS: 30, CORRECTION: 30, ERASURE: 30, GRIEVANCE: 30 },

  // How long each kind of record is kept once its purpose is served.
  // CLINICAL is three years from the last entry, which is what the NMC
  // Ethics Regulations require of an outpatient record; inpatient and
  // medico-legal records are kept longer by local rule, so the sweep
  // REPORTS and never deletes on its own.
  RETENTION_DAYS: {
    CLINICAL: 1095,          // 3 years — NMC Ethics Regulations
    FINANCIAL: 2920,         // 8 years — Income Tax Act s.44AA/44AB practice
    SESSION: 30,             // sign-in sessions
    AUDIT: 1095,             // the audit trail itself
    MARKETING: 365,          // a number kept only to send reminders
    SHARED_LINK: 30          // a document put on Drive for a patient
  }
};

function dpdp_str_(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }
function dpdp_now_() { return new Date(); }
function dpdp_fmt_(d) {
  if (!d) return '';
  var dt = (d instanceof Date) ? d : new Date(d);
  if (isNaN(dt.getTime())) return '';
  return Utilities.formatDate(dt, DPDP_CFG.TZ, 'yyyy-MM-dd HH:mm');
}

/**
 * The purposes consent is asked for, separately.
 *
 * s.6 requires consent to be SPECIFIC. One "I agree" covering treatment,
 * insurance, reminders and research is not specific, and in practice it is
 * the marketing purpose that a patient wants to refuse while still wanting
 * to be treated — which is exactly why they must be separable.
 *
 * TREATMENT is marked `legitimate` because care of a patient who has
 * approached a clinic is a legitimate use under s.7(a)-(b): it is recorded
 * so the record shows it was considered, not gated behind a tick box that
 * would stop someone being treated.
 */
var DPDP_PURPOSES = [
  { key: 'TREATMENT', legitimate: true, required: true,
    label: 'Care and treatment',
    detail: 'Recording your history, examination, investigations, prescriptions ' +
            'and admissions so that you can be treated and followed up.' },
  { key: 'BILLING', legitimate: true, required: true,
    label: 'Billing and accounts',
    detail: 'Raising bills and receipts, and keeping them for the period tax ' +
            'law requires.' },
  { key: 'INSURANCE', legitimate: false, required: false,
    label: 'Insurance and TPA claims',
    detail: 'Sharing your records with your insurer or third-party ' +
            'administrator so a claim can be settled.' },
  { key: 'COMMUNICATION', legitimate: false, required: false,
    label: 'Reports and reminders by WhatsApp or email',
    detail: 'Sending your prescription, lab report or invoice to your mobile ' +
            'number or email address.' },
  { key: 'MARKETING', legitimate: false, required: false,
    label: 'Health camps and offers',
    detail: 'Telling you about camps, screenings and services. You may refuse ' +
            'this and still be treated exactly as before.' },
  { key: 'RESEARCH', legitimate: false, required: false,
    label: 'De-identified research and audit',
    detail: 'Using your data with your name and contact details removed, for ' +
            'clinical audit or research.' }
];

function dpdp_purpose_(key) {
  var k = dpdp_str_(key).toUpperCase();
  for (var i = 0; i < DPDP_PURPOSES.length; i++) {
    if (DPDP_PURPOSES[i].key === k) return DPDP_PURPOSES[i];
  }
  return null;
}

// ---------------------------------------------------------------------------
// SHEETS
// ---------------------------------------------------------------------------

function dpdp_consentSheet_() {
  return dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), DPDP_CFG.CONSENT, [
    'Entry_ID', 'Patient_ID', 'Patient_Name', 'Purpose', 'Decision',
    'Notice_Version', 'Method', 'Recorded_By', 'Recorded_At',
    'Withdrawn_At', 'Withdrawn_Reason', 'Guardian_Name', 'Guardian_Relation', 'Notes'
  ]);
}

function dpdp_requestSheet_() {
  return dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), DPDP_CFG.REQUESTS, [
    'Request_ID', 'Patient_ID', 'Patient_Name', 'Type', 'Details',
    'Raised_By', 'Raised_At', 'Due_By', 'Status', 'Handled_By',
    'Closed_At', 'Outcome'
  ]);
}

function dpdp_sharedSheet_() {
  return dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), DPDP_CFG.SHARED, [
    'Entry_ID', 'File_ID', 'File_Name', 'Doc_Type', 'Patient_ID',
    'Shared_At', 'Shared_By', 'Expires_At', 'Status', 'Revoked_At', 'Link'
  ]);
}

/** ONE-OFF. Creates every register. Safe to re-run. */
function dpdpSetup() {
  dpdp_consentSheet_();
  dpdp_requestSheet_();
  dpdp_sharedSheet_();
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty(DPDP_CFG.NOTICE_PROP)) {
    props.setProperty(DPDP_CFG.NOTICE_PROP, 'v1-' +
      Utilities.formatDate(new Date(), DPDP_CFG.TZ, 'yyyyMMdd'));
  }
  return 'DPDP registers ready: ' + DPDP_CFG.CONSENT + ', ' + DPDP_CFG.REQUESTS +
         ', ' + DPDP_CFG.SHARED + '.\n' +
         'Now set the grievance officer:\n' +
         '  dpdpSetGrievanceOfficer("Dr. …", "officer@clinic.in", "+91 …")\n' +
         'and run dpdpReadinessCheck().';
}

/**
 * s.13 — a Data Fiduciary must publish the contact of the person who
 * answers a data principal's grievance. Not having one named is itself a
 * finding, so this is stored where the notice can read it.
 */
function dpdpSetGrievanceOfficer(name, email, phone) {
  var v = { name: dpdp_str_(name), email: dpdp_str_(email), phone: dpdp_str_(phone) };
  if (!v.name || !(v.email || v.phone)) {
    return 'Give a name and at least one of email or phone. ' +
           'This is what a patient is told to contact.';
  }
  PropertiesService.getScriptProperties()
    .setProperty(DPDP_CFG.OFFICER_PROP, JSON.stringify(v));
  return 'Grievance officer set: ' + v.name + ' (' + (v.email || v.phone) + ').';
}

function dpdp_officer_() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(DPDP_CFG.OFFICER_PROP);
    return raw ? JSON.parse(raw) : null;
  } catch (e) { return null; }
}

function dpdp_noticeVersion_() {
  try {
    return PropertiesService.getScriptProperties().getProperty(DPDP_CFG.NOTICE_PROP) || 'v1';
  } catch (e) { return 'v1'; }
}

// ---------------------------------------------------------------------------
// SECTION 5 — THE NOTICE
// ---------------------------------------------------------------------------

/**
 * FRONTEND ENTRY. What the patient is told, in the words they are told it.
 *
 * Unauthenticated on purpose: a notice a person has to sign in to read is
 * not a notice given "at or before" collection.
 */
function getDPDPNotice() {
  var clinic = (typeof cresc_clinic_ === 'function')
    ? cresc_clinic_() : { name: 'this clinic', address: '', phone: '', email: '' };
  var officer = dpdp_officer_();

  return {
    success: true,
    version: dpdp_noticeVersion_(),
    fiduciary: clinic.name,
    contact: { address: clinic.address, phone: clinic.phone, email: clinic.email },
    purposes: DPDP_PURPOSES,
    rights: [
      { key: 'ACCESS',     label: 'A copy of what is held about you',
        detail: 'You may ask for a summary of your personal data and who it has been shared with.' },
      { key: 'CORRECTION', label: 'Correction of anything wrong',
        detail: 'You may ask for inaccurate or incomplete data to be corrected or completed.' },
      { key: 'ERASURE',    label: 'Erasure',
        detail: 'You may ask for your data to be erased. Clinical records the law requires ' +
                'the clinic to keep — normally three years from your last visit — cannot be ' +
                'erased until that period ends; you will be told which those are.' },
      { key: 'WITHDRAW',   label: 'Withdrawal of consent',
        detail: 'You may withdraw any consent that was not required for your treatment, ' +
                'as easily as you gave it.' },
      { key: 'GRIEVANCE',  label: 'Complaint',
        detail: 'You may complain to the person named below, and afterwards to the ' +
                'Data Protection Board of India.' },
      { key: 'NOMINATE',   label: 'Nomination',
        detail: 'You may nominate someone to exercise these rights if you die or ' +
                'become incapable of exercising them.' }
    ],
    grievanceOfficer: officer,
    // A missing officer is a finding, and it is shown rather than hidden.
    warnings: officer ? [] : [
      'No grievance officer has been named. Section 13 of the DPDP Act requires ' +
      'a Data Fiduciary to publish one. Run dpdpSetGrievanceOfficer().'
    ],
    responseDays: DPDP_CFG.RESPONSE_DAYS
  };
}

// ---------------------------------------------------------------------------
// SECTION 6 — CONSENT
// ---------------------------------------------------------------------------

/**
 * FRONTEND ENTRY. Records a consent decision, per purpose.
 *
 * APPEND-ONLY. A consent that can be edited is a consent that cannot be
 * proved, and s.6(10) puts the burden of proof on the Data Fiduciary.
 * Changing your mind writes another row; the latest row for a purpose is
 * the current answer.
 *
 * @param {{patientId, decisions:Object, method, guardianName, guardianRelation}} payload
 *        decisions is { PURPOSE_KEY: true|false }
 * @param {string} sessionToken
 */
function recordConsent(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    payload = payload || {};

    var actor = crescRequire_(sessionToken, ['patient.register', 'patient.write']);

    var patientId = dpdp_str_(payload.patientId).toUpperCase();
    if (!patientId) return { success: false, message: 'No patient was named.' };

    var decisions = payload.decisions || {};
    var keys = Object.keys(decisions);
    if (!keys.length) return { success: false, message: 'No consent decisions were given.' };

    var profile = null;
    try { profile = pt_readProfile_(patientId); } catch (e) { profile = null; }
    if (!profile) return { success: false, message: 'No patient with the ID ' + patientId + '.' };

    // s.9 — a child's data needs verifiable consent from a parent or
    // guardian. The clinic cannot give that consent on the child's behalf,
    // so the guardian has to be named.
    var age = parseInt(profile.age, 10);
    var isChild = isFinite(age) && age < 18;
    if (isChild && !dpdp_str_(payload.guardianName)) {
      return { success: false, code: 'GUARDIAN_REQUIRED',
               message: profile.name + ' is under 18. Section 9 of the DPDP Act ' +
                        'requires consent from a parent or lawful guardian — record ' +
                        'their name and relationship.' };
    }

    var sh = dpdp_consentSheet_();
    var now = dpdp_now_();
    var version = dpdp_noticeVersion_();
    var rows = [], bad = [];

    keys.forEach(function (k) {
      var purpose = dpdp_purpose_(k);
      if (!purpose) { bad.push(k); return; }
      rows.push([
        'CNS-' + Utilities.formatDate(now, DPDP_CFG.TZ, 'yyMMdd-HHmmss') + '-' +
          Utilities.getUuid().substring(0, 4).toUpperCase(),
        patientId,
        profile.name,
        purpose.key,
        decisions[k] ? 'GIVEN' : 'REFUSED',
        version,
        dpdp_str_(payload.method) || 'IN_PERSON',
        actor.username,
        now,
        '', '',
        dpdp_str_(payload.guardianName),
        dpdp_str_(payload.guardianRelation),
        dpdp_str_(payload.notes)
      ]);
    });

    if (bad.length) {
      return { success: false,
               message: 'Unknown consent purpose(s): ' + bad.join(', ') + '.' };
    }

    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    dc_invalidate_(DPDP_CFG.CONSENT);
    SpreadsheetApp.flush();

    try {
      logAudit_({ username: actor.username, role: actor.role },
                'DPDP_CONSENT_RECORDED', 'Patient', patientId,
                { decisions: decisions, noticeVersion: version, child: isChild });
    } catch (e) { /* best effort */ }

    return { success: true, recorded: rows.length, noticeVersion: version,
             message: rows.length + ' consent decision(s) recorded against notice ' +
                      version + '.' };

  } catch (err) {
    var m = String((err && err.message) || err);
    return { success: false, message: m.replace('FORBIDDEN: ', '') };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * FRONTEND ENTRY. Where a patient's consent currently stands, per purpose.
 *
 * Reports NOT_ASKED separately from REFUSED. They look the same to a system
 * that only stores positives, and they are not the same fact: one is a
 * decision the patient made and the other is a question the clinic never
 * put to them.
 */
function getConsentStatus(patientId, sessionToken) {
  try {
    crescRequire_(sessionToken, 'patient.read');
    var want = dpdp_str_(patientId).toUpperCase();
    if (!want) return { success: false, message: 'No patient was named.' };

    var sh = dpdp_consentSheet_();
    var data = dc_sheetValues_(sh);
    var m = dc_headerMap_(sh);

    var latest = {};
    for (var i = 1; i < (data ? data.length : 0); i++) {
      if (dpdp_str_(data[i][m['Patient_ID']]).toUpperCase() !== want) continue;
      var p = dpdp_str_(data[i][m['Purpose']]).toUpperCase();
      // Later rows win: the register is append-only, so the last word is the
      // current one.
      latest[p] = {
        decision: dpdp_str_(data[i][m['Decision']]).toUpperCase(),
        at: dpdp_str_(data[i][m['Recorded_At']]),
        by: dpdp_str_(data[i][m['Recorded_By']]),
        noticeVersion: dpdp_str_(data[i][m['Notice_Version']]),
        withdrawnAt: dpdp_str_(data[i][m['Withdrawn_At']]),
        guardian: dpdp_str_(data[i][m['Guardian_Name']])
      };
    }

    var current = dpdp_noticeVersion_();
    var out = DPDP_PURPOSES.map(function (p) {
      var rec = latest[p.key];
      var state = !rec ? 'NOT_ASKED'
                : rec.withdrawnAt ? 'WITHDRAWN'
                : rec.decision;
      return {
        key: p.key, label: p.label, detail: p.detail,
        required: p.required, legitimate: p.legitimate,
        state: state,
        at: rec ? rec.at : '', by: rec ? rec.by : '',
        guardian: rec ? rec.guardian : '',
        noticeVersion: rec ? rec.noticeVersion : '',
        // Consent given against an older notice is not consent to the new
        // one. s.5 requires the notice to describe what is actually done.
        stale: !!(rec && rec.noticeVersion && rec.noticeVersion !== current)
      };
    });

    var unanswered = out.filter(function (o) { return o.state === 'NOT_ASKED'; }).length;
    return { success: true, patientId: want, purposes: out,
             noticeVersion: current, unanswered: unanswered, message: '' };

  } catch (err) {
    var m2 = String((err && err.message) || err);
    return { success: false, message: m2.replace('FORBIDDEN: ', '') };
  }
}

/**
 * FRONTEND ENTRY. Withdraws a consent.
 *
 * s.6(6): withdrawal must be AS EASY AS GIVING. So this takes the same
 * shape as recordConsent and refuses nothing except withdrawal of the two
 * purposes that are not consent-based in the first place — care and
 * billing — where the honest answer is that the clinic processes those
 * under s.7 and the remedy is erasure, not withdrawal.
 */
function withdrawConsent(patientId, purpose, reason, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var actor = crescRequire_(sessionToken, ['patient.write', 'portal.self']);

    var want = dpdp_str_(patientId).toUpperCase();
    var p = dpdp_purpose_(purpose);
    if (!want) return { success: false, message: 'No patient was named.' };
    if (!p) return { success: false, message: 'Unknown purpose "' + purpose + '".' };

    // A patient may withdraw only their own consent.
    if (actor.role === 'patient' && dpdp_str_(actor.username).toUpperCase() !== want) {
      return { success: false, message: 'You can only change your own consent.' };
    }

    if (p.legitimate) {
      return { success: false, code: 'LEGITIMATE_USE',
               message: '"' + p.label + '" is not processed on consent — it is a ' +
                        'legitimate use under section 7, because you approached the ' +
                        'clinic for it. It cannot be withdrawn, but you may ask for ' +
                        'erasure once the record is past the period the law requires ' +
                        'it to be kept.' };
    }

    var sh = dpdp_consentSheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    var now = dpdp_now_();
    var marked = 0;

    for (var i = 1; i < (data ? data.length : 0); i++) {
      if (dpdp_str_(data[i][m['Patient_ID']]).toUpperCase() !== want) continue;
      if (dpdp_str_(data[i][m['Purpose']]).toUpperCase() !== p.key) continue;
      if (dpdp_str_(data[i][m['Withdrawn_At']])) continue;
      sh.getRange(i + 1, m['Withdrawn_At'] + 1).setValue(now);
      sh.getRange(i + 1, m['Withdrawn_Reason'] + 1).setValue(dpdp_str_(reason));
      marked++;
    }

    // The withdrawal itself is a decision, and it is recorded as one.
    sh.appendRow([
      'CNS-' + Utilities.formatDate(now, DPDP_CFG.TZ, 'yyMMdd-HHmmss') + '-W',
      want, '', p.key, 'WITHDRAWN', dpdp_noticeVersion_(),
      'WITHDRAWAL', actor.username, now, now, dpdp_str_(reason), '', '', ''
    ]);
    dc_invalidate_(DPDP_CFG.CONSENT);
    SpreadsheetApp.flush();

    try {
      logAudit_({ username: actor.username, role: actor.role },
                'DPDP_CONSENT_WITHDRAWN', 'Patient', want,
                { purpose: p.key, reason: dpdp_str_(reason) });
    } catch (e) {}

    return { success: true, marked: marked,
             message: 'Consent for "' + p.label + '" withdrawn. ' +
                      'Stop using the data for that purpose from now.' };

  } catch (err) {
    var m3 = String((err && err.message) || err);
    return { success: false, message: m3.replace('FORBIDDEN: ', '') };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ---------------------------------------------------------------------------
// SECTIONS 11-13 — DATA PRINCIPAL RIGHTS
// ---------------------------------------------------------------------------

/**
 * FRONTEND ENTRY. Logs an access / correction / erasure / grievance request
 * and starts its clock.
 *
 * THE CLOCK IS THE POINT. A right with no deadline attached to it is a
 * request that sits in an inbox. Every row carries a Due_By, and
 * listDPDPRequests() reports what is overdue.
 */
function raiseDPDPRequest(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    payload = payload || {};
    var actor = crescRequire_(sessionToken, ['patient.read', 'portal.self']);

    var type = dpdp_str_(payload.type).toUpperCase();
    if (!DPDP_CFG.RESPONSE_DAYS.hasOwnProperty(type)) {
      return { success: false,
               message: 'Type must be one of: ' +
                        Object.keys(DPDP_CFG.RESPONSE_DAYS).join(', ') + '.' };
    }
    var patientId = dpdp_str_(payload.patientId).toUpperCase();
    if (!patientId) return { success: false, message: 'No patient was named.' };
    if (actor.role === 'patient' && dpdp_str_(actor.username).toUpperCase() !== patientId) {
      return { success: false, message: 'You can only raise a request about yourself.' };
    }
    var details = dpdp_str_(payload.details);
    if (details.length < 10) {
      return { success: false, message: 'Say what is being asked for, in a sentence.' };
    }

    var profile = null;
    try { profile = pt_readProfile_(patientId); } catch (e) { profile = null; }

    var now = dpdp_now_();
    var due = new Date(now.getTime() + DPDP_CFG.RESPONSE_DAYS[type] * 86400000);
    var id = 'DPR-' + Utilities.formatDate(now, DPDP_CFG.TZ, 'yyMMdd-HHmmss') + '-' +
             Utilities.getUuid().substring(0, 4).toUpperCase();

    dpdp_requestSheet_().appendRow([
      id, patientId, profile ? profile.name : '', type, details,
      actor.username, now, due, 'OPEN', '', '', ''
    ]);
    dc_invalidate_(DPDP_CFG.REQUESTS);
    SpreadsheetApp.flush();

    try {
      logAudit_({ username: actor.username, role: actor.role },
                'DPDP_REQUEST_RAISED', 'Patient', patientId, { type: type, id: id });
    } catch (e) {}

    return { success: true, requestId: id, dueBy: dpdp_fmt_(due),
             message: type.charAt(0) + type.slice(1).toLowerCase() +
                      ' request ' + id + ' logged. It is due by ' +
                      Utilities.formatDate(due, DPDP_CFG.TZ, 'dd-MMM-yyyy') + '.' };

  } catch (err) {
    var m = String((err && err.message) || err);
    return { success: false, message: m.replace('FORBIDDEN: ', '') };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/** FRONTEND ENTRY. Open requests, overdue first. */
function listDPDPRequests(sessionToken, opts) {
  try {
    opts = opts || {};
    crescRequire_(sessionToken, ['admin.audit', 'admin.config']);
    var sh = dpdp_requestSheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    if (!data || data.length < 2) {
      return { success: true, rows: [], overdue: 0, message: '' };
    }

    var now = Date.now(), rows = [];
    for (var i = 1; i < data.length; i++) {
      var status = dpdp_str_(data[i][m['Status']]).toUpperCase();
      if (opts.openOnly && status !== 'OPEN') continue;
      var due = data[i][m['Due_By']];
      var dueDate = (typeof cresc_parseDate_ === 'function') ? cresc_parseDate_(due) : new Date(due);
      var daysLeft = (dueDate && !isNaN(dueDate.getTime()))
        ? Math.ceil((dueDate.getTime() - now) / 86400000) : null;
      rows.push({
        requestId: dpdp_str_(data[i][m['Request_ID']]),
        patientId: dpdp_str_(data[i][m['Patient_ID']]),
        patientName: dpdp_str_(data[i][m['Patient_Name']]),
        type: dpdp_str_(data[i][m['Type']]),
        details: dpdp_str_(data[i][m['Details']]),
        raisedAt: dpdp_str_(data[i][m['Raised_At']]),
        dueBy: dpdp_str_(data[i][m['Due_By']]),
        daysLeft: daysLeft,
        overdue: (status === 'OPEN' && daysLeft !== null && daysLeft < 0),
        status: status,
        outcome: dpdp_str_(data[i][m['Outcome']])
      });
    }
    rows.sort(function (a, b) {
      if (a.overdue !== b.overdue) return a.overdue ? -1 : 1;
      return (a.daysLeft === null ? 9999 : a.daysLeft) - (b.daysLeft === null ? 9999 : b.daysLeft);
    });
    return { success: true, rows: rows,
             overdue: rows.filter(function (r) { return r.overdue; }).length, message: '' };
  } catch (err) {
    var m2 = String((err && err.message) || err);
    return { success: false, rows: [], message: m2.replace('FORBIDDEN: ', '') };
  }
}

/** FRONTEND ENTRY. Closes a request with what was actually done. */
function closeDPDPRequest(requestId, outcome, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var actor = crescRequire_(sessionToken, ['admin.audit', 'admin.config']);
    var want = dpdp_str_(requestId).toUpperCase();
    var text = dpdp_str_(outcome);
    if (text.length < 10) {
      return { success: false,
               message: 'Say what was done. "Closed" is not an answer to a ' +
                        'statutory request.' };
    }

    var sh = dpdp_requestSheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    for (var i = 1; i < (data ? data.length : 0); i++) {
      if (dpdp_str_(data[i][m['Request_ID']]).toUpperCase() !== want) continue;
      sh.getRange(i + 1, m['Status'] + 1).setValue('CLOSED');
      sh.getRange(i + 1, m['Handled_By'] + 1).setValue(actor.username);
      sh.getRange(i + 1, m['Closed_At'] + 1).setValue(dpdp_now_());
      sh.getRange(i + 1, m['Outcome'] + 1).setValue(text);
      dc_invalidate_(DPDP_CFG.REQUESTS);
      SpreadsheetApp.flush();
      return { success: true, message: 'Request ' + want + ' closed.' };
    }
    return { success: false, message: 'No request with the id ' + requestId + '.' };
  } catch (err) {
    var m3 = String((err && err.message) || err);
    return { success: false, message: m3.replace('FORBIDDEN: ', '') };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * FRONTEND ENTRY. Everything held about one patient, as one object.
 *
 * s.11 gives the data principal the right to a summary of their personal
 * data AND of the identities it has been shared with. Assembling that by
 * hand across nine sheets is what makes a thirty-day deadline hard to meet,
 * so it is assembled here.
 *
 * Each section fails soft: a missing sheet reduces the export rather than
 * failing it, and says which section could not be read.
 */
function exportPatientData(patientId, sessionToken) {
  try {
    var actor = crescRequire_(sessionToken, ['patient.read', 'portal.self']);
    var want = dpdp_str_(patientId).toUpperCase();
    if (!want) return { success: false, message: 'No patient was named.' };
    if (actor.role === 'patient' && dpdp_str_(actor.username).toUpperCase() !== want) {
      return { success: false, message: 'You can only export your own record.' };
    }

    var out = { patientId: want, generatedAt: dpdp_fmt_(dpdp_now_()),
                sections: {}, problems: [] };

    function section(name, fn) {
      try { out.sections[name] = fn(); }
      catch (e) { out.problems.push(name + ': ' + e.message); }
    }

    section('demographics', function () {
      var p = pt_readProfile_(want);
      if (!p) throw new Error('no patient record');
      // The portal password is never in pt_readProfile_'s output and must
      // never be in an export either.
      return p;
    });

    section('consent', function () {
      var res = getConsentStatus(want, sessionToken);
      return res.success ? res.purposes : [];
    });

    // Every sheet that holds rows about a patient, and the column their id
    // is in. Listed rather than guessed: a scan of every sheet would export
    // rows that merely mention a similar string.
    var SOURCES = [
      { sheet: 'Appointments',        col: 'Patient_ID',  label: 'appointments' },
      { sheet: 'OP_Encounters',       col: 'Patient_ID',  label: 'consultations' },
      { sheet: 'IP_Admissions',       col: 'Patient_ID',  label: 'admissions' },
      { sheet: 'IP_CaseSheets_DB',    col: 'Patient_ID',  label: 'caseSheets' },
      { sheet: 'IP_Timeline_DB',      col: 'Patient_ID',  label: 'wardNotes' },
      { sheet: 'LAB_ORDERS',          col: 'PatientID',   label: 'labOrders' },
      { sheet: 'LAB_BILLING',         col: 'PatientID',   label: 'labBills' },
      { sheet: 'Pharmacy_Invoices',   col: 'Patient_ID',  label: 'pharmacyBills' },
      { sheet: 'Hospital_Invoices',   col: 'Patient_ID',  label: 'hospitalBills' },
      { sheet: 'OP_Referrals',        col: 'Patient_ID',  label: 'referrals' },
      { sheet: 'Insurance_Claims_Ledger', col: 'Patient_Corp_Name', label: 'insuranceClaims' }
    ];

    SOURCES.forEach(function (src) {
      section(src.label, function () {
        var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(src.sheet);
        if (!sh) return [];
        var values = dc_sheetValues_(sh);
        if (!values || values.length < 2) return [];
        var idx = -1;
        for (var c = 0; c < values[0].length; c++) {
          if (dpdp_str_(values[0][c]) === src.col) { idx = c; break; }
        }
        if (idx === -1) return [];
        var rows = [];
        for (var i = 1; i < values.length && rows.length < 500; i++) {
          if (dpdp_str_(values[i][idx]).toUpperCase() !== want) continue;
          var obj = {};
          for (var j = 0; j < values[0].length; j++) {
            var h = dpdp_str_(values[0][j]);
            if (h) obj[h] = values[i][j];
          }
          rows.push(obj);
        }
        return rows;
      });
    });

    // s.11(1)(b): who it has been shared with. Guessed from the register of
    // documents actually put on Drive plus the consent decisions.
    section('disclosures', function () {
      var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DPDP_CFG.SHARED);
      if (!sh) return [];
      var values = dc_sheetValues_(sh);
      if (!values || values.length < 2) return [];
      var m = dc_headerMap_(sh);
      var rows = [];
      for (var i = 1; i < values.length; i++) {
        if (dpdp_str_(values[i][m['Patient_ID']]).toUpperCase() !== want) continue;
        rows.push({
          document: dpdp_str_(values[i][m['File_Name']]),
          type: dpdp_str_(values[i][m['Doc_Type']]),
          sharedAt: dpdp_str_(values[i][m['Shared_At']]),
          sharedBy: dpdp_str_(values[i][m['Shared_By']]),
          status: dpdp_str_(values[i][m['Status']])
        });
      }
      return rows;
    });

    try {
      logAudit_({ username: actor.username, role: actor.role },
                'DPDP_DATA_EXPORTED', 'Patient', want,
                { sections: Object.keys(out.sections).length });
    } catch (e) {}

    return { success: true, data: out, message: 'Export assembled.' };

  } catch (err) {
    var m = String((err && err.message) || err);
    return { success: false, message: m.replace('FORBIDDEN: ', '') };
  }
}

// ---------------------------------------------------------------------------
// SECTION 8(7) — RETENTION, AND THE LINKS NOBODY REVOKED
// ---------------------------------------------------------------------------

/**
 * Registers a file that has just been shared on Drive.
 *
 * WHY THIS MATTERS MORE THAN IT LOOKS. Five places in this project do
 *
 *     file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, ...)
 *
 * to hand a patient a lab report, a prescription or an invoice over
 * WhatsApp. Those files carry the patient's name, ID, diagnoses and results,
 * the link never expires, nobody records that it exists, and a WhatsApp
 * message is forwarded. There is no way to un-share what cannot be listed.
 *
 * Called by each of those five sites, this makes the set finite:
 * dpdpExpireSharedLinks() can then revoke the ones past their window.
 */
function dpdpRegisterSharedFile(file, docType, patientId, sharedBy) {
  try {
    if (!file) return '';
    var sh = dpdp_sharedSheet_();
    var now = dpdp_now_();
    var expires = new Date(now.getTime() + DPDP_CFG.RETENTION_DAYS.SHARED_LINK * 86400000);
    var id = 'SHR-' + Utilities.formatDate(now, DPDP_CFG.TZ, 'yyMMdd-HHmmss') + '-' +
             Utilities.getUuid().substring(0, 4).toUpperCase();
    sh.appendRow([
      id, file.getId(), file.getName(), dpdp_str_(docType),
      dpdp_str_(patientId).toUpperCase(), now, dpdp_str_(sharedBy) || 'SYSTEM',
      expires, 'SHARED', '', file.getUrl()
    ]);
    dc_invalidate_(DPDP_CFG.SHARED);
    return id;
  } catch (e) {
    Logger.log('dpdpRegisterSharedFile: ' + e.message);
    return '';
  }
}

/**
 * Revokes public access to registered files past their window.
 *
 * Run on a time-driven trigger, daily. Reports what it did; a file already
 * deleted or already private is counted, not treated as a failure.
 *
 * @param {boolean} [dryRun]  true to report without changing anything
 */
function dpdpExpireSharedLinks(dryRun) {
  var sh = dpdp_sharedSheet_();
  var m = dc_headerMap_(sh);
  var data = dc_sheetValues_(sh);
  if (!data || data.length < 2) return 'Nothing in the shared-document register.';

  var now = Date.now(), revoked = 0, gone = 0, kept = 0, problems = [];

  for (var i = 1; i < data.length; i++) {
    if (dpdp_str_(data[i][m['Status']]).toUpperCase() !== 'SHARED') continue;
    var expires = data[i][m['Expires_At']];
    var when = (typeof cresc_parseDate_ === 'function') ? cresc_parseDate_(expires) : new Date(expires);
    if (!when || isNaN(when.getTime()) || when.getTime() > now) { kept++; continue; }

    if (dryRun) { revoked++; continue; }

    try {
      var f = DriveApp.getFileById(dpdp_str_(data[i][m['File_ID']]));
      f.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
      sh.getRange(i + 1, m['Status'] + 1).setValue('REVOKED');
      sh.getRange(i + 1, m['Revoked_At'] + 1).setValue(dpdp_now_());
      revoked++;
    } catch (e) {
      // A file the user deleted is not a failure; it is the same outcome.
      sh.getRange(i + 1, m['Status'] + 1).setValue('GONE');
      sh.getRange(i + 1, m['Revoked_At'] + 1).setValue(dpdp_now_());
      gone++;
    }
  }
  dc_invalidate_(DPDP_CFG.SHARED);

  var msg = (dryRun ? 'DRY RUN — ' : '') +
    revoked + ' link(s) ' + (dryRun ? 'would be' : '') + ' revoked, ' +
    gone + ' file(s) already gone, ' + kept + ' still within their window.' +
    (problems.length ? '\n' + problems.join('\n') : '');
  Logger.log(msg);
  return msg;
}

/**
 * REPORTS what is past its retention period. NEVER DELETES.
 *
 * Erasure under s.8(7) is a decision with legal consequences — a medico-legal
 * case, an insurance dispute or a consumer complaint can require a record
 * years after the ordinary period — so this produces a list for a human to
 * act on. A retention sweep that deletes on its own is how a clinic loses
 * the record it is about to be asked for.
 */
function dpdpRetentionReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = Date.now();
  var lines = ['DPDP retention report — ' + dpdp_fmt_(dpdp_now_()), ''];

  var TARGETS = [
    { sheet: 'Sessions',      dateCol: 'Issued_At',  days: DPDP_CFG.RETENTION_DAYS.SESSION,
      what: 'sign-in sessions' },
    { sheet: 'Audit_Log',     dateCol: 'Timestamp',  days: DPDP_CFG.RETENTION_DAYS.AUDIT,
      what: 'audit entries' },
    { sheet: 'OP_Encounters', dateCol: 'Timestamp',  days: DPDP_CFG.RETENTION_DAYS.CLINICAL,
      what: 'consultations' },
    { sheet: 'Appointments',  dateCol: 'Timestamp',  days: DPDP_CFG.RETENTION_DAYS.CLINICAL,
      what: 'appointments' },
    { sheet: DPDP_CFG.SHARED, dateCol: 'Shared_At',  days: DPDP_CFG.RETENTION_DAYS.SHARED_LINK,
      what: 'shared document links' }
  ];

  TARGETS.forEach(function (t) {
    var sh = ss.getSheetByName(t.sheet);
    if (!sh) { lines.push('  ' + t.sheet + ': absent.'); return; }
    var values = dc_sheetValues_(sh);
    if (!values || values.length < 2) { lines.push('  ' + t.sheet + ': empty.'); return; }

    var idx = -1;
    for (var c = 0; c < values[0].length; c++) {
      if (dpdp_str_(values[0][c]) === t.dateCol) { idx = c; break; }
    }
    if (idx === -1) { lines.push('  ' + t.sheet + ': no "' + t.dateCol + '" column.'); return; }

    var cutoff = now - t.days * 86400000;
    var past = 0, unreadable = 0;
    for (var i = 1; i < values.length; i++) {
      var d = (typeof cresc_parseDate_ === 'function')
        ? cresc_parseDate_(values[i][idx]) : new Date(values[i][idx]);
      if (!d || isNaN(d.getTime())) { unreadable++; continue; }
      if (d.getTime() < cutoff) past++;
    }
    lines.push('  ' + t.sheet + ': ' + past + ' of ' + (values.length - 1) + ' ' + t.what +
               ' are older than ' + t.days + ' days' +
               (unreadable ? ' (' + unreadable + ' with an unreadable date)' : '') + '.');
  });

  lines.push('');
  lines.push('Nothing has been deleted. Erasure is a decision with legal');
  lines.push('consequences — a medico-legal case or an insurance dispute can');
  lines.push('require a record years later — so act on this list deliberately.');
  lines.push('Clinical records: NMC Ethics Regulations require 3 years from the');
  lines.push('last entry; inpatient and medico-legal records are usually kept');
  lines.push('longer. Financial records: 8 years under income tax practice.');

  var report = lines.join('\n');
  Logger.log(report);
  return report;
}

// ---------------------------------------------------------------------------
// THE SELF-AUDIT
// ---------------------------------------------------------------------------

/**
 * Where this deployment actually stands. Run it from the script editor.
 *
 * Checks the things that are checkable from inside the script: the
 * deployment settings, the registers, the notice, the officer, the password
 * storage, and the shared-link backlog. The rest is in
 * docs/DPDP_READINESS.md, because it is organisational rather than
 * technical and no function can assert it.
 */
function dpdpReadinessCheck() {
  var findings = [];
  function add(severity, area, text, fix) {
    findings.push({ severity: severity, area: area, text: text, fix: fix });
  }

  // --- deployment ---------------------------------------------------------
  // The single highest-risk setting in the project, and it is not in code —
  // which is why this cannot be checked, only reminded about. appsscript.json
  // in the repository now says access "ANYONE" (Google sign-in required), but
  // the manifest only takes effect on a NEW DEPLOYMENT: an /exec URL that was
  // published before the change keeps its old setting until it is redeployed.
  add('HIGH', 'Deployment',
      'The manifest asks for access "ANYONE" and executeAs USER_DEPLOYING, so ' +
      'every google.script.run endpoint still runs with the owner’s full ' +
      'spreadsheet access — now for signed-in callers only, and only if this ' +
      'deployment was published AFTER the manifest changed.',
      'Deploy > Manage deployments > edit > New version. Then open the /exec ' +
      'URL in a private window: if it answers without asking you to sign in, ' +
      'the old ANYONE_ANONYMOUS deployment is still live. Treat any period it ' +
      'was anonymous as potentially breached (s.8(6)).');

  // --- registers ----------------------------------------------------------
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  [DPDP_CFG.CONSENT, DPDP_CFG.REQUESTS, DPDP_CFG.SHARED].forEach(function (name) {
    if (!ss.getSheetByName(name)) {
      add('HIGH', 'Registers', name + ' does not exist.', 'Run dpdpSetup().');
    }
  });

  // --- grievance officer (s.13) ------------------------------------------
  if (!dpdp_officer_()) {
    add('HIGH', 'Section 13',
        'No grievance officer is named, so the notice cannot tell a patient ' +
        'whom to complain to.',
        'dpdpSetGrievanceOfficer("name", "email", "phone")');
  }

  // --- consent coverage ---------------------------------------------------
  try {
    var pts = ss.getSheetByName('Patients');
    var consented = {};
    var cs = ss.getSheetByName(DPDP_CFG.CONSENT);
    if (cs) {
      var cv = dc_sheetValues_(cs);
      var cm = dc_headerMap_(cs);
      for (var i = 1; i < (cv ? cv.length : 0); i++) {
        consented[dpdp_str_(cv[i][cm['Patient_ID']]).toUpperCase()] = true;
      }
    }
    if (pts) {
      var total = Math.max(0, pts.getLastRow() - 1);
      var have = Object.keys(consented).length;
      if (have < total) {
        add(have === 0 ? 'HIGH' : 'MEDIUM', 'Section 6',
            (total - have) + ' of ' + total + ' patients have no consent record ' +
            'at all.',
            'Capture consent at registration and backfill at the next visit. ' +
            'recordConsent() is the endpoint.');
      }
    }
  } catch (e) { /* the check itself must not fail the report */ }

  // --- password storage ---------------------------------------------------
  // Counted rather than sampled: one hashed row at the top of the sheet says
  // nothing about the two hundred below it, and it was a sampled check that
  // made this look fixed the first time.
  try {
    [['Users', 'staff logins', 'CRITICAL'],
     ['Patients', 'patient portal logins', 'HIGH']].forEach(function (t) {
      var sh = ss.getSheetByName(t[0]);
      if (!sh || sh.getLastRow() < 2) return;
      var col = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getDisplayValues();
      var plain = 0;
      col.forEach(function (r) {
        var v = dpdp_str_(r[0]);
        if (v && !(typeof crescPwdIsHashed_ === 'function' && crescPwdIsHashed_(v))) plain++;
      });
      if (plain) {
        add(t[2], 'Section 8(5)',
            plain + ' of ' + col.length + ' ' + t[1] + ' are still stored in ' +
            'PLAIN TEXT on the ' + t[0] + ' sheet. Everyone who can open the ' +
            'spreadsheet can read them, and people reuse passwords. Sign-in ' +
            'refuses them, so those accounts cannot be used at all until they ' +
            'are reset.',
            'Run crescMigrateCredentials() from the script editor: it issues a ' +
            'fresh random password per account, stores only the digest, prints ' +
            'the list once, and forces a change at first sign-in.');
      }
    });
  } catch (e) { /* ditto */ }

  // --- shared links -------------------------------------------------------
  try {
    var shSh = ss.getSheetByName(DPDP_CFG.SHARED);
    if (shSh) {
      var sv = dc_sheetValues_(shSh);
      var sm = dc_headerMap_(shSh);
      var open = 0, overdue = 0, nowMs = Date.now();
      for (var j = 1; j < (sv ? sv.length : 0); j++) {
        if (dpdp_str_(sv[j][sm['Status']]).toUpperCase() !== 'SHARED') continue;
        open++;
        var exp = (typeof cresc_parseDate_ === 'function')
          ? cresc_parseDate_(sv[j][sm['Expires_At']]) : new Date(sv[j][sm['Expires_At']]);
        if (exp && !isNaN(exp.getTime()) && exp.getTime() < nowMs) overdue++;
      }
      if (overdue) {
        add('HIGH', 'Section 8(7)',
            overdue + ' patient document(s) are still publicly readable by link ' +
            'past their retention window.',
            'Run dpdpExpireSharedLinks() and add a daily time-driven trigger for it.');
      }
      if (!open && !overdue) {
        add('MEDIUM', 'Section 8(5)',
            'The shared-document register is empty. If documents are still being ' +
            'put on Drive with ANYONE_WITH_LINK and not registered here, they ' +
            'cannot be revoked because nothing lists them.',
            'Confirm every setSharing() call also calls dpdpRegisterSharedFile().');
      }
    }
  } catch (e) { /* ditto */ }

  // --- open requests ------------------------------------------------------
  try {
    var rq = ss.getSheetByName(DPDP_CFG.REQUESTS);
    if (rq) {
      var rv = dc_sheetValues_(rq);
      var rm = dc_headerMap_(rq);
      var late = 0, nowMs2 = Date.now();
      for (var k = 1; k < (rv ? rv.length : 0); k++) {
        if (dpdp_str_(rv[k][rm['Status']]).toUpperCase() !== 'OPEN') continue;
        var due = (typeof cresc_parseDate_ === 'function')
          ? cresc_parseDate_(rv[k][rm['Due_By']]) : new Date(rv[k][rm['Due_By']]);
        if (due && !isNaN(due.getTime()) && due.getTime() < nowMs2) late++;
      }
      if (late) {
        add('HIGH', 'Sections 11-13',
            late + ' data-principal request(s) are past their response date.',
            'listDPDPRequests() shows them, overdue first.');
      }
    }
  } catch (e) { /* ditto */ }

  var order = { CRITICAL: 0, HIGH: 1, MEDIUM: 2, LOW: 3 };
  findings.sort(function (a, b) { return order[a.severity] - order[b.severity]; });

  var lines = ['DPDP readiness — ' + dpdp_fmt_(dpdp_now_()),
               findings.length + ' finding(s), worst first.', ''];
  findings.forEach(function (f, i) {
    lines.push((i + 1) + '. [' + f.severity + '] ' + f.area);
    lines.push('   ' + f.text);
    lines.push('   FIX: ' + f.fix);
    lines.push('');
  });
  lines.push('This checks only what is checkable from inside the script.');
  lines.push('The organisational half — the Consent Notice, the retention');
  lines.push('schedule, the breach procedure, the processor agreements — is in');
  lines.push('docs/DPDP_READINESS.md.');

  var report = lines.join('\n');
  Logger.log(report);
  return { success: true, findings: findings, report: report };
}
