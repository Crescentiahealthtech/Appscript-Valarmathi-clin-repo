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
  NOMINEES:  'Nomination_Register',
  NOTICE_PROP: 'DPDP_NOTICE_VERSION',
  OFFICER_PROP: 'DPDP_GRIEVANCE_OFFICER',
  TZ: 'Asia/Kolkata',

  // Statutory-ish clocks. The Act sets the obligation; the Rules set the
  // period. These are the clinic's committed response times, shown to the
  // patient and used to flag an overdue request.
  RESPONSE_DAYS: { ACCESS: 30, CORRECTION: 30, ERASURE: 30, GRIEVANCE: 30 },

  /**
   * How the clinic satisfied itself that the person asking is the person the
   * data is about — finding M5.
   *
   * Answering an access request to the wrong person is itself a disclosure,
   * and it is the disclosure an attacker will choose because it arrives
   * gift-wrapped in a statutory right. Every request row records which of
   * these was used; UNVERIFIED exists so that a request can be TAKEN before
   * it is verified, and so that the gap is visible rather than assumed away.
   */
  VERIFICATION: {
    IN_PERSON_ID:  'Photo ID checked at the clinic, in person',
    REGISTERED_MOBILE: 'Call-back to the mobile number already on the record',
    PORTAL_SESSION: 'Asked through the patient portal, signed in',
    GUARDIAN_ID:   'Parent or guardian identified in person (s.9)',
    NOMINEE_PROOF: 'Nominee under s.14, with the death or incapacity evidenced',
    UNVERIFIED:    'NOT YET VERIFIED — do not answer this request until it is'
  },

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
    'Closed_At', 'Outcome',
    // Finding M5. Answering an access request to the wrong person is a
    // disclosure dressed as compliance, so how the requester was identified
    // is part of the record, not part of the memory of whoever was at the desk.
    'Verification_Method', 'Verified_By', 'Verified_At', 'Channel'
  ]);
}

function dpdp_sharedSheet_() {
  return dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), DPDP_CFG.SHARED, [
    'Entry_ID', 'File_ID', 'File_Name', 'Doc_Type', 'Patient_ID',
    'Shared_At', 'Shared_By', 'Expires_At', 'Status', 'Revoked_At', 'Link'
  ]);
}

/**
 * s.14 — the nominee.
 *
 * A separate register rather than a purpose on the consent sheet, because a
 * nomination is not a consent: it is an instruction about who may exercise
 * these rights if the patient dies or becomes incapable of exercising them
 * themselves. It is append-only for the same reason consent is — the
 * question "who was nominated on the day they died" has exactly one right
 * answer and it is not "whoever the row says now".
 *
 * The emergency contact already on the patient record is NOT this. That was
 * collected to reach somebody in a hurry; it was never given for this
 * purpose and the patient never chose it for this purpose.
 */
function dpdp_nomineeSheet_() {
  return dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), DPDP_CFG.NOMINEES, [
    'Entry_ID', 'Patient_ID', 'Patient_Name', 'Nominee_Name', 'Relationship',
    'Nominee_Contact', 'Scope', 'Notice_Version', 'Recorded_By', 'Recorded_At',
    'Revoked_At', 'Revoked_Reason', 'Notes'
  ]);
}

/** ONE-OFF. Creates every register. Safe to re-run. */
function dpdpSetup() {
  dpdp_consentSheet_();
  dpdp_requestSheet_();
  dpdp_sharedSheet_();
  dpdp_nomineeSheet_();
  if (typeof dpdp_grantSheet_ === 'function') dpdp_grantSheet_();
  if (typeof dpdp_breachSheet_ === 'function') dpdp_breachSheet_();
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty(DPDP_CFG.NOTICE_PROP)) {
    props.setProperty(DPDP_CFG.NOTICE_PROP, 'v1-' +
      Utilities.formatDate(new Date(), DPDP_CFG.TZ, 'yyyyMMdd'));
  }
  return 'DPDP registers ready: ' + [DPDP_CFG.CONSENT, DPDP_CFG.REQUESTS,
           DPDP_CFG.SHARED, DPDP_CFG.NOMINEES, 'Document_Grants',
           'Breach_Register'].join(', ') + '.\n' +
         'Now, in order:\n' +
         '  1. dpdpSetGrievanceOfficer("Dr. …", "officer@clinic.in", "+91 …")  — s.13\n' +
         '  2. dpdpInstallTriggers()        — the daily and weekly jobs\n' +
         '  3. crescMigrateCredentials()    — hash every password, force a reset\n' +
         '  4. dpdpReadinessCheck()         — where you stand afterwards';
}

/**
 * s.13 — a Data Fiduciary must publish the contact of the person who
 * answers a data principal's grievance. Not having one named is itself a
 * finding, so this is stored where the notice can read it.
 */
function dpdpSetGrievanceOfficer(name, email, phone) {
  crescEditorOnly_('dpdpSetGrievanceOfficer', ['dpdp.manage', 'admin.config']);
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
    ? cresc_clinic_() : { name: 'this clinic', address: '', phone: '', email: '',
                          website: '' };
  var officer = dpdp_officer_();

  return {
    success: true,
    version: dpdp_noticeVersion_(),
    fiduciary: clinic.name,
    // s.5(1) requires the notice to identify the Data Fiduciary. A patient
    // who has put the paper down has to be able to find the clinic again,
    // so the notice carries the clinic's own web address — not the
    // deployment URL this page happens to be served from.
    contact: { address: clinic.address, phone: clinic.phone, email: clinic.email,
               website: clinic.website || '' },
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

    // s.6(6): withdrawal has to be as easy as giving, and the patient portal
    // is where a patient can do either without telephoning the clinic during
    // working hours. A patient reaches this with 'portal.self' and may only
    // ever act on their own record — checked immediately below, because
    // 'portal.self' says "the portal", not "any patient in it".
    var actor = crescRequire_(sessionToken,
                              ['patient.register', 'patient.write', 'portal.self']);

    var patientId = dpdp_str_(payload.patientId).toUpperCase();
    if (!patientId) return { success: false, message: 'No patient was named.' };
    if (actor.role === 'patient' && dpdp_str_(actor.username).toUpperCase() !== patientId) {
      return { success: false, message: 'You can only change your own consent.' };
    }

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
    // A patient may read their own consent record, and nobody else's.
    var actor = crescRequire_(sessionToken, ['patient.read', 'portal.self']);
    var want = dpdp_str_(patientId).toUpperCase();
    if (!want) return { success: false, message: 'No patient was named.' };
    if (actor.role === 'patient' && dpdp_str_(actor.username).toUpperCase() !== want) {
      return { success: false, message: 'You can only see your own consent record.' };
    }

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

    // s.11-13 with finding M5. A patient asking through the portal has already
    // proved who they are — their session IS the verification, and asking the
    // desk to type that again invites them to type something else. Anyone
    // else has to say how the person in front of them was identified, and
    // UNVERIFIED is allowed so that the request can be taken now and the gap
    // stays visible instead of being assumed away.
    var method = dpdp_str_(payload.verification).toUpperCase();
    if (actor.role === 'patient') method = 'PORTAL_SESSION';
    if (!method) method = 'UNVERIFIED';
    if (!DPDP_CFG.VERIFICATION.hasOwnProperty(method)) {
      return { success: false,
               message: 'Verification must be one of: ' +
                        Object.keys(DPDP_CFG.VERIFICATION).join(', ') + '.' };
    }

    var profile = null;
    try { profile = pt_readProfile_(patientId); } catch (e) { profile = null; }

    var now = dpdp_now_();
    var due = new Date(now.getTime() + DPDP_CFG.RESPONSE_DAYS[type] * 86400000);
    var id = 'DPR-' + Utilities.formatDate(now, DPDP_CFG.TZ, 'yyMMdd-HHmmss') + '-' +
             Utilities.getUuid().substring(0, 4).toUpperCase();

    dpdp_requestSheet_().appendRow([
      id, patientId, profile ? profile.name : '', type, details,
      actor.username, now, due, 'OPEN', '', '', '',
      method, method === 'UNVERIFIED' ? '' : actor.username,
      method === 'UNVERIFIED' ? '' : now,
      dpdp_str_(payload.channel) || (actor.role === 'patient' ? 'PORTAL' : 'DESK')
    ]);
    dc_invalidate_(DPDP_CFG.REQUESTS);
    SpreadsheetApp.flush();

    try {
      logAudit_({ username: actor.username, role: actor.role },
                'DPDP_REQUEST_RAISED', 'Patient', patientId, { type: type, id: id });
    } catch (e) {}

    return { success: true, requestId: id, dueBy: dpdp_fmt_(due),
             verification: method,
             message: type.charAt(0) + type.slice(1).toLowerCase() +
                      ' request ' + id + ' logged. It is due by ' +
                      Utilities.formatDate(due, DPDP_CFG.TZ, 'dd-MMM-yyyy') + '.' +
                      (method === 'UNVERIFIED'
                        ? ' IT IS NOT VERIFIED: confirm who is asking before you ' +
                          'answer it, and record how (dpdpVerifyRequester).'
                        : '') };

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

      // An access request answered to the wrong person is a disclosure, and
      // the statutory right is what an attacker would use to ask for it. So a
      // request nobody verified cannot be closed as answered.
      var how = dpdp_str_(data[i][m['Verification_Method']]).toUpperCase();
      if (!how || how === 'UNVERIFIED') {
        return { success: false, code: 'UNVERIFIED',
                 message: 'This request has not been verified. Confirm the person ' +
                          'asking is the person the data is about, record how with ' +
                          'dpdpVerifyRequester("' + want + '", METHOD, token), and ' +
                          'then close it.' };
      }
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
 * FRONTEND ENTRY. Records HOW the clinic satisfied itself that the person who
 * asked is the person the data is about — finding M5.
 *
 * Separate from raising the request on purpose: the request usually arrives
 * before the proof does. A call back to the number already on the record is
 * the cheapest method the clinic already has, and it is the one thing the
 * person impersonating a patient cannot arrange.
 */
function dpdpVerifyRequester(requestId, method, note, sessionToken) {
  try {
    var actor = crescRequire_(sessionToken, ['dpdp.manage', 'patient.read']);
    var how = dpdp_str_(method).toUpperCase();
    if (!DPDP_CFG.VERIFICATION.hasOwnProperty(how) || how === 'UNVERIFIED') {
      return { success: false,
               message: 'Verification must be one of: ' +
                        Object.keys(DPDP_CFG.VERIFICATION)
                          .filter(function (k) { return k !== 'UNVERIFIED'; })
                          .join(', ') + '.' };
    }

    var sh = dpdp_requestSheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    var want = dpdp_str_(requestId).toUpperCase();

    for (var i = 1; i < (data ? data.length : 0); i++) {
      if (dpdp_str_(data[i][m['Request_ID']]).toUpperCase() !== want) continue;
      sh.getRange(i + 1, m['Verification_Method'] + 1).setValue(how);
      sh.getRange(i + 1, m['Verified_By'] + 1).setValue(actor.username);
      sh.getRange(i + 1, m['Verified_At'] + 1).setValue(dpdp_now_());
      if (dpdp_str_(note)) {
        sh.getRange(i + 1, m['Details'] + 1).setValue(
          dpdp_str_(data[i][m['Details']]) + '\n[verification: ' + dpdp_str_(note) + ']');
      }
      dc_invalidate_(DPDP_CFG.REQUESTS);
      SpreadsheetApp.flush();
      try {
        logAudit_({ username: actor.username, role: actor.role },
                  'DPDP_REQUESTER_VERIFIED', 'Request', want, { method: how });
      } catch (e) {}
      return { success: true, message: 'Recorded: ' + DPDP_CFG.VERIFICATION[how] + '.' };
    }
    return { success: false, message: 'No request with the id ' + requestId + '.' };
  } catch (err) {
    return { success: false, message: String(err.message || err).replace('FORBIDDEN: ', '') };
  }
}

/**
 * PUBLIC BY DESIGN. A data principal asks without holding a staff login.
 *
 * s.11-13 give the right to a person, not to a person with a password.
 * Requiring somebody to come to the desk to ask for their own data — or to
 * get past a login they may have forgotten — is a way of receiving fewer
 * requests, not of answering them.
 *
 * WHAT MAKES THIS SAFE TO LEAVE OPEN:
 *   * it WRITES a request and returns nothing about anybody;
 *   * every row lands UNVERIFIED, and closeDPDPRequest() refuses to answer an
 *     unverified request, so the form cannot be used to have data posted to
 *     a stranger;
 *   * it is rate-limited per browser, because an open write endpoint that is
 *     not rate-limited is a way to fill a spreadsheet.
 */
function dpdpSubmitPublicRequest(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    payload = payload || {};

    var type = dpdp_str_(payload.type).toUpperCase();
    if (!DPDP_CFG.RESPONSE_DAYS.hasOwnProperty(type)) {
      return { success: false,
               message: 'Choose what you are asking for: ' +
                        Object.keys(DPDP_CFG.RESPONSE_DAYS).join(', ') + '.' };
    }
    var name = dpdp_str_(payload.name);
    var contact = dpdp_str_(payload.contact);
    var details = dpdp_str_(payload.details);
    if (name.length < 2 || contact.length < 6) {
      return { success: false,
               message: 'We need your name and a phone number or email address to ' +
                        'reply to — and to check that it is you asking.' };
    }
    if (details.length < 10) {
      return { success: false, message: 'Tell us in a sentence what you would like.' };
    }

    // Crude but sufficient: four requests per hour from one browser. The cache
    // key is the best identifier a web app gets — Apps Script does not expose
    // the caller's IP address — so this slows a person down rather than
    // stopping a determined script, and is said plainly rather than dressed up.
    try {
      var cache = CacheService.getScriptCache();
      var seen = parseInt(cache.get('DPDP_PUB_' + dpdp_str_(payload.formId)), 10) || 0;
      if (seen >= 4) {
        return { success: false,
                 message: 'Several requests have already been sent from this page. ' +
                          'Please call the clinic instead.' };
      }
      cache.put('DPDP_PUB_' + dpdp_str_(payload.formId), String(seen + 1), 3600);
    } catch (e) { /* the cache is best effort; never block a statutory right on it */ }

    var now = dpdp_now_();
    var due = new Date(now.getTime() + DPDP_CFG.RESPONSE_DAYS[type] * 86400000);
    var id = 'DPR-' + Utilities.formatDate(now, DPDP_CFG.TZ, 'yyMMdd-HHmmss') + '-' +
             Utilities.getUuid().substring(0, 4).toUpperCase();

    // The patient ID is whatever they typed, and it is NOT trusted — it is a
    // lead for the person who verifies them, not an identification.
    dpdp_requestSheet_().appendRow([
      id, dpdp_str_(payload.patientId).toUpperCase(), name, type,
      details + '\n[submitted through the public form by ' + name + ', ' + contact + ']',
      'PUBLIC_FORM', now, due, 'OPEN', '', '', '',
      'UNVERIFIED', '', '', 'PUBLIC_FORM'
    ]);
    dc_invalidate_(DPDP_CFG.REQUESTS);
    SpreadsheetApp.flush();

    try {
      logAudit_({ username: 'PUBLIC_FORM', role: '' },
                'DPDP_REQUEST_RAISED', 'Request', id, { type: type, channel: 'PUBLIC' });
    } catch (e) {}

    var officer = dpdp_officer_();
    if (officer && officer.email) {
      try {
        MailApp.sendEmail({
          to: officer.email,
          subject: 'Data principal request ' + id + ' (' + type + ')',
          body: [
            'A request has come in through the public form.',
            '',
            'Reference:  ' + id,
            'Type:       ' + type,
            'From:       ' + name + ' (' + contact + ')',
            'Patient ID they gave: ' + (dpdp_str_(payload.patientId) || '(none)'),
            'Due by:     ' + dpdp_fmt_(due),
            '',
            details,
            '',
            'IT IS NOT VERIFIED. Confirm who is asking before answering — a call ' +
            'back to the number already on their record is enough and is the one ' +
            'thing somebody impersonating them cannot arrange. Record it with ' +
            'dpdpVerifyRequester().'
          ].join('\n')
        });
      } catch (e) { /* the row is the record; the email is a convenience */ }
    }

    return { success: true, requestId: id,
             dueBy: Utilities.formatDate(due, DPDP_CFG.TZ, 'dd-MMM-yyyy'),
             message: 'Your request has been logged as ' + id + '. We will reply by ' +
                      Utilities.formatDate(due, DPDP_CFG.TZ, 'dd MMMM yyyy') + '. ' +
                      'We will contact you first to check it is you asking.' };
  } catch (err) {
    return { success: false, message: 'The request could not be sent: ' + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ---------------------------------------------------------------------------
// SECTION 14 — NOMINATION
// ---------------------------------------------------------------------------

/**
 * FRONTEND ENTRY. Records who the patient nominates to exercise their rights
 * if they die or become incapable of exercising them — finding M4.
 *
 * NOT the emergency contact. That was collected so somebody could be reached
 * in a hurry; it was never given for this purpose, the patient never chose it
 * for this purpose, and s.14 requires the patient's own act. Same affirmative
 * action as a consent, recorded the same way, against the same notice
 * version.
 */
function recordNomination(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    payload = payload || {};
    var actor = crescRequire_(sessionToken, ['patient.write', 'patient.register', 'portal.self']);

    var patientId = dpdp_str_(payload.patientId).toUpperCase();
    if (!patientId) return { success: false, message: 'No patient was named.' };
    if (actor.role === 'patient' && dpdp_str_(actor.username).toUpperCase() !== patientId) {
      return { success: false, message: 'You can only nominate on your own record.' };
    }

    var nominee = dpdp_str_(payload.nomineeName);
    var relation = dpdp_str_(payload.relationship);
    var contact = dpdp_str_(payload.contact);
    if (nominee.length < 2 || !relation || contact.length < 6) {
      return { success: false,
               message: 'A nomination needs the nominee\'s name, their relationship ' +
                        'to the patient, and a way to reach them.' };
    }

    var profile = null;
    try { profile = pt_readProfile_(patientId); } catch (e) { profile = null; }
    if (!profile) return { success: false, message: 'No patient with the ID ' + patientId + '.' };

    var now = dpdp_now_();
    var sh = dpdp_nomineeSheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);

    // A new nomination replaces the last one, and the old row stays, revoked.
    // "Who was nominated when they died" has one right answer and it is not
    // "whoever the current row says".
    for (var i = 1; i < (data ? data.length : 0); i++) {
      if (dpdp_str_(data[i][m['Patient_ID']]).toUpperCase() !== patientId) continue;
      if (dpdp_str_(data[i][m['Revoked_At']])) continue;
      sh.getRange(i + 1, m['Revoked_At'] + 1).setValue(now);
      sh.getRange(i + 1, m['Revoked_Reason'] + 1).setValue('Replaced by a later nomination');
    }

    sh.appendRow([
      'NOM-' + Utilities.formatDate(now, DPDP_CFG.TZ, 'yyMMdd-HHmmss') + '-' +
        Utilities.getUuid().substring(0, 4).toUpperCase(),
      patientId, profile.name, nominee, relation, contact,
      dpdp_str_(payload.scope) || 'ALL_RIGHTS',
      dpdp_noticeVersion_(), actor.username, now, '', '', dpdp_str_(payload.notes)
    ]);
    dc_invalidate_(DPDP_CFG.NOMINEES);
    SpreadsheetApp.flush();

    try {
      logAudit_({ username: actor.username, role: actor.role },
                'DPDP_NOMINATION_RECORDED', 'Patient', patientId, { relationship: relation });
    } catch (e) {}

    return { success: true,
             message: nominee + ' is recorded as ' + profile.name + '\'s nominee under ' +
                      'section 14. Tell the patient that a nomination can be changed or ' +
                      'withdrawn at any time.' };
  } catch (err) {
    return { success: false, message: String(err.message || err).replace('FORBIDDEN: ', '') };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/** FRONTEND ENTRY. The nomination in force, and the ones it replaced. */
function getNomination(patientId, sessionToken) {
  try {
    var actor = crescRequire_(sessionToken, ['patient.read', 'portal.self']);
    var want = dpdp_str_(patientId).toUpperCase();
    if (actor.role === 'patient' && dpdp_str_(actor.username).toUpperCase() !== want) {
      return { success: false, nomination: null, history: [],
               message: 'You can only see your own nomination.' };
    }
    var sh = dpdp_nomineeSheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    var current = null, history = [];

    for (var i = 1; i < (data ? data.length : 0); i++) {
      if (dpdp_str_(data[i][m['Patient_ID']]).toUpperCase() !== want) continue;
      var row = {
        entryId: dpdp_str_(data[i][m['Entry_ID']]),
        nominee: dpdp_str_(data[i][m['Nominee_Name']]),
        relationship: dpdp_str_(data[i][m['Relationship']]),
        contact: dpdp_str_(data[i][m['Nominee_Contact']]),
        scope: dpdp_str_(data[i][m['Scope']]),
        recordedAt: dpdp_str_(data[i][m['Recorded_At']]),
        recordedBy: dpdp_str_(data[i][m['Recorded_By']]),
        revokedAt: dpdp_str_(data[i][m['Revoked_At']])
      };
      if (row.revokedAt) history.push(row); else current = row;
    }
    return { success: true, patientId: want, nomination: current,
             history: history.reverse(), message: '' };
  } catch (err) {
    return { success: false, nomination: null, history: [],
             message: String(err.message || err).replace('FORBIDDEN: ', '') };
  }
}

/** FRONTEND ENTRY. Withdraws a nomination. As easy as making one. */
function revokeNomination(patientId, reason, sessionToken) {
  try {
    var actor = crescRequire_(sessionToken, ['patient.write', 'portal.self']);
    var want = dpdp_str_(patientId).toUpperCase();
    if (actor.role === 'patient' && dpdp_str_(actor.username).toUpperCase() !== want) {
      return { success: false, message: 'You can only change your own nomination.' };
    }
    var sh = dpdp_nomineeSheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    var n = 0;
    for (var i = 1; i < (data ? data.length : 0); i++) {
      if (dpdp_str_(data[i][m['Patient_ID']]).toUpperCase() !== want) continue;
      if (dpdp_str_(data[i][m['Revoked_At']])) continue;
      sh.getRange(i + 1, m['Revoked_At'] + 1).setValue(dpdp_now_());
      sh.getRange(i + 1, m['Revoked_Reason'] + 1).setValue(dpdp_str_(reason) || 'Withdrawn by the patient');
      n++;
    }
    dc_invalidate_(DPDP_CFG.NOMINEES);
    SpreadsheetApp.flush();
    try {
      logAudit_({ username: actor.username, role: actor.role },
                'DPDP_NOMINATION_REVOKED', 'Patient', want, {});
    } catch (e) {}
    return { success: n > 0,
             message: n ? 'Nomination withdrawn.' : 'There was no nomination on this record.' };
  } catch (err) {
    return { success: false, message: String(err.message || err).replace('FORBIDDEN: ', '') };
  }
}

/**
 * Records that somebody READ a clinical document or a patient record —
 * finding M2.
 *
 * logAudit_() was called on writes only, so "who looked at this patient's
 * file" — the question every incident actually asks — could not be answered
 * at all. Reads are far more numerous than writes, so this logs the ones that
 * matter: a whole document assembled, printed, exported or searched for, not
 * every cell fetched to paint a list.
 *
 * Never throws. An audit that can break the thing it audits is worse than no
 * audit: nobody should be unable to print a discharge summary because the log
 * sheet hit a quota.
 */
function dpdpLogRead_(actor, what, entityId, details) {
  try {
    var a = actor || {};
    logAudit_({ username: dpdp_str_(a.username) || '(unknown)',
                role: dpdp_str_(a.role), doctorId: dpdp_str_(a.doctorId) },
              'CLINICAL_DOCUMENT_READ', dpdp_str_(what) || 'Document',
              dpdp_str_(entityId), details || {});
  } catch (e) {
    Logger.log('dpdpLogRead_ failed: ' + e.message);
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

    // s.14 — the person the patient nominated to exercise these rights.
    section('nomination', function () {
      var res = getNomination(want, sessionToken);
      return res.success ? { current: res.nomination, previous: res.history } : null;
    });

    // s.11(1)(b) — every document link issued for this patient, whether it is
    // still live or not. This is the modern half of "who has my data been
    // given to"; the Shared_Documents section below is the legacy Drive half.
    section('documentLinks', function () {
      var res = dpdpListDocumentLinks(want, sessionToken);
      return res.success ? res.rows : [];
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
 * LEGACY. Registers a file shared on Drive with ANYONE_WITH_LINK.
 *
 * NOTHING CALLS THIS ANY MORE, and that is the point. The five places that
 * used to publish a report, a prescription or an invoice to the open web now
 * call dpdpIssueDocumentLink_() (DPDP_Documents.gs), which keeps the file
 * private and hands out an expiring key instead.
 *
 * It is kept for two reasons: dpdpExpireSharedLinks() still has to work
 * through the backlog of files published before that change, and a deployment
 * that adds another Drive-sharing integration later should register it here
 * rather than inventing a second register. If you find yourself calling this,
 * look at dpdpIssueDocumentLink_() first — publishing to Drive is almost
 * never the answer.
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
 * ADMIN. Erases the three columns that were collected for no stated purpose —
 * marital status, occupation and education — from every existing patient row.
 *
 * WHY THIS ONE DOES DELETE, WHEN THE RETENTION SWEEP DOES NOT. Everything the
 * retention report lists is a CLINICAL or FINANCIAL record: a medico-legal
 * case or an insurance dispute can require it years after its ordinary
 * period, so deleting it on a schedule is how a clinic loses the file it is
 * about to be asked for. These three fields are the opposite case. Nothing
 * reads them, no purpose was ever stated for them, and s.6(1) permits
 * collection only for a specified purpose — so keeping them has no upside to
 * weigh against, and erasing them is the plain s.8(7) answer.
 *
 * Dry run by default. Pass true to actually clear them.
 *
 * @param {boolean} [confirm]  false/omitted reports; true erases
 */
function dpdpEraseUnusedFields(confirm) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Patients');
  if (!sh || sh.getLastRow() < 2) return 'No patient rows.';

  var COLUMNS = [
    { col: 13, name: 'Marital_Status' },   // 1-based: M
    { col: 15, name: 'Occupation' },       // O
    { col: 16, name: 'Education' }         // P
  ];
  var rows = sh.getLastRow() - 1;
  var lines = ['Unused-field erasure' + (confirm ? '' : ' — DRY RUN') + ' — ' +
               dpdp_fmt_(dpdp_now_()), ''];
  var total = 0;

  COLUMNS.forEach(function (c) {
    if (sh.getLastColumn() < c.col) { lines.push('  ' + c.name + ': column absent.'); return; }
    var range = sh.getRange(2, c.col, rows, 1);
    var values = range.getValues();
    var filled = values.filter(function (r) { return dpdp_str_(r[0]) !== ''; }).length;
    total += filled;
    if (confirm && filled) range.clearContent();
    lines.push('  ' + c.name + ': ' + filled + ' value(s) ' +
               (confirm ? 'erased.' : 'would be erased.'));
  });

  lines.push('');
  lines.push(confirm
    ? total + ' value(s) erased from ' + rows + ' patient row(s). The columns are ' +
      'left in place so the sheet layout, which every column index in this ' +
      'project depends on, does not move.'
    : 'Nothing has been changed. Run dpdpEraseUnusedFields(true) to erase.');

  var report = lines.join('\n');
  Logger.log(report);
  if (confirm) {
    try {
      logAudit_({ username: 'ADMIN', role: '' }, 'DPDP_FIELDS_ERASED', 'Patients', 'ALL',
                { columns: COLUMNS.map(function (c) { return c.name; }), values: total });
    } catch (e) {}
  }
  return report;
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
// THE CONSOLE'S OWN ENDPOINTS
// ---------------------------------------------------------------------------

/**
 * FRONTEND ENTRY. Names the person a patient complains to — s.13.
 *
 * The editor version (dpdpSetGrievanceOfficer) stays, because this is the
 * kind of setting that gets changed once by whoever is holding the laptop.
 * This one exists so it can be changed by the clinic rather than by someone
 * who knows how to open Apps Script.
 */
function dpdpSaveGrievanceOfficer(payload, sessionToken) {
  try {
    var actor = crescRequire_(sessionToken, ['dpdp.manage', 'admin.config']);
    payload = payload || {};
    var msg = dpdpSetGrievanceOfficer(payload.name, payload.email, payload.phone);
    var ok = msg.indexOf('set:') !== -1;
    if (ok) {
      try {
        logAudit_({ username: actor.username, role: actor.role },
                  'DPDP_OFFICER_SET', 'Config', 'GRIEVANCE_OFFICER',
                  { name: dpdp_str_(payload.name) });
      } catch (e) {}
    }
    return { success: ok, message: msg };
  } catch (err) {
    return { success: false, message: String(err.message || err).replace('FORBIDDEN: ', '') };
  }
}

/**
 * FRONTEND ENTRY. Everything the privacy console shows on first paint, in one
 * round trip — the posture, the queue, the register and the officer.
 *
 * One call rather than six because this screen is opened by somebody checking
 * whether anything needs them today, and six sequential Apps Script round
 * trips is most of a minute.
 */
function dpdpConsoleSnapshot(sessionToken) {
  try {
    var actor = crescRequire_(sessionToken, ['dpdp.manage', 'admin.config', 'admin.audit']);

    // The FAST half only. dpdpConsoleHousekeeping() below carries the
    // readiness check, the audit scan and the trigger list, which is what
    // the screen used to block on.
    var out = { success: true, actor: actor.displayName || actor.username,
                officer: dpdp_officer_(), noticeVersion: dpdp_noticeVersion_(),
                requests: [], overdue: 0, breaches: [], unassessed: 0,
                message: '' };

    try {
      var rq = listDPDPRequests(sessionToken, { openOnly: false });
      if (rq.success) { out.requests = rq.rows; out.overdue = rq.overdue; }
    } catch (e) { out.message += 'Requests unavailable. '; }

    try {
      var br = dpdpListBreaches(sessionToken);
      if (br.success) { out.breaches = br.rows; out.unassessed = br.unassessed; }
    } catch (e) { out.message += 'Breach register unavailable. '; }

    return out;
  } catch (err) {
    return { success: false, requests: [], breaches: [], findings: [], anomalies: [],
             message: String(err.message || err).replace('FORBIDDEN: ', '') };
  }
}

/**
 * FRONTEND ENTRY. The slow half of the console, fetched after first paint.
 *
 * THE SCREEN USED TO WAIT FOR ALL OF THIS BEFORE SHOWING ANYTHING. One call
 * did five jobs: the request queue, the breach register, the full readiness
 * check, a seven-day scan of the audit log and the trigger list. The first
 * two are a couple of sheet reads; the last three walk most of the
 * spreadsheet and take tens of seconds on a real clinic's data. The console
 * painted nothing until the slowest of them finished, and the four counters
 * at the top sat on "–" the whole time, which is what "poor loading" looks
 * like from the desk.
 *
 * Now the queue and the register come back on their own and the screen is
 * usable immediately; this runs behind it and fills the Readiness tab in.
 * Each part still fails soft and says which one failed.
 *
 * @return {{success:boolean, findings:Array, anomalies:Array,
 *           triggers:string, message:string}}
 */
function dpdpConsoleHousekeeping(sessionToken) {
  try {
    crescRequire_(sessionToken, ['dpdp.manage', 'admin.config', 'admin.audit']);
    var out = { success: true, findings: [], anomalies: [], triggers: '', message: '' };

    try {
      var rd = dpdpReadinessCheck();
      if (rd && rd.findings) out.findings = rd.findings;
    } catch (e) { out.message += 'Readiness check failed (' + e.message + '). '; }

    try {
      var an = dpdp_anomalyScan_(7);
      out.anomalies = (an && an.findings) || [];
    } catch (e) { out.message += 'Audit review failed (' + e.message + '). '; }

    try {
      out.triggers = (typeof dpdpTriggerStatus === 'function') ? dpdpTriggerStatus() : '';
    } catch (e) { out.triggers = 'Could not read the trigger list: ' + e.message; }

    return out;
  } catch (err) {
    return { success: false, findings: [], anomalies: [], triggers: '',
             message: String(err.message || err).replace('FORBIDDEN: ', '') };
  }
}

/**
 * FRONTEND ENTRY. May dictation leave this building?
 *
 * Finding M3. Voice typing uses the browser's Web Speech API, which in Chrome
 * and Edge sends the AUDIO — a clinician saying a patient's history out loud —
 * to the browser vendor's speech service. The per-browser consent dialog in
 * Voice_Input.html is the right thing to show the person using it, but it is
 * the wrong place to make the decision: one clinician clicking Enable on one
 * laptop is not the clinic deciding that recorded clinical speech may go to a
 * third party.
 *
 * THE CLINIC HAS NOW DECIDED: dictation is allowed, and that is the default.
 * The per-person browser dialog is gone with it — a clinician who opens the
 * consult and presses the microphone is not being asked to make a policy
 * decision they were never in a position to make, and being asked once per
 * browser meant being asked again on every new machine, every cleared
 * profile and every private window.
 *
 * FORBIDDEN still works and still switches dictation off everywhere, for a
 * clinic that wants it off; it is simply no longer the thing an unset
 * property means. The readiness check no longer raises a finding for an
 * unset property either, because unset now means allowed, which is a
 * decision rather than an omission.
 */
function dpdpVoicePolicy(sessionToken) {
  try {
    crescRequire_(sessionToken, ['emr.write', 'ward.write', 'rx.write', 'admin.config',
                                 'dpdp.manage', 'patient.read']);
    var raw = '';
    try {
      raw = PropertiesService.getScriptProperties().getProperty('CRESC_VOICE_POLICY') || '';
    } catch (e) {}
    // Unset means ALLOWED. See the note above.
    var policy = dpdp_str_(raw).toUpperCase() || 'ALLOWED';
    return {
      success: true,
      policy: policy,
      allowed: policy !== 'FORBIDDEN',
      decided: true,
      message: policy === 'FORBIDDEN'
        ? 'This clinic has decided that dictated audio must not leave the building. ' +
          'Voice typing is switched off here — type the note instead.'
        : ''
    };
  } catch (err) {
    // Still fail CLOSED when the SESSION cannot be resolved: a dictation
    // feature that works when the server cannot be asked is a feature that
    // works for someone who is not signed in. This is about the caller, not
    // about the policy — the policy's own default is ALLOWED.
    return { success: false, policy: 'UNSET', allowed: false, decided: false,
             message: 'Voice typing could not confirm your session, so it is off. ' +
                      'Sign in again and retry.' };
  }
}

/** FRONTEND ENTRY. Records the clinic's decision about dictation. */
function dpdpSetVoicePolicy(policy, sessionToken) {
  try {
    var actor = crescRequire_(sessionToken, ['dpdp.manage', 'admin.config']);
    var want = dpdp_str_(policy).toUpperCase();
    if (want !== 'ALLOWED' && want !== 'FORBIDDEN') {
      return { success: false, message: 'Policy must be ALLOWED or FORBIDDEN.' };
    }
    PropertiesService.getScriptProperties().setProperty('CRESC_VOICE_POLICY', want);
    try {
      logAudit_({ username: actor.username, role: actor.role },
                'DPDP_VOICE_POLICY_SET', 'Config', 'CRESC_VOICE_POLICY', { policy: want });
    } catch (e) {}
    return { success: true, policy: want,
             message: want === 'ALLOWED'
               ? 'Voice typing is allowed, which is also the default.'
               : 'Voice typing is off for the whole clinic. No dictated audio leaves ' +
                 'the building through this application.' };
  } catch (err) {
    return { success: false, message: String(err.message || err).replace('FORBIDDEN: ', '') };
  }
}

/** FRONTEND ENTRY. The retention report, for the console. Reports, never deletes. */
function dpdpRetentionReportUI(sessionToken) {
  try {
    crescRequire_(sessionToken, ['dpdp.manage', 'admin.config']);
    return { success: true, report: dpdpRetentionReport() };
  } catch (err) {
    return { success: false, report: '',
             message: String(err.message || err).replace('FORBIDDEN: ', '') };
  }
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
  // The single highest-risk setting in the project, and it is not in code.
  //
  // This USED to be an unconditional HIGH built by reading appsscript.json —
  // which describes what the NEXT version would be published with, not what
  // the live /exec URL does. So it stayed HIGH on a clinic that had fixed it
  // months ago, and its fix was "go and look in a private window", which
  // nobody does twice. A finding that cannot go away is a finding that stops
  // being read.
  //
  // depDeploymentFinding() answers it from EVIDENCE instead: whether the page
  // loads actually arriving at this application carry an identified Google
  // user, which is a direct consequence of the live access mode. See
  // Deployment_Probe.gs.
  try {
    var dep = depDeploymentFinding();
    add(dep.severity, 'Deployment', dep.text, dep.fix);
  } catch (e) {
    add('HIGH', 'Deployment',
        'The manifest asks for access "ANYONE" and executeAs USER_DEPLOYING, so ' +
        'every google.script.run endpoint runs with the owner’s full ' +
        'spreadsheet access — and the deployment probe could not be read (' +
        e.message + '), so whether the LIVE deployment requires a sign-in is ' +
        'not known.',
        'Deploy > Manage deployments > edit > New version. Then open the /exec ' +
        'URL in a private window: if it answers without asking you to sign in, ' +
        'the old ANYONE_ANONYMOUS deployment is still live. Treat any period it ' +
        'was anonymous as potentially breached (s.8(6)).');
  }

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
  //
  // TWO SEPARATE FACTS, reported separately, because the remedies are not the
  // same and the old single count conflated them.
  //
  //   "no record at all" is an accountability gap: nothing says the basis for
  //   processing this patient was ever considered. It is closeable by the
  //   clinic on its own — dpdpBackfillConsent() writes the two s.7 legitimate
  //   uses, which is a record of a decision the clinic did make.
  //
  //   "the optional purposes have not been put to them" is a gap only the
  //   PATIENT can close, by being asked. No function can close it and no
  //   report should imply one could; what it needs is a worklist, which is
  //   dpdpConsentQueue().
  //
  // Reporting them as one number invited exactly the fix that must not
  // happen: a batch write of GIVEN against everything, which would clear the
  // finding by putting a false consent on the record.
  try {
    var pts = ss.getSheetByName('Patients');
    var anyRow = {}, optionalDone = {};
    var OPTIONAL = ['INSURANCE', 'COMMUNICATION', 'MARKETING', 'RESEARCH'];
    var cs = ss.getSheetByName(DPDP_CFG.CONSENT);
    if (cs) {
      var cv = dc_sheetValues_(cs);
      var cm = dc_headerMap_(cs);
      for (var i = 1; i < (cv ? cv.length : 0); i++) {
        var pid = dpdp_str_(cv[i][cm['Patient_ID']]).toUpperCase();
        if (!pid) continue;
        anyRow[pid] = true;
        var pk = dpdp_str_(cv[i][cm['Purpose']]).toUpperCase();
        if (OPTIONAL.indexOf(pk) === -1) continue;
        if (!optionalDone[pid]) optionalDone[pid] = {};
        optionalDone[pid][pk] = true;
      }
    }
    if (pts) {
      var total = Math.max(0, pts.getLastRow() - 1);
      var have = Object.keys(anyRow).length;

      if (have < total) {
        add(have === 0 ? 'HIGH' : 'MEDIUM', 'Section 6',
            (total - have) + ' of ' + total + ' patients have no consent ' +
            'record at all — nothing says the basis for processing them was ' +
            'ever considered.',
            'Run dpdpBackfillConsent({mode:"LEGITIMATE"}) — or Privacy > ' +
            'Consent > "Record the s.7 basis for every patient". It writes ' +
            'care and billing as the legitimate uses they are, and ' +
            'deliberately does NOT invent the four optional consents.');
      }

      // How many have been asked about ALL FOUR optional purposes.
      var asked = 0;
      Object.keys(optionalDone).forEach(function (pid) {
        var n = 0;
        OPTIONAL.forEach(function (k) { if (optionalDone[pid][k]) n++; });
        if (n === OPTIONAL.length) asked++;
      });
      if (asked < total) {
        add('MEDIUM', 'Section 6',
            (total - asked) + ' of ' + total + ' patients have not been asked ' +
            'about insurance, reports by WhatsApp or email, health camps or ' +
            'research. Documents cannot be sent to them until they are — the ' +
            'dispatch gate refuses.',
            'This one cannot be fixed by a function: the patient has to be ' +
            'asked. dpdpConsentQueue() is the worklist, Privacy > Consent ' +
            'shows it, and the WhatsApp button now puts the question at the ' +
            'counter and records the answer.');
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

  // --- dictation (finding M3) ---------------------------------------------
  // No finding any more. Dictation is allowed by default, which is a decision
  // the clinic has taken, so an unset property is no longer an open question.
  // dpdpSetVoicePolicy("FORBIDDEN") still turns it off everywhere. What the
  // consent notice says about dictation is a wording question, covered by the
  // notice-version check above rather than by a finding of its own.

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
      // The empty-register case is no longer a finding on its own: since
      // DPDP_Documents.gs there are no setSharing(ANYONE_WITH_LINK) calls
      // left, so an empty legacy register means the backlog is cleared, which
      // is the good outcome rather than a suspicious one.
    }
  } catch (e) { /* ditto */ }

  // --- document grants (the register that replaced the public links) ------
  try {
    var gs = ss.getSheetByName('Document_Grants');
    if (!gs) {
      add('MEDIUM', 'Section 8(5)',
          'Document_Grants does not exist, so no document has been shared with a ' +
          'patient since private links replaced public Drive links — or the ' +
          'registers were never created.',
          'Run dpdpSetup(). If documents ARE being sent, check that ' +
          'dpdpIssueDocumentLink_() is what sends them: grep the project for ' +
          'setSharing to be sure nothing publishes to Drive directly.');
    } else {
      var gv = dc_sheetValues_(gs);
      var gm = dc_headerMap_(gs);
      var liveOverdue = 0, heavilyOpened = 0, nowG = Date.now();
      for (var g = 1; g < (gv ? gv.length : 0); g++) {
        if (dpdp_str_(gv[g][gm['Status']]).toUpperCase() !== 'ACTIVE') continue;
        var gexp = (typeof cresc_parseDate_ === 'function')
          ? cresc_parseDate_(gv[g][gm['Expires_At']]) : new Date(gv[g][gm['Expires_At']]);
        if (gexp && !isNaN(gexp.getTime()) && gexp.getTime() < nowG) liveOverdue++;
        if ((parseInt(gv[g][gm['Opens']], 10) || 0) >= 6) heavilyOpened++;
      }
      if (liveOverdue) {
        add('MEDIUM', 'Section 8(7)',
            liveOverdue + ' document link(s) are past their expiry but still marked ' +
            'active, which means the daily job is not running.',
            'dpdpInstallTriggers(), then dpdpTriggerStatus() to confirm.');
      }
      if (heavilyOpened) {
        add('MEDIUM', 'Section 8(6)',
            heavilyOpened + ' shared document(s) have been opened six times or more. ' +
            'A patient reading their own report does not do that; a forwarded ' +
            'message does.',
            'Privacy console > A patient > withdraw the link, and consider whether ' +
            'the patient should be told (docs/BREACH_PROCEDURE.md).');
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
