// ============================================================================
// DPDP_Consent_Backfill.gs — Crescentia HealthTech
// The existing patients, brought into the consent register — honestly.
// ----------------------------------------------------------------------------
// THE FINDING THIS ANSWERS
//
// dpdpReadinessCheck() reports, under Section 6:
//
//     "23 of 25 patients have no consent record at all."
//
// which is true, and is a real gap: s.6(10) puts the burden of proving
// consent on the Data Fiduciary, and 23 patients are being treated, billed
// and messaged with nothing on record either way.
//
// WHAT CANNOT BE DONE ABOUT IT, AND WHY THIS FILE REFUSES TO
//
// The obvious "fix" is a button that writes GIVEN against all six purposes
// for all 25 patients. That would clear the finding and would be worse than
// the gap it closes. s.6(1) requires consent to be free, specific, informed
// and given by a clear affirmative action; a row written by a script on
// behalf of somebody who was never asked is a FALSE RECORD, and a false
// consent record is the single most damaging artefact a data-protection
// audit can find — it converts "we had not got round to asking" into "we
// claimed they agreed". It also actively harms the patient, because the
// clinic then sends their reports over WhatsApp on the strength of it.
//
// So dpdpBackfillConsent() will not invent an answer. It does two separate
// things, and the caller has to say which.
//
// MODE 1 — LEGITIMATE (the default, and safe to run unattended)
//
// TREATMENT and BILLING are marked `legitimate` in DPDP_PURPOSES: the clinic
// processes them as legitimate uses under s.7 because the patient approached
// it to be treated. They are NOT consent-gated, which is why the portal shows
// them as statements rather than switches. What the register was missing for
// them is not permission — it is the RECORD that the basis was considered for
// this patient, which s.5 and the accountability principle want written down.
//
// This mode writes those two, for every patient who has no row, with method
// BACKFILL_LEGITIMATE_USE and a note saying exactly that. It does NOT touch
// INSURANCE, COMMUNICATION, MARKETING or RESEARCH, which are real consents
// and stay NOT_ASKED — a state getConsentStatus() already reports separately
// from REFUSED for precisely this reason.
//
// After it runs, the Section 6 finding changes from "no record at all" to
// "the four optional purposes have not been put to them", which is an
// accurate description of where the clinic stands and a shorter list to work
// through.
//
// MODE 2 — RECORD (for consent the clinic has ACTUALLY collected)
//
// If the clinic has a signed paper consent register, or a stack of
// registration forms with ticked boxes, those are consents and they belong in
// the register. This mode writes them — but it makes the caller state, per
// patient or for the batch, WHICH purposes were agreed, HOW they were
// collected, and WHERE the evidence is. Without the evidence sentence it
// refuses. That sentence is what an audit asks for and it is the only thing
// distinguishing this mode from fabrication.
//
// WHAT TO DO WITH THE REST
//
// dpdpConsentQueue() lists the patients whose optional purposes are still
// unanswered, newest visit first, so the front desk works through them as
// people come in. That is what "backfill at the next visit" means in
// practice, and it needs a worklist rather than a button.
// ============================================================================

var DPDP_BACKFILL = {
  /** The two purposes processed under s.7, which this file may write. */
  LEGITIMATE: ['TREATMENT', 'BILLING'],

  /** The four that are genuine consents, which it may not invent. */
  CONSENTED: ['INSURANCE', 'COMMUNICATION', 'MARKETING', 'RESEARCH'],

  METHOD_LEGITIMATE: 'BACKFILL_LEGITIMATE_USE',

  NOTE_LEGITIMATE:
    'Recorded by backfill as a legitimate use under s.7 — the patient ' +
    'approached the clinic for treatment. This is NOT a consent and was not ' +
    'obtained as one; it records that the basis for processing was considered ' +
    'for this patient.',

  /** Rows written in one call, to stay inside the execution limit. */
  BATCH_CAP: 500
};

/**
 * Every patient id and name on file, in sheet order.
 * @return {Array<{id:string, name:string, age:string, row:number}>}
 */
function dpdp_allPatients_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Patients');
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getDataRange().getDisplayValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][0] || '').trim().toUpperCase();
    if (!id) continue;
    out.push({ id: id, name: String(data[i][2] || '').trim(),
               age: String(data[i][3] || '').trim(), row: i + 1 });
  }
  return out;
}

/**
 * patientId -> { PURPOSE: state } for the whole register, in one read.
 *
 * dpdp_consentState_() (DPDP_Dispatch.gs) answers one patient and one purpose
 * and re-reads the sheet each time. For 25 patients × 6 purposes that is 150
 * reads of the same range, and this file is the only place that wants all of
 * it at once.
 */
function dpdp_consentMatrix_() {
  var out = {};
  var sh = dpdp_consentSheet_();
  var m = dc_headerMap_(sh);
  var data = dc_sheetValues_(sh);
  for (var i = 1; i < (data ? data.length : 0); i++) {
    var pid = dpdp_str_(data[i][m['Patient_ID']]).toUpperCase();
    if (!pid) continue;
    var purpose = dpdp_str_(data[i][m['Purpose']]).toUpperCase();
    if (!purpose) continue;
    if (!out[pid]) out[pid] = {};
    // Append-only register: later rows win.
    out[pid][purpose] = dpdp_str_(data[i][m['Withdrawn_At']])
      ? 'WITHDRAWN'
      : (dpdp_str_(data[i][m['Decision']]).toUpperCase() || 'NOT_ASKED');
  }
  return out;
}

/**
 * FRONTEND ENTRY / ADMIN. Brings existing patients into the consent register.
 *
 * @param {{mode:string, decisions:Object, method:string, evidence:string,
 *          patientIds:Array, dryRun:boolean}} payload
 *
 *        mode       'LEGITIMATE' (default) or 'RECORD'
 *        decisions  RECORD mode only: { COMMUNICATION:true, MARKETING:false, … }
 *        method     RECORD mode only: PAPER_FORM | IN_PERSON | PHONE | PORTAL
 *        evidence   RECORD mode only: where the signed answers are kept
 *        patientIds optional; omit for every patient with a gap
 *        dryRun     report what would be written, write nothing
 *
 * @param {string} sessionToken
 * @return {{success:boolean, written:number, patients:number, skipped:number,
 *           rows:Array, message:string}}
 */
function dpdpBackfillConsent(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    payload = payload || {};

    var actor = crescRequire_(sessionToken, ['dpdp.manage', 'admin.config']);
    var mode = dpdp_str_(payload.mode).toUpperCase() || 'LEGITIMATE';
    var dryRun = payload.dryRun === true;

    if (mode !== 'LEGITIMATE' && mode !== 'RECORD') {
      return { success: false, written: 0, patients: 0, skipped: 0, rows: [],
               message: 'mode must be LEGITIMATE or RECORD.' };
    }

    // ---- what is being written, and is it the caller's to write? ---------
    var purposes = [], decisions = {}, method = '', note = '';

    if (mode === 'LEGITIMATE') {
      purposes = DPDP_BACKFILL.LEGITIMATE.slice();
      purposes.forEach(function (p) { decisions[p] = true; });
      method = DPDP_BACKFILL.METHOD_LEGITIMATE;
      note = DPDP_BACKFILL.NOTE_LEGITIMATE;

    } else {
      var given = payload.decisions || {};
      var keys = Object.keys(given);
      if (!keys.length) {
        return { success: false, written: 0, patients: 0, skipped: 0, rows: [],
                 message: 'RECORD mode needs the decisions the patients ' +
                          'actually gave, e.g. { COMMUNICATION: true, ' +
                          'MARKETING: false }.' };
      }
      var unknown = keys.filter(function (k) { return !dpdp_purpose_(k); });
      if (unknown.length) {
        return { success: false, written: 0, patients: 0, skipped: 0, rows: [],
                 message: 'Unknown purpose(s): ' + unknown.join(', ') + '.' };
      }

      // The gate. A batch of consents with no stated source is indistinguishable
      // from a batch of invented ones, and the difference is the whole point.
      var evidence = dpdp_str_(payload.evidence);
      if (evidence.length < 20) {
        return { success: false, written: 0, patients: 0, skipped: 0, rows: [],
                 message: 'Say where these answers came from, in a sentence — ' +
                          'e.g. "signed consent forms in the registration file, ' +
                          'Jan-Sep 2026, held at reception". Every row written ' +
                          'carries it, and it is the only thing that tells an ' +
                          'audit these are recorded consents rather than ' +
                          'assumed ones. Nothing was written.' };
      }
      method = dpdp_str_(payload.method).toUpperCase() || 'PAPER_FORM';
      note = 'Backfilled from consent already collected. Source: ' + evidence;
      purposes = keys;
      keys.forEach(function (k) { decisions[dpdp_purpose_(k).key] = (given[k] === true); });
      purposes = Object.keys(decisions);
    }

    // ---- who ------------------------------------------------------------
    var wanted = null;
    if (Array.isArray(payload.patientIds) && payload.patientIds.length) {
      wanted = {};
      payload.patientIds.forEach(function (p) {
        wanted[String(p || '').trim().toUpperCase()] = true;
      });
    }

    var patients = dpdp_allPatients_();
    if (!patients.length) {
      return { success: true, written: 0, patients: 0, skipped: 0, rows: [],
               message: 'No patients on file.' };
    }

    var matrix = dpdp_consentMatrix_();
    var version = dpdp_noticeVersion_();
    var now = dpdp_now_();
    var stamp = Utilities.formatDate(now, DPDP_CFG.TZ, 'yyMMdd-HHmmss');

    var rows = [], report = [], touched = 0, skipped = 0, guardianNeeded = [];

    for (var i = 0; i < patients.length; i++) {
      var p = patients[i];
      if (wanted && !wanted[p.id]) continue;

      var have = matrix[p.id] || {};
      var missing = purposes.filter(function (k) {
        // Only purposes with NO row are written. A patient who refused
        // marketing must not have that refusal overwritten by a batch job,
        // and a patient who already agreed does not need a second row.
        return !have[k];
      });
      if (!missing.length) { skipped++; continue; }

      // s.9 — a child's consent comes from a parent or guardian, and a batch
      // job is not a guardian. LEGITIMATE mode is exempt because it records a
      // s.7 basis rather than a consent; RECORD mode is not, and the child is
      // named so the clinic can go and ask properly.
      var age = parseInt(p.age, 10);
      if (mode === 'RECORD' && isFinite(age) && age < 18 &&
          !dpdp_str_(payload.guardianName)) {
        guardianNeeded.push(p.id + ' (' + p.name + ', ' + age + ')');
        skipped++;
        continue;
      }

      if (rows.length + missing.length > DPDP_BACKFILL.BATCH_CAP) {
        report.push('Stopped at the batch cap of ' + DPDP_BACKFILL.BATCH_CAP +
                    ' rows. Run it again to continue.');
        break;
      }

      missing.forEach(function (k) {
        rows.push([
          'CNS-' + stamp + '-' + Utilities.getUuid().substring(0, 4).toUpperCase(),
          p.id, p.name, k,
          decisions[k] ? 'GIVEN' : 'REFUSED',
          version, method, actor.username, now,
          '', '',
          dpdp_str_(payload.guardianName),
          dpdp_str_(payload.guardianRelation),
          note
        ]);
      });
      touched++;
      report.push(p.id + '  ' + missing.join(', '));
    }

    if (!rows.length) {
      return { success: true, written: 0, patients: 0, skipped: skipped, rows: [],
               message: 'Nothing to write — every patient already has a row for ' +
                        purposes.join(', ') + '.' +
                        (guardianNeeded.length
                          ? '\n\n' + guardianNeeded.length + ' patient(s) are under 18 ' +
                            'and need a named guardian (s.9): ' +
                            guardianNeeded.join('; ')
                          : '') };
    }

    if (dryRun) {
      return { success: true, written: 0, patients: touched, skipped: skipped,
               rows: report,
               message: 'DRY RUN — ' + rows.length + ' row(s) would be written for ' +
                        touched + ' patient(s). Nothing was changed.' };
    }

    var sh = dpdp_consentSheet_();
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    dc_invalidate_(DPDP_CFG.CONSENT);
    SpreadsheetApp.flush();

    try {
      logAudit_({ username: actor.username, role: actor.role },
                'DPDP_CONSENT_BACKFILL', 'Patient', 'BATCH',
                { mode: mode, purposes: purposes, patients: touched,
                  rows: rows.length, method: method,
                  noticeVersion: version,
                  evidence: dpdp_str_(payload.evidence) });
    } catch (e) {}

    var tail = (mode === 'LEGITIMATE')
      ? '\n\nThese are the two purposes the clinic processes as LEGITIMATE USES ' +
        'under s.7, not consents. The four that are real consents — insurance, ' +
        'reports by WhatsApp or email, health camps, and research — are ' +
        'deliberately untouched and still read "not asked yet". ' +
        'dpdpConsentQueue() is the list to work through at the desk.'
      : '';

    return {
      success: true, written: rows.length, patients: touched, skipped: skipped,
      rows: report,
      message: rows.length + ' consent row(s) written for ' + touched +
               ' patient(s) against notice ' + version + '. ' + skipped +
               ' patient(s) already had a row and were left alone.' +
               (guardianNeeded.length
                 ? '\n\n' + guardianNeeded.length + ' patient(s) under 18 were ' +
                   'SKIPPED — s.9 needs a parent or guardian named: ' +
                   guardianNeeded.join('; ')
                 : '') + tail
    };

  } catch (err) {
    return { success: false, written: 0, patients: 0, skipped: 0, rows: [],
             message: String((err && err.message) || err).replace('FORBIDDEN: ', '') };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * FRONTEND ENTRY. The patients whose real consents have still not been put to
 * them, and which ones are outstanding.
 *
 * This is what "backfill at the next visit" needs in order to happen: a
 * worklist the front desk can see, not an instruction in a readiness report.
 *
 * @return {{success, rows:Array, counts:Object, message:string}}
 */
function dpdpConsentQueue(sessionToken) {
  try {
    crescRequire_(sessionToken, ['dpdp.manage', 'admin.config', 'patient.register',
                                 'patient.write']);
    var patients = dpdp_allPatients_();
    var matrix = dpdp_consentMatrix_();
    var current = dpdp_noticeVersion_();

    var rows = [], counts = { complete: 0, partial: 0, none: 0 };

    patients.forEach(function (p) {
      var have = matrix[p.id] || {};
      var outstanding = DPDP_BACKFILL.CONSENTED.filter(function (k) { return !have[k]; });
      var anyRow = Object.keys(have).length > 0;

      if (!outstanding.length) { counts.complete++; return; }
      if (anyRow) counts.partial++; else counts.none++;

      rows.push({
        patientId: p.id,
        name: p.name,
        age: p.age,
        outstanding: outstanding,
        outstandingLabels: outstanding.map(function (k) {
          var d = dpdp_purpose_(k);
          return d ? d.label : k;
        }),
        hasAnyRecord: anyRow,
        // Under 18 needs a guardian present (s.9), which changes who the desk
        // has to ask — worth knowing before they call the patient over.
        needsGuardian: (function () {
          var a = parseInt(p.age, 10);
          return isFinite(a) && a < 18;
        })()
      });
    });

    // Most outstanding first: a patient with nothing on record is the one to
    // catch, and a patient missing one purpose can wait.
    rows.sort(function (a, b) { return b.outstanding.length - a.outstanding.length; });

    return {
      success: true, rows: rows, counts: counts, noticeVersion: current,
      purposes: DPDP_BACKFILL.CONSENTED.map(function (k) {
        var d = dpdp_purpose_(k);
        return { key: k, label: d ? d.label : k, detail: d ? d.detail : '' };
      }),
      message: rows.length
        ? rows.length + ' patient(s) still have consents that have not been put ' +
          'to them. Ask when they next come in.'
        : 'Every patient has been asked about every optional purpose.'
    };
  } catch (err) {
    return { success: false, rows: [],
             message: String((err && err.message) || err).replace('FORBIDDEN: ', '') };
  }
}

// ---------------------------------------------------------------------------
// RUNNABLE FROM THE SCRIPT EDITOR
//
// Each one mints an admin session, so each one starts with crescEditorOnly_()
// (see RUN_Setup.gs). Without it they were endpoints that handed that
// session's powers — the consent worklist with every patient's name, and a
// real write to the consent register — to anyone holding the URL.
// ---------------------------------------------------------------------------

/**
 * A short-lived admin session for a job run from the script editor.
 *
 * Same pattern as RUN_90_resetOnePassword in RUN_Setup.gs: the guard on
 * these endpoints is real and is not bypassed for the editor — a session is
 * issued, used, and revoked when the job ends, so the row it writes to the
 * consent register carries a username an audit can follow.
 */
function dpdp_editorSession_(fn) {
  var token = issueSession_({ username: 'SCRIPT_OWNER', role: 'admin',
                              doctorId: '', name: 'Script owner (editor)' });
  try {
    return fn(token);
  } finally {
    try { revokeSession(token); } catch (e) {}
  }
}

/**
 * ADMIN, from the editor. What the LEGITIMATE backfill would write. Writes
 * nothing.
 */
function RUN_consentBackfill_DRYRUN() {
  crescEditorOnly_('RUN_consentBackfill_DRYRUN');
  var r = dpdp_editorSession_(function (t) {
    return dpdpBackfillConsent({ mode: 'LEGITIMATE', dryRun: true }, t);
  });
  var out = r.message + (r.rows && r.rows.length ? '\n\n' + r.rows.join('\n') : '');
  Logger.log(out);
  return out;
}

/** ADMIN, from the editor. Runs the LEGITIMATE backfill for real. */
function RUN_consentBackfill_FOR_REAL() {
  crescEditorOnly_('RUN_consentBackfill_FOR_REAL');
  var r = dpdp_editorSession_(function (t) {
    return dpdpBackfillConsent({ mode: 'LEGITIMATE' }, t);
  });
  var out = r.message + (r.rows && r.rows.length ? '\n\n' + r.rows.join('\n') : '');
  Logger.log(out);
  return out;
}

/** ADMIN, from the editor. The worklist for the front desk. */
function RUN_consentQueue() {
  crescEditorOnly_('RUN_consentQueue');
  var r = dpdp_editorSession_(function (t) { return dpdpConsentQueue(t); });
  var lines = (r.rows || []).map(function (x) {
    return '  ' + x.patientId + '  ' + x.name +
           (x.needsGuardian ? '  [under 18 — guardian needed]' : '') +
           '\n      still to ask: ' + x.outstandingLabels.join('; ');
  });
  var out = r.message + (lines.length ? '\n\n' + lines.join('\n') : '');
  Logger.log(out);
  return out;
}
