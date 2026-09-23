// ============================================================================
// Auth_Reset.gs — Crescentia HealthTech
// Forgot password: a temporary one, by email, without telephoning the clinic.
// ----------------------------------------------------------------------------
// WHAT THERE WAS
//
// crescAdminResetPassword() (Auth_Credentials.gs) — which works, and is the
// wrong shape for the situation it is used in. It requires an administrator
// to be signed in, at a computer, and to then READ THE TEMPORARY PASSWORD
// OUT to the person who needs it. That has three costs:
//
//   * A nurse who cannot sign in at 2am has no route at all until somebody
//     with admin.users wakes up.
//   * The temporary password travels by voice, WhatsApp or a note on the
//     desk. Every one of those is a worse channel than the mailbox the
//     account already has on file.
//   * The administrator knows the password. For the period before it is
//     changed, two people can sign in as that account and the audit log
//     cannot tell them apart.
//
// A patient was worse off still: crescChangePassword refuses a legacy portal
// credential and tells them to "ask at reception", which for a portal whose
// whole purpose is not having to come in is close to useless.
//
// WHAT THIS DOES
//
// The person names their account. A random temporary password is stored
// BESIDE the current one (which keeps working) and EMAILED to the address
// already on the record —
// never to an address supplied in the request, which would turn this into a
// way to take over any account whose id you can guess. They sign in with it
// and the existing must-change flow forces them to set their own.
//
// THREE THINGS IT IS CAREFUL ABOUT
//
// 1. IT DOES NOT SAY WHETHER THE ACCOUNT EXISTS.
//    Patient ids in this system are sequential and printed on every barcode
//    label. An endpoint that answers "no such account" for PT-0104 and "check
//    your email" for PT-0103 is a directory of who is a patient here, open to
//    anyone with the URL. Every outcome returns the same sentence. The
//    DIFFERENCE is in the audit log, where it belongs.
//
// 2. IT IS RATE LIMITED, PER ACCOUNT AND PER DEPLOYMENT.
//    It no longer invalidates the real password (see "THE TEMPORARY PASSWORD
//    LIVES BESIDE THE REAL ONE" below), so a stranger clicking it cannot lock
//    anybody out — but it still sends mail. Three per account per hour, and a
//    deployment-wide ceiling so a script cannot walk the patient id range.
//
// 3. THE TEMPORARY PASSWORD IS SHORT-LIVED BY POLICY, NOT BY HOPE.
//    The must-change flag means it cannot be left in place, and the reset is
//    audited with the address it went to, so "who asked for this" is
//    answerable afterwards.
//
// WHY GMAIL AND NOT SMS
//
// Apps Script carries GmailApp on the free tier with a daily quota this
// clinic will not approach; SMS needs a paid gateway, an account and a
// template approval. The address is already on the record for both staff
// (Users column E) and patients (Patients column Q), and nothing else has to
// be bought or signed up for.
// ============================================================================

var CRESC_RESET = {
  /** Per-account resets allowed inside the window below. */
  MAX_PER_ACCOUNT: 3,

  /** Resets allowed across the whole deployment in the same window. */
  MAX_GLOBAL: 20,

  WINDOW_S: 3600,

  /** Length of the temporary password. Same alphabet as every other one. */
  TEMP_LENGTH: 12,

  /** How long an emailed temporary password stays usable. */
  TEMP_HOURS: 24,

  /**
   * The one answer every outcome gets.
   *
   * Deliberately a statement about what the SYSTEM did, not about what was
   * found: "if that account exists" is doing the work that stops this being
   * an account-existence oracle.
   */
  SAME_ANSWER:
    'If that account exists and has an email address on file, a temporary ' +
    'password has just been sent to it. Check the inbox — and the spam ' +
    'folder — then sign in with it and choose a new password. Your current ' +
    'password keeps working until you do. If nothing arrives within a few ' +
    'minutes, please contact the clinic.'
};

function cresc_resetKey_(username) {
  return 'PWRESET_' + String(username || '').toUpperCase()
           .replace(/[^A-Z0-9_\-]/g, '').substring(0, 100);
}

/**
 * Counts this request against both caps.
 * @return {string} '' when it may proceed, otherwise why not
 */
function cresc_resetThrottle_(username) {
  try {
    var cache = CacheService.getScriptCache();

    var gKey = 'PWRESET_ALL';
    var g = parseInt(cache.get(gKey), 10) || 0;
    if (g >= CRESC_RESET.MAX_GLOBAL) {
      return 'Too many password resets have been requested across the clinic ' +
             'in the last hour. Please contact the clinic directly.';
    }

    var uKey = cresc_resetKey_(username);
    var u = parseInt(cache.get(uKey), 10) || 0;
    if (u >= CRESC_RESET.MAX_PER_ACCOUNT) {
      // Said plainly. This one DOES leak that the id has been asked about
      // recently, which is information the person doing the asking already
      // has, and the alternative is letting them keep going.
      return 'A temporary password has already been sent for this account ' +
             'more than once in the last hour. Check the inbox, or contact ' +
             'the clinic.';
    }

    cache.put(uKey, String(u + 1), CRESC_RESET.WINDOW_S);
    cache.put(gKey, String(g + 1), CRESC_RESET.WINDOW_S);
    return '';
  } catch (e) {
    // A cache that cannot be read must not become an uncapped reset endpoint.
    return 'Password resets are temporarily unavailable. Please contact the clinic.';
  }
}

// ---------------------------------------------------------------------------
// THE TEMPORARY PASSWORD LIVES BESIDE THE REAL ONE, NOT IN ITS PLACE
//
// A reset used to overwrite the password the moment it was asked for. Anyone
// who knew a colleague's user id could therefore lock them out from the
// signed-out page, three times an hour, and the clinic-wide cap meant twenty
// such clicks stopped everyone else resetting at all.
//
// Now the temporary password is stored in its own columns with an expiry.
// The old password keeps working. Whichever is used first decides:
//   * the temporary one  -> the person must choose a new password, and that
//                           replaces the old one (crescChangePassword);
//   * the old one        -> the pending temporary password is discarded.
// ---------------------------------------------------------------------------

function cresc_resetCols_(sheet, isPatient) {
  return {
    temp: ensureColumn_(sheet, isPatient ? 'Portal_Reset_Temp' : 'Reset_Temp'),
    exp:  ensureColumn_(sheet, isPatient ? 'Portal_Reset_Expires' : 'Reset_Expires')
  };
}

/** Stores a pending temporary password; the real one is untouched. */
function cresc_writePendingReset_(sheet, rowNo, temp, isPatient) {
  var c = cresc_resetCols_(sheet, isPatient);
  sheet.getRange(rowNo, c.temp).setNumberFormat('@').setValue(crescPwdEncode_(temp));
  sheet.getRange(rowNo, c.exp)
       .setValue(new Date(Date.now() + CRESC_RESET.TEMP_HOURS * 3600 * 1000));
}

/** Does `plain` match an unexpired pending temporary password on this row? */
function cresc_pendingResetMatches_(sheet, rowNo, plain, isPatient) {
  try {
    var tc = cresc_colOf_(sheet, isPatient ? 'Portal_Reset_Temp' : 'Reset_Temp');
    var ec = cresc_colOf_(sheet, isPatient ? 'Portal_Reset_Expires' : 'Reset_Expires');
    if (tc === -1 || ec === -1 || !plain) return false;
    var stored = sheet.getRange(rowNo, tc).getValue();
    if (!stored || !crescPwdIsHashed_(stored)) return false;
    var exp = sheet.getRange(rowNo, ec).getValue();
    exp = (exp instanceof Date) ? exp : new Date(exp);
    if (isNaN(exp.getTime()) || exp.getTime() < Date.now()) {
      cresc_clearPendingReset_(sheet, rowNo, isPatient);
      return false;
    }
    return crescPwdVerify_(plain, stored).ok === true;
  } catch (e) {
    return false;
  }
}

/** Discards any pending temporary password on this row. Never throws. */
function cresc_clearPendingReset_(sheet, rowNo, isPatient) {
  try {
    var tc = cresc_colOf_(sheet, isPatient ? 'Portal_Reset_Temp' : 'Reset_Temp');
    var ec = cresc_colOf_(sheet, isPatient ? 'Portal_Reset_Expires' : 'Reset_Expires');
    if (tc !== -1 && sheet.getRange(rowNo, tc).getValue() !== '') sheet.getRange(rowNo, tc).setValue('');
    if (ec !== -1 && sheet.getRange(rowNo, ec).getValue() !== '') sheet.getRange(rowNo, ec).setValue('');
  } catch (e) {}
}

/** A plausible email address, for deciding whether there is anywhere to send. */
function cresc_looksLikeEmail_(v) {
  return /^[^@\s]+@[^@\s.]+\.[^@\s]+$/.test(String(v || '').trim());
}

/**
 * FRONTEND ENTRY, deliberately reachable WITHOUT a session — a person who
 * cannot sign in is exactly who needs it.
 *
 * @param {{username:string}} payload
 * @return {{success:boolean, message:string}}
 */
function crescRequestPasswordReset(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    payload = payload || {};
    var want = String(payload.username || '').trim();

    if (!want) {
      return { success: false, message: 'Enter your user ID or patient ID first.' };
    }
    // An id nothing could match is rejected before it costs a sheet read, and
    // before it counts against anybody's throttle.
    if (want.length > 40 || /[\r\n<>]/.test(want)) {
      return { success: false, message: 'That does not look like a user ID.' };
    }

    var held = cresc_resetThrottle_(want);
    if (held) return { success: false, message: held };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var temp = crescRandomPassword_(CRESC_RESET.TEMP_LENGTH);
    var clinic = (typeof cresc_clinic_ === 'function')
      ? cresc_clinic_() : { name: 'Crescentia HealthTech' };

    // ---- staff -----------------------------------------------------------
    var users = ss.getSheetByName('Users');
    if (users && users.getLastRow() > 1) {
      var udata = users.getDataRange().getValues();
      for (var i = 1; i < udata.length; i++) {
        if (!udata[i][0]) continue;
        if (String(udata[i][0]).trim().toUpperCase() !== want.toUpperCase()) continue;

        var status = String(udata[i][3] || '').trim().toUpperCase();
        if (status && status !== 'ACTIVE') {
          // A disabled account is not reset and not distinguished in the
          // reply: re-enabling it is an administrator's decision, and saying
          // "that account is disabled" tells an outsider it exists.
          crescAuthAudit_(CRESC_AUTH_EVENTS.PASSWORD_RESET_REFUSED, want,
                          String(udata[i][2] || ''), { reason: 'account not active' });
          return { success: true, message: CRESC_RESET.SAME_ANSWER };
        }

        var uEmail = String(udata[i][4] || '').trim();
        if (!cresc_looksLikeEmail_(uEmail)) {
          crescAuthAudit_(CRESC_AUTH_EVENTS.PASSWORD_RESET_REFUSED, want,
                          String(udata[i][2] || ''), { reason: 'no email on record' });
          return { success: true, message: CRESC_RESET.SAME_ANSWER };
        }

        // Written BEFORE the send: a temporary password that was mailed but
        // not stored would be one nobody can use. It sits beside the real
        // password, which keeps working until one or the other is used.
        cresc_writePendingReset_(cresc_usersSheet_(), i + 1, temp, false);
        SpreadsheetApp.flush();

        var sent = cresc_sendResetEmail_(uEmail, want, temp, clinic,
                                         String(udata[i][2] || 'staff'));
        crescAuthAudit_(CRESC_AUTH_EVENTS.PASSWORD_RESET, want,
                        String(udata[i][2] || ''),
                        { by: 'self-service', emailed: sent.ok,
                          to: cresc_maskEmail_(uEmail),
                          problem: sent.ok ? '' : sent.message });

        if (!sent.ok) {
          // Nothing was lost — the old password still works — but the person
          // is waiting for an email that is not coming, and should be told.
          // It is about this system failing, not about whether the account
          // exists, so it is the one case with a different answer.
          cresc_clearPendingReset_(cresc_usersSheet_(), i + 1, false);
          return { success: false,
                   message: 'The email could not be sent (' + sent.message + '). ' +
                            'Your current password still works; if you have lost ' +
                            'it, contact the clinic.' };
        }
        return { success: true, message: CRESC_RESET.SAME_ANSWER };
      }
    }

    // ---- patient portal --------------------------------------------------
    var patients = ss.getSheetByName('Patients');
    if (patients && patients.getLastRow() > 1) {
      var pdata = patients.getDataRange().getValues();
      for (var j = 1; j < pdata.length; j++) {
        if (!pdata[j][0]) continue;
        if (String(pdata[j][0]).trim().toUpperCase() !== want.toUpperCase()) continue;

        // Column Q (index 16) is the patient's email, the same column
        // buildLongitudinalTimeline reads for the consult dispatch address.
        var pEmail = String(pdata[j][16] || '').trim();
        if (!cresc_looksLikeEmail_(pEmail)) {
          crescAuthAudit_(CRESC_AUTH_EVENTS.PASSWORD_RESET_REFUSED, want, 'patient',
                          { reason: 'no email on record' });
          return { success: true, message: CRESC_RESET.SAME_ANSWER };
        }

        cresc_writePendingReset_(cresc_patientsSheet_(), j + 1, temp, true);
        SpreadsheetApp.flush();

        var psent = cresc_sendResetEmail_(pEmail, want, temp, clinic, 'patient');
        crescAuthAudit_(CRESC_AUTH_EVENTS.PASSWORD_RESET, want, 'patient',
                        { by: 'self-service', emailed: psent.ok,
                          to: cresc_maskEmail_(pEmail),
                          problem: psent.ok ? '' : psent.message });

        if (!psent.ok) {
          cresc_clearPendingReset_(cresc_patientsSheet_(), j + 1, true);
          return { success: false,
                   message: 'The email could not be sent (' + psent.message + '). ' +
                            'Your current password still works; if you have lost ' +
                            'it, please contact the clinic.' };
        }
        return { success: true, message: CRESC_RESET.SAME_ANSWER };
      }
    }

    // No such account. Same answer, different audit row.
    crescAuthAudit_(CRESC_AUTH_EVENTS.PASSWORD_RESET_REFUSED, want, '',
                    { reason: 'no such account' });
    return { success: true, message: CRESC_RESET.SAME_ANSWER };

  } catch (err) {
    return { success: false,
             message: 'The reset could not be completed: ' + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/** j***@gmail.com — enough to recognise, not enough to harvest. */
function cresc_maskEmail_(email) {
  var s = String(email || '').trim();
  var at = s.indexOf('@');
  if (at < 1) return '';
  return s.charAt(0) + '***' + s.substring(at);
}

/**
 * Sends the temporary password.
 *
 * PLAIN TEXT AND HTML, both saying the same thing: a mail client that strips
 * HTML must not strip the password. No link back into the application — a
 * "click here to reset" link in an email about a password is the shape every
 * phishing message takes, and this one has nothing to gain from it.
 *
 * @return {{ok:boolean, message:string}}
 */
function cresc_sendResetEmail_(to, username, temp, clinic, role) {
  var name = (clinic && clinic.name) || 'Crescentia HealthTech';
  var site = (clinic && clinic.website) ? String(clinic.website) : '';
  var phone = (clinic && clinic.phone) ? String(clinic.phone) : '';

  var subject = 'Your temporary password — ' + name;

  var plain =
    'A temporary password was requested for the account ' + username + '.\n\n' +
    'Temporary password:  ' + temp + '\n\n' +
    'Sign in with it and you will be asked to choose a new password straight ' +
    'away. It works once, and only for the next ' + CRESC_RESET.TEMP_HOURS + ' hours.\n\n' +
    'IF YOU DID NOT ASK FOR THIS, you can ignore this email: your current ' +
    'password still works, and signing in with it cancels this temporary one. ' +
    'If it keeps happening, tell the clinic' + (phone ? ' on ' + phone : '') + '.\n\n' +
    'We will never ask you for your password, and this email contains no ' +
    'links to click.\n\n' +
    name + (site ? '\n' + site : '');

  var html =
    '<div style="font-family:Arial,Helvetica,sans-serif;color:#111827;max-width:560px;' +
      'margin:0 auto;border:1px solid #e5e7eb;border-radius:8px;overflow:hidden;">' +
      '<div style="background:#0369a1;color:#fff;padding:18px 22px;">' +
        '<div style="font-size:16px;font-weight:700;letter-spacing:.04em;">' +
          cresc_htmlEsc_(name.toUpperCase()) + '</div></div>' +
      '<div style="padding:24px 22px;font-size:14px;line-height:1.6;">' +
        '<p style="margin:0 0 14px;">A temporary password was requested for the ' +
          'account <strong>' + cresc_htmlEsc_(username) + '</strong>.</p>' +
        '<div style="background:#f8fafc;border:1px dashed #94a3b8;border-radius:8px;' +
          'padding:16px;text-align:center;margin:18px 0;">' +
          '<div style="font-size:11px;color:#64748b;text-transform:uppercase;' +
            'letter-spacing:.08em;">Temporary password</div>' +
          '<div style="font-family:monospace;font-size:22px;font-weight:700;' +
            'letter-spacing:.12em;margin-top:6px;">' + cresc_htmlEsc_(temp) + '</div>' +
        '</div>' +
        '<p style="margin:0 0 14px;">Sign in with it and you will be asked to choose ' +
          'a new password straight away. It works once, and only for the next ' +
          CRESC_RESET.TEMP_HOURS + ' hours.</p>' +
        '<p style="margin:0 0 14px;padding:12px 14px;background:#f8fafc;' +
          'border-left:4px solid #64748b;color:#334155;">' +
          '<strong>If you did not ask for this</strong>, you can ignore this email: ' +
          'your current password still works, and signing in with it cancels this ' +
          'temporary one. If it keeps happening, tell the clinic' +
          (phone ? ' on ' + cresc_htmlEsc_(phone) : '') + '.</p>' +
        '<p style="margin:0;font-size:12px;color:#6b7280;">We will never ask you ' +
          'for your password. This email contains no links to click.</p>' +
      '</div>' +
      '<div style="padding:12px 22px;background:#f8fafc;border-top:1px solid #e5e7eb;' +
        'font-size:11px;color:#94a3b8;">' + cresc_htmlEsc_(name) +
        (site ? ' &middot; ' + cresc_htmlEsc_(site) : '') + '</div>' +
    '</div>';

  try {
    GmailApp.sendEmail(to, subject, plain, { htmlBody: html, name: name });
    return { ok: true, message: '' };
  } catch (e) {
    // The commonest real cause is the daily quota, and "Service invoked too
    // many times" tells the person at the keyboard nothing they can act on.
    var m = String(e.message || e);
    if (/quota|too many times/i.test(m)) {
      m = "today's email quota is used up";
    }
    return { ok: false, message: m };
  }
}

function cresc_htmlEsc_(v) {
  return String(v === null || v === undefined ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * ADMIN, from the script editor. Who could not use this if they tried.
 *
 * An account with no email address on file gets the same reassuring answer as
 * everybody else and no email, which is the correct behaviour at the sign-in
 * screen and a problem the clinic should fix before somebody needs it at 2am.
 */
function crescResetReadiness() {
  crescEditorOnly_('crescResetReadiness');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = [], staffMissing = 0, patientMissing = 0, staffOk = 0, patientOk = 0;

  var users = ss.getSheetByName('Users');
  if (users && users.getLastRow() > 1) {
    var u = users.getDataRange().getValues();
    for (var i = 1; i < u.length; i++) {
      if (!u[i][0]) continue;
      var st = String(u[i][3] || '').trim().toUpperCase();
      if (st && st !== 'ACTIVE') continue;
      if (cresc_looksLikeEmail_(u[i][4])) { staffOk++; continue; }
      staffMissing++;
      out.push('  STAFF    ' + String(u[i][0]).trim() + '  (' +
               String(u[i][2] || '') + ') — no email');
    }
  }

  var patients = ss.getSheetByName('Patients');
  if (patients && patients.getLastRow() > 1) {
    var p = patients.getDataRange().getValues();
    for (var j = 1; j < p.length; j++) {
      if (!p[j][0]) continue;
      if (cresc_looksLikeEmail_(p[j][16])) { patientOk++; continue; }
      patientMissing++;
      if (patientMissing <= 30) {
        out.push('  PATIENT  ' + String(p[j][0]).trim() + '  ' +
                 String(p[j][2] || '') + ' — no email');
      }
    }
  }

  var report =
    'Forgot-password readiness\n' +
    '  staff:    ' + staffOk + ' can reset by email, ' + staffMissing + ' cannot\n' +
    '  patients: ' + patientOk + ' can reset by email, ' + patientMissing + ' cannot\n' +
    (out.length ? '\n' + out.join('\n') : '\nEvery account has an email address.') +
    (patientMissing > 30 ? '\n  … and ' + (patientMissing - 30) + ' more patients.' : '');
  Logger.log(report);
  return report;
}
