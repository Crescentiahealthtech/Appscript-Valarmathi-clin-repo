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
// The person names their account. A random temporary password is written,
// flagged must-change, and EMAILED to the address already on the record —
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
//    Without a cap this is a button that mails a working credential and
//    invalidates the real one. Anybody who knows a colleague's user id could
//    lock them out repeatedly, from a signed-out page, for as long as they
//    cared to keep clicking. Three per account per hour, and a deployment-wide
//    ceiling so a script cannot walk the patient id range.
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
    'folder — then sign in with it and choose a new password. If nothing ' +
    'arrives within a few minutes, please contact the clinic.'
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

        // Written BEFORE the send, then rolled forward: a temporary password
        // that was mailed but not stored is an account nobody can get into.
        cresc_writeCredential_(cresc_usersSheet_(), i + 1, temp, true,
                               'Must_Change', 'Password_Updated_At');
        SpreadsheetApp.flush();

        var sent = cresc_sendResetEmail_(uEmail, want, temp, clinic,
                                         String(udata[i][2] || 'staff'));
        crescAuthAudit_(CRESC_AUTH_EVENTS.PASSWORD_RESET, want,
                        String(udata[i][2] || ''),
                        { by: 'self-service', emailed: sent.ok,
                          to: cresc_maskEmail_(uEmail),
                          problem: sent.ok ? '' : sent.message });

        if (!sent.ok) {
          // The password HAS changed, so the old one no longer works and
          // saying "check your email" would leave them locked out with no
          // idea why. This is the one case that gets a different answer,
          // because it is about this system failing rather than about
          // whether the account exists.
          return { success: false,
                   message: 'Your password was reset but the email could not be ' +
                            'sent (' + sent.message + '). Contact the clinic — ' +
                            'an administrator can give you the new one.' };
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

        cresc_writeCredential_(cresc_patientsSheet_(), j + 1, temp, true,
                               'Portal_Must_Change', 'Portal_Password_Updated_At');
        SpreadsheetApp.flush();

        var psent = cresc_sendResetEmail_(pEmail, want, temp, clinic, 'patient');
        crescAuthAudit_(CRESC_AUTH_EVENTS.PASSWORD_RESET, want, 'patient',
                        { by: 'self-service', emailed: psent.ok,
                          to: cresc_maskEmail_(pEmail),
                          problem: psent.ok ? '' : psent.message });

        if (!psent.ok) {
          return { success: false,
                   message: 'Your password was reset but the email could not be ' +
                            'sent (' + psent.message + '). Please contact the clinic.' };
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
    'away. This temporary one stops working as soon as you do.\n\n' +
    'IF YOU DID NOT ASK FOR THIS, your old password has already stopped ' +
    'working and somebody else may have requested it. Contact the clinic' +
    (phone ? ' on ' + phone : '') + ' as soon as you can.\n\n' +
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
          'a new password straight away. This temporary one stops working as soon ' +
          'as you do.</p>' +
        '<p style="margin:0 0 14px;padding:12px 14px;background:#fef2f2;' +
          'border-left:4px solid #dc2626;color:#7f1d1d;">' +
          '<strong>If you did not ask for this</strong>, your old password has ' +
          'already stopped working and somebody else may have requested it. ' +
          'Contact the clinic' + (phone ? ' on ' + cresc_htmlEsc_(phone) : '') +
          ' as soon as you can.</p>' +
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
