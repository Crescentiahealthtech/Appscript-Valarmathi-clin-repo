// ============================================================================
// Auth_MFA_Admin.gs — Crescentia HealthTech / CresRx
// Two-step sign-in, set up from the application instead of the editor.
// ----------------------------------------------------------------------------
// WHY THIS EXISTS
//
// The only way to give somebody an authenticator was enrolMFA("nurse1") in the
// Apps Script editor. That meant the owner, at a laptop, reading a 32-letter
// secret aloud or pasting an otpauth:// URL into a QR website — which hands
// the second factor to a stranger's server. And nothing checked that the phone
// had actually taken the secret, so a mistyped letter was discovered at the
// user's next sign-in, when they were locked out.
//
// Here:
//   * the QR is drawn on THIS server (DS_QR_Lib.gs), so the secret never
//     leaves the deployment;
//   * the secret is written to the Users sheet only after the phone has read
//     it back as a valid six-digit code, so a failed scan changes nothing;
//   * a member of staff can set up their own phone (password required), and
//     an administrator can set up or remove anybody they may manage — the
//     same doctor/administrator boundary as the staff list.
//
// The secret is shown once, during enrolment, to the person holding the
// phone. It is never readable afterwards; diagnoseMFA() says whether it is
// valid without printing it.
// ============================================================================

/** How long a started enrolment waits for its confirming code. */
var CRESC_MFA_ENROL_S = 600;

/** Wrong codes allowed against one enrolment before it must be restarted. */
var CRESC_MFA_ENROL_TRIES = 5;

/**
 * 160 random bits as Base32 (RFC 4226's recommended length).
 *
 * From Utilities.getUuid(), not Math.random(): the second is a predictable
 * PRNG, and this secret is the whole second factor. One byte per character,
 * and 256 is a multiple of 32, so the mapping has no bias.
 */
function mfa_newSecret_() {
  var alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567';
  var secret = '';
  while (secret.length < 32) {
    var raw = Utilities.getUuid().replace(/-/g, '');
    for (var k = 0; k + 1 < raw.length && secret.length < 32; k += 2) {
      secret += alphabet.charAt(parseInt(raw.substr(k, 2), 16) % alphabet.length);
    }
  }
  return secret;
}

/** The name the authenticator app shows above the code. */
function mfa_issuer_() {
  var issuer = 'CresRx';
  try {
    issuer = PropertiesService.getScriptProperties().getProperty('CLINIC_NAME') || issuer;
  } catch (e) {}
  return issuer;
}

function mfa_otpauthUrl_(username, secret) {
  var issuer = mfa_issuer_();
  return 'otpauth://totp/' + encodeURIComponent(issuer + ':' + username) +
         '?secret=' + secret +
         '&issuer=' + encodeURIComponent(issuer) +
         '&algorithm=SHA1&digits=6&period=30';
}

/** The otpauth URL as a QR image (data: URI), drawn server-side. '' on failure. */
function mfa_qrDataUrl_(url) {
  try {
    if (typeof DSX_QR !== 'function') return '';
    var qr = DSX_QR(0, 'M');
    qr.addData(url);
    qr.make();
    return qr.createDataURL(5, 4);
  } catch (e) {
    return '';
  }
}

/** First Users row for a username: {rowNo, row} or null. */
function mfa_findUser_(sh, username) {
  var want = String(username || '').trim().toUpperCase();
  if (!want) return null;
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var u = data[i][0];
    if (u === null || u === undefined || String(u).trim() === '') continue;
    if (String(u).trim().toUpperCase() === want) return { rowNo: i + 1, row: data[i] };
  }
  return null;
}

/**
 * Whether ACTOR may set up or remove MFA on TARGET. null when allowed,
 * otherwise the refusal to return.
 *
 * Yourself: always (a password is asked for separately).
 * Anybody else: admin.users, and for a doctor or administrator the owner's
 * admin.users.elevated — the same line crescListStaffAccounts draws.
 */
function mfa_refusal_(actor, targetUsername, targetRole, what) {
  var self = String(actor.username || '').toUpperCase() === String(targetUsername || '').toUpperCase();
  if (self) return null;
  if (actor.permissions.indexOf('admin.users') === -1) {
    return { success: false, code: 'FORBIDDEN',
             message: 'You can set up two-step sign-in for your own account only.' };
  }
  if (cresc_roleIsElevated_(targetRole) &&
      actor.permissions.indexOf('admin.users.elevated') === -1) {
    return cresc_elevationRefusal_(what, targetRole);
  }
  return null;
}

function mfa_roleOf_(row) {
  return (typeof crescRole_ === 'function') ? crescRole_(row[2]) : String(row[2] || '').toLowerCase();
}

/**
 * FRONTEND ENTRY. Starts setting up an authenticator.
 *
 * Nothing is written yet. The new secret waits in the cache until the phone
 * proves it has it (mfaConfirmEnrolment); until then the account keeps
 * whatever it had, so a scan that goes wrong costs nothing.
 *
 * @param {{username?:string, currentPassword?:string}} payload
 *        username  whose phone; blank means the caller's own.
 *        currentPassword  required when it is the caller's own account — a
 *        session left open on a desk must not be enough to move the second
 *        factor onto somebody else's phone.
 * @param {string} sessionToken
 * @return {{success:boolean, ticket:string, username:string, secret:string,
 *           otpauthUrl:string, qr:string, issuer:string, replacing:boolean}}
 */
function mfaBeginEnrolment(payload, sessionToken) {
  try {
    var actor = crescRequire_(sessionToken);
    var p = payload || {};
    var target = String(p.username || actor.username || '').trim();
    var self = target.toUpperCase() === String(actor.username || '').toUpperCase();

    var sh = cresc_usersSheet_();
    var hit = mfa_findUser_(sh, target);
    if (!hit) {
      return { success: false, code: 'NO_USER',
               message: self ? 'Two-step sign-in is for staff accounts.'
                             : 'There is no staff account "' + target + '".' };
    }
    target = String(hit.row[0]).trim();
    var role = mfa_roleOf_(hit.row);

    var refused = mfa_refusal_(actor, target, role, 'set up two-step sign-in for');
    if (refused) return refused;

    if (self) {
      var held = crescAuthGuard_(target);
      if (held) return held;
      var pv = crescPwdVerify_(String(p.currentPassword || ''), hit.row[1]);
      if (!pv.ok) {
        var warn = crescAuthFailed_(target, role, 'bad password at MFA setup');
        return { success: false, code: 'BAD_PASSWORD',
                 message: (warn.indexOf('locked') !== -1 || warn.indexOf('left') !== -1)
                   ? warn : 'That is not your current password.' };
      }
    }

    var secret = mfa_newSecret_();
    var ticket = Utilities.getUuid();
    CacheService.getScriptCache().put('MFAE_' + ticket, JSON.stringify({
      u: target, s: secret, by: actor.username, tries: 0
    }), CRESC_MFA_ENROL_S);

    var url = mfa_otpauthUrl_(target, secret);
    return {
      success: true,
      ticket: ticket,
      username: target,
      issuer: mfa_issuer_(),
      secret: secret.replace(/(.{4})/g, '$1 ').trim(),
      otpauthUrl: url,
      qr: mfa_qrDataUrl_(url),
      replacing: !!String(hit.row[5] || '').replace(/\s/g, ''),
      expiresInS: CRESC_MFA_ENROL_S,
      message: ''
    };
  } catch (err) {
    return { success: false, message: String((err && err.message) || err).replace('FORBIDDEN: ', '') };
  }
}

/**
 * FRONTEND ENTRY. Finishes an enrolment: the six digits the phone now shows.
 *
 * Only a code that verifies against the new secret writes it to column F.
 * The started enrolment belongs to whoever started it — another session
 * cannot finish it.
 *
 * @param {string} ticket        from mfaBeginEnrolment
 * @param {string} code          six digits
 * @param {string} sessionToken
 */
function mfaConfirmEnrolment(ticket, code, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    var actor = crescRequire_(sessionToken);
    var cache = CacheService.getScriptCache();
    var key = 'MFAE_' + String(ticket || '');
    var pending = null;
    try { pending = ticket ? JSON.parse(cache.get(key) || 'null') : null; } catch (e) { pending = null; }
    if (!pending || String(pending.by).toUpperCase() !== String(actor.username || '').toUpperCase()) {
      return { success: false, code: 'EXPIRED',
               message: 'This set-up has expired. Start again.' };
    }

    var check = processTOTP_(pending.s, code);
    if (!check.success) {
      pending.tries = (pending.tries || 0) + 1;
      if (pending.tries >= CRESC_MFA_ENROL_TRIES) {
        cache.remove(key);
        return { success: false, code: 'EXPIRED',
                 message: 'Too many codes that did not match. Start again, and ' +
                          'delete the half-added entry in the authenticator app first.' };
      }
      cache.put(key, JSON.stringify(pending), CRESC_MFA_ENROL_S);
      return { success: false, code: check.code || 'BAD_CODE',
               message: check.code === 'BAD_CODE'
                 ? 'That code does not match. Check the phone clock is set automatically ' +
                   'and enter the code now showing.'
                 : check.message };
    }

    lock.waitLock(10000);
    var sh = cresc_usersSheet_();
    var hit = mfa_findUser_(sh, pending.u);
    if (!hit) return { success: false, code: 'NO_USER', message: 'The account no longer exists.' };
    // Re-checked at the write: the actor's rights may have changed in ten minutes.
    var refused = mfa_refusal_(actor, pending.u, mfa_roleOf_(hit.row), 'set up two-step sign-in for');
    if (refused) return refused;

    if (sh.getLastColumn() < 6) sh.getRange(1, 6).setValue('MFA_Secret');
    // Text format, or a secret that happens to look numeric is reformatted by
    // Sheets and no longer matches what the phone was given.
    sh.getRange(hit.rowNo, 6).setNumberFormat('@').setValue(pending.s);
    SpreadsheetApp.flush();
    cache.remove(key);

    var self = String(pending.u).toUpperCase() === String(actor.username || '').toUpperCase();
    crescAuthAudit_(CRESC_AUTH_EVENTS.MFA_ENROLLED, pending.u, mfa_roleOf_(hit.row),
                    { by: self ? 'self' : actor.username });
    return { success: true, username: pending.u,
             message: 'Two-step sign-in is on for ' + pending.u + '. From the next sign-in ' +
                      'the code from the phone will be asked for after the password.' };
  } catch (err) {
    return { success: false, message: String((err && err.message) || err).replace('FORBIDDEN: ', '') };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * FRONTEND ENTRY. Switches two-step sign-in off for somebody who has lost
 * their phone, so they can sign in with the password and set up a new one.
 *
 * Administrators only, and never on your own account: a session that can
 * remove its own second factor makes the second factor decorative. The owner
 * can still use the editor (enrolMFA) for their own.
 *
 * @param {string} username
 * @param {string} reason        recorded in the audit — "lost phone"
 * @param {string} sessionToken
 */
function mfaRemove(username, reason, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    var actor = crescRequire_(sessionToken, 'admin.users');
    var why = String(reason || '').trim();
    if (why.length < 3) return { success: false, message: 'Give a reason — it goes in the audit log.' };
    if (String(username || '').trim().toUpperCase() === String(actor.username || '').toUpperCase()) {
      return { success: false, code: 'SELF',
               message: 'Another administrator has to remove your own two-step sign-in.' };
    }

    lock.waitLock(10000);
    var sh = cresc_usersSheet_();
    var hit = mfa_findUser_(sh, username);
    if (!hit) return { success: false, code: 'NO_USER', message: 'There is no staff account "' + username + '".' };
    var target = String(hit.row[0]).trim();
    var role = mfa_roleOf_(hit.row);
    var refused = mfa_refusal_(actor, target, role, 'remove two-step sign-in from');
    if (refused) return refused;

    if (!String(hit.row[5] || '').replace(/\s/g, '')) {
      return { success: true, username: target, message: target + ' did not have two-step sign-in.' };
    }
    sh.getRange(hit.rowNo, 6).setValue('');
    SpreadsheetApp.flush();
    crescAuthAudit_(CRESC_AUTH_EVENTS.MFA_REMOVED, target, role,
                    { by: actor.username, reason: why.slice(0, 200) });
    return { success: true, username: target,
             message: 'Two-step sign-in removed for ' + target + '. They sign in with the ' +
                      'password alone until a new phone is set up.' };
  } catch (err) {
    return { success: false, message: String((err && err.message) || err).replace('FORBIDDEN: ', '') };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * FRONTEND ENTRY. Whether the caller's own account has two-step sign-in, for
 * the account menu.
 */
function mfaMyStatus(sessionToken) {
  try {
    var actor = crescRequire_(sessionToken);
    var hit = mfa_findUser_(cresc_usersSheet_(), actor.username);
    if (!hit) return { success: true, staff: false, enrolled: false };
    var raw = String(hit.row[5] || '');
    return { success: true, staff: true,
             enrolled: !!raw.replace(/\s/g, ''),
             valid: mfa_normaliseSecret_(raw).ok };
  } catch (err) {
    return { success: false, message: String((err && err.message) || err).replace('FORBIDDEN: ', '') };
  }
}
