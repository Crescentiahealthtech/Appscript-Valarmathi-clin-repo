// ==========================================
// 🧠 CENTRAL AUTH + RBAC ENGINE
// ==========================================

/**
 * Every outcome of this function is now recorded in Audit_Log, and five
 * failures inside fifteen minutes hold the account. See Auth_Audit.gs for
 * why, and for what deliberately is NOT recorded.
 *
 * One message covers "no such user" and "wrong password", on purpose. The
 * two used to be distinguishable — "User ID not found." versus "Incorrect
 * staff password." — which turns this function into an oracle for testing
 * whether a patient ID is real, and patient IDs are printed on every bill
 * and every barcode label. The audit row still records which it actually
 * was, so support can tell the difference and an attacker cannot.
 */
function verifyLogin(credentials) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usernameInput = credentials.username.toString().trim();
  const passwordInput = credentials.password.toString().trim();

  // Before any comparison: an account being guessed at costs one cache read.
  const held = crescAuthGuard_(usernameInput);
  if (held) return held;

  const GENERIC_FAIL = 'Incorrect ID or password.';

  // ==========================================
  // 1. HOSPITAL STAFF LOGIN
  // ==========================================
  const userSheet = ss.getSheetByName('Users');

  if (userSheet) {
    const userData = userSheet.getDataRange().getValues();

    for (let i = 1; i < userData.length; i++) {
      const storedUsername = userData[i][0];
      const storedPassword = userData[i][1];
      const storedRole = userData[i][2];
      const isActive = userData[i][3];

      if (!storedUsername) continue;

      if (storedUsername.toString().trim().toUpperCase() === usernameInput.toUpperCase()) {
        
        // Check Active Status
        if (isActive && isActive.toString().toLowerCase() !== 'active') {
          crescAuthAudit_(CRESC_AUTH_EVENTS.DISABLED, storedUsername, storedRole,
                          { status: String(isActive) });
          return { success: false, message: "Account disabled." };
        }

        // Password validation — against a stored digest, never a stored
        // password. See Auth_Credentials.gs for why a plain-text cell is
        // refused outright rather than accepted once and hashed after.
        const check = crescPwdVerify_(passwordInput, storedPassword);

        if (check.legacy) {
          crescAuthAudit_(CRESC_AUTH_EVENTS.LEGACY_REFUSED, storedUsername, storedRole, {});
          return { success: false, code: 'LEGACY_CREDENTIAL',
                   message: 'This account still has an old, unprotected password. ' +
                            'An administrator must reset it before you can sign in ' +
                            '— ask them to run crescMigrateCredentials().' };
        }

        if (check.ok) {
          const role = storedRole.toString().trim().toLowerCase();

          // A temporary password gets you exactly one screen: the one that
          // replaces it. Issuing a session here would mean a password handed
          // over at a desk, or sitting in a migration log, stayed usable for
          // as long as nobody got round to changing it.
          if (cresc_mustChange_(userSheet, i + 1, 'Must_Change')) {
            crescAuthAudit_(CRESC_AUTH_EVENTS.MUST_CHANGE, storedUsername, storedRole, {});
            return { success: false, code: 'MUST_CHANGE',
                     username: storedUsername.toString().trim(),
                     message: 'Your password was set for you and has to be changed ' +
                              'before you can sign in.' };
          }

          // The real password worked, so a reset somebody else may have
          // asked for is no longer needed: the temporary one stops here.
          if (typeof cresc_clearPendingReset_ === 'function') {
            cresc_clearPendingReset_(userSheet, i + 1, false);
          }
          return cresc_staffSignIn_(storedUsername, role, userData[i][5], 'password');
        } else if (typeof cresc_pendingResetMatches_ === 'function' &&
                   cresc_pendingResetMatches_(userSheet, i + 1, passwordInput, false)) {
          // The emailed temporary password. It opens exactly one screen: the
          // one that replaces it. The old password kept working until now.
          crescAuthAudit_(CRESC_AUTH_EVENTS.MUST_CHANGE, storedUsername, storedRole,
                          { via: 'reset email' });
          return { success: false, code: 'MUST_CHANGE',
                   username: storedUsername.toString().trim(),
                   message: 'You signed in with a temporary password. Choose your ' +
                            'own password to continue.' };
        } else {
          const warn = crescAuthFailed_(storedUsername, storedRole, 'bad staff password');
          // The counter's warning is worth showing; which half was wrong is not.
          return { success: false,
                   message: warn.indexOf('locked') !== -1 || warn.indexOf('left') !== -1
                            ? warn : GENERIC_FAIL };
        }
      }
    }
  }

  // ==========================================
  // 2. PATIENT LOGIN (DYNAMIC PASSWORD ENGINE)
  // ==========================================
  const patientSheet = ss.getSheetByName('Patients');
  if (!patientSheet) {
    return { success: false, message: "Patients database missing." };
  }

  const patientData = patientSheet.getDataRange().getValues();

  for (let i = 1; i < patientData.length; i++) {
    const patientID = patientData[i][0];
    if (!patientID) continue;

    if (patientID.toString().trim().toUpperCase() === usernameInput.toUpperCase()) {
      
      let rawName = patientData[i][2] ? patientData[i][2].toString().trim() : "XXX";

      // THE DERIVATION IS GONE. This block used to rebuild the password from
      // the first three letters of the name and the birth year and compare
      // it. Both inputs are printed on the patient's own documents and the
      // patient ID is the barcode on the same page, so the portal was open to
      // anyone who had ever held a prescription. Portal passwords are now
      // random, stored as a digest, and changed at first sign-in.
      const pcheck = crescPwdVerify_(passwordInput, patientData[i][1]);

      if (pcheck.legacy) {
        crescAuthAudit_(CRESC_AUTH_EVENTS.LEGACY_REFUSED, patientID, 'patient', {});
        return { success: false, code: 'LEGACY_CREDENTIAL',
                 message: 'Your portal password has to be reset at the clinic before ' +
                          'you can sign in. Please ask at reception.' };
      }

      if (!pcheck.ok && typeof cresc_pendingResetMatches_ === 'function' &&
          cresc_pendingResetMatches_(patientSheet, i + 1, passwordInput, true)) {
        crescAuthAudit_(CRESC_AUTH_EVENTS.MUST_CHANGE, patientID, 'patient', { via: 'reset email' });
        return { success: false, code: 'MUST_CHANGE',
                 username: patientID.toString().trim().toUpperCase(),
                 message: 'You signed in with a temporary password. Choose your ' +
                          'own password to continue.' };
      }

      if (pcheck.ok) {
        if (typeof cresc_clearPendingReset_ === 'function') {
          cresc_clearPendingReset_(patientSheet, i + 1, true);
        }
        if (cresc_mustChange_(patientSheet, i + 1, 'Portal_Must_Change')) {
          crescAuthAudit_(CRESC_AUTH_EVENTS.MUST_CHANGE, patientID, 'patient', {});
          return { success: false, code: 'MUST_CHANGE',
                   username: patientID.toString().trim().toUpperCase(),
                   message: 'The password you were given at the clinic has to be ' +
                            'changed before you can sign in.' };
        }
        // Patients get a real session token too. Without one the portal had no
        // way to prove who it was, so getUserProfile() could not be session-
        // checked. The session username IS the patient ID, which is what
        // getUserProfile() compares against to keep a patient to their own row.
        const patientKey = patientID.toString().trim().toUpperCase();
        const patientToken = issueSession_({
          username: patientKey,
          role: 'patient',
          doctorId: "",
          name: rawName
        });
        crescAuthPassed_(patientKey, 'patient', 'portal');
        return {
          success: true,
          role: 'patient',
          portal: 'patient',
          username: patientKey,
          sessionToken: patientToken,
          message: "Welcome Patient"
        };
      } else {
        // The format hint went when the format did. It used to print the
        // derivation rule outright — first three letters of the name plus the
        // birth year — to anyone holding a patient ID, which is anyone holding
        // a printed bill.
        const warn = crescAuthFailed_(patientID, 'patient', 'bad portal password');
        return { success: false,
                 message: warn.indexOf('locked') !== -1 || warn.indexOf('left') !== -1
                          ? warn : GENERIC_FAIL };
      }
    }
  }

  // ==========================================
  // 3. USER NOT FOUND
  // ==========================================
  // Counted like any other failure, so walking a range of patient IDs runs
  // into the same lock a password attack does, and answered with the same
  // wording so the walk learns nothing from the replies.
  crescAuthFailed_(usernameInput, '', 'no such user');
  crescAuthAudit_(CRESC_AUTH_EVENTS.UNKNOWN_USER, usernameInput, '', {});
  return { success: false, message: GENERIC_FAIL };
}

// ==========================================
// GOOGLE SSO AUTHENTICATION ENGINE
// ==========================================
function verifyGoogleLogin(payload) {
  try {
    // IT USED TO TAKE AN EMAIL ADDRESS. The browser did the Google sign-in and
    // then told this function who the user was, and this function believed
    // it: google.script.run.verifyGoogleLogin('admin@clinic') from any
    // browser console returned an administrator's session — no password, no
    // second factor. The browser now sends the ID token Google gave it, and
    // Google (not the browser) says whose it is.
    var idToken = (payload && typeof payload === 'object') ? String(payload.idToken || '') : '';
    if (!idToken) {
      return { success: false, message: 'Google sign-in did not complete. Please try again.' };
    }
    var who = cresc_verifyGoogleIdToken_(idToken);
    if (!who.ok) {
      crescAuthAudit_(CRESC_AUTH_EVENTS.FAILED, '', '', { method: 'google', reason: who.message });
      return { success: false, message: who.message };
    }
    var userEmail = who.email;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName('Users');
    if (!userSheet) throw new Error("Users staff registry tab is missing.");

    const userData = userSheet.getDataRange().getValues();

    // Check email alignment against column E (index 4) of your Users dataset
    for (let i = 1; i < userData.length; i++) {
      const storedUsername = userData[i][0];
      const storedRole = userData[i][2];
      const isActive = userData[i][3];
      const storedEmail = userData[i][4];

      if (!storedEmail) continue;

      if (storedEmail.toString().trim().toLowerCase() === userEmail) {

        if (isActive && isActive.toString().toLowerCase() !== 'active') {
          crescAuthAudit_(CRESC_AUTH_EVENTS.DISABLED, storedUsername, storedRole,
                          { method: 'google', email: String(userEmail), status: String(isActive) });
          return { success: false, message: "Access Denied: This staff credential context is deactivated." };
        }

        const role = storedRole.toString().trim().toLowerCase();
        return cresc_staffSignIn_(storedUsername, role, userData[i][5], 'google');
      }
    }

    // A Google address that is not on the Users sheet is a real sign-in
    // attempt by a real person and belongs in the record; unlike a password
    // guess, naming the address back is safe — whoever is reading it owns it.
    crescAuthAudit_(CRESC_AUTH_EVENTS.UNKNOWN_USER, String(userEmail), '',
                    { method: 'google' });
    return {
      success: false,
      message: "Access Denied: The email " + userEmail + " is not registered in CresRx. Contact Admin."
    };

  } catch (error) {
    return { success: false, message: "System Security Fault: " + error.toString() };
  }
}

/**
 * Asks Google whose ID token this is. accounts:lookup checks the signature,
 * the expiry and that the token was issued for THIS Firebase project (the one
 * the API key belongs to), so a token from another site is refused.
 *
 * The web API key is the same public value Auth.html already ships; override
 * it with the script property FIREBASE_API_KEY if the project changes.
 *
 * @return {{ok:boolean, email:string, message:string}}
 */
var CRESC_FIREBASE_API_KEY_DEFAULT = 'AIzaSyDJSMsLhlbsGwC1cMR9mNM4Vy2kQ1MmGy4';

function cresc_verifyGoogleIdToken_(idToken) {
  var key = '';
  try { key = PropertiesService.getScriptProperties().getProperty('FIREBASE_API_KEY') || ''; } catch (e) {}
  key = key || CRESC_FIREBASE_API_KEY_DEFAULT;
  try {
    var resp = UrlFetchApp.fetch(
      'https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + encodeURIComponent(key), {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ idToken: idToken }),
        muteHttpExceptions: true
      });
    if (resp.getResponseCode() !== 200) {
      return { ok: false, email: '', message: 'Google could not confirm this sign-in. Please try again.' };
    }
    var body = JSON.parse(resp.getContentText() || '{}');
    var u = (body.users || [])[0];
    if (!u || !u.email) {
      return { ok: false, email: '', message: 'Google did not return an email address for this account.' };
    }
    if (u.emailVerified !== true) {
      return { ok: false, email: '', message: 'This Google account\'s email address is not verified.' };
    }
    if (u.disabled === true) {
      return { ok: false, email: '', message: 'This Google account is disabled.' };
    }
    return { ok: true, email: String(u.email).trim().toLowerCase(), message: '' };
  } catch (e) {
    return { ok: false, email: '', message: 'Google sign-in could not be verified: ' + e.message };
  }
}

// ==========================================
// STAFF SIGN-IN: ONE PLACE A SESSION IS ISSUED
// ------------------------------------------
// verifyLogin used to issue the session BEFORE the second factor. The browser
// then showed the MFA screen and waited — but it already held a working
// token, so a stolen password alone was a full sign-in for anyone who read
// the reply instead of typing six digits. Now an account with MFA enrolled
// gets a short-lived ticket, and only verifyMFA turns a ticket plus a valid
// code into a session.
// ==========================================

var CRESC_MFA_TICKET_S = 300;        // five minutes to type the code
var CRESC_MFA_TICKET_TRIES = 5;      // wrong codes before the ticket is spent

function cresc_hasMfa_(rawSecret) {
  return !!String(rawSecret || '').replace(/\s/g, '');
}

function cresc_issueStaffSession_(username, role, method) {
  var doc = resolveDoctorByUsername_(username);
  var token = issueSession_({
    username: username,
    role: role,
    doctorId: doc ? doc.doctorId : "",
    name: doc ? doc.name : username
  });
  crescAuthPassed_(username, role, method);
  return {
    success: true,
    role: role,
    portal: 'hospital',
    username: String(username),
    displayName: doc ? doc.name : String(username),
    doctorId: doc ? doc.doctorId : "",
    sessionToken: token,
    message: "Welcome " + role
  };
}

function cresc_staffSignIn_(username, role, rawSecret, method) {
  if (!cresc_hasMfa_(rawSecret)) return cresc_issueStaffSession_(username, role, method);

  var ticket = Utilities.getUuid();
  CacheService.getScriptCache().put('MFAT_' + ticket, JSON.stringify({
    u: String(username).trim(), role: role, method: method, tries: 0
  }), CRESC_MFA_TICKET_S);
  return {
    success: true,
    mfaRequired: true,
    mfaTicket: ticket,
    role: role,
    portal: 'hospital',
    username: String(username).trim(),
    message: 'Enter the 6-digit code from your authenticator.'
  };
}

// ==========================================
// MFA / TOTP ENGINE (Google Authenticator)
// ------------------------------------------
// Users sheet layout this engine relies on:
//   A Username | B Password | C Role | D Status | E Email | F MFA_Secret
//
// WHY THIS WAS REWRITTEN
//   One staff login (a nurse) could not get past "Invalid or Expired MFA
//   Code" while the doctor and admin logins on the same device and the same
//   clock worked. The algorithm below is unchanged and correct, so the fault
//   was never the maths — it was the SECRET, and the old code could not say
//   so, because:
//
//     * base32ToBytes() silently skipped every character outside the Base32
//       alphabet. A secret typed or pasted with 0/1/8/9 in it (the four
//       digits Base32 does not use, and the four most commonly mistyped for
//       O/I/B/g) therefore produced a SHORTER key that was wrong in a way
//       nobody could see. The phone, given the same string, rejects those
//       characters too — so the two sides derive different keys and every
//       code is "invalid" for ever.
//     * processTOTP() caught every error and returned the single message
//       "Verification error.", which the UI then overwrote with "Invalid or
//       Expired MFA Code" — so a malformed secret and a mistyped code looked
//       identical to the user.
//     * An empty or whitespace-only cell in column F was truthy often enough
//       (a stray space) to enter TOTP verification with no key at all.
//
//   The engine now normalises the secret, validates it, distinguishes "your
//   enrolment is broken" from "that code is wrong", and ships
//   diagnoseMFA()/enrolMFA() so an administrator can see and fix the stored
//   secret without guessing.
// ==========================================

/** Drift windows accepted either side of now. 1 = ±30s, 2 = ±60s.
 *  Google's servers keep exact time; the phone is what drifts, and a cheap
 *  handset that has not synced in a week is routinely 40-60 seconds out. */
var MFA_DRIFT_WINDOWS = 2;

/**
 * Canonical form of a stored Base32 secret.
 * Google Authenticator ignores case, spaces and "=" padding, so we must too —
 * secrets are routinely stored as "abcd efgh ijkl mnop".
 *
 * @return {{ok:boolean, secret:string, message:string}}
 */
function mfa_normaliseSecret_(raw) {
  var s = String(raw === null || raw === undefined ? '' : raw)
    .replace(/[\s\-_]/g, '')      // spaces, dashes and underscores are display only
    .replace(/=+$/, '')           // padding carries no bits
    .toUpperCase();

  if (!s) return { ok: false, secret: '', message: 'No authenticator secret is enrolled.' };

  var bad = s.replace(/[A-Z2-7]/g, '');
  if (bad) {
    // Named explicitly: these are the characters that silently corrupted the key.
    return { ok: false, secret: '',
             message: 'The stored authenticator secret contains character(s) that are ' +
                      'not valid Base32 (' + bad.split('').filter(function (c, i, a) {
                        return a.indexOf(c) === i;
                      }).join(' ') + '). Base32 uses A-Z and 2-7 only — it never ' +
                      'contains 0, 1, 8 or 9. Re-enrol this user.' };
  }
  if (s.length < 16) {
    return { ok: false, secret: '',
             message: 'The stored authenticator secret is too short (' + s.length +
                      ' characters). Re-enrol this user.' };
  }
  return { ok: true, secret: s, message: '' };
}

/**
 * Verifies a 6-digit authenticator code for a staff username.
 * Contract: { success:boolean, message:string, code?:string }.
 *
 * On success the reply IS the sign-in: it carries the session token, which
 * verifyLogin / verifyGoogleLogin withhold from any account with MFA enrolled.
 *
 * code 'TICKET_EXPIRED' no password step before this, or it timed out / was spent
 * code 'BAD_SECRET'     the enrolment itself is broken — an admin must fix it
 * code 'BAD_CODE'       the secret is fine, the six digits are not
 */
function verifyMFA(username, userCode, mfaTicket) {
  try {
    // A code proves nothing without the password (or Google) step before it.
    var cache = CacheService.getScriptCache();
    var tkey = 'MFAT_' + String(mfaTicket || '');
    var ticket = null;
    try { ticket = mfaTicket ? JSON.parse(cache.get(tkey) || 'null') : null; } catch (e) { ticket = null; }
    if (!ticket || String(ticket.u).toUpperCase() !== String(username || '').trim().toUpperCase()) {
      return { success: false, code: 'TICKET_EXPIRED',
               message: 'This sign-in has expired. Enter your password again.' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userSheet = ss.getSheetByName('Users');
    if (!userSheet) {
      return { success: false, code: 'NO_SHEET', message: 'Staff registry is missing.' };
    }

    var want = String(username || '').trim().toUpperCase();
    if (!want) return { success: false, code: 'NO_USER', message: 'No user to verify.' };

    var userData = userSheet.getDataRange().getValues();
    for (var i = 1; i < userData.length; i++) {
      var stored = userData[i][0];
      if (stored === null || stored === undefined || String(stored).trim() === '') continue;
      if (String(stored).trim().toUpperCase() !== want) continue;

      var norm = mfa_normaliseSecret_(userData[i][5]);   // Column F — MFA_Secret

      // Not enrolled at all: MFA is optional per user, so this is a pass.
      // A blank cell, and a cell holding only spaces, must behave identically.
      if (!norm.ok && !String(userData[i][5] || '').replace(/\s/g, '')) {
        // Unenrolled between the password and the code: the password stands.
        cache.remove(tkey);
        return cresc_issueStaffSession_(String(stored).trim(), ticket.role, ticket.method);
      }
      if (!norm.ok) {
        return { success: false, code: 'BAD_SECRET', message: norm.message };
      }

      // The second factor is the half of the sign-in an attacker has to get
      // past once they already hold the password, so a failure here is the
      // most interesting row in the whole audit — it means the first factor
      // has already gone.
      var totp = processTOTP(norm.secret, userCode);
      crescAuthAudit_(totp && totp.success ? CRESC_AUTH_EVENTS.MFA_PASSED
                                           : CRESC_AUTH_EVENTS.MFA_FAILED,
                      String(stored).trim(), String(userData[i][2] || ''),
                      { code: (totp && totp.code) || '' });
      if (!(totp && totp.success)) {
        crescAuthFailed_(String(stored).trim(), '', 'bad MFA code');
        ticket.tries = (ticket.tries || 0) + 1;
        if (ticket.tries >= CRESC_MFA_TICKET_TRIES) {
          cache.remove(tkey);
          return { success: false, code: 'TICKET_EXPIRED',
                   message: 'Too many wrong codes. Enter your password again.' };
        }
        cache.put(tkey, JSON.stringify(ticket), CRESC_MFA_TICKET_S);
        return totp;
      }
      cache.remove(tkey);
      return cresc_issueStaffSession_(String(stored).trim(), ticket.role, ticket.method);
    }
    crescAuthAudit_(CRESC_AUTH_EVENTS.UNKNOWN_USER, want, '', { stage: 'mfa' });
    return { success: false, code: 'NO_USER', message: 'User not found for MFA verification.' };
  } catch (e) {
    return { success: false, code: 'ERROR', message: 'MFA verification failed: ' + e.message };
  }
}

/**
 * @param {string} secretBase32  already normalised by mfa_normaliseSecret_,
 *                               but re-normalised here because DS_Workflow.gs
 *                               calls this directly with a raw sheet value.
 * @param {string} userToken     the six digits the user typed
 * @return {{success:boolean, message:string, code?:string}}
 */
function processTOTP(secretBase32, userToken) {
  var norm = mfa_normaliseSecret_(secretBase32);
  if (!norm.ok) return { success: false, code: 'BAD_SECRET', message: norm.message };

  // Keep leading zeros: a code of "012345" is not the number 12345. Strip
  // spaces the way a phone displays them ("012 345").
  var token = String(userToken === null || userToken === undefined ? '' : userToken)
                .replace(/\s/g, '');
  if (!/^\d{6}$/.test(token)) {
    return { success: false, code: 'BAD_CODE', message: 'Enter the 6-digit code from your authenticator.' };
  }

  try {
    var keyBytes = base32ToBytes(norm.secret);
    if (!keyBytes.length) {
      return { success: false, code: 'BAD_SECRET',
               message: 'The stored authenticator secret produced no key. Re-enrol this user.' };
    }

    var timeWindow = Math.floor(Date.now() / 1000 / 30);
    for (var i = -MFA_DRIFT_WINDOWS; i <= MFA_DRIFT_WINDOWS; i++) {
      if (generateTOTPAlgorithm(keyBytes, timeWindow + i) === token) {
        return { success: true, message: 'MFA Verified' };
      }
    }
    return { success: false, code: 'BAD_CODE',
             message: 'That code is not valid right now. Check that your phone clock is ' +
                      'set automatically, then try the next code.' };
  } catch (e) {
    return { success: false, code: 'ERROR', message: 'Verification error: ' + e.message };
  }
}

function generateTOTPAlgorithm(keyBytes, timeValue) {
  var timeBytes = new Array(8);
  for (var i = 7; i >= 0; i--) {
    // Sign-extend to the -128..127 range Java's byte[] uses. Apps Script will
    // cast 0..255 for us, but being explicit keeps the two byte arrays (time
    // and key) built the same way.
    timeBytes[i] = ((timeValue & 0xff) << 24) >> 24;
    timeValue = Math.floor(timeValue / 256);
  }
  var hmac = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_1, timeBytes, keyBytes);
  var offset = hmac[hmac.length - 1] & 0x0f;
  var binary = ((hmac[offset] & 0x7f) << 24) | ((hmac[offset + 1] & 0xff) << 16) |
               ((hmac[offset + 2] & 0xff) << 8) | (hmac[offset + 3] & 0xff);
  var otp = (binary % 1000000).toString();
  while (otp.length < 6) { otp = '0' + otp; }
  return otp;
}

function base32ToBytes(base32) {
  var alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";
  var bits = 0, value = 0, output = [];
  var s = String(base32 || '').toUpperCase();
  for (var i = 0; i < s.length; i++) {
    var val = alphabet.indexOf(s.charAt(i));
    if (val === -1) continue;
    value = (value << 5) | val;
    bits += 5;
    if (bits >= 8) {
      // Sign-extended so the array is a true Java byte[] on both sides.
      output.push(((((value >>> (bits - 8)) & 255) << 24) >> 24));
      bits -= 8;
    }
  }
  return output;
}

// ==========================================
// MFA ADMINISTRATION
// Run these from the Apps Script editor. They are the supported way to fix a
// login that cannot get past two-factor.
// ==========================================

/**
 * ADMIN. Reports exactly what is wrong with one user's MFA enrolment, without
 * revealing the secret. Run diagnoseMFA("nurse1") in the editor.
 */
function diagnoseMFA(username) {
  crescEditorOnly_('diagnoseMFA');
  var out = [];
  var want = String(username || '').trim().toUpperCase();
  out.push('MFA DIAGNOSIS — ' + (want || '(no username given)'));
  out.push('');

  if (!want) { out.push('Pass the username, e.g. diagnoseMFA("nurse1").'); }
  else {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    if (!sh) out.push('FAIL  There is no Users sheet.');
    else {
      var data = sh.getDataRange().getValues();
      var hits = [];
      for (var i = 1; i < data.length; i++) {
        var u = data[i][0];
        if (u === null || u === undefined || String(u).trim() === '') continue;
        if (String(u).trim().toUpperCase() === want) hits.push(i + 1);
      }

      if (!hits.length) {
        out.push('FAIL  No row in Users has username "' + want + '".');
      } else {
        if (hits.length > 1) {
          out.push('WARN  ' + hits.length + ' rows share this username (rows ' +
                   hits.join(', ') + '). Login always uses the FIRST one, row ' +
                   hits[0] + '. Delete the duplicates.');
        }
        var row = data[hits[0] - 1];
        out.push('row      ' + hits[0]);
        out.push('role     ' + String(row[2] || '(blank)'));
        out.push('status   ' + String(row[3] || '(blank — treated as active)'));

        var raw = row[5];
        if (!String(raw || '').replace(/\s/g, '')) {
          out.push('secret   (empty) — MFA is not enrolled, so login skips it.');
          out.push('');
          out.push('If this user is being asked for a code anyway, they are being');
          out.push('stopped by something other than MFA. Check the Status column.');
        } else {
          var norm = mfa_normaliseSecret_(raw);
          out.push('secret   ' + String(raw).length + ' characters stored, ' +
                   (norm.ok ? norm.secret.length + ' after normalising' : 'INVALID'));
          if (!norm.ok) {
            out.push('FAIL  ' + norm.message);
            out.push('');
            out.push('FIX   Run enrolMFA("' + want + '") to issue a fresh secret, then');
            out.push('      have the user delete the old entry in their authenticator');
            out.push('      app and scan/enter the new one.');
          } else {
            var expect = '';
            try {
              expect = generateTOTPAlgorithm(base32ToBytes(norm.secret),
                                             Math.floor(Date.now() / 1000 / 30));
            } catch (e) { expect = 'could not be computed: ' + e.message; }
            out.push('PASS  The secret is valid Base32.');
            out.push('now   The code this server expects this instant is ' + expect + '.');
            out.push('');
            out.push('If the user\'s phone shows a different number, their');
            out.push('authenticator holds a DIFFERENT secret from the one in the');
            out.push('sheet. Run enrolMFA("' + want + '") and re-enrol the device.');
          }
        }
      }
    }
  }

  var report = out.join('\n');
  Logger.log(report);
  return report;
}

/**
 * ADMIN. Issues a fresh, valid Base32 secret for one user, writes it to
 * column F, and returns the otpauth:// URL to enrol the device with. This is
 * the only supported way to create a secret: a hand-typed one is exactly how
 * an unscannable enrolment gets into the sheet.
 *
 * @param {string} username
 * @return {string} the otpauth URL, also logged
 */
function enrolMFA(username) {
  crescEditorOnly_('enrolMFA');
  var want = String(username || '').trim();
  if (!want) throw new Error('Pass the username, e.g. enrolMFA("nurse1").');

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sh) throw new Error('There is no Users sheet.');

  var data = sh.getDataRange().getValues();
  var rowNo = -1;
  for (var i = 1; i < data.length; i++) {
    var u = data[i][0];
    if (u === null || u === undefined || String(u).trim() === '') continue;
    if (String(u).trim().toUpperCase() === want.toUpperCase()) { rowNo = i + 1; break; }
  }
  if (rowNo === -1) throw new Error('No row in Users has username "' + want + '".');

  // 160 bits, the RFC 4226 recommendation, drawn from the alphabet only.
  // From Utilities.getUuid(), not Math.random(): the second is a predictable
  // PRNG, and this secret is the whole second factor. One byte per character,
  // and 256 is a multiple of 32, so the mapping has no bias.
  var alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567';
  var secret = '';
  while (secret.length < 32) {
    var raw = Utilities.getUuid().replace(/-/g, '');
    for (var k = 0; k + 1 < raw.length && secret.length < 32; k += 2) {
      secret += alphabet.charAt(parseInt(raw.substr(k, 2), 16) % alphabet.length);
    }
  }

  if (sh.getLastColumn() < 6) sh.getRange(1, 6).setValue('MFA_Secret');
  // Force text, or a secret that happens to look numeric is reformatted by
  // Sheets and no longer matches what the phone was given.
  sh.getRange(rowNo, 6).setNumberFormat('@').setValue(secret);
  SpreadsheetApp.flush();

  var issuer = 'CresRx';
  try {
    issuer = PropertiesService.getScriptProperties().getProperty('CLINIC_NAME') || issuer;
  } catch (e) {}

  var url = 'otpauth://totp/' + encodeURIComponent(issuer + ':' + want) +
            '?secret=' + secret +
            '&issuer=' + encodeURIComponent(issuer) +
            '&algorithm=SHA1&digits=6&period=30';

  var report = [
    'MFA re-enrolled for ' + want + ' (Users row ' + rowNo + ').',
    '',
    'Secret (type this into the app if the QR cannot be scanned):',
    '  ' + secret.replace(/(.{4})/g, '$1 ').trim(),
    '',
    'Or build a QR from this URL:',
    '  ' + url,
    '',
    'The user MUST delete their old CresRx entry in the authenticator app',
    'first: two entries with the same name is how the wrong code gets typed.',
    'Verify with diagnoseMFA("' + want + '") once the device is set up.'
  ].join('\n');
  Logger.log(report);
  return url;
}