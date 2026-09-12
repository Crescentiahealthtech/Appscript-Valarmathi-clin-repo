// ==========================================
// 🧠 CENTRAL AUTH + RBAC ENGINE
// ==========================================

function verifyLogin(credentials) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usernameInput = credentials.username.toString().trim();
  const passwordInput = credentials.password.toString().trim();

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
          return { success: false, message: "Account disabled." };
        }

        // Password Validation
        if (storedPassword.toString().trim() === passwordInput) {
          const role = storedRole.toString().trim().toLowerCase();
          const doc  = resolveDoctorByUsername_(storedUsername);
          const token = issueSession_({
            username: storedUsername,
            role: role,
            doctorId: doc ? doc.doctorId : "",
            name: doc ? doc.name : storedUsername
          });
          return {
            success: true,
            role: role,
            portal: 'hospital',
            username: storedUsername,
            displayName: doc ? doc.name : storedUsername,
            doctorId: doc ? doc.doctorId : "",
            sessionToken: token,
            message: "Welcome " + storedRole
          };
        } else {
          return { success: false, message: "Incorrect staff password." };
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
      let rawDOB = patientData[i][5];

      // Format Name Part (First 3 chars)
      let namePart = rawName.replace(/[^a-zA-Z]/g, '');
      if (namePart.length < 3) {
        namePart = (namePart + "XXX").substring(0, 3);
      } else {
        namePart = namePart.substring(0, 3);
      }
      namePart = namePart.charAt(0).toUpperCase() + namePart.substring(1).toLowerCase();

      // Format Year Part (4 Digits)
      let yearPart = "0000";
      if (rawDOB instanceof Date) {
        yearPart = rawDOB.getFullYear().toString();
      } else if (rawDOB) {
        let dobStr = rawDOB.toString().trim();
        let yearMatch = dobStr.match(/\b(19|20)\d{2}\b/);
        if (yearMatch) {
          yearPart = yearMatch[0];
        } else {
          yearPart = dobStr.length >= 4 ? dobStr.slice(-4) : "0000";
        }
      }

      const expectedPassword = namePart + yearPart;

      // Validate Password
      if (passwordInput.trim().toLowerCase() === expectedPassword.toLowerCase()) {
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
        return {
          success: true,
          role: 'patient',
          portal: 'patient',
          username: patientKey,
          sessionToken: patientToken,
          message: "Welcome Patient"
        };
      } else {
        return { success: false, message: "Incorrect password. Format: Name(3 chars) + Birth Year." };
      }
    }
  }

  // ==========================================
  // 3. USER NOT FOUND
  // ==========================================
  return { success: false, message: "User ID not found." };
}

// ==========================================
// GOOGLE SSO AUTHENTICATION ENGINE
// ==========================================
function verifyGoogleLogin(userEmail) {
  try {
    if (!userEmail) {
      return { success: false, message: "Authentication payload missing email link." };
    }

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

      if (storedEmail.toString().trim().toLowerCase() === userEmail.toString().trim().toLowerCase()) {
        
        if (isActive && isActive.toString().toLowerCase() !== 'active') {
          return { success: false, message: "Access Denied: This staff credential context is deactivated." };
        }

        const role = storedRole.toString().trim().toLowerCase();
        const doc  = resolveDoctorByUsername_(storedUsername);
        const token = issueSession_({
          username: storedUsername,
          role: role,
          doctorId: doc ? doc.doctorId : "",
          name: doc ? doc.name : storedUsername
        });
        return {
          success: true,
          role: role,
          portal: 'hospital',
          username: storedUsername,
          displayName: doc ? doc.name : storedUsername,
          doctorId: doc ? doc.doctorId : "",
          sessionToken: token,
          message: "Welcome back " + storedUsername
        };
      }
    }

    return { 
      success: false, 
      message: "Access Denied: The email " + userEmail + " is not registered in CresRx. Contact Admin." 
    };

  } catch (error) {
    return { success: false, message: "System Security Fault: " + error.toString() };
  }
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
 * code 'NOT_ENROLLED'  MFA is not set up for this user — the caller lets them in
 * code 'BAD_SECRET'    the enrolment itself is broken — an admin must fix it
 * code 'BAD_CODE'      the secret is fine, the six digits are not
 */
function verifyMFA(username, userCode) {
  try {
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
        return { success: true, code: 'NOT_ENROLLED', message: 'MFA is not enrolled for this user.' };
      }
      if (!norm.ok) {
        return { success: false, code: 'BAD_SECRET', message: norm.message };
      }

      return processTOTP(norm.secret, userCode);
    }
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
  var alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567';
  var secret = '';
  for (var k = 0; k < 32; k++) {
    secret += alphabet.charAt(Math.floor(Math.random() * alphabet.length));
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
