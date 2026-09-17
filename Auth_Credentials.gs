// ============================================================================
// Auth_Credentials.gs — Crescentia HealthTech / CresRx
// Passwords that the spreadsheet cannot read back.
// ----------------------------------------------------------------------------
// WHAT WAS WRONG (DPDP_READINESS.md finding C3, CRITICAL)
//
// Staff passwords sat in column B of the Users sheet in clear text and were
// compared with `===`. Everyone with edit access to the spreadsheet — every
// staff member, every Google account it had ever been shared with, anyone
// holding a copy or an export — could read all of them. People reuse
// passwords, so the blast radius was never this system.
//
// Patient portal passwords were worse than plain text: they were DERIVED.
// registerPatient() generated them as the first three letters of the name
// plus the birth year — Mei2001 — and printed both inputs on the registration
// slip. Anyone holding a patient's prescription could compute their password,
// and the patient ID needed to use it is printed as a barcode on the same
// document. That is not a credential; it is a formula with a public input.
//
// WHAT THIS FILE DOES
//
//   * Stores a salted, iterated SHA-256 digest, never the password.
//   * Compares in constant time, so a wrong answer takes the same time as a
//     nearly-right one.
//   * REFUSES a legacy plain-text credential at sign-in rather than quietly
//     accepting it. Hashing a plaintext column in place would keep every
//     password that has already leaked valid; the only honest migration is a
//     forced reset, so that is what crescMigrateCredentials() does.
//   * Generates random passwords — for a new patient, for a reset — from an
//     alphabet with the look-alikes removed, and marks them for change at
//     first sign-in.
//
// WHY NOT bcrypt / scrypt / Argon2
//   Apps Script has no native slow KDF and cannot load one: Utilities exposes
//   computeDigest and computeHmacSignature and nothing else. Iterated HMAC-
//   SHA-256 (PBKDF2's construction, with the iteration count below) is what
//   the platform permits. It is far short of a memory-hard KDF and is said so
//   here rather than implied to be more: the real defences are the lockout in
//   Auth_Audit.gs, MFA, and the fact that the digest is useless to anyone who
//   opens the sheet.
//
// SETUP, IN ORDER
//   1. crescCredentialStatus()    — see what is stored today, in the log
//   2. crescMigrateCredentials()  — issue a fresh random password to every
//                                   account, hash it, and print the list ONCE
//   3. Hand each person their temporary password. They change it at first
//      sign-in; the system will not let them past until they do.
// ============================================================================

/** Marker, algorithm, iteration count, salt, digest — all in one cell. */
var CRESC_PWD_PREFIX = 'pbkdf2$sha256$';

/**
 * Iterations of HMAC-SHA-256. Every one is a call across the Apps Script
 * bridge, so this is a straight trade of sign-in latency for the cost of
 * guessing offline, and how that trade lands depends on the runtime you are
 * actually on. DO NOT GUESS: run crescPwdBenchmark() in the editor, which
 * times it and tells you what to set. 10,000 is the default; the
 * CRESC_PWD_ITERATIONS script property overrides it, and this function
 * refuses anything below 1,000 however the property is set.
 */
var CRESC_PWD_ITERATIONS_DEFAULT = 10000;

/** Password rules. Length does more than any character class ever did. */
var CRESC_PWD_MIN_LENGTH = 10;

/**
 * The alphabet random passwords are drawn from: no O/0, I/l/1, S/5, B/8.
 * A password read aloud across a reception desk, or copied off a printed
 * slip, is mistyped on exactly those characters — and a user who cannot type
 * their temporary password writes it down somewhere worse.
 */
var CRESC_PWD_ALPHABET = 'ACDEFGHJKMNPQRTUVWXYZabcdefghijkmnpqrtuvwxyz23479';

function cresc_pwdIterations_() {
  try {
    var v = parseInt(PropertiesService.getScriptProperties()
      .getProperty('CRESC_PWD_ITERATIONS'), 10);
    if (isFinite(v) && v >= 1000) return v;
  } catch (e) {}
  return CRESC_PWD_ITERATIONS_DEFAULT;
}

// ---------------------------------------------------------------------------
// SECTION A — THE PRIMITIVES
// ---------------------------------------------------------------------------

/** Bytes -> base64, for storing a digest in a spreadsheet cell. */
function cresc_b64_(bytes) { return Utilities.base64Encode(bytes); }

/** A fresh 128-bit salt. Per user, never reused, stored beside the digest. */
function cresc_newSalt_() {
  // getUuid() is a v4 UUID: 122 bits of randomness from the platform's own
  // generator. Math.random() is not a credible source for a salt and is not
  // used for one.
  return Utilities.getUuid().replace(/-/g, '');
}

/**
 * The digest itself: HMAC-SHA-256 chained `iterations` times, salt as the key.
 *
 * @param {string} plain
 * @param {string} salt
 * @param {number} iterations
 * @return {string} base64 of the final 32 bytes
 */
function cresc_pwdDigest_(plain, salt, iterations) {
  // Both arguments as byte arrays on every call: Utilities overloads this on
  // (String, String) and (Byte[], Byte[]), and mixing the two is how a hash
  // that worked in the editor stops working under a different runtime.
  var key = Utilities.newBlob(String(salt)).getBytes();
  var bytes = Utilities.computeHmacSha256Signature(
    Utilities.newBlob(String(plain)).getBytes(), key);
  for (var i = 1; i < iterations; i++) {
    bytes = Utilities.computeHmacSha256Signature(bytes, key);
  }
  return cresc_b64_(bytes);
}

/**
 * Encodes a password for storage. The whole record is one string so that a
 * future change of algorithm or iteration count can live beside the old one
 * instead of needing a second column and a migration.
 *
 * @return {string} pbkdf2$sha256$<iterations>$<salt>$<digest>
 */
function crescPwdEncode_(plain) {
  var iters = cresc_pwdIterations_();
  var salt = cresc_newSalt_();
  return CRESC_PWD_PREFIX + iters + '$' + salt + '$' +
         cresc_pwdDigest_(plain, salt, iters);
}

/** Is this cell a stored digest, or a password somebody can just read? */
function crescPwdIsHashed_(stored) {
  return String(stored === null || stored === undefined ? '' : stored)
           .trim().indexOf(CRESC_PWD_PREFIX) === 0;
}

/**
 * Constant-time comparison.
 *
 * `a === b` on two strings returns as soon as it finds a difference, so the
 * time it takes reveals how many leading characters were right. That is a
 * real attack on a remote digest comparison and it costs nothing to close:
 * always look at every character.
 */
function crescPwdEquals_(a, b) {
  var x = String(a || ''), y = String(b || '');
  var diff = x.length ^ y.length;
  var n = Math.max(x.length, y.length);
  for (var i = 0; i < n; i++) {
    diff |= (x.charCodeAt(i) || 0) ^ (y.charCodeAt(i) || 0);
  }
  return diff === 0;
}

/**
 * Checks a password against what is stored.
 *
 * @return {{ok:boolean, legacy:boolean}}
 *   ok      the password matches a STORED DIGEST
 *   legacy  the cell holds plain text — the answer is never `ok`, whatever
 *           was typed, because accepting it would keep a leaked password
 *           working. The caller tells the user their password must be reset.
 */
function crescPwdVerify_(plain, stored) {
  var cell = String(stored === null || stored === undefined ? '' : stored).trim();
  if (!cell) return { ok: false, legacy: false };
  if (!crescPwdIsHashed_(cell)) return { ok: false, legacy: true };

  var parts = cell.split('$');           // pbkdf2, sha256, iters, salt, digest
  if (parts.length !== 5) return { ok: false, legacy: false };
  var iters = parseInt(parts[2], 10);
  if (!isFinite(iters) || iters < 1) return { ok: false, legacy: false };

  return { ok: crescPwdEquals_(cresc_pwdDigest_(plain, parts[3], iters), parts[4]),
           legacy: false };
}

/**
 * A random password, drawn from the platform's random source rather than
 * Math.random(), which is neither seeded nor documented for this purpose.
 */
function crescRandomPassword_(length) {
  var n = length || 12;
  var out = '';
  while (out.length < n) {
    // Each UUID contributes 32 hex characters of platform randomness; they are
    // mapped into the alphabet rather than used directly so the result is not
    // restricted to 0-9a-f.
    var raw = Utilities.getUuid().replace(/-/g, '');
    for (var i = 0; i < raw.length && out.length < n; i += 2) {
      var v = parseInt(raw.substr(i, 2), 16);
      out += CRESC_PWD_ALPHABET.charAt(v % CRESC_PWD_ALPHABET.length);
    }
  }
  return out;
}

/**
 * Is this new password allowed? Length first, because it is the only rule
 * that reliably helps; then the two patterns this clinic has actually been
 * bitten by — the derived portal password, and the username itself.
 *
 * @return {{ok:boolean, message:string}}
 */
function crescPwdPolicy_(plain, username, patientName) {
  var p = String(plain || '');
  if (p.length < CRESC_PWD_MIN_LENGTH) {
    return { ok: false, message: 'Use at least ' + CRESC_PWD_MIN_LENGTH +
             ' characters. A longer ordinary phrase is stronger than a short ' +
             'one with symbols in it.' };
  }
  if (p.length > 100) {
    return { ok: false, message: 'That is longer than 100 characters.' };
  }
  var lower = p.toLowerCase();
  if (username && lower.indexOf(String(username).toLowerCase()) !== -1) {
    return { ok: false, message: 'Your password cannot contain your user ID.' };
  }
  // The old portal rule: three letters of the name plus a year. Anyone who has
  // seen the patient's prescription can type that, which is why it is gone.
  var namePart = String(patientName || '').replace(/[^a-zA-Z]/g, '').substring(0, 3);
  if (namePart.length === 3 &&
      new RegExp('^' + namePart + '\\s*(19|20)\\d{2}$', 'i').test(p)) {
    return { ok: false, message: 'That is the old automatic password pattern — ' +
             'the first letters of the name and a year. It is guessable by ' +
             'anyone holding one of your documents. Choose something else.' };
  }
  if (/^(.)\1+$/.test(p)) {
    return { ok: false, message: 'One repeated character is not a password.' };
  }
  return { ok: true, message: '' };
}

/**
 * ADMIN, from the script editor. How long a sign-in actually costs here.
 *
 * The iteration count is the only security parameter in this file that has a
 * user-visible price, and the price is different on every Apps Script runtime
 * and every day. So it is measured rather than asserted: run this, read the
 * milliseconds, and set CRESC_PWD_ITERATIONS to whatever gives you a hashing
 * cost you are willing to pay on every sign-in.
 *
 * Roughly: under 400ms is unnoticeable next to the sheet reads a sign-in
 * already does. Over 1500ms and people will say the system is slow, and a
 * system people say is slow is one they leave signed in on a shared machine —
 * which is a worse security outcome than a lower iteration count.
 */
function crescPwdBenchmark() {
  var salt = cresc_newSalt_();
  var sizes = [1000, 5000, 10000, 20000];
  var lines = ['Password hashing cost on this runtime', ''];

  sizes.forEach(function (n) {
    var t = Date.now();
    cresc_pwdDigest_('a representative password', salt, n);
    var ms = Date.now() - t;
    lines.push('  ' + String(n).padStart(6) + ' iterations   ' +
               String(ms).padStart(5) + ' ms' +
               (ms < 400 ? '   comfortable'
                : ms < 1500 ? '   noticeable but fine'
                : '   too slow — people will complain'));
  });

  var current = cresc_pwdIterations_();
  lines.push('');
  lines.push('In use now: ' + current + ' iterations.');
  lines.push('');
  lines.push('To change it:');
  lines.push('  PropertiesService.getScriptProperties()');
  lines.push('    .setProperty("CRESC_PWD_ITERATIONS", "8000");');
  lines.push('');
  lines.push('Passwords already stored keep their own iteration count — it is');
  lines.push('written into each stored value — so changing this does not lock');
  lines.push('anybody out. New and changed passwords use the new number.');

  var report = lines.join('\n');
  Logger.log(report);
  return report;
}

// ---------------------------------------------------------------------------
// SECTION B — WHERE CREDENTIALS LIVE
// ---------------------------------------------------------------------------

/**
 * The Users sheet gained two columns: whether a password must be changed at
 * next sign-in, and when it was last set. Both are created on demand so an
 * existing deployment needs no manual schema edit.
 *
 * A -> Username | B Password | C Role | D Status | E Email | F MFA_Secret
 * G Must_Change | H Password_Updated_At
 */
function cresc_usersSheet_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sh) throw new Error('There is no Users sheet.');
  if (sh.getLastColumn() < 6) sh.getRange(1, 6).setValue('MFA_Secret');
  ensureColumn_(sh, 'Must_Change');
  ensureColumn_(sh, 'Password_Updated_At');
  return sh;
}

/** The same two columns on the patient master, for the portal credential. */
function cresc_patientsSheet_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Patients');
  if (!sh) throw new Error('There is no Patients sheet.');
  ensureColumn_(sh, 'Portal_Must_Change');
  ensureColumn_(sh, 'Portal_Password_Updated_At');
  return sh;
}

/** 1-based column index of a header, or -1. Never creates. */
function cresc_colOf_(sheet, header) {
  try {
    var row = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    for (var i = 0; i < row.length; i++) {
      if (String(row[i]).trim() === header) return i + 1;
    }
  } catch (e) {}
  return -1;
}

/** Is this cell's "must change" flag set? Blank means no. */
function cresc_mustChange_(sheet, rowNo, header) {
  var col = cresc_colOf_(sheet, header);
  if (col === -1) return false;
  var v = String(sheet.getRange(rowNo, col).getDisplayValue() || '').trim().toUpperCase();
  return v === 'YES' || v === 'TRUE' || v === 'Y' || v === '1';
}

/** Writes a new credential and stamps when. `mustChange` for a temporary one. */
function cresc_writeCredential_(sheet, rowNo, plain, mustChange, flagHeader, stampHeader) {
  // Force text, or Sheets reformats a digest that happens to look numeric and
  // the stored value stops matching what was written.
  sheet.getRange(rowNo, 2).setNumberFormat('@').setValue(crescPwdEncode_(plain));
  var flagCol = cresc_colOf_(sheet, flagHeader);
  if (flagCol !== -1) sheet.getRange(rowNo, flagCol).setValue(mustChange ? 'YES' : 'NO');
  var stampCol = cresc_colOf_(sheet, stampHeader);
  if (stampCol !== -1) sheet.getRange(rowNo, stampCol).setValue(new Date());
}

// ---------------------------------------------------------------------------
// SECTION C — CHANGING A PASSWORD
// ---------------------------------------------------------------------------

/**
 * FRONTEND ENTRY, and deliberately reachable without a session: the commonest
 * time a password has to change is the moment the system has just refused to
 * let someone in with a temporary one. It proves the CURRENT password before
 * it writes anything, so it grants nothing a sign-in would not.
 *
 * Counted by the same lockout as a sign-in, so it cannot be used as a
 * password oracle that sidesteps the five-attempt hold.
 *
 * @param {{username, currentPassword, newPassword}} payload
 * @return {{success:boolean, message:string, code?:string}}
 */
function crescChangePassword(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    payload = payload || {};
    var username = String(payload.username || '').trim();
    var current  = String(payload.currentPassword || '');
    var next     = String(payload.newPassword || '');

    if (!username || !current || !next) {
      return { success: false, message: 'Give your user ID, your current password and the new one.' };
    }

    var held = crescAuthGuard_(username);
    if (held) return held;

    if (current === next) {
      return { success: false, message: 'The new password is the same as the old one.' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ---- staff --------------------------------------------------------
    var users = ss.getSheetByName('Users');
    if (users) {
      var udata = users.getDataRange().getValues();
      for (var i = 1; i < udata.length; i++) {
        if (!udata[i][0]) continue;
        if (String(udata[i][0]).trim().toUpperCase() !== username.toUpperCase()) continue;

        var uv = crescPwdVerify_(current, udata[i][1]);
        if (uv.legacy) {
          // Deliberately NOT a self-service migration path. A plain-text
          // password has been readable by everyone with access to the
          // spreadsheet; letting it authorise its own replacement would let
          // whoever copied it take the account over instead of the owner.
          // An administrator resets it — crescMigrateCredentials() does every
          // account at once, crescAdminResetPassword() does one.
          return { success: false, code: 'LEGACY_CREDENTIAL',
                   message: 'This account is still on an old, unprotected password ' +
                            'and has to be reset by an administrator before it can ' +
                            'be used or changed.' };
        }
        if (!uv.ok) {
          var uwarn = crescAuthFailed_(username, String(udata[i][2] || ''), 'bad password at change');
          return { success: false, message: uwarn.indexOf('locked') !== -1 ||
                   uwarn.indexOf('left') !== -1 ? uwarn : 'Incorrect ID or password.' };
        }

        var upol = crescPwdPolicy_(next, username, '');
        if (!upol.ok) return { success: false, message: upol.message };

        var ush = cresc_usersSheet_();
        cresc_writeCredential_(ush, i + 1, next, false, 'Must_Change', 'Password_Updated_At');
        SpreadsheetApp.flush();
        crescAuthPassed_(username, String(udata[i][2] || ''), 'password-change');
        crescAuthAudit_(CRESC_AUTH_EVENTS.PASSWORD_CHANGED, username, String(udata[i][2] || ''),
                        { by: 'self' });
        return { success: true, message: 'Password changed. Sign in with the new one.' };
      }
    }

    // ---- patient portal ------------------------------------------------
    var patients = ss.getSheetByName('Patients');
    if (patients) {
      var pdata = patients.getDataRange().getValues();
      for (var j = 1; j < pdata.length; j++) {
        if (!pdata[j][0]) continue;
        if (String(pdata[j][0]).trim().toUpperCase() !== username.toUpperCase()) continue;

        var pv = crescPwdVerify_(current, pdata[j][1]);
        if (pv.legacy) {
          // The old portal password was derivable from the patient's name and
          // birth year, both printed on documents they carry. Accepting it
          // here would let anyone holding a prescription set a new password
          // and take the portal account. The front desk resets it instead.
          return { success: false, code: 'LEGACY_CREDENTIAL',
                   message: 'Your portal password has to be reset at the clinic ' +
                            'before it can be used. Ask at reception.' };
        }
        if (!pv.ok) {
          var pwarn = crescAuthFailed_(username, 'patient', 'bad password at change');
          return { success: false, message: pwarn.indexOf('locked') !== -1 ||
                   pwarn.indexOf('left') !== -1 ? pwarn : 'Incorrect ID or password.' };
        }

        var ppol = crescPwdPolicy_(next, username, pdata[j][2]);
        if (!ppol.ok) return { success: false, message: ppol.message };

        var psh = cresc_patientsSheet_();
        cresc_writeCredential_(psh, j + 1, next, false,
                               'Portal_Must_Change', 'Portal_Password_Updated_At');
        SpreadsheetApp.flush();
        crescAuthPassed_(username, 'patient', 'password-change');
        crescAuthAudit_(CRESC_AUTH_EVENTS.PASSWORD_CHANGED, username, 'patient', { by: 'self' });
        return { success: true, message: 'Password changed. Sign in with the new one.' };
      }
    }

    // Same wording as a wrong password: whether an ID exists is not something
    // this endpoint is willing to confirm.
    crescAuthFailed_(username, '', 'no such user at change');
    return { success: false, message: 'Incorrect ID or password.' };

  } catch (err) {
    return { success: false, message: 'Could not change the password: ' + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * FRONTEND ENTRY. An administrator resets somebody's password and is shown
 * the temporary one ONCE, to hand over.
 *
 * The new password is random and marked for change, so the administrator
 * cannot keep using it and the user cannot leave it in place.
 */
function crescAdminResetPassword(username, sessionToken) {
  try {
    var actor = crescRequire_(sessionToken, 'admin.users');
    var want = String(username || '').trim();
    if (!want) return { success: false, message: 'Name the account to reset.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var temp = crescRandomPassword_(12);

    var users = ss.getSheetByName('Users');
    if (users) {
      var udata = users.getDataRange().getValues();
      for (var i = 1; i < udata.length; i++) {
        if (!udata[i][0]) continue;
        if (String(udata[i][0]).trim().toUpperCase() !== want.toUpperCase()) continue;
        cresc_writeCredential_(cresc_usersSheet_(), i + 1, temp, true,
                               'Must_Change', 'Password_Updated_At');
        SpreadsheetApp.flush();
        crescAuthAudit_(CRESC_AUTH_EVENTS.PASSWORD_RESET, want, String(udata[i][2] || ''),
                        { by: actor.username });
        return { success: true, temporaryPassword: temp,
                 message: 'Temporary password for ' + want + ': ' + temp +
                          '\nIt is shown once. They must change it at first sign-in.' };
      }
    }

    var patients = ss.getSheetByName('Patients');
    if (patients) {
      var pdata = patients.getDataRange().getValues();
      for (var j = 1; j < pdata.length; j++) {
        if (!pdata[j][0]) continue;
        if (String(pdata[j][0]).trim().toUpperCase() !== want.toUpperCase()) continue;
        cresc_writeCredential_(cresc_patientsSheet_(), j + 1, temp, true,
                               'Portal_Must_Change', 'Portal_Password_Updated_At');
        SpreadsheetApp.flush();
        crescAuthAudit_(CRESC_AUTH_EVENTS.PASSWORD_RESET, want, 'patient', { by: actor.username });
        return { success: true, temporaryPassword: temp,
                 message: 'Temporary portal password for ' + want + ': ' + temp +
                          '\nIt is shown once. They must change it at first sign-in.' };
      }
    }

    return { success: false, message: 'No account with the ID ' + want + '.' };
  } catch (err) {
    return { success: false, message: String(err.message || err).replace('FORBIDDEN: ', '') };
  }
}

// ---------------------------------------------------------------------------
// SECTION C2 — CREATING A STAFF ACCOUNT
// ---------------------------------------------------------------------------

/** The roles a STAFF account may hold. Everything the matrix knows, less
 *  'patient' — see crescCreateStaffAccount for why. */
function crescStaffRoles_() {
  return Object.keys(CRESC_ROLE_MATRIX).filter(function (r) { return r !== 'patient'; });
}

/**
 * WHY THIS EXISTS.
 *
 * Adding a doctor, a nurse or an administrator meant typing a row into the
 * Users sheet by hand — and the password column holds a PBKDF2 digest, which
 * cannot be typed. So a hand-added account either got a plain-text password
 * (which AuthLogin refuses outright, as LOGIN_LEGACY_CREDENTIAL) or it got
 * whatever the person pasted, and the only way to make it work was to run
 * crescMigrateCredentials and reset everybody. There was no supported way to
 * add a user at all.
 *
 * A doctor also needs a row on the Doctors sheet, linked by Linked_Username,
 * or nothing attributes their consultations to them: the schedule, the
 * signature line, the referral list and the case-sheet author all read
 * Doctors, not Users. Creating one without the other is the half-made account
 * that looks fine until the first prescription prints unsigned.
 *
 * So this does both, once, in the right order, with a hashed credential.
 *
 * @param {{username:string, role:string, displayName:string, email:string,
 *          specialty:string, regNo:string, qualification:string,
 *          consultFee:number, isVisiting:boolean, canViewAll:boolean,
 *          password:string}} payload
 *          password is optional; a random one is generated and returned when
 *          it is left out. Either way the account must change it at first
 *          sign-in.
 * @param {string} sessionToken
 * @return {{success:boolean, username:string, doctorId:string,
 *           temporaryPassword:string, message:string}}
 */
function crescCreateStaffAccount(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    payload = payload || {};
    var actor = crescRequire_(sessionToken, 'admin.users');

    var username = String(payload.username || '').trim();
    var role     = String(payload.role || '').trim().toLowerCase();
    var display  = String(payload.displayName || '').trim();

    if (!username) return { success: false, message: 'Give a username.' };
    if (!/^[A-Za-z0-9._-]{3,40}$/.test(username)) {
      return { success: false,
               message: 'A username may use letters, digits, dot, underscore ' +
                        'and hyphen, and must be 3 to 40 characters. Spaces ' +
                        'break the sign-in form.' };
    }
    if (!display) return { success: false, message: 'Give the person\'s name.' };

    // The role has to be one the permission matrix knows, or the account
    // signs in and can do nothing — crescRequire_ refuses every endpoint with
    // "the role X on your account is not one this system knows", which reads
    // as a broken system rather than a typo at creation time.
    var known = (typeof crescRole_ === 'function') ? crescRole_(role) : role;
    if (!known || !CRESC_ROLE_MATRIX[known]) {
      return { success: false,
               message: '"' + (payload.role || '') + '" is not a role this system ' +
                        'knows. Use one of: ' + crescStaffRoles_().join(', ') + '.' };
    }
    // 'patient' is in the matrix but is not a staff account: a patient signs
    // in with their patient ID against the Patients sheet, and a Users row
    // claiming that role would be an account with portal permissions and no
    // record behind it.
    if (known === 'patient') {
      return { success: false,
               message: 'Patients are not created here. Register the patient, ' +
                        'and the portal login is their patient ID.' };
    }

    var users = cresc_usersSheet_();
    var udata = users.getDataRange().getValues();
    for (var i = 1; i < udata.length; i++) {
      if (String(udata[i][0] || '').trim().toUpperCase() === username.toUpperCase()) {
        return { success: false, code: 'EXISTS',
                 message: username + ' already exists. Reset its password ' +
                          'instead of creating it again.' };
      }
    }

    // A supplied password is held to the same policy a user's own change is;
    // a generated one is random and long enough not to need checking.
    var temp = String(payload.password || '').trim();
    if (temp) {
      var pol = crescPwdPolicy_(temp, username, display);
      if (!pol.ok) return { success: false, message: pol.message };
    } else {
      temp = crescRandomPassword_(12);
    }

    // ---- the Users row ---------------------------------------------------
    var m = dc_headerMap_(users);
    var row = new Array(users.getLastColumn()).fill('');
    var put = function (h, v) { if (m[h] !== undefined) row[m[h]] = v; };
    row[0] = username;
    row[1] = '';                       // written by cresc_writeCredential_ below
    row[2] = known;
    row[3] = 'ACTIVE';
    put('Email Address', String(payload.email || '').trim());
    put('Must_Change', 'YES');
    users.appendRow(row);
    var rowNo = users.getLastRow();

    // Written through the same helper the reset path uses, so a created
    // account and a reset account are stored identically — including the
    // plain-text number format, without which Sheets reformats a digest that
    // happens to look numeric and the stored value stops matching.
    cresc_writeCredential_(users, rowNo, temp, true, 'Must_Change', 'Password_Updated_At');
    SpreadsheetApp.flush();

    // ---- the Doctors row, for a role that consults -----------------------
    var doctorId = '';
    var clinical = (known === 'doctor');
    if (clinical) {
      try {
        doctorId = cresc_createDoctorRow_(username, display, payload, known);
      } catch (e) {
        // The login exists and works; only the clinical profile failed. Say
        // so precisely rather than reporting the whole thing as a failure and
        // inviting a second attempt that hits EXISTS.
        return { success: true, username: username, doctorId: '',
                 temporaryPassword: temp,
                 message: 'The login for ' + username + ' was created, but the ' +
                          'entry on the Doctors sheet was not (' + e.message +
                          '). Add it before they consult, or their ' +
                          'prescriptions print unattributed. Temporary ' +
                          'password: ' + temp };
      }
    }

    try {
      logAudit_({ username: actor.username, role: actor.role, doctorId: actor.doctorId },
                'STAFF_ACCOUNT_CREATED', 'User', username.toUpperCase(),
                { role: known, doctorId: doctorId, visiting: !!payload.isVisiting });
    } catch (e) {}

    return {
      success: true,
      username: username,
      doctorId: doctorId,
      temporaryPassword: temp,
      message: display + ' can sign in as "' + username + '" with the ' +
               'temporary password ' + temp + '. It is shown once, and they ' +
               'must change it at first sign-in.' +
               (doctorId ? ' Doctor profile ' + doctorId + ' created.' : '')
    };
  } catch (err) {
    return { success: false,
             message: String((err && err.message) || err).replace('FORBIDDEN: ', '') };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * The Doctors row that makes a login into a clinician.
 *
 * Linked_Username is the join every other module uses to get from a session
 * to a doctor, so it is the one field that must not be left blank.
 *
 * @return {string} the new Doctor_ID
 */
function cresc_createDoctorRow_(username, display, payload, role) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Doctors');
  if (!sh) {
    if (typeof setupDoctorsSheet === 'function') {
      setupDoctorsSheet();
      sh = ss.getSheetByName('Doctors');
    }
    if (!sh) throw new Error('there is no Doctors sheet');
  }
  if (typeof dc_ensureColumn_ === 'function') dc_ensureColumn_(sh, 'Is_Visiting');

  var m = dc_headerMap_(sh);
  var data = sh.getDataRange().getValues();

  // Next free DOCnnn, so this never collides with the ids already in use.
  var maxN = 0;
  for (var i = 1; i < data.length; i++) {
    var mm = /^DOC(\d+)$/i.exec(String(data[i][0] || '').trim());
    if (mm) maxN = Math.max(maxN, parseInt(mm[1], 10));
  }
  var newId = 'DOC' + ('000' + (maxN + 1)).slice(-3);

  var regNo = String(payload.regNo || '').trim();
  var qual  = String(payload.qualification || '').trim();
  var spec  = String(payload.specialty || '').trim();

  // The signature line is what prints under a prescription. Built from what
  // was given rather than left empty, because an empty one prints a blank
  // where the registration number belongs.
  var sigParts = [display];
  if (qual) sigParts.push(qual);
  if (spec) sigParts.push(spec);
  if (regNo) sigParts.push('Reg. No. ' + regNo);
  var signature = sigParts.join(', ');

  var row = new Array(sh.getLastColumn()).fill('');
  var put = function (h, v) { if (m[h] !== undefined) row[m[h]] = v; };
  row[0] = newId;
  put('Doctor_ID', newId);
  put('Tenant_ID', (typeof getTenantId_ === 'function') ? getTenantId_() : '');
  put('Display_Name', display);
  put('Specialty', spec);
  put('Reg_No', regNo);
  put('Signature_Line', signature);
  put('Linked_Username', username);
  put('Status', 'ACTIVE');
  put('Can_View_All', payload.canViewAll ? 'TRUE' : 'FALSE');
  put('Default_Consult_Fee', Number(payload.consultFee) || 0);
  put('Is_Visiting', payload.isVisiting ? 'TRUE' : 'FALSE');

  sh.appendRow(row);
  if (typeof dc_invalidate_ === 'function') dc_invalidate_('Doctors');
  SpreadsheetApp.flush();
  return newId;
}

// ---------------------------------------------------------------------------
// SECTION D — THE MIGRATION
// ---------------------------------------------------------------------------

/**
 * ADMIN, run from the script editor. What is stored today, without printing
 * a single password.
 */
function crescCredentialStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lines = ['Credential storage — ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), ''];

  [['Users', 'staff logins'], ['Patients', 'patient portal logins']].forEach(function (pair) {
    var sh = ss.getSheetByName(pair[0]);
    if (!sh || sh.getLastRow() < 2) { lines.push('  ' + pair[0] + ': absent or empty.'); return; }
    var col = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getDisplayValues();
    var hashed = 0, plain = 0, blank = 0;
    col.forEach(function (r) {
      var v = String(r[0] || '').trim();
      if (!v) blank++;
      else if (crescPwdIsHashed_(v)) hashed++;
      else plain++;
    });
    lines.push('  ' + pair[0] + ' (' + pair[1] + '): ' + hashed + ' hashed, ' +
               plain + ' in PLAIN TEXT, ' + blank + ' blank.');
  });

  lines.push('');
  lines.push('A plain-text credential is REFUSED at sign-in — it is not silently');
  lines.push('accepted and it is not silently hashed, because hashing a password');
  lines.push('that has already been readable keeps the compromise. Run');
  lines.push('crescMigrateCredentials() to issue fresh random passwords, then hand');
  lines.push('them out. Iterations in use: ' + cresc_pwdIterations_() + '.');

  var report = lines.join('\n');
  Logger.log(report);
  return report;
}

/**
 * ADMIN, run from the script editor ONCE at deployment.
 *
 * Issues a fresh random password to every account whose cell is not already a
 * digest, stores the digest, marks it for change at first sign-in, and prints
 * the list — the only time any of them is ever visible.
 *
 * WHY IT DOES NOT JUST HASH WHAT IS THERE. Every one of those passwords has
 * been readable by everyone with access to the spreadsheet for as long as the
 * sheet has existed. Hashing them in place would protect them from here on
 * and leave every already-copied password working — which is the compromise,
 * not a fix for it.
 *
 * @param {boolean} [dryRun] true to report who WOULD be reset, changing nothing
 */
function crescMigrateCredentials(dryRun) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = ['Credential migration' + (dryRun ? ' — DRY RUN' : '') + ' — ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), ''];
  var changed = 0, already = 0;

  function sweep(sheetName, ensure, flag, stamp, label) {
    var sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) { out.push(sheetName + ': absent or empty.'); return; }
    if (!dryRun) sh = ensure();
    var data = sh.getDataRange().getValues();
    out.push(label + ':');
    for (var i = 1; i < data.length; i++) {
      var id = String(data[i][0] || '').trim();
      if (!id) continue;
      if (crescPwdIsHashed_(data[i][1])) { already++; continue; }
      var temp = crescRandomPassword_(12);
      if (!dryRun) cresc_writeCredential_(sh, i + 1, temp, true, flag, stamp);
      out.push('  ' + id + '   ' + (dryRun ? '(would be reset)' : temp));
      changed++;
    }
  }

  sweep('Users', cresc_usersSheet_, 'Must_Change', 'Password_Updated_At', 'STAFF');
  sweep('Patients', cresc_patientsSheet_, 'Portal_Must_Change',
        'Portal_Password_Updated_At', 'PATIENT PORTAL');

  if (!dryRun) SpreadsheetApp.flush();

  out.push('');
  out.push(changed + ' credential(s) ' + (dryRun ? 'would be reset' : 'reset') +
           ', ' + already + ' already hashed.');
  out.push('');
  out.push('THIS LOG IS THE ONLY COPY. Hand each person their password by a route');
  out.push('that is not the same one they will use to sign in, and do not paste');
  out.push('this list into the spreadsheet, a chat, or an email to everyone.');
  out.push('Each account must change it at first sign-in before it can do anything.');

  var report = out.join('\n');
  Logger.log(report);
  // Not audited with the passwords in it, for the obvious reason. The fact
  // that a migration ran is worth recording; its output is not.
  try {
    if (!dryRun) crescAuthAudit_(CRESC_AUTH_EVENTS.PASSWORD_MIGRATION, 'SYSTEM', '', { reset: changed });
  } catch (e) {}
  return report;
}
