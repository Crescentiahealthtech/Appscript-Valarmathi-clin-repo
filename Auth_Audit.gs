// ============================================================================
// Auth_Audit.gs — Crescentia HealthTech / CresRx
// Sign-in audit, and a lockout that makes the audit worth reading.
// ----------------------------------------------------------------------------
// WHAT WAS MISSING
//
// AuthLogin.gs recorded nothing. Not a success, not a failure, not a sign-out.
// logAudit_() already existed and eleven other modules were writing to
// Audit_Log through it; authentication — the one event every audit starts
// from — wrote nothing at all.
//
// So the two questions an incident actually asks could not be answered:
//
//   "Did anyone else sign in as Dr. Rekha last Tuesday?"     — unanswerable
//   "Is someone working through our patient IDs right now?"  — invisible
//
// The second one matters more than it looks. Patient passwords are derived
// from the first three letters of the name plus the birth year, the patient
// ID is printed on every bill and every barcode label, and there was no limit
// on failed attempts anywhere in the sign-in path. An unlimited guess rate
// against a four-digit-plus-three-letter secret is not a lock.
//
// THIS FILE
//
//   crescAuthAudit_()   one row per attempt, into the same Audit_Log
//   crescAuthGuard_()   refuse an account that has just failed five times
//   crescAuthFailed_()  count a failure
//   crescAuthPassed_()  clear the count
//   crescGetLoginAudit()  read it back, for an administrator
//
// Under the DPDP Act the clinic is accountable for access to patient data,
// and "we keep no record of who signed in" is not a position it is possible
// to defend. This is the minimum that makes it defensible.
// ============================================================================

/** Failures allowed inside the window before the account is held. */
var CRESC_LOGIN_FAIL_MAX = 5;

/** The window, and the hold, in seconds. Fifteen minutes each. */
var CRESC_LOGIN_FAIL_WINDOW_S = 900;

/** Events this file writes. Kept as constants so a query can rely on them. */
var CRESC_AUTH_EVENTS = {
  SUCCESS:      'LOGIN_SUCCESS',
  FAILED:       'LOGIN_FAILED',
  LOCKED:       'LOGIN_LOCKED',
  BLOCKED:      'LOGIN_BLOCKED',      // attempt refused while locked
  DISABLED:     'LOGIN_DISABLED',     // account marked not active
  UNKNOWN_USER: 'LOGIN_UNKNOWN_USER',
  MFA_FAILED:   'LOGIN_MFA_FAILED',
  MFA_PASSED:   'LOGIN_MFA_PASSED',
  SIGNOUT:      'LOGOUT',

  // Credential lifecycle (Auth_Credentials.gs). A password that changes is a
  // fact an incident review needs: "when did this account's password last
  // change" is the second question asked after "who signed in".
  PASSWORD_CHANGED:   'PASSWORD_CHANGED',
  PASSWORD_RESET:     'PASSWORD_RESET',
  PASSWORD_MIGRATION: 'PASSWORD_MIGRATION',
  LEGACY_REFUSED:     'LOGIN_LEGACY_CREDENTIAL',   // plain-text cell, refused
  MUST_CHANGE:        'LOGIN_MUST_CHANGE'          // correct password, expired by policy
};


// ---------------------------------------------------------------------------
// SECTION A — THE RECORD
// ---------------------------------------------------------------------------

/**
 * One audit row per authentication event.
 *
 * Never throws. An audit that can break the thing it is auditing is worse
 * than no audit: a sheet quota error must not become "nobody can sign in".
 *
 * WHAT IS DELIBERATELY NOT RECORDED: the password, correct or otherwise, in
 * any form. A failed attempt that logs what was typed turns the audit log
 * into a list of near-miss passwords, and the audit log is readable by more
 * people than the Users sheet is.
 *
 * WHAT CANNOT BE RECORDED: the caller's IP address. Apps Script does not
 * expose it to a web app, so "signed in from where" is out of reach here and
 * would need a reverse proxy in front of the deployment. Said plainly rather
 * than left as a gap somebody assumes is covered.
 *
 * @param {string} event     one of CRESC_AUTH_EVENTS
 * @param {string} username  the identifier that was TYPED, whether or not it exists
 * @param {string} role      the role, when known
 * @param {Object} [detail]  method, reason, counters — no secrets
 */
function crescAuthAudit_(event, username, role, detail) {
  try {
    var d = detail || {};
    // Session.getActiveUser() is populated only when the visitor is signed in
    // to a Google account in the same domain as the deployment; anonymous
    // access leaves it empty, which is itself worth recording.
    try { d.googleAccount = Session.getActiveUser().getEmail() || '(anonymous)'; }
    catch (e) { d.googleAccount = '(unavailable)'; }
    d.at = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    logAudit_({ username: crescStr_(username) || '(blank)', role: crescStr_(role), doctorId: '' },
              event, 'Session', crescStr_(username).toUpperCase(), d);
  } catch (e) {
    Logger.log('crescAuthAudit_ failed (' + event + '): ' + e.message);
  }
}

/** FRONTEND ENTRY. Records a sign-out. Best effort — never blocks the client. */
function crescLogSignOut(token) {
  try {
    var actor = crescActor_(token);
    crescAuthAudit_(CRESC_AUTH_EVENTS.SIGNOUT,
                    actor ? actor.username : '(expired session)',
                    actor ? actor.role : '', {});
    // revokeSession() has carried the comment "Call from your sign-out
    // handler to revoke server-side. Optional." since it was written, and
    // nothing ever called it. So a sign-out only cleared localStorage: the
    // token stayed valid in the cache and in the durable Sessions sheet for
    // its full eight hours, and anyone who had it could keep using it after
    // the user believed they had left.
    if (typeof revokeSession === 'function') {
      try { revokeSession(token); } catch (e) {}
    }
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  }
}


// ---------------------------------------------------------------------------
// SECTION B — THE LOCKOUT
// ---------------------------------------------------------------------------

/**
 * CacheService, not a sheet.
 *
 * The counter only has to survive the fifteen minutes an attack takes, a
 * cache read costs nothing on a path every sign-in runs, and a sheet write
 * per failed attempt is exactly what an attacker would want us to do. If the
 * entry is evicted the account unlocks early — which is the right way for
 * this to fail, because the alternative is locking a clinic out of its own
 * system on a cache hiccup.
 */
function crescLoginFailKey_(username) {
  return 'LOGINFAIL_' + crescStr_(username).toUpperCase().replace(/[^A-Z0-9_\-]/g, '').substring(0, 100);
}

/** How many failures are on record for this identifier right now. */
function crescLoginFailCount_(username) {
  try {
    var raw = CacheService.getScriptCache().get(crescLoginFailKey_(username));
    return raw ? (parseInt(raw, 10) || 0) : 0;
  } catch (e) { return 0; }
}

/**
 * Refuses the attempt when the account is being guessed at.
 *
 * Returns a ready-to-return failure object, or null to carry on. Called
 * BEFORE any password comparison, so a locked account costs an attacker a
 * cache read and nothing else.
 */
function crescAuthGuard_(username) {
  var n = crescLoginFailCount_(username);
  if (n < CRESC_LOGIN_FAIL_MAX) return null;

  crescAuthAudit_(CRESC_AUTH_EVENTS.BLOCKED, username, '', { failures: n });
  return {
    success: false,
    locked: true,
    message: 'Too many failed attempts. This account is locked for ' +
             Math.round(CRESC_LOGIN_FAIL_WINDOW_S / 60) + ' minutes. ' +
             'Contact an administrator if this was not you.'
  };
}

/**
 * Counts one failure and returns the message to show.
 *
 * The remaining-attempts warning starts only at the halfway mark. Telling
 * somebody on their first slip that they have four tries left mostly teaches
 * an attacker the threshold; telling a real user on their fourth that the
 * next one locks the account is a genuine warning.
 */
function crescAuthFailed_(username, role, reason) {
  var n = crescLoginFailCount_(username) + 1;
  try {
    CacheService.getScriptCache()
      .put(crescLoginFailKey_(username), String(n), CRESC_LOGIN_FAIL_WINDOW_S);
  } catch (e) { /* the attempt still fails; only the counting is best effort */ }

  var locked = n >= CRESC_LOGIN_FAIL_MAX;
  crescAuthAudit_(locked ? CRESC_AUTH_EVENTS.LOCKED : CRESC_AUTH_EVENTS.FAILED,
                  username, role, { reason: reason || 'bad credentials', failures: n });

  if (locked) {
    return 'Too many failed attempts. This account is locked for ' +
           Math.round(CRESC_LOGIN_FAIL_WINDOW_S / 60) + ' minutes.';
  }
  var left = CRESC_LOGIN_FAIL_MAX - n;
  return left <= 2
    ? 'Incorrect password. ' + left + ' attempt' + (left === 1 ? '' : 's') + ' left before this account locks.'
    : 'Incorrect password.';
}

/** A good sign-in clears the count and records the success. */
function crescAuthPassed_(username, role, method) {
  try { CacheService.getScriptCache().remove(crescLoginFailKey_(username)); } catch (e) {}
  crescAuthAudit_(CRESC_AUTH_EVENTS.SUCCESS, username, role, { method: method || 'password' });
}

/**
 * FRONTEND ENTRY. Clears a lock ahead of time. Administrators only.
 *
 * Without this the only remedy for a locked-out nurse mid-shift is to wait
 * fifteen minutes, and an administrator who cannot lift a lock will simply
 * ask for the lock to be removed altogether.
 */
function crescUnlockAccount(token, username) {
  try {
    var actor = crescRequire_(token, 'admin.users');
    try { CacheService.getScriptCache().remove(crescLoginFailKey_(username)); } catch (e) {}
    logAudit_({ username: actor.username, role: actor.role, doctorId: actor.doctorId },
              'LOGIN_UNLOCKED', 'Session', crescStr_(username).toUpperCase(), {});
    return { success: true, message: crescStr_(username) + ' can sign in again.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}


// ---------------------------------------------------------------------------
// SECTION C — READING IT BACK
// ---------------------------------------------------------------------------

/**
 * FRONTEND ENTRY. The sign-in history, newest first.
 *
 * Reads the tail of Audit_Log rather than the whole sheet. Audit_Log is
 * append-only and is the sheet in this project most certain to grow without
 * limit, so getDataRange().getValues() on it is a timeout waiting for a busy
 * month — see docs/SAAS_DATA_STRATEGY.md.
 *
 * @param {string} token
 * @param {{days?:number, username?:string, limit?:number}} [opts]
 */
function crescGetLoginAudit(token, opts) {
  try {
    crescRequire_(token, 'admin.audit');
    var o = opts || {};
    var limit = Math.min(Math.max(parseInt(o.limit, 10) || 200, 1), 1000);
    var days  = Math.min(Math.max(parseInt(o.days, 10) || 30, 1), 365);
    var wantUser = crescStr_(o.username).toUpperCase();
    var since = new Date(Date.now() - days * 86400000);

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Audit_Log');
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, rows: [], count: 0, scanned: 0,
               message: 'No audit entries yet.' };
    }

    // Scan backwards in blocks. A sign-in question is nearly always about the
    // recent past, and the recent past is the end of an append-only sheet.
    var last = sheet.getLastRow();
    var cols = sheet.getLastColumn();
    var header = sheet.getRange(1, 1, 1, cols).getValues()[0];
    var col = {};
    header.forEach(function (h, i) { col[String(h).trim()] = i; });

    var AUTH = {};
    Object.keys(CRESC_AUTH_EVENTS).forEach(function (k) { AUTH[CRESC_AUTH_EVENTS[k]] = true; });
    AUTH['LOGIN_UNLOCKED'] = true;

    var rows = [], scanned = 0, block = 500, cursor = last;
    while (cursor > 1 && rows.length < limit && scanned < 20000) {
      var from = Math.max(2, cursor - block + 1);
      var data = sheet.getRange(from, 1, cursor - from + 1, cols).getValues();
      for (var i = data.length - 1; i >= 0 && rows.length < limit; i--) {
        scanned++;
        var r = data[i];
        var ev = String(r[col['Event']] || '');
        if (!AUTH[ev]) continue;
        var ts = r[col['Timestamp']];
        var when = (ts instanceof Date) ? ts : new Date(ts);
        if (!(when instanceof Date) || isNaN(when.getTime())) continue;
        if (when < since) { cursor = 1; break; }      // older than the window
        var user = String(r[col['Actor_Username']] || '');
        if (wantUser && user.toUpperCase() !== wantUser) continue;
        var detail = {};
        try { detail = JSON.parse(r[col['Details_JSON']] || '{}'); } catch (e) {}
        rows.push({
          at:      Utilities.formatDate(when, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
          event:   ev,
          username: user,
          role:    String(r[col['Actor_Role']] || ''),
          method:  String(detail.method || ''),
          reason:  String(detail.reason || ''),
          failures: detail.failures || 0,
          googleAccount: String(detail.googleAccount || '')
        });
      }
      cursor = from - 1;
    }

    return { success: true, rows: rows, count: rows.length, scanned: scanned,
             windowDays: days, message: '' };
  } catch (err) {
    return { success: false, rows: [], count: 0, message: err.message };
  }
}
