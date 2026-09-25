// ============================================================================
// Doctor_Session_Store.gs  —  Crescentia HealthTech
// FIXES: sessions expiring long before the advertised timeout.
// ----------------------------------------------------------------------------
// ROOT CAUSE
//   issueSession_() stores the token in CacheService. Apps Script caps cache
//   TTL at 6 hours (a longer value is silently truncated), and cache entries
//   can be evicted at ANY time under memory pressure. A cache is not a session
//   store. Losing the entry logs the user out mid-clinic with no explanation.
//
// FIX
//   A durable "Sessions" sheet backs the cache. Nothing in Auth.html or
//   Doctors_Engine.gs needs to change: the row is written lazily on the first
//   validation after login, while the cache is still warm.
//
// AFTER ADDING THIS FILE, make two edits in Doctor_Core.gs:
//   resolveScope_()        : var sess = validateSession_(sessionToken);
//                          -> var sess = dc_validateSession_(sessionToken);
//   resolveWriteDoctor_()  : var sess = validateSession_(sessionToken);
//                          -> var sess = dc_validateSession_(sessionToken);
//   (deleteScheduleException and bookAppointmentScoped call it too — swap
//    every validateSession_ inside Doctor_Core.gs and
//    Doctor_Schedule_Engine.gs for dc_validateSession_.)
// ============================================================================

var DS_SESSION_HOURS = 8;              // real session lifetime, sheet-backed
var DS_SESSION_MAX_HOURS = 16;         // however active: a double shift, then sign in again
var DS_CACHE_SECONDS = 21600;          // 6h — the Apps Script hard ceiling

function ds_sessionSheet_() {
  return dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), "Sessions", [
    "Token", "Username", "Role", "Doctor_ID", "Display_Name",
    "Issued_At", "Last_Seen", "Expires_At", "Status"
  ]);
}

/**
 * Durable session validation. Drop-in replacement for validateSession_().
 * Order of resolution:
 *   1. CacheService  -> fast path; persist to the sheet if not yet stored
 *   2. Sessions sheet -> cache was evicted; rehydrate it
 *   3. null           -> genuinely expired or revoked (fails closed)
 * Every successful validation slides Last_Seen and Expires_At forward.
 */
/**
 * The human-readable name for a session, whichever field it was written with.
 * issueSession_() sets `name`; the durable Sessions row sets `displayName`.
 * Reading only one of them yields the login username instead of the person.
 */
function dc_sessionName_(sess) {
  if (!sess) return "";
  return dc_str_(sess.displayName) || dc_str_(sess.name) || dc_str_(sess.username);
}

/**
 * THE ACCOUNT, AS IT IS NOW. A session remembers who someone was when they
 * signed in; this asks RBAC.gs whether that account is still switched on and
 * what its role is today, so switching someone off — or changing their role —
 * takes effect on their very next call instead of when the session expires.
 *
 * Returns the session with its role brought up to date, or null when the
 * account may no longer act (its sessions are ended as well).
 */
function ds_applyAccountState_(token, sess) {
  if (!sess || typeof crescAccountState_ !== 'function') return sess;
  var st = crescAccountState_(sess);
  if (!st.ok) {
    try { if (typeof ds_revokeUserSessions_ === 'function') ds_revokeUserSessions_(sess.username, ''); }
    catch (e) {}
    try { CacheService.getScriptCache().remove("SESS_" + token); } catch (e) {}
    return null;
  }
  if (st.role && dc_str_(st.role).toLowerCase() !== dc_str_(sess.role).toLowerCase()) {
    sess.role = dc_str_(st.role).toLowerCase();
  }
  return sess;
}

function dc_validateSession_(sessionToken) {
  var token = dc_str_(sessionToken);
  if (!token) return null;

  // ---- 1. fast path: the cache still holds it -----------------------------
  var sess = null;
  try { sess = validateSession_(token); } catch (e) { sess = null; }

  if (sess) {
    sess = ds_applyAccountState_(token, sess);
    if (!sess) return null;
    ds_touchSession_(token, sess);     // writes the row if it is missing
    if (typeof crescNoteSession_ === 'function') crescNoteSession_(sess);
    return sess;
  }

  // ---- 2. cache miss: fall back to the durable row ------------------------
  try {
    var sh = ds_sessionSheet_();
    if (sh.getLastRow() < 2) return null;

    var data = sh.getDataRange().getDisplayValues();
    var now = new Date();

    for (var i = 1; i < data.length; i++) {
      if (dc_str_(data[i][0]) !== token) continue;
      if (dc_upper_(data[i][8]) !== "ACTIVE") return null;

      // DISPLAY values, so in an Indian-locale spreadsheet these read
      // "25/09/2026 10:00:00". new Date() reads that month-first — or not at
      // all when the day is over 12 — so a live session was judged expired
      // the moment the cache let it go, which is the very logout this file
      // exists to prevent. cresc_parseDate_ reads it day-first.
      var parseAt = function (v) {
        var d = (typeof cresc_parseDate_ === 'function') ? cresc_parseDate_(v) : new Date(v);
        return (d && !isNaN(d.getTime())) ? d : null;
      };
      var expires = parseAt(data[i][7]);
      var issued = parseAt(data[i][5]);
      // An absolute ceiling as well as the sliding one: a token in steady
      // use — or stolen and kept warm — does not live for ever.
      var tooOld = issued && (now.getTime() - issued.getTime()) > DS_SESSION_MAX_HOURS * 3600 * 1000;
      if (!expires || expires < now || tooOld) {
        var m0 = dc_headerMap_(sh);
        sh.getRange(i + 1, m0["Status"] + 1).setValue("EXPIRED");
        return null;
      }

      var revived = {
        username:    dc_str_(data[i][1]),
        role:        dc_str_(data[i][2]),
        doctorId:    dc_str_(data[i][3]),
        // Both spellings: issueSession_() writes `name`, the Sessions sheet
        // stores Display_Name, and callers read whichever they were written
        // against. Keeping them in step here is cheaper than auditing every
        // reader.
        name:        dc_str_(data[i][4]),
        displayName: dc_str_(data[i][4])
      };

      revived = ds_applyAccountState_(token, revived);
      if (!revived) return null;

      // Rehydrate the cache so the next call takes the fast path.
      try {
        CacheService.getScriptCache()
          .put("SESS_" + token, JSON.stringify(revived), DS_CACHE_SECONDS);
      } catch (e) { /* cache is best-effort; the sheet is the truth */ }

      ds_slideExpiry_(sh, i + 1);
      ds_markTouched_(token);
      if (typeof crescNoteSession_ === 'function') crescNoteSession_(revived);
      return revived;
    }
    return null;

  } catch (e) {
    return null;                       // fail closed, never permissive
  }
}

/**
 * How long a session's Sessions row is trusted as fresh before the next call
 * slides its expiry again.
 *
 * THIS WAS THE SLOWEST LINE IN THE APPLICATION. Every validated call — every
 * lab queue refresh, every billing screen, every keystroke that searched a
 * patient — read the ENTIRE Sessions sheet and then wrote two cells to it,
 * before doing the work the user asked for. Two cell writes and a full-sheet
 * read are several hundred milliseconds on a good day, paid on every call by
 * every desk, on a sheet that only grows. The expiry is eight hours; sliding
 * it every two minutes instead of every call loses nothing.
 */
var DS_TOUCH_SECONDS = 120;

function ds_markTouched_(token) {
  try { CacheService.getScriptCache().put("SESSTOUCH_" + token, "1", DS_TOUCH_SECONDS); }
  catch (e) {}
}

/** Writes the row on first sight; otherwise slides the expiry forward. */
function ds_touchSession_(token, sess) {
  try {
    // Touched within the last DS_TOUCH_SECONDS: the row exists and its
    // expiry is hours away. Nothing to do.
    try {
      if (CacheService.getScriptCache().get("SESSTOUCH_" + token)) return;
    } catch (e) { /* no cache: fall through to the sheet, as before */ }

    var sh = ds_sessionSheet_();
    var data = sh.getDataRange().getDisplayValues();

    for (var i = 1; i < data.length; i++) {
      if (dc_str_(data[i][0]) === token) {
        ds_slideExpiry_(sh, i + 1);
        ds_markTouched_(token);
        return;
      }
    }

    var now = new Date();
    var exp = new Date(now.getTime() + DS_SESSION_HOURS * 3600 * 1000);
    sh.appendRow([
      String(token),
      String(dc_str_(sess.username)),
      String(dc_str_(sess.role)),
      String(dc_str_(sess.doctorId)),
      String(dc_sessionName_(sess)),
      now, now, exp, "ACTIVE"
    ]);
    ds_markTouched_(token);
  } catch (e) { /* never block the request over session bookkeeping */ }
}

/** Sliding expiry: an actively used session does not die mid-shift. */
function ds_slideExpiry_(sh, rowNumber) {
  try {
    var m = dc_headerMap_(sh);
    var now = new Date();
    var exp = new Date(now.getTime() + DS_SESSION_HOURS * 3600 * 1000);
    // One write when the two columns sit side by side, as they do on every
    // sheet this file creates; two when somebody has moved them.
    if (m["Expires_At"] === m["Last_Seen"] + 1) {
      sh.getRange(rowNumber, m["Last_Seen"] + 1, 1, 2).setValues([[now, exp]]);
    } else {
      sh.getRange(rowNumber, m["Last_Seen"] + 1).setValue(now);
      sh.getRange(rowNumber, m["Expires_At"] + 1).setValue(exp);
    }
  } catch (e) { /* best effort */ }
}

/** Call from your sign-out handler to revoke server-side. Optional. */
function revokeSession_(sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var token = dc_str_(sessionToken);
    if (!token) return { success: true, message: "Signed out." };

    try { CacheService.getScriptCache().remove("SESS_" + token); } catch (e) {}

    var sh = ds_sessionSheet_();
    var data = sh.getDataRange().getDisplayValues();
    var m = dc_headerMap_(sh);
    for (var i = 1; i < data.length; i++) {
      if (dc_str_(data[i][0]) === token) {
        sh.getRange(i + 1, m["Status"] + 1).setValue("REVOKED");
        break;
      }
    }
    return { success: true, message: "Signed out." };
  } catch (e) {
    return { success: false, message: "Could not revoke session: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Housekeeping. The Sessions sheet grows one row per login; without this it
 * becomes the largest sheet in the workbook within a year.
 * Attach to a weekly time-driven trigger.
 */
function purgeExpiredSessions_() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sh = ds_sessionSheet_();
    if (sh.getLastRow() < 2) return "No sessions to purge.";

    var data = sh.getDataRange().getValues();
    var cutoff = new Date(Date.now() - 7 * 24 * 3600 * 1000);
    var removed = 0;

    for (var i = data.length - 1; i >= 1; i--) {
      var exp = new Date(data[i][7]);
      if (isNaN(exp.getTime()) || exp < cutoff) { sh.deleteRow(i + 1); removed++; }
    }
    return "Purged " + removed + " expired session row(s).";
  } catch (e) {
    return "Purge failed: " + e.message;
  } finally {
    lock.releaseLock();
  }
}

/**
 * FRONTEND ENTRY. "Stay signed in": validating the token is what slides its
 * expiry forward (dc_validateSession_ -> ds_touchSession_), so this is all a
 * renewal needs. { success:false, expired:true } tells the browser the server
 * has already let the session go.
 */
function crescTouchSession(sessionToken) {
  var sess = dc_validateSession_(sessionToken);
  if (!sess) return { success: false, expired: true, message: "Your session has expired." };
  return { success: true, username: dc_str_(sess.username) };
}

/**
 * Signs a user out everywhere except (optionally) one session. Used when a
 * password changes: whoever else was signed in with the old password stops
 * being signed in. Never throws; returns how many sessions ended.
 */
function ds_revokeUserSessions_(username, keepToken) {
  try {
    var who = dc_upper_(username);
    if (!who) return 0;
    var sh = ds_sessionSheet_();
    var data = sh.getDataRange().getDisplayValues();
    var m = dc_headerMap_(sh);
    var n = 0;
    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][m["Username"]]) !== who) continue;
      if (dc_upper_(data[i][m["Status"]]) !== "ACTIVE") continue;
      var tok = dc_str_(data[i][m["Token"]]);
      if (keepToken && tok === dc_str_(keepToken)) continue;
      sh.getRange(i + 1, m["Status"] + 1).setValue("REVOKED");
      try { CacheService.getScriptCache().remove("SESS_" + tok); } catch (e) {}
      n++;
    }
    return n;
  } catch (e) { return 0; }
}
