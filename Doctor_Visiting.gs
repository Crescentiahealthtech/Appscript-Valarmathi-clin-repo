// ============================================================================
// Doctor_Visiting.gs  —  Crescentia HealthTech
// The visiting consultant: one shared profile, a different real person on it
// each time, and a record of which.
// ----------------------------------------------------------------------------
// THE PROBLEM THIS SOLVES
//
// A visiting consultant comes in for a session, sees patients, and leaves.
// Giving each one a permanent Doctors row and a login means an administrator
// has to create and later disable an account for a person who may never come
// back — and in practice nobody does, so the ward writes everything under
// "Dr. Duty MO" and the record no longer says who saw the patient.
//
// So there is ONE visiting slot (Is_Visiting on the Doctors sheet — DOC004 by
// default), it has no name and no registration number of its own, and the
// person using it TYPES THEIRS IN AFTER SIGNING IN. Until they do, the slot
// cannot write anything.
//
// WHY THE DECLARATION IS PER SESSION, NOT PER NOTE
//
// A consultant types their name once, at the start of their session, and
// every note, prescription and case sheet they write for the rest of it is
// signed with it. Asking per note guarantees it is eventually skipped or
// mistyped; asking once and holding it makes the signature a property of who
// is sitting there.
//
// WHAT IS RECORDED
//
// Every declaration is a row on Visiting_Sessions: the name, the registration
// number, who was signed in, and when. That row is the audit trail — it is
// what answers "who was Dr. X on the fourteenth" months later, when the
// session is long gone. It is append-only; ending a session closes the row
// rather than deleting it.
//
// THE REGISTRATION NUMBER IS NOT DECORATION. It is what the signature on a
// prescription means, and under the NMC regulations it belongs on one. It is
// required, it is written into the signature line, and it is printed.
// ============================================================================

var DV_SHEET = "Visiting_Sessions";
var DV_HEADERS = [
  "Entry_ID", "Tenant_ID", "Doctor_ID", "Session_Token",
  "Consultant_Name", "Reg_No", "Qualification", "Specialty",
  "Declared_By", "Declared_At", "Ended_At", "Status"
];

/** Cache key for a session's declared identity. */
function dv_cacheKey_(token) { return "VISIT_" + String(token || ""); }

function dv_sheet_() {
  var sh = dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), DV_SHEET, DV_HEADERS);
  sh.setFrozenRows(1);
  return sh;
}

/** '' for anything unusable. */
function dv_str_(v) { return (v === null || v === undefined) ? "" : String(v).trim(); }

/**
 * Is this doctor row a shared visiting slot?
 *
 * Driven by the Is_Visiting column, not by the id: a clinic that wants two
 * visiting slots, or that numbered theirs differently, needs no code change.
 */
function dv_isVisitingSlot_(doctorId) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Doctors");
    if (!sh) return false;
    var m = dc_headerMap_(sh);
    if (m["Is_Visiting"] === undefined) return false;
    var data = dc_sheetValues_(sh);
    var id = dc_upper_(doctorId);
    if (!id) return false;
    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][0]) !== id) continue;
      return ["TRUE", "YES", "1", "Y"].indexOf(dc_upper_(data[i][m["Is_Visiting"]])) !== -1;
    }
    return false;
  } catch (e) { return false; }
}

/**
 * The identity declared for this session, or null.
 *
 * Cache first, sheet second — the same shape as dc_validateSession_, and for
 * the same reason: CacheService can evict at any time and a consultant being
 * asked to retype their registration number mid-clinic is the failure this
 * file exists to prevent.
 *
 * @return {{name,regNo,qualification,specialty,signature,doctorId}|null}
 */
function dv_identityFor_(sessionToken) {
  var token = dv_str_(sessionToken);
  if (!token) return null;

  try {
    var cached = CacheService.getScriptCache().get(dv_cacheKey_(token));
    if (cached) return JSON.parse(cached);
  } catch (e) { /* fall through to the sheet */ }

  try {
    var sh = dv_sheet_();
    var data = dc_sheetValues_(sh);
    if (!data || data.length < 2) return null;
    var m = dc_headerMap_(sh);

    // Last match wins: a consultant may re-declare (a corrected spelling, a
    // second person taking over the slot) and the newest row is the truth.
    for (var i = data.length - 1; i >= 1; i--) {
      if (dv_str_(data[i][m["Session_Token"]]) !== token) continue;
      if (dc_upper_(data[i][m["Status"]]) !== "ACTIVE") return null;
      var rec = dv_recordFrom_(data[i], m);
      try {
        CacheService.getScriptCache().put(dv_cacheKey_(token), JSON.stringify(rec), 21600);
      } catch (e2) { /* best effort */ }
      return rec;
    }
    return null;
  } catch (e) { return null; }
}

/** One sheet row -> the identity object every caller uses. */
function dv_recordFrom_(row, m) {
  var name = dv_str_(row[m["Consultant_Name"]]);
  var reg  = dv_str_(row[m["Reg_No"]]);
  var qual = dv_str_(row[m["Qualification"]]);
  return {
    doctorId:      dv_str_(row[m["Doctor_ID"]]),
    name:          name,
    regNo:         reg,
    qualification: qual,
    specialty:     dv_str_(row[m["Specialty"]]),
    // The signature line as it will be printed. Built once, here, so a
    // prescription and a case sheet cannot disagree about it.
    signature:     dv_signature_(name, qual, reg),
    declaredAt:    dv_str_(row[m["Declared_At"]])
  };
}

/** "Dr. A Kumar, MD (Reg. No. 12345)" — whatever parts exist. */
function dv_signature_(name, qualification, regNo) {
  var out = dv_str_(name);
  if (!out) return "";
  if (!/^dr\.?\s/i.test(out)) out = "Dr. " + out;
  var q = dv_str_(qualification);
  if (q) out += ", " + q;
  var r = dv_str_(regNo);
  if (r) out += " (Reg. No. " + r + ")";
  return out;
}

// ---------------------------------------------------------------------------
// WHEN A DECLARATION ENDS
//
// A row used to close only when the consultant pressed "End session". Signing
// out, or simply letting the session run out, left it ACTIVE for ever — so the
// register said "still signed in" about somebody who had gone home hours ago,
// and the next note written for the slot on their behalf would have carried
// their name. A declaration now lives exactly as long as the sign-in it was
// made under: sign-out ends it (crescLogSignOut), and an expired or revoked
// session closes it the next time anything reads the register.
// ---------------------------------------------------------------------------

/**
 * Which of these session tokens are still signed in. One read of Sessions.
 * @return {Object} token -> {live:boolean, endedAt:Date|null}
 */
function dv_sessionStates_(tokens) {
  var out = {};
  var want = {};
  (tokens || []).forEach(function (t) { if (t) want[t] = true; });
  if (!Object.keys(want).length) return out;
  var now = Date.now();
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sessions");
    var data = sh ? sh.getDataRange().getValues() : [];
    var m = sh ? dc_headerMap_(sh) : {};
    for (var i = 1; i < data.length; i++) {
      var tok = dv_str_(data[i][m["Token"] !== undefined ? m["Token"] : 0]);
      if (!want[tok]) continue;
      var status = dc_upper_(data[i][m["Status"]]);
      var exp = data[i][m["Expires_At"]];
      exp = (exp instanceof Date) ? exp : new Date(exp);
      var seen = data[i][m["Last_Seen"]];
      seen = (seen instanceof Date) ? seen : new Date(seen);
      var live = status === "ACTIVE" && !isNaN(exp.getTime()) && exp.getTime() > now;
      out[tok] = { live: live,
                   endedAt: live ? null : (!isNaN(seen.getTime()) ? seen : (isNaN(exp.getTime()) ? new Date() : exp)) };
    }
  } catch (e) { /* unreadable: say nothing rather than guess */ return {}; }
  // A token the Sessions sheet has never seen may still be in the cache.
  Object.keys(want).forEach(function (t) {
    if (out[t]) return;
    var cached = null;
    try { cached = CacheService.getScriptCache().get("SESS_" + t); } catch (e) {}
    out[t] = cached ? { live: true, endedAt: null } : { live: false, endedAt: new Date() };
  });
  return out;
}

/** Closes every ACTIVE declaration whose sign-in has ended. Returns how many. */
function dv_closeStale_() {
  try {
    var sh = dv_sheet_();
    var data = dc_sheetValues_(sh);
    if (!data || data.length < 2) return 0;
    var m = dc_headerMap_(sh);
    var active = [];
    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][m["Status"]]) === "ACTIVE") active.push(i);
    }
    if (!active.length) return 0;
    var states = dv_sessionStates_(active.map(function (i) { return dv_str_(data[i][m["Session_Token"]]); }));
    var closed = 0;
    active.forEach(function (i) {
      var st = states[dv_str_(data[i][m["Session_Token"]])];
      if (!st || st.live) return;
      sh.getRange(i + 1, m["Status"] + 1).setValue("EXPIRED");
      sh.getRange(i + 1, m["Ended_At"] + 1).setValue(st.endedAt || new Date());
      closed++;
    });
    if (closed) dc_invalidate_(DV_SHEET);
    return closed;
  } catch (e) { return 0; }
}

/** Ends this token's declaration, if it has one. Called at sign-out. Never throws. */
function dv_endForToken_(token, why) {
  try {
    token = dv_str_(token);
    if (!token) return 0;
    var sh = dv_sheet_();
    var data = dc_sheetValues_(sh);
    var m = dc_headerMap_(sh);
    var closed = 0, now = new Date();
    for (var i = 1; i < data.length; i++) {
      if (dv_str_(data[i][m["Session_Token"]]) !== token) continue;
      if (dc_upper_(data[i][m["Status"]]) !== "ACTIVE") continue;
      sh.getRange(i + 1, m["Status"] + 1).setValue(why || "ENDED");
      sh.getRange(i + 1, m["Ended_At"] + 1).setValue(now);
      closed++;
    }
    if (closed) dc_invalidate_(DV_SHEET);
    try { CacheService.getScriptCache().remove(dv_cacheKey_(token)); } catch (e) {}
    return closed;
  } catch (e) { return 0; }
}

/**
 * Who is on (or was last on) this visiting slot — for somebody ELSE recording
 * for it: the ward adding the slot's note on the consultant's behalf, an
 * administrator transcribing a verbal order.
 *
 * The consultant signed in and declared themselves; the recorder should not
 * be asked to type that again. Prefers a declaration whose sign-in is still
 * live; otherwise the most recent one in the last `hours` (default 24), so a
 * note transcribed after the consultant left still carries their name.
 *
 * @return {{name,regNo,qualification,specialty,signature,doctorId,live:boolean}|null}
 */
function dv_latestIdentityForSlot_(doctorId, hours) {
  try {
    dv_closeStale_();
    var sh = dv_sheet_();
    var data = dc_sheetValues_(sh);
    if (!data || data.length < 2) return null;
    var m = dc_headerMap_(sh);
    var id = dc_upper_(doctorId);
    var cutoff = Date.now() - (hours || 24) * 3600 * 1000;
    var recent = null;
    for (var i = data.length - 1; i >= 1; i--) {
      if (dc_upper_(data[i][m["Doctor_ID"]]) !== id) continue;
      var st = dc_upper_(data[i][m["Status"]]);
      var at = data[i][m["Declared_At"]];
      at = (at instanceof Date) ? at : new Date(at);
      if (st === "ACTIVE") {
        var live = dv_recordFrom_(data[i], m); live.live = true; return live;
      }
      if (!recent && !isNaN(at.getTime()) && at.getTime() >= cutoff && st !== "SUPERSEDED") {
        recent = dv_recordFrom_(data[i], m); recent.live = false;
      }
    }
    return recent;
  } catch (e) { return null; }
}

// ---------------------------------------------------------------------------
// FRONTEND ENTRY POINTS
// ---------------------------------------------------------------------------

/**
 * Does this session need to declare who it is, and has it?
 *
 * Called once after sign-in. Every answer is a complete instruction to the
 * client, so the browser never has to infer anything from a role name.
 *
 * @return {{success, required:boolean, declared:boolean, identity:Object,
 *           doctorId:string, slotName:string, message:string}}
 */
function getVisitingConsultantState(sessionToken) {
  try {
    var sess = dc_validateSession_(sessionToken);
    if (!sess) {
      return { success: false, required: false, declared: false, identity: null,
               doctorId: "", slotName: "",
               message: "Your session has expired. Please sign in again." };
    }
    var doctorId = dv_str_(sess.doctorId);
    if (!doctorId || !dv_isVisitingSlot_(doctorId)) {
      return { success: true, required: false, declared: false, identity: null,
               doctorId: doctorId, slotName: "", message: "" };
    }

    var slot = dc_getDoctorById_(doctorId);
    var ident = dv_identityFor_(sessionToken);
    return {
      success: true,
      required: true,
      declared: !!ident,
      identity: ident,
      doctorId: doctorId,
      slotName: slot ? slot.name : "Visiting Consultant",
      message: ""
    };
  } catch (err) {
    return { success: false, required: false, declared: false, identity: null,
             doctorId: "", slotName: "", message: err.message };
  }
}

/**
 * Records who is on the visiting slot for this session.
 *
 * VALIDATION IS NOT A FORMALITY HERE. This name and this number go onto
 * prescriptions, so "asdf" and a registration number of "1" are worse than
 * an empty box — they look like a signed document. Both are checked, and the
 * reasons are specific enough to act on.
 *
 * @param {{name,regNo,qualification,specialty}} payload
 * @param {string} sessionToken
 */
function declareVisitingConsultant(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    payload = payload || {};

    var sess = dc_validateSession_(sessionToken);
    if (!sess) {
      return { success: false, message: "Your session has expired. Please sign in again." };
    }
    var doctorId = dv_str_(sess.doctorId);
    if (!doctorId || !dv_isVisitingSlot_(doctorId)) {
      return { success: false,
               message: "This login is not a visiting consultant slot, so there is " +
                        "nothing to declare." };
    }

    var name = dv_str_(payload.name).replace(/\s+/g, " ");
    var reg  = dv_str_(payload.regNo).toUpperCase();
    var qual = dv_str_(payload.qualification);
    var spec = dv_str_(payload.specialty);

    // The name has to be a name.
    var bare = name.replace(/^dr\.?\s+/i, "").trim();
    if (bare.length < 3) {
      return { success: false, message: "Enter the consultant's full name." };
    }
    if (!/[A-Za-z]{2}/.test(bare)) {
      return { success: false, message: "The name does not look like a name." };
    }

    // The registration number has to be a registration number. State medical
    // council numbers vary in shape, so the rule is deliberately loose — but
    // it must contain digits, and a one- or two-digit "number" is a
    // placeholder somebody typed to get past the box.
    var digits = reg.replace(/[^0-9]/g, "");
    if (!reg) {
      return { success: false,
               message: "Enter your medical council registration number. It is " +
                        "printed on every prescription you sign here." };
    }
    if (digits.length < 4) {
      return { success: false,
               message: "\"" + reg + "\" does not look like a registration number " +
                        "(it needs at least four digits). Enter the number exactly " +
                        "as it appears on your council registration." };
    }

    var sh = dv_sheet_();
    var m = dc_headerMap_(sh);
    var now = new Date();

    // Close whatever this session declared before, so dv_identityFor_ never
    // has to choose between two ACTIVE rows for one token.
    var data = dc_sheetValues_(sh);
    for (var i = 1; i < data.length; i++) {
      if (dv_str_(data[i][m["Session_Token"]]) !== dv_str_(sessionToken)) continue;
      if (dc_upper_(data[i][m["Status"]]) !== "ACTIVE") continue;
      sh.getRange(i + 1, m["Status"] + 1).setValue("SUPERSEDED");
      sh.getRange(i + 1, m["Ended_At"] + 1).setValue(now);
    }

    var entryId = "VIS-" + Utilities.formatDate(now, "Asia/Kolkata", "yyMMdd-HHmmss") +
                  "-" + Utilities.getUuid().substring(0, 4).toUpperCase();

    var row = [];
    row[m["Entry_ID"]]        = entryId;
    row[m["Tenant_ID"]]       = getTenantId_();
    row[m["Doctor_ID"]]       = doctorId;
    row[m["Session_Token"]]   = dv_str_(sessionToken);
    row[m["Consultant_Name"]] = name;
    row[m["Reg_No"]]          = reg;
    row[m["Qualification"]]   = qual;
    row[m["Specialty"]]       = spec;
    row[m["Declared_By"]]     = dv_str_(sess.username);
    row[m["Declared_At"]]     = now;
    row[m["Ended_At"]]        = "";
    row[m["Status"]]          = "ACTIVE";
    for (var c = 0; c < DV_HEADERS.length; c++) if (row[c] === undefined) row[c] = "";

    sh.appendRow(row);
    dc_invalidate_(DV_SHEET);

    var ident = {
      doctorId: doctorId, name: name, regNo: reg, qualification: qual,
      specialty: spec, signature: dv_signature_(name, qual, reg),
      declaredAt: Utilities.formatDate(now, "Asia/Kolkata", "yyyy-MM-dd HH:mm")
    };
    try {
      CacheService.getScriptCache().put(dv_cacheKey_(sessionToken), JSON.stringify(ident), 21600);
    } catch (e) { /* the sheet is the durable copy */ }

    try {
      logAudit_(sess, "VISITING_CONSULTANT_DECLARED", "Doctor", doctorId,
                { name: name, regNo: reg, entryId: entryId });
    } catch (e) { /* audit must not break the declaration */ }

    SpreadsheetApp.flush();
    return { success: true, identity: ident,
             message: "Recording as " + ident.signature + "." };

  } catch (err) {
    return { success: false, message: "Could not record the consultant: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Ends the declaration — the consultant has finished their session and the
 * slot is free for the next person. The row stays; only its status changes.
 */
function endVisitingConsultant(sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sess = dc_validateSession_(sessionToken);
    if (!sess) return { success: false, message: "Your session has expired." };

    var sh = dv_sheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    var now = new Date(), closed = 0;

    for (var i = 1; i < data.length; i++) {
      if (dv_str_(data[i][m["Session_Token"]]) !== dv_str_(sessionToken)) continue;
      if (dc_upper_(data[i][m["Status"]]) !== "ACTIVE") continue;
      sh.getRange(i + 1, m["Status"] + 1).setValue("ENDED");
      sh.getRange(i + 1, m["Ended_At"] + 1).setValue(now);
      closed++;
    }
    dc_invalidate_(DV_SHEET);
    try { CacheService.getScriptCache().remove(dv_cacheKey_(sessionToken)); } catch (e) {}

    SpreadsheetApp.flush();
    return { success: true, message: closed
      ? "Visiting session ended. The next consultant will be asked for their details."
      : "There was no active visiting session to end." };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * The log, for an administrator. Newest first.
 * Requires accounts-grade oversight rather than any clinical login: it names
 * every person who has worked under the slot.
 */
function listVisitingConsultants(sessionToken, limit) {
  try {
    var actor = crescRequire_(sessionToken, ["admin.audit", "admin.users"]);
    dv_closeStale_();
    var sh = dv_sheet_();
    var data = dc_sheetValues_(sh);
    if (!data || data.length < 2) return { success: true, rows: [], message: "" };
    var m = dc_headerMap_(sh);

    var rows = [];
    for (var i = data.length - 1; i >= 1 && rows.length < (limit || 100); i--) {
      if (!dv_str_(data[i][m["Entry_ID"]])) continue;
      rows.push({
        entryId:    dv_str_(data[i][m["Entry_ID"]]),
        doctorId:   dv_str_(data[i][m["Doctor_ID"]]),
        name:       dv_str_(data[i][m["Consultant_Name"]]),
        regNo:      dv_str_(data[i][m["Reg_No"]]),
        qualification: dv_str_(data[i][m["Qualification"]]),
        specialty:  dv_str_(data[i][m["Specialty"]]),
        declaredBy: dv_str_(data[i][m["Declared_By"]]),
        declaredAt: dv_str_(data[i][m["Declared_At"]]),
        endedAt:    dv_str_(data[i][m["Ended_At"]]),
        status:     dv_str_(data[i][m["Status"]])
      });
    }
    return { success: true, rows: rows, actor: actor.username, message: "" };
  } catch (err) {
    return { success: false, rows: [], message: err.message };
  }
}

// ---------------------------------------------------------------------------
// SETUP
// ---------------------------------------------------------------------------

/**
 * ONE-OFF. Adds the Is_Visiting column and creates the visiting slot.
 * Safe to re-run.
 *
 * The slot gets a login of its own (`visiting`) so that several visiting
 * consultants over a week each sign in the same way and each declare
 * themselves. Set its password in whatever the Users sheet uses, exactly as
 * for any other staff login.
 */
function setupVisitingConsultant() {
  crescEditorOnly_('setupVisitingConsultant');
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName("Doctors");
    if (!sh) {
      setupDoctorsSheet();
      sh = ss.getSheetByName("Doctors");
    }
    dc_ensureColumn_(sh, "Is_Visiting");
    dv_sheet_();

    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);

    // Already have a visiting slot? Leave it exactly as the clinic set it up.
    for (var i = 1; i < data.length; i++) {
      if (["TRUE", "YES", "1", "Y"].indexOf(dc_upper_(data[i][m["Is_Visiting"]])) !== -1) {
        return "Visiting slot already present: " + dv_str_(data[i][0]) + " (" +
               dv_str_(data[i][2]) + "). " + DV_SHEET + " is ready.";
      }
    }

    // Next free DOCnnn, so this does not collide with a clinic that already
    // has four doctors.
    var maxN = 0;
    for (var j = 1; j < data.length; j++) {
      var mm = /^DOC(\d+)$/i.exec(dv_str_(data[j][0]));
      if (mm) maxN = Math.max(maxN, parseInt(mm[1], 10));
    }
    var newId = "DOC" + ("000" + (maxN + 1)).slice(-3);

    var row = [];
    row[0] = newId;
    row[1] = getTenantId_();
    row[2] = "Visiting Consultant";
    row[3] = "Visiting";
    row[4] = "";                           // Reg_No: typed in per session
    row[5] = "";                           // Signature_Line: built per session
    row[6] = "visiting";                   // Linked_Username
    row[7] = "ACTIVE";
    row[m["Is_Visiting"]] = "TRUE";
    for (var c = 0; c < sh.getLastColumn(); c++) if (row[c] === undefined) row[c] = "";

    sh.appendRow(row);
    dc_invalidate_("Doctors");
    SpreadsheetApp.flush();

    return "Visiting slot created: " + newId + " (Visiting Consultant), login " +
           "\"visiting\". It has no name and no registration number of its own — " +
           "whoever signs in on it types theirs after logging in, and every note " +
           "they write is signed with it. Create the \"visiting\" login on the " +
           "Users sheet with the role \"doctor\".";

  } catch (e) {
    return "setupVisitingConsultant failed: " + e.message;
  } finally {
    lock.releaseLock();
  }
}
