// ============================================================================
// Doctor_Core.gs  —  Crescentia HealthTech
// PHASE 0 + 1 : Multi-Doctor Schema Migration, Header Resolver, Scope Engine
// ----------------------------------------------------------------------------
// DEPENDS ON (already present in Doctors_Engine.gs — do NOT redefine here):
//   getTenantId_()            validateSession_(token)
//   issueSession_(obj)        resolveDoctorByUsername_(username)
//   logAudit_(sess, ...)      getActiveDoctors()
//
// ADD THIS AS A NEW FILE. It does not overwrite anything.
// ============================================================================

var DC_DEFAULT_DOCTOR   = "DOC001";          // legacy attribution target
var DC_ROLES_VIEW_ALL   = ["admin", "receptionist", "reception", "nurse",
                           "accounts", "accountant", "pharmacy", "pharmacist",
                           "lab", "lab technician"];

// ============================================================================
// SECTION A — HEADER-DRIVEN COLUMN RESOLUTION
// Never hard-code an index for a column added by this migration.
// ============================================================================

/** Returns { headerName : zeroBasedIndex } for a sheet's row 1. */
function dc_headerMap_(sheet) {
  if (!sheet || sheet.getLastColumn() === 0) return {};
  var hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  for (var i = 0; i < hdr.length; i++) {
    var key = String(hdr[i] || "").trim();
    if (key && map[key] === undefined) map[key] = i;
  }
  return map;
}

/** Zero-based index of a header, or -1. */
function dc_col_(sheet, headerName) {
  var m = dc_headerMap_(sheet);
  return (m[headerName] === undefined) ? -1 : m[headerName];
}

/** Appends a header column if absent. Returns its ZERO-based index. */
function dc_ensureColumn_(sheet, headerName) {
  var idx = dc_col_(sheet, headerName);
  if (idx !== -1) return idx;
  var newCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, newCol)
       .setValue(headerName)
       .setFontWeight("bold")
       .setBackground("#d9ead3");
  return newCol - 1;
}

/** Creates a sheet with headers if absent. Idempotent. */
function dc_ensureSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
         .setFontWeight("bold")
         .setBackground("#d9ead3");
    sheet.setFrozenRows(1);
  } else {
    headers.forEach(function (h) { dc_ensureColumn_(sheet, h); });
  }
  return sheet;
}

// ============================================================================
// SECTION B — TYPE + TIME CASTING (Data Strictness rule)
// ============================================================================

function dc_str_(v)   { return (v === null || v === undefined) ? "" : String(v).trim(); }
function dc_money_(v) { var n = parseFloat(v); return isNaN(n) ? 0.0 : n; }
function dc_int_(v)   { var n = parseInt(v, 10); return isNaN(n) ? 0 : n; }
function dc_upper_(v) { return dc_str_(v).toUpperCase(); }

/** Any time representation -> "HH:mm" 24h. Survives Sheets' Date coercion. */
function dc_to24_(t) {
  if (!t && t !== 0) return "";
  if (t instanceof Date) {
    return Utilities.formatDate(t, Session.getScriptTimeZone(), "HH:mm");
  }
  var s = dc_upper_(t).replace(/\./g, "").replace(/([0-9])(AM|PM)/, "$1 $2");
  var m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?$/);
  if (!m) return s.substring(0, 5);
  var h = parseInt(m[1], 10);
  var mi = m[2];
  var ap = m[3];
  if (ap === "PM" && h < 12) h += 12;
  if (ap === "AM" && h === 12) h = 0;
  return (h < 10 ? "0" : "") + h + ":" + mi;
}

/** "17:15" -> "05:15 PM" (matches the legacy Appointments display format). */
function dc_to12_(t24) {
  var s = dc_to24_(t24);
  if (!s) return "";
  var parts = s.split(":");
  var h = parseInt(parts[0], 10);
  var mi = parts[1];
  var ap = (h >= 12) ? "PM" : "AM";
  var h12 = h % 12; if (h12 === 0) h12 = 12;
  return (h12 < 10 ? "0" : "") + h12 + ":" + mi + " " + ap;
}

function dc_minutes_(t24) {
  var s = dc_to24_(t24);
  if (!s) return -1;
  var p = s.split(":");
  return (parseInt(p[0], 10) * 60) + parseInt(p[1], 10);
}

function dc_fromMinutes_(mins) {
  var h = Math.floor(mins / 60), m = mins % 60;
  return (h < 10 ? "0" : "") + h + ":" + (m < 10 ? "0" : "") + m;
}

/** Any date cell -> "yyyy-MM-dd". Kills the GMT+0530 drift bug. */
function dc_dateKey_(v) {
  if (!v) return "";
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  var s = dc_str_(v);
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  var d = new Date(s);
  if (isNaN(d.getTime())) return s.substring(0, 10);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

// ============================================================================
// SECTION C — PHASE 0 : SCHEMA MIGRATION  (run once from the Apps Script IDE)
// Idempotent. Safe to re-run. Deletes nothing.
// ============================================================================

/**
 * ENTRY POINT. Run this manually from the editor BEFORE deploying scoped reads.
 * Returns a human-readable report. Verify the report before shipping Phase 3.
 */
function runMultiDoctorMigration() {
  var lock = LockService.getScriptLock();
  var report = [];
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tenant = getTenantId_();

    // --- C1. Doctors master: extend, do not rebuild -------------------------
    var docSheet = ss.getSheetByName("Doctors");
    if (!docSheet) {
      setupDoctorsSheet();                       // from Doctors_Engine.gs
      docSheet = ss.getSheetByName("Doctors");
      report.push("Doctors sheet created and seeded.");
    }
    ["Can_View_All", "Default_Consult_Fee", "Colour_Tag"].forEach(function (h) {
      dc_ensureColumn_(docSheet, h);
    });
    report.push("Doctors: extended columns ensured.");

    // --- C2. New relational sheets -----------------------------------------
    dc_ensureSheet_(ss, "Doctor_Schedules", [
      "Schedule_ID", "Tenant_ID", "Doctor_ID", "Weekday", "Session_Label",
      "Start_Time", "End_Time", "Slot_Minutes", "Max_Per_Slot", "Location", "Status"
    ]);
    dc_ensureSheet_(ss, "Doctor_Schedule_Exceptions", [
      "Exception_ID", "Tenant_ID", "Doctor_ID", "Date", "Type",
      "Start_Time", "End_Time", "Reason", "Created_By", "Created_At"
    ]);
    dc_ensureSheet_(ss, "IP_Care_Team", [
      "Entry_ID", "Tenant_ID", "IP_Number", "Patient_ID", "Doctor_ID",
      "Team_Role", "Active", "Added_By", "Added_At", "Removed_At"
    ]);
    report.push("Doctor_Schedules / Doctor_Schedule_Exceptions / IP_Care_Team ready.");

    // --- C3. Append-only columns on existing clinical sheets ----------------
    var additions = [
      { sheet: "Appointments",     cols: ["Doctor_ID", "Doctor_Name_Snapshot", "Booked_By", "Attribution_Source"] },
      { sheet: "OP_Encounters",    cols: ["Doctor_ID", "Doctor_Signature_Snapshot"] },
      { sheet: "IP_Admissions",    cols: ["Primary_Doctor_ID"] },
      { sheet: "IP_CaseSheets_DB", cols: ["Doctor_ID", "Author_Signature_Snapshot"] },
      { sheet: "IP_Timeline_DB",   cols: ["Author_Doctor_ID", "Author_Signature_Snapshot", "Author_Username"] }
    ];
    additions.forEach(function (spec) {
      var sh = ss.getSheetByName(spec.sheet);
      if (!sh) { report.push("SKIPPED (absent): " + spec.sheet); return; }
      spec.cols.forEach(function (c) { dc_ensureColumn_(sh, c); });
      report.push(spec.sheet + ": columns appended.");
    });

    SpreadsheetApp.flush();
    report.push("--- Columns done. Now run backfillDoctorAttribution() ---");
    return report.join("\n");

  } catch (e) {
    return "MIGRATION FAILED: " + e.message;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Stamps historical rows with DC_DEFAULT_DOCTOR and marks the attribution
 * as inferred, never as recorded. Run AFTER runMultiDoctorMigration().
 * Only writes to cells that are currently blank — safe to re-run.
 */
function backfillDoctorAttribution() {
  var lock = LockService.getScriptLock();
  var report = [];
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var doc = dc_getDoctorById_(DC_DEFAULT_DOCTOR);
    var docName = doc ? doc.name : DC_DEFAULT_DOCTOR;
    var docSig  = doc ? doc.signature : docName;

    report.push(dc_backfillColumn_(ss, "Appointments", "Doctor_ID", DC_DEFAULT_DOCTOR));
    report.push(dc_backfillColumn_(ss, "Appointments", "Doctor_Name_Snapshot", docName));
    report.push(dc_backfillColumn_(ss, "Appointments", "Attribution_Source", "BACKFILL"));
    report.push(dc_backfillColumn_(ss, "OP_Encounters", "Doctor_ID", DC_DEFAULT_DOCTOR));
    report.push(dc_backfillColumn_(ss, "OP_Encounters", "Doctor_Signature_Snapshot", docSig));
    report.push(dc_backfillColumn_(ss, "IP_Admissions", "Primary_Doctor_ID", DC_DEFAULT_DOCTOR));
    report.push(dc_backfillColumn_(ss, "IP_CaseSheets_DB", "Doctor_ID", DC_DEFAULT_DOCTOR));

    SpreadsheetApp.flush();
    report.push("NOTE: rows marked Attribution_Source=BACKFILL are INFERRED, not recorded.");
    return report.join("\n");
  } catch (e) {
    return "BACKFILL FAILED: " + e.message;
  } finally {
    lock.releaseLock();
  }
}

/** Batch-fills blank cells in one column. Single setValues() write. */
function dc_backfillColumn_(ss, sheetName, header, value) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return sheetName + ": absent, skipped.";
  var last = sh.getLastRow();
  if (last < 2) return sheetName + "." + header + ": no data rows.";
  var idx = dc_col_(sh, header);
  if (idx === -1) return sheetName + "." + header + ": column missing, run migration first.";

  var range = sh.getRange(2, idx + 1, last - 1, 1);
  var vals = range.getValues();
  var touched = 0;
  for (var i = 0; i < vals.length; i++) {
    if (dc_str_(vals[i][0]) === "") { vals[i][0] = value; touched++; }
  }
  if (touched > 0) range.setValues(vals);
  return sheetName + "." + header + ": " + touched + " row(s) filled.";
}

/**
 * DESTRUCTIVE. Do not run until the new slot grid is verified against the old.
 * Removes the legacy 'Blocked' pseudo-appointment rows once schedules are live.
 */
function purgeLegacyBlockedRows() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Appointments");
    if (!sh) return "Appointments sheet absent.";
    var data = sh.getDataRange().getValues();
    var removed = 0;
    for (var i = data.length - 1; i >= 1; i--) {
      if (dc_str_(data[i][6]) === "Blocked" && dc_upper_(data[i][1]) === "ADMIN") {
        sh.deleteRow(i + 1);
        removed++;
      }
    }
    SpreadsheetApp.flush();
    return "Removed " + removed + " legacy blocked row(s).";
  } catch (e) {
    return "PURGE FAILED: " + e.message;
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// SECTION D — DOCTOR MASTER READS
// ============================================================================

/** Full doctor record incl. signature + view-all flag. Null if not found. */
function dc_getDoctorById_(doctorId) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Doctors");
    if (!sh) return null;
    var data = sh.getDataRange().getDisplayValues();
    var m = dc_headerMap_(sh);
    var id = dc_upper_(doctorId);
    if (!id) return null;

    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][0]) !== id) continue;
      var canAll = (m["Can_View_All"] !== undefined)
        ? ["TRUE", "YES", "1", "Y"].indexOf(dc_upper_(data[i][m["Can_View_All"]])) !== -1
        : false;
      return {
        doctorId:   dc_str_(data[i][0]),
        tenantId:   dc_str_(data[i][1]),
        name:       dc_str_(data[i][2]),
        specialty:  dc_str_(data[i][3]),
        regNo:      dc_str_(data[i][4]),
        signature:  dc_str_(data[i][5]) || dc_str_(data[i][2]),
        username:   dc_str_(data[i][6]),
        status:     dc_upper_(data[i][7]),
        canViewAll: canAll,
        consultFee: (m["Default_Consult_Fee"] !== undefined)
                      ? dc_money_(data[i][m["Default_Consult_Fee"]]) : 0.0,
        colour:     (m["Colour_Tag"] !== undefined)
                      ? dc_str_(data[i][m["Colour_Tag"]]) : ""
      };
    }
    return null;
  } catch (e) { return null; }
}

/**
 * FRONTEND ENTRY. Feeds every doctor-picker in the SPA.
 * Returns the doctors this session may see + whether the picker is locked.
 */
function getDoctorPickerContext(sessionToken) {
  try {
    var scope = resolveScope_(sessionToken, null);
    if (!scope.ok) return { success: false, message: scope.message, doctors: [] };

    var all = getActiveDoctors();
    var visible = (scope.mode === "ALL")
      ? all
      : all.filter(function (d) { return scope.doctorIds.indexOf(d.doctorId) !== -1; });

    return {
      success: true,
      mode: scope.mode,
      locked: (scope.mode !== "ALL"),
      selfDoctorId: scope.selfDoctorId,
      doctors: visible.map(function (d) {
        var full = dc_getDoctorById_(d.doctorId);
        return {
          doctorId:  d.doctorId,
          name:      d.name,
          specialty: d.specialty,
          colour:    full ? full.colour : ""
        };
      })
    };
  } catch (e) {
    return { success: false, message: "Doctor list unavailable: " + e.message, doctors: [] };
  }
}

// ============================================================================
// SECTION E — SCOPE ENGINE  (the security boundary — fails CLOSED)
// ============================================================================

/**
 * Resolves what this session is allowed to SEE.
 * @param {string} sessionToken   from window.sessionToken
 * @param {string} requestedDoctorId  optional narrowing request from the UI
 * @return {{ok, mode:'ALL'|'SET', doctorIds:[], selfDoctorId, role, username, message}}
 *
 * mode 'ALL'  -> caller may read every doctor's data
 * mode 'SET'  -> caller may read only doctorIds[]
 * Never returns ALL on an invalid session.
 */
function resolveScope_(sessionToken, requestedDoctorId) {
  var sess = validateSession_(sessionToken);
  if (!sess) {
    return { ok: false, mode: "SET", doctorIds: [], selfDoctorId: "",
             role: "", username: "",
             message: "Your session has expired. Please sign in again." };
  }

  var role = dc_str_(sess.role).toLowerCase();
  var self = dc_str_(sess.doctorId);
  var req  = dc_str_(requestedDoctorId);

  var base = { ok: true, role: role, username: dc_str_(sess.username),
               selfDoctorId: self, message: "" };

  // Patients never reach doctor-scoped endpoints.
  if (role === "patient") {
    base.ok = false; base.mode = "SET"; base.doctorIds = [];
    base.message = "Not authorised.";
    return base;
  }

  // Non-clinician operational roles see the whole clinic.
  if (DC_ROLES_VIEW_ALL.indexOf(role) !== -1) {
    if (req) { base.mode = "SET"; base.doctorIds = [req]; }
    else     { base.mode = "ALL"; base.doctorIds = []; }
    return base;
  }

  if (role === "doctor") {
    if (!self) {
      base.ok = false; base.mode = "SET"; base.doctorIds = [];
      base.message = "Your login is not linked to a doctor profile. " +
                     "Ask the administrator to set Linked_Username in the Doctors sheet.";
      return base;
    }
    var prof = dc_getDoctorById_(self);
    var canAll = prof ? prof.canViewAll : false;

    if (!req) {
      if (canAll) { base.mode = "ALL"; base.doctorIds = []; }
      else        { base.mode = "SET"; base.doctorIds = [self]; }
      return base;
    }
    if (req === self || canAll) {
      base.mode = "SET"; base.doctorIds = [req];
      return base;
    }
    base.ok = false; base.mode = "SET"; base.doctorIds = [self];
    base.message = "You can only view your own records.";
    return base;
  }

  // Unknown role -> fail closed.
  base.ok = false; base.mode = "SET"; base.doctorIds = [];
  base.message = "Role '" + role + "' has no defined clinical scope.";
  return base;
}

/** True if this scope permits reading rows belonging to doctorId. */
function dc_inScope_(scope, doctorId) {
  if (!scope || !scope.ok) return false;
  if (scope.mode === "ALL") return true;
  if (!doctorId) return false;            // unattributed rows hidden from scoped views
  return scope.doctorIds.indexOf(dc_str_(doctorId)) !== -1;
}

/**
 * Resolves what this session is allowed to WRITE against.
 * Doctors always write as themselves. Reception/admin may write on behalf of
 * a named doctor, but must name one.
 * @return {{ok, doctorId, signature, name, sess, message}}
 */
function resolveWriteDoctor_(sessionToken, targetDoctorId) {
  var sess = validateSession_(sessionToken);
  if (!sess) {
    return { ok: false, message: "Your session has expired. Please sign in again." };
  }
  var role = dc_str_(sess.role).toLowerCase();
  var self = dc_str_(sess.doctorId);
  var target = dc_str_(targetDoctorId);

  var chosen = "";
  if (role === "doctor") {
    if (!self) return { ok: false, message: "Your login is not linked to a doctor profile." };
    var prof = dc_getDoctorById_(self);
    if (target && target !== self && !(prof && prof.canViewAll)) {
      return { ok: false, message: "You cannot record entries under another doctor." };
    }
    chosen = target || self;
  } else if (DC_ROLES_VIEW_ALL.indexOf(role) !== -1) {
    if (!target) return { ok: false, message: "Select a doctor before saving." };
    chosen = target;
  } else {
    return { ok: false, message: "Role '" + role + "' cannot record clinical entries." };
  }

  var d = dc_getDoctorById_(chosen);
  if (!d) return { ok: false, message: "Doctor '" + chosen + "' not found." };
  if (d.status !== "ACTIVE") return { ok: false, message: d.name + " is not an active doctor." };

  return {
    ok: true, sess: sess,
    doctorId: d.doctorId,
    name: d.name,
    signature: d.signature,       // SNAPSHOT at time of write — never re-derived
    message: ""
  };
}

// ============================================================================
// SECTION F — IP CARE TEAM (cross-consults)
// ============================================================================

function dc_careTeamSheet_() {
  return dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), "IP_Care_Team", [
    "Entry_ID", "Tenant_ID", "IP_Number", "Patient_ID", "Doctor_ID",
    "Team_Role", "Active", "Added_By", "Added_At", "Removed_At"
  ]);
}

/** Active care-team members for an admission. */
function getIPCareTeam(ipNumber, sessionToken) {
  try {
    var scope = resolveScope_(sessionToken, null);
    if (!scope.ok) return { success: false, message: scope.message, team: [] };

    var sh = dc_careTeamSheet_();
    var data = sh.getDataRange().getDisplayValues();
    var ip = dc_upper_(ipNumber);
    var team = [];
    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][2]) !== ip) continue;
      if (dc_upper_(data[i][6]) !== "TRUE") continue;
      var d = dc_getDoctorById_(data[i][4]);
      team.push({
        entryId:  dc_str_(data[i][0]),
        doctorId: dc_str_(data[i][4]),
        name:     d ? d.name : dc_str_(data[i][4]),
        specialty: d ? d.specialty : "",
        teamRole: dc_upper_(data[i][5])
      });
    }
    return { success: true, team: team };
  } catch (e) {
    return { success: false, message: "Care team unavailable: " + e.message, team: [] };
  }
}

/** Adds a cross-consult / surgeon / anaesthetist to an admission. */
function addIPCareTeamMember(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var w = resolveWriteDoctor_(sessionToken, payload.doctorId);
    if (!w.ok) return { success: false, message: w.message };

    var ip = dc_upper_(payload.ipNumber);
    if (!ip) return { success: false, message: "IP number is required." };

    var teamRole = dc_upper_(payload.teamRole) || "CROSS_CONSULT";
    var valid = ["PRIMARY", "CROSS_CONSULT", "SURGEON", "ANAESTHETIST"];
    if (valid.indexOf(teamRole) === -1) {
      return { success: false, message: "Invalid team role." };
    }

    var sh = dc_careTeamSheet_();
    var data = sh.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][2]) === ip &&
          dc_upper_(data[i][4]) === dc_upper_(w.doctorId) &&
          dc_upper_(data[i][6]) === "TRUE") {
        return { success: false, message: w.name + " is already on this care team." };
      }
    }

    sh.appendRow([
      String("CT-" + Utilities.getUuid().substring(0, 8).toUpperCase()),
      String(getTenantId_()),
      String(ip),
      String(dc_upper_(payload.patientId)),
      String(w.doctorId),
      String(teamRole),
      "TRUE",
      String(w.sess.username),
      new Date(),
      ""
    ]);

    SpreadsheetApp.flush();
    logAudit_(w.sess, "CARE_TEAM_ADD", "IP_Admission", ip,
              { doctorId: w.doctorId, teamRole: teamRole });
    return { success: true, message: w.name + " added to the care team." };

  } catch (e) {
    return { success: false, message: "Could not add care-team member: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/** Soft-removes a care-team member. Rows are never deleted (audit trail). */
function removeIPCareTeamMember(entryId, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sess = validateSession_(sessionToken);
    if (!sess) return { success: false, message: "Your session has expired." };

    var sh = dc_careTeamSheet_();
    var m = dc_headerMap_(sh);
    var data = sh.getDataRange().getDisplayValues();
    var target = dc_upper_(entryId);

    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][0]) !== target) continue;
      if (dc_upper_(data[i][5]) === "PRIMARY") {
        return { success: false, message: "Change the primary consultant on the admission instead." };
      }
      sh.getRange(i + 1, m["Active"] + 1).setValue("FALSE");
      sh.getRange(i + 1, m["Removed_At"] + 1).setValue(new Date());
      SpreadsheetApp.flush();
      logAudit_(sess, "CARE_TEAM_REMOVE", "IP_Admission", dc_str_(data[i][2]),
                { entryId: entryId });
      return { success: true, message: "Removed from care team." };
    }
    return { success: false, message: "Care-team entry not found." };
  } catch (e) {
    return { success: false, message: "Could not remove member: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/** True if doctorId may write clinical notes on this admission. */
function dc_isOnCareTeam_(ipNumber, doctorId) {
  try {
    var sh = dc_careTeamSheet_();
    var data = sh.getDataRange().getDisplayValues();
    var ip = dc_upper_(ipNumber), d = dc_upper_(doctorId);
    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][2]) === ip &&
          dc_upper_(data[i][4]) === d &&
          dc_upper_(data[i][6]) === "TRUE") return true;
    }
    return false;
  } catch (e) { return false; }
}