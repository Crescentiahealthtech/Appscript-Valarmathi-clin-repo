// ============================================================================
// IP_Clinical_Access.gs  —  Crescentia HealthTech
// PHASE 5 + 6 : IP Care Team resolution, note authorship, section-level RBAC
// ----------------------------------------------------------------------------
// This is the single security boundary for the IP Casesheet and IP Notes
// modules. Nothing in those modules may write a clinical row without first
// passing through resolveIPWrite_(), and nothing may read an admission
// without passing through ipc_mayReadAdmission_().
//
// DEPENDS ON Doctor_Core.gs        : dc_*, resolveScope_, resolveWriteDoctor_
//            Doctor_Session_Store  : dc_validateSession_
//            Doctors_Engine.gs     : getTenantId_, logAudit_
//
// DESIGN RULES
//   1. Identity is never taken from the client. payload.author is ignored;
//      the author is derived from the session token, server side.
//   2. A doctor may only write on an admission they are on the care team for.
//      Reading is governed by resolveScope_ as everywhere else.
//   3. Sections a role may not author are STRIPPED, not rejected — a nurse
//      submitting vitals plus a stray medOrders block still gets her vitals
//      saved, and the med orders never reach the pharmacy queue.
//   4. Every gate fails closed.
// ============================================================================

// ---------------------------------------------------------------------------
// SECTION A — SCHEMA
// ---------------------------------------------------------------------------

var IPC_TIMELINE_HEADERS = [
  "Timestamp", "IP_Number", "Patient_ID", "Role_Type",
  "Note_Data_JSON", "Author", "Shift", "Flags",
  "Note_ID", "Author_Doctor_ID", "Author_Signature_Snapshot", "Author_Username"
];

/** IP_Timeline_DB with the full Phase 6 authorship schema. Idempotent. */
function ipc_timelineSheet_() {
  var sh = dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(),
                           "IP_Timeline_DB", IPC_TIMELINE_HEADERS);
  sh.setFrozenRows(1);
  return sh;
}

var IPC_CASESHEET_HEADERS = [
  "Encounter_ID", "IP_Number", "Ward", "Bed", "Patient_ID", "Timestamp",
  "Patient Name", "Age", "Sex",
  "Sys_BP", "Dia_BP", "PR", "SpO2", "Temp", "Height", "Weight",
  "Chief_Complaints", "History",
  "Pallor", "Icterus", "Cyanosis", "Clubbing", "Edema", "Other GE findings",
  "CVS", "RS", "PA", "CNS", "Primary Diagnosis",
  "Prescription_JSON", "Lab_Orders_JSON",
  "Outside Lab Records", "Radiological records", "Advice",
  "Doctor's Name", "Doctor_ID", "Author_Signature_Snapshot", "Author_Username"
];

/** IP_CaseSheets_DB with the full Phase 5 attribution schema. Idempotent. */
function ipc_casesheetSheet_() {
  var sh = dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(),
                           "IP_CaseSheets_DB", IPC_CASESHEET_HEADERS);
  sh.setFrozenRows(1);
  return sh;
}

// ---------------------------------------------------------------------------
// SECTION B — RBAC MATRIX
// ---------------------------------------------------------------------------

/**
 * Which note types each application role may AUTHOR.
 * A role absent from this map may author nothing.
 *
 * Deliberately strict: only a doctor signs a clinical progress note, a
 * consultant opinion or a procedure record, because each of those carries
 * prescribing authority downstream (meds reach IP_Pharmacy_Queue, lab orders
 * reach the Lab engine). Admins can run the ward and read everything, but an
 * admin login cannot manufacture a doctor's signature.
 */
var IPC_ROLE_NOTE_TYPES = {
  "doctor":       ["DOCTOR", "CONSULTANT", "PROCEDURE", "QUICK", "HANDOVER"],
  "nurse":        ["NURSE", "QUICK", "HANDOVER"],
  "admin":        ["QUICK", "HANDOVER"],
  "receptionist": ["QUICK"],
  "reception":    ["QUICK"]
};

/**
 * Which noteData sections each role may write, per note type.
 * "*" means every section. Anything not listed is stripped before the row is
 * written and before any side effect (pharmacy / lab routing) is evaluated.
 */
var IPC_SECTION_RBAC = {
  "DOCTOR": {
    "doctor": ["subjectiveObjective", "vitalsReview", "sysExam", "assessment",
               "medOrders", "investigationOrders", "adviceText", "diagnosis",
               "plan", "alertText"]
  },
  "NURSE": {
    "nurse":  ["vitals", "intervention", "observations", "markedMeds",
               "intakeOutput", "alertText"],
    // A doctor recording nursing observations directly (night duty, no nurse
    // on the floor) may log vitals and observations, but never mark a
    // medication administered — that is a nursing act with its own liability.
    "doctor": ["vitals", "intervention", "observations", "intakeOutput", "alertText"]
  },
  "CONSULTANT": {
    "doctor": ["consultantName", "specialty", "findings", "recommendations",
               "medOrders", "investigationOrders", "alertText"]
  },
  "PROCEDURE": {
    "doctor": ["procedureName", "operator", "findings", "complications",
               "anaesthesia", "alertText"]
  },
  "QUICK":    { "*": ["text"] },
  "HANDOVER": { "*": ["handoverText"] }
};

/** Note types whose med/lab orders are allowed to reach downstream queues. */
var IPC_PRESCRIBING_TYPES = ["DOCTOR", "CONSULTANT"];

/** True if this role may author this note type. */
function ipc_roleMayAuthor_(role, roleType) {
  var allowed = IPC_ROLE_NOTE_TYPES[dc_str_(role).toLowerCase()];
  if (!allowed) return false;
  return allowed.indexOf(dc_upper_(roleType)) !== -1;
}

/**
 * Returns a copy of noteData holding only the sections this role may author.
 * @return {{data:Object, stripped:Array<string>}}
 */
function ipc_filterSections_(roleType, role, noteData) {
  var out = {}, stripped = [];
  var src = noteData || {};
  var rule = IPC_SECTION_RBAC[dc_upper_(roleType)];
  if (!rule) return { data: out, stripped: Object.keys(src) };

  var allowed = rule[dc_str_(role).toLowerCase()] || rule["*"] || [];
  var allowAll = (allowed.length === 1 && allowed[0] === "*");

  Object.keys(src).forEach(function (k) {
    if (allowAll || allowed.indexOf(k) !== -1) out[k] = src[k];
    else stripped.push(k);
  });
  return { data: out, stripped: stripped };
}

// ---------------------------------------------------------------------------
// SECTION C — CARE TEAM RESOLUTION (Phase 5)
// ---------------------------------------------------------------------------

/** Doctor_ID for a consultant display name ("Dr. Valarmathi" -> DOC001). */
function ipc_doctorIdByName_(name) {
  var needle = dc_upper_(name).replace(/^DR\.?\s+/, "");
  if (!needle) return "";
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Doctors");
    if (!sh) return "";
    var data = sh.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
      var rowName = dc_upper_(data[i][2]).replace(/^DR\.?\s+/, "");
      if (rowName && rowName === needle) return dc_str_(data[i][0]);
    }
  } catch (e) { /* fall through */ }
  return "";
}

/** The admission row for an IP number, as display values, or null. */
function ipc_admissionRow_(ipNumber) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("IP_Admissions");
    if (!sh) return null;
    var data = sh.getDataRange().getDisplayValues();
    var ip = dc_upper_(ipNumber);
    if (!ip) return null;
    var m = dc_headerMap_(sh);
    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][0]) !== ip) continue;
      return { row: data[i], rowNumber: i + 1, headers: m, sheet: sh };
    }
  } catch (e) { /* fall through */ }
  return null;
}

/**
 * The primary consultant's Doctor_ID for an admission.
 * Prefers the Primary_Doctor_ID column; falls back to resolving the legacy
 * Consultant name column, and writes the resolved id back so the lookup
 * happens once per admission rather than once per note.
 */
function ipc_primaryDoctorId_(ipNumber) {
  var adm = ipc_admissionRow_(ipNumber);
  if (!adm) return "";

  var idxPrimary = adm.headers["Primary_Doctor_ID"];
  if (idxPrimary !== undefined) {
    var stored = dc_str_(adm.row[idxPrimary]);
    if (stored) return stored;
  }

  var resolved = ipc_doctorIdByName_(adm.row[9]);   // [9] Consultant
  if (resolved && idxPrimary !== undefined) {
    try {
      adm.sheet.getRange(adm.rowNumber, idxPrimary + 1).setValue(String(resolved));
    } catch (e) { /* best effort — never block the caller */ }
  }
  return resolved;
}

/**
 * Guarantees the admission's primary consultant sits on IP_Care_Team.
 * Legacy admissions predate the care-team sheet entirely; without this every
 * one of them would reject its own consultant's notes.
 */
function ipc_ensurePrimaryOnCareTeam_(ipNumber) {
  var primary = ipc_primaryDoctorId_(ipNumber);
  if (!primary) return "";
  if (dc_isOnCareTeam_(ipNumber, primary)) return primary;

  try {
    var adm = ipc_admissionRow_(ipNumber);
    dc_careTeamSheet_().appendRow([
      String("CT-" + Utilities.getUuid().substring(0, 8).toUpperCase()),
      String(getTenantId_()),
      String(dc_upper_(ipNumber)),
      String(adm ? dc_upper_(adm.row[1]) : ""),
      String(primary),
      "PRIMARY",
      "TRUE",
      "SYSTEM",
      new Date(),
      ""
    ]);
  } catch (e) { /* best effort */ }
  return primary;
}

/**
 * May this doctor write clinical entries on this admission?
 * Primary consultant, any active care-team member, or a doctor flagged
 * Can_View_All (the on-call / covering consultant).
 */
function ipc_mayWriteOnAdmission_(ipNumber, doctorId) {
  var d = dc_upper_(doctorId);
  if (!d) return false;

  var primary = ipc_ensurePrimaryOnCareTeam_(ipNumber);
  if (primary && dc_upper_(primary) === d) return true;
  if (dc_isOnCareTeam_(ipNumber, doctorId)) return true;

  var prof = dc_getDoctorById_(doctorId);
  return !!(prof && prof.canViewAll);
}

/**
 * May this scope READ this admission?
 * Operational roles and view-all doctors see the whole ward. A scoped doctor
 * sees admissions they are the primary for or are consulting on.
 */
function ipc_mayReadAdmission_(scope, ipNumber) {
  if (!scope || !scope.ok) return false;
  if (scope.mode === "ALL") return true;
  for (var i = 0; i < scope.doctorIds.length; i++) {
    if (ipc_mayWriteOnAdmission_(ipNumber, scope.doctorIds[i])) return true;
  }
  return false;
}

/**
 * Pre-computes read visibility for a whole ward in one pass.
 * ipc_mayReadAdmission_ re-reads IP_Care_Team per admission; on a 40-bed ward
 * that is 40 full-sheet scans and the request times out. This reads the
 * care-team sheet and the admissions sheet once and returns a lookup.
 * @return {function(string):boolean}
 */
function ipc_wardVisibilityFilter_(scope) {
  if (!scope || !scope.ok) return function () { return false; };
  if (scope.mode === "ALL") return function () { return true; };

  var mine = {};
  scope.doctorIds.forEach(function (d) {
    var prof = dc_getDoctorById_(d);
    if (prof && prof.canViewAll) mine.__ALL__ = true;
    mine[dc_upper_(d)] = true;
  });
  if (mine.__ALL__) return function () { return true; };

  var visible = {};

  // Care-team membership — one scan.
  try {
    var ct = dc_careTeamSheet_().getDataRange().getDisplayValues();
    for (var i = 1; i < ct.length; i++) {
      if (dc_upper_(ct[i][6]) !== "TRUE") continue;
      if (mine[dc_upper_(ct[i][4])]) visible[dc_upper_(ct[i][2])] = true;
    }
  } catch (e) { /* fall through to the admissions pass */ }

  // Primary consultant on the admission itself — one scan, covers legacy rows
  // that were never migrated onto the care-team sheet.
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("IP_Admissions");
    if (sh) {
      var m = dc_headerMap_(sh);
      var idxPrimary = m["Primary_Doctor_ID"];
      var data = sh.getDataRange().getDisplayValues();
      var nameCache = {};
      for (var j = 1; j < data.length; j++) {
        var ipKey = dc_upper_(data[j][0]);
        if (!ipKey || visible[ipKey]) continue;

        var pid = (idxPrimary !== undefined) ? dc_str_(data[j][idxPrimary]) : "";
        if (!pid) {
          var nm = dc_upper_(data[j][9]);
          if (nameCache[nm] === undefined) nameCache[nm] = ipc_doctorIdByName_(nm);
          pid = nameCache[nm];
        }
        if (pid && mine[dc_upper_(pid)]) visible[ipKey] = true;
      }
    }
  } catch (e) { /* fail closed for anything unresolved */ }

  return function (ipNumber) { return !!visible[dc_upper_(ipNumber)]; };
}

// ---------------------------------------------------------------------------
// SECTION D — THE WRITE GATE (Phase 6)
// ---------------------------------------------------------------------------

/**
 * Resolves the authenticated author for an IP clinical write and applies
 * section-level RBAC. This is the only function the IP modules may use to
 * establish who is writing.
 *
 * @param {string} sessionToken
 * @param {string} ipNumber
 * @param {string} roleType        DOCTOR | NURSE | CONSULTANT | PROCEDURE | QUICK | HANDOVER
 * @param {Object} noteData        raw client sections (filtered on the way out)
 * @return {{ok, message, sess, role, username, displayName, doctorId,
 *           signature, authorLabel, noteData, stripped, mayPrescribe}}
 */
function resolveIPWrite_(sessionToken, ipNumber, roleType, noteData) {
  var fail = function (msg) { return { ok: false, message: msg }; };

  var sess = dc_validateSession_(sessionToken);
  if (!sess) return fail("Your session has expired. Please sign in again.");

  var role = dc_str_(sess.role).toLowerCase();
  var type = dc_upper_(roleType) || "QUICK";

  if (role === "patient") return fail("Not authorised.");
  if (!ipc_roleMayAuthor_(role, type)) {
    return fail("Your role (" + role + ") cannot author a " + type + " note.");
  }

  var ip = dc_upper_(ipNumber);
  if (!ip) return fail("An IP number is required.");
  if (!ipc_admissionRow_(ip)) return fail("Admission " + ip + " was not found.");

  // ---- doctor identity + care-team gate ---------------------------------
  var doctorId = "", signature = "", displayName = dc_str_(sess.displayName) || dc_str_(sess.username);

  if (role === "doctor") {
    doctorId = dc_str_(sess.doctorId);
    if (!doctorId) {
      return fail("Your login is not linked to a doctor profile. " +
                  "Ask the administrator to set Linked_Username in the Doctors sheet.");
    }
    var prof = dc_getDoctorById_(doctorId);
    if (!prof) return fail("Doctor profile '" + doctorId + "' not found.");
    if (prof.status !== "ACTIVE") return fail(prof.name + " is not an active doctor.");

    if (!ipc_mayWriteOnAdmission_(ip, doctorId)) {
      return fail("You are not on the care team for " + ip +
                  ". Ask the primary consultant to add you as a cross-consult.");
    }
    // Signature is SNAPSHOTTED here and never re-derived at print time.
    signature   = prof.signature;
    displayName = prof.name;
  }

  var authorLabel = (role === "doctor")   ? ("Dr. " + displayName.replace(/^Dr\.?\s+/i, ""))
                  : (role === "nurse")    ? ("Staff Nurse " + displayName)
                  : displayName;

  var filtered = ipc_filterSections_(type, role, noteData);

  return {
    ok: true,
    message: "",
    sess: sess,
    role: role,
    username: dc_str_(sess.username),
    displayName: displayName,
    doctorId: doctorId,
    signature: signature,
    authorLabel: authorLabel,
    noteData: filtered.data,
    stripped: filtered.stripped,
    mayPrescribe: (role === "doctor" && IPC_PRESCRIBING_TYPES.indexOf(type) !== -1)
  };
}

/** Read gate shared by every IP fetcher. */
function resolveIPRead_(sessionToken, ipNumber) {
  var scope = resolveScope_(sessionToken, null);
  if (!scope.ok) return { ok: false, message: scope.message, scope: scope };
  if (ipNumber && !ipc_mayReadAdmission_(scope, ipNumber)) {
    return { ok: false, message: "This admission is not under your care.", scope: scope };
  }
  return { ok: true, message: "", scope: scope };
}

// ---------------------------------------------------------------------------
// SECTION E — CARE TEAM UI ENDPOINT (Phase 5)
// ---------------------------------------------------------------------------

/**
 * FRONTEND ENTRY. Everything the IP Notes / Casesheet care-team panel needs:
 * the current team, the doctors that could be added, and what this session
 * is permitted to do on this admission.
 */
function getIPCareTeamPanel(ipNumber, sessionToken) {
  try {
    var gate = resolveIPRead_(sessionToken, ipNumber);
    if (!gate.ok) return { success: false, message: gate.message, team: [] };

    ipc_ensurePrimaryOnCareTeam_(ipNumber);

    var team = getIPCareTeam(ipNumber, sessionToken);
    if (!team.success) return team;

    var sess = dc_validateSession_(sessionToken);
    var role = dc_str_(sess && sess.role).toLowerCase();
    var selfId = dc_str_(sess && sess.doctorId);
    var onTeam = team.team.filter(function (t) {
      return dc_upper_(t.doctorId) === dc_upper_(selfId);
    }).length > 0;

    var already = {};
    team.team.forEach(function (t) { already[dc_upper_(t.doctorId)] = true; });

    var picker = getDoctorPickerContext(sessionToken);
    var addable = (picker.success ? picker.doctors : []).filter(function (d) {
      return !already[dc_upper_(d.doctorId)];
    });

    return {
      success: true,
      ipNumber: dc_upper_(ipNumber),
      team: team.team,
      addable: addable,
      selfDoctorId: selfId,
      role: role,
      onCareTeam: onTeam,
      canWrite: (role !== "doctor") || ipc_mayWriteOnAdmission_(ipNumber, selfId),
      // Only the primary consultant and admins manage team membership.
      canManageTeam: (role === "admin") ||
                     (role === "doctor" && selfId &&
                      dc_upper_(ipc_primaryDoctorId_(ipNumber)) === dc_upper_(selfId))
    };
  } catch (e) {
    return { success: false, message: "Care team unavailable: " + e.message, team: [] };
  }
}

/**
 * Adds a cross-consult to an admission. Wraps addIPCareTeamMember so the
 * patient id is taken from the admission rather than trusted from the client,
 * and so only the primary consultant or an admin may extend the team.
 */
function addIPCrossConsult(payload, sessionToken) {
  try {
    var ip = dc_upper_(payload && payload.ipNumber);
    var panel = getIPCareTeamPanel(ip, sessionToken);
    if (!panel.success) return { success: false, message: panel.message };
    if (!panel.canManageTeam) {
      return { success: false,
               message: "Only the primary consultant or an administrator can add a cross-consult." };
    }

    var adm = ipc_admissionRow_(ip);
    return addIPCareTeamMember({
      ipNumber:  ip,
      patientId: adm ? adm.row[1] : "",
      doctorId:  dc_str_(payload.doctorId),
      teamRole:  dc_upper_(payload.teamRole) || "CROSS_CONSULT"
    }, sessionToken);
  } catch (e) {
    return { success: false, message: "Could not add cross-consult: " + e.message };
  }
}

/** Removes a cross-consult. Same gate as adding one. */
function removeIPCrossConsult(payload, sessionToken) {
  try {
    var panel = getIPCareTeamPanel(dc_upper_(payload && payload.ipNumber), sessionToken);
    if (!panel.success) return { success: false, message: panel.message };
    if (!panel.canManageTeam) {
      return { success: false,
               message: "Only the primary consultant or an administrator can change the care team." };
    }
    return removeIPCareTeamMember(dc_str_(payload.entryId), sessionToken);
  } catch (e) {
    return { success: false, message: "Could not remove cross-consult: " + e.message };
  }
}

/**
 * FRONTEND ENTRY. What the IP module may render for this session — used to
 * hide the note buttons a role cannot use, rather than letting the user fill
 * in a form the server will reject.
 */
function getIPNotePermissions(sessionToken) {
  try {
    var sess = dc_validateSession_(sessionToken);
    if (!sess) return { success: false, message: "Your session has expired.", noteTypes: [] };
    var role = dc_str_(sess.role).toLowerCase();
    return {
      success: true,
      role: role,
      doctorId: dc_str_(sess.doctorId),
      displayName: dc_str_(sess.displayName) || dc_str_(sess.username),
      noteTypes: (IPC_ROLE_NOTE_TYPES[role] || []).slice(),
      canPrescribe: (role === "doctor")
    };
  } catch (e) {
    return { success: false, message: e.message, noteTypes: [] };
  }
}

// ---------------------------------------------------------------------------
// SECTION F — PHASE 5/6 SHIP GATE
// Run these from the Apps Script editor BEFORE deploying. They write nothing.
// ---------------------------------------------------------------------------

/**
 * Reports header drift on the two IP clinical sheets.
 *
 * ipc_timelineSheet_() and ipc_casesheetSheet_() append any header they do
 * not find. If a live sheet spells a column differently — a trailing space, a
 * renamed "Doctor's Name" — the ensure step would silently create a SECOND
 * column and every subsequent write would land in the wrong one. Eyeball this
 * report first; it is the Phase 5 ship gate.
 */
function verifyIPClinicalSchema() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = [];

  [{ name: "IP_Timeline_DB",   want: IPC_TIMELINE_HEADERS },
   { name: "IP_CaseSheets_DB", want: IPC_CASESHEET_HEADERS }].forEach(function (spec) {

    out.push("=== " + spec.name + " ===");
    var sh = ss.getSheetByName(spec.name);
    if (!sh) { out.push("  absent — will be created cleanly on first write."); return; }

    var actual = (sh.getLastColumn() === 0)
      ? []
      : sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h); });

    var have = {};
    actual.forEach(function (h) { have[h.trim()] = true; });

    var missing = spec.want.filter(function (h) { return !have[h]; });
    out.push(missing.length
      ? "  WILL APPEND: " + missing.join(", ")
      : "  all expected headers present.");

    // A header that differs only by whitespace or case is drift, not a new
    // column, and appending its canonical twin would split the data.
    var suspect = [];
    actual.forEach(function (h) {
      var t = h.trim();
      if (spec.want.indexOf(t) !== -1) {
        if (h !== t) suspect.push("'" + h + "' has stray whitespace");
        return;
      }
      spec.want.forEach(function (w) {
        if (t.toLowerCase() === w.toLowerCase()) {
          suspect.push("'" + h + "' differs from '" + w + "' only by case");
        }
      });
    });
    if (suspect.length) {
      out.push("  ⚠ DRIFT — fix these by hand before deploying:");
      suspect.forEach(function (x) { out.push("     " + x); });
    }

    var seen = {}, dupes = [];
    actual.forEach(function (h) {
      var t = h.trim();
      if (!t) return;
      if (seen[t]) dupes.push(t); else seen[t] = true;
    });
    if (dupes.length) out.push("  ⚠ DUPLICATE HEADERS: " + dupes.join(", "));
  });

  var report = out.join("\n");
  Logger.log(report);
  return report;
}

/**
 * Reports which active admissions have no resolvable primary consultant.
 * Those admissions accept no doctor's notes until the Consultant name on the
 * admission matches a row in the Doctors sheet, or Primary_Doctor_ID is set
 * by hand. Run it after verifyIPClinicalSchema().
 */
function verifyIPCareTeamCoverage() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("IP_Admissions");
  if (!sh) return "IP_Admissions sheet absent.";

  var data = sh.getDataRange().getDisplayValues();
  var m = dc_headerMap_(sh);
  var idxPrimary = m["Primary_Doctor_ID"];
  var orphans = [], ok = 0;

  for (var i = 1; i < data.length; i++) {
    if (dc_upper_(data[i][11]) !== "ACTIVE") continue;
    var pid = (idxPrimary !== undefined) ? dc_str_(data[i][idxPrimary]) : "";
    if (!pid) pid = ipc_doctorIdByName_(data[i][9]);
    if (pid) { ok++; continue; }
    orphans.push("  " + dc_str_(data[i][0]) + "  consultant='" + dc_str_(data[i][9]) + "'");
  }

  var report = ok + " active admission(s) have a resolvable primary consultant.\n" +
    (orphans.length
      ? orphans.length + " have NONE — no doctor can write on these until fixed:\n" +
        orphans.join("\n") +
        "\nFix by matching the Consultant name to Doctors.Doctor_Name, or by " +
        "setting Primary_Doctor_ID on the admission row."
      : "Every active admission is covered.");

  Logger.log(report);
  return report;
}
