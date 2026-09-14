// ============================================================================
// Patient_Profile_Edit.gs  —  Crescentia HealthTech
// Correcting a patient's demographics after registration.
// ----------------------------------------------------------------------------
// WHAT WAS MISSING
//
// Registration wrote twenty-two columns and nothing could change any of
// them. The profile modal in Patient Search was read-only, and the only
// other writer — saveUserProfile() in CodeMV.gs — is the patient PORTAL's
// own editor: it takes no session token at all, keys on a `username` the
// caller supplies, and touches eight columns. On an anonymous web app that
// is an unauthenticated write to any patient's record; it is not something
// to route staff edits through.
//
// So a mistyped mobile number, a misspelt name, a wrong date of birth, a
// blood group entered in the wrong box — all of them were permanent, and
// the workaround is a second registration, which is how a clinic ends up
// with the same patient under two IDs and half their history under each.
//
// WHAT THIS DOES
//
// One guarded write for the fields a receptionist can legitimately correct,
// with a before/after diff on the audit log. What it deliberately does NOT
// touch:
//
//   Column A (Patient_ID) — the identity. Printed on barcodes, referenced
//     by every admission, bill, order and note. It is not editable here and
//     should not be editable anywhere.
//   Column B (Password)   — the portal credential. Never read back to the
//     browser by pt_readProfile_, and never written from here either.
//   Column K (Registration_Date) — when they first came. A historical fact.
//
// EVERY CHANGE IS LOGGED WITH ITS OLD VALUE. Demographics are what a
// clinical record is matched on; an unlogged edit to a date of birth is
// indistinguishable from having always been wrong.
// ============================================================================

/**
 * The fields this endpoint will write, and the column each one lives in.
 * Zero-based, matching pt_readProfile_() in CodeMV.gs — the two must agree,
 * so they are written the same way round.
 */
var PPE_FIELDS = {
  name:            { col: 2,  label: "Name" },
  age:             { col: 3,  label: "Age" },
  gender:          { col: 4,  label: "Gender" },
  dob:             { col: 5,  label: "Date of birth" },
  mobile:          { col: 6,  label: "Mobile" },
  whatsapp:        { col: 7,  label: "WhatsApp" },
  address:         { col: 8,  label: "Address" },
  comorb:          { col: 9,  label: "Co-morbidities" },
  salutation:      { col: 11, label: "Salutation" },
  maritalStatus:   { col: 12, label: "Marital status" },
  bloodGroup:      { col: 13, label: "Blood group" },
  occupation:      { col: 14, label: "Occupation" },
  education:       { col: 15, label: "Education" },
  email:           { col: 16, label: "Email" },
  relationType:    { col: 17, label: "Relation type" },
  relationName:    { col: 18, label: "Relative / guardian" },
  emergencyName:   { col: 19, label: "Emergency contact" },
  emergencyNumber: { col: 20, label: "Emergency number" },
  referredBy:      { col: 21, label: "Referred by" }
};

function ppe_str_(v) { return (v === null || v === undefined) ? "" : String(v).trim(); }

/**
 * Validates the fields where a wrong value is worse than a blank one.
 *
 * Returns a list of problems, one sentence each, naming the field. A single
 * "Invalid input" sends the user hunting through nineteen boxes.
 */
function ppe_validate_(patch) {
  var problems = [];

  if (patch.hasOwnProperty("name")) {
    var n = ppe_str_(patch.name);
    if (n.length < 2) problems.push("The name cannot be blank.");
    else if (!/[A-Za-z]/.test(n)) problems.push("The name does not contain any letters.");
  }

  if (patch.hasOwnProperty("age")) {
    var a = ppe_str_(patch.age);
    if (a !== "") {
      var n2 = Number(a);
      // The dose calculator bands paediatric doses on this number, and the
      // one thing that has actually landed in an age box in this system is
      // a ten-digit mobile number.
      if (!isFinite(n2) || n2 < 0 || n2 > 130) {
        problems.push("\"" + a + "\" is not an age. Enter it in whole years.");
      }
    }
  }

  ["mobile", "whatsapp", "emergencyNumber"].forEach(function (k) {
    if (!patch.hasOwnProperty(k)) return;
    var raw = ppe_str_(patch[k]);
    if (!raw) return;
    var digits = raw.replace(/[^0-9]/g, "");
    if (digits.length < 6 || digits.length > 15) {
      problems.push(PPE_FIELDS[k].label + " \"" + raw + "\" is not a usable phone number.");
    }
  });

  if (patch.hasOwnProperty("email")) {
    var e = ppe_str_(patch.email);
    if (e && !/^[^@\s]+@[^@\s]+\.[^@\s]{2,}$/.test(e)) {
      problems.push("\"" + e + "\" is not an email address.");
    }
  }

  if (patch.hasOwnProperty("dob")) {
    var d = ppe_str_(patch.dob);
    if (d) {
      var parsed = (typeof cresc_parseDate_ === "function") ? cresc_parseDate_(d) : new Date(d);
      if (!parsed || isNaN(parsed.getTime())) {
        problems.push("The date of birth \"" + d + "\" could not be read. Use yyyy-MM-dd.");
      } else if (parsed.getTime() > Date.now()) {
        problems.push("The date of birth is in the future. Check the year.");
      }
    }
  }

  if (patch.hasOwnProperty("bloodGroup")) {
    var bg = ppe_str_(patch.bloodGroup).toUpperCase().replace(/\s+/g, "");
    if (bg && !/^(A|B|AB|O)[+-]$/.test(bg) && bg !== "UNKNOWN") {
      problems.push("\"" + ppe_str_(patch.bloodGroup) + "\" is not a blood group. " +
                    "Use A+, A-, B+, B-, AB+, AB-, O+ or O-.");
    }
  }

  return problems;
}

/**
 * FRONTEND ENTRY. Updates a patient's demographics.
 *
 * Only the keys actually present in `patch` are written, so a screen that
 * edits four fields cannot blank the other fifteen by omitting them — which
 * is exactly what a full-row write from a partial form does.
 *
 * @param {string} patientId
 * @param {Object} patch   any subset of PPE_FIELDS
 * @param {string} sessionToken
 * @return {{success:boolean, changed:Array, message:string, profile:Object}}
 */
function updatePatientProfile(patientId, patch, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var actor = crescRequire_(sessionToken, "patient.write");

    var want = ppe_str_(patientId).toUpperCase();
    if (!want) return { success: false, changed: [], message: "No patient was named." };
    patch = patch || {};

    // Unknown keys are refused rather than ignored: silently dropping a
    // field the caller believed it had saved is the worse failure.
    var unknown = Object.keys(patch).filter(function (k) { return !PPE_FIELDS.hasOwnProperty(k); });
    if (unknown.length) {
      return { success: false, changed: [],
               message: "These fields cannot be edited here: " + unknown.join(", ") + "." };
    }
    if (!Object.keys(patch).length) {
      return { success: false, changed: [], message: "Nothing was changed." };
    }

    var problems = ppe_validate_(patch);
    if (problems.length) {
      return { success: false, changed: [], message: problems.join(" ") };
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Patients");
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: false, changed: [], message: "The Patients sheet is empty." };
    }

    var cell = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
      .createTextFinder(want).matchEntireCell(true).matchCase(false).findNext();
    if (!cell) {
      return { success: false, changed: [], message: "No patient with the ID " + patientId + "." };
    }

    var rowNum = cell.getRow();
    var width = Math.max(22, sheet.getLastColumn());
    var before = sheet.getRange(rowNum, 1, 1, width).getValues()[0];

    // Build the write as one contiguous range rather than a setValue per
    // field: nineteen separate writes is nineteen service round trips, and a
    // half-applied row if one of them throws.
    var after = before.slice();
    var changed = [];

    Object.keys(patch).forEach(function (key) {
      var spec = PPE_FIELDS[key];
      var newVal = ppe_str_(patch[key]);

      // Normalise the two values that are matched on rather than read.
      if (key === "bloodGroup") newVal = newVal.toUpperCase().replace(/\s+/g, "");
      if (key === "email") newVal = newVal.toLowerCase();
      if (key === "dob" && newVal) {
        var d = (typeof cresc_parseDate_ === "function") ? cresc_parseDate_(newVal) : new Date(newVal);
        if (d && !isNaN(d.getTime())) {
          newVal = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
      }

      var oldRaw = before[spec.col];
      var oldVal = (oldRaw instanceof Date)
        ? Utilities.formatDate(oldRaw, Session.getScriptTimeZone(), "yyyy-MM-dd")
        : ppe_str_(oldRaw);

      if (oldVal === newVal) return;          // not a change; not logged as one
      after[spec.col] = newVal;
      changed.push({ field: spec.label, from: oldVal, to: newVal });
    });

    if (!changed.length) {
      return { success: true, changed: [], message: "Nothing was different.",
               profile: pt_readProfile_(want) };
    }

    sheet.getRange(rowNum, 1, 1, width).setValues([after]);

    // The old values are the point of the entry. "Profile edited" tells a
    // reviewer nothing; "Mobile 9876543210 -> 9876543211" tells them
    // everything, including whether it was a correction or a mistake.
    try {
      logAudit_(actor && actor.username ? { username: actor.username, role: actor.role } : null,
                "PATIENT_PROFILE_EDITED", "Patient", want,
                { changes: changed });
    } catch (e) { /* audit must not block the correction */ }

    try {
      acc_audit_(actor.username, "PATIENT_PROFILE_EDITED", "Patients", want,
                 JSON.stringify(changed.map(function (c) { return c.field + ": " + c.from; })),
                 JSON.stringify(changed.map(function (c) { return c.field + ": " + c.to; })),
                 "Demographic correction");
    } catch (e) { /* the second log is best effort */ }

    SpreadsheetApp.flush();

    return {
      success: true,
      changed: changed,
      profile: pt_readProfile_(want),
      message: changed.length + " field" + (changed.length === 1 ? "" : "s") + " updated."
    };

  } catch (err) {
    var m = String((err && err.message) || err);
    if (m.indexOf("FORBIDDEN:") === 0) {
      return { success: false, changed: [], message: m.replace("FORBIDDEN: ", "") };
    }
    return { success: false, changed: [], message: "Could not save the profile: " + m };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * The edit history for one patient, newest first.
 *
 * Read from Audit_Log, which is append-only. It answers "was this always the
 * date of birth, or did somebody change it" — which is the question that
 * actually gets asked, usually at the worst possible moment.
 */
function getPatientProfileHistory(patientId, sessionToken) {
  try {
    crescRequire_(sessionToken, "patient.read");
    var want = ppe_str_(patientId).toUpperCase();
    if (!want) return { success: false, rows: [], message: "No patient was named." };

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Audit_Log");
    if (!sh || sh.getLastRow() < 2) return { success: true, rows: [], message: "" };

    var data = sh.getDataRange().getValues();
    var rows = [];
    for (var i = data.length - 1; i >= 1 && rows.length < 25; i--) {
      if (ppe_str_(data[i][6]) !== "PATIENT_PROFILE_EDITED") continue;
      if (ppe_str_(data[i][8]).toUpperCase() !== want) continue;
      var details = {};
      try { details = JSON.parse(data[i][9] || "{}"); } catch (e) { details = {}; }
      rows.push({
        at: (data[i][1] instanceof Date)
              ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), "dd-MMM-yyyy HH:mm")
              : ppe_str_(data[i][1]),
        by: ppe_str_(data[i][3]),
        role: ppe_str_(data[i][4]),
        changes: details.changes || []
      });
    }
    return { success: true, rows: rows, message: "" };
  } catch (err) {
    var m = String((err && err.message) || err);
    return { success: false, rows: [], message: m.replace("FORBIDDEN: ", "") };
  }
}
