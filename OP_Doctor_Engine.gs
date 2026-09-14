// ============================================================================
// OP_Doctor_Engine.gs  —  Crescentia HealthTech
// PHASE 4 : OPD multi-doctor — scoped queue, encounter attribution, referrals
// ----------------------------------------------------------------------------
// REQUIRES: Doctor_Core.gs, Doctor_Schedule_Engine.gs, Doctor_Session_Store.gs
//
// DESIGN
//   A doctor's queue is their OWN patients. Cross-doctor access happens only
//   through an explicit referral, so there is always a record of who asked,
//   who was asked, and why.
//
//   OP_Encounters.Doctor_ID = who actually consulted (drives the signature).
//   OP_Encounters.Booking_Doctor_ID = who the appointment was made with.
//   When they differ the row is flagged SUBSTITUTE rather than silently
//   rewritten — a prescription must never claim a doctor who did not sign it.
//
// This file WRAPS your existing saveOPEncounter(). It does not replace it.
// ============================================================================

// ============================================================================
// SECTION A — SHEETS
// ============================================================================

function op_referralSheet_() {
  // The first sixteen columns are read BY INDEX elsewhere in this file, so
  // everything added since is appended and read by header name only.
  //
  //   Urgency          a routine second opinion and a patient who needs
  //                    seeing now looked identical in the queue. They are
  //                    not the same request and should not sort together.
  //   Clinical_Summary what the referring doctor already knows. Without it
  //                    the receiving doctor opens the chart from scratch,
  //                    and the reason — "please see" — carries nothing.
  //   Referral_Type    INTERNAL, or EXTERNAL for a referral out of the
  //                    clinic. An external referral had nowhere to go at
  //                    all, so it was written on paper and left no record.
  //   External_To /
  //   External_Facility  who and where, for an external referral.
  return dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), "OP_Referrals", [
    "Referral_ID", "Tenant_ID", "Date", "Patient_ID", "Patient_Name",
    "From_Doctor_ID", "From_Doctor_Name", "To_Doctor_ID", "Reason",
    "Source_Encounter_ID", "Status", "Created_By", "Created_At",
    "Responded_At", "Response_Encounter_ID", "Response_Note",
    "Urgency", "Clinical_Summary", "Referral_Type",
    "External_To", "External_Facility", "Specialty_Sought"
  ]);
}

/** Urgencies, worst first — the order a queue should show them in. */
var OP_REFERRAL_URGENCY = ["EMERGENCY", "URGENT", "ROUTINE"];

/** Ensures the attribution columns exist on OP_Encounters. Idempotent. */
function op_ensureEncounterColumns_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("OP_Encounters");
  if (!sh) return null;
  ["Doctor_ID", "Doctor_Signature_Snapshot", "Booking_Doctor_ID",
   "Referral_ID", "Attribution_Source"].forEach(function (h) {
    dc_ensureColumn_(sh, h);
  });
  return sh;
}

// ============================================================================
// SECTION B — SCOPED OP QUEUE
// ============================================================================

/**
 * FRONTEND ENTRY. The consulting doctor's working queue for a date.
 * Returns two lists, deliberately separate:
 *   own       — appointments booked with this doctor
 *   referred  — patients referred TO this doctor by a colleague
 *
 * Reception/admin (scope ALL) with no doctorId gets the whole clinic.
 */
function getMyOPQueue(dateStr, doctorId, sessionToken) {
  try {
    var scope = resolveScope_(sessionToken, doctorId);
    if (!scope.ok) {
      return { success: false, message: scope.message, own: [], referred: [] };
    }

    var dateKey = dc_dateKey_(dateStr);
    if (!dateKey) return { success: false, message: "Select a valid date.", own: [], referred: [] };

    var target = dc_str_(doctorId) || scope.selfDoctorId;

    // ---- own appointments (reuses the scoped ledger) ---------------------
    var ledger = fetchDailyLedgerScoped(dateKey, target, sessionToken);
    if (!ledger.success) {
      return { success: false, message: ledger.message, own: [], referred: [] };
    }

    var LIVE = ["Booked", "Arrived", "In-Progress"];
    var own = ledger.ledger.filter(function (r) {
      return LIVE.indexOf(r.status) !== -1 || r.status === "Completed";
    }).map(function (r) {
      return {
        apptId: r.apptId, time12: r.time12, time24: r.time24,
        patientId: r.patientId, patientName: r.patientName,
        age: r.age, sex: r.sex, purpose: r.purpose, status: r.status,
        doctorId: r.doctorId, doctorName: r.doctorName,
        kind: "OWN"
      };
    });

    // ---- referrals addressed to this doctor -----------------------------
    var referred = [];
    if (target) {
      var sh = op_referralSheet_();
      var data = sh.getDataRange().getDisplayValues();
      var docU = dc_upper_(target);

      for (var i = 1; i < data.length; i++) {
        if (dc_upper_(data[i][7]) !== docU) continue;            // To_Doctor_ID, cheap
        var status = dc_upper_(data[i][10]);
        if (status !== "PENDING" && status !== "ACCEPTED") continue;
        if (dc_dateKey_(data[i][2]) !== dateKey) continue;

        referred.push({
          referralId:   dc_str_(data[i][0]),
          patientId:    dc_str_(data[i][3]),
          patientName:  dc_str_(data[i][4]),
          fromDoctorId: dc_str_(data[i][5]),
          fromDoctorName: dc_str_(data[i][6]),
          reason:       dc_str_(data[i][8]),
          sourceEncounterId: dc_str_(data[i][9]),
          status:       status,
          kind:         "REFERRAL"
        });
      }
    }

    // ---- referrals this doctor SENT that are still open ------------------
    var sentOpen = 0;
    if (target) {
      var d2 = op_referralSheet_().getDataRange().getDisplayValues();
      for (var j = 1; j < d2.length; j++) {
        if (dc_upper_(d2[j][5]) !== dc_upper_(target)) continue;
        if (dc_upper_(d2[j][10]) !== "PENDING") continue;
        if (dc_dateKey_(d2[j][2]) !== dateKey) continue;
        sentOpen++;
      }
    }

    own.sort(function (a, b) { return dc_minutes_(a.time24) - dc_minutes_(b.time24); });

    return {
      success: true,
      date: dateKey,
      doctorId: target,
      scopeMode: scope.mode,
      own: own,
      referred: referred,
      sentReferralsOpen: sentOpen,
      message: (own.length || referred.length) ? "" :
               "No patients in your queue for this date."
    };
  } catch (e) {
    return { success: false, message: "Could not load the queue: " + e.message,
             own: [], referred: [] };
  }
}

// ============================================================================
// SECTION C — ATTRIBUTED ENCOUNTER SAVE
// ============================================================================

/**
 * FRONTEND ENTRY. Replaces the direct saveOPEncounter() call from the UI.
 * Runs the existing save, then stamps who actually consulted.
 *
 * payload = <your existing OP payload> plus optional:
 *           doctorId    — only reception/admin may name someone else
 *           referralId  — set when consulting off a colleague's referral
 */
function saveOPEncounterScoped(payload, sessionToken) {
  try {
    var w = resolveWriteDoctor_(sessionToken, payload && payload.doctorId);
    if (!w.ok) return { success: false, message: w.message };

    var patientId = dc_upper_(payload && payload.patientId);
    if (!patientId || patientId === "--") {
      return { success: false, message: "Fetch a patient before saving." };
    }

    // ---- referral gate ---------------------------------------------------
    var referralId = dc_str_(payload.referralId);
    var referralRow = -1;
    if (referralId) {
      var rSh = op_referralSheet_();
      var rData = rSh.getDataRange().getDisplayValues();
      var found = false;
      for (var i = 1; i < rData.length; i++) {
        if (dc_upper_(rData[i][0]) !== dc_upper_(referralId)) continue;
        found = true;
        if (dc_upper_(rData[i][7]) !== dc_upper_(w.doctorId)) {
          return { success: false, message: "This referral is addressed to another doctor." };
        }
        if (dc_upper_(rData[i][3]) !== patientId) {
          return { success: false, message: "This referral is for a different patient." };
        }
        referralRow = i + 1;
        break;
      }
      if (!found) return { success: false, message: "Referral not found." };
    }

    // ---- who was this booked with? --------------------------------------
    var bookingDoctorId = op_findBookingDoctor_(patientId, new Date());

    // ---- run the existing save engine unchanged --------------------------
    var res = saveOPEncounter(payload);
    if (!res || !res.success) return res;

    // ---- stamp attribution ----------------------------------------------
    var sh = op_ensureEncounterColumns_();
    if (sh) {
      var m = dc_headerMap_(sh);
      var data = sh.getDataRange().getDisplayValues();
      var rowIdx = -1;
      for (var k = data.length - 1; k >= 1; k--) {
        if (dc_str_(data[k][0]) === dc_str_(res.encounterId)) { rowIdx = k + 1; break; }
      }

      if (rowIdx > -1) {
        var source = referralId ? "REFERRAL"
                   : (bookingDoctorId && dc_upper_(bookingDoctorId) !== dc_upper_(w.doctorId)
                        ? "SUBSTITUTE" : "RECORDED");

        // Signature is SNAPSHOTTED here, never re-derived at print time.
        if (m["Doctor_ID"] !== undefined)
          sh.getRange(rowIdx, m["Doctor_ID"] + 1).setValue(String(w.doctorId));
        if (m["Doctor_Signature_Snapshot"] !== undefined)
          sh.getRange(rowIdx, m["Doctor_Signature_Snapshot"] + 1).setValue(String(w.signature));
        if (m["Booking_Doctor_ID"] !== undefined)
          sh.getRange(rowIdx, m["Booking_Doctor_ID"] + 1).setValue(String(bookingDoctorId || ""));
        if (m["Referral_ID"] !== undefined)
          sh.getRange(rowIdx, m["Referral_ID"] + 1).setValue(String(referralId || ""));
        if (m["Attribution_Source"] !== undefined)
          sh.getRange(rowIdx, m["Attribution_Source"] + 1).setValue(source);

        res.attributionSource = source;
        if (source === "SUBSTITUTE") {
          res.message = res.message +
            " Recorded under " + w.name + " (booked with another doctor).";
        }
      }
    }

    // ---- close the referral ---------------------------------------------
    if (referralRow > -1) {
      var rs = op_referralSheet_();
      var rm = dc_headerMap_(rs);
      rs.getRange(referralRow, rm["Status"] + 1).setValue("COMPLETED");
      rs.getRange(referralRow, rm["Responded_At"] + 1).setValue(new Date());
      rs.getRange(referralRow, rm["Response_Encounter_ID"] + 1).setValue(String(res.encounterId));
      if (payload.referralNote) {
        rs.getRange(referralRow, rm["Response_Note"] + 1).setValue(String(payload.referralNote));
      }
    }

    SpreadsheetApp.flush();
    logAudit_(w.sess, "OP_ENCOUNTER_SAVE", "OP_Encounter", res.encounterId, {
      doctorId: w.doctorId, patientId: patientId,
      bookingDoctorId: bookingDoctorId, referralId: referralId
    });

    res.doctorId = w.doctorId;
    res.doctorName = w.name;
    return res;

  } catch (e) {
    return { success: false, message: "Could not save the consultation: " + e.message };
  }
}

/** Doctor_ID on this patient's appointment for the given date, or "". */
function op_findBookingDoctor_(patientId, dateObj) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Appointments");
    if (!sh || sh.getLastRow() < 2) return "";

    var m = dc_headerMap_(sh);
    var docCol = (m["Doctor_ID"] === undefined) ? -1 : m["Doctor_ID"];
    if (docCol === -1) return "";

    var dateKey = dc_fmtDate_(dateObj);
    var data = sh.getDataRange().getDisplayValues();
    var pid = dc_upper_(patientId);

    for (var i = data.length - 1; i >= 1; i--) {
      // Cheap comparisons first — this sheet is large.
      if (dc_upper_(data[i][1]) !== pid) continue;
      var status = data[i][6];
      if (status === "Blocked" || status === "DELETE" || status === "Cancelled") continue;
      if (dc_dateKey_(data[i][3]) !== dateKey) continue;
      return dc_str_(data[i][docCol]);
    }
    return "";
  } catch (e) { return ""; }
}

// ============================================================================
// SECTION D — INTERNAL REFERRAL / CROSS-CONSULT
// ============================================================================

/**
 * FRONTEND ENTRY. Dr. A asks Dr. B to see a patient.
 * payload = { patientId, patientName, toDoctorId, reason, sourceEncounterId, date }
 */
function createOPReferral(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var w = resolveWriteDoctor_(sessionToken, payload && payload.fromDoctorId);
    if (!w.ok) return { success: false, message: w.message };

    var patientId = dc_upper_(payload.patientId);
    var toId = dc_str_(payload.toDoctorId);
    var reason = dc_str_(payload.reason);
    var kind = dc_upper_(payload.referralType) === "EXTERNAL" ? "EXTERNAL" : "INTERNAL";
    var urgency = dc_upper_(payload.urgency);
    if (OP_REFERRAL_URGENCY.indexOf(urgency) === -1) urgency = "ROUTINE";

    if (!patientId) return { success: false, message: "Patient ID is required." };
    if (!reason) {
      return { success: false, message: "State the reason for referral — it becomes part of the record." };
    }
    // A one-word reason is the reason nobody can act on. "Please see" is
    // what the receiving doctor gets instead of a question.
    if (reason.replace(/\s+/g, " ").length < 10) {
      return { success: false,
               message: "Say what you want their opinion on, in a sentence. " +
                        "\"" + reason + "\" does not give them a question to answer." };
    }

    var toDoc = null, externalTo = "", externalFacility = "";

    if (kind === "EXTERNAL") {
      // Referring OUT had nowhere to go at all, so it was written on paper
      // and left no record: nothing in the chart said the patient had been
      // sent anywhere, and nothing chased whether they went.
      externalTo = dc_str_(payload.externalTo);
      externalFacility = dc_str_(payload.externalFacility);
      if (!externalTo && !externalFacility) {
        return { success: false,
                 message: "Name the doctor or the hospital you are referring to." };
      }
    } else {
      if (!toId) return { success: false, message: "Select the doctor you are referring to." };
      if (dc_upper_(toId) === dc_upper_(w.doctorId)) {
        return { success: false, message: "You cannot refer a patient to yourself." };
      }
      toDoc = dc_getDoctorById_(toId);
      if (!toDoc) return { success: false, message: "Doctor not found." };
      if (toDoc.status !== "ACTIVE") {
        return { success: false, message: toDoc.name + " is not currently active." };
      }
    }

    var dateKey = dc_dateKey_(payload.date) || dc_fmtDate_(new Date());

    // Don't stack duplicate open referrals for the same pair on the same day.
    var sh = op_referralSheet_();
    var m = dc_headerMap_(sh);
    var data = sh.getDataRange().getDisplayValues();
    if (kind === "INTERNAL") {
      for (var i = 1; i < data.length; i++) {
        if (dc_upper_(data[i][3]) !== patientId) continue;
        if (dc_upper_(data[i][7]) !== dc_upper_(toId)) continue;
        var st = dc_upper_(data[i][10]);
        if (st !== "PENDING" && st !== "ACCEPTED") continue;
        if (dc_dateKey_(data[i][2]) !== dateKey) continue;
        return { success: false,
                 message: "An open referral to " + toDoc.name + " already exists for this patient today." };
      }
    }

    var refId = "REF-" + Utilities.getUuid().substring(0, 8).toUpperCase();

    // Written by header name from column 17 on, because those columns were
    // appended to a sheet whose first sixteen are read by index.
    var row = [
      String(refId),
      String(getTenantId_()),
      String(dateKey),
      String(patientId),
      String(dc_str_(payload.patientName)),
      String(w.doctorId),
      String(w.name),
      String(toDoc ? toDoc.doctorId : ""),
      String(reason),
      String(dc_str_(payload.sourceEncounterId)),
      "PENDING",
      String(w.sess.username),
      new Date(),
      "", "", ""
    ];
    row[m["Urgency"]]          = urgency;
    row[m["Clinical_Summary"]] = dc_str_(payload.clinicalSummary);
    row[m["Referral_Type"]]    = kind;
    row[m["External_To"]]      = externalTo;
    row[m["External_Facility"]]= externalFacility;
    row[m["Specialty_Sought"]] = dc_str_(payload.specialty);
    for (var c = 0; c < sh.getLastColumn(); c++) if (row[c] === undefined) row[c] = "";

    sh.appendRow(row);
    dc_invalidate_("OP_Referrals");

    SpreadsheetApp.flush();
    logAudit_(w.sess, "OP_REFERRAL_CREATE", "Patient", patientId, {
      from: w.doctorId, to: toDoc ? toDoc.doctorId : (externalTo || externalFacility),
      type: kind, urgency: urgency, reason: reason
    });

    var who = toDoc ? toDoc.name : (externalTo || externalFacility);
    return { success: true, referralId: refId, referralType: kind, urgency: urgency,
             message: kind === "EXTERNAL"
               ? "Referral to " + who + " recorded. Print the letter for the patient to carry."
               : "Referred to " + who + "." +
                 (urgency !== "ROUTINE" ? " Marked " + urgency.toLowerCase() + "." : "") };

  } catch (e) {
    return { success: false, message: "Could not create the referral: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/** FRONTEND ENTRY. The receiving doctor accepts or declines. */
function respondToOPReferral(referralId, decision, note, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var sess = dc_validateSession_(sessionToken);
    if (!sess) return { success: false, message: "Your session has expired." };

    var verdict = dc_upper_(decision);
    if (["ACCEPTED", "DECLINED"].indexOf(verdict) === -1) {
      return { success: false, message: "Invalid response." };
    }
    if (verdict === "DECLINED" && !dc_str_(note)) {
      return { success: false, message: "Give a reason when declining a referral." };
    }

    var sh = op_referralSheet_();
    var m = dc_headerMap_(sh);
    var data = sh.getDataRange().getDisplayValues();
    var target = dc_upper_(referralId);

    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][0]) !== target) continue;

      var w = resolveWriteDoctor_(sessionToken, dc_str_(data[i][7]));
      if (!w.ok) return { success: false, message: w.message };
      if (dc_upper_(w.doctorId) !== dc_upper_(data[i][7])) {
        return { success: false, message: "Only the receiving doctor can respond." };
      }
      if (dc_upper_(data[i][10]) !== "PENDING") {
        return { success: false, message: "This referral has already been actioned." };
      }

      sh.getRange(i + 1, m["Status"] + 1).setValue(verdict);
      sh.getRange(i + 1, m["Responded_At"] + 1).setValue(new Date());
      if (note) sh.getRange(i + 1, m["Response_Note"] + 1).setValue(String(note));

      SpreadsheetApp.flush();
      logAudit_(sess, "OP_REFERRAL_" + verdict, "Patient", dc_str_(data[i][3]),
                { referralId: referralId });

      return { success: true,
               message: verdict === "ACCEPTED"
                 ? "Referral accepted. The patient is now in your queue."
                 : "Referral declined." };
    }
    return { success: false, message: "Referral not found." };

  } catch (e) {
    return { success: false, message: "Could not update the referral: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * FRONTEND ENTRY. Referral history for one patient — both directions.
 * Shown in the EMR timeline so the trail is visible where clinicians look.
 */
function getPatientReferralHistory(patientId, sessionToken) {
  try {
    var scope = resolveScope_(sessionToken, null);
    if (!scope.ok) return { success: false, message: scope.message, rows: [] };

    var sh = op_referralSheet_();
    var data = sh.getDataRange().getDisplayValues();
    var pid = dc_upper_(patientId);
    var rows = [];

    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][3]) !== pid) continue;
      rows.push({
        referralId: dc_str_(data[i][0]),
        date: dc_dateKey_(data[i][2]),
        fromDoctorId: dc_str_(data[i][5]),
        fromDoctorName: dc_str_(data[i][6]),
        toDoctorId: dc_str_(data[i][7]),
        toDoctorName: (dc_getDoctorById_(data[i][7]) || {}).name || dc_str_(data[i][7]),
        reason: dc_str_(data[i][8]),
        status: dc_upper_(data[i][10]),
        responseNote: dc_str_(data[i][15]),
        responseEncounterId: dc_str_(data[i][14])
      });
    }
    rows.sort(function (a, b) { return a.date < b.date ? 1 : -1; });

    return { success: true, rows: rows,
             message: rows.length ? "" : "No referrals recorded for this patient." };
  } catch (e) {
    return { success: false, message: "Could not load referrals: " + e.message, rows: [] };
  }
}
// ============================================================================
// SECTION E — REFERRALS A DOCTOR CAN ACTUALLY FOLLOW
// ----------------------------------------------------------------------------
// getOPDoctorQueue() reports referrals ADDRESSED TO a doctor in full, and
// referrals they SENT as a single number: `sentReferralsOpen: 3`. Three what?
// To whom, for which patient, asked when, and did anybody answer?
//
// A referral is a question. Sending one and never seeing the reply is the
// same as not asking, and the commonest failure of a referral system is not
// that the request is lost — it is that nobody notices it was never answered.
// ============================================================================

/**
 * FRONTEND ENTRY. Referrals this doctor sent, newest first, with what
 * happened to each.
 *
 * @param {string} sessionToken
 * @param {{status:string, limit:number}} [opts]
 */
function listMyOPReferrals(sessionToken, opts) {
  try {
    opts = opts || {};
    var w = resolveWriteDoctor_(sessionToken, opts.doctorId);
    if (!w.ok) return { success: false, rows: [], message: w.message };

    var sh = op_referralSheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    if (!data || data.length < 2) return { success: true, rows: [], message: "" };

    function g(row, name) { return (m[name] === undefined) ? "" : dc_str_(row[m[name]]); }

    var want = dc_upper_(opts.status);
    var rows = [];
    for (var i = data.length - 1; i >= 1 && rows.length < (opts.limit || 60); i--) {
      if (dc_upper_(data[i][5]) !== dc_upper_(w.doctorId)) continue;   // From_Doctor_ID
      var status = dc_upper_(data[i][10]) || "PENDING";
      if (want && want !== "ALL" && status !== want) continue;

      var toId = dc_str_(data[i][7]);
      var toDoc = toId ? dc_getDoctorById_(toId) : null;

      rows.push({
        referralId: dc_str_(data[i][0]),
        date:       dc_str_(data[i][2]),
        patientId:  dc_str_(data[i][3]),
        patientName:dc_str_(data[i][4]),
        toDoctorId: toId,
        toName:     toDoc ? toDoc.name
                          : (g(data[i], "External_To") || g(data[i], "External_Facility")),
        toSpecialty: toDoc ? toDoc.specialty : g(data[i], "Specialty_Sought"),
        external:   dc_upper_(g(data[i], "Referral_Type")) === "EXTERNAL",
        facility:   g(data[i], "External_Facility"),
        urgency:    dc_upper_(g(data[i], "Urgency")) || "ROUTINE",
        reason:     dc_str_(data[i][8]),
        summary:    g(data[i], "Clinical_Summary"),
        status:     status,
        respondedAt:dc_str_(data[i][13]),
        responseNote: dc_str_(data[i][15]),
        // The number that makes an unanswered referral visible. A question
        // asked eleven days ago and never answered is the failure mode this
        // whole listing exists for.
        daysOpen: (function () {
          if (status !== "PENDING" && status !== "ACCEPTED") return 0;
          var d = dc_dateKey_(data[i][2]);
          if (!d) return 0;
          var then = (typeof cresc_parseDate_ === "function") ? cresc_parseDate_(d) : new Date(d);
          if (!then || isNaN(then.getTime())) return 0;
          return Math.max(0, Math.floor((Date.now() - then.getTime()) / 86400000));
        })()
      });
    }

    // Emergency before urgent before routine, and within each, oldest first:
    // the one that has been waiting longest is the one to chase.
    rows.sort(function (a, b) {
      var ua = OP_REFERRAL_URGENCY.indexOf(a.urgency);
      var ub = OP_REFERRAL_URGENCY.indexOf(b.urgency);
      if (ua !== ub) return ua - ub;
      return b.daysOpen - a.daysOpen;
    });

    var openCount = rows.filter(function (r) {
      return r.status === "PENDING" || r.status === "ACCEPTED";
    }).length;
    var stale = rows.filter(function (r) { return r.daysOpen >= 3; }).length;

    return { success: true, rows: rows, openCount: openCount, staleCount: stale,
             message: "" };
  } catch (e) {
    return { success: false, rows: [], message: "Referrals unavailable: " + e.message };
  }
}

/**
 * FRONTEND ENTRY. The referral as a letter the patient can carry.
 *
 * An external referral used to be written on a pad, so nothing in the record
 * said where the patient had been sent and the receiving doctor got whatever
 * handwriting fitted on the page. This is the same document, generated from
 * the referral that was actually recorded.
 *
 * Rendered through the shared letterhead in Clinic_Profile.gs so it carries
 * the same name, address and registration as every other document the clinic
 * hands out.
 */
function getOPReferralLetter(referralId, sessionToken) {
  try {
    var sess = dc_validateSession_(sessionToken);
    if (!sess) return { success: false, message: "Your session has expired." };

    var sh = op_referralSheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    var want = dc_upper_(referralId);
    var row = null;
    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][0]) === want) { row = data[i]; break; }
    }
    if (!row) return { success: false, message: "Referral " + referralId + " not found." };

    function g(name) { return (m[name] === undefined) ? "" : dc_str_(row[m[name]]); }
    function esc(v) {
      return String(v == null ? "" : v)
        .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    }

    var clinic = (typeof cresc_clinic_ === "function")
      ? cresc_clinic_() : { name: "Crescentia HealthTech", address: "", phone: "", gstin: "", footer: "" };

    var toDoc = dc_str_(row[7]) ? dc_getDoctorById_(dc_str_(row[7])) : null;
    var toName = toDoc ? toDoc.name : (g("External_To") || g("External_Facility") || "The treating doctor");
    var toLine = toDoc
      ? esc(toDoc.name) + (toDoc.specialty ? ", " + esc(toDoc.specialty) : "")
      : esc(g("External_To")) + (g("External_Facility") ? "<br>" + esc(g("External_Facility")) : "");

    var patient = null;
    try { patient = pt_readProfile_(dc_str_(row[3])); } catch (e) { patient = null; }

    var urgency = dc_upper_(g("Urgency")) || "ROUTINE";
    var fromDoc = dc_str_(row[5]) ? dc_getDoctorById_(dc_str_(row[5])) : null;
    var signature = fromDoc ? (fromDoc.signature || fromDoc.name) : dc_str_(row[6]);

    var html =
      '<div style="max-width:760px;margin:16px auto;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;' +
        'color:#111827;background:#fff;border:1px solid #e5e7eb;border-radius:10px;overflow:hidden;">' +

      '<div style="padding:18px 24px;border-bottom:2px solid #0369a1;display:flex;' +
        'justify-content:space-between;align-items:flex-start;gap:16px;">' +
        '<div><div style="font-size:20px;font-weight:800;color:#0369a1;text-transform:uppercase;">' +
          esc(clinic.name) + '</div>' +
          (clinic.address ? '<div style="font-size:12px;color:#6b7280;margin-top:4px;">' + esc(clinic.address) + '</div>' : '') +
          (clinic.phone ? '<div style="font-size:12px;color:#6b7280;">Phone: ' + esc(clinic.phone) + '</div>' : '') +
        '</div>' +
        '<div style="text-align:right;">' +
          '<div style="font-size:16px;font-weight:800;letter-spacing:1px;">REFERRAL LETTER</div>' +
          '<div style="font-size:12px;color:#6b7280;margin-top:4px;">Ref: <strong>' + esc(row[0]) + '</strong></div>' +
          '<div style="font-size:11px;color:#6b7280;">Date: ' + esc(row[2]) + '</div>' +
          (urgency !== "ROUTINE"
            ? '<div style="margin-top:6px;display:inline-block;border:2px solid #b91c1c;color:#b91c1c;' +
              'padding:2px 10px;border-radius:4px;font-weight:800;font-size:11px;">' + esc(urgency) + '</div>'
            : '') +
        '</div>' +
      '</div>' +

      '<div style="padding:16px 24px;font-size:13px;line-height:1.6;">' +
        '<p style="margin:0 0 14px;"><strong>To:</strong><br>' + toLine + '</p>' +

        '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:14px;">' +
          '<tr><td style="padding:3px 0;color:#6b7280;width:130px;">Patient</td>' +
            '<td style="padding:3px 0;"><strong>' + esc(row[4]) + '</strong> &bull; ' + esc(row[3]) + '</td></tr>' +
          (patient
            ? '<tr><td style="padding:3px 0;color:#6b7280;">Age / Sex</td><td style="padding:3px 0;">' +
                esc(patient.age) + ' / ' + esc(patient.gender) + '</td></tr>' +
              (patient.mobile ? '<tr><td style="padding:3px 0;color:#6b7280;">Contact</td><td style="padding:3px 0;">' +
                esc(patient.mobile) + '</td></tr>' : '') +
              (patient.comorb ? '<tr><td style="padding:3px 0;color:#6b7280;">Co-morbidities</td><td style="padding:3px 0;">' +
                esc(patient.comorb) + '</td></tr>' : '')
            : '') +
          (g("Specialty_Sought") ? '<tr><td style="padding:3px 0;color:#6b7280;">Opinion sought</td>' +
            '<td style="padding:3px 0;">' + esc(g("Specialty_Sought")) + '</td></tr>' : '') +
        '</table>' +

        '<div style="margin-bottom:14px;">' +
          '<div style="font-size:11px;color:#6b7280;text-transform:uppercase;font-weight:700;margin-bottom:4px;">Reason for referral</div>' +
          '<div style="white-space:pre-wrap;">' + esc(row[8]) + '</div>' +
        '</div>' +

        (g("Clinical_Summary")
          ? '<div style="margin-bottom:14px;">' +
              '<div style="font-size:11px;color:#6b7280;text-transform:uppercase;font-weight:700;margin-bottom:4px;">Clinical summary</div>' +
              '<div style="white-space:pre-wrap;">' + esc(g("Clinical_Summary")) + '</div>' +
            '</div>'
          : '') +

        '<p style="margin:18px 0 0;">Thank you for seeing this patient. ' +
          'I should be grateful for your opinion and for a note of your management.</p>' +

        '<div style="margin-top:34px;">' +
          '<div style="border-top:1px solid #111827;display:inline-block;padding-top:4px;min-width:240px;">' +
            '<strong>' + esc(signature) + '</strong>' +
          '</div>' +
        '</div>' +
      '</div>' +

      '<div style="padding:12px 24px;background:#f9fafb;border-top:1px solid #e5e7eb;' +
        'font-size:11px;color:#6b7280;">' + esc(clinic.footer || "") + '</div>' +
      '</div>';

    return { success: true, html: html, referralId: dc_str_(row[0]), toName: toName };

  } catch (e) {
    return { success: false, message: "Could not build the letter: " + e.message };
  }
}
