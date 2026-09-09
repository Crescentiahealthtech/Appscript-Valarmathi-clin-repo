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
  return dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), "OP_Referrals", [
    "Referral_ID", "Tenant_ID", "Date", "Patient_ID", "Patient_Name",
    "From_Doctor_ID", "From_Doctor_Name", "To_Doctor_ID", "Reason",
    "Source_Encounter_ID", "Status", "Created_By", "Created_At",
    "Responded_At", "Response_Encounter_ID", "Response_Note"
  ]);
}

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

    if (!patientId) return { success: false, message: "Patient ID is required." };
    if (!toId)      return { success: false, message: "Select the doctor you are referring to." };
    if (dc_upper_(toId) === dc_upper_(w.doctorId)) {
      return { success: false, message: "You cannot refer a patient to yourself." };
    }
    if (!reason) {
      return { success: false, message: "State the reason for referral — it becomes part of the record." };
    }

    var toDoc = dc_getDoctorById_(toId);
    if (!toDoc) return { success: false, message: "Doctor not found." };
    if (toDoc.status !== "ACTIVE") {
      return { success: false, message: toDoc.name + " is not currently active." };
    }

    var dateKey = dc_dateKey_(payload.date) || dc_fmtDate_(new Date());

    // Don't stack duplicate open referrals for the same pair on the same day.
    var sh = op_referralSheet_();
    var data = sh.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][3]) !== patientId) continue;
      if (dc_upper_(data[i][7]) !== dc_upper_(toId)) continue;
      var st = dc_upper_(data[i][10]);
      if (st !== "PENDING" && st !== "ACCEPTED") continue;
      if (dc_dateKey_(data[i][2]) !== dateKey) continue;
      return { success: false,
               message: "An open referral to " + toDoc.name + " already exists for this patient today." };
    }

    var refId = "REF-" + Utilities.getUuid().substring(0, 8).toUpperCase();
    sh.appendRow([
      String(refId),
      String(getTenantId_()),
      String(dateKey),
      String(patientId),
      String(dc_str_(payload.patientName)),
      String(w.doctorId),
      String(w.name),
      String(toDoc.doctorId),
      String(reason),
      String(dc_str_(payload.sourceEncounterId)),
      "PENDING",
      String(w.sess.username),
      new Date(),
      "", "", ""
    ]);

    SpreadsheetApp.flush();
    logAudit_(w.sess, "OP_REFERRAL_CREATE", "Patient", patientId, {
      from: w.doctorId, to: toDoc.doctorId, reason: reason
    });

    return { success: true, referralId: refId,
             message: "Referred to " + toDoc.name + "." };

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