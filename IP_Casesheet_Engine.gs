// ============================================================================
// IP_Casesheet_Engine.gs  —  Crescentia HealthTech
// Full rewrite of the IP admission casesheet backend, at OP-module parity.
// Replaces IP_Casesheet_Logic.gs.
// ----------------------------------------------------------------------------
// DESIGN NOTE — why this file is short.
//
// The previous IP casesheet carried its own drug master, lab master and
// template reader, each a weaker copy of the OP equivalent: the lab list came
// from a flat Lab_Test_Master scan with no catalog IDs, so IP lab orders could
// not be billed or resulted like OP ones. Those duplicates are gone.
//
// The OP engines are not OP-specific — they key on doctorId and patientId, not
// on the module that called them. This file reuses them directly:
//
//   getComposerContext / getScopedTemplates  -> phrases, bundles, exam defaults
//   getAssistForDiagnosis                    -> diagnosis-driven assist
//   checkRxSafety / getPatientAllergies      -> interaction + allergy checks
//   listConsultTemplates / applyConsultTemplate
//   buildTaperPlan / suggestPaediatricDose
//   getOPDOrderableTests / createLabRequest   -> the real lab catalog
//   fetchOPDrugMaster / fetchUniversalDrugs  -> formulary
//
// What is genuinely IP-specific lives here: the admission context, the
// care-team write gate (Phase 5), the ward round's baseline row, and routing
// admission orders to IP_Pharmacy_Queue rather than the OP queue.
//
// DEPENDS ON Doctor_Core.gs, IP_Clinical_Access.gs, OP_Rx_Engine.gs,
//            OP_Templates_Engine.gs, OP_Database_Engine.gs, Lab_OPD_Bridge.gs
// ============================================================================

// ---------------------------------------------------------------------------
// SECTION A — WARD ROSTER
// ---------------------------------------------------------------------------

/**
 * Active admissions this session may open, for the casesheet ward picker.
 * Doctor-scoped: a consultant sees their own beds and their cross-consults.
 */
function getIPCasesheetWard(sessionToken) {
  try {
    var gate = resolveIPRead_(sessionToken, null);
    if (!gate.ok) return { success: false, message: gate.message, data: [] };

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("IP_Admissions");
    if (!sh) return { success: true, data: [] };

    var data = sh.getDataRange().getDisplayValues();
    if (data.length <= 1) return { success: true, data: [] };

    var canSee = ipc_wardVisibilityFilter_(gate.scope);
    var out = [];
    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][11]) !== "ACTIVE") continue;
      if (!canSee(data[i][0])) continue;

      var ageSex = dc_str_(data[i][3]);
      var parts = ageSex.split("/");
      out.push({
        ipNumber:   dc_str_(data[i][0]),
        patientId:  dc_str_(data[i][1]),
        name:       dc_str_(data[i][2]),
        ageSex:     ageSex,
        age:        dc_str_(parts[0]).replace(/[^0-9]/g, ""),
        sex:        dc_str_(parts[1]) || "",
        doa:        dc_dateKey_(data[i][4]),
        ward:       dc_str_(data[i][7]),
        bed:        dc_str_(data[i][8]),
        consultant: dc_str_(data[i][9]),
        diagnosis:  dc_str_(data[i][10])
      });
    }
    return { success: true, data: out.reverse(), scopeMode: gate.scope.mode };
  } catch (e) {
    return { success: false, message: "Ward unavailable: " + e.message, data: [] };
  }
}

// ---------------------------------------------------------------------------
// SECTION B — ONE-SHOT COMPOSER CONTEXT
// ---------------------------------------------------------------------------

/**
 * Everything the casesheet needs to open on a patient, in ONE round trip.
 *
 * The old module made seven separate google.script.run calls on load — drug
 * master, universal drugs, lab master, templates, patient, allergies, ward —
 * each a cold Apps Script invocation. On a ward tablet that was most of the
 * time-to-first-keystroke.
 */
function getIPCasesheetContext(ipNumber, sessionToken) {
  try {
    var gate = resolveIPRead_(sessionToken, ipNumber);
    if (!gate.ok) return { success: false, message: gate.message };

    var adm = ipc_admissionRow_(ipNumber);
    if (!adm) return { success: false, message: "Admission " + ipNumber + " was not found." };

    var sess = dc_validateSession_(sessionToken);
    var role = dc_str_(sess && sess.role).toLowerCase();
    var selfDoctorId = dc_str_(sess && sess.doctorId);

    // ---- who is writing, and may they? ----------------------------------
    var doctorName = dc_sessionName_(sess);
    var signature = "";
    if (role === "doctor" && selfDoctorId) {
      var prof = dc_getDoctorById_(selfDoctorId);
      if (prof) { doctorName = prof.name || doctorName; signature = prof.signature; }
    }
    var canWrite = (role === "doctor") && ipc_mayWriteOnAdmission_(ipNumber, selfDoctorId);

    var ageSex = dc_str_(adm.row[3]);
    var sexParts = ageSex.split("/");
    var patientId = dc_str_(adm.row[1]);

    var ctx = {
      success: true,
      canWrite: canWrite,
      role: role,
      doctorId: selfDoctorId,
      doctorName: doctorName,
      authorLabel: ipc_authorLabel_(role, doctorName),
      signature: signature,
      patient: {
        ipNumber:  dc_str_(adm.row[0]),
        patientId: patientId,
        name:      dc_str_(adm.row[2]),
        ageSex:    ageSex,
        age:       dc_str_(sexParts[0]).replace(/[^0-9]/g, ""),
        sex:       dc_str_(sexParts[1]) || "",
        doa:       dc_dateKey_(adm.row[4]),
        ward:      dc_str_(adm.row[7]),
        bed:       dc_str_(adm.row[8]),
        consultant: dc_str_(adm.row[9]),
        diagnosis: dc_str_(adm.row[10])
      }
    };

    // ---- reused OP engines, each failing soft ---------------------------
    ctx.composer   = ipc_safe_(function () { return getComposerContext(selfDoctorId, sessionToken); },
                               { phrases: { CC: [], HX: [], ADVICE: [] }, bundles: [], workup: {} });
    ctx.drugs      = ipc_safe_(function () { return fetchOPDrugMaster(); }, []);
    ctx.extDrugs   = ipc_safe_(function () { return fetchUniversalDrugs(); }, []);
    ctx.labCatalog = ipc_safe_(function () { return getOPDOrderableTests(); },
                               { panels: [], tests: [], packages: [] });
    ctx.templates  = ipc_safe_(function () {
                       var r = listConsultTemplates(selfDoctorId, sessionToken);
                       return r && r.success ? r.templates : [];
                     }, []);
    ctx.allergies  = ipc_safe_(function () {
                       var r = getPatientAllergies(patientId, sessionToken);
                       return r && r.success ? r.allergies : "";
                     }, "");
    ctx.careTeam   = ipc_safe_(function () {
                       var r = getIPCareTeam(ipNumber, sessionToken);
                       return r && r.success ? r.team : [];
                     }, []);
    ctx.lastCasesheet = ipc_safe_(function () { return ipc_lastCasesheet_(ipNumber); }, null);

    return ctx;
  } catch (e) {
    return { success: false, message: "Could not open the casesheet: " + e.message };
  }
}

/** Runs a reused engine, swallowing its failure so one dead sheet cannot
 *  blank the whole composer. */
function ipc_safe_(fn, fallback) {
  try {
    var v = fn();
    return (v === null || v === undefined) ? fallback : v;
  } catch (e) {
    Logger.log("IP casesheet context section failed: " + e.message);
    return fallback;
  }
}

/** The most recent casesheet for an admission, for "copy previous". */
function ipc_lastCasesheet_(ipNumber) {
  var sh = ipc_casesheetSheet_();
  var m = dc_headerMap_(sh);
  var data = sh.getDataRange().getDisplayValues();
  var ip = dc_upper_(ipNumber);

  for (var i = data.length - 1; i >= 1; i--) {
    if (dc_upper_(data[i][1]) !== ip) continue;
    var get = function (h) {
      return (m[h] === undefined) ? "" : dc_str_(data[i][m[h]]);
    };
    var parse = function (h) {
      try { return JSON.parse(get(h) || "[]"); } catch (e) { return []; }
    };
    return {
      encounterId: get("Encounter_ID"),
      timestamp:   get("Timestamp"),
      doctorName:  get("Doctor's Name"),
      complaints:  parse("Chief_Complaints"),
      history:     parse("History"),
      diagnosis:   get("Primary Diagnosis"),
      advice:      get("Advice"),
      meds:        parse("Prescription_JSON"),
      sysExam:     { cvs: get("CVS"), rs: get("RS"), pa: get("PA"), cns: get("CNS") }
    };
  }
  return null;
}

// ---------------------------------------------------------------------------
// SECTION C — SAVE
// ---------------------------------------------------------------------------

/**
 * Writes the admission casesheet and routes its orders.
 *
 * Care-team gated (Phase 5): the author must be an active doctor on this
 * admission's team. The signature is snapshotted into the row, so a later
 * rename in the Doctors sheet cannot rewrite a record that is already signed.
 */
function saveIPCasesheet(payload, sessionToken) {
  payload = payload || {};

  // The row write holds the script lock. Everything downstream —
  // createLabRequest and learnTemplatesScoped_ — takes that same lock itself,
  // so it runs AFTER the lock is released. Nesting the acquisitions would
  // stall the execution until waitLock timed out and threw.
  var saved = ipc_writeCasesheetRow_(payload, sessionToken);
  if (!saved.success) return saved;

  var result = {
    success: true,
    encounterId: saved.encounterId,
    doctorName: saved.doctorName,
    signature: saved.signature,
    warnings: saved.warnings || []
  };

  // ---- lab orders -> the real catalog, with billable test IDs ------------
  try {
    var labMsg = ipc_routeLabOrders_(payload, saved.patientId, saved.encounterId, saved.writer);
    if (labMsg) result.warnings.push(labMsg);
  } catch (e) {
    result.warnings.push("Lab orders were not routed: " + e.message);
  }

  // ---- the doctor's own phrase habits ------------------------------------
  try {
    if (payload.templateLearn && payload.templateLearn.length) {
      learnTemplatesScoped_(payload.templateLearn, saved.writer.doctorId);
    }
  } catch (e) { Logger.log("template learn skipped: " + e.message); }

  result.message = result.warnings.length
    ? "Casesheet saved, with warnings: " + result.warnings.join(" ")
    : "Casesheet saved and orders routed.";
  return result;
}

/** The locked section: validate, append the row, queue the ward's meds. */
function ipc_writeCasesheetRow_(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var w = resolveIPWrite_(sessionToken || payload.sessionToken, payload.ipNumber, "DOCTOR", {});
    if (!w.ok) return { success: false, message: w.message };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ipc_casesheetSheet_();

    var adm = ipc_admissionRow_(payload.ipNumber);
    if (!adm) return { success: false, message: "Admission not found." };

    var patientId = dc_upper_(payload.patientId) || dc_upper_(adm.row[1]);
    var timestamp = new Date();
    var dateStr = Utilities.formatDate(timestamp, ss.getSpreadsheetTimeZone(), "MM/dd/yyyy hh:mm a");
    var encounterId = "IP-ENC-" + patientId + "-" + timestamp.getTime().toString().slice(-6);

    var vitals  = payload.vitals  || {};
    var genExam = payload.genExam || {};
    var sysExam = payload.sysExam || {};
    var flags   = Array.isArray(genExam.flags) ? genExam.flags : [];
    var hasGE = function (f) { return flags.indexOf(f) !== -1 ? "Yes" : "No"; };

    // Header-driven write: the row is addressed by column name, never by the
    // position of a value in an array literal.
    var m = dc_headerMap_(sheet);
    var row = new Array(sheet.getLastColumn()).fill("");
    var put = function (h, v) {
      if (m[h] !== undefined) row[m[h]] = (v === undefined || v === null) ? "" : v;
    };

    put("Encounter_ID", encounterId);
    put("IP_Number", dc_upper_(payload.ipNumber));
    put("Ward", payload.ward || adm.row[7]);
    put("Bed", payload.bed || adm.row[8]);
    put("Patient_ID", patientId);
    put("Timestamp", dateStr);
    put("Patient Name", payload.patientName || adm.row[2]);
    put("Age", payload.age);
    put("Sex", payload.sex);
    put("Sys_BP", vitals.sys);   put("Dia_BP", vitals.dia);
    put("PR", vitals.hr);        put("SpO2", vitals.spo2);
    put("Temp", vitals.temp);    put("Height", vitals.height);
    put("Weight", vitals.weight);
    put("Chief_Complaints", JSON.stringify(payload.complaints || []));
    put("History", JSON.stringify(payload.history || []));
    put("Pallor", hasGE("Pallor"));     put("Icterus", hasGE("Icterus"));
    put("Cyanosis", hasGE("Cyanosis")); put("Clubbing", hasGE("Clubbing"));
    put("Edema", hasGE("Edema"));
    put("Other GE findings", genExam.notes);
    put("CVS", sysExam.cvs); put("RS", sysExam.rs);
    put("PA", sysExam.pa);   put("CNS", sysExam.cns);
    put("Primary Diagnosis", payload.provDiagnosis);
    put("Prescription_JSON", JSON.stringify(payload.meds || []));
    put("Lab_Orders_JSON", JSON.stringify(payload.labs || []));
    put("Outside Lab Records", JSON.stringify(payload.outsideLabs || []));
    put("Radiological records", payload.radiology);
    put("Advice", payload.advice);
    put("Doctor's Name", w.displayName);
    put("Doctor_ID", w.doctorId);
    put("Author_Signature_Snapshot", w.signature);
    put("Author_Username", w.username);

    sheet.appendRow(row);

    var warnings = [];

    // Admission orders share this lock: they are the same clinical act.
    try {
      ipc_routeAdmissionOrders_(ss, payload, patientId, encounterId, w, timestamp);
    } catch (e) {
      warnings.push("Medication orders were not queued: " + e.message);
    }

    // Keep the admission's working diagnosis current.
    try {
      var dx = dc_str_(payload.provDiagnosis);
      if (dx && dc_str_(adm.row[10]) !== dx) {
        adm.sheet.getRange(adm.rowNumber, 11).setValue(dx);
      }
    } catch (e) { /* the casesheet is saved; this is a convenience */ }

    try {
      logAudit_(w.sess, "IP_CASESHEET_SAVE", "IP_CaseSheet", encounterId, {
        patientId: patientId, ipNumber: payload.ipNumber,
        diagnosis: payload.provDiagnosis, doctorId: w.doctorId
      });
    } catch (e) { /* auditing must never fail a clinical save */ }

    SpreadsheetApp.flush();
    return {
      success: true,
      encounterId: encounterId,
      patientId: patientId,
      doctorName: w.displayName,
      signature: w.signature,
      writer: w,
      warnings: warnings
    };

  } catch (e) {
    return { success: false, message: "Could not save the casesheet: " + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/** Admission medication orders into the IP pharmacy queue. */
function ipc_routeAdmissionOrders_(ss, payload, patientId, encounterId, w, timestamp) {
  var meds = (payload.meds || []).filter(function (md) {
    return dc_str_(md.drugName) &&
           dc_upper_(md.source || "INTERNAL") !== "EXTERNAL";
  });
  if (!meds.length) return;

  var sheet = _ensureIPPharmacyQueueSheet_(ss);
  meds.forEach(function (md) {
    sheet.appendRow([
      "IPQ-" + Utilities.getUuid().substring(0, 8).toUpperCase(),
      dc_upper_(payload.ipNumber),
      patientId,
      (dc_str_(md.strength) + " " + dc_str_(md.drugName)).trim(),
      dc_str_(md.dose) || dc_str_(md.strength),
      dc_str_(md.freq) || dc_str_(md.sig),
      dc_str_(md.route) || "Oral",
      dc_str_(md.instructions) || dc_str_(md.comments),
      "Pending_Dispense",
      w.authorLabel,
      timestamp,
      "ACTIVE",
      "", "",
      encounterId
    ]);
  });
}

/**
 * Lab orders through the OPD bridge so IP orders carry real catalog IDs and
 * are billable and resultable exactly like OP ones. The old path passed bare
 * test NAMES, which the lab module could not price or match to a panel.
 */
function ipc_routeLabOrders_(payload, patientId, encounterId, w) {
  var orders = (payload.labs || []).filter(function (l) {
    return dc_upper_(l.type) === "ORDER" && dc_upper_(l.source || "INTERNAL") !== "EXTERNAL";
  });
  if (!orders.length) return "";

  var testIds = orders.map(function (l) { return dc_str_(l.testId); }).filter(Boolean);
  var hasStat = orders.some(function (l) { return dc_upper_(l.priority) === "STAT"; });

  if (testIds.length) {
    // createLabRequest directly, not the OPD bridge: that bridge hard-codes
    // sourceModule 'OPD', and an IP order booked as OPD asks the ward for
    // cash at sample collection instead of posting to the IP account.
    var res = createLabRequest({
      patientId:          patientId,
      sourceModule:       "IP_CASESHEET",
      visitId:            encounterId,
      admissionId:        dc_upper_(payload.ipNumber),
      testIds:            testIds,
      orderingDoctorId:   w.doctorId,
      orderingDoctorName: w.displayName,
      priority:           hasStat ? "STAT" : "ROUTINE",
      clinicalNote:       dc_str_(payload.provDiagnosis)
    });
    if (res && res.success === false) return dc_str_(res.message);
  }

  // Free-typed tests have no catalog id; they still belong on the record.
  var unlisted = orders.filter(function (l) { return !dc_str_(l.testId); })
                       .map(function (l) { return dc_str_(l.testName); })
                       .filter(Boolean);
  if (unlisted.length) {
    return "Not in the lab catalog, so not billable: " + unlisted.join(", ") + ".";
  }
  return "";
}

// ---------------------------------------------------------------------------
// SECTION D — PRINT
// ---------------------------------------------------------------------------

/**
 * Print-ready casesheet HTML.
 * Read by header and fully escaped: clinical free text routinely contains
 * "<" and "&", and the doctor column moved when Phase 5 added attribution.
 */
function getIPCasesheetPrintHtml(encounterId, sessionToken) {
  try {
    var sheet = ipc_casesheetSheet_();
    var m = dc_headerMap_(sheet);
    var data = sheet.getDataRange().getDisplayValues();

    var rec = null;
    for (var i = data.length - 1; i >= 1; i--) {
      if (dc_str_(data[i][0]) === dc_str_(encounterId)) { rec = data[i]; break; }
    }
    if (!rec) return { success: false, message: "Casesheet record not found." };

    var f = function (h) { return (m[h] === undefined) ? "" : dc_str_(rec[m[h]]); };

    var gate = resolveIPRead_(sessionToken, f("IP_Number"));
    if (!gate.ok) return { success: false, message: gate.message };

    var e = function (v) {
      return String(v === null || v === undefined ? "" : v)
        .replace(/&/g, "&amp;").replace(/</g, "&lt;")
        .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
    };
    var fe = function (h) { return e(f(h)); };
    var parse = function (h) { try { return JSON.parse(f(h) || "[]"); } catch (err) { return []; } };

    var complaints = parse("Chief_Complaints");
    var history    = parse("History");
    var meds       = parse("Prescription_JSON");
    var labs       = parse("Lab_Orders_JSON");
    var outside    = parse("Outside Lab Records");

    var listText = function (arr) {
      return arr.map(function (c) {
        if (typeof c === "string") return e(c);
        var t = e(c.condition || c.text || "");
        if (c.prefix) t = e(c.prefix) + " " + t;
        return c.duration ? t + " (" + e(c.duration) + ")" : t;
      }).filter(Boolean).join(" | ");
    };

    var medsHtml = meds.length
      ? '<table style="width:100%; border-collapse:collapse; font-size:0.85rem;">' +
        '<thead><tr style="background:#f1f5f9;">' +
        '<th style="text-align:left; padding:6px; border:1px solid #e2e8f0;">#</th>' +
        '<th style="text-align:left; padding:6px; border:1px solid #e2e8f0;">Drug</th>' +
        '<th style="text-align:left; padding:6px; border:1px solid #e2e8f0;">Dose / Frequency</th>' +
        '<th style="text-align:left; padding:6px; border:1px solid #e2e8f0;">Route</th>' +
        '<th style="text-align:left; padding:6px; border:1px solid #e2e8f0;">Duration</th>' +
        '</tr></thead><tbody>' +
        meds.map(function (md, i) {
          return '<tr>' +
            '<td style="padding:6px; border:1px solid #e2e8f0;">' + (i + 1) + '</td>' +
            '<td style="padding:6px; border:1px solid #e2e8f0;"><strong>' +
              e(md.type || "") + " " + e(md.strength || "") + " " + e(md.drugName || "") +
            '</strong></td>' +
            '<td style="padding:6px; border:1px solid #e2e8f0;">' +
              e(md.dose || "") + " " + e(md.freq || md.sig || "") + '</td>' +
            '<td style="padding:6px; border:1px solid #e2e8f0;">' + e(md.route || "Oral") + '</td>' +
            '<td style="padding:6px; border:1px solid #e2e8f0;">' + e(md.duration || "") + '</td>' +
          '</tr>';
        }).join("") +
        '</tbody></table>'
      : '<span style="color:#777;">No admission medication ordered.</span>';

    var orderedLabs = labs.filter(function (l) { return dc_upper_(l.type) === "ORDER"; })
                          .map(function (l) { return e(l.testName); }).join(", ");
    var outsideHtml = outside.length
      ? outside.map(function (l) { return e(l.test) + ": " + e(l.val); }).join(" | ")
      : "None";

    var signature = f("Author_Signature_Snapshot") || f("Doctor's Name") || "Doctor's Signature";

    var html =
    '<html><head><meta charset="utf-8"><title>IP Casesheet ' + fe("IP_Number") + '</title>' +
    '<style>@media print{@page{margin:14mm;} .noprint{display:none;}}' +
    'body{font-family:Arial,Helvetica,sans-serif;color:#000;margin:0;padding:20px;}' +
    'h5{margin:0 0 6px;padding-bottom:4px;border-bottom:1px solid #cbd5e1;color:#0369a1;font-size:0.9rem;}' +
    '.sec{margin-bottom:18px;font-size:0.88rem;}' +
    '.grid{display:flex;justify-content:space-between;gap:18px;}' +
    '</style></head><body>' +
    '<div style="max-width:820px;margin:auto;">' +
      '<div style="border-bottom:2px solid #0369a1;padding-bottom:10px;margin-bottom:18px;text-align:center;">' +
        '<h2 style="margin:0;text-transform:uppercase;color:#0369a1;">Valarmathi Clinic</h2>' +
        '<p style="margin:0;font-size:0.85rem;color:#555;">Premium Healthcare Services | Ph: +91 88387 23513</p>' +
        '<h4 style="margin:8px 0 0;">IP Admission Casesheet</h4>' +
      '</div>' +

      '<div class="grid" style="background:#f8fafc;padding:14px;border:1px solid #e2e8f0;border-radius:6px;margin-bottom:18px;font-size:0.88rem;">' +
        '<div><strong>Patient:</strong> ' + fe("Patient Name") + '<br>' +
          '<strong>PID:</strong> ' + fe("Patient_ID") + ' &nbsp;|&nbsp; <strong>IP No:</strong> ' + fe("IP_Number") + '<br>' +
          '<strong>Age/Sex:</strong> ' + fe("Age") + ' / ' + fe("Sex") + '<br>' +
          '<strong>Ward/Bed:</strong> ' + fe("Ward") + ' - ' + fe("Bed") + '</div>' +
        '<div style="text-align:right;"><strong>Date:</strong> ' + fe("Timestamp") + '<br>' +
          '<strong>Doctor:</strong> ' + fe("Doctor's Name") + '<br>' +
          '<span style="font-size:0.78rem;color:#555;">' + fe("Doctor_ID") + '</span></div>' +
      '</div>' +

      '<div class="sec"><h5>Vitals on Admission</h5>' +
        '<strong>BP:</strong> ' + fe("Sys_BP") + '/' + fe("Dia_BP") + ' mmHg &nbsp;|&nbsp; ' +
        '<strong>Pulse:</strong> ' + fe("PR") + ' bpm &nbsp;|&nbsp; ' +
        '<strong>SpO2:</strong> ' + fe("SpO2") + '% &nbsp;|&nbsp; ' +
        '<strong>Temp:</strong> ' + fe("Temp") + ' &deg;F &nbsp;|&nbsp; ' +
        '<strong>Wt:</strong> ' + fe("Weight") + ' kg' +
      '</div>' +

      '<div class="sec"><h5>Clinical History &amp; Examination</h5>' +
        '<p><strong>Chief Complaints:</strong> ' + (listText(complaints) || "--") + '</p>' +
        '<p><strong>History:</strong> ' + (listText(history) || "--") + '</p>' +
        '<p><strong>General Exam:</strong> Pallor: ' + fe("Pallor") + ', Icterus: ' + fe("Icterus") +
          ', Cyanosis: ' + fe("Cyanosis") + ', Clubbing: ' + fe("Clubbing") + ', Edema: ' + fe("Edema") +
          '<br><em>Notes:</em> ' + (fe("Other GE findings") || "--") + '</p>' +
        '<p><strong>Systemic Exam:</strong> CVS: ' + fe("CVS") + ' | RS: ' + fe("RS") +
          ' | P/A: ' + fe("PA") + ' | CNS: ' + fe("CNS") + '</p>' +
        '<p><strong>Primary Diagnosis:</strong> <span style="font-size:1.05em;font-weight:bold;">' +
          (fe("Primary Diagnosis") || "--") + '</span></p>' +
      '</div>' +

      '<div class="sec"><h5>Admission Orders</h5>' + medsHtml + '</div>' +

      '<div class="sec grid">' +
        '<div style="width:48%;">' +
          '<h5>Lab Orders</h5><p>' + (orderedLabs || "None") + '</p>' +
          '<h5 style="margin-top:14px;">External Lab Results</h5><p>' + outsideHtml + '</p>' +
        '</div>' +
        '<div style="width:48%;">' +
          '<h5>Radiology / Scans</h5><p style="white-space:pre-wrap;">' +
            (fe("Radiological records") || "None") + '</p>' +
          '<h5 style="margin-top:14px;">Advice / Plan</h5><p style="white-space:pre-wrap;">' +
            (fe("Advice") || "Standard ward protocol.") + '</p>' +
        '</div>' +
      '</div>' +

      '<div style="text-align:right;margin-top:56px;">' +
        '<div style="border-top:1px solid #000;display:inline-block;padding-top:5px;width:230px;text-align:center;">' +
          '<strong>' + e(signature) + '</strong>' +
        '</div>' +
      '</div>' +
    '</div></body></html>';

    return { success: true, html: html };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/** Previous casesheets for this admission, for the recall panel. */
function getIPCasesheetHistory(ipNumber, sessionToken) {
  try {
    var gate = resolveIPRead_(sessionToken, ipNumber);
    if (!gate.ok) return { success: false, message: gate.message, data: [] };

    var sh = ipc_casesheetSheet_();
    var m = dc_headerMap_(sh);
    var data = sh.getDataRange().getDisplayValues();
    var ip = dc_upper_(ipNumber);
    var out = [];

    for (var i = data.length - 1; i >= 1; i--) {
      if (dc_upper_(data[i][1]) !== ip) continue;
      var g = function (h) { return (m[h] === undefined) ? "" : dc_str_(data[i][m[h]]); };
      out.push({
        encounterId: g("Encounter_ID"),
        timestamp:   g("Timestamp"),
        doctorName:  g("Doctor's Name"),
        diagnosis:   g("Primary Diagnosis")
      });
    }
    return { success: true, data: out };
  } catch (e) {
    return { success: false, message: e.message, data: [] };
  }
}
