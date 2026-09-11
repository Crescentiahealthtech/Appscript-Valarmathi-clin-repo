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
    // One admission carries one casesheet. Once it exists, the bed leaves this
    // picker — the sheet is amended from IP Records, not written again.
    var done = ipc_admissionsWithCasesheet_();
    var out = [];
    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][11]) !== "ACTIVE") continue;
      if (!canSee(data[i][0])) continue;
      if (done[dc_upper_(data[i][0])]) continue;

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

/** { IP_NUMBER : true } for every admission that already has a CURRENT casesheet. */
function ipc_admissionsWithCasesheet_() {
  var out = {};
  try {
    var sh = ipc_casesheetSheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    var iStatus = m["Status"];
    for (var i = 1; i < data.length; i++) {
      var ip = dc_upper_(data[i][1]);
      if (!ip) continue;
      // Rows written before the Status column existed are current by default.
      var st = (iStatus === undefined) ? "" : dc_upper_(data[i][iStatus]);
      if (st === "SUPERSEDED") continue;
      out[ip] = true;
    }
  } catch (e) { /* if this fails the picker simply shows every bed */ }
  return out;
}

/** The CURRENT casesheet row for an admission, or null. */
function ipc_currentCasesheet_(ipNumber) {
  var sh = ipc_casesheetSheet_();
  var m = dc_headerMap_(sh);
  var data = dc_sheetValues_(sh);
  var ip = dc_upper_(ipNumber);
  var iStatus = m["Status"];

  for (var i = data.length - 1; i >= 1; i--) {
    if (dc_upper_(data[i][1]) !== ip) continue;
    var st = (iStatus === undefined) ? "" : dc_upper_(data[i][iStatus]);
    if (st === "SUPERSEDED") continue;
    return { row: data[i], rowNumber: i + 1, headers: m, sheet: sh };
  }
  return null;
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

    // An admission carries ONE casesheet. Saving again is an amendment, which
    // goes through amendIPCasesheet() so it records a reason and keeps the
    // superseded version. Without this guard, re-opening a completed chart and
    // pressing Save wrote a duplicate row and re-queued every drug order.
    var existing = ipc_currentCasesheet_(payload.ipNumber);
    if (existing && !payload.__amending) {
      return {
        success: false,
        alreadyExists: true,
        encounterId: dc_str_(existing.row[0]),
        message: "This admission already has a casesheet (" +
                 dc_str_(existing.row[0]) + "). Open it from IP Records to amend it."
      };
    }

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
    put("Status", "CURRENT");
    put("Version", payload.__version || 1);
    put("Amended_At", payload.__amending ? timestamp : "");
    put("Amended_By", payload.__amending ? w.authorLabel : "");
    put("Amend_Reason", payload.__amending ? dc_str_(payload.__amendReason) : "");

    sheet.appendRow(row);
    dc_invalidate_("IP_CaseSheets_DB");

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
      // The drug NAME only. This used to be written as strength + name
      // ("500 Tab Azithral") while IP Notes sends the name as displayed
      // ("Tab Azithral 500"), so a STOP order never matched its own row and
      // the drug stayed on the Continue list for every later note.
      dc_str_(md.drugName),
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
 *
 * Read by header (the doctor column moved when Phase 5 added attribution) and
 * fully escaped: clinical free text routinely contains "<" and "&".
 * Composed through IP_Print_Kit so the casesheet, the progress record and the
 * full-file print share one letterhead, one grid and one set of page-break
 * rules — previously each had its own inline CSS and they had drifted apart.
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

    return {
      success: true,
      html: ipc_composeCasesheetHtml_(f, { standalone: true })
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * The casesheet body. Shared by the standalone print and the full-file print,
 * so the admission chart never has two different layouts.
 *
 * @param {function(string):string} f      header-addressed field reader
 * @param {Object} opts   { standalone:true } wraps the sections in a document;
 *                        otherwise only the sections are returned.
 */
function ipc_composeCasesheetHtml_(f, opts) {
  opts = opts || {};
  var e  = ipp_esc_;
  var fe = function (h) { return e(f(h)); };
  var parse = function (h) { try { return JSON.parse(f(h) || "[]"); } catch (err) { return []; } };

  var complaints = parse("Chief_Complaints");
  var history    = parse("History");
  var meds       = parse("Prescription_JSON");
  var labs       = parse("Lab_Orders_JSON");
  var outside    = parse("Outside Lab Records");

  // Complaints and history are structured {prefix, condition, duration}. One
  // per line, so a long problem list reads as a list instead of a paragraph.
  var listLines = function (arr) {
    return arr.map(function (c) {
      if (typeof c === "string") return e(c);
      var t = e(c.condition || c.text || "");
      if (!t) return "";
      if (c.prefix)   t = e(c.prefix) + " " + t;
      if (c.duration) t += ' <span class="muted">(' + e(c.duration) + ')</span>';
      return t;
    }).filter(Boolean).join("<br>");
  };

  // ---- vitals: only what was measured -------------------------------------
  var bp = (f("Sys_BP") || f("Dia_BP"))
    ? (f("Sys_BP") || "--") + "/" + (f("Dia_BP") || "--") + " mmHg" : "";
  // Two columns: six vitals stacked in one made the block taller than the
  // examination findings below it, for no reason.
  var vitalsKv = ipp_cols_(
    ipp_kv_([
      ["Blood Pressure", e(bp)],
      ["Pulse",  f("PR")   ? fe("PR") + " bpm"  : ""],
      ["SpO2",   f("SpO2") ? fe("SpO2") + " %"  : ""]
    ], { narrow: true }),
    ipp_kv_([
      ["Temp",   f("Temp")   ? fe("Temp") + " °F"   : ""],
      ["Height", f("Height") ? fe("Height") + " cm" : ""],
      ["Weight", f("Weight") ? fe("Weight") + " kg" : ""]
    ], { narrow: true })
  );

  // ---- general examination: positive findings first ------------------------
  var geFlags = ["Pallor", "Icterus", "Cyanosis", "Clubbing", "Edema"];
  var positive = geFlags.filter(function (k) { return dc_upper_(f(k)) === "YES"; });
  var negative = geFlags.filter(function (k) { return dc_upper_(f(k)) === "NO"; });
  var geKv = ipp_kv_([
    ["Positive", positive.length ? '<strong>' + e(positive.join(", ")) + '</strong>' : ""],
    ["Negative", negative.length ? '<span class="muted">' + e(negative.join(", ")) + '</span>' : ""],
    ["Notes",    ipp_escMultiline_(f("Other GE findings"))]
  ], { narrow: true });

  // ---- systemic examination: one system per row ----------------------------
  var seKv = ipp_kv_([
    ["CVS", fe("CVS")], ["RS", fe("RS")], ["P/A", fe("PA")], ["CNS", fe("CNS")]
  ], { narrow: true });

  // ---- medications ---------------------------------------------------------
  var medsTable = ipp_table_(
    [{ label: "#", cls: "ctr" }, "Drug", "Dose / Frequency", { label: "Route", cls: "ctr" }, { label: "Duration", cls: "ctr" }],
    meds.map(function (md, i) {
      var name = [md.type, md.drugName || md.name, md.strength]
        .map(dc_str_).filter(Boolean).join(" ");
      var note = dc_str_(md.comment || md.instructions);
      return [
        String(i + 1),
        '<strong>' + e(name) + '</strong>' +
          (note ? '<br><span class="muted">' + e(note) + '</span>' : ''),
        e([md.dose, md.freq || md.sig].map(dc_str_).filter(Boolean).join(" ")),
        e(md.route || "Oral"),
        e(md.duration || md.days || "")
      ];
    }),
    ["10mm", "auto", "38mm", "20mm", "22mm"]
  ) || '<span class="muted">No admission medication ordered.</span>';

  // ---- orders and results --------------------------------------------------
  var orderedLabs = labs.filter(function (l) { return dc_upper_(l.type) === "ORDER"; })
                        .map(function (l) { return e(l.testName); }).filter(Boolean).join(", ");
  var outsideHtml = outside.map(function (l) {
      var t = e(l.test), v = e(l.val);
      return t ? ('<strong>' + t + ':</strong> ' + (v || "--")) : "";
    }).filter(Boolean).join("<br>");

  // ---- version / amendment provenance -------------------------------------
  var amendKv = "";
  if (dc_str_(f("Amend_Reason")) || dc_upper_(f("Status")) === "SUPERSEDED" ||
      (parseInt(f("Version"), 10) || 1) > 1) {
    amendKv = ipp_kv_([
      ["Version",  fe("Version") || "1"],
      ["Status",   fe("Status") || "CURRENT"],
      ["Amended",  f("Amended_At") ? fe("Amended_At") + (f("Amended_By") ? " by " + fe("Amended_By") : "") : ""],
      ["Reason",   ipp_escMultiline_(f("Amend_Reason"))],
      ["Superseded by", fe("Superseded_By")]
    ], { narrow: true });
  }

  var body =
    ipp_sec_("Vitals on Admission", vitalsKv) +
    ipp_sec_("Presenting Complaints & History",
      ipp_cols_(
        ipp_kv_([["Chief Complaints", listLines(complaints)]], { narrow: true }),
        ipp_kv_([["History", listLines(history)]], { narrow: true })
      )) +
    ipp_sec_("Examination", ipp_cols_(
      '<div style="font-weight:600;color:#334155;margin-bottom:3px;font-size:9pt;">General</div>' + (geKv || '<span class="muted">Not recorded.</span>'),
      '<div style="font-weight:600;color:#334155;margin-bottom:3px;font-size:9pt;">Systemic</div>' + (seKv || '<span class="muted">Not recorded.</span>')
    )) +
    ipp_sec_("Provisional Diagnosis",
      f("Primary Diagnosis") ? '<div class="dx">' + fe("Primary Diagnosis") + '</div>' : "") +
    ipp_sec_("Admission Medication Orders", medsTable, { loose: true, keepEmpty: true }) +
    ipp_sec_("Investigations & Results", ipp_cols_(
      ipp_kv_([
        ["Lab Orders",    orderedLabs],
        ["External Labs", outsideHtml]
      ], { narrow: true }),
      ipp_kv_([
        ["Radiology", ipp_escMultiline_(f("Radiological records"))]
      ], { narrow: true })
    )) +
    ipp_sec_("Advice / Plan",
      ipp_escMultiline_(f("Advice")) || '<span class="muted">Standard ward protocol.</span>',
      { keepEmpty: true }) +
    ipp_sec_("Record Provenance", amendKv);

  if (!opts.standalone) return body;

  var wardBed = [f("Ward"), f("Bed")].filter(function (x) { return !!x; }).join(" / ");
  var ageSex  = [f("Age"), f("Sex")].filter(function (x) { return !!x; }).join(" / ");
  var signature = f("Author_Signature_Snapshot") || f("Doctor's Name");

  return ipp_doc_({
    docTitle: "Inpatient Admission Casesheet",
    patient: {
      name:       f("Patient Name"),
      pid:        f("Patient_ID"),
      ipNumber:   f("IP_Number"),
      ageSex:     ageSex,
      wardBed:    wardBed,
      consultant: f("Doctor's Name"),
      diagnosis:  f("Primary Diagnosis")
    },
    bodyHtml: ipp_kv_([
        ["Recorded", fe("Timestamp")],
        ["Encounter", fe("Encounter_ID")]
      ], { narrow: true }) +
      '<div style="height:10px;"></div>' + body +
      ipp_sig_(signature, "Admitting Doctor" + (f("Doctor_ID") ? " · " + f("Doctor_ID") : "")),
    footNote: "Encounter " + f("Encounter_ID")
  });
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

// ---------------------------------------------------------------------------
// SECTION E — AMENDMENTS
//
// A casesheet is a signed clinical record. It is never edited in place and
// never re-entered: an amendment writes a NEW current version, marks the
// previous one SUPERSEDED with a pointer to its replacement, and records who
// changed it and why. The whole chain stays readable.
// ---------------------------------------------------------------------------

/** The current casesheet for an admission, in the shape the composer loads. */
function getIPCasesheetForEdit(ipNumber, sessionToken) {
  try {
    var gate = resolveIPRead_(sessionToken, ipNumber);
    if (!gate.ok) return { success: false, message: gate.message };

    var cur = ipc_currentCasesheet_(ipNumber);
    if (!cur) return { success: false, message: "No casesheet on file for " + ipNumber + "." };

    var m = cur.headers;
    var g = function (h) { return (m[h] === undefined) ? "" : dc_str_(cur.row[m[h]]); };
    var parse = function (h) { try { return JSON.parse(g(h) || "[]"); } catch (e) { return []; } };

    var flags = [];
    ["Pallor", "Icterus", "Cyanosis", "Clubbing", "Edema"].forEach(function (f) {
      if (dc_upper_(g(f)) === "YES") flags.push(f);
    });

    return {
      success: true,
      encounterId: g("Encounter_ID"),
      version:     dc_int_(g("Version")) || 1,
      status:      g("Status") || "CURRENT",
      recordedBy:  g("Doctor's Name"),
      recordedAt:  g("Timestamp"),
      data: {
        ipNumber:   g("IP_Number"),
        patientId:  g("Patient_ID"),
        patientName: g("Patient Name"),
        age: g("Age"), sex: g("Sex"),
        ward: g("Ward"), bed: g("Bed"),
        vitals: {
          sys: g("Sys_BP"), dia: g("Dia_BP"), hr: g("PR"), spo2: g("SpO2"),
          temp: g("Temp"), height: g("Height"), weight: g("Weight")
        },
        complaints:  parse("Chief_Complaints"),
        history:     parse("History"),
        genExam:     { flags: flags, notes: g("Other GE findings") },
        sysExam:     { cvs: g("CVS"), rs: g("RS"), pa: g("PA"), cns: g("CNS") },
        provDiagnosis: g("Primary Diagnosis"),
        meds:        parse("Prescription_JSON"),
        labs:        parse("Lab_Orders_JSON"),
        outsideLabs: parse("Outside Lab Records"),
        radiology:   g("Radiological records"),
        advice:      g("Advice")
      }
    };
  } catch (e) {
    return { success: false, message: "Could not load the casesheet: " + e.message };
  }
}

/**
 * Records an amended casesheet. Requires a reason — an unexplained change to a
 * signed record is worse than no change at all.
 *
 * Medication and lab orders are NOT re-routed: the ward already has them, and
 * re-queueing on every amendment is what produced the duplicate drug cards.
 * Order changes belong in a progress note, which is where the ward looks.
 */
function amendIPCasesheet(payload, sessionToken) {
  try {
    payload = payload || {};
    var reason = dc_str_(payload.amendReason);
    if (reason.length < 5) {
      return { success: false, message: "Give a reason for the amendment (at least a few words)." };
    }

    var cur = ipc_currentCasesheet_(payload.ipNumber);
    if (!cur) return { success: false, message: "No casesheet to amend for " + payload.ipNumber + "." };

    var m = cur.headers;
    var prevId = dc_str_(cur.row[0]);
    var prevVersion = (m["Version"] === undefined) ? 1 : (dc_int_(cur.row[m["Version"]]) || 1);

    // Write the new version first: if this fails, the old one is still current.
    var amended = {};
    Object.keys(payload).forEach(function (k) { amended[k] = payload[k]; });
    amended.__amending = true;
    amended.__version = prevVersion + 1;
    amended.__amendReason = reason;
    amended.meds = [];        // orders are not re-queued from an amendment
    amended.labs = (payload.labs || []).map(function (l) {
      var c = {}; Object.keys(l).forEach(function (k) { c[k] = l[k]; });
      c.type = "Result";      // keep them on the record, out of the order path
      return c;
    });

    var saved = ipc_writeCasesheetRow_(amended, sessionToken);
    if (!saved.success) return saved;

    // Then retire the previous version.
    try {
      var sh = ipc_casesheetSheet_();
      if (m["Status"] !== undefined) {
        sh.getRange(cur.rowNumber, m["Status"] + 1).setValue("SUPERSEDED");
      }
      if (m["Superseded_By"] !== undefined) {
        sh.getRange(cur.rowNumber, m["Superseded_By"] + 1).setValue(saved.encounterId);
      }
      dc_invalidate_("IP_CaseSheets_DB");
      SpreadsheetApp.flush();
    } catch (e) {
      return { success: false,
               message: "The amendment was written as " + saved.encounterId +
                        " but the previous version could not be retired: " + e.message +
                        " Two versions are now marked current — fix this before continuing." };
    }

    // The amendment is itself a clinical event, so it belongs on the timeline.
    try {
      saveIPNote({
        ipNumber:  payload.ipNumber,
        patientId: saved.patientId,
        roleType:  "DOCTOR",
        flags:     "CASESHEET_AMENDED",
        noteData:  {
          assessment: "Admission casesheet amended (v" + prevVersion +
                      " → v" + (prevVersion + 1) + ").",
          plan: "Reason: " + reason
        }
      }, sessionToken);
    } catch (e) { /* the amendment is saved; the timeline note is a courtesy */ }

    return {
      success: true,
      encounterId: saved.encounterId,
      supersededId: prevId,
      version: prevVersion + 1,
      message: "Casesheet amended. Version " + (prevVersion + 1) +
               " is now current; v" + prevVersion + " is kept as superseded."
    };
  } catch (e) {
    return { success: false, message: "Could not amend the casesheet: " + e.message };
  }
}

/** Full version history for an admission, newest first. */
function getIPCasesheetVersions(ipNumber, sessionToken) {
  try {
    var gate = resolveIPRead_(sessionToken, ipNumber);
    if (!gate.ok) return { success: false, message: gate.message, data: [] };

    var sh = ipc_casesheetSheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    var ip = dc_upper_(ipNumber);
    var out = [];

    for (var i = data.length - 1; i >= 1; i--) {
      if (dc_upper_(data[i][1]) !== ip) continue;
      var g = function (h) { return (m[h] === undefined) ? "" : dc_str_(data[i][m[h]]); };
      out.push({
        encounterId:  g("Encounter_ID"),
        version:      dc_int_(g("Version")) || 1,
        status:       g("Status") || "CURRENT",
        supersededBy: g("Superseded_By"),
        recordedAt:   g("Timestamp"),
        recordedBy:   g("Doctor's Name"),
        amendedAt:    g("Amended_At"),
        amendedBy:    g("Amended_By"),
        amendReason:  g("Amend_Reason"),
        diagnosis:    g("Primary Diagnosis")
      });
    }
    return { success: true, data: out };
  } catch (e) {
    return { success: false, message: e.message, data: [] };
  }
}
