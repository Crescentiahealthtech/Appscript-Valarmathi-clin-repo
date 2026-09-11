// ============================================================================
// IP_Records_Logic.gs  —  Crescentia HealthTech
// The IP master record: the historical ledger and the complete patient file.
// ----------------------------------------------------------------------------
// WHAT CHANGED AND WHY
//
//  1. Progress notes now come from IP_Timeline_DB. This module used to read a
//     sheet called "IP_Notes" with a four-column schema that no writer in the
//     system has ever produced — the notes engine writes IP_Timeline_DB. The
//     Progress Notes tab therefore said "No progress notes recorded yet" for
//     every admission in the hospital, no matter how many notes were signed.
//
//  2. The ledger no longer lists superseded casesheets. Phase 7 made a
//     casesheet amendment write a new CURRENT row and mark the old one
//     SUPERSEDED; this module listed every row, so an amended chart appeared
//     two or three times and there was no way to tell which was live.
//
//  3. Every endpoint is gated. These functions returned any patient's chart to
//     any caller with no session token at all, bypassing the care-team scope
//     the rest of the IP modules enforce through resolveIPRead_.
//
//  4. Columns are addressed by header name, not by position. The hard-coded
//     indices here predated the Phase 5 attribution columns and the Phase 7
//     versioning columns.
//
//  5. Admissions with no casesheet are listed too, flagged as a gap. A missing
//     casesheet is precisely the record you need the ledger to show you.
// ============================================================================

var IPR_NOTE_LABELS = {
  DOCTOR: "Progress Note", NURSE: "Nursing Note", CONSULTANT: "Consultant Opinion",
  PROCEDURE: "Procedure Note", HANDOVER: "Shift Handover", QUICK: "Quick Note",
  INVESTIGATION: "Investigation"
};

// ---------------------------------------------------------------------------
// SECTION A — THE LEDGER
// ---------------------------------------------------------------------------

/**
 * Every IP encounter this session may read, newest first.
 *
 * Joins three sheets in one pass each: IP_Admissions (the spine),
 * IP_CaseSheets_DB (current versions only) and IP_Timeline_DB (note counts).
 * Per-admission lookups would be one full sheet scan per bed.
 *
 * @return {{success:boolean, data:Array, message:string}}
 */
function getIPRecordsLedger(sessionToken) {
  try {
    var gate = resolveIPRead_(sessionToken, null);
    if (!gate.ok) return { success: false, message: gate.message, data: [] };
    var mayRead = ipc_wardVisibilityFilter_(gate.scope);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var byIp = {};

    // ---- 1. admissions: the spine -----------------------------------------
    var admSh = ss.getSheetByName("IP_Admissions");
    if (admSh) {
      var admData = dc_sheetValues_(admSh);
      for (var i = 1; i < admData.length; i++) {
        var r = admData[i];
        var ip = dc_upper_(r[IPA_COL.IP]);
        if (!ip || !mayRead(ip)) continue;
        var wb = ipa_resolveWardBed_(r[IPA_COL.WARD], r[IPA_COL.BED]);
        byIp[ip] = {
          ipNumber:   dc_str_(r[IPA_COL.IP]),
          patientId:  dc_str_(r[IPA_COL.PATIENT_ID]),
          name:       dc_str_(r[IPA_COL.NAME]) || "Unknown",
          ageSex:     dc_str_(r[IPA_COL.AGE_SEX]),
          ward:       wb.ward, bed: wb.bed,
          consultant: dc_str_(r[IPA_COL.CONSULTANT]),
          diagnosis:  dc_str_(r[IPA_COL.DIAGNOSIS]),
          admittedOn: ipr_when_(r[IPA_COL.DOA]),
          admittedTs: ipr_ms_(r[IPA_COL.DOA]),
          dischargedOn: ipr_when_(r[IPA_COL.DOD]),
          status:     dc_upper_(r[IPA_COL.STATUS]) || "UNKNOWN",
          encounterId: "", casesheetOn: "", version: 0,
          noteCount: 0, lastNoteOn: "", lastNoteTs: 0
        };
      }
    }

    // ---- 2. casesheets: current versions only ------------------------------
    var csSh = ipc_casesheetSheet_();
    var csM  = dc_headerMap_(csSh);
    var csData = dc_sheetValues_(csSh);
    for (var c = 1; c < csData.length; c++) {
      var cr = csData[c];
      var g = function (h) { return (csM[h] === undefined) ? "" : dc_str_(cr[csM[h]]); };
      if (dc_upper_(g("Status")) === "SUPERSEDED") continue;

      var cip = dc_upper_(g("IP_Number"));
      if (!cip || !mayRead(cip)) continue;

      // An admission row may have been purged; the casesheet still stands on
      // its own, so seed an entry from it rather than dropping the record.
      var entry = byIp[cip] || (byIp[cip] = {
        ipNumber: g("IP_Number"), patientId: g("Patient_ID"),
        name: g("Patient Name") || "Unknown",
        ageSex: [g("Age"), g("Sex")].filter(Boolean).join(" / "),
        ward: g("Ward"), bed: g("Bed"),
        consultant: g("Doctor's Name"), diagnosis: g("Primary Diagnosis"),
        admittedOn: "", admittedTs: 0, dischargedOn: "", status: "UNKNOWN",
        encounterId: "", casesheetOn: "", version: 0,
        noteCount: 0, lastNoteOn: "", lastNoteTs: 0
      });

      entry.encounterId = g("Encounter_ID");
      entry.casesheetOn = g("Timestamp");
      entry.version     = parseInt(g("Version"), 10) || 1;
      if (!entry.diagnosis)  entry.diagnosis  = g("Primary Diagnosis");
      if (!entry.consultant) entry.consultant = g("Doctor's Name");
      if (!entry.admittedTs) entry.admittedTs = ipr_ms_(g("Timestamp"));
    }

    // ---- 3. timeline: note counts and recency ------------------------------
    try {
      var tlData = dc_sheetValues_(ipc_timelineSheet_());
      for (var t = 1; t < tlData.length; t++) {
        var tip = dc_upper_(tlData[t][1]);
        var ent = byIp[tip];
        if (!ent) continue;
        ent.noteCount++;
        var ms = ipr_ms_(tlData[t][0]);
        if (ms > ent.lastNoteTs) {
          ent.lastNoteTs = ms;
          ent.lastNoteOn = ipr_when_(tlData[t][0]);
        }
      }
    } catch (e) { /* a ledger without note counts still beats no ledger */ }

    var now = Date.now();
    var out = Object.keys(byIp).map(function (k) {
      var r = byIp[k];
      r.gaps = ipr_gaps_(r, now);
      r.sortTs = Math.max(r.lastNoteTs || 0, r.admittedTs || 0);
      return r;
    }).sort(function (a, b) { return b.sortTs - a.sortTs; });

    return { success: true, data: out, message: "" };
  } catch (e) {
    return { success: false, message: e.toString(), data: [] };
  }
}

/**
 * Record-completeness flags for one encounter.
 *
 * The point of a master-records screen is to find the chart that is missing
 * something before an auditor, an insurer or a court does. These are the three
 * omissions that actually matter on a ward.
 */
function ipr_gaps_(r, now) {
  var gaps = [];
  if (!r.encounterId) gaps.push("No casesheet");
  if (dc_upper_(r.status) === "ACTIVE") {
    if (!r.noteCount) gaps.push("No progress notes");
    else if (r.lastNoteTs && (now - r.lastNoteTs) > 24 * 3600 * 1000) {
      gaps.push("No note in " + Math.floor((now - r.lastNoteTs) / (3600 * 1000)) + "h");
    }
  }
  return gaps;
}

/** "dd MMM yyyy, hh:mm a", or "" for anything unparseable. */
function ipr_when_(v, pattern) {
  try {
    if (v === null || v === undefined || v === "") return "";
    var d = (v instanceof Date) ? v : new Date(v);
    if (isNaN(d.getTime())) return dc_str_(v);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), pattern || "dd MMM yyyy, hh:mm a");
  } catch (e) { return dc_str_(v); }
}

/** Epoch milliseconds, or 0. */
function ipr_ms_(v) {
  try {
    if (v === null || v === undefined || v === "") return 0;
    var d = (v instanceof Date) ? v : new Date(v);
    var ms = d.getTime();
    return isNaN(ms) ? 0 : ms;
  } catch (e) { return 0; }
}

// ---------------------------------------------------------------------------
// SECTION B — THE COMPLETE FILE
// ---------------------------------------------------------------------------

/**
 * Everything held on one encounter: casesheet, chronological notes, the
 * observation series and whatever discharge record exists.
 *
 * @param {string} encounterId  an Encounter_ID, or an IP number
 */
function getIPRecordFile(encounterId, sessionToken) {
  try {
    var found = ipr_locate_(encounterId);
    if (!found.ipNumber) {
      return { success: false, message: "No IP record matches " + dc_str_(encounterId) + "." };
    }

    var gate = resolveIPRead_(sessionToken, found.ipNumber);
    if (!gate.ok) return { success: false, message: gate.message };

    var adm = ipc_admissionRow_(found.ipNumber);
    var patient = ipn_printPatient_(found.ipNumber, adm);

    var out = {
      success:   true,
      ipNumber:  found.ipNumber,
      encounterId: found.encounterId,
      patient:   patient,
      casesheet: found.f ? ipr_casesheetView_(found.f) : null,
      notes:     [],
      vitals:    [],
      discharge: ipr_discharge_(found.ipNumber, found.encounterId, adm),
      gaps:      []
    };

    // Notes, oldest first — a file is read forwards. Sorted on the timestamp
    // rather than reversed, so a late entry cannot appear out of sequence.
    var tl = getIPTimeline(found.ipNumber, sessionToken);
    if (tl.success) {
      out.notes = tl.data.slice().sort(ipn_byTimeAsc_).map(function (n) {
        var d = n.noteData || {};
        return {
          noteId:    n.noteId,
          timestamp: n.timestamp,
          rawTs:     n.rawTs,
          roleType:  dc_upper_(n.roleType),
          label:     IPR_NOTE_LABELS[dc_upper_(n.roleType)] || n.roleType,
          author:    n.author,
          signature: n.signature,
          shift:     n.shift,
          // The composer's pre-canonical key still lives in older rows.
          noteData:  (d.systemExam && !d.sysExam)
                       ? ipr_withSysExam_(d)
                       : d
        };
      });
      out.vitals = out.notes.filter(function (n) { return n.noteData && n.noteData.vitals; })
        .map(function (n) { return { label: n.timestamp, ts: n.rawTs, vitals: n.noteData.vitals }; });
    }

    out.gaps = ipr_gaps_({
      encounterId: found.encounterId,
      status: adm ? dc_upper_(adm.row[IPA_COL.STATUS]) : "",
      noteCount: out.notes.length,
      lastNoteTs: out.notes.length ? ipr_ms_(out.notes[out.notes.length - 1].rawTs) : 0
    }, Date.now());

    return out;
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/** Copy of a noteData that carries the legacy key under the canonical one. */
function ipr_withSysExam_(d) {
  var copy = {};
  Object.keys(d).forEach(function (k) { copy[k] = d[k]; });
  copy.sysExam = d.systemExam;
  return copy;
}

/**
 * Resolves an Encounter_ID or an IP number to the current casesheet row.
 * @return {{ipNumber:string, encounterId:string, f:(function(string):string|null)}}
 */
function ipr_locate_(key) {
  var k = dc_upper_(key);
  var sh = ipc_casesheetSheet_();
  var m  = dc_headerMap_(sh);
  var data = dc_sheetValues_(sh);
  var iEnc = m["Encounter_ID"], iIp = m["IP_Number"], iStatus = m["Status"];

  var hit = null;
  for (var i = data.length - 1; i >= 1; i--) {
    var row = data[i];
    var isMatch = (iEnc !== undefined && dc_upper_(row[iEnc]) === k) ||
                  (iIp  !== undefined && dc_upper_(row[iIp])  === k);
    if (!isMatch) continue;
    // Prefer the live version; fall back to a superseded one if the caller
    // asked for that specific encounter by ID.
    var superseded = (iStatus !== undefined) && dc_upper_(row[iStatus]) === "SUPERSEDED";
    if (superseded && !(iEnc !== undefined && dc_upper_(row[iEnc]) === k)) continue;
    hit = row;
    if (!superseded) break;
  }

  if (hit) {
    var f = function (h) { return (m[h] === undefined) ? "" : dc_str_(hit[m[h]]); };
    return { ipNumber: f("IP_Number"), encounterId: f("Encounter_ID"), f: f };
  }

  // No casesheet — but the admission may still exist and carry notes.
  var adm = ipc_admissionRow_(k);
  if (adm) return { ipNumber: dc_str_(adm.row[IPA_COL.IP]), encounterId: "", f: null };
  return { ipNumber: "", encounterId: "", f: null };
}

/** The casesheet in the shape the records screen renders. */
function ipr_casesheetView_(f) {
  var parse = function (h) { try { return JSON.parse(f(h) || "[]"); } catch (e) { return []; } };
  var geFlags = ["Pallor", "Icterus", "Cyanosis", "Clubbing", "Edema"];

  return {
    encounterId: f("Encounter_ID"),
    ipNumber:    f("IP_Number"),
    patientId:   f("Patient_ID"),
    name:        f("Patient Name"),
    ageSex:      [f("Age"), f("Sex")].filter(Boolean).join(" / "),
    recordedOn:  f("Timestamp"),
    doctor:      f("Doctor's Name"),
    doctorId:    f("Doctor_ID"),
    signature:   f("Author_Signature_Snapshot"),
    vitals: {
      sys: f("Sys_BP"), dia: f("Dia_BP"), pr: f("PR"), spo2: f("SpO2"),
      temp: f("Temp"), height: f("Height"), weight: f("Weight")
    },
    complaints:  parse("Chief_Complaints"),
    history:     parse("History"),
    genExam: {
      positive: geFlags.filter(function (k) { return dc_upper_(f(k)) === "YES"; }),
      negative: geFlags.filter(function (k) { return dc_upper_(f(k)) === "NO"; }),
      notes:    f("Other GE findings")
    },
    sysExam:     { cvs: f("CVS"), rs: f("RS"), pa: f("PA"), cns: f("CNS") },
    diagnosis:   f("Primary Diagnosis"),
    meds:        parse("Prescription_JSON"),
    labs:        parse("Lab_Orders_JSON"),
    outsideLabs: parse("Outside Lab Records"),
    radiology:   f("Radiological records"),
    advice:      f("Advice"),
    status:      f("Status") || "CURRENT",
    version:     parseInt(f("Version"), 10) || 1,
    amendedAt:   f("Amended_At"),
    amendedBy:   f("Amended_By"),
    amendReason: f("Amend_Reason")
  };
}

/**
 * Whatever discharge record exists for this encounter.
 *
 * Read by header rather than by position: two modules in this project read a
 * "Discharge_Summary" sheet with two incompatible column orders, and nothing
 * in the project writes it. Rather than guess, this matches on whichever of
 * IP_Number / Encounter_ID / Patient_ID the sheet actually has, and falls back
 * to the admission row's own discharge status — which IS authoritative.
 */
function ipr_discharge_(ipNumber, encounterId, adm) {
  var fallback = null;
  if (adm && dc_upper_(adm.row[IPA_COL.STATUS]) === "DISCHARGED") {
    fallback = {
      source:  "admission",
      date:    ipr_when_(adm.row[IPA_COL.DOD], "dd MMM yyyy"),
      outcome: "Discharged",
      summary: "",
      diagnosis: dc_str_(adm.row[IPA_COL.DIAGNOSIS])
    };
  }

  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Discharge_Summary");
    if (!sh || sh.getLastRow() < 2) return fallback;

    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    var pick = function (row, names) {
      for (var i = 0; i < names.length; i++) {
        if (m[names[i]] !== undefined) {
          var v = dc_str_(row[m[names[i]]]);
          if (v) return v;
        }
      }
      return "";
    };

    for (var r = data.length - 1; r >= 1; r--) {
      var row = data[r];
      var key = dc_upper_(pick(row, ["Encounter_ID", "IP_Number"]));
      if (key !== dc_upper_(encounterId) && key !== dc_upper_(ipNumber)) continue;
      return {
        source:    "summary",
        date:      ipr_when_(pick(row, ["Discharge_Date", "Date", "Timestamp"]), "dd MMM yyyy"),
        outcome:   pick(row, ["Outcome", "Condition", "Status"]) || "Discharged",
        summary:   pick(row, ["Summary", "Discharge_Summary", "Notes"]),
        diagnosis: pick(row, ["Final_Diagnosis", "Diagnosis"])
      };
    }
  } catch (e) { /* a malformed summary sheet must not hide the admission fact */ }

  return fallback;
}

// ---------------------------------------------------------------------------
// SECTION C — THE COMPLETE FILE, PRINTED
//
// One continuous document: casesheet, observation chart, every progress note
// in order, and the discharge record. This is the artefact a ward actually
// needs — for the patient's file, for an insurer, for a referral — and until
// now it could only be produced by printing three screens separately and
// stapling them, each with a different letterhead and margin.
// ---------------------------------------------------------------------------

/**
 * @param {string} encounterId  an Encounter_ID or an IP number
 * @param {Object} [opts]  { casesheet, notes, trend, discharge } booleans;
 *                         omitted means include everything.
 */
function getIPRecordPrintHtml(encounterId, sessionToken, opts) {
  try {
    opts = opts || {};
    var want = function (k) { return opts[k] !== false; };

    var file = getIPRecordFile(encounterId, sessionToken);
    if (!file.success) return file;

    var sections = [];

    if (want("casesheet")) {
      sections.push(file.casesheet
        ? ipr_casesheetSection_(file.encounterId)
        : ipp_sec_("Admission Casesheet",
            '<span class="muted">No casesheet was recorded for this admission.</span>',
            { keepEmpty: true }));
    }

    if (want("trend") && file.vitals.length >= 2) {
      sections.push(ipp_sec_("Observation Chart", ipp_vitalsTrend_(file.vitals), { loose: true }));
    }

    if (want("notes")) {
      sections.push(ipp_sec_("Progress Notes",
        file.notes.length
          ? file.notes.map(ipn_printNote_).join("")
          : '<span class="muted">No progress notes were recorded.</span>',
        { loose: true, keepEmpty: true }));
    }

    if (want("discharge") && file.discharge) {
      sections.push(ipp_sec_("Discharge", ipp_kv_([
        ["Discharged on",   ipp_esc_(file.discharge.date)],
        ["Outcome",         ipp_esc_(file.discharge.outcome)],
        ["Final Diagnosis", ipp_esc_(file.discharge.diagnosis)],
        ["Summary",         ipp_escMultiline_(file.discharge.summary)]
      ], { narrow: true })));
    }

    var html = ipp_doc_({
      docTitle: "Complete Inpatient Record",
      patient:  file.patient,
      bodyHtml: sections.filter(Boolean).join("") +
                ipp_sig_(file.patient.consultant || "Consultant", "Treating Consultant"),
      footNote: (file.encounterId ? "Encounter " + file.encounterId + " · " : "") +
                file.notes.length + " note(s)"
    });

    return { success: true, html: html, noteCount: file.notes.length };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/** The casesheet sections, composed by the casesheet engine's own renderer. */
function ipr_casesheetSection_(encounterId) {
  var sh = ipc_casesheetSheet_();
  var m  = dc_headerMap_(sh);
  var data = sh.getDataRange().getDisplayValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (dc_str_(data[i][0]) !== dc_str_(encounterId)) continue;
    var row = data[i];
    var f = function (h) { return (m[h] === undefined) ? "" : dc_str_(row[m[h]]); };
    return ipp_sec_("Admission Casesheet",
      ipp_kv_([["Recorded", ipp_esc_(f("Timestamp"))],
               ["By",       ipp_esc_(f("Doctor's Name"))]], { narrow: true }) +
      '<div style="height:8px;"></div>' +
      ipc_composeCasesheetHtml_(f, { standalone: false }),
      { loose: true, keepEmpty: true });
  }
  return "";
}

// ---------------------------------------------------------------------------
// SECTION D — LEGACY ENTRY POINTS
//
// Kept so a browser tab left open on the old build degrades to an honest error
// instead of silently rendering an unauthorised chart.
// ---------------------------------------------------------------------------

/** @deprecated Use getIPRecordsLedger(sessionToken). */
function fetchIPRecordsLedger() {
  return JSON.stringify({ error: "This build requires a signed-in session. Reload the page." });
}

/** @deprecated Use getIPRecordFile(encounterId, sessionToken). */
function fetchFullIPRecord() {
  return JSON.stringify({ error: "This build requires a signed-in session. Reload the page." });
}
