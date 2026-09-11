// ==========================================
// 🏥 IP NOTES & TIMELINE MODULE — BACKEND
// Crescentia HealthTech | Valarmathi Clinic
// ==========================================
// DB Schema (IP_Timeline_DB) — Phase 6 authorship columns appended:
// [0] Timestamp | [1] IP_Number | [2] Patient_ID | [3] Role_Type
// [4] Note_Data_JSON | [5] Author | [6] Shift | [7] Flags
// [8] Note_ID | [9] Author_Doctor_ID | [10] Author_Signature_Snapshot
// [11] Author_Username
//
// AUTHORSHIP RULE (Phase 6): payload.author from the client is IGNORED.
// Every write resolves its author from the session token through
// resolveIPWrite_() in IP_Clinical_Access.gs, which also strips any note
// section the caller's role may not author.
//
// DB Schema (IP_Pharmacy_Queue):
// [0] Queue_ID | [1] IP_Number | [2] Patient_ID | [3] Drug_Name
// [4] Dose | [5] Frequency | [6] Route | [7] Instructions
// [8] Status | [9] Ordered_By | [10] Ordered_At | [11] Action_Flag
// [12] Modified_At | [13] Modified_By | [14] Encounter_Note_ID
// ==========================================

// ── HELPERS ──────────────────────────────────────────────

function _ensureIPTimelineSheet_(ss) {
  // Delegates to the Phase 6 schema owner so the authorship columns are
  // guaranteed present on every write path, including legacy sheets that
  // were created with only the original eight columns.
  return ipc_timelineSheet_();
}

function _ensureIPPharmacyQueueSheet_(ss) {
  let sheet = ss.getSheetByName('IP_Pharmacy_Queue');
  if (!sheet) {
    sheet = ss.insertSheet('IP_Pharmacy_Queue');
    sheet.appendRow([
      "Queue_ID", "IP_Number", "Patient_ID", "Drug_Name",
      "Dose", "Frequency", "Route", "Instructions",
      "Status", "Ordered_By", "Ordered_At", "Action_Flag",
      "Modified_At", "Modified_By", "Encounter_Note_ID"
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function _ensureLabQueueSheet_(ss) {
  let sheet = ss.getSheetByName('Lab_Queue_DB');
  if (!sheet) {
    sheet = ss.insertSheet('Lab_Queue_DB');
    sheet.appendRow([
      "Order_ID", "Encounter_ID", "Patient_ID", "IP_Number",
      "Ordered_At", "Test_Name", "Priority", "Status",
      "Result_JSON", "Reported_At"
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Canonical form of a drug name for matching queue rows.
 * Rows have been written with the strength ahead of the name, behind it, and
 * with inconsistent spacing and case. Comparing raw strings meant a STOP order
 * silently failed to find the drug it was stopping.
 */
function _normDrug_(name) {
  return String(name || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim()
    .split(" ")
    .sort()            // "azithral tab 500" === "500 azithral tab"
    .join(" ");
}

function _getCurrentShift_() {
  const hour = new Date().getHours();
  if (hour >= 6 && hour < 14) return "MORNING";
  if (hour >= 14 && hour < 22) return "EVENING";
  return "NIGHT";
}

function _generateNoteId_(prefix) {
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyMMddHHmm");
  const rand = Math.random().toString(36).substr(2, 4).toUpperCase();
  return prefix + ts + "-" + rand;
}

// ── 1. FETCH ACTIVE WARD ROSTER ───────────────────────────

/**
 * Returns all ACTIVE admissions for the IP Notes ward roster.
 * IP_Admissions schema:
 * [0]IP_Number [1]Patient_ID [2]Name [3]Age/Sex [4]DOA [5]TOA
 * [6]Type [7]Ward [8]Bed [9]Consultant [10]Diagnosis [11]Status [12]DOD
 */
function getActiveIPAdmissionsForNotes(sessionToken) {
  try {
    var gate = resolveIPRead_(sessionToken, null);
    if (!gate.ok) return { success: false, message: gate.message, data: [] };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('IP_Admissions');
    if (!sheet) return { success: false, message: "IP_Admissions sheet not found.", data: [] };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: [], scopeMode: gate.scope.mode };

    // One pre-computed visibility lookup for the whole ward. Calling
    // ipc_mayReadAdmission_ per row would rescan IP_Care_Team per bed.
    const canSee = ipc_wardVisibilityFilter_(gate.scope);

    const activeAdmissions = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      const status = row[11] ? row[11].toString().trim().toUpperCase() : "";
      if (status !== 'ACTIVE') continue;
      if (!canSee(row[0])) continue;

      let doaFormatted = "--";
      try {
        doaFormatted = Utilities.formatDate(
          new Date(row[4]),
          Session.getScriptTimeZone(),
          "dd-MMM-yyyy"
        );
      } catch(e) { doaFormatted = row[4] ? row[4].toString() : "--"; }

      activeAdmissions.push({
        ipNumber:   String(row[0]).trim(),
        patientId:  String(row[1] || "").trim(),
        name:       String(row[2] || "Unknown").trim(),
        ageSex:     String(row[3] || "--").trim(),
        doa:        doaFormatted,
        wardBed:    row[7] && row[8] ? `Ward ${row[7]} - Bed ${row[8]}` : (row[7] ? `Ward ${row[7]}` : "Unassigned"),
        consultant: String(row[9] || "--").trim(),
        diagnosis:  String(row[10] || "Pending").trim()
      });
    }

    return {
      success: true,
      data: activeAdmissions.reverse(),
      scopeMode: gate.scope.mode,
      scopedTo: gate.scope.mode === 'ALL' ? '' : gate.scope.doctorIds.join(', ')
    };
  } catch (error) {
    return { success: false, message: "Error fetching roster: " + error.toString(), data: [] };
  }
}

// ── 2. FETCH FULL CLINICAL CONTEXT FOR PATIENT ────────────

/**
 * Fetches everything needed to populate the 3-panel detail view:
 * - Patient demographics (from IP_Admissions)
 * - Active diagnosis (from IP_CaseSheets_DB baseline)
 * - Latest vitals (from most recent NURSE note in IP_Timeline_DB)
 * - Active medications (from IP_Pharmacy_Queue, Action_Flag = ACTIVE)
 * - Running infusions (from IP_Pharmacy_Queue, Route = IV, Action_Flag = ACTIVE)
 * - Allergies (from IP_CaseSheets_DB)
 */
function getClinicalContext(ipNumber, sessionToken) {
  try {
    var gate = resolveIPRead_(sessionToken, ipNumber);
    if (!gate.ok) return { success: false, message: gate.message };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const result = {
      success: true,
      demographics: {},
      diagnosis: [],
      allergies: [],
      latestVitals: { bp: "--/--", pulse: "--", spo2: "--", temp: "--", recorded: null },
      activeMeds: [],
      runningIV: [],
      dueMeds: [],
      handoverAlert: null
    };

    // ── Demographics from IP_Admissions ──
    const admSheet = ss.getSheetByName('IP_Admissions');
    if (admSheet) {
      const admData = admSheet.getDataRange().getValues();
      for (let i = 1; i < admData.length; i++) {
        if (String(admData[i][0]).trim() === String(ipNumber).trim()) {
          result.demographics = {
            name: String(admData[i][2] || "").trim(),
            ageSex: String(admData[i][3] || "--").trim(),
            wardBed: admData[i][7] && admData[i][8] ? `Ward ${admData[i][7]} - Bed ${admData[i][8]}` : "--",
            consultant: String(admData[i][9] || "--").trim(),
            patientId: String(admData[i][1] || "").trim()
          };
          break;
        }
      }
    }

    // ── Baseline from IP_CaseSheets_DB ──
    const csSheet = ss.getSheetByName('IP_CaseSheets_DB');
    if (csSheet) {
      // Header-driven: the casesheet gained attribution columns in Phase 5,
      // and a positional read would silently drift the day one more is added.
      const csMap  = dc_headerMap_(csSheet);
      const csData = csSheet.getDataRange().getValues();
      const iDx    = csMap["Primary Diagnosis"];
      const iAdv   = csMap["Advice"];
      const iDoc   = csMap["Doctor's Name"];

      for (let i = csData.length - 1; i > 0; i--) {
        if (String(csData[i][1]).trim() !== String(ipNumber).trim()) continue;

        const dx = (iDx !== undefined && csData[i][iDx]) ? String(csData[i][iDx]).trim() : "";
        if (dx) result.diagnosis = dx.split(',').map(d => d.trim()).filter(Boolean);

        result.baseline = {
          admittingDoctor: (iDoc !== undefined) ? String(csData[i][iDoc] || "") : "",
          admissionAdvice: (iAdv !== undefined) ? String(csData[i][iAdv] || "") : ""
        };
        break;
      }
    }

    // ── Latest Vitals from most recent NURSE note ──
    const tlSheet = ss.getSheetByName('IP_Timeline_DB');
    if (tlSheet) {
      const tlData = tlSheet.getDataRange().getValues();
      for (let i = tlData.length - 1; i > 0; i--) {
        if (String(tlData[i][1]).trim() === String(ipNumber).trim() &&
            tlData[i][3] === 'NURSE') {
          try {
            const nd = JSON.parse(tlData[i][4] || "{}");
            if (nd.vitals) {
              result.latestVitals = {
                bp:    nd.vitals.bp    || "--/--",
                pulse: nd.vitals.pulse || "--",
                spo2:  nd.vitals.spo2  || "--",
                temp:  nd.vitals.temp  || "--",
                recorded: Utilities.formatDate(
                  new Date(tlData[i][0]),
                  Session.getScriptTimeZone(),
                  "hh:mm a"
                )
              };
              break;
            }
          } catch(e) {}
        }
      }

      // ── Handover Alert — most recent HANDOVER flag ──
      for (let i = tlData.length - 1; i > 0; i--) {
        if (String(tlData[i][1]).trim() === String(ipNumber).trim() &&
            tlData[i][7] && tlData[i][7].toString().includes('HANDOVER')) {
          try {
            const nd = JSON.parse(tlData[i][4] || "{}");
            result.handoverAlert = nd.handoverText || null;
          } catch(e) {}
          break;
        }
      }
    }

    // ── Active Meds & Running IV from IP_Pharmacy_Queue ──
    const pqSheet = ss.getSheetByName('IP_Pharmacy_Queue');
    if (pqSheet) {
      const pqData = pqSheet.getDataRange().getValues();
      const now = new Date();

      // The same drug can hold several ACTIVE rows: a double-save, or a
      // MODIFY whose retire step failed. Show one card per drug — the newest —
      // rather than repeating it down the panel.
      const seenDrug = {};
      for (let i = pqData.length - 1; i > 0; i--) {
        if (String(pqData[i][1]).trim() !== String(ipNumber).trim()) continue;
        const actionFlag = String(pqData[i][11] || "ACTIVE").toUpperCase();
        if (actionFlag !== 'ACTIVE') continue;

        const dedupeKey = _normDrug_(pqData[i][3]);
        if (dedupeKey && seenDrug[dedupeKey]) continue;
        if (dedupeKey) seenDrug[dedupeKey] = true;

        const route = String(pqData[i][6] || "").toUpperCase().trim();
        const drug = {
          queueId:      String(pqData[i][0]),
          drugName:     String(pqData[i][3] || ""),
          dose:         String(pqData[i][4] || ""),
          freq:         String(pqData[i][5] || ""),
          route:        String(pqData[i][6] || ""),
          instructions: String(pqData[i][7] || ""),
          orderedBy:    String(pqData[i][9] || "")
        };

        const isIV = route === 'IV' || route.startsWith('IV') || route === 'INFUSION' || route.includes('FLUID');
        if (isIV) {
          result.runningIV.push(drug);
        } else {
          result.activeMeds.push(drug);
        }
      }
    }

    // ── Allergies ──
    // result.allergies was initialised and then never filled, so the panel
    // always read "None on record" however much had been entered elsewhere.
    try {
      const alg = getPatientAllergies(result.demographics.patientId, sessionToken);
      result.allergies = (alg && alg.allergies) || [];
      result.allergiesRecorded = !!(alg && alg.recorded);
      result.allergiesRaw = (alg && alg.raw) || "";
    } catch (e) {
      result.allergies = [];
      result.allergiesRecorded = false;
    }

    // ── Care team (Phase 5) ──
    try {
      ipc_ensurePrimaryOnCareTeam_(ipNumber);
      const teamRes = getIPCareTeam(ipNumber, sessionToken);
      result.careTeam = teamRes.success ? teamRes.team : [];
      result.primaryDoctorId = ipc_primaryDoctorId_(ipNumber);
    } catch (e) {
      result.careTeam = [];
      result.primaryDoctorId = "";
    }

    return result;
  } catch (error) {
    return { success: false, message: "Error fetching clinical context: " + error.toString() };
  }
}

// ── 3. FETCH TIMELINE ────────────────────────────────────

/**
 * Fetches all timeline notes for a specific IP admission, newest first.
 */
function getIPTimeline(ipNumber, sessionToken) {
  try {
    var gate = resolveIPRead_(sessionToken, ipNumber);
    if (!gate.ok) return { success: false, message: gate.message, data: [] };

    const sheet = ipc_timelineSheet_();
    const m = dc_headerMap_(sheet);
    const data = sheet.getDataRange().getValues();
    const timeline = [];

    const at = function (row, header, fallbackIndex) {
      const idx = (m[header] === undefined) ? fallbackIndex : m[header];
      return (idx === undefined || idx < 0) ? "" : row[idx];
    };

    for (let i = data.length - 1; i > 0; i--) {
      if (String(data[i][1]).trim() !== String(ipNumber).trim()) continue;

      let noteData = {};
      try { noteData = JSON.parse(data[i][4] || "{}"); } catch(e) {}

      let tsFormatted = "--";
      try {
        tsFormatted = Utilities.formatDate(
          new Date(data[i][0]),
          Session.getScriptTimeZone(),
          "dd-MMM-yyyy hh:mm a"
        );
      } catch(e) {}

      timeline.push({
        noteId:    String(at(data[i], "Note_ID", 8) || "") ||
                   (data[i][0] ? String(data[i][0].getTime ? data[i][0].getTime() : data[i][0]) : "--"),
        timestamp: tsFormatted,
        rawTs:     data[i][0] ? new Date(data[i][0]).toISOString() : null,
        ipNumber:  String(data[i][1] || ""),
        patientId: String(data[i][2] || ""),
        roleType:  String(data[i][3] || ""),
        noteData:  noteData,
        author:    String(data[i][5] || ""),
        shift:     String(data[i][6] || ""),
        flags:     String(data[i][7] || ""),
        authorDoctorId: String(at(data[i], "Author_Doctor_ID", 9) || ""),
        // The signature stored at the moment of writing — never re-derived,
        // so a doctor's later rename cannot silently rewrite old notes.
        signature:      String(at(data[i], "Author_Signature_Snapshot", 10) || ""),
        authorUsername: String(at(data[i], "Author_Username", 11) || "")
      });
    }

    return { success: true, data: timeline };
  } catch (error) {
    return { success: false, message: "Error fetching timeline: " + error.toString(), data: [] };
  }
}

// ── 4. SAVE IP NOTE (CORE WRITE) ────────────────────────

/**
 * Universal note saver. Accepts any roleType:
 * DOCTOR | NURSE | CONSULTANT | PROCEDURE | QUICK | INVESTIGATION
 */
function saveIPNote(payload, sessionToken) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    payload = payload || {};

    // ---- Phase 6 gate: identity, care team, section-level RBAC ----------
    // Everything the client claimed about who is writing is discarded here.
    const w = resolveIPWrite_(sessionToken, payload.ipNumber,
                              payload.roleType, payload.noteData);
    if (!w.ok) return { success: false, message: w.message };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ipc_timelineSheet_();

    const roleType  = String(payload.roleType || "QUICK").trim().toUpperCase();
    const timestamp = new Date();
    const shift     = _getCurrentShift_();
    const flags     = String(payload.flags || "");
    const noteId    = _generateNoteId_("IPN");
    const noteData  = w.noteData;                    // filtered, not raw
    const noteDataString = JSON.stringify(noteData);

    // Header-driven write: the authorship columns were appended by the
    // migration and must never be addressed by a hard-coded index.
    const m = dc_headerMap_(sheet);
    const row = new Array(sheet.getLastColumn()).fill("");
    const put = function (header, value) {
      if (m[header] !== undefined) row[m[header]] = value;
    };
    put("Timestamp", timestamp);
    put("IP_Number", String(payload.ipNumber).trim().toUpperCase());
    put("Patient_ID", String(payload.patientId || "").trim().toUpperCase());
    put("Role_Type", roleType);
    put("Note_Data_JSON", noteDataString);
    put("Author", w.authorLabel);
    put("Shift", shift);
    put("Flags", flags);
    put("Note_ID", noteId);
    put("Author_Doctor_ID", w.doctorId);
    put("Author_Signature_Snapshot", w.signature);
    put("Author_Username", w.username);
    sheet.appendRow(row);
    dc_invalidate_("IP_Timeline_DB");

    // ── Side Effects by Note Type ──────────────────────
    // Gated on w.mayPrescribe, not on roleType alone: a nurse cannot reach
    // the pharmacy queue even if she posts a note labelled DOCTOR, because
    // resolveIPWrite_ would have rejected the note type outright, and a
    // doctor's nursing-style note has had its medOrders stripped already.

    // A) Prescribing note: sync med orders to IP_Pharmacy_Queue
    if (w.mayPrescribe && noteData.medOrders && noteData.medOrders.length) {
      _syncMedOrdersToPharmacyQueue_(
        ss, payload.ipNumber, payload.patientId,
        noteData.medOrders, w.authorLabel, timestamp, noteId, w.doctorId
      );
    }

    // B) Prescribing note: route investigation orders to Lab
    if (w.mayPrescribe && noteData.investigationOrders && noteData.investigationOrders.length) {
      _routeInvestigationOrders_(
        ss, payload.ipNumber, payload.patientId,
        noteData.investigationOrders, w.authorLabel, timestamp, w.doctorId
      );
    }

    // C) Nurse note: mark administered meds. Scoped to THIS admission so a
    //    guessed queue id cannot touch another patient's chart.
    if (roleType === 'NURSE' && w.role === 'nurse' &&
        noteData.markedMeds && noteData.markedMeds.length) {
      _markMedsAdministered_(ss, noteData.markedMeds, w.authorLabel, timestamp,
                             payload.ipNumber);
    }

    try {
      logAudit_(w.sess, "IP_NOTE_SAVE", "IP_Timeline", noteId, {
        ipNumber: payload.ipNumber, roleType: roleType,
        doctorId: w.doctorId, stripped: w.stripped
      });
    } catch (e) { /* auditing must never fail a clinical save */ }

    SpreadsheetApp.flush();

    var msg = "Note saved to clinical timeline.";
    if (w.stripped.length) {
      msg += " The following were not recorded because your role cannot author them: " +
             w.stripped.join(", ") + ".";
    }
    return {
      success: true,
      message: msg,
      noteId: noteId,
      author: w.authorLabel,
      signature: w.signature,
      stripped: w.stripped
    };
  } catch (error) {
    return { success: false, message: "Failed to save note: " + error.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ── 5. SYNC MED ORDERS TO PHARMACY QUEUE ─────────────────

/**
 * Called when a Doctor note is saved.
 * action = 'NEW' → append | 'CONT' → no-op | 'STOP'/'HOLD' → flag
 * 'MODIFY' → mark old MODIFIED, append new ACTIVE
 */
function _syncMedOrdersToPharmacyQueue_(ss, ipNumber, patientId, medOrders, author, timestamp, noteId, doctorId) {
  const sheet = _ensureIPPharmacyQueueSheet_(ss);
  const ip = String(ipNumber).trim().toUpperCase();

  // Row-level edits below change Action_Flag, so the snapshot has to be
  // re-read rather than reused: two STOP orders for the same drug in one
  // note would otherwise both land on the same row.
  const readRows = function () { return sheet.getDataRange().getValues(); };

  const appendOrder = function (order) {
    sheet.appendRow([
      _generateNoteId_("IPQ"),
      ip,
      String(patientId || "").trim().toUpperCase(),
      String(order.drugName || "").trim(),
      String(order.dose || "").trim(),
      String(order.freq || "").trim(),
      String(order.route || "Oral").trim(),
      String(order.instructions || "").trim(),
      "Pending_Dispense",
      String(author).trim(),
      timestamp,
      "ACTIVE",
      "", "",
      // Encounter_Note_ID ties the queue row back to the note that ordered
      // it, so pharmacy can always show who signed for a drug.
      String(noteId || "")
    ]);
  };

  /** Newest ACTIVE row for this drug on THIS admission, or -1. */
  const findActive = function (drugName) {
    const data = readRows();
    const needle = _normDrug_(drugName);
    if (!needle) return -1;
    for (let i = data.length - 1; i > 0; i--) {
      if (String(data[i][1]).trim().toUpperCase() !== ip) continue;
      if (_normDrug_(data[i][3]) !== needle) continue;
      if (String(data[i][11]).toUpperCase() !== 'ACTIVE') continue;
      return i + 1;
    }
    return -1;
  };

  const stamp = function (rowNumber, flag) {
    sheet.getRange(rowNumber, 12).setValue(flag);
    sheet.getRange(rowNumber, 13).setValue(timestamp);
    sheet.getRange(rowNumber, 14).setValue(String(author));
  };

  medOrders.forEach(function(order) {
    const action = String(order.action || "NEW").toUpperCase();

    if (action === 'NEW') {
      appendOrder(order);
    } else if (action === 'STOP' || action === 'HOLD') {
      // Every active row for this drug, not just the newest: a duplicate left
      // by an earlier double-save would otherwise stay on the Continue list
      // after the doctor had stopped the drug.
      let r, guard = 0;
      while ((r = findActive(order.drugName)) > 0 && guard++ < 20) {
        stamp(r, action === 'STOP' ? 'STOPPED' : 'HOLD');
      }
    } else if (action === 'MODIFY') {
      const r = findActive(order.drugName);
      if (r > 0) stamp(r, 'MODIFIED');
      appendOrder(order);
    }
    // CONT → no write needed
  });
}

// ── 6. ROUTE INVESTIGATION ORDERS TO LAB ──────────────────

function _routeInvestigationOrders_(ss, ipNumber, patientId, orders, author, timestamp, doctorId) {
  // Bridge to new Lab Integration Engine — routes via LAB_ORDERS schema
  if (!orders || !orders.length) return;
  try {
    const testNames = orders.map(function(o){ return String(o.testName||'').trim(); }).filter(Boolean);
    const testIds   = orders.map(function(o){ return String(o.testId||'').trim(); }).filter(Boolean);
    const hasStat   = orders.some(function(o){ return String(o.priority||'').toUpperCase() === 'STAT'; });
    if (!testNames.length && !testIds.length) return;
    createLabRequest({
      patientId:          String(patientId).trim(),
      admissionId:        String(ipNumber).trim(),
      sourceModule:       'IP_NOTES',
      // Real catalog IDs when the picker supplied them; names only as a
      // fallback for a test typed by hand.
      testIds:            testIds,
      testNames:          testNames,
      priority:           hasStat ? 'STAT' : 'ROUTINE',
      orderingDoctorName: String(author || ''),
      orderingDoctorId:   String(doctorId || '')
    });
  } catch (e) {
    Logger.log('IP Notes → Lab bridge failed: ' + e.message);
    // Fallback: write to legacy sheet so no order is lost
    try {
      const sheet = _ensureLabQueueSheet_(ss);
      orders.forEach(function(order) {
        sheet.appendRow([
          _generateNoteId_("IPL"), '', String(patientId).trim(), String(ipNumber).trim(),
          timestamp, String(order.testName||'').trim(),
          String(order.priority||'ROUTINE').toUpperCase(), 'Pending_IP_Sample', '', ''
        ]);
      });
    } catch(e2) { Logger.log('IP Notes legacy fallback also failed: ' + e2.message); }
  }
}

// ── 7. MARK MEDS ADMINISTERED (NURSE NOTE) ──────────────

function _markMedsAdministered_(ss, markedMeds, nurse, timestamp, ipNumber) {
  const sheet = ss.getSheetByName('IP_Pharmacy_Queue');
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const ip = String(ipNumber || "").trim().toUpperCase();

  markedMeds.forEach(function(med) {
    if (!med.given) return;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() !== String(med.queueId).trim()) continue;
      // A queue id alone is not authority: the row must belong to the
      // admission the nurse is charting on.
      if (ip && String(data[i][1]).trim().toUpperCase() !== ip) break;
      sheet.getRange(i + 1, 9).setValue("Dispensed_IP");
      sheet.getRange(i + 1, 13).setValue(timestamp);
      sheet.getRange(i + 1, 14).setValue(String(nurse));
      break;
    }
  });
}

// ── 8. UPDATE MED ORDER ACTION (standalone) ──────────────

/**
 * Allows a doctor to change a single med's action flag directly.
 */
function updateIPMedAction(payload, sessionToken) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    payload = payload || {};
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('IP_Pharmacy_Queue');
    if (!sheet) return { success: false, message: "IP_Pharmacy_Queue not found." };

    const data = sheet.getDataRange().getValues();
    const newAction = String(payload.newAction || "ACTIVE").toUpperCase();
    const timestamp = new Date();
    let found = false;

    // Locate the row first so the admission it belongs to — not the client —
    // decides which care team must authorise the change.
    let targetRow = -1;
    for (let k = 1; k < data.length; k++) {
      if (String(data[k][0]).trim() === String(payload.queueId).trim()) { targetRow = k; break; }
    }
    if (targetRow === -1) return { success: false, message: "Queue entry not found." };

    // Changing a drug order is a prescribing act: it goes through the same
    // gate as writing the note that would have ordered it.
    const w = resolveIPWrite_(sessionToken, data[targetRow][1], "DOCTOR", {});
    if (!w.ok) return { success: false, message: w.message };
    const actor = w.authorLabel;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(payload.queueId).trim()) {
        if (newAction === 'MODIFY') {
          sheet.getRange(i + 1, 12).setValue('MODIFIED');
          sheet.getRange(i + 1, 13).setValue(timestamp);
          sheet.getRange(i + 1, 14).setValue(actor);
          const newId = _generateNoteId_("IPQ");
          sheet.appendRow([
            newId, data[i][1], data[i][2],
            String(payload.newDrugName     || data[i][3]),
            String(payload.newDose         || data[i][4]),
            String(payload.newFreq         || data[i][5]),
            String(payload.newRoute        || data[i][6]),
            String(payload.newInstructions || data[i][7]),
            "Pending_Dispense",
            actor,
            timestamp, "ACTIVE", "", "", String(data[i][14] || "")
          ]);
        } else {
          sheet.getRange(i + 1, 12).setValue(newAction);
          sheet.getRange(i + 1, 13).setValue(timestamp);
          sheet.getRange(i + 1, 14).setValue(actor);
        }
        found = true;
        break;
      }
    }

    if (!found) return { success: false, message: "Queue entry not found." };
    try {
      logAudit_(w.sess, "IP_MED_ACTION", "IP_Pharmacy_Queue",
                String(payload.queueId), { newAction: newAction, doctorId: w.doctorId });
    } catch (e) { /* never fail the order on audit */ }
    SpreadsheetApp.flush();
    return { success: true, message: `Medication status updated to ${newAction}.` };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ── 9. GENERATE HANDOVER SUMMARY ─────────────────────────

/**
 * Generates a structured shift handover summary note.
 */
function generateIPHandoverSummary(payload, sessionToken) {
  // No lock here: this function only reads, then delegates the single write
  // to saveIPNote(), which takes the script lock itself. Taking it twice in
  // one execution would deadlock.
  try {
    payload = payload || {};
    const w = resolveIPWrite_(sessionToken, payload.ipNumber, "HANDOVER", {});
    if (!w.ok) return { success: false, message: w.message };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tlSheet = ss.getSheetByName('IP_Timeline_DB');
    const pqSheet = ss.getSheetByName('IP_Pharmacy_Queue');

    const cutoffMs = 12 * 60 * 60 * 1000;
    const now = new Date();

    let vitalsList = [];
    let doctorNotes = [];
    let nurseNotes = [];
    let alerts = [];
    let events = [];        // diagnosis revisions, amendments, procedures
    let consults = [];

    if (tlSheet) {
      const tlData = tlSheet.getDataRange().getValues();
      const tlMap  = dc_headerMap_(tlSheet);
      const iFlags = (tlMap["Flags"] === undefined) ? 7 : tlMap["Flags"];

      for (let i = 1; i < tlData.length; i++) {
        if (String(tlData[i][1]).trim() !== String(payload.ipNumber).trim()) continue;
        const rowTs = new Date(tlData[i][0]);
        if ((now - rowTs) > cutoffMs) continue;

        const roleType = String(tlData[i][3] || "");
        const author   = String(tlData[i][5] || "");
        const flags    = String(tlData[i][iFlags] || "");
        let nd = {};
        try { nd = JSON.parse(tlData[i][4] || "{}"); } catch(e) {}

        const timeStr = Utilities.formatDate(rowTs, Session.getScriptTimeZone(), "hh:mm a");

        if (roleType === 'NURSE' && nd.vitals) {
          vitalsList.push(`${timeStr}: BP ${nd.vitals.bp || "--"} | P ${nd.vitals.pulse || "--"} | SpO2 ${nd.vitals.spo2 || "--"}% | T ${nd.vitals.temp || "--"}`);
        }

        // A doctor note is more than its subjective line. Recording only
        // nd.subjectiveObjective is why a diagnosis revised at 20:27 never
        // reached the handover — that change lives in assessment/diagnosis.
        if (roleType === 'DOCTOR') {
          const body = [nd.subjectiveObjective, nd.assessment, nd.adviceText]
            .map(function (x) { return String(x || "").trim(); })
            .filter(Boolean).join(" — ");
          if (body) doctorNotes.push(`${timeStr} [${author}]: ${body}`);
        }

        if (roleType === 'NURSE' && nd.observations) {
          nurseNotes.push(`${timeStr} [${author}]: ${nd.observations}`);
        }

        if (roleType === 'CONSULTANT') {
          consults.push(`${timeStr} [${author}] ${nd.specialty ? nd.specialty + ": " : ""}` +
                        `${String(nd.recommendations || nd.findings || "opinion recorded").trim()}`);
        }

        if (roleType === 'PROCEDURE') {
          events.push(`${timeStr} Procedure — ${String(nd.procedureName || "unnamed").trim()}` +
                      `${nd.complications ? " (complications: " + nd.complications + ")" : ""} [${author}]`);
        }

        // Flagged clinical events: a diagnosis revision or a casesheet
        // amendment is exactly what the incoming shift needs to know.
        if (flags.indexOf('DIAGNOSIS') !== -1) {
          events.push(`${timeStr} Diagnosis revised — ${String(nd.diagnosis || nd.assessment || "").trim()} [${author}]`);
        }
        if (flags.indexOf('CASESHEET_AMENDED') !== -1) {
          events.push(`${timeStr} Casesheet amended — ${String(nd.plan || "").trim()} [${author}]`);
        }
        if (flags.indexOf('ALERT') !== -1) {
          alerts.push(`\u26A0 ${timeStr} ${nd.alertText || "Clinical alert triggered."} [${author}]`);
        }
      }
    }

    // Current working diagnosis, so the handover states where things stand and
    // not merely what changed.
    let currentDx = "";
    try {
      const adm = ipc_admissionRow_(payload.ipNumber);
      if (adm) currentDx = dc_str_(adm.row[10]);
    } catch (e) { /* the summary is still worth producing without it */ }

    const summaryLines = [
      `=== ${payload.shift || _getCurrentShift_()} SHIFT HANDOVER SUMMARY ===`,
      `Generated: ${Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MMM-yyyy hh:mm a")}`,
      `Generated by: ${w.authorLabel}`,
      "",
      "WORKING DIAGNOSIS:",
      currentDx || "Not recorded.",
      "",
      "CHANGES THIS PERIOD:",
      events.length ? events.join("\n") : "No diagnosis changes, procedures or amendments.",
      "",
      "ALERTS:",
      alerts.length ? alerts.join("\n") : "No alerts in this period.",
      "",
      "VITALS TREND (Last 12h):",
      vitalsList.length ? vitalsList.join("\n") : "No vitals recorded.",
      "",
      "ACTIVE MEDICATIONS:",
      activeMedsList.length ? activeMedsList.join("\n") : "No active medications.",
      "",
      "DOCTOR NOTES:",
      doctorNotes.length ? doctorNotes.join("\n") : "No doctor notes.",
      "",
      "CONSULTANT OPINIONS:",
      consults.length ? consults.join("\n") : "No consultant opinions this period.",
      "",
      "NURSING OBSERVATIONS:",
      nurseNotes.length ? nurseNotes.join("\n") : "No nursing notes."
    ];
    const handoverText = summaryLines.join("\n");

    // Route the handover through the same authored-write path as any other
    // note so it carries a Note_ID and a real signature.
    const saved = saveIPNote({
      ipNumber:  payload.ipNumber,
      patientId: payload.patientId,
      roleType:  "HANDOVER",
      flags:     "HANDOVER",
      noteData:  { handoverText: handoverText }
    }, sessionToken);
    if (!saved.success) return { success: false, message: saved.message };

    return { success: true, summary: handoverText,
             noteId: saved.noteId, author: w.authorLabel };
  } catch (error) {
    return { success: false, message: "Handover generation failed: " + error.toString() };
  }
}

// ── 10. FETCH IP PHARMACY QUEUE FOR A PATIENT ──────────

/**
 * For the Pharmacy module: fetches all ACTIVE/Pending_Dispense IP orders.
 */
function getIPPharmacyQueue(ipNumber, sessionToken) {
  try {
    var gate = resolveIPRead_(sessionToken, ipNumber);
    if (!gate.ok) return { success: false, message: gate.message, data: [] };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('IP_Pharmacy_Queue');
    if (!sheet) return { success: true, data: [] };

    const data = sheet.getDataRange().getValues();
    const results = [];

    for (let i = 1; i < data.length; i++) {
      const ipFilter = ipNumber ? String(data[i][1]).trim() === String(ipNumber).trim() : true;
      if (!ipFilter) continue;

      const actionFlag = String(data[i][11] || "ACTIVE").toUpperCase();
      if (actionFlag !== 'ACTIVE') continue;

      let orderedAtFormatted = "--";
      try {
        orderedAtFormatted = Utilities.formatDate(
          new Date(data[i][10]),
          Session.getScriptTimeZone(),
          "dd-MMM hh:mm a"
        );
      } catch(e) {}

      results.push({
        queueId:      String(data[i][0]),
        ipNumber:     String(data[i][1]),
        patientId:    String(data[i][2]),
        drugName:     String(data[i][3]),
        dose:         String(data[i][4]),
        freq:         String(data[i][5]),
        route:        String(data[i][6]),
        instructions: String(data[i][7]),
        status:       String(data[i][8]),
        orderedBy:    String(data[i][9]),
        orderedAt:    orderedAtFormatted,
        actionFlag:   actionFlag
      });
    }

    return { success: true, data: results };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ── 11. FETCH STAFF ROSTER FOR AUTHOR DROPDOWNS ──────────

/**
 * Returns all active staff from Users sheet, grouped by role.
 */
function getStaffRoster(sessionToken) {
  try {
    // The staff list is directory data, not public data — it names every
    // clinician and their registration number.
    var gate = resolveIPRead_(sessionToken, null);
    if (!gate.ok) return { success: false, message: gate.message, data: [] };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Users');
    if (!sheet) return { success: false, data: [] };

    const data = sheet.getDataRange().getValues();
    const staff = [];

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const username    = String(data[i][0] || "").trim();
      const role        = String(data[i][2] || "").trim().toLowerCase();
      const isActive    = String(data[i][3] || "active").trim().toLowerCase();
      const displayName = data[i][4] ? String(data[i][4]).trim() : username;
      const regNo       = data[i][5] ? String(data[i][5]).trim() : "";

      if (isActive !== 'active' && isActive !== '') continue;

      let label = displayName;
      if (role === 'doctor') {
        label = 'Dr. ' + displayName + (regNo ? ' (Reg: ' + regNo + ')' : '');
      } else if (role === 'nurse') {
        label = 'Staff Nurse ' + displayName + (regNo ? ' (' + regNo + ')' : '');
      } else if (role === 'admin') {
        label = displayName + ' (Admin)';
      }

      staff.push({ username, displayName, role, regNo, label });
    }

    return { success: true, data: staff };
  } catch (error) {
    return { success: false, message: error.toString(), data: [] };
  }
}

// ── 12. FETCH PHARMACY MASTER FOR AUTOCOMPLETE ────────────

function fetchPharmacyMasterForIP() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Pharmacy_Inventory");
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const drugs = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][1]) continue;
      drugs.push({
        brand:   String(data[i][1] || ""),
        generic: String(data[i][2] || ""),
        type:    String(data[i][3] || ""),
        unit:    String(data[i][5] || ""),
        stock:   parseInt(data[i][4]) || 0
      });
    }
    return drugs;
  } catch(e) {
    return [];
  }
}

// ── 13. FETCH LAB RESULTS FOR PATIENT (IP context) ───────

function getIPLabResults(ipNumber, patientId, sessionToken) {
  try {
    var gate = resolveIPRead_(sessionToken, ipNumber);
    if (!gate.ok) return { success: false, message: gate.message, data: [] };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Lab_Queue_DB');
    if (!sheet) return { success: true, data: [] };

    const data = sheet.getDataRange().getValues();
    const results = [];

    for (let i = 1; i < data.length; i++) {
      const rowIP  = String(data[i][3] || "").trim();
      const rowPID = String(data[i][2] || "").trim();
      if (rowIP !== String(ipNumber).trim() && rowPID !== String(patientId).trim()) continue;

      let orderedAtFmt = "--";
      try { orderedAtFmt = Utilities.formatDate(new Date(data[i][4]), Session.getScriptTimeZone(), "dd-MMM hh:mm a"); } catch(e) {}

      let resultData = {};
      try { resultData = JSON.parse(data[i][8] || "{}"); } catch(e) {}

      results.push({
        orderId:    String(data[i][0]),
        orderedAt:  orderedAtFmt,
        testName:   String(data[i][5] || ""),
        priority:   String(data[i][6] || ""),
        status:     String(data[i][7] || ""),
        result:     resultData,
        reportedAt: String(data[i][9] || "--")
      });
    }

    results.sort((a, b) => b.orderId.localeCompare(a.orderId));
    return { success: true, data: results };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
// ── 14. PRINTABLE PROGRESS RECORD ─────────────────────────

/**
 * Print-ready HTML for an admission's notes.
 * @param {string} ipNumber
 * @param {Object} opts  { from, to } ISO dates, or blank for the whole stay
 */
/**
 * Print-ready HTML for an admission's progress record.
 *
 * Composed with IP_Print_Kit so the sheet matches the casesheet typographically
 * and lines up column-for-column. Every section a note can carry is rendered
 * through an aligned label/value grid rather than a run of inline <p> tags —
 * that is what makes a stack of ward-round notes readable on paper.
 *
 * @param {string} ipNumber
 * @param {Object} opts  { from, to } ISO dates for a date-bounded extract,
 *                       { roleTypes: [...] } to print only certain note kinds,
 *                       { trend: false } to suppress the vitals chart
 */
function getIPNotesPrintHtml(ipNumber, opts, sessionToken) {
  try {
    var gate = resolveIPRead_(sessionToken, ipNumber);
    if (!gate.ok) return { success: false, message: gate.message };

    opts = opts || {};
    var tl = getIPTimeline(ipNumber, sessionToken);
    if (!tl.success) return { success: false, message: tl.message };

    var adm = ipc_admissionRow_(ipNumber);

    var from = opts.from ? new Date(opts.from) : null;
    var to   = opts.to   ? new Date(opts.to)   : null;
    if (to) to.setHours(23, 59, 59, 999);

    var wanted = null;
    if (opts.roleTypes && opts.roleTypes.length) {
      wanted = {};
      opts.roleTypes.forEach(function (r) { wanted[dc_upper_(r)] = true; });
    }

    // A printed record reads chronologically. Sorting on the timestamp rather
    // than reversing the sheet order matters: a note entered late, or a row
    // written during an amendment, sits out of sequence in the sheet and would
    // otherwise print out of sequence too.
    var rows = tl.data.filter(function (n) {
      if (wanted && !wanted[dc_upper_(n.roleType)]) return false;
      if (!n.rawTs) return !from && !to;
      var d = new Date(n.rawTs);
      if (from && d < from) return false;
      if (to && d > to) return false;
      return true;
    }).sort(ipn_byTimeAsc_);

    if (!rows.length) {
      return { success: false, message: "No notes in that period to print." };
    }

    var patient = ipn_printPatient_(ipNumber, adm);
    var body = ipp_sec_("Observation Chart", ipn_trendSection_(rows, opts), { loose: true }) +
               ipp_sec_("Progress Notes",
                        rows.map(ipn_printNote_).join(""),
                        { loose: true });

    var span = (from || to)
      ? "Extract: " + (from ? ipp_when_(from, "dd-MMM-yyyy") : "admission") +
        " to " + (to ? ipp_when_(to, "dd-MMM-yyyy") : "today")
      : "Whole stay";

    var html = ipp_doc_({
      docTitle: "Inpatient Progress Record",
      patient:  patient,
      bodyHtml: body + ipp_sig_(patient.consultant || "Consultant", "Treating Consultant"),
      footNote: rows.length + " note(s) · " + span
    });

    return { success: true, html: html, noteCount: rows.length };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/** Oldest first; rows with no usable timestamp keep their relative order last. */
function ipn_byTimeAsc_(a, b) {
  var ta = a.rawTs ? new Date(a.rawTs).getTime() : NaN;
  var tb = b.rawTs ? new Date(b.rawTs).getTime() : NaN;
  if (isNaN(ta) && isNaN(tb)) return 0;
  if (isNaN(ta)) return 1;
  if (isNaN(tb)) return -1;
  return ta - tb;
}

/**
 * Identity block for an admission, shared by the notes and full-file printers.
 * IP_Admissions is the frozen 13-column schema documented in
 * IP_Admissions_Logic.gs: [1]Patient_ID [2]Patient_Name [3]Age_Sex
 * [7]Ward_Bed [8]Bed [9]Consultant [10]Diagnosis.
 */
function ipn_printPatient_(ipNumber, adm) {
  var r = adm ? adm.row : [];
  // Historic rows wrote "B - B201" into Ward_Bed and left Bed stale; the
  // admissions module's own resolver is the authority on which is real.
  var wb = ipa_resolveWardBed_(r[7], r[8]);
  return {
    name:       dc_str_(r[2]),
    pid:        dc_str_(r[1]),
    ipNumber:   dc_str_(ipNumber),
    ageSex:     dc_str_(r[3]),
    wardBed:    [wb.ward, wb.bed].filter(function (x) { return !!x; }).join(" / "),
    consultant: dc_str_(r[9]),
    diagnosis:  dc_str_(r[10])
  };
}

/** The vitals chart, built from whichever notes carry a vitals block. */
function ipn_trendSection_(rows, opts) {
  if (opts && opts.trend === false) return "";
  var series = rows.filter(function (n) {
    return n.noteData && n.noteData.vitals;
  }).map(function (n) {
    return { ts: n.rawTs, label: n.timestamp, vitals: n.noteData.vitals };
  });
  if (series.length < 2) return "";   // one reading is not a trend
  return ipp_vitalsTrend_(series);
}

var IPN_PRINT_LABELS = {
  DOCTOR: "Clinical Progress Note", NURSE: "Nursing Note",
  CONSULTANT: "Consultant Opinion", PROCEDURE: "Procedure Note",
  HANDOVER: "Shift Handover", QUICK: "Quick Note",
  INVESTIGATION: "Investigation"
};

/** One timeline entry as an aligned, unbreakable block. */
function ipn_printNote_(n) {
  var d = n.noteData || {};
  var e = ipp_esc_;
  var kv = [];
  var push = function (label, valueHtml) {
    var v = String(valueHtml === null || valueHtml === undefined ? "" : valueHtml).trim();
    if (v) kv.push([label, v]);
  };

  push("Subjective / Objective", ipp_escMultiline_(d.subjectiveObjective));

  if (d.vitals) {
    var v = d.vitals;
    // Only the parameters actually recorded. The old printer emitted
    // "BP -- | Pulse -- | SpO2 --% | Temp --" for a note with no vitals at all.
    var vit = [
      dc_str_(v.bp)    ? "BP " + e(v.bp) + " mmHg"   : "",
      dc_str_(v.pulse) ? "Pulse " + e(v.pulse) + " bpm" : "",
      dc_str_(v.spo2)  ? "SpO2 " + e(v.spo2) + "%"   : "",
      dc_str_(v.temp)  ? "Temp " + e(v.temp) + " °F" : ""
    ].filter(Boolean).join(" &nbsp;&middot;&nbsp; ");
    push("Vitals", vit);
  }

  // Systemic examination. Older rows may still hold the pre-canonical key.
  var se = d.sysExam || d.systemExam;
  if (se) {
    var exam = [
      dc_str_(se.cvs) ? ["CVS", se.cvs] : null,
      dc_str_(se.rs)  ? ["RS",  se.rs]  : null,
      dc_str_(se.pa)  ? ["P/A", se.pa]  : null,
      dc_str_(se.cns) ? ["CNS", se.cns] : null
    ].filter(Boolean);
    if (exam.length) {
      // Each system on its own line with a fixed-width gutter, so CVS/RS/P-A/CNS
      // read down the page instead of running together on one wrapped line.
      push("Systemic Exam",
        '<table class="kv narrow" style="margin:-2px 0;"><tbody>' +
        exam.map(function (x) {
          return '<tr><th style="width:16mm;font-weight:600;">' + e(x[0]) +
                 '</th><td>' + e(x[1]) + '</td></tr>';
        }).join("") + '</tbody></table>');
    }
  }

  push("Assessment",      ipp_escMultiline_(d.assessment));
  push("Diagnosis",       ipp_escMultiline_(d.diagnosis));
  push("Plan",            ipp_escMultiline_(d.plan));
  push("Advice",          ipp_escMultiline_(d.adviceText));
  push("Intervention",    ipp_escMultiline_(d.intervention));
  push("Observations",    ipp_escMultiline_(d.observations));
  push("Consultant",      e(d.consultantName));
  push("Specialty",       e(d.specialty));
  push("Findings",        ipp_escMultiline_(d.findings));
  push("Recommendations", ipp_escMultiline_(d.recommendations));
  push("Procedure",       e(d.procedureName));
  push("Operator",        e(d.operator));
  push("Anaesthesia",     e(d.anaesthesia));
  push("Complications",   ipp_escMultiline_(d.complications));
  push("Note",            ipp_escMultiline_(d.text));
  push("Alert",           ipp_escMultiline_(d.alertText));

  if (d.intakeOutput && (dc_str_(d.intakeOutput.intake) || dc_str_(d.intakeOutput.output))) {
    push("Intake / Output",
      "In " + e(d.intakeOutput.intake || "--") + " &nbsp;&middot;&nbsp; Out " +
      e(d.intakeOutput.output || "--"));
  }

  var blocks = [];
  if (kv.length) blocks.push(ipp_kv_(kv));

  if (d.medOrders && d.medOrders.length) {
    blocks.push('<div style="height:5px;"></div>' + ipp_table_(
      [{ label: "Action", cls: "ctr" }, "Drug", "Dose / Frequency", "Route"],
      d.medOrders.map(function (m) {
        return [
          e(m.action || "NEW"),
          e(m.drugName || m.newDrugName || ""),
          [e(m.dose || ""), e(m.freq || "")].filter(Boolean).join(" "),
          e(m.route || "")
        ];
      }),
      ["20mm", "auto", "42mm", "22mm"]
    ));
  }

  if (d.investigationOrders && d.investigationOrders.length) {
    var tests = d.investigationOrders.map(function (o) {
      var name = dc_str_(o.testName);
      if (!name) return "";
      var prio = dc_upper_(o.priority);
      return e(name) + (prio && prio !== "ROUTINE" ? ' <strong>(' + e(prio) + ')</strong>' : "");
    }).filter(Boolean);
    if (tests.length) {
      blocks.push('<div style="height:5px;"></div>' +
                  ipp_kv_([["Investigations", tests.join(", ")]]));
    }
  }

  if (d.markedMeds && d.markedMeds.length) {
    var given = d.markedMeds.filter(function (m) { return m.given; })
                            .map(function (m) { return e(m.drugName); }).filter(Boolean);
    if (given.length) {
      blocks.push('<div style="height:5px;"></div>' +
                  ipp_kv_([["Medication Given", given.join(", ")]]));
    }
  }

  if (dc_str_(d.handoverText)) {
    blocks.push('<div style="white-space:pre-wrap;overflow-wrap:anywhere;margin-top:3px;">' +
                ipp_esc_(d.handoverText) + '</div>');
  }

  if (!blocks.length) {
    blocks.push('<span class="muted">No content recorded.</span>');
  }

  var author = dc_str_(n.author);
  var sig    = dc_str_(n.signature);

  return '<div class="note">' +
    '<div class="nh clearfix">' +
      '<span><span class="tag">' + ipp_esc_(IPN_PRINT_LABELS[dc_upper_(n.roleType)] || n.roleType) + '</span>' +
        ipp_esc_(author) + (sig && sig !== author ? ' <em>(' + ipp_esc_(sig) + ')</em>' : '') + '</span>' +
      '<span class="r">' + ipp_esc_(n.timestamp) +
        (dc_str_(n.shift) ? ' &middot; ' + ipp_esc_(n.shift) + ' shift' : '') + '</span>' +
    '</div>' + blocks.join("") +
  '</div>';
}

// ── 15. LAB CATALOGUE FOR IP NOTES ────────────────────────

/**
 * The same catalogue the OP module and the IP casesheet order from, so an
 * investigation ordered on a ward round carries a real catalog ID and is
 * billable and resultable. IP Notes previously took free-typed test names,
 * which the lab module could neither price nor match to a panel.
 */
function getIPNotesLabCatalog(sessionToken) {
  try {
    var gate = resolveIPRead_(sessionToken, null);
    if (!gate.ok) return { success: false, message: gate.message, panels: [], tests: [], packages: [] };
    return getOPDOrderableTests();
  } catch (e) {
    return { success: false, message: e.message, panels: [], tests: [], packages: [] };
  }
}
