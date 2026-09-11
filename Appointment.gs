var APPT_ALLOWED_STATUSES = ['Booked', 'Arrived', 'In-Progress', 'Completed', 'Cancelled'];
var APPT_STATUS_WRITERS   = ['admin', 'doctor', 'receptionist', 'reception', 'nurse'];

// ==========================================
// 🚀 APPOINTMENT MODULE ENGINE
// ==========================================

// Helper to prevent Google Sheets from corrupting Time formats into hidden Date objects
function formatTimeSafely(timeVal) {
  if(!timeVal) return "";
  if(timeVal instanceof Date) {
    return Utilities.formatDate(timeVal, Session.getScriptTimeZone(), "hh:mm a").toUpperCase();
  }
  let t = String(timeVal).trim().toUpperCase();
  t = t.replace(/([0-9])(AM|PM)/, "$1 $2"); // Ensures space before AM/PM
  return t;
}

// 1. DOCTOR AVAILABILITY ENGINE
function getAvailableTimeSlots(dateStr) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Appointments');
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const standardSlots = [
    "10:00 AM", "10:15 AM", "10:30 AM", "10:45 AM", "11:00 AM", "11:15 AM", "11:30 AM", "11:45 AM",
    "12:00 PM", "12:15 PM", "12:30 PM", "12:45 PM", "05:00 PM", "05:15 PM", "05:30 PM", "05:45 PM",
    "06:00 PM", "06:15 PM", "06:30 PM", "06:45 PM", "07:00 PM", "07:15 PM", "07:30 PM", "07:45 PM",
    "08:00 PM", "08:15 PM", "08:30 PM", "08:45 PM"
  ];

  const takenSlots = [];
  for (let i = 1; i < data.length; i++) {
    let dObj = data[i][3];
    let rowDate = (dObj instanceof Date) ? Utilities.formatDate(dObj, Session.getScriptTimeZone(), "yyyy-MM-dd") : dObj.toString().substring(0,10);

    if (rowDate === dateStr && data[i][6] !== 'Cancelled' && data[i][6] !== 'DELETE') {
      takenSlots.push(formatTimeSafely(data[i][4]));
    }
  }

  return standardSlots.filter(slot => !takenSlots.includes(slot));
}

// Retrieves accurate schedule array for the Frontend Checkboxes
function getAppointmentsByDate(dateStr) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Appointments');
    if(!sheet) return [];
    const data = sheet.getDataRange().getValues();
    let appts = [];
    for(let i = 1; i < data.length; i++) {
      let dObj = data[i][3];
      let rowDate = "";
      if(dObj instanceof Date) rowDate = Utilities.formatDate(dObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
      else if(dObj) rowDate = dObj.toString().substring(0,10);

      if(rowDate === dateStr) {
        // BUG 2 FIX: formatTimeSafely forces identical strings so the UI toggle arrays match perfectly
        appts.push({ 
            apptId: data[i][0], 
            patientId: data[i][1], 
            patientName: data[i][2], 
            time: formatTimeSafely(data[i][4]), 
            purpose: data[i][5], 
            status: data[i][6], 
            fee: data[i][7] 
        });
      }
    }
    return appts;
  } catch(e) {
    return [];
  }
}

// 2. PATIENT AUTO-FETCH DEMOGRAPHICS
function getPatientDemographics(patientId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Patients');
  if (!sheet || !patientId) return null;
  const data = sheet.getDataRange().getValues();
  const want = patientId.toString().trim().toUpperCase();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toUpperCase() === want) {
      // gender + sex both returned: the booking modal reads gender, the ledger reads sex
      return { id: data[i][0], name: data[i][2], age: data[i][3], sex: data[i][4], gender: data[i][4], mobile: data[i][6] || "" };
    }
  }
  return null;
}

// 3. ADMIN DAILY LEDGER
function fetchDailyLedger(dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const apptSheet = ss.getSheetByName('Appointments');
  const patientSheet = ss.getSheetByName('Patients');
  if (!apptSheet || !patientSheet) return [];
  const apptData = apptSheet.getDataRange().getValues();

  // Collect the ids this date actually needs BEFORE touching the Patients sheet,
  // so a 20-row ledger does not pay for a 10,000-row patient master.
  const needed = {};
  const rows = [];
  for (let i = 1; i < apptData.length; i++) {
    let dObj = apptData[i][3];
    let rowDate = (dObj instanceof Date)
      ? Utilities.formatDate(dObj, Session.getScriptTimeZone(), "yyyy-MM-dd")
      : (dObj ? dObj.toString().substring(0, 10) : "");
    if (rowDate !== dateStr) continue;
    let pId = apptData[i][1] ? apptData[i][1].toString().toUpperCase() : "";
    if (pId === "ADMIN") continue;
    needed[pId] = true;
    rows.push({ r: apptData[i], pId: pId });
  }
  if (!rows.length) return [];

  const patientData = patientSheet.getDataRange().getValues();
  const patientMap = {};
  for (let i = 1; i < patientData.length; i++) {
    let key = patientData[i][0] ? patientData[i][0].toString().toUpperCase() : "";
    if (needed[key]) patientMap[key] = { name: patientData[i][2], age: patientData[i][3], sex: patientData[i][4] };
  }

  return rows.map(function (x) {
    const p = patientMap[x.pId];
    const nm = x.r[2] || (p ? p.name : '-');
    return {
      apptId: x.r[0],
      time: formatTimeSafely(x.r[4]),
      patientId: x.pId,
      // both keys on purpose: older screens read .name, newer ones read .patientName
      name: nm,
      patientName: nm,
      age: p ? p.age : '-',
      sex: p ? p.sex : '-',
      purpose: x.r[5],
      status: x.r[6]
    };
  });
}

// 4. BOOK APPOINTMENT WRITER (With Security Overrides)
function submitNewAppointment(apptObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Appointments');
    const data = sheet.getDataRange().getValues();

    if (apptObj.patientId !== 'ADMIN' && apptObj.patientId !== 'DIRECT' && apptObj.patientId !== 'WALK-IN') {
      for (let i = 1; i < data.length; i++) {
        let dObj = data[i][3];
        let rowDate = (dObj instanceof Date)
          ? Utilities.formatDate(dObj, Session.getScriptTimeZone(), "yyyy-MM-dd")
          : (dObj ? dObj.toString().substring(0, 10) : "");
        if (rowDate !== apptObj.date) continue;
        if (data[i][1] === apptObj.patientId) {
          let status = data[i][6];
          if (status === 'Booked' || status === 'Arrived' || status === 'In-Progress') {
            return { success: false, message: "This patient already has an active appointment on that date." };
          }
        }
        if (formatTimeSafely(data[i][4]) === apptObj.time && data[i][6] !== 'Cancelled' && data[i][6] !== 'DELETE') {
          return { success: false, message: "Slot collision. That time was just booked by another user." };
        }
      }
    }

    // Timestamp-derived, not row-count-derived: deleting a row must never let the
    // next booking reuse an id that is already printed on a bill.
    const newId = "APT-" + Date.now().toString().slice(-8);
    sheet.appendRow([
      newId, apptObj.patientId, apptObj.patientName, apptObj.date, apptObj.time,
      apptObj.purpose, apptObj.status || 'Booked', apptObj.fee || 0, new Date().toISOString()
    ]);
    SpreadsheetApp.flush();
    return { success: true, apptId: newId };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// 5. UPDATE APPOINTMENT STATUS (Queue advancement / Deletion)
/**
 * Appointment ledger status writer. SESSION REQUIRED.
 *
 * Previously this ran with no lock, no session check and no scope check: any
 * caller with the /exec URL could rewrite — or permanently delete — any
 * appointment row. It now matches the guarantees the scan check-in path
 * already had (Barcode_Engine.gs), so both writers are equally safe.
 *
 * DELETE still removes the row, because that is what the ledger's delete
 * button has always meant. It is now audited WITH a snapshot of the deleted
 * row, so a mistaken delete is recoverable from Audit_Log.
 */
function updateAppointmentStatus(apptId, newStatus, sessionToken) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const sess = dc_validateSession_(sessionToken);
    if (!sess) return "Your session has expired. Please sign in again.";

    const role = dc_str_(sess.role).toLowerCase();
    if (APPT_STATUS_WRITERS.indexOf(role) === -1) {
      return "Your role cannot change appointment status.";
    }

    const id = dc_str_(apptId);
    const target = dc_str_(newStatus);
    if (!id) return "Appointment ID is missing.";
    if (target !== "DELETE" && APPT_ALLOWED_STATUSES.indexOf(target) === -1) {
      return "Unknown status: " + target;
    }
    if (target === "DELETE" && role !== "admin" && role !== "receptionist" && role !== "reception") {
      return "Your role cannot delete appointments.";
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Appointments');
    if (!sheet || sheet.getLastRow() < 2) return "Appointments sheet is missing.";

    const cell = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
      .createTextFinder(id).matchEntireCell(true).findNext();
    if (!cell) return "Error updating.";

    const rowNum = cell.getRow();
    const lastCol = sheet.getLastColumn();
    const row = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];

    // Doctors may only touch their own column of the ledger.
    const m = dc_headerMap_(sheet);
    const docIdx = (m['Doctor_ID'] === undefined) ? -1 : m['Doctor_ID'];
    const rowDoc = (docIdx === -1) ? DC_DEFAULT_DOCTOR : (dc_str_(row[docIdx]) || DC_DEFAULT_DOCTOR);
    const scope = resolveScope_(sessionToken, null);
    if (!dc_inScope_(scope, rowDoc)) return "This appointment belongs to another doctor.";

    const previous = dc_str_(row[6]);

    if (target === "DELETE") {
      // Snapshot first: once the row is gone the audit entry is the only record.
      logAudit_(sess, 'APPOINTMENT_DELETED', 'Appointment', id, {
        patientId: dc_upper_(row[1]),
        status: previous,
        row: row.map(function (v) { return (v instanceof Date) ? v.toISOString() : String(v); })
      });
      sheet.deleteRow(rowNum);
    } else {
      sheet.getRange(rowNum, 7).setValue(target);
      logAudit_(sess, 'APPOINTMENT_STATUS', 'Appointment', id,
                { from: previous, to: target, patientId: dc_upper_(row[1]) });
    }

    SpreadsheetApp.flush();
    return "Status updated!";
  } catch (e) {
    return "Error updating: " + e.message;
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

// 6. BATCH OPTIMIZED AVAILABILITY SAVER
// ==========================================
// 🚀 ENTERPRISE AVAILABILITY ENGINE (FIXED)
// ==========================================

function saveEnterpriseAvailability(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const apptSheet = ss.getSheetByName('Appointments');
    const data = apptSheet.getDataRange().getValues();
    
    // Feature 4: Setup Audit Logging (Safe Check)
    let auditSheet = ss.getSheetByName('Audit_Logs');
    if (!auditSheet) {
      auditSheet = ss.insertSheet('Audit_Logs');
      auditSheet.appendRow(['Timestamp', 'User', 'Module', 'Action', 'Target Dates']);
    }

    const { startDate, endDate, blockedSlots, overrideConflicts } = payload;
    let conflicts = [];
    let datesToProcess = [];

    // FIX 1: Failsafe Date Loop to prevent Infinite Spinning
    const start = new Date(startDate);
    const end = new Date(endDate || startDate);
    let loopDate = new Date(start.getTime());
    let safetyCounter = 0; // Prevent infinite loops

    while (loopDate <= end && safetyCounter < 365) {
      datesToProcess.push(Utilities.formatDate(loopDate, Session.getScriptTimeZone(), "yyyy-MM-dd"));
      loopDate.setDate(loopDate.getDate() + 1);
      safetyCounter++;
    }

    // Feature 1: Conflict Resolution Engine
    for (let i = 1; i < data.length; i++) {
      let dObj = data[i][3];
      let rowDate = (dObj instanceof Date) ? Utilities.formatDate(dObj, Session.getScriptTimeZone(), "yyyy-MM-dd") : dObj.toString().substring(0,10);
      let pId = data[i][1];
      let time = formatTimeSafely(data[i][4]);
      let status = data[i][6];

      if (datesToProcess.includes(rowDate) && blockedSlots.includes(time) && pId !== 'ADMIN' && status !== 'Cancelled' && status !== 'DELETE') {
        conflicts.push({ date: rowDate, time: time, patient: data[i][2] });
      }
    }

    if (conflicts.length > 0 && !overrideConflicts) {
      return { success: false, isConflict: true, conflicts: conflicts };
    }

    // Process the Blocks 
    let rowsToDelete = [];
    for(let i = data.length - 1; i >= 1; i--) {
      let dObj = data[i][3];
      let rowDate = (dObj instanceof Date) ? Utilities.formatDate(dObj, Session.getScriptTimeZone(), "yyyy-MM-dd") : dObj.toString().substring(0,10);
      
      if(datesToProcess.includes(rowDate) && data[i][6] === 'Blocked') {
        rowsToDelete.push(i + 1);
      }
    }
    
    // Delete old blocks
    rowsToDelete.forEach(r => apptSheet.deleteRow(r));

    // Batch append new blocks
    if (blockedSlots.length > 0) {
      const newRows = [];
      let startId = apptSheet.getLastRow();
      
      datesToProcess.forEach(dateStr => {
        blockedSlots.forEach((slot) => {
          startId++;
          let apptId = "APT-" + startId.toString().padStart(4, '0');
          newRows.push([apptId, 'ADMIN', 'BLOCKED', dateStr, slot, 'Doctor Unavailable', 'Blocked', 0]);
        });
      });
      if(newRows.length > 0) apptSheet.getRange(apptSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }

    // Log the Audit Trail (Wrapped in try/catch to prevent it crashing the main save)
    try {
      let activeUser = Session.getActiveUser().getEmail() || 'System Admin';
      let logAction = blockedSlots.length === 0 ? "Cleared Schedule" : `Blocked ${blockedSlots.length} slots per day`;
      auditSheet.appendRow([new Date(), activeUser, 'Availability', logAction, `${startDate} to ${endDate || startDate}`]);
    } catch(err) { /* Ignore audit log failure if permissions block it */ }

    SpreadsheetApp.flush();
    return { success: true, message: 'Availability successfully updated!' };

  } catch(e) {
    // FIX 2: Correctly returning backend errors
    return { success: false, message: 'Server Error: ' + e.message };
  }
}