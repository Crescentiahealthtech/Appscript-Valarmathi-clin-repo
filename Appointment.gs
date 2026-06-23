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
  if(!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toUpperCase() === patientId.trim().toUpperCase()) {
      return { name: data[i][2], age: data[i][3], sex: data[i][4] };
    }
  }
  return null;
}

// 3. ADMIN DAILY LEDGER
function fetchDailyLedger(dateStr) {
  const apptSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Appointments');
  const patientSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Patients');
  if(!apptSheet || !patientSheet) return [];
  const apptData = apptSheet.getDataRange().getValues();
  const patientData = patientSheet.getDataRange().getValues();
  const patientMap = {};
  for (let i = 1; i < patientData.length; i++) {
    patientMap[patientData[i][0].toString().toUpperCase()] = { age: patientData[i][3], sex: patientData[i][4] };
  }

  const ledger = [];
  for (let i = 1; i < apptData.length; i++) {
    let dObj = apptData[i][3];
    let rowDate = (dObj instanceof Date) ? Utilities.formatDate(dObj, Session.getScriptTimeZone(), "yyyy-MM-dd") : dObj.toString().substring(0,10);

    if (rowDate === dateStr) {
      let pId = apptData[i][1].toString().toUpperCase();
      if (pId === "ADMIN") continue;

      ledger.push({
        apptId: apptData[i][0],
        time: formatTimeSafely(apptData[i][4]),
        patientId: pId,
        name: apptData[i][2],
        age: patientMap[pId] ? patientMap[pId].age : '-',
        sex: patientMap[pId] ? patientMap[pId].sex : '-',
        purpose: apptData[i][5],
        status: apptData[i][6]
      });
    }
  }
  return ledger;
}

// 4. BOOK APPOINTMENT WRITER (With Security Overrides)
function submitNewAppointment(apptObj) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Appointments');
    const data = sheet.getDataRange().getValues();

    // Check double booking for actual patients (ignores walk-ins and admin blocks)
    if (apptObj.patientId !== 'ADMIN' && apptObj.patientId !== 'DIRECT' && apptObj.patientId !== 'WALK-IN') {
      for (let i = 1; i < data.length; i++) {
        let dObj = data[i][3];
        let rowDate = (dObj instanceof Date) ? Utilities.formatDate(dObj, Session.getScriptTimeZone(), "yyyy-MM-dd") : dObj.toString().substring(0,10);
        if (rowDate === apptObj.date && data[i][1] === apptObj.patientId) {
          let status = data[i][6];
          if(status === 'Booked' || status === 'Arrived' || status === 'In-Progress') {
            return { success: false, message: "You already have an active appointment scheduled for this date. Please wait until it is completed." };
          }
        }
      }
    }

    const newId = "APT-" + (sheet.getLastRow()).toString().padStart(4, '0');
    sheet.appendRow([
      newId, apptObj.patientId, apptObj.patientName, apptObj.date, apptObj.time, apptObj.purpose, apptObj.status || 'Booked', apptObj.fee || 0
    ]);
    
    SpreadsheetApp.flush(); // Force save to datastore immediately
    return { success: true, apptId: newId };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// 5. UPDATE APPOINTMENT STATUS (Queue advancement / Deletion)
function updateAppointmentStatus(apptId, newStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Appointments');
  const data = sheet.getDataRange().getValues();
  for(let i = 1; i < data.length; i++) {
    if(data[i][0] == apptId) {
      if(newStatus === "DELETE") {
        sheet.deleteRow(i + 1);
      } else {
        sheet.getRange(i + 1, 7).setValue(newStatus);
      }
      SpreadsheetApp.flush(); 
      return "Status updated!";
    }
  }
  return "Error updating.";
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