// ==========================================
// 🩺 EMR RECORDS MODULE (COMMON & IP WARD)
// ==========================================

function searchPatientForEMR(patientId) {
  try {
    // Server-side caller: uses the internal reader (CodeMV.gs). getUserProfile()
    // is now session-checked and is for browser calls only.
    const data = pt_readProfile_(patientId); 
    
    if (data) {
      return {
        success: true,
        data: { 
          id: patientId, 
          name: data.name || "Unknown Patient", 
          age: data.age || "-", 
          gender: data.gender || "-", 
          history: data.history || "" 
        }
      };
    }
    return { success: false, message: "Patient ID not found in database." };
  } catch (error) {
    return { success: false, message: "System Error: " + error.toString() };
  }
}

function getPatientEMRHistory(patientId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('OP_Encounters'); 
    if(!sheet) return [];
    
    const data = sheet.getDataRange().getDisplayValues();
    let history = [];
    
    // Scan backward to show the most recent visits first
    for(let i = data.length - 1; i > 0; i--) {
      if(data[i][1].toString().trim().toUpperCase() === patientId.toString().trim().toUpperCase()) { 
        history.push({
          encounterId: data[i][0],
          date: data[i][2], // Timestamp
          bp: data[i][3] + '/' + data[i][4], // Sys_BP / Dia_BP
          hr: data[i][5] || '--',
          spo2: data[i][6] || '--',
          weight: data[i][8] || '--',
          diagnosis: data[i][21] || 'Pending Dx'
        });
      }
    }
    return history;
  } catch (e) {
    return []; // Return empty array safely on crash
  }
}

function saveClinicalEncounter(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); 
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const encountersSheet = ss.getSheetByName("Encounters_DB");
    const labSheet = ss.getSheetByName("Lab_Queue_DB");
    const pharmacySheet = ss.getSheetByName("Pharmacy_Queue_DB");

    const timestamp = new Date();
    const dateStr = Utilities.formatDate(timestamp, ss.getSpreadsheetTimeZone(), "MM/dd/yyyy HH:mm:ss");
    const uniqueHash = timestamp.getTime().toString().slice(-6);
    const encounterId = "ENC-" + payload.patientId + "-" + uniqueHash;
    
    // STRICTLY bypasses OP. Only triggers for WARD/IP cases.
    if (encountersSheet && payload.visitType !== 'OP') {
      encountersSheet.appendRow([
        encounterId, payload.patientId, payload.apptId || "WARD", dateStr,
        payload.visitType, payload.vitals.bp, payload.vitals.hr, payload.vitals.spo2,
        payload.vitals.weight, payload.complaints, payload.diagnosis, 'Admitted'
      ]);
    }

    if (labSheet && payload.labs && payload.labs.length > 0) {
      payload.labs.forEach(function(lab, index) {
        labSheet.appendRow([
          "LAB-" + uniqueHash + "-" + index, encounterId, payload.patientId, dateStr, 
          lab.testName, lab.priority || "Routine", "Pending_Sample", "", "", ""
        ]);
      });
    }

    if (pharmacySheet && payload.meds && payload.meds.length > 0) {
      payload.meds.forEach(function(med, index) {
        pharmacySheet.appendRow([
          "RX-" + uniqueHash + "-" + index, encounterId, payload.patientId, 
          med.drugName, med.strength, med.sig, med.duration, med.qty, "Pending"
        ]);
      });
    }

    SpreadsheetApp.flush();
    return { success: true, message: "IP Encounter routed successfully!", encounterId: encounterId };

  } catch (error) {
    return { success: false, message: "Database routing failed: " + error.toString() };
  } finally {
    lock.releaseLock();
  }
}