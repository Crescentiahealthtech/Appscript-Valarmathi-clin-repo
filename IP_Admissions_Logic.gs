// ==========================================
// IP_Admissions_Logic.gs
// ==========================================

function fetchPatientForAdmit(patientId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Patients"); 
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim().toUpperCase() === patientId.trim().toUpperCase()) {
      return { name: data[i][2], age: data[i][3], sex: data[i][4] };
    }
  }
  return null;
}

function saveNewAdmissionLedger(pay) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyMM");
  const randomSeq = Math.floor(1000 + Math.random() * 9000); 
  const newIpNumber = `IP${dateStr}-${randomSeq}`;

  let admitSheet = ss.getSheetByName("IP_Admissions");
  if(!admitSheet) { 
    admitSheet = ss.insertSheet("IP_Admissions"); 
    // Add Headers if it's a new sheet
    admitSheet.appendRow(["IP Number", "Patient ID", "Patient Name", "Age/Sex", "DOA", "TOA", "Type", "Ward", "Bed", "Consultant", "Diagnosis", "Status", "DOD"]);
  }
  
  // Appending ALL payload data matching the headers above
  admitSheet.appendRow([
    newIpNumber, 
    payload.patientId, 
    payload.patientName, 
    payload.ageSex,
    payload.doa, 
    payload.toa,
    payload.type,
    payload.ward,
    payload.bed,
    payload.consultant,
    payload.diagnosis, 
    "ACTIVE",
    "" // DOD (Date of Discharge) left blank initially
  ]);

  let bedSheet = ss.getSheetByName("Master_Beds");
  if(bedSheet) {
    const bData = bedSheet.getDataRange().getValues();
    for(let i=1; i<bData.length; i++) {
      if(bData[i][0] === payload.bed) {
        bedSheet.getRange(i+1, 3).setValue("Occupied");
        bedSheet.getRange(i+1, 4).setValue(payload.patientId);
        bedSheet.getRange(i+1, 5).setValue(payload.patientName); 
        bedSheet.getRange(i+1, 6).setValue(payload.doa);
        bedSheet.getRange(i+1, 7).setValue(newIpNumber);
        break;
      }
    }
  }
  return newIpNumber;
}

function processBedTransfer(ipNumber, oldBedId, newWard, newBedId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let admitSheet = ss.getSheetByName("IP_Admissions");
  let patientId = "", patientName = "", doa = "";
  
  if (admitSheet) {
    const aData = admitSheet.getDataRange().getValues();
    for (let i = 1; i < aData.length; i++) {
      if (aData[i][0] === ipNumber) {
        admitSheet.getRange(i+1, 8).setValue(`${newWard} - ${newBedId}`);
        patientId = aData[i][1];
        patientName = aData[i][2];
        doa = aData[i][4];
        break;
      }
    }
  }

  let bedSheet = ss.getSheetByName("Master_Beds");
  if (bedSheet) {
    const bData = bedSheet.getDataRange().getValues();
    for (let i = 1; i < bData.length; i++) {
      if (bData[i][0] === oldBedId) {
        bedSheet.getRange(i+1, 3).setValue("Cleaning");
        bedSheet.getRange(i+1, 4).clearContent();
        bedSheet.getRange(i+1, 5).clearContent();
        bedSheet.getRange(i+1, 6).clearContent();
        bedSheet.getRange(i+1, 7).clearContent();
        break;
      }
    }
    for (let i = 1; i < bData.length; i++) {
      if (bData[i][0] === newBedId) {
        bedSheet.getRange(i+1, 3).setValue("Occupied");
        bedSheet.getRange(i+1, 4).setValue(patientId);
        bedSheet.getRange(i+1, 5).setValue(patientName);
        bedSheet.getRange(i+1, 6).setValue(doa);
        bedSheet.getRange(i+1, 7).setValue(ipNumber);
        break;
      }
    }
  }
  return true;
}

function getAvailableBedsByWard(ward) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = initializeBedsIfEmpty(ss); // Calls the initializer safely
  
  const data = sheet.getDataRange().getValues();
  let availableBeds = [];
  
  for (let i = 1; i < data.length; i++) {
    const rowWard = data[i][1] ? data[i][1].toString().trim() : "";
    const rowStatus = data[i][2] ? data[i][2].toString().trim().toUpperCase() : "";
    
    if (rowWard === ward && rowStatus === "AVAILABLE") {
      availableBeds.push(data[i][0]);
    }
  }
  return availableBeds;
}

function getIPLedgerData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("IP_Admissions");
    if (!sheet || sheet.getLastRow() < 2) {
      return JSON.stringify({ success: true, data: [] });
    }
    const tz = Session.getScriptTimeZone();
    const fmt = function (v, pat) {
      if (v instanceof Date) return Utilities.formatDate(v, tz, pat);
      return (v == null ? "" : v).toString();
    };
    const data = sheet.getDataRange().getValues();
    const ledger = [];
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (!r[0]) continue; // skip blank/trailing rows
      ledger.push({
        ipNumber:   (r[0]  || "").toString(),
        patientId:  (r[1]  || "").toString(),
        patientName:(r[2]  || "").toString(),
        ageSex:     (r[3]  || "").toString(),
        doa:        fmt(r[4], "dd MMM yyyy"),
        toa:        fmt(r[5], "hh:mm a"),
        type:       (r[6]  || "").toString(),
        ward:       (r[7]  || "").toString(),
        bed:        (r[8]  || "").toString(),
        consultant: (r[9]  || "").toString(),
        diagnosis:  (r[10] || "").toString(),
        status:     (r[11] || "UNKNOWN").toString(),
        dod:        fmt(r[12], "dd MMM yyyy")
      });
    }
    return JSON.stringify({ success: true, data: ledger.reverse() });
  } catch (e) {
    return JSON.stringify({ success: false, message: e.message });
  }
}

function processPatientDischarge(ipNumber, bedId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  // 1. Update Admission Ledger Status
  let admitSheet = ss.getSheetByName("IP_Admissions");
  if (admitSheet) {
    const aData = admitSheet.getDataRange().getValues();
    for (let i = 1; i < aData.length; i++) {
      if (aData[i][0] === ipNumber) {
        admitSheet.getRange(i+1, 12).setValue("DISCHARGED"); // Status Col
        admitSheet.getRange(i+1, 13).setValue(dateStr);      // DOD Col
        break;
      }
    }
  }

  // 2. Free up the Bed (Set to Cleaning)
  let bedSheet = ss.getSheetByName("Master_Beds");
  if (bedSheet) {
    const bData = bedSheet.getDataRange().getValues();
    for (let i = 1; i < bData.length; i++) {
      if (bData[i][0] === bedId) {
        bedSheet.getRange(i+1, 3).setValue("Cleaning");
        bedSheet.getRange(i+1, 4).clearContent();
        bedSheet.getRange(i+1, 5).clearContent();
        bedSheet.getRange(i+1, 6).clearContent();
        bedSheet.getRange(i+1, 7).clearContent();
        break;
      }
    }
  }
  return true;
}