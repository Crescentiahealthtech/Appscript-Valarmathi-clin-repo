/**
 * Fetches the high-level patient directory for the Ledger view.
 */
function getMasterPatientDirectory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Patients");
  if (!sheet) return JSON.stringify([]);

  const data = sheet.getDataRange().getValues();
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    // Assuming: Col 0=ID, Col 2=Name, Col 3=Age, Col 4=Sex, Col 6=Mobile
    result.push({
      id: data[i][0] ? data[i][0].toString() : "",
      name: data[i][2] || "Unknown",
      age: data[i][3] || "--",
      sex: data[i][4] || "--",
      mobile: data[i][6] ? data[i][6].toString() : ""
    });
  }
  // Reverse to show newest registered patients at the top of the ledger
  return JSON.stringify(result.reverse());
}

/**
 * The Aggregator: Pulls OP and IP records, merges them, and sorts chronologically.
 */
function buildLongitudinalTimeline(patientId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pIdUpper = patientId.trim().toUpperCase();
  
  const response = {
    found: false,
    patient: {},
    events: [] 
  };

  // 1. Get Patient Demographics
  const patSheet = ss.getSheetByName("Patients") || ss.getSheetByName("Patients_DB");
  if (patSheet) {
    const pData = patSheet.getDataRange().getValues();
    for (let i = 1; i < pData.length; i++) {
      if (pData[i][0].toString().trim().toUpperCase() === pIdUpper) {
        response.found = true;
        response.patient = {
          id: pData[i][0],
          name: pData[i][2],
          age: pData[i][3],
          sex: pData[i][4],
          mobile: pData[i][6]
        };
        break;
      }
    }
  }

  if (!response.found) return response;

  // 2. Extract OP Encounters (NEW DYNAMIC COLUMN MAPPING)
  const opSheet = ss.getSheetByName("OP_Encounters");
  if (opSheet) {
    const opData = opSheet.getDataRange().getValues();
    const headers = opData[0]; // Capture Row 1 headers to dynamically label the exams!

    for (let i = 1; i < opData.length; i++) {
      if (opData[i][1] && opData[i][1].toString().trim().toUpperCase() === pIdUpper) { // Col B [1]: Patient_ID
        let rawDate = new Date(opData[i][2]); // Col C [2]: Timestamp
        if (isNaN(rawDate)) continue; 
        
        // --- Assemble Vitals Object (Cols D through J | Indices 3 through 9) ---
        let vitalsObj = {};
        // Combine Sys_BP and Dia_BP into one string
        if (opData[i][3] || opData[i][4]) vitalsObj.bp = (opData[i][3] || '--') + '/' + (opData[i][4] || '--');
        if (opData[i][5]) vitalsObj.hr = opData[i][5];       // F
        if (opData[i][6]) vitalsObj.spo2 = opData[i][6];     // G
        if (opData[i][7]) vitalsObj.temp = opData[i][7];     // H
        if (opData[i][8]) vitalsObj.weight = opData[i][8];   // I
        if (opData[i][9]) vitalsObj.bmi = opData[i][9];      // J

        // --- Assemble Gen Exam (Cols M to Q | Indices 12 to 16) ---
        let genExamObj = {};
        for (let c = 12; c <= 16; c++) {
          if (opData[i][c] && opData[i][c] !== "") {
            // Uses the actual column header from your sheet as the label!
            genExamObj[headers[c] || `GenExam_${c}`] = opData[i][c];
          }
        }

        // --- Assemble Sys Exam (Cols R to U | Indices 17 to 20) ---
        let sysExamObj = {};
        for (let c = 17; c <= 20; c++) {
          if (opData[i][c] && opData[i][c] !== "") {
            sysExamObj[headers[c] || `SysExam_${c}`] = opData[i][c];
          }
        }

        // --- Format Follow-up Date (Col Y or Z) ---
        // Checking if it's a valid date object to format it beautifully
        let rawFollowUp = opData[i][25]; // Col Z [25]: Follow Up Date
        let followUpStr = "None";
        if (rawFollowUp instanceof Date && !isNaN(rawFollowUp)) {
            followUpStr = Utilities.formatDate(rawFollowUp, Session.getScriptTimeZone(), "dd-MMM-yyyy");
        } else if (rawFollowUp) {
            followUpStr = rawFollowUp.toString();
        }

        response.events.push({
          type: 'OP',
          rawDateObj: rawDate.getTime(), 
          date: Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "dd-MMM-yyyy | hh:mm a"),
          apptId: opData[i][0] || "Unknown",                // Col A [0]: Appt_ID
          vitals: JSON.stringify(vitalsObj),                // Converted Vitals
          complaints: opData[i][10] || "Not recorded",      // Col K [10]: Complaints
          history: opData[i][11] || "None",                 // Col L [11]: History
          genExam: JSON.stringify(genExamObj),              // Converted Gen Exam
          sysExam: JSON.stringify(sysExamObj),              // Converted Sys Exam
          diagnosis: opData[i][21] || "Pending",            // Col V [21]: Diagnosis
          rx: opData[i][22] || "[]",                        // Col W [22]: Rx
          labs: opData[i][23] || "None",                    // Col X [23]: Labs
          advice: opData[i][24] || "None",                  // Col Y [24]: Advice
          followUp: followUpStr                             // Col Z [25]: Follow Up
        });
      }
    }
  }

  // 3. Extract IP Admissions (Discharge Summary)
  const ipSheet = ss.getSheetByName("Discharge_Summary");
  if (ipSheet) {
    const ipData = ipSheet.getDataRange().getValues();
    for (let i = 1; i < ipData.length; i++) {
      if (ipData[i][1] && ipData[i][1].toString().trim().toUpperCase() === pIdUpper) {
        let rawDate = new Date(ipData[i][2]); 
        if (isNaN(rawDate)) continue;
        
        response.events.push({
          type: 'IP',
          rawDateObj: rawDate.getTime(),
          date: Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "dd-MMM-yyyy"),
          ipNumber: ipData[i][0] || "IP-N/A", 
          ward: ipData[i][3] || "Discharged", 
          diagnosis: ipData[i][4] || "Final Diagnosis Pending" 
        });
      }
    }
  }

  // Sort Newest First
  response.events.sort((a, b) => b.rawDateObj - a.rawDateObj);
  return response;
}