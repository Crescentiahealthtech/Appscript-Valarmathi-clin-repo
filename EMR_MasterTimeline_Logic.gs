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

  // 3. Admissions and their discharge summaries.
  //
  // This used to read a sheet called "Discharge_Summary", which NOTHING in
  // the project writes — see the same note in ipr_discharge_(). Every signed
  // summary produced by the discharge module lives in DS_Summaries, so the
  // patient's timeline showed no inpatient episode at all, however many
  // times they had been admitted and discharged.
  //
  // It now reads the admissions register, which is authoritative for who was
  // admitted when, and enriches each stay with its summary from the discharge
  // engine: the status, the final diagnosis, who signed it and the hash that
  // identifies the signed version.
  mt_pushAdmissions_(ss, pIdUpper, response.events);

  // Sort Newest First
  response.events.sort((a, b) => b.rawDateObj - a.rawDateObj);
  return response;
}

/**
 * One timeline event per admission, carrying its discharge summary when the
 * discharge module has one.
 *
 * @param {Spreadsheet} ss
 * @param {string} pIdUpper       canonical patient id
 * @param {Array} events          appended to in place
 */
function mt_pushAdmissions_(ss, pIdUpper, events) {
  const sheet = ss.getSheetByName("IP_Admissions");
  if (!sheet || sheet.getLastRow() < 2) return;

  const tz = Session.getScriptTimeZone();
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[1] || r[1].toString().trim().toUpperCase() !== pIdUpper) continue;   // Patient_ID

    // A stay is placed on the timeline by its admission date, which always
    // exists; the discharge date may not yet.
    const doa = new Date(r[4]);
    if (isNaN(doa)) continue;

    const ipNumber = (r[0] || "").toString().trim();
    const dod = new Date(r[12]);
    const status = (r[11] || "").toString().trim().toUpperCase();

    const ev = {
      type: 'IP',
      rawDateObj: doa.getTime(),
      date: Utilities.formatDate(doa, tz, "dd-MMM-yyyy"),
      ipNumber: ipNumber || "IP-N/A",
      ward: (r[7] || r[8] || "").toString().trim() || "—",
      admitted: Utilities.formatDate(doa, tz, "dd-MMM-yyyy"),
      discharged: isNaN(dod) ? "" : Utilities.formatDate(dod, tz, "dd-MMM-yyyy"),
      status: status || "ADMITTED",
      diagnosis: (r[10] || "").toString().trim() || "Final Diagnosis Pending",
      consultant: (r[9] || "").toString().trim(),
      // filled in below when the discharge module has a summary
      summaryId: "", summaryStatus: "", signedAt: "", signedBy: "",
      signerRegNo: "", shortHash: "", summarySigned: false
    };

    // ipr_dischargeSummaryBlock_ (IP_Records_Logic.gs) already knows how to
    // read the discharge engine for one admission, including the signed
    // snapshot list and the final diagnosis. Reusing it keeps the timeline
    // and the IP record file saying the same thing about the same stay.
    try {
      if (ipNumber && typeof ipr_dischargeSummaryBlock_ === "function") {
        const ds = ipr_dischargeSummaryBlock_(ipNumber);
        if (ds) {
          ev.summaryId     = ds.summaryId || "";
          ev.summaryStatus = ds.status || "";
          ev.summarySigned = ds.status === "SIGNED";
          ev.signedBy      = ds.signedBy || "";
          ev.signerRegNo   = ds.signerRegNo || "";
          ev.shortHash     = ds.shortHash || "";
          ev.signedAt      = ds.date || "";
          if (ds.diagnosis) ev.diagnosis = ds.diagnosis;
          if (ds.outcome) ev.status = ds.outcome;
        }
      }
    } catch (e) { /* the discharge module may not be on this deployment */ }

    events.push(ev);
  }
}