// ==========================================
// 🏥 IP RECORDS LOGIC (Historical Ledger & Details)
// ==========================================

function fetchIPRecordsLedger() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("IP_CaseSheets_DB");
    if (!sheet) return JSON.stringify([]);

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return JSON.stringify([]);

    const result = [];
    // Schema: Encounter_ID[0], IP_Number[1], Ward[2], Bed[3], Patient_ID[4], Timestamp[5], Name[6], Age[7], Sex[8]
    for (let i = data.length - 1; i > 0; i--) { // Reverse loop for newest first
      result.push({
        encounterId: data[i][0],
        ipNumber: data[i][1] || "--",
        ward: data[i][2] || "--",
        bed: data[i][3] || "--",
        patientId: data[i][4],
        date: Utilities.formatDate(new Date(data[i][5]), Session.getScriptTimeZone(), "dd MMM yyyy, hh:mm a"),
        name: data[i][6] || "Unknown",
        ageSex: (data[i][7] || "-") + " / " + (data[i][8] || "-")
      });
    }
    return JSON.stringify(result);
  } catch (e) {
    return JSON.stringify({ error: e.toString() });
  }
}

function fetchFullIPRecord(encounterId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const result = { casesheet: null, notes: [], discharge: null };

    // 1. Fetch Casesheet (Upgraded to fetch all 34+ columns)
    const csSheet = ss.getSheetByName("IP_CaseSheets_DB");
    if (csSheet) {
      const csData = csSheet.getDataRange().getValues();
      const row = csData.find(r => r[0] === encounterId);
      if (row) {
        result.casesheet = {
          encounterId: row[0], ipNumber: row[1], patientId: row[4],
          name: row[6], ageSex: row[7] + "/" + row[8],
          vitals: `BP: ${row[9]}/${row[10]} | PR: ${row[11]} | SpO2: ${row[12]}% | Temp: ${row[13]} | Wt: ${row[15]}kg`,
          chiefComplaints: row[16] || "--",
          history: row[17] || "--",
          generalExam: `Pallor: ${row[18]} | Icterus: ${row[19]} | Cyanosis: ${row[20]} | Clubbing: ${row[21]} | Edema: ${row[22]} | Other: ${row[23]}`,
          systemicExam: `CVS: ${row[24]} | RS: ${row[25]} | PA: ${row[26]} | CNS: ${row[27]}`,
          diagnosis: row[28] || "--",
          prescriptions: row[29] || "[]", 
          labOrders: row[30] || "[]",     
          advice: row[33] || "--",
          doctor: row[34] || "--"
        };
      }
    }

    // 2. Fetch Progress Notes
    const notesSheet = ss.getSheetByName("IP_Notes");
    if (notesSheet) {
      const notesData = notesSheet.getDataRange().getValues();
      for (let i = 1; i < notesData.length; i++) {
        if (notesData[i][1] === encounterId) {
          result.notes.push({
            date: Utilities.formatDate(new Date(notesData[i][0]), Session.getScriptTimeZone(), "dd MMM, hh:mm a"),
            doctor: notesData[i][3] || "Duty Doctor",
            vitals: notesData[i][4] || "--",
            note: notesData[i][5] || "--"
          });
        }
      }
    }

    // 3. Fetch Discharge Summary
    const dsSheet = ss.getSheetByName("Discharge_Summary");
    if (dsSheet) {
      const dsData = dsSheet.getDataRange().getValues();
      const dsRow = dsData.find(r => r[0] === encounterId);
      if (dsRow) {
        result.discharge = {
          date: Utilities.formatDate(new Date(dsRow[1]), Session.getScriptTimeZone(), "dd MMM yyyy"),
          outcome: dsRow[4] || "--",
          summary: dsRow[5] || "--"
        };
      }
    }

    return JSON.stringify(result);
  } catch (e) {
    return JSON.stringify({ error: e.toString() });
  }
}