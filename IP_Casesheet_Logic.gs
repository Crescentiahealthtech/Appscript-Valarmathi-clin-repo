// ==========================================
// IP_Casesheet_Logic.gs (Upgraded OP-Parity Engine)
// ==========================================

function getIPAdmissions() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("IP_Admissions"); 
    
    if (!sheet) return JSON.stringify([]); // clean empty state — never fabricate patients

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return JSON.stringify([]);

    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][11] === "ACTIVE") { 
        result.push({
          ipNumber: data[i][0],
          id: data[i][1],
          name: data[i][2],
          age: data[i][3] || "--",
          sex: data[i][4] || "--",
          doa: data[i][5] ? Utilities.formatDate(new Date(data[i][5]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "--",
          ward: data[i][7] || "Ward",
          bed: data[i][8],
          triage: data[i][13] || "Stable" 
        });
      }
    }
    return JSON.stringify(result);
  } catch (e) {
    return JSON.stringify({ error: e.toString() });
  }
}

function saveIPRecord(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); 

  // --- Server-side identity & authorization (ABDM attribution) ---
    const sess = validateSession_(payload.sessionToken);
    if (!sess)                  return { success: false, message: "Session expired. Please log in again." };
    if (sess.role !== 'doctor') return { success: false, message: "Only a logged-in doctor can author a casesheet." };
    if (!sess.doctorId)         return { success: false, message: "Your account is not linked to a doctor profile. Contact admin." };
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = 'IP_CaseSheets_DB';
    let sheet = ss.getSheetByName(sheetName);
    const rxSheet = ss.getSheetByName("IP_Pharmacy_Queue"); // Route IP meds here
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow([
        "Encounter_ID", "IP_Number", "Ward", "Bed", "Patient_ID", "Timestamp", "Patient Name", "Age", "Sex", 
        "Sys_BP", "Dia_BP", "PR", "SpO2", "Temp", "Height", "Weight", 
        "Chief_Complaints", "History", "Pallor", "Icterus", "Cyanosis", "Clubbing", "Edema", "Other GE findings", 
        "CVS", "RS", "PA", "CNS", "Primary Diagnosis", "Prescription_JSON", "Lab_Orders_JSON", 
        "Outside Lab Records", "Radiological records", "Advice", "Doctor's Name", "Doctor_ID"
      ]);
      sheet.getRange("A1:AI1").setFontWeight("bold").setBackground("#d9ead3");
    }

    const timestamp = new Date();
    const dateStr = Utilities.formatDate(timestamp, ss.getSpreadsheetTimeZone(), "MM/dd/yyyy hh:mm a");
    const uniqueHash = timestamp.getTime().toString().slice(-6);
    const encounterId = "IP-ENC-" + payload.patientId + "-" + uniqueHash;

    const hasGE = (flag) => payload.genExam.flags.includes(flag) ? "Yes" : "No";

    const medsJSON = JSON.stringify(payload.meds || []);
    const labsJSON = JSON.stringify(payload.labs || []);

    const rowData = [
      encounterId,
      payload.ipNumber,
      payload.ward,
      payload.bed,
      payload.patientId,
      dateStr,
      payload.patientName,
      payload.age,
      payload.sex,
      payload.vitals.sys,
      payload.vitals.dia,
      payload.vitals.hr,
      payload.vitals.spo2,
      payload.vitals.temp,
      payload.vitals.height,
      payload.vitals.weight,
      JSON.stringify(payload.complaints),
      JSON.stringify(payload.history),
      hasGE('Pallor'),
      hasGE('Icterus'),
      hasGE('Cyanosis'),
      hasGE('Clubbing'),
      hasGE('Edema'),
      payload.genExam.notes,
      payload.sysExam.cvs,
      payload.sysExam.rs,
      payload.sysExam.pa,
      payload.sysExam.cns,
      payload.provDiagnosis,
      medsJSON,
      labsJSON,
      JSON.stringify(payload.outsideLabs), 
      payload.radiology,
      payload.advice,
      sess.name
    ];

    ensureColumn_(sheet, "Doctor_ID");
    rowData.push(sess.doctorId);
    sheet.appendRow(rowData);

    sheet.appendRow(rowData);

    // 1. PHARMACY ROUTING (Internal Drugs Only)
    if (rxSheet && payload.meds && payload.meds.length > 0) {
      payload.meds.forEach((med, index) => {
        if ((med.source || "INTERNAL").toUpperCase() === "EXTERNAL") return;
        // Schema: OrderID | IP_Number | Patient_ID | Timestamp | Item | Sig | Doctor | Status | Duration
        rxSheet.appendRow([
          `IP-RX-${uniqueHash}-${index}`, 
          payload.ipNumber, 
          payload.patientId, 
          dateStr, 
          `${med.strength || ""} ${med.drugName || ""}`.trim(),
          med.sig || "", 
          sess.name || "Doctor", 
          "Pending", 
          med.duration || ""
        ]);
      });
    }

    // 2. LAB ROUTING (Internal Tests Only)
    const labOrders = (payload.labs || []).filter(l => l.type === "Order" && (l.source || "INTERNAL").toUpperCase() !== "EXTERNAL");
    if (labOrders.length > 0) {
      try {
        const testNameList = labOrders.map(l => l.testName).filter(Boolean);
        const hasStat = labOrders.some(l => (l.priority || '').toUpperCase() === 'STAT');
        createLabRequest({
          patientId:          payload.patientId,
          patientName:        payload.patientName || '',
          age:                payload.age || '',
          gender:             payload.sex || '',
          admissionId:        payload.ipNumber || encounterId,
          sourceModule:       'IP_CASESHEET',
          visitId:            encounterId,
          orderingDoctorName: sess.name || '',
          testNames:          testNameList,
          priority:           hasStat ? 'STAT' : 'ROUTINE',
          clinicalNote:       payload.provDiagnosis || ''
        });
      } catch (labErr) {
        Logger.log('IP Casesheet → Lab bridge failed: ' + labErr.message);
      }
    }

    // 3. TEMPLATE LEARNING
    try { if (payload.templateLearn) learnTemplates(payload.templateLearn); } 
    catch (tErr) { Logger.log("template learn skipped: " + tErr.message); }

    logAudit_(sess, "CASESHEET_SAVE", "IP_CaseSheet", encounterId, {
      patientId: payload.patientId, ipNumber: payload.ipNumber, diagnosis: payload.provDiagnosis
    });

    SpreadsheetApp.flush();
    return { success: true, message: "IP Casesheet Locked & Saved to DB Successfully!", encounterId: encounterId };

  } catch (error) {
    return { success: false, message: "Database Error: " + error.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// INTEGRATION FETCHERS (Mirrored from OP)
// ==========================================
function fetchIPDrugMaster() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pharmacy_Inventory");
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const drugs = [];
    for (let i = 1; i < data.length; i++) {
      const brand = String(data[i][1] || "").trim();
      if (!brand) continue;
      const stock      = parseInt(data[i][4], 10) || 0;        
      const unit       = String(data[i][5] || "").trim();      
      const refDose    = String(data[i][6] || "").trim();      
      const adultDose  = String(data[i][7] || "").trim();      
      const reorder    = parseInt(data[i][8], 10) || 5;        
      const status     = stock <= 0 ? "out" : (stock <= reorder ? "low" : "ok");
      drugs.push({
        brand: brand, generic: String(data[i][2] || "").trim(), type: String(data[i][3] || "Tab").trim(),
        stock: stock, unit: unit, refDose: refDose, adultDose: adultDose, status: status, source: "INTERNAL"
      });
    }
    return drugs;
  } catch (e) { return []; }
}

function fetchIPUniversalDrugs() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Drug_Master_Universal");
    if (!sheet) return [];
    const data = sheet.getDataRange().getDisplayValues();
    const out = [];
    for (let i = 1; i < data.length; i++) {
      const brand = (data[i][0] || "").trim();
      if (!brand) continue;
      out.push({
        brand: brand, generic: (data[i][1] || "").trim(), type: (data[i][2] || "Tab").trim(),
        stock: null, status: "external", source: "EXTERNAL"
      });
    }
    return out;
  } catch (e) { return []; }
}

function fetchIPLabTestMaster() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lab_Test_Master");
    if (!sheet) return { internal: [], external: [] };
    const data = sheet.getDataRange().getDisplayValues();
    const internal = [], external = [];
    for (let i = 1; i < data.length; i++) {
      const name = (data[i][0] || "").trim();
      if (!name) continue;
      const item = { testName: name, panel: (data[i][1] || "").trim(), sample: (data[i][3] || "").trim(), tat: (data[i][4] || "").trim() };
      if (String(data[i][2] || "").trim().toUpperCase() === "Y") { item.source = "INTERNAL"; internal.push(item); }
      else { item.source = "EXTERNAL"; external.push(item); }
    }
    return { internal: internal, external: external };
  } catch (e) { return { internal: [], external: [] }; }
}

function fetchIPClinicalTemplates() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Clinical_Templates");
    if (!sheet) return { CC: [], HX: [], ADVICE: [] };
    const data = sheet.getDataRange().getDisplayValues();
    const buckets = { CC: [], HX: [], ADVICE: [] };
    for (let i = 1; i < data.length; i++) {
      const cat = (data[i][0] || "").trim().toUpperCase();
      const text = (data[i][1] || "").trim();
      const count = parseInt(data[i][2], 10) || 0;
      if (!text || !buckets[cat]) continue;
      buckets[cat].push({ text: text, count: count });
    }
    Object.keys(buckets).forEach(k => buckets[k].sort((a, b) => b.count - a.count));
    return buckets;
  } catch (e) { return { CC: [], HX: [], ADVICE: [] }; }
}

// ==========================================
// IP PRINT HTML GENERATOR
// ==========================================
function getIPCasesheetHtml(encounterId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("IP_CaseSheets_DB");
    if (!sheet) return { success: false, message: "IP_CaseSheets_DB not found." };

    const data = sheet.getDataRange().getDisplayValues();
    let record = null;
    
    // Find the record
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === encounterId) {
        record = data[i];
        break;
      }
    }
    
    if (!record) return { success: false, message: "Casesheet record not found." };

    // Extract JSON Arrays
    let complaints = [], history = [], meds = [], labs = [], outsideLabs = [];
    try { complaints = JSON.parse(record[16]); } catch(e){}
    try { history = JSON.parse(record[17]); } catch(e){}
    try { meds = JSON.parse(record[29]); } catch(e){}
    try { labs = JSON.parse(record[30]); } catch(e){}
    try { outsideLabs = JSON.parse(record[31]); } catch(e){}

    // Format Lists
    let ccStr = complaints.map(c => `${c.condition} (${c.duration})`).join(" | ");
    let hxStr = history.map(h => `${h.prefix} ${h.condition} (${h.duration})`).join(" | ");

    let medsHtml = meds.length ? `<ol style="margin:0; padding-left: 20px;">` + meds.map(m => `<li style="margin-bottom:6px;"><strong>${m.type} ${m.drugName}</strong><br><span style="color:#555; font-size:0.85em;">${m.sig} | ${m.duration} | ${m.comments}</span></li>`).join('') + `</ol>` : `<span style="color:#777;">No admission medication ordered.</span>`;

    let internalLabs = labs.filter(l => l.source === 'INTERNAL' || l.type === 'Order').map(l => l.testName).join(", ");
    let extLabsHtml = outsideLabs.length ? outsideLabs.map(l => `${l.test}: ${l.val}`).join(" | ") : "None";

    let html = `
    <div style="font-family: Arial, sans-serif; color: #000; padding: 20px; max-width: 800px; margin: auto;">
        <div style="border-bottom: 2px solid #0369a1; padding-bottom: 10px; margin-bottom: 20px; text-align: center;">
            <h2 style="margin:0; text-transform:uppercase; font-weight:bold; color: #0369a1;">Valarmathi Clinic</h2>
            <p style="margin:0; font-size: 0.9rem; color: #555;">Premium Healthcare Services | Ph: +91 88387 23513</p>
            <h4 style="margin-top:10px; color: #111;">IP Admission Casesheet</h4>
        </div>

        <div style="display: flex; justify-content: space-between; margin-bottom: 20px; font-size: 0.95rem; background: #f9fafb; padding: 15px; border: 1px solid #e5e7eb; border-radius: 6px;">
            <div>
                <strong>Patient:</strong> ${record[6]}<br>
                <strong>PID:</strong> ${record[4]} | <strong>IP No:</strong> ${record[1]}<br>
                <strong>Age/Sex:</strong> ${record[7]} Y / ${record[8]}<br>
                <strong>Ward/Bed:</strong> ${record[2]} - ${record[3]}
            </div>
            <div style="text-align: right;">
                <strong>Date:</strong> ${record[5]}<br>
                <strong>Doctor:</strong> ${record[34] || '--'}
            </div>
        </div>

        <div style="margin-bottom: 20px; font-size: 0.9rem;">
            <h5 style="border-bottom: 1px solid #ccc; padding-bottom: 5px; color: #0369a1;">Vitals on Admission</h5>
            <p style="margin: 5px 0;"><strong>BP:</strong> ${record[9]}/${record[10]} mmHg &nbsp;|&nbsp; <strong>Pulse:</strong> ${record[11]} bpm &nbsp;|&nbsp; <strong>SpO2:</strong> ${record[12]}% &nbsp;|&nbsp; <strong>Temp:</strong> ${record[13]} °F &nbsp;|&nbsp; <strong>Wt:</strong> ${record[15]} kg</p>
        </div>

        <div style="margin-bottom: 20px; font-size: 0.9rem;">
            <h5 style="border-bottom: 1px solid #ccc; padding-bottom: 5px; color: #0369a1;">Clinical History & Examination</h5>
            <p style="margin: 5px 0;"><strong>Chief Complaints:</strong> ${ccStr || '--'}</p>
            <p style="margin: 5px 0;"><strong>History:</strong> ${hxStr || '--'}</p>
            <p style="margin: 5px 0;"><strong>General Exam:</strong> Pallor: ${record[18]}, Icterus: ${record[19]}, Cyanosis: ${record[20]}, Clubbing: ${record[21]}, Edema: ${record[22]}<br><em>Notes:</em> ${record[23] || '--'}</p>
            <p style="margin: 5px 0;"><strong>Systemic Exam:</strong> CVS: ${record[24]} | RS: ${record[25]} | P/A: ${record[26]} | CNS: ${record[27]}</p>
            <p style="margin: 5px 0;"><strong>Primary Diagnosis:</strong> <span style="font-size:1.1em; font-weight:bold;">${record[28] || '--'}</span></p>
        </div>

        <div style="margin-bottom: 20px; font-size: 0.9rem;">
            <h5 style="border-bottom: 1px solid #ccc; padding-bottom: 5px; color: #0369a1;">Admission Orders (Rx / Diet / Nursing)</h5>
            ${medsHtml}
        </div>

        <div style="display:flex; justify-content: space-between; font-size: 0.9rem; margin-bottom: 20px;">
            <div style="width: 48%;">
                <h5 style="border-bottom: 1px solid #ccc; padding-bottom: 5px; color: #0369a1;">Lab Orders</h5>
                <p style="margin: 5px 0;">${internalLabs || 'None'}</p>
                <h5 style="border-bottom: 1px solid #ccc; padding-bottom: 5px; margin-top: 15px; color: #0369a1;">External Lab Results</h5>
                <p style="margin: 5px 0;">${extLabsHtml}</p>
            </div>
            <div style="width: 48%;">
                <h5 style="border-bottom: 1px solid #ccc; padding-bottom: 5px; color: #0369a1;">Radiology / Scans</h5>
                <p style="margin: 5px 0; white-space: pre-wrap;">${record[32] || 'None'}</p>
                <h5 style="border-bottom: 1px solid #ccc; padding-bottom: 5px; margin-top: 15px; color: #0369a1;">Advice / Plan</h5>
                <p style="margin: 5px 0; white-space: pre-wrap;">${record[33] || 'Standard ward protocol.'}</p>
            </div>
        </div>

        <div style="text-align: right; margin-top: 60px;">
            <div style="border-top: 1px solid #000; display: inline-block; padding-top: 5px; width: 200px; text-align: center;">
                <strong>${record[34] || "Doctor's Signature"}</strong>
            </div>
        </div>
    </div>
    `;

    return { success: true, html: html };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}