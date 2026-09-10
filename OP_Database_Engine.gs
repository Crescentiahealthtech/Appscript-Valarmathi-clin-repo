// ==========================================
// OP DATABASE ENGINE: FETCH, UPSERT & PRINT  (clean, integrated)
// Replace your ENTIRE OP_Database_Engine.gs with this file.
// ==========================================

// 1. FETCH DETAILS
function fetchPatientContext(patientId) {
  const result = searchPatientForEMR(patientId);
  if (result.success) {
    return {
      success: true,
      patientId: result.data.id,
      name: result.data.name,
      age: result.data.age,
      sex: result.data.gender,
      history: result.data.history
    };
  }
  return result;
}

// 2. MAIN OP ENCOUNTER SAVE / UPSERT
function saveOPEncounter(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const opSheet = ss.getSheetByName("OP_Encounters");
    const rxSheet = ss.getSheetByName("Pharmacy_Queue_DB");
    const labSheet = ss.getSheetByName("Lab_Queue_DB");
    const ledgerSheet = ss.getSheetByName("Appointments");

    if (!opSheet) return { success: false, message: "Error: 'OP_Encounters' sheet missing." };

    const timestamp = new Date();
    const dateStr = Utilities.formatDate(timestamp, ss.getSpreadsheetTimeZone(), "MM/dd/yyyy hh:mm a");

    let encounterId = payload.encounterId;
    let isNewEncounter = false;

    if (!encounterId || encounterId === "") {
      const uniqueHash = timestamp.getTime().toString().slice(-5);
      encounterId = `OP-${payload.patientId}-${uniqueHash}`;
      isNewEncounter = true;
    }

    const medsJSON = JSON.stringify(payload.meds || []);
    const labsJSON = JSON.stringify(payload.labs || []);

    const rowData = [
      encounterId,               // A (0)
      payload.patientId,         // B (1)
      dateStr,                   // C (2)
      payload.sysBp || "",       // D (3)
      payload.diaBp || "",       // E (4)
      payload.hr || "",          // F (5)
      payload.spo2 || "",        // G (6)
      payload.temp || "",        // H (7)
      payload.weight || "",      // I (8)
      payload.bmi || "",         // J (9)
      payload.complaints || "",  // K (10)
      payload.history || "",     // L (11)
      payload.pallor || "",      // M (12)
      payload.icterus || "",     // N (13)
      payload.cyanosis || "",    // O (14)
      payload.clubbing || "",    // P (15)
      payload.edema || "",       // Q (16)
      payload.cvs || "",         // R (17)
      payload.rs || "",          // S (18)
      payload.pa || "",          // T (19)
      payload.cns || "",         // U (20)
      payload.diagnosis || "",   // V (21)
      medsJSON,                  // W (22)
      labsJSON,                  // X (23)
      payload.advice || "",      // Y (24)
      payload.reviewDate || ""   // Z (25)
    ];

    let rowIndex = -1;
    if (!isNewEncounter) {
      const data = opSheet.getDataRange().getDisplayValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] !== encounterId) continue;

        // An encounter belongs to exactly one patient. A stale encounterId sent
        // with a different patient's chart used to overwrite the stored row,
        // destroying the first patient's consultation. Refuse, and let the
        // caller save it as a new encounter instead.
        const owner = String(data[i][1] || "").trim().toUpperCase();
        const claimed = String(payload.patientId || "").trim().toUpperCase();
        if (owner && claimed && owner !== claimed) {
          return {
            success: false,
            message: "Encounter " + encounterId + " belongs to patient " + data[i][1] +
                     ", not " + payload.patientId + ". Nothing was saved. Start a new " +
                     "consultation for this patient and save again."
          };
        }
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex > -1) {
      opSheet.getRange(rowIndex, 1, 1, 26).setValues([rowData]);
    } else {
      opSheet.appendRow(rowData);
    }

    // Pharmacy Routing — INTERNAL drugs only (external Rx prints but is never billed)
    if (rxSheet && payload.meds && payload.meds.length > 0) {
      clearExistingQueueRows(rxSheet, encounterId, 2);
      payload.meds.forEach((med, index) => {
        if ((med.source || "INTERNAL").toUpperCase() === "EXTERNAL") return; // skip external Rx
        rxSheet.appendRow([
          `RX-${encounterId}-${index}`,                        // A
          payload.patientId,                                   // B
          encounterId,                                         // C
          dateStr,                                             // D
          `${med.strength || ""} ${med.drugName || ""}`.trim(),// E
          med.sig || "",                                       // F
          "Doctor",                                            // G
          "Pending",                                           // H
          med.duration || ""                                   // I  (days, for billing qty)
        ]);
      });
    }

    // Lab Routing — INTERNAL tests only route to the lab engine
    const labOrders = (payload.labs || []).filter(l =>
      l.type === "Order" && (l.source || "INTERNAL").toUpperCase() !== "EXTERNAL"
    );
    if (labOrders.length > 0) {
      try {
        const testIdList = labOrders.map(l => l.testId).filter(Boolean);
        const hasStat = labOrders.some(l => (l.priority || '').toUpperCase() === 'STAT');
        createOPDLabOrder({
          patientId:   payload.patientId,
          encounterId: encounterId,
          doctorId:    payload.doctorId || '',
          doctorName:  payload.doctorName || '',
          testIds:     testIdList,
          priority:    hasStat ? 'STAT' : 'ROUTINE',
          clinicalNote: payload.diagnosis || ''
        });
      } catch (labErr) {
        Logger.log('OPD → Lab bridge failed: ' + labErr.message); // non-fatal
      }
    }

    // Appointments ledger — mark today's visit completed
    if (ledgerSheet && payload.patientId) {
      const ledgerData = ledgerSheet.getDataRange().getDisplayValues();
      const todayStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
      for (let i = 1; i < ledgerData.length; i++) {
        let rowDateStr = "";
        if (ledgerData[i][3]) {
          let d = new Date(ledgerData[i][3]);
          if (!isNaN(d.getTime())) rowDateStr = Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
        }
        if (ledgerData[i][1] === payload.patientId && rowDateStr === todayStr && ledgerData[i][6] !== "Completed") {
          ledgerSheet.getRange(i + 1, 7).setValue("Completed");
          break;
        }
      }
    }

    // Self-learning templates (complaints / history / advice) — non-fatal
    try { if (payload.templateLearn) learnTemplates(payload.templateLearn); }
    catch (tErr) { Logger.log("template learn skipped: " + tErr.message); }

    SpreadsheetApp.flush();
    return { success: true, message: "OP Record saved successfully!", encounterId: encounterId };

  } catch (error) {
    return { success: false, message: "System Error: " + error.toString() };
  } finally {
    lock.releaseLock();
  }
}

function clearExistingQueueRows(sheet, encounterId, colIndex) {
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][colIndex] === encounterId) sheet.deleteRow(i + 1);
  }
}

// 3. OP PRINT FETCHER
function getEncounterForPrint(encounterId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("OP_Encounters");
    if (!sheet) return { success: false, message: "OP_Encounters sheet missing." };

    const data = sheet.getDataRange().getDisplayValues();

    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === encounterId) {
        const ptData = searchPatientForEMR(data[i][1]);
        const ptName = ptData.success ? ptData.data.name : "Unknown Patient";
        const ptAgeSex = ptData.success ? `${ptData.data.age} / ${ptData.data.gender}` : "-";

        let parsedMeds = [];
        let parsedLabs = [];
        try { parsedMeds = data[i][22] ? JSON.parse(data[i][22]) : []; } catch (e) {}
        try { parsedLabs = data[i][23] ? JSON.parse(data[i][23]) : []; } catch (e) {}

        // Consulting doctor. Doctor_ID and Doctor_Signature_Snapshot are
        // APPENDED columns (op_ensureEncounterColumns_), so they must be read
        // by header name — a fixed index would silently read the wrong cell.
        const hdr = {};
        data[0].forEach((h, c) => { hdr[String(h).trim()] = c; });
        const readCol = (name) => (hdr[name] !== undefined ? String(data[i][hdr[name]] || "").trim() : "");

        const doctorId = readCol("Doctor_ID");
        const sigSnapshot = readCol("Doctor_Signature_Snapshot");

        let doctorName = "";
        let doctorRegNo = "";
        let doctorSpecialty = "";
        if (doctorId && typeof dc_getDoctorById_ === "function") {
          const prof = dc_getDoctorById_(doctorId);
          if (prof) {
            doctorName = prof.name || "";
            doctorRegNo = prof.regNo || "";
            doctorSpecialty = prof.specialty || "";
          }
        }
        // The snapshot is what the doctor signed under at the time of the
        // consult. It wins over anything the Doctors sheet says today.
        if (sigSnapshot) doctorName = sigSnapshot;

        return {
          success: true,
          data: {
            encounterId: data[i][0],
            patientId: data[i][1],
            patientName: ptName,
            patientAgeSex: ptAgeSex,
            date: data[i][2],
            vitals: {
              sysBp: data[i][3], diaBp: data[i][4], hr: data[i][5],
              spo2: data[i][6], temp: data[i][7], weight: data[i][8], bmi: data[i][9]
            },
            clinical: { complaints: data[i][10], history: data[i][11], diagnosis: data[i][21] },
            meds: parsedMeds,
            labs: parsedLabs,
            advice: data[i][24],
            reviewDate: data[i][25],
            doctorId: doctorId,
            doctorName: doctorName,
            doctorRegNo: doctorRegNo,
            doctorSpecialty: doctorSpecialty
          }
        };
      }
    }
    return { success: false, message: "Encounter ID not found." };
  } catch (e) {
    return { success: false, message: "Print fetch error: " + e.message };
  }
}

// ─────────────────────────────────────────────────────────────
// 4. PHARMACY STOCK  (Brand[1] Generic[2] Type[3] Stock[4] Reorder[5])
//    Single definition — the old plain version has been removed.
// ─────────────────────────────────────────────────────────────
function fetchPharmacyInventoryForOP() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pharmacy_Inventory");
    if (!sheet) return [];
    const data = sheet.getDataRange().getDisplayValues();
    const drugs = [];
    for (let i = 1; i < data.length; i++) {
      const brand = (data[i][1] || "").trim();
      if (!brand) continue;
      const stock = parseInt(data[i][4], 10) || 0;
      const reorder = parseInt(data[i][5], 10) || 5; // col F optional; default low threshold = 5
      let status = "ok";
      if (stock <= 0) status = "out";
      else if (stock <= reorder) status = "low";
      drugs.push({
        brand: brand,
        generic: (data[i][2] || "").trim(),
        type: (data[i][3] || "Tab").trim(),
        stock: stock,
        reorder: reorder,
        status: status,       // "ok" | "low" | "out"
        source: "INTERNAL"
      });
    }
    return drugs;
  } catch (e) { return []; }
}

// ─────────────────────────────────────────────────────────────
// 5. UNIVERSAL DRUG MASTER  (Brand[0] Generic[1] Type[2]) — external fallback
// ─────────────────────────────────────────────────────────────
function fetchUniversalDrugs() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Drug_Master_Universal");
    if (!sheet) return [];
    const data = sheet.getDataRange().getDisplayValues();
    const out = [];
    for (let i = 1; i < data.length; i++) {
      const brand = (data[i][0] || "").trim();
      if (!brand) continue;
      out.push({
        brand: brand,
        generic: (data[i][1] || "").trim(),
        type: (data[i][2] || "Tab").trim(),
        stock: null,
        status: "external",
        source: "EXTERNAL"
      });
    }
    return out;
  } catch (e) { return []; }
}

// ─────────────────────────────────────────────────────────────
// 6. LAB TEST MASTER  (Name[0] Panel[1] Internal Y/N[2] Sample[3] TAT[4])
// ─────────────────────────────────────────────────────────────
function fetchLabTestMaster() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lab_Test_Master");
    if (!sheet) return { internal: [], external: [] };
    const data = sheet.getDataRange().getDisplayValues();
    const internal = [], external = [];
    for (let i = 1; i < data.length; i++) {
      const name = (data[i][0] || "").trim();
      if (!name) continue;
      const item = {
        testName: name,
        panel: (data[i][1] || "").trim(),
        sample: (data[i][3] || "").trim(),
        tat: (data[i][4] || "").trim()
      };
      if (String(data[i][2] || "").trim().toUpperCase() === "Y") { item.source = "INTERNAL"; internal.push(item); }
      else { item.source = "EXTERNAL"; external.push(item); }
    }
    return { internal: internal, external: external };
  } catch (e) { return { internal: [], external: [] }; }
}

// ─────────────────────────────────────────────────────────────
// 7. CLINICAL TEMPLATES  (Category[0] Text[1] UseCount[2]) — self-learning
// ─────────────────────────────────────────────────────────────
function fetchClinicalTemplates() {
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

// Increments use-count / inserts new phrases. Called from saveOPEncounter.
function learnTemplates(items) {
  if (!items || !items.length) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Clinical_Templates");
    if (!sheet) { sheet = ss.insertSheet("Clinical_Templates"); sheet.appendRow(["Category", "Text", "UseCount"]); }
    const data = sheet.getDataRange().getValues();
    const index = {};
    for (let i = 1; i < data.length; i++) {
      index[String(data[i][0]).trim().toUpperCase() + "|" + String(data[i][1]).trim().toLowerCase()] = i + 1;
    }
    items.forEach(it => {
      const cat = (it.category || "").trim().toUpperCase();
      const text = (it.text || "").trim();
      if (!text || ["CC", "HX", "ADVICE"].indexOf(cat) === -1) return;
      const key = cat + "|" + text.toLowerCase();
      if (index[key]) {
        const r = index[key];
        sheet.getRange(r, 3).setValue((parseInt(sheet.getRange(r, 3).getValue(), 10) || 0) + 1);
      } else {
        sheet.appendRow([cat, text, 1]);
        index[key] = sheet.getLastRow();
      }
    });
    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log("learnTemplates error: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
//  ADD THIS FUNCTION to your OP_Database_Engine.gs (anywhere).
//  Reads Pharmacy_Inventory and includes optional reference-dose columns.
//  Columns:  B Brand | C Generic | D Type | E Stock | F Unit
//            G RefDose(mg/kg/day, optional) | H AdultDose(optional) | I Reorder(optional)
//  F stays "Unit" to match your working fetchPharmacyMasterForIP().
// ============================================================
function fetchOPDrugMaster() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pharmacy_Inventory");
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const drugs = [];
    for (let i = 1; i < data.length; i++) {
      const brand = String(data[i][1] || "").trim();
      if (!brand) continue;
      const stock     = parseInt(data[i][4], 10) || 0;        // E
      const unit      = String(data[i][5] || "").trim();      // F (unit)
      const refDose   = String(data[i][6] || "").trim();      // G (optional)
      const adultDose = String(data[i][7] || "").trim();      // H (optional)
      const reorder   = parseInt(data[i][8], 10) || 5;        // I (optional)
      const status = stock <= 0 ? "out" : (stock <= reorder ? "low" : "ok");
      drugs.push({
        brand: brand,
        generic: String(data[i][2] || "").trim(),
        type: String(data[i][3] || "Tab").trim(),
        stock: stock,
        unit: unit,
        refDose: refDose,
        adultDose: adultDose,
        reorder: reorder,
        status: status,
        source: "INTERNAL"
      });
    }
    return drugs;
  } catch (e) { return []; }
}

// =========================================================================
// 🖨️ OP PRESCRIPTION - BACKEND HTML GENERATOR
// =========================================================================

function getOPPrescriptionHtml(encounterId) {
  try {
    // 1. Fetch the data using your existing fetcher
    const fetchRes = getEncounterForPrint(encounterId);
    if (!fetchRes.success) return { success: false, message: fetchRes.message };
    
    const data = fetchRes.data;

    // 1b. A prescription must name the doctor who signed it. Printing an
    // anonymous one is not an option, so fail loudly rather than quietly.
    const doctorName = String(data.doctorName || "").trim();
    if (!doctorName) {
      return {
        success: false,
        message: "This encounter (" + (data.encounterId || encounterId) +
                 ") has no consulting doctor recorded, so a prescription cannot " +
                 "be issued. Run runMultiDoctorMigration() to backfill Doctor_ID on " +
                 "legacy rows, or re-save the consult with a doctor selected."
      };
    }
    const doctorRegNo = String(data.doctorRegNo || "").trim();      // optional
    const doctorSpecialty = String(data.doctorSpecialty || "").trim(); // optional

    // 2. Format Dates & Nulls
    let printDate = data.date || "--";
    try {
      let d = new Date(data.date);
      if (!isNaN(d.getTime())) printDate = Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy, hh:mm a");
    } catch (e) {}

    const vitals = data.vitals || {};
    const clin = data.clinical || {};
    const bp = (vitals.sysBp || vitals.diaBp) ? `${vitals.sysBp || '-'}/${vitals.diaBp || '-'}` : "--";
    
    // 3. Build Medication Table Rows
    let medsHtml = "";
    if (!data.meds || data.meds.length === 0) {
      medsHtml = `<tr><td colspan="5" style="text-align:center; padding:15px; color:#6b7280;">No medications prescribed.</td></tr>`;
    } else {
      data.meds.forEach((m, i) => {
        medsHtml += `
          <tr style="page-break-inside: avoid;">
            <td style="padding: 8px; border-bottom: 1px solid #e5e7eb; vertical-align: top;">${i + 1}</td>
            <td style="padding: 8px; border-bottom: 1px solid #e5e7eb; vertical-align: top;"><strong>${m.strength || ''} ${m.drugName || ''}</strong></td>
            <td style="padding: 8px; border-bottom: 1px solid #e5e7eb; vertical-align: top;">${m.sig || '-'}</td>
            <td style="padding: 8px; border-bottom: 1px solid #e5e7eb; vertical-align: top;">${m.duration || '-'}</td>
            <td style="padding: 8px; border-bottom: 1px solid #e5e7eb; vertical-align: top; color: #6b7280;">${m.comments || '-'}</td>
          </tr>`;
      });
    }

    // 4. Build Labs & Results
    let orderNames = [];
    let resultStrs = [];
    if (data.labs && data.labs.length > 0) {
      data.labs.forEach(l => {
        if (l.type === "Result") resultStrs.push(`${l.testName}: ${l.result}`);
        else if (l.type === "Order") orderNames.push(l.testName);
      });
    }
    const ordersHtml = orderNames.length > 0 ? orderNames.join(", ") : "None";
    const resultsHtml = resultStrs.length > 0 ? `<div style="font-size:11px; color:#6b7280; margin-top:6px;">RECORDED FINDINGS:</div><div style="font-weight:600;">${resultStrs.join(" | ")}</div>` : "";

    // 5. Advice
    let advStr = data.advice || "Follow prescribed treatment.";
    if (data.reviewDate) advStr += `<br><br><strong>Next Review:</strong> ${data.reviewDate}`;

    // 6. Assemble the Master HTML String
    const htmlString = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <title>Prescription - ${data.patientName}</title>
        <style>
          @page { size: A4 portrait; margin: 0; }
          body { 
            font-family: 'Helvetica Neue', Arial, sans-serif; 
            margin: 0; padding: 0; 
            background: #fff; 
            color: #111827;
            -webkit-print-color-adjust: exact; 
            color-adjust: exact; 
          }
          .page { 
            width: 210mm; 
            min-height: 297mm; 
            padding: 15mm 20mm; 
            box-sizing: border-box; 
          }
          .section-break { margin-bottom: 5mm; border-bottom: 1px solid #e5e7eb; padding-bottom: 4mm; }
          .label { color: #6b7280; font-weight: normal; font-size: 10pt; }
        </style>
      </head>
      <body>
        <div class="page">
          
          <div style="display:flex; justify-content:space-between; border-bottom: 2px solid #0369a1; padding-bottom: 4mm; margin-bottom: 4mm;">
            <div>
              <div style="font-size: 22pt; font-weight: 900; color: #0369a1; text-transform: uppercase; margin: 0;">Valarmathi Clinic</div>
              <div style="font-size: 10pt; color: #6b7280; margin-top: 2px;">Premium Healthcare Services</div>
              <div style="font-size: 10pt; color: #6b7280;">Phone: +91 88387 23513</div>
            </div>
            <div style="text-align: right;">
              <div style="font-size: 16pt; font-weight: 800; letter-spacing: 1px; color: #111827;">OPD PRESCRIPTION</div>
              <div style="font-size: 10pt; color: #4b5563; margin-top: 4px;">Date: <strong>${printDate}</strong></div>
              <div style="font-size: 10pt; color: #4b5563;">Encounter: ${data.encounterId || '--'}</div>
              <div style="font-size: 10pt; color: #4b5563; margin-top: 2px;">Consulting Doctor: <strong>${doctorName}</strong></div>
            </div>
          </div>

          <div class="section-break" style="display:flex; justify-content:space-between; font-size:11pt; background:#f9fafb; padding:3mm; border-radius:4px;">
            <div><span class="label">Patient:</span> <strong>${data.patientName || 'Unknown'}</strong> &bull; ${data.patientAgeSex || '--'}</div>
            <div style="text-align: right;"><span class="label">Patient ID:</span> <strong>${data.patientId || '--'}</strong></div>
          </div>

          <div class="section-break" style="display:flex; gap:4mm; flex-wrap:wrap; font-size:10pt;">
            <div><span class="label">BP:</span> <strong>${bp}</strong> mmHg</div>
            <div><span class="label">Pulse:</span> <strong>${vitals.hr || '--'}</strong> bpm</div>
            <div><span class="label">SpO2:</span> <strong>${vitals.spo2 || '--'}</strong> %</div>
            <div><span class="label">Temp:</span> <strong>${vitals.temp || '--'}</strong> °F</div>
            <div><span class="label">Wt:</span> <strong>${vitals.weight || '--'}</strong> kg</div>
          </div>

          <div class="section-break">
            <div style="font-size:11pt; margin-bottom:2mm;"><span class="label">Complaints & History:</span> ${clin.complaints || 'Routine Checkup'}</div>
            <div style="font-size:11pt;"><span class="label">Clinical Diagnosis:</span> <strong>${clin.diagnosis || 'Pending Diagnosis'}</strong></div>
          </div>

          <div class="section-break">
            <div style="font-size: 11pt; font-weight: 800; color: #0369a1; margin-bottom: 3mm; text-transform: uppercase;">TREATMENT PLAN (Rx)</div>
            <table style="width: 100%; border-collapse: collapse; font-size: 10pt;">
              <thead>
                <tr>
                  <th style="background: #f3f4f6; padding: 8px; text-align: left; border-bottom: 2px solid #d1d5db; color: #374151; width:5%;">#</th>
                  <th style="background: #f3f4f6; padding: 8px; text-align: left; border-bottom: 2px solid #d1d5db; color: #374151; width:40%;">MEDICINE</th>
                  <th style="background: #f3f4f6; padding: 8px; text-align: left; border-bottom: 2px solid #d1d5db; color: #374151; width:20%;">DOSAGE</th>
                  <th style="background: #f3f4f6; padding: 8px; text-align: left; border-bottom: 2px solid #d1d5db; color: #374151; width:15%;">DURATION</th>
                  <th style="background: #f3f4f6; padding: 8px; text-align: left; border-bottom: 2px solid #d1d5db; color: #374151; width:20%;">NOTES</th>
                </tr>
              </thead>
              <tbody>
                ${medsHtml}
              </tbody>
            </table>
          </div>

          <div class="section-break" style="display:flex; justify-content:space-between; background:#f9fafb; padding:3mm; border-radius:4px;">
            <div style="width:48%; font-size:10pt;">
               <div style="font-size: 11pt; font-weight: 800; color: #0369a1; margin-bottom: 3mm; text-transform: uppercase;">ORDERS & RESULTS</div>
               <div>${ordersHtml}</div>
               ${resultsHtml}
            </div>
            <div style="width:48%; font-size:10pt;">
               <div style="font-size: 11pt; font-weight: 800; color: #0369a1; margin-bottom: 3mm; text-transform: uppercase;">ADVICE & FOLLOW-UP</div>
               <div>${advStr}</div>
            </div>
          </div>

          <div style="margin-top: 25mm; text-align: right; page-break-inside: avoid;">
            <div style="display: inline-block; text-align: center; min-width: 60mm;">
              <div style="border-top: 1px solid #9ca3af; padding-top: 2mm;">
                <div style="font-size: 11pt; font-weight: 800; color: #111827;">${doctorName}</div>
                ${doctorSpecialty ? `<div style="font-size: 9pt; color: #6b7280; margin-top: 1mm;">${doctorSpecialty}</div>` : ``}
                ${doctorRegNo ? `<div style="font-size: 9pt; color: #4b5563; margin-top: 1mm;">Reg. No: <strong>${doctorRegNo}</strong></div>` : ``}
                <div style="font-size: 9pt; color: #6b7280; margin-top: 2mm;">Doctor's Signature / Seal</div>
              </div>
            </div>
            <div style="font-size: 8pt; color: #9ca3af; text-align: center; margin-top: 15mm;">
              This is a computer-generated medical record.
            </div>
          </div>

        </div>
      </body>
      </html>
    `;

    return { success: true, html: htmlString };

  } catch (error) {
    return { success: false, message: "Backend HTML Error: " + error.toString() };
  }
}

// =========================================================================
// 🚀 OP PRESCRIPTION DISPATCH ENGINE (WHATSAPP & EMAIL)
// =========================================================================

/**
 * Generates OP Prescription PDF, saves to Drive, and returns public link for WhatsApp.
 */
function generateAndStoreOPPrescriptionPDF(encounterId) {
  try {
    // 1. Generate HTML using your existing OP engine
    const reportResponse = getOPPrescriptionHtml(encounterId); 
    if (!reportResponse.success) {
      return { success: false, message: "Could not generate HTML: " + reportResponse.message };
    }

    // 2. Convert to PDF Blob
    const htmlBlob = Utilities.newBlob(reportResponse.html, 'text/html', 'prescription.html');
    const pdfBlob = htmlBlob.getAs('application/pdf');
    pdfBlob.setName("Valarmathi_Prescription_" + encounterId + ".pdf");

    // 3. Drive Folder Architecture
    const rootFolderName = "Valarmathi_OP_Prescriptions";
    let rootFolder;
    const rootFolders = DriveApp.getFoldersByName(rootFolderName);
    
    if (rootFolders.hasNext()) {
      rootFolder = rootFolders.next();
    } else {
      rootFolder = DriveApp.createFolder(rootFolderName);
    }

    const now = new Date();
    const yearStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy");
    const monthStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMMM");

    let yearFolder;
    const yearFolders = rootFolder.getFoldersByName(yearStr);
    if (yearFolders.hasNext()) {
      yearFolder = yearFolders.next();
    } else {
      yearFolder = rootFolder.createFolder(yearStr);
    }

    let monthFolder;
    const monthFolders = yearFolder.getFoldersByName(monthStr);
    if (monthFolders.hasNext()) {
      monthFolder = monthFolders.next();
    } else {
      monthFolder = yearFolder.createFolder(monthStr);
    }

    // 4. Save File & Set Permissions
    const file = monthFolder.createFile(pdfBlob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // 5. Return the Secure Drive Link to Frontend
    return { success: true, link: file.getUrl() };

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Generates OP Prescription PDF and emails it directly via GMAIL API.
 */
function emailOPPrescriptionPDF(encounterId, patientEmail) {
  try {
    const reportResponse = getOPPrescriptionHtml(encounterId); 
    
    if (!reportResponse.success) {
      return { success: false, message: "Could not generate HTML: " + reportResponse.message };
    }

    const htmlBlob = Utilities.newBlob(reportResponse.html, 'text/html', 'prescription.html');
    const pdfBlob = htmlBlob.getAs('application/pdf');
    pdfBlob.setName("Valarmathi_Prescription_" + encounterId + ".pdf");

    const subject = "Your OPD Prescription - Valarmathi Clinic";
    const plainBody = "Dear Patient, please find your prescription attached to this email. Regards, Valarmathi Clinic.";

    const richHtmlBody = `
      <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden;">
        <div style="background-color: #0369a1; color: white; padding: 20px; text-align: center;">
          <h2 style="margin: 0; letter-spacing: 1px;">VALARMATHI CLINIC</h2>
        </div>
        <div style="padding: 30px;">
          <p style="font-size: 16px;">Dear Patient,</p>
          <p style="font-size: 15px; line-height: 1.5;">Please find your official OPD consultation prescription (Encounter: <strong>${encounterId}</strong>) attached to this email.</p>
          
          <div style="background-color: #f0fdf4; border-left: 4px solid #10b981; padding: 15px; margin: 25px 0;">
            <p style="margin: 0; color: #065f46;"><strong>Secure PDF Attached:</strong> You may download it to share with your pharmacy or keep it for your records.</p>
          </div>
          
          <p style="font-size: 14px; color: #4b5563;">Wishing you a speedy recovery,<br><br><strong>The Care Team</strong><br>Valarmathi Clinic</p>
          
          <hr style="border: 0; border-top: 1px solid #e5e7eb; margin: 30px 0 15px 0;">
          <p style="font-size: 11px; color: #9ca3af; text-align: center; margin: 0;">This is an automatically generated dispatch. Please do not reply to this email.</p>
        </div>
      </div>
    `;

    GmailApp.sendEmail(patientEmail, subject, plainBody, {
      htmlBody: richHtmlBody,
      attachments: [pdfBlob],
      name: "Valarmathi Clinic Desk"
    });

    return { success: true, message: "Email sent successfully." };

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}