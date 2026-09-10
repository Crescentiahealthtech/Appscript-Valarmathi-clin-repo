// ==========================================
// IP_Casesheet_Logic.gs (Upgraded OP-Parity Engine)
// ==========================================

/**
 * Ward list for the casesheet patient picker.
 * Doctor-scoped: a consultant sees the beds they are responsible for or
 * consulting on; operational roles see the whole ward.
 * Returns a JSON string for backwards compatibility with the existing UI.
 */
function getIPAdmissions(sessionToken) {
  try {
    var gate = resolveIPRead_(sessionToken, null);
    if (!gate.ok) return JSON.stringify({ error: gate.message });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("IP_Admissions");

    if (!sheet) return JSON.stringify([]); // clean empty state — never fabricate patients

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return JSON.stringify([]);

    const canSee = ipc_wardVisibilityFilter_(gate.scope);
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][11] !== "ACTIVE") continue;
      if (!canSee(data[i][0])) continue;
      result.push({
        ipNumber: data[i][0],
        id: data[i][1],
        name: data[i][2],
        age: data[i][3] || "--",
        sex: data[i][4] || "--",
        doa: data[i][5] ? Utilities.formatDate(new Date(data[i][5]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "--",
        ward: data[i][7] || "Ward",
        bed: data[i][8],
        triage: data[i][13] || "Stable",
        consultant: data[i][9] || ""
      });
    }
    return JSON.stringify(result);
  } catch (e) {
    return JSON.stringify({ error: e.toString() });
  }
}

/**
 * Writes the admission casesheet.
 *
 * Phase 5: the author must be an active doctor ON THE CARE TEAM for this
 * admission — the primary consultant, a cross-consult, or a covering doctor
 * flagged Can_View_All. The signature is snapshotted into the row so a later
 * rename in the Doctors sheet cannot rewrite a signed record.
 *
 * @param {Object} payload
 * @param {string} sessionToken  preferred; payload.sessionToken still accepted
 *                               so an older cached client keeps working.
 */
function saveIPRecord(payload, sessionToken) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    payload = payload || {};
    const token = sessionToken || payload.sessionToken;

    // --- Server-side identity, care-team gate & attribution -------------
    const w = resolveIPWrite_(token, payload.ipNumber, "DOCTOR", {});
    if (!w.ok) return { success: false, message: w.message };
    if (w.role !== 'doctor') {
      return { success: false, message: "Only a logged-in doctor can author a casesheet." };
    }
    const sess = w.sess;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Owns the schema, including the Phase 5 attribution columns. Creating
    // the sheet inline here is what let the header drift out of step with
    // the row being written.
    const sheet = ipc_casesheetSheet_();
    const rxSheet = ss.getSheetByName("IP_Pharmacy_Queue"); // Route IP meds here

    const timestamp = new Date();
    const dateStr = Utilities.formatDate(timestamp, ss.getSpreadsheetTimeZone(), "MM/dd/yyyy hh:mm a");
    const uniqueHash = timestamp.getTime().toString().slice(-6);
    const encounterId = "IP-ENC-" + payload.patientId + "-" + uniqueHash;

    const vitals  = payload.vitals  || {};
    const genExam = payload.genExam || {};
    const sysExam = payload.sysExam || {};
    const geFlags = Array.isArray(genExam.flags) ? genExam.flags : [];
    const hasGE = (flag) => geFlags.indexOf(flag) !== -1 ? "Yes" : "No";

    const medsJSON = JSON.stringify(payload.meds || []);
    const labsJSON = JSON.stringify(payload.labs || []);

    // Header-driven write. Column order in IP_CaseSheets_DB is no longer
    // implied by the position of a value in an array literal.
    const m = dc_headerMap_(sheet);
    const row = new Array(sheet.getLastColumn()).fill("");
    const put = function (header, value) {
      if (m[header] !== undefined) row[m[header]] = (value === undefined || value === null) ? "" : value;
    };

    put("Encounter_ID", encounterId);
    put("IP_Number", payload.ipNumber);
    put("Ward", payload.ward);
    put("Bed", payload.bed);
    put("Patient_ID", payload.patientId);
    put("Timestamp", dateStr);
    put("Patient Name", payload.patientName);
    put("Age", payload.age);
    put("Sex", payload.sex);
    put("Sys_BP", vitals.sys);
    put("Dia_BP", vitals.dia);
    put("PR", vitals.hr);
    put("SpO2", vitals.spo2);
    put("Temp", vitals.temp);
    put("Height", vitals.height);
    put("Weight", vitals.weight);
    put("Chief_Complaints", JSON.stringify(payload.complaints || []));
    put("History", JSON.stringify(payload.history || []));
    put("Pallor", hasGE('Pallor'));
    put("Icterus", hasGE('Icterus'));
    put("Cyanosis", hasGE('Cyanosis'));
    put("Clubbing", hasGE('Clubbing'));
    put("Edema", hasGE('Edema'));
    put("Other GE findings", genExam.notes);
    put("CVS", sysExam.cvs);
    put("RS", sysExam.rs);
    put("PA", sysExam.pa);
    put("CNS", sysExam.cns);
    put("Primary Diagnosis", payload.provDiagnosis);
    put("Prescription_JSON", medsJSON);
    put("Lab_Orders_JSON", labsJSON);
    put("Outside Lab Records", JSON.stringify(payload.outsideLabs || []));
    put("Radiological records", payload.radiology);
    put("Advice", payload.advice);
    put("Doctor's Name", w.displayName);
    put("Doctor_ID", w.doctorId);
    put("Author_Signature_Snapshot", w.signature);
    put("Author_Username", w.username);

    sheet.appendRow(row);

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
          w.authorLabel,
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
          orderingDoctorName: w.displayName,
          orderingDoctorId:   w.doctorId,
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

    try {
      logAudit_(sess, "CASESHEET_SAVE", "IP_CaseSheet", encounterId, {
        patientId: payload.patientId, ipNumber: payload.ipNumber,
        diagnosis: payload.provDiagnosis, doctorId: w.doctorId
      });
    } catch (e) { /* auditing must never fail a clinical save */ }

    SpreadsheetApp.flush();
    return {
      success: true,
      message: "IP Casesheet Locked & Saved to DB Successfully!",
      encounterId: encounterId,
      doctorName: w.displayName,
      signature: w.signature
    };

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
function getIPCasesheetHtml(encounterId, sessionToken) {
  try {
    const sheet = ipc_casesheetSheet_();
    const m = dc_headerMap_(sheet);
    const data = sheet.getDataRange().getDisplayValues();

    let record = null;
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === encounterId) { record = data[i]; break; }
    }
    if (!record) return { success: false, message: "Casesheet record not found." };

    // Read by header, never by position: the attribution columns are appended
    // by the migration and their index differs between deployments.
    const f = function (header) {
      const idx = m[header];
      return (idx === undefined) ? "" : String(record[idx] || "");
    };

    // A signed casesheet is only printable by someone entitled to the
    // admission it belongs to.
    const gate = resolveIPRead_(sessionToken, f("IP_Number"));
    if (!gate.ok) return { success: false, message: gate.message };

    // Every value below is interpolated into markup — escape it. Free-text
    // fields such as Advice and the diagnosis are typed by clinicians and
    // routinely contain "<" and "&".
    const e = function (v) {
      return String(v === null || v === undefined ? '' : v)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;')
        .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    };
    const fe = function (header) { return e(f(header)); };

    let complaints = [], history = [], meds = [], labs = [], outsideLabs = [];
    try { complaints  = JSON.parse(f("Chief_Complaints")    || "[]"); } catch(err){}
    try { history     = JSON.parse(f("History")             || "[]"); } catch(err){}
    try { meds        = JSON.parse(f("Prescription_JSON")   || "[]"); } catch(err){}
    try { labs        = JSON.parse(f("Lab_Orders_JSON")     || "[]"); } catch(err){}
    try { outsideLabs = JSON.parse(f("Outside Lab Records") || "[]"); } catch(err){}

    const ccStr = complaints.map(c => `${e(c.condition)} (${e(c.duration)})`).join(" | ");
    const hxStr = history.map(h => `${e(h.prefix)} ${e(h.condition)} (${e(h.duration)})`).join(" | ");

    const medsHtml = meds.length
      ? `<ol style="margin:0; padding-left: 20px;">` + meds.map(md =>
          `<li style="margin-bottom:6px;"><strong>${e(md.type)} ${e(md.drugName)}</strong><br>` +
          `<span style="color:#555; font-size:0.85em;">${e(md.sig)} | ${e(md.duration)} | ${e(md.comments)}</span></li>`
        ).join('') + `</ol>`
      : `<span style="color:#777;">No admission medication ordered.</span>`;

    const internalLabs = labs.filter(l => l.source === 'INTERNAL' || l.type === 'Order')
                             .map(l => e(l.testName)).join(", ");
    const extLabsHtml = outsideLabs.length
      ? outsideLabs.map(l => `${e(l.test)}: ${e(l.val)}`).join(" | ")
      : "None";

    // The signature snapshotted at save time, falling back to the recorded
    // name for rows written before the Phase 5 migration.
    const signature = f("Author_Signature_Snapshot") || f("Doctor's Name") || "Doctor's Signature";

    const html = `
    <div style="font-family: Arial, sans-serif; color: #000; padding: 20px; max-width: 800px; margin: auto;">
        <div style="border-bottom: 2px solid #0369a1; padding-bottom: 10px; margin-bottom: 20px; text-align: center;">
            <h2 style="margin:0; text-transform:uppercase; font-weight:bold; color: #0369a1;">Valarmathi Clinic</h2>
            <p style="margin:0; font-size: 0.9rem; color: #555;">Premium Healthcare Services | Ph: +91 88387 23513</p>
            <h4 style="margin-top:10px; color: #111;">IP Admission Casesheet</h4>
        </div>

        <div style="display: flex; justify-content: space-between; margin-bottom: 20px; font-size: 0.95rem; background: #f9fafb; padding: 15px; border: 1px solid #e5e7eb; border-radius: 6px;">
            <div>
                <strong>Patient:</strong> ${fe("Patient Name")}<br>
                <strong>PID:</strong> ${fe("Patient_ID")} | <strong>IP No:</strong> ${fe("IP_Number")}<br>
                <strong>Age/Sex:</strong> ${fe("Age")} Y / ${fe("Sex")}<br>
                <strong>Ward/Bed:</strong> ${fe("Ward")} - ${fe("Bed")}
            </div>
            <div style="text-align: right;">
                <strong>Date:</strong> ${fe("Timestamp")}<br>
                <strong>Doctor:</strong> ${fe("Doctor's Name") || '--'}<br>
                <span style="font-size:0.8em; color:#555;">${fe("Doctor_ID")}</span>
            </div>
        </div>

        <div style="margin-bottom: 20px; font-size: 0.9rem;">
            <h5 style="border-bottom: 1px solid #ccc; padding-bottom: 5px; color: #0369a1;">Vitals on Admission</h5>
            <p style="margin: 5px 0;"><strong>BP:</strong> ${fe("Sys_BP")}/${fe("Dia_BP")} mmHg &nbsp;|&nbsp; <strong>Pulse:</strong> ${fe("PR")} bpm &nbsp;|&nbsp; <strong>SpO2:</strong> ${fe("SpO2")}% &nbsp;|&nbsp; <strong>Temp:</strong> ${fe("Temp")} °F &nbsp;|&nbsp; <strong>Wt:</strong> ${fe("Weight")} kg</p>
        </div>

        <div style="margin-bottom: 20px; font-size: 0.9rem;">
            <h5 style="border-bottom: 1px solid #ccc; padding-bottom: 5px; color: #0369a1;">Clinical History &amp; Examination</h5>
            <p style="margin: 5px 0;"><strong>Chief Complaints:</strong> ${ccStr || '--'}</p>
            <p style="margin: 5px 0;"><strong>History:</strong> ${hxStr || '--'}</p>
            <p style="margin: 5px 0;"><strong>General Exam:</strong> Pallor: ${fe("Pallor")}, Icterus: ${fe("Icterus")}, Cyanosis: ${fe("Cyanosis")}, Clubbing: ${fe("Clubbing")}, Edema: ${fe("Edema")}<br><em>Notes:</em> ${fe("Other GE findings") || '--'}</p>
            <p style="margin: 5px 0;"><strong>Systemic Exam:</strong> CVS: ${fe("CVS")} | RS: ${fe("RS")} | P/A: ${fe("PA")} | CNS: ${fe("CNS")}</p>
            <p style="margin: 5px 0;"><strong>Primary Diagnosis:</strong> <span style="font-size:1.1em; font-weight:bold;">${fe("Primary Diagnosis") || '--'}</span></p>
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
                <p style="margin: 5px 0; white-space: pre-wrap;">${fe("Radiological records") || 'None'}</p>
                <h5 style="border-bottom: 1px solid #ccc; padding-bottom: 5px; margin-top: 15px; color: #0369a1;">Advice / Plan</h5>
                <p style="margin: 5px 0; white-space: pre-wrap;">${fe("Advice") || 'Standard ward protocol.'}</p>
            </div>
        </div>

        <div style="text-align: right; margin-top: 60px;">
            <div style="border-top: 1px solid #000; display: inline-block; padding-top: 5px; width: 220px; text-align: center;">
                <strong>${e(signature)}</strong>
            </div>
        </div>
    </div>
    `;

    return { success: true, html: html };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}
