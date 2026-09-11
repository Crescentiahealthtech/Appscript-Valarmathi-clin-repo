// ==========================================
// 🧠 SYSTEM CORE & ROUTER - MV WORKSPACE
// ==========================================

function doGet(e) {
  // 0. Discharge summary verification. Anonymous by design: anyone holding a
  //    printed summary can confirm it is genuine. The page carries no clinical
  //    content — see dsx_verifyPage_ in DS_Print.gs.
  if (e && e.parameter && e.parameter.verifyDS) {
    if (typeof dsx_verifyPage_ === 'function') {
      return dsx_verifyPage_(e.parameter.verifyDS);
    }
    return HtmlService.createHtmlOutput('Verification is not available on this deployment.');
  }

  // 1. WhatsApp / Patient Mobile Interceptor
  if (e && e.parameter && e.parameter.viewReport) {
    const orderId = e.parameter.viewReport;
    const reportData = getLabReportHtml(orderId);
    
    if (reportData.success) {
       // Convert HTML to PDF, then to Base64 String for mobile downloading
       const htmlBlob = Utilities.newBlob(reportData.html, 'text/html', 'report.html');
       const pdfBlob = htmlBlob.getAs('application/pdf');
       const base64Data = Utilities.base64Encode(pdfBlob.getBytes());
       const fileName = "Valarmathi_Report_" + orderId + ".pdf";

       // Ultra-lightweight Mobile Download Page (Will NOT crash on phones)
       const mobileDownloadHtml = `
         <!DOCTYPE html>
         <html>
         <head>
           <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
           <title>Download Report</title>
           <style>
             body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: #f8fafc; padding: 40px 20px; margin: 0; display: flex; justify-content: center; align-items: center; min-height: 100vh; }
             .card { background: white; padding: 40px 25px; border-radius: 16px; box-shadow: 0 4px 20px rgba(0,0,0,0.08); text-align: center; max-width: 350px; width: 100%; border: 1px solid #e5e7eb; }
             .icon { font-size: 50px; margin-bottom: 15px; }
             .btn { background: #10b981; color: white; border: none; padding: 18px 20px; border-radius: 12px; font-size: 16px; font-weight: bold; width: 100%; cursor: pointer; margin-top: 25px; display: block; box-sizing: border-box; box-shadow: 0 4px 6px rgba(16,185,129,0.2); }
             .btn:active { transform: scale(0.96); background: #059669; }
           </style>
         </head>
         <body>
           <div class="card">
             <div class="icon">📄</div>
             <h2 style="color:#0f172a; margin: 0 0 10px 0;">Report Ready</h2>
             <p style="color:#64748b; font-size: 14px; margin: 0 0 10px 0;">Order ID: <strong>${orderId}</strong></p>
             <p style="color:#94a3b8; font-size: 12px; margin: 0 0 20px 0;">Valarmathi Clinic & Diagnostics</p>
             <button id="dlBtn" class="btn">📥 Download PDF Report</button>
           </div>
           
           <script>
             document.getElementById('dlBtn').addEventListener('click', function() {
                this.innerText = 'Downloading...';
                // Trigger native file download on the mobile device
                const link = document.createElement('a');
                link.href = "data:application/pdf;base64,${base64Data}";
                link.download = "${fileName}";
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                
                setTimeout(() => { this.innerText = '✅ Download Complete'; }, 2500);
             });
           </script>
         </body>
         </html>
       `;
       
       return HtmlService.createHtmlOutput(mobileDownloadHtml)
         .setTitle('Secure Lab Report')
         .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
         
    } else {
       return HtmlService.createHtmlOutput("<h2 style='text-align:center; font-family:Arial; margin-top:50px; color:#ef4444;'>Report Not Found or Invalid Link</h2>");
    }
  }

  // 2. Normal App Load for Staff
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Valarmathi Clinic Enterprise')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) { 
  return HtmlService.createHtmlOutputFromFile(filename).getContent(); 
}

function registerPatient(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Patients');
    
    // 1. Generate Patient ID — Barcode_Engine.gs
    // Never reuses an ID after row deletion (the old getLastRow() scheme did,
    // which would make a printed patient barcode open the wrong record).
    const newId = bc_nextPatientId_(sheet);
    
    // 2. BACKEND PASSWORD GENERATOR (Name 3 char + YYYY)
    let namePart = data.name ? data.name.toString().trim().replace(/[^a-zA-Z]/g, '') : "UNK";
    if (namePart.length < 3) namePart = (namePart + "XXX").substring(0, 3);
    else namePart = namePart.substring(0, 3);
    namePart = namePart.charAt(0).toUpperCase() + namePart.substring(1).toLowerCase();
    
    let yearPart = "0000";
    if (data.dob) yearPart = data.dob.toString().substring(0, 4);
    let generatedPassword = namePart + yearPart;

    // 3. Current Timestamp for Registration Date
    const regDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");

    // 4. APPEND ROW - STRICTLY MAPPED TO PRESERVE COLUMNS A THROUGH J, NEW FIELDS K TO V
    sheet.appendRow([
      newId,                      // A: ID
      generatedPassword,          // B: Password
      data.name || "",            // C: Name
      data.age || "",             // D: Age
      data.gender || "",          // E: Gender
      data.dob || "",             // F: DOB
      data.mobile || "",          // G: Mobile
      data.whatsapp || "",        // H: WhatsApp
      data.address || "",         // I: Address
      data.comorb || "Nil",       // J: Conditions
      regDate,                    // K: Registration_Date (NEW)
      data.salutation || "",      // L: Salutation (NEW)
      data.maritalStatus || "",   // M: Marital_Status (NEW)
      data.bloodGroup || "",      // N: Blood_Group (NEW)
      data.occupation || "",      // O: Occupation (NEW)
      data.education || "",       // P: Education (NEW)
      data.email || "",           // Q: Email (NEW)
      data.relationType || "",    // R: Relation_Type (NEW)
      data.relationName || "",    // S: Relation_Name (NEW)
      data.emergencyName || "",   // T: Emergency_Contact_Name (NEW)
      data.emergencyNumber || "", // U: Emergency_Number (NEW)
      data.referredBy || ""       // V: Referred_By (NEW)
    ]);
    
    SpreadsheetApp.flush(); 
    return { 
      success: true, 
      patientId: newId,          // so the registration screen can print the card
      message: `Patient Registered Successfully!\nID: ${newId}\nPassword: ${generatedPassword}` 
    };
    
  } catch(e) { 
    return { success: false, message: "Error: " + e.message }; 
  } finally { 
    lock.releaseLock(); 
  }
}

function getAllPatients() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Patients');
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  
  let roster = [];
  // Start at 1 to skip headers
  for(let i = 1; i < data.length; i++) {
    if(data[i][0]) {
      roster.push({ 
        id: data[i][0].toString(),            // A: Patient ID
        name: data[i][2].toString(),          // C: Name
        age: data[i][3].toString(),           // D: Age
        gender: data[i][4].toString(),        // E: Gender
        mobile: data[i][6].toString(),        // G: Mobile
        address: data[i][8] ? data[i][8].toString() : "" // I: Address
      });
    }
  }
  return roster;
}

function saveAdminAvailability(dateStr, blockedSlots) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Appointments');
    const data = sheet.getDataRange().getValues();
    const rowsToDelete = [];
    for(let i = data.length - 1; i >= 1; i--) {
      let dObj = data[i][3];
      let rowDate = (dObj instanceof Date) ? Utilities.formatDate(dObj, Session.getScriptTimeZone(), "yyyy-MM-dd") : dObj.toString().substring(0,10);
      if(rowDate === dateStr && data[i][6] === 'Blocked') rowsToDelete.push(i + 1);
    }
    rowsToDelete.forEach(r => sheet.deleteRow(r));
    if (blockedSlots.length > 0) {
      const newRows = [];
      let startId = sheet.getLastRow();
      let timestamp = new Date().toISOString();
      blockedSlots.forEach((slot, index) => {
        let apptId = "APT-" + (startId + index).toString().padStart(4, '0');
        newRows.push([apptId, 'ADMIN', 'BLOCKED', dateStr, slot, 'Doctor Unavailable', 'Blocked', 0, timestamp]);
      });
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }
    SpreadsheetApp.flush(); return {success: true, message: 'Availability Updated!'};
  } catch(e) { return {success: false, message: 'Failed to save availability.'}; } finally { lock.releaseLock(); }
}

function saveEMRRecord(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EMR_Records');
  if(!sheet) return {success: false, message: "Create 'EMR_Records' sheet first."};
  let timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
  sheet.appendRow([timestamp, data.apptId, data.patientId, data.bp, data.pulse, data.weight, data.complaints, data.diagnosis, data.plan]);
  SpreadsheetApp.flush(); return {success: true, message: "Clinical Notes Saved Successfully!"};
}

// Drop these updated functions into CodeMV.gs

function initializeDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = [
    { name: "Users", headers: ["Username", "Password", "Role"] },
    // NEW SCHEMA: Password is now explicitly Column B
    { name: "Patients", headers: ["Patient ID", "Password", "Name", "Age", "Gender", "DOB", "Mobile", "WhatsApp", "Address", "Comorbidities"] },
    { name: "Appointments", headers: ["Appt ID", "Patient ID", "Patient Name", "Date", "Time", "Purpose", "Status", "Fee", "Timestamp"] },
    { name: "Blocked_Slots", headers: ["Date", "Blocked Times (Comma Separated)"] },
    { name: "EMR_Records", headers: ["Timestamp", "Appt ID", "Patient ID", "BP", "Pulse", "Weight", "Complaints", "Diagnosis", "Plan"] },
    { name: "Lab_Orders", headers: ["Order ID", "Date", "Appt ID", "Patient ID", "Tests", "Status", "Total Cost"] },
    { name: "Pharmacy_Inventory", headers: ["Drug ID", "Drug Name", "Stock", "Price"] },
    { name: "Billing_Ledger", headers: ["Invoice ID", "Date", "Appt ID", "Patient ID", "Consult Fee", "Pharmacy Fee", "Lab Fee", "Grand Total", "Status"] }
  ];

  requiredSheets.forEach(schema => {
    let sheet = ss.getSheetByName(schema.name);
    if (!sheet) {
      sheet = ss.insertSheet(schema.name);
      sheet.appendRow(schema.headers);
      sheet.getRange(1, 1, 1, schema.headers.length).setFontWeight("bold").setBackground("#22262d").setFontColor("#ffffff");
    }
  });
}

/**
 * Internal patient-profile reader. NO session check — every caller must
 * enforce its own access rule. The password column (B) is never read here,
 * so no caller can leak it by accident.
 */
function pt_readProfile_(patientId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Patients");
    if (!sheet || sheet.getLastRow() < 2) return null;
    const want = String(patientId || "").trim().toUpperCase();
    if (!want) return null;

    // TextFinder instead of getDataRange(): a 10,000-row patient master is not
    // read into memory to answer a single-ID lookup.
    const cell = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
      .createTextFinder(want).matchEntireCell(true).matchCase(false).findNext();
    if (!cell) return null;

    const row = sheet.getRange(cell.getRow(), 1, 1, Math.max(22, sheet.getLastColumn())).getValues()[0];
    return {
      id: row[0],                 // A
      // Column B is the patient portal password. Deliberately NOT returned.
      name: row[2],               // C
      age: row[3],                // D
      gender: row[4],             // E
      dob: (row[5] instanceof Date) ? Utilities.formatDate(row[5], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(row[5] || ""), // F
      mobile: row[6],             // G
      whatsapp: row[7],           // H
      address: row[8],            // I
      comorb: row[9],             // J
      regDate: row[10] ? row[10].toString() : "", // K
      salutation: row[11] || "",                  // L
      maritalStatus: row[12] || "",               // M
      bloodGroup: row[13] || "",                  // N
      occupation: row[14] || "",                  // O
      education: row[15] || "",                   // P
      email: row[16] || "",                       // Q
      relationType: row[17] || "",                // R
      relationName: row[18] || "",                // S
      emergencyName: row[19] || "",               // T
      emergencyNumber: row[20] || "",             // U
      referredBy: row[21] || ""                   // V
    };
  } catch (e) {
    return null;
  }
}

/**
 * Patient profile for the browser. SESSION REQUIRED.
 *
 * Previously this took only a patient ID, ran with no session check, and
 * returned the portal password alongside DOB, mobile and address. With the web
 * app deployed as "anyone, even anonymous" and patient IDs sequential, anyone
 * holding the /exec URL could walk LMTVS0001, 0002 ... from the console and
 * harvest credentials. Printed patient barcodes make those IDs public, so this
 * is now closed:
 *   - staff session   -> any patient, minus the password column
 *   - patient session -> their own record only
 *   - no valid session -> null
 */
function getUserProfile(patientId, sessionToken) {
  try {
    const sess = dc_validateSession_(sessionToken);
    if (!sess) return null;

    const role = String(sess.role || "").trim().toLowerCase();
    const want = String(patientId || "").trim().toUpperCase();
    if (!want) return null;

    // A patient may read only themselves. Their session username IS their ID.
    if (role === "patient" && String(sess.username || "").trim().toUpperCase() !== want) {
      return null;
    }
    return pt_readProfile_(want);
  } catch (e) {
    return null;
  }
}

function saveUserProfile(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Patients');
    const records = sheet.getDataRange().getValues();
    for (let i = 1; i < records.length; i++) {
      if (records[i][0].toString().toUpperCase() === data.username.toUpperCase()) {
        // Shifted columns by +1 to account for Password in Column B
        sheet.getRange(i + 1, 3).setValue(data.name); 
        sheet.getRange(i + 1, 4).setValue(data.age);
        sheet.getRange(i + 1, 5).setValue(data.gender); 
        sheet.getRange(i + 1, 6).setValue(data.dob);
        sheet.getRange(i + 1, 7).setValue(data.mobile); 
        sheet.getRange(i + 1, 8).setValue(data.whatsapp);
        sheet.getRange(i + 1, 9).setValue(data.address); 
        sheet.getRange(i + 1, 10).setValue(data.comorb);
        SpreadsheetApp.flush(); return "Profile Updated Successfully.";
      }
    }
    return "Error: User not found.";
  } catch (e) { return "Error: " + e.message; } finally { lock.releaseLock(); }
}