// ==========================================
// 🧠 SYSTEM CORE & ROUTER - MV WORKSPACE
// ==========================================

function doGet(e) {
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

function getUserProfile(patientId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Patients");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().toUpperCase() === patientId.toUpperCase()) {
        return {
          id: data[i][0],
          name: data[i][2], 
          age: data[i][3], 
          gender: data[i][4], 
          dob: (data[i][5] instanceof Date) ? Utilities.formatDate(data[i][5], Session.getScriptTimeZone(), "yyyy-MM-dd") : data[i][5].toString(),
          mobile: data[i][6], 
          whatsapp: data[i][7], 
          address: data[i][8], 
          comorb: data[i][9],
          email: data[i][10] || "" // <-- NEW: Grabs Column K (Index 10)
        };
      }
    }
    return null;
  } catch (e) { 
    return null; 
  }
}

function registerPatient(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Patients');
    
    // 1. Generate Patient ID
    // (If row is 1, it becomes LMTVS0001, etc.)
    const newId = "LMTVS" + (sheet.getLastRow()).toString().padStart(4, '0');
    
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

function formatTimeSafely(timeVal) {
  if(!timeVal) return "";
  if(timeVal instanceof Date) { return Utilities.formatDate(timeVal, Session.getScriptTimeZone(), "hh:mm a").toUpperCase(); }
  let t = String(timeVal).trim().toUpperCase();
  t = t.replace(/([0-9])(AM|PM)/, "$1 $2"); return t;
}

function getPatientDashboardStats(patientId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const apptSheet = ss.getSheetByName('Appointments');
    let lastVisit = "None Recorded";
    let upcomingBookings = [];
    
    if (apptSheet) {
      const apptData = apptSheet.getDataRange().getValues();
      let today = new Date(); today.setHours(0,0,0,0);
      for(let i = 1; i < apptData.length; i++) {
        if(apptData[i][1] === patientId) {
          let rawDate = apptData[i][3];
          let apptDateObj = (rawDate instanceof Date) ? new Date(rawDate) : new Date(rawDate);
          apptDateObj.setHours(0,0,0,0);
          let dateStr = (rawDate instanceof Date) ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "dd MMM yyyy") : String(rawDate).substring(0,10);
          let timeStr = formatTimeSafely(apptData[i][4]);
          let status = apptData[i][6];
          
          if(apptDateObj >= today && (status === 'Booked' || status === 'Arrived' || status === 'In-Progress')) {
            upcomingBookings.push({ dateVal: apptDateObj, display: `${dateStr} • ${timeStr}` });
          }
          if(apptDateObj < today && status === 'Completed') lastVisit = dateStr; 
        }
      }
      upcomingBookings.sort((a,b) => a.dateVal - b.dateVal);
    }
    
    const emrSheet = ss.getSheetByName('EMR_Records');
    let emrNextVisit = "Awaiting Doctor's Update";
    let emrNextLab = "Awaiting Doctor's Update";
    
    if(emrSheet) {
        const emrData = emrSheet.getDataRange().getValues();
        for(let i = emrData.length - 1; i >= 1; i--) {
            if(emrData[i][2] === patientId) {
                let lines = String(emrData[i][8]).split('\n');
                lines.forEach(line => {
                    let l = line.toLowerCase();
                    if(l.includes('follow-up') || l.includes('next visit')) emrNextVisit = line.replace(/follow-up|next visit|:/gi, '').trim() || "See EMR Plan";
                    if(l.includes('lab') || l.includes('blood test') || l.includes('investigation')) emrNextLab = line.replace(/lab visit|lab date|next lab|investigations|:/gi, '').trim() || "See EMR Plan";
                }); break;
            }
        }
    }
    return { lastVisit: lastVisit, upcomingBookings: upcomingBookings.map(b => b.display), emrNextVisit: emrNextVisit, emrNextLab: emrNextLab };
  } catch (e) { return { lastVisit: "Error", upcomingBookings: [], emrNextVisit: "Error", emrNextLab: "Error" }; }
}

function getAvailableTimeSlots(dateStr) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Appointments');
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const standardSlots = [
    "10:00 AM", "10:15 AM", "10:30 AM", "10:45 AM", "11:00 AM", "11:15 AM", "11:30 AM", "11:45 AM", "12:00 PM", "12:15 PM", "12:30 PM", "12:45 PM", "05:00 PM", "05:15 PM", "05:30 PM", "05:45 PM", "06:00 PM", "06:15 PM", "06:30 PM", "06:45 PM", "07:00 PM", "07:15 PM", "07:30 PM", "07:45 PM", "08:00 PM", "08:15 PM", "08:30 PM", "08:45 PM"
  ];
  const takenSlots = [];
  for (let i = 1; i < data.length; i++) {
    let dObj = data[i][3];
    let rowDate = (dObj instanceof Date) ? Utilities.formatDate(dObj, Session.getScriptTimeZone(), "yyyy-MM-dd") : dObj.toString().substring(0,10);
    if (rowDate === dateStr && data[i][6] !== 'Cancelled' && data[i][6] !== 'DELETE') takenSlots.push(formatTimeSafely(data[i][4]));
  }
  return standardSlots.filter(slot => !takenSlots.includes(slot));
}

function getAppointmentsByDate(dateStr) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Appointments');
    if(!sheet) return [];
    const data = sheet.getDataRange().getValues();
    let appts = [];
    for(let i = 1; i < data.length; i++) {
      let dObj = data[i][3];
      let rowDate = (dObj instanceof Date) ? Utilities.formatDate(dObj, Session.getScriptTimeZone(), "yyyy-MM-dd") : (dObj ? dObj.toString().substring(0,10) : "");
      if(rowDate === dateStr) {
        appts.push({ apptId: data[i][0], patientId: data[i][1], patientName: data[i][2], time: formatTimeSafely(data[i][4]), purpose: data[i][5], status: data[i][6], fee: data[i][7] });
      }
    } return appts;
  } catch(e) { return []; }
}

function fetchDailyLedger(dateStr) {
  const apptSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Appointments');
  const patientSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Patients');
  if(!apptSheet || !patientSheet) return [];
  const apptData = apptSheet.getDataRange().getValues();
  const patientData = patientSheet.getDataRange().getValues();
  
  const patientMap = {};
  for (let i = 1; i < patientData.length; i++) {
    patientMap[patientData[i][0].toString().toUpperCase()] = { name: patientData[i][2], age: patientData[i][3], sex: patientData[i][4] };
  }

  const ledger = [];
  for (let i = 1; i < apptData.length; i++) {
    let dObj = apptData[i][3];
    let rowDate = (dObj instanceof Date) ? Utilities.formatDate(dObj, Session.getScriptTimeZone(), "yyyy-MM-dd") : dObj.toString().substring(0,10);
    if (rowDate === dateStr) {
      let pId = apptData[i][1].toString().toUpperCase();
      if (pId === "ADMIN") continue;
      ledger.push({
        apptId: apptData[i][0], time: formatTimeSafely(apptData[i][4]), patientId: pId,
        patientName: apptData[i][2] || (patientMap[pId] ? patientMap[pId].name : '-'),
        age: patientMap[pId] ? patientMap[pId].age : '-', sex: patientMap[pId] ? patientMap[pId].sex : '-',
        purpose: apptData[i][5], status: apptData[i][6]
      });
    }
  } return ledger;
}

function submitNewAppointment(apptObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Appointments');
    const data = sheet.getDataRange().getValues();

    if (apptObj.patientId !== 'ADMIN' && apptObj.patientId !== 'DIRECT' && apptObj.patientId !== 'WALK-IN') {
      for (let i = 1; i < data.length; i++) {
        let dObj = data[i][3];
        let rowDate = (dObj instanceof Date) ? Utilities.formatDate(dObj, Session.getScriptTimeZone(), "yyyy-MM-dd") : dObj.toString().substring(0,10);
        if (rowDate === apptObj.date && data[i][1] === apptObj.patientId) {
          let status = data[i][6];
          if(status === 'Booked' || status === 'Arrived' || status === 'In-Progress') return { success: false, message: "You already have an active appointment scheduled for this date." };
        }
        if (rowDate === apptObj.date && formatTimeSafely(data[i][4]) === apptObj.time) return { success: false, message: "Slot collision. Time was just booked by another user." };
      }
    }
    const newId = "APT-" + Date.now().toString().slice(-6);
    sheet.appendRow([newId, apptObj.patientId, apptObj.patientName, apptObj.date, apptObj.time, apptObj.purpose, apptObj.status || 'Booked', apptObj.fee || 0, new Date().toISOString()]);
    SpreadsheetApp.flush(); return { success: true, apptId: newId };
  } catch(e) { return { success: false, message: e.message }; } finally { lock.releaseLock(); }
}

function updateAppointmentStatus(apptId, newStatus) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Appointments');
  const data = sheet.getDataRange().getValues();
  for(let i = 1; i < data.length; i++) {
    if(data[i][0] == apptId) {
      if(newStatus === "DELETE") sheet.deleteRow(i + 1);
      else sheet.getRange(i + 1, 7).setValue(newStatus);
      SpreadsheetApp.flush(); return "Status updated!";
    }
  } return "Error updating.";
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

function getPatientDemographics(patientId) { return getUserProfile(patientId); }

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

function getUserProfile(patientId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Patients");
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      // Find the row where Column A matches the Patient ID
      if (data[i][0].toString().toUpperCase() === patientId.toUpperCase()) {
        return {
          id: data[i][0],             // A
          password: data[i][1],       // B
          name: data[i][2],           // C
          age: data[i][3],            // D
          gender: data[i][4],         // E
          dob: (data[i][5] instanceof Date) ? Utilities.formatDate(data[i][5], Session.getScriptTimeZone(), "yyyy-MM-dd") : data[i][5].toString(), // F
          mobile: data[i][6],         // G
          whatsapp: data[i][7],       // H
          address: data[i][8],        // I
          comorb: data[i][9],         // J
          
          // === NEW EXTENDED FIELDS ===
          regDate: data[i][10] ? data[i][10].toString() : "", // K
          salutation: data[i][11] || "",                      // L
          maritalStatus: data[i][12] || "",                   // M
          bloodGroup: data[i][13] || "",                      // N
          occupation: data[i][14] || "",                      // O
          education: data[i][15] || "",                       // P
          email: data[i][16] || "",                           // Q
          relationType: data[i][17] || "",                    // R
          relationName: data[i][18] || "",                    // S
          emergencyName: data[i][19] || "",                   // T
          emergencyNumber: data[i][20] || "",                 // U
          referredBy: data[i][21] || ""                       // V
        };
      }
    }
    return null;
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
