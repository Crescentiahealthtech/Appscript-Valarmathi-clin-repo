// ==========================================
// 🗄️ LAB RECORDS ARCHIVE LOGIC (Search-First)
// ==========================================

// ==========================================
// 🗄️ LAB RECORDS ARCHIVE LOGIC (Search-First Relational)
// ==========================================

function searchLabRecords(query) {
  try {
    if (!query || !String(query).trim()) return { success:false, message:'Enter a Patient ID, Name, or Mobile.' };
    var q = String(query).trim().toUpperCase();
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ---------------------------------------------------------
    // STEP 1: RELATIONAL SEARCH IN PATIENTS DB (For Mobile/Name)
    // ---------------------------------------------------------
    var pSheet = ss.getSheetByName("Patients");
    if (!pSheet || pSheet.getLastRow() < 2) return { success:false, message: "Patients database missing." };

    var pData = pSheet.getDataRange().getValues();
    var pHeaders = pData[0];

    // Find columns dynamically (with fallbacks to your schema)
    var pColId = pHeaders.indexOf("Patient_ID") > -1 ? pHeaders.indexOf("Patient_ID") : 0;
    var pColName = pHeaders.indexOf("Name") > -1 ? pHeaders.indexOf("Name") : 2;
    var pColAge = pHeaders.indexOf("Age") > -1 ? pHeaders.indexOf("Age") : 3;
    var pColGen = pHeaders.indexOf("Gender") > -1 ? pHeaders.indexOf("Gender") : 4;
    var pColMob = pHeaders.indexOf("Mobile") > -1 ? pHeaders.indexOf("Mobile") : 6;
    var pColWa = pHeaders.indexOf("WhatsApp") > -1 ? pHeaders.indexOf("WhatsApp") : 7;
    var pColEmail = pHeaders.indexOf("Email") > -1 ? pHeaders.indexOf("Email") : 16;

    var matchingPids = new Set();
    var patHeader = null;

    for (var i = 1; i < pData.length; i++) {
      var idStr = String(pData[i][pColId] || '').trim().toUpperCase();
      var nameStr = String(pData[i][pColName] || '').trim().toUpperCase();
      var mobStr = String(pData[i][pColMob] || '').trim().toUpperCase();
      var waStr = String(pData[i][pColWa] || '').trim().toUpperCase();

      // Does the query match ID, Name, Mobile, or WhatsApp?
      if (idStr === q || nameStr.indexOf(q) > -1 || mobStr.indexOf(q) > -1 || waStr.indexOf(q) > -1) {
        matchingPids.add(idStr);
        
        // Grab the first match as the Master Header for UI
        if (!patHeader) { 
          patHeader = {
            patientId: String(pData[i][pColId] || ''),
            name: String(pData[i][pColName] || ''),
            age: String(pData[i][pColAge] || ''),
            gender: String(pData[i][pColGen] || '').toUpperCase().charAt(0) || 'M',
            mobile: String(pData[i][pColMob] || ''),
            whatsapp: String(pData[i][pColWa] || pData[i][pColMob] || ''),
            email: String(pData[i][pColEmail] || '')
          };
        }
      }
    }

    // If no patient matches the phone/name, stop here.
    if (matchingPids.size === 0) {
       return { success: true, records: [], patient: null };
    }

    // ---------------------------------------------------------
    // STEP 2: FETCH MATCHING ORDERS 
    // ---------------------------------------------------------
    var oSheet = ss.getSheetByName("LAB_ORDERS");
    if (!oSheet || oSheet.getLastRow() < 2) return { success:true, records:[], patient:patHeader };

    var oMap  = labHeaderMap(oSheet);
    var oData = oSheet.getRange(2,1,oSheet.getLastRow()-1,oSheet.getLastColumn()).getValues();
    var matched = {};

    var REPORTABLE = ['VERIFIED','REPORT_DISPATCHED','AMENDED'];

    oData.forEach(function(r){
      var pidClean = String(r[oMap['PatientID']] || '').trim().toUpperCase();

      // Only process orders that belong to our matched Patient IDs
      if (!matchingPids.has(pidClean)) return;
      if (REPORTABLE.indexOf(String(r[oMap['OrderStatus']]||'')) === -1) return;

      var oid = String(r[oMap['OrderID']]);
      matched[oid] = {
        orderId: oid,
        date:    String(r[oMap['CreatedAt']]||''),
        testNames: String(r[oMap['TestNames']]||''),
        source:  String(r[oMap['SourceModule']]||''),
        doctor:  String(r[oMap['OrderingDoctorName']]||''),
        status:  String(r[oMap['OrderStatus']]||''),
        verifiedBy: '',
        results: []
      };
    });

    var oids = Object.keys(matched);
    if (!oids.length) return { success:true, records:[], patient:patHeader };

    // ---------------------------------------------------------
    // STEP 3: MAP RESULTS TO ORDERS
    // ---------------------------------------------------------
    var rSheet = ss.getSheetByName("LAB_RESULTS");
    if (rSheet && rSheet.getLastRow() >= 2) {
      var rMap  = labHeaderMap(rSheet);
      var rData = rSheet.getRange(2,1,rSheet.getLastRow()-1,rSheet.getLastColumn()).getValues();
      
      rData.forEach(function(r){
        var oid = String(r[rMap['OrderID']]);
        var rec = matched[oid];
        if (!rec) return;

        var isLatest = (r[rMap['IsLatest']]===true || String(r[rMap['IsLatest']]).toUpperCase()==='TRUE');
        var isDraft  = (r[rMap['IsDraft']]===true  || String(r[rMap['IsDraft']]).toUpperCase()==='TRUE');
        if (!isLatest || isDraft) return;

        rec.results.push({
          parameterName: String(r[rMap['ParameterName']]||''),
          value:         String(r[rMap['ResultValue']]||''),
          unit:          String(r[rMap['Unit']]||''),
          flag:          String(r[rMap['Flag']]||''),
          refRangeText:  String(r[rMap['RefRangeText']]||'')
        });
        if (!rec.verifiedBy) rec.verifiedBy = String(r[rMap['VerifiedBy']]||'');
      });
    }

    var records = oids.map(function(k){ return matched[k]; })
                      .sort(function(a,b){ return a.date < b.date ? 1 : -1; });
                      
    return { success:true, records:records, patient:patHeader };

  } catch(err) {
    return { success:false, message:'searchLabRecords failed: ' + err.message };
  }
}

/**
 * Server-Side function: Generates the Lab Report as a PDF and emails it.
 */
/**
 * Server-Side function: Generates the Lab Report as a PDF and emails it via GMAIL API.
 */
function emailLabReportPDF(orderId, patientEmail) {
  try {
    const reportResponse = getLabReportHtml(orderId); 
    
    if (!reportResponse.success) {
      return { success: false, message: "Could not generate report HTML: " + reportResponse.message };
    }

    // Convert to PDF Blob
    const htmlBlob = Utilities.newBlob(reportResponse.html, 'text/html', 'report.html');
    const pdfBlob = htmlBlob.getAs('application/pdf');
    pdfBlob.setName("Crescentia_Clinic_Lab_Report_" + orderId + ".pdf");

    const subject = "Your Verified Lab Report - Crescentia Clinic & Diagnostics";
    const plainBody = "Dear Patient, your verified lab report is attached to this email. Regards, Crescentia Clinic.";

    // Premium HTML Email Layout for GmailApp
    const richHtmlBody = `
      <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden;">
        <div style="background-color: #1e3a8a; color: white; padding: 20px; text-align: center;">
          <h2 style="margin: 0; letter-spacing: 1px;">CRESCENTIA CLINIC & DIAGNOSTICS</h2>
        </div>
        <div style="padding: 30px;">
          <p style="font-size: 16px;">Dear Patient,</p>
          <p style="font-size: 15px; line-height: 1.5;">Your laboratory investigation report (<strong>${orderId}</strong>) is ready.</p>
          <p style="font-size: 15px; line-height: 1.5;">This report has been digitally signed and carries an NABL/ISO compliant Attestation Hash to guarantee its authenticity.</p>
          
          <div style="background-color: #f0fdf4; border-left: 4px solid #10b981; padding: 15px; margin: 25px 0;">
            <p style="margin: 0; color: #065f46;"><strong>Secure PDF Attached:</strong> Please find your official report attached to this email. You may download it to share with your consulting doctor.</p>
          </div>
          
          <p style="font-size: 14px; color: #4b5563;">Wishing you the best of health,<br><br><strong>The Care Team</strong><br>Crescentia Clinic & Diagnostics</p>
          
          <hr style="border: 0; border-top: 1px solid #e5e7eb; margin: 30px 0 15px 0;">
          <p style="font-size: 11px; color: #9ca3af; text-align: center; margin: 0;">This is an automatically generated dispatch. Please do not reply to this email.</p>
        </div>
      </div>
    `;

    // Dispatch via Gmail API
    GmailApp.sendEmail(patientEmail, subject, plainBody, {
      htmlBody: richHtmlBody,
      attachments: [pdfBlob],
      name: "Crescentia Diagnostics"
    });

    return { success: true, message: "Email sent successfully via Gmail API." };

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Server-Side function: Generates PDF, saves to Drive, and returns public link.
 */
function generateAndStoreLabReportPDF(orderId) {
  try {
    console.log("1. Starting PDF generation for Order: " + orderId);

    // 1. Generate HTML using your existing engine
    const reportResponse = getLabReportHtml(orderId); 
    if (!reportResponse.success) {
      console.error("HTML Generation Failed: " + reportResponse.message);
      return { success: false, message: "Could not generate HTML: " + reportResponse.message };
    }

    console.log("2. HTML generated successfully. Converting to PDF Blob...");

    // 2. Convert to PDF Blob
    const htmlBlob = Utilities.newBlob(reportResponse.html, 'text/html', 'report.html');
    const pdfBlob = htmlBlob.getAs('application/pdf');
    pdfBlob.setName("Crescentia_Report_" + orderId + ".pdf");

    console.log("3. PDF Created. Searching for Drive Folders...");

    // 3. Drive Folder Architecture
    const rootFolderName = "Crescentia_Lab_Reports";
    let rootFolder;
    const rootFolders = DriveApp.getFoldersByName(rootFolderName);
    
    if (rootFolders.hasNext()) {
      rootFolder = rootFolders.next();
    } else {
      rootFolder = DriveApp.createFolder(rootFolderName);
    }

    // Safely format Dates using Apps Script Native Utilities
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

    console.log("4. Folders mapped. Saving file to Drive...");

    // 4. Save File & Set Permissions
    const file = monthFolder.createFile(pdfBlob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    console.log("5. Success! File URL: " + file.getUrl());

    // 5. Return the Secure Drive Link to Frontend
    return { 
      success: true, 
      link: file.getUrl() 
    };

  } catch (error) {
    console.error("CRITICAL BACKEND ERROR: " + error.toString());
    return { success: false, message: error.toString() };
  }
}

function forceDriveAuthorization() {
  // Run this function ONCE from the editor to approve Drive permissions
  DriveApp.createFolder("Test_Crescentia_Init");
  Logger.log("Drive Authorization Successful!");
}

// ==========================================
// 🗄️ LAB RECORDS INVOICE DISPATCH ENGINE
// ==========================================

function getArchiveInvoiceLink(orderId) {
  try {
    // 1. Get the HTML Invoice (Uses the existing function that handles IP logic)
    const reportResponse = getLabBillHtml(orderId); 
    if (!reportResponse.success) throw new Error("Invoice HTML Generation Failed");

    // 2. Convert to PDF
    const htmlBlob = Utilities.newBlob(reportResponse.html, 'text/html', 'invoice.html');
    const pdfBlob = htmlBlob.getAs('application/pdf');
    pdfBlob.setName("Crescentia_Lab_Invoice_" + orderId + ".pdf");

    // 3. Save to the dedicated Invoices Folder
    const rootFolders = DriveApp.getFoldersByName("Crescentia_Lab_Invoices");
    const rootFolder = rootFolders.hasNext() ? rootFolders.next() : DriveApp.createFolder("Crescentia_Lab_Invoices");

    const now = new Date();
    const yearStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy");
    const monthStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMMM");

    let yearFolder = rootFolder.getFoldersByName(yearStr).hasNext() ? rootFolder.getFoldersByName(yearStr).next() : rootFolder.createFolder(yearStr);
    let monthFolder = yearFolder.getFoldersByName(monthStr).hasNext() ? yearFolder.getFoldersByName(monthStr).next() : yearFolder.createFolder(monthStr);

    const file = monthFolder.createFile(pdfBlob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return { success: true, link: file.getUrl() };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function emailArchiveInvoice(orderId, patientEmail) {
  try {
    const reportResponse = getLabBillHtml(orderId); 
    if (!reportResponse.success) throw new Error("Invoice HTML Generation Failed");

    const htmlBlob = Utilities.newBlob(reportResponse.html, 'text/html', 'invoice.html');
    const pdfBlob = htmlBlob.getAs('application/pdf');
    pdfBlob.setName("Crescentia_Lab_Invoice_" + orderId + ".pdf");

    const subject = "Your Lab Invoice - Crescentia Clinic & Diagnostics";
    const plainBody = "Dear Patient, please find your lab invoice attached.";

    const richHtmlBody = `
      <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden;">
        <div style="background-color: #0d9488; color: white; padding: 20px; text-align: center;">
          <h2 style="margin: 0; letter-spacing: 1px;">CRESCENTIA CLINIC & DIAGNOSTICS</h2>
        </div>
        <div style="padding: 30px;">
          <p style="font-size: 16px;">Dear Patient,</p>
          <p style="font-size: 15px; line-height: 1.5;">Please find your official payment receipt/invoice (Order: <strong>${orderId}</strong>) attached to this email.</p>
          <p style="font-size: 14px; color: #4b5563;">Wishing you the best of health,<br><br><strong>The Care Team</strong></p>
        </div>
      </div>
    `;

    GmailApp.sendEmail(patientEmail, subject, plainBody, {
      htmlBody: richHtmlBody,
      attachments: [pdfBlob],
      name: "Crescentia Billing Desk"
    });

    return { success: true, message: "Invoice emailed successfully." };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}