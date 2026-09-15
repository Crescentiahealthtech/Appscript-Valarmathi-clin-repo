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

  // 0b. Lab report verification. The same contract as the discharge summary
  //     above: anonymous, no clinical content, one question answered. The QR
  //     printed on every released report points here. It used to point at
  //     https://valarmathi.clinic/verify, which does not exist, and was drawn
  //     by a Google endpoint that was switched off years ago — so the block
  //     on the report was an empty frame beside a promise nothing kept. See
  //     _labVerifyBlock_ in LabIntegrationEngine.gs.
  if (e && e.parameter && e.parameter.verifyLab) {
    if (typeof labVerifyPage_ === 'function') {
      return labVerifyPage_(e.parameter.verifyLab, e.parameter.c);
    }
    return HtmlService.createHtmlOutput('Verification is not available on this deployment.');
  }

  // 0c. The privacy notice and the request form: ?privacy
  //     Section 5 requires the notice AT OR BEFORE collection, and ss.11-13
  //     give the rights to a person, not to a person with a staff login. A
  //     form only reachable from inside the application is a way of receiving
  //     fewer requests, not of answering them. Print this URL on the
  //     registration slip and the invoice footer.
  if (e && e.parameter && (e.parameter.privacy !== undefined ||
                           e.parameter.rights !== undefined)) {
    return HtmlService.createHtmlOutputFromFile('DPDP_Privacy_Page')
      .setTitle('Your information and your rights')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // 1. A document shared with a patient: ?doc=<grant>&k=<key>
  //    The file itself stays private in Drive. This route holds the expiry,
  //    the open counter and the audit row. See DPDP_Documents.gs.
  if (e && e.parameter && e.parameter.doc) {
    if (typeof dpdpServeDocument_ === 'function') {
      return dpdpServeDocument_(e.parameter.doc, e.parameter.k);
    }
    return HtmlService.createHtmlOutput('Document delivery is not available on this deployment.');
  }

  // 1b. The old WhatsApp report link: ?viewReport=<orderId>
  //
  //     WITHDRAWN, and deliberately not made to work again. It took an order
  //     id — sequential, printed on the patient's own paperwork — and
  //     returned that order's COMPLETE LAB REPORT as a PDF to anyone who
  //     asked, with no key, no session and no expiry. Walking the ids
  //     returned the lab's whole output.
  //
  //     Anyone still holding one of those messages gets this instead, and the
  //     clinic re-sends the document as an expiring, revocable link.
  if (e && e.parameter && e.parameter.viewReport) {
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><meta name="viewport" content="width=device-width, initial-scale=1">' +
      '<title>This link has been withdrawn</title></head>' +
      '<body style="font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Roboto,sans-serif;' +
      'background:#f8fafc;margin:0;padding:40px 20px;text-align:center;color:#334155;">' +
      '<div style="max-width:380px;margin:0 auto;background:#fff;padding:36px 24px;' +
      'border:1px solid #e5e7eb;border-radius:16px;">' +
      '<h2 style="font-size:1.15rem;color:#0f172a;margin:0 0 10px;">This link has been withdrawn</h2>' +
      '<p style="font-size:.9rem;line-height:1.55;color:#64748b;margin:0;">' +
      'Report links of this kind never expired, so we have stopped honouring them. ' +
      'Please contact the clinic and we will send your report again as a private ' +
      'link that only works for a short time.</p></div></body></html>')
      .setTitle('This link has been withdrawn');
  }

  // 2. Normal App Load for Staff
  //
  // The probe goes HERE and not at the top of doGet, deliberately. The four
  // routes above are anonymous BY DESIGN — a printed verification QR, the
  // public privacy notice, a patient's document link — so an unidentified
  // caller on any of them says nothing about the deployment's access mode.
  // This route is the application itself, and whether its callers arrive
  // identified is the one measurable fact about the live deployment's access
  // setting. See Deployment_Probe.gs.
  try { if (typeof depProbeRecord_ === 'function') depProbeRecord_(); } catch (e) {}

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Valarmathi Clinic Enterprise')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) { 
  return HtmlService.createHtmlOutputFromFile(filename).getContent(); 
}

/**
 * FRONTEND ENTRY. Registers a patient, gives them a portal credential, and
 * records what they were asked and what they answered.
 *
 * THREE THINGS CHANGED HERE, ALL OF THEM DPDP FINDINGS.
 *
 * C1/C2 — it took no session token and checked nothing. On a web app
 * deployed as ANYONE_ANONYMOUS that is an unauthenticated write to the
 * patient master by anyone holding the URL.
 *
 * C3 — the portal password was `Mei2001`: the first three letters of the name
 * and the birth year, stored in clear in column B and printed on the
 * registration slip. Both inputs are on documents the patient carries, so it
 * was not a secret. It is now random, stored only as a digest, shown to the
 * desk ONCE to hand over, and has to be changed at first sign-in.
 *
 * H2/M1 — nothing recorded that consent had ever been asked for, and three of
 * the twenty-two fields collected (education, occupation, marital status) were
 * read by nothing in the application. Section 6(1) permits collection for a
 * SPECIFIED purpose; a field that exists because the form had a box has no
 * purpose to specify. They are no longer collected. Consent decisions taken at
 * the desk are written to the Consent_Register in the same call, so the notice
 * version they were given against is the one on the row.
 *
 * @param {Object} data              the registration form
 * @param {string} sessionToken      the desk's session
 * @param {Object} [consent]         { decisions:{PURPOSE:bool}, guardianName,
 *                                     guardianRelation, method }
 */
function registerPatient(data, sessionToken, consent) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const actor = crescRequire_(sessionToken, 'patient.register');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Patients');

    // 1. Generate Patient ID — Barcode_Engine.gs
    // Never reuses an ID after row deletion (the old getLastRow() scheme did,
    // which would make a printed patient barcode open the wrong record).
    const newId = bc_nextPatientId_(sheet);

    // 2. A random portal password, stored as a salted digest and never again
    //    readable from the sheet. Returned once, below, for the desk to hand
    //    over — and flagged so the patient must replace it at first sign-in.
    const portalPassword = crescRandomPassword_(10);

    // 3. Current Timestamp for Registration Date
    const regDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");

    // 4. APPEND ROW - STRICTLY MAPPED TO PRESERVE COLUMNS A THROUGH J, NEW FIELDS K TO V
    //    M, O and P are written empty on purpose: see the note above.
    sheet.appendRow([
      newId,                      // A: ID
      crescPwdEncode_(portalPassword), // B: Password — a digest, not a password
      data.name || "",            // C: Name
      data.age || "",             // D: Age
      data.gender || "",          // E: Gender
      data.dob || "",             // F: DOB
      data.mobile || "",          // G: Mobile
      data.whatsapp || "",        // H: WhatsApp
      data.address || "",         // I: Address
      data.comorb || "Nil",       // J: Conditions
      regDate,                    // K: Registration_Date
      data.salutation || "",      // L: Salutation
      "",                         // M: Marital_Status — no longer collected (DPDP M1)
      data.bloodGroup || "",      // N: Blood_Group
      "",                         // O: Occupation — no longer collected (DPDP M1)
      "",                         // P: Education — no longer collected (DPDP M1)
      data.email || "",           // Q: Email
      data.relationType || "",    // R: Relation_Type
      data.relationName || "",    // S: Relation_Name
      data.emergencyName || "",   // T: Emergency_Contact_Name
      data.emergencyNumber || "", // U: Emergency_Number
      data.referredBy || ""       // V: Referred_By
    ]);

    // The digest column must be text, or Sheets reformats a value that starts
    // with a recognisable pattern and the stored credential stops matching.
    const newRow = sheet.getLastRow();
    sheet.getRange(newRow, 2).setNumberFormat('@');
    try {
      const psheet = cresc_patientsSheet_();       // adds the two flag columns
      const flagCol = cresc_colOf_(psheet, 'Portal_Must_Change');
      if (flagCol !== -1) psheet.getRange(newRow, flagCol).setValue('YES');
      const stampCol = cresc_colOf_(psheet, 'Portal_Password_Updated_At');
      if (stampCol !== -1) psheet.getRange(newRow, stampCol).setValue(new Date());
    } catch (e) { /* the credential is already stored; the flag is best effort */ }

    SpreadsheetApp.flush();

    // 5. Consent, against the notice version in force right now. Section
    //    6(10) puts the burden of proving consent on the clinic, and a
    //    consent nobody wrote down cannot be proved.
    let consentResult = null;
    if (consent && consent.decisions) {
      try {
        consentResult = recordConsent({
          patientId: newId,
          decisions: consent.decisions,
          method: consent.method || 'IN_PERSON',
          guardianName: consent.guardianName,
          guardianRelation: consent.guardianRelation,
          notes: 'Captured at registration'
        }, sessionToken);
      } catch (e) {
        consentResult = { success: false, message: e.message };
      }
    }

    try {
      logAudit_({ username: actor.username, role: actor.role },
                'PATIENT_REGISTERED', 'Patient', newId,
                { consentRecorded: !!(consentResult && consentResult.success) });
    } catch (e) { /* best effort */ }

    return {
      success: true,
      patientId: newId,          // so the registration screen can print the card
      portalPassword: portalPassword,
      consent: consentResult,
      message: 'Patient registered. ID: ' + newId +
               '\nOne-time portal password: ' + portalPassword +
               '\nIt is shown once and must be changed at first sign-in.'
    };

  } catch(e) {
    return { success: false, message: String(e.message || e).replace('FORBIDDEN: ', '') };
  } finally {
    lock.releaseLock();
  }
}

/**
 * The whole patient register.
 *
 * THIS TOOK NO SESSION TOKEN AND CHECKED NOTHING. appsscript.json deploys
 * this web app as executeAs USER_DEPLOYING / access ANYONE_ANONYMOUS, so
 * anybody holding the /exec URL could open a console and call
 *
 *     google.script.run.withSuccessHandler(console.log).getAllPatients()
 *
 * without signing in, and receive every patient's name, age, sex, mobile
 * number and home address — running with the deploying account's full
 * access to the spreadsheet, not the caller's.
 *
 * Under the DPDP Act 2023 that is a failure of section 8(5) (reasonable
 * security safeguards) over the entire patient register, and the resulting
 * disclosure is a reportable personal data breach.
 *
 * It is now behind patient.read, like every other patient endpoint. The
 * caller's screen already had a session token; it simply never sent it.
 *
 * @param {string} sessionToken
 */
function getAllPatients(sessionToken) {
  try {
    crescRequire_(sessionToken, 'patient.read');
    // Finding M2: a read of the ENTIRE register is the one read that must
    // never be silent. Who pulled the directory, and when.
    dpdpLogRead_(crescActor_(sessionToken), 'PatientRegister', 'ALL',
                 { endpoint: 'getAllPatients' });
  } catch (e) {
    // The directory screen renders whatever array it gets, so an empty one
    // with a message is the shape it can already handle.
    return [];
  }

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
      let timestamp = new Date().toISOString();
      blockedSlots.forEach((slot, index) => {
        // Row-count-derived ids, generated immediately after deleting rows
        // from the same sheet: getLastRow() has just dropped, so the next
        // block hands out ids that are already on real bookings. Appointment
        // .gs's generator is unique and does not care how many rows exist.
        let apptId = apt_newId_();
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

  // Hospital billing owns its own schema and seeds the service tariff, so it
  // sets itself up rather than being listed above. Without this the billing
  // desk opens on "Service_Master is empty" after a fresh install.
  try { if (typeof setupHospitalBilling === 'function') setupHospitalBilling(); } catch (e) {}
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

/**
 * The patient portal's own profile editor.
 *
 * THIS TOOK NO SESSION TOKEN EITHER, and keyed on a `username` the caller
 * supplied. On an anonymously deployed web app that is an unauthenticated
 * WRITE to any patient's name, age, date of birth, mobile number and
 * address — patient IDs are sequential and printed on barcodes, so the
 * identifier needed to target it is public.
 *
 * It is now what it always claimed to be: a patient editing THEIR OWN
 * record. Staff corrections go through updatePatientProfile()
 * (Patient_Profile_Edit.gs), which is guarded by patient.write and logs
 * every change with its previous value.
 *
 * @param {Object} data           the fields to write
 * @param {string} sessionToken   the portal session
 */
function saveUserProfile(data, sessionToken) {
  const lock = LockService.getScriptLock();
  try {
    data = data || {};

    var actor = null;
    try { actor = crescActor_(sessionToken); } catch (e) { actor = null; }
    if (!actor) return "Error: your session has expired. Please sign in again.";

    // A patient may write only their own record; their session username IS
    // their patient ID. Staff must use the audited endpoint.
    if (actor.role === 'patient') {
      if (String(actor.username || '').trim().toUpperCase() !==
          String(data.username || '').trim().toUpperCase()) {
        return "Error: you can only change your own profile.";
      }
    } else if (actor.permissions.indexOf('patient.write') === -1) {
      return "Error: your role cannot edit patient records.";
    }

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