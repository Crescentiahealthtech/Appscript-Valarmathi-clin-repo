// ==========================================
// Doctors_Engine.gs — Identity, Session, Audit foundation
// Multi-doctor master + server-side session + append-only audit
// ==========================================

// Single-tenant today; swap this one function when Phase 2 multi-tenant lands.
function getTenantId_() { return "VALARMATHI"; }

// ---------- DOCTOR MASTER ----------
function getActiveDoctors() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Doctors");
    if (!sheet) return [];
    const data = sheet.getDataRange().getDisplayValues();
    const tenant = getTenantId_();
    const out = [];
    for (let i = 1; i < data.length; i++) {
      const id = String(data[i][0] || "").trim();
      const status = String(data[i][7] || "").trim().toUpperCase();
      if (!id || status !== "ACTIVE") continue;
      if (String(data[i][1] || "").trim() !== tenant) continue;
      out.push({
        doctorId:  id,
        name:      String(data[i][2] || "").trim(),
        specialty: String(data[i][3] || "").trim(),
        regNo:     String(data[i][4] || "").trim()
      });
    }
    return out;
  } catch (e) { return []; }
}

// Resolve the doctor profile bound to a login username (Linked_Username col, index 6)
function resolveDoctorByUsername_(username) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Doctors");
    if (!sheet) return null;
    const data = sheet.getDataRange().getDisplayValues();
    const u = String(username || "").trim().toUpperCase();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][6] || "").trim().toUpperCase() === u &&
          String(data[i][7] || "").trim().toUpperCase() === "ACTIVE") {
        return { doctorId: String(data[i][0]).trim(), name: String(data[i][2]).trim() };
      }
    }
    return null;
  } catch (e) { return null; }
}

// Admin-run once: creates the sheet and seeds your two existing names so
// historical records keep resolving. Fill Linked_Username after, to bind logins.
function setupDoctorsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Doctors");
  if (!sheet) {
    sheet = ss.insertSheet("Doctors");
    sheet.appendRow(["Doctor_ID","Tenant_ID","Display_Name","Specialty","Reg_No","Signature_Line","Linked_Username","Status"]);
    sheet.getRange("A1:H1").setFontWeight("bold").setBackground("#d9ead3");
  }
  if (sheet.getLastRow() <= 1) {
    const t = getTenantId_();
    sheet.appendRow(["DOC001", t, "Dr. Valarmathi", "General Medicine", "", "Dr. Valarmathi, MBBS", "", "ACTIVE"]);
    sheet.appendRow(["DOC002", t, "Dr. Duty MO",    "Duty Medical Officer", "", "Dr. Duty MO", "", "ACTIVE"]);
  }
  return "Doctors sheet ready. Set Linked_Username for each doctor to bind their login.";
}

// ---------- SERVER-SIDE SESSION ----------
function issueSession_(obj) {
  const token = Utilities.getUuid();
  const payload = JSON.stringify({
    username: obj.username || "",
    role:     (obj.role || "").toLowerCase(),
    doctorId: obj.doctorId || "",
    name:     obj.name || obj.username || "",
    tenantId: getTenantId_()
  });
  CacheService.getScriptCache().put("SESS_" + token, payload, 21600); // 6h
  return token;
}

function validateSession_(token) {
  try {
    if (!token) return null;
    const raw = CacheService.getScriptCache().get("SESS_" + token);
    return raw ? JSON.parse(raw) : null;
  } catch (e) { return null; }
}

// ---------- APPEND-ONLY AUDIT (never throws) ----------
function logAudit_(sess, event, entityType, entityId, details) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Audit_Log");
    if (!sheet) {
      sheet = ss.insertSheet("Audit_Log");
      sheet.appendRow(["Audit_ID","Timestamp","Tenant_ID","Actor_Username","Actor_Role","Actor_Doctor_ID","Event","Entity_Type","Entity_ID","Details_JSON"]);
      sheet.getRange("A1:J1").setFontWeight("bold").setBackground("#fce5cd");
    }
    sheet.appendRow([
      Utilities.getUuid(),
      new Date(),
      getTenantId_(),
      (sess && sess.username) || "UNKNOWN",
      (sess && sess.role) || "",
      (sess && sess.doctorId) || "",
      String(event || ""),
      String(entityType || ""),
      String(entityId || ""),
      JSON.stringify(details || {})
    ]);
  } catch (e) { Logger.log("audit failed: " + e.message); }
}

// Adds a header column to an existing sheet if missing; returns its 1-based index.
function ensureColumn_(sheet, headerName) {
  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = header.indexOf(headerName);
  if (idx !== -1) return idx + 1;
  sheet.getRange(1, lastCol + 1).setValue(headerName).setFontWeight("bold");
  return lastCol + 1;
}