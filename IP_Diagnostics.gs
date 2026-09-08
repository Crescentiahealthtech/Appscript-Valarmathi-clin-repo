// ==========================================
// IP_Diagnostics.gs — READ ONLY. No writes, no lock needed.
// Run auditIPAdmissionsSheet() from the Apps Script editor,
// then open View > Logs.
// ==========================================

function auditIPAdmissionsSheet() {
  const out = { success: true, sheetExists: false, headers: [], rowCount: 0,
                statusTally: {}, patients: [], warnings: [] };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("IP_Admissions");
    if (!sheet) {
      out.warnings.push("Sheet 'IP_Admissions' does not exist.");
      Logger.log(JSON.stringify(out, null, 2));
      return out;
    }
    out.sheetExists = true;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    out.headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
                       .map(function (h) { return String(h).trim(); });

    const EXPECTED = ["IP Number","Patient ID","Patient Name","Age/Sex","DOA","TOA",
                      "Type","Ward","Bed","Consultant","Diagnosis","Status","DOD"];
    for (var e = 0; e < EXPECTED.length; e++) {
      if (String(out.headers[e] || "").toLowerCase() !== EXPECTED[e].toLowerCase()) {
        out.warnings.push("Col " + (e + 1) + ": expected '" + EXPECTED[e] +
                          "', found '" + (out.headers[e] || "(empty)") + "'");
      }
    }

    if (lastRow < 2) {
      out.warnings.push("Sheet has headers but zero data rows. Ledger is genuinely empty.");
      Logger.log(JSON.stringify(out, null, 2));
      return out;
    }

    const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const seen = {};
    for (var i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      out.rowCount++;

      const ip = String(r[0]).trim();
      const status = String(r[11] || "(blank)").trim().toUpperCase();
      out.statusTally[status] = (out.statusTally[status] || 0) + 1;

      if (seen[ip]) out.warnings.push("DUPLICATE IP Number '" + ip +
                                      "' on sheet rows " + seen[ip] + " and " + (i + 2));
      else seen[ip] = i + 2;

      out.patients.push({
        row:        i + 2,
        ipNumber:   ip,
        patientId:  String(r[1] || "").trim(),
        name:       String(r[2] || "").trim(),
        ageSex:     String(r[3] || "").trim(),
        doaRaw:     String(r[4]),
        doaIsDate:  (r[4] instanceof Date),
        ward:       String(r[7] || "").trim(),
        bed:        String(r[8] || "").trim(),
        consultant: String(r[9] || "").trim(),
        status:     status
      });
    }
    Logger.log(JSON.stringify(out, null, 2));
    return out;
  } catch (err) {
    out.success = false;
    out.message = err.message;
    Logger.log("AUDIT FAILED: " + err.message);
    return out;
  }
}