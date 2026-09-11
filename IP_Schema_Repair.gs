// ============================================================================
// IP_Schema_Repair.gs  —  Crescentia HealthTech
// One-off repairs for schema drift already present in the live workbook.
// Run from the Apps Script editor. Every function reports what it changed and
// is safe to re-run.
// ============================================================================

/**
 * REPAIRS IP_CaseSheets_DB HEADER DRIFT.
 *
 * The live sheet carries:
 *   ... | Patient Name | Age/Sex | Timestamp | Sys_BP | ...
 * where the code expects "Age" and "Sex". Because those two names were absent,
 * dc_ensureSheet_ appended fresh Age and Sex columns at the END of the row, so
 * new casesheets write age and sex to columns 38/39 while every historical row
 * still holds a combined "24 / Male" in column 8. That is why patients show as
 * "- / -" in the IP records list.
 *
 * There is also a SECOND column literally headed "Timestamp". dc_headerMap_
 * keeps the first occurrence, so the duplicate is dead weight that silently
 * receives nothing.
 *
 * This function:
 *   1. splits the legacy "Age/Sex" values into the real Age and Sex columns,
 *      without overwriting anything already there;
 *   2. renames the two orphaned headers to Legacy_* so the header map is
 *      unambiguous and dc_ensureSheet_ stops treating them as candidates.
 * Nothing is deleted — the legacy columns keep their data.
 */
function repairCasesheetHeaderDrift() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName("IP_CaseSheets_DB");
    if (!sh) return "IP_CaseSheets_DB not found.";

    var lastCol = sh.getLastColumn();
    var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0]
                .map(function (h) { return String(h || "").trim(); });
    var report = [];

    // --- locate the players -------------------------------------------------
    var idxAgeSex = hdr.indexOf("Age/Sex");
    var idxAge    = hdr.indexOf("Age");
    var idxSex    = hdr.indexOf("Sex");

    // The SECOND "Timestamp", if there is one.
    var firstTs = hdr.indexOf("Timestamp");
    var dupTs   = (firstTs === -1) ? -1 : hdr.indexOf("Timestamp", firstTs + 1);

    if (idxAge === -1 || idxSex === -1) {
      return "Age and/or Sex columns are missing. Open the casesheet once so " +
             "the schema is ensured, then re-run this.";
    }

    // --- 1. backfill Age and Sex from the combined legacy column -----------
    var lastRow = sh.getLastRow();
    if (idxAgeSex !== -1 && lastRow > 1) {
      var n = lastRow - 1;
      var legacy = sh.getRange(2, idxAgeSex + 1, n, 1).getValues();
      var ages   = sh.getRange(2, idxAge + 1, n, 1).getValues();
      var sexes  = sh.getRange(2, idxSex + 1, n, 1).getValues();
      var filled = 0;

      for (var i = 0; i < n; i++) {
        var combined = String(legacy[i][0] || "").trim();
        if (!combined) continue;
        var haveAge = String(ages[i][0] || "").trim();
        var haveSex = String(sexes[i][0] || "").trim();
        if (haveAge && haveSex) continue;      // already migrated

        // "24 / Male", "24/Male", "24 Y / Male" all appear in the wild.
        var parts = combined.split("/");
        var a = String(parts[0] || "").replace(/[^0-9]/g, "");
        var s = String(parts[1] || "").trim();

        if (!haveAge && a) { ages[i][0]  = a; filled++; }
        if (!haveSex && s) { sexes[i][0] = s; }
      }
      if (filled) {
        sh.getRange(2, idxAge + 1, n, 1).setValues(ages);
        sh.getRange(2, idxSex + 1, n, 1).setValues(sexes);
      }
      report.push("Age/Sex: backfilled " + filled + " row(s) from the legacy column.");
    } else if (idxAgeSex === -1) {
      report.push("Age/Sex: no legacy column present, nothing to migrate.");
    }

    // --- 2. rename the orphans so the header map is unambiguous ------------
    if (idxAgeSex !== -1) {
      sh.getRange(1, idxAgeSex + 1).setValue("Legacy_Age_Sex");
      report.push("Renamed 'Age/Sex' -> 'Legacy_Age_Sex' (data kept).");
    }
    if (dupTs !== -1) {
      sh.getRange(1, dupTs + 1).setValue("Legacy_Timestamp_2");
      report.push("Renamed the duplicate 'Timestamp' -> 'Legacy_Timestamp_2' (data kept).");
    }

    dc_invalidate_("IP_CaseSheets_DB");
    SpreadsheetApp.flush();

    if (report.length === 0) report.push("Nothing needed repair.");
    var out = report.join("\n");
    Logger.log(out);
    return out;

  } catch (e) {
    return "Repair failed: " + e.message;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Runs every IP verification and repair in the right order and returns one
 * report. This is the ship gate: read it before letting a ward round onto a
 * new deployment.
 */
function runIPHealthCheck() {
  var out = [];
  out.push("========== IP SCHEMA ==========");
  out.push(verifyIPClinicalSchema());
  out.push("");
  out.push("========== CASESHEET HEADER DRIFT ==========");
  out.push(repairCasesheetHeaderDrift());
  out.push("");
  out.push("========== CARE TEAM DUPLICATES ==========");
  out.push(repairDuplicateCareTeamRows());
  out.push("");
  out.push("========== DUPLICATE DRUG ORDERS ==========");
  out.push(repairDuplicatePharmacyQueueRows());
  out.push("");
  out.push("========== CARE TEAM COVERAGE ==========");
  out.push(verifyIPCareTeamCoverage());
  var report = out.join("\n");
  Logger.log(report);
  return report;
}

/**
 * Deactivates duplicate ACTIVE rows in IP_Pharmacy_Queue, keeping the newest
 * for each drug on each admission. These come from a casesheet saved twice
 * before the double-submit guard existed, and from STOP orders that failed to
 * match because the name was written with the strength in front of it.
 * Rows are marked DUPLICATE, never deleted.
 */
function repairDuplicatePharmacyQueueRows() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("IP_Pharmacy_Queue");
    if (!sh) return "IP_Pharmacy_Queue not found.";

    var data = sh.getDataRange().getValues();
    if (data.length < 2) return "No pharmacy queue rows.";

    var seen = {}, fixed = [];
    // Newest first, so the row that survives is the most recent order.
    for (var i = data.length - 1; i > 0; i--) {
      if (String(data[i][11] || "ACTIVE").toUpperCase() !== "ACTIVE") continue;
      var key = String(data[i][1]).trim().toUpperCase() + "||" + _normDrug_(data[i][3]);
      if (!seen[key]) { seen[key] = true; continue; }

      sh.getRange(i + 1, 12).setValue("DUPLICATE");
      sh.getRange(i + 1, 13).setValue(new Date());
      sh.getRange(i + 1, 14).setValue("SYSTEM_REPAIR");
      fixed.push(String(data[i][0]) + "  " + String(data[i][3]) + "  (" + String(data[i][1]) + ")");
    }
    dc_invalidate_("IP_Pharmacy_Queue");
    SpreadsheetApp.flush();

    var report = fixed.length
      ? "Deactivated " + fixed.length + " duplicate order row(s):\n  " + fixed.join("\n  ")
      : "No duplicate pharmacy queue rows found.";
    Logger.log(report);
    return report;
  } catch (e) {
    return "Repair failed: " + e.message;
  } finally {
    lock.releaseLock();
  }
}