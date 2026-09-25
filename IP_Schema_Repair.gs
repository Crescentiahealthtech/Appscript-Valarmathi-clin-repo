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
  crescEditorOnly_('repairCasesheetHeaderDrift');
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
  crescEditorOnly_('runIPHealthCheck');
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
  crescEditorOnly_('repairDuplicatePharmacyQueueRows');
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
/**
 * NORMALISES EVERY DATE CELL THAT IS STILL TEXT.
 *
 * The parser in Shared_Dates.gs reads a text date correctly on the way out,
 * so every screen is already right. This puts the workbook right too, which
 * matters for the three things that read the sheet without going through
 * our code: the user sorting a column by hand, a spreadsheet formula, and a
 * CSV export. A column that mixes real dates with "13/09/2026" sorts all the
 * text below all the dates, whichever direction you pick.
 *
 * Safe to re-run. Reports every cell it changed and every cell it could not
 * read, and CHANGES NOTHING on a dry run - call it with no argument first
 * and read the report before letting it write.
 *
 *   normaliseSheetDates()       -> report only, writes nothing
 *   normaliseSheetDates(true)   -> applies the same changes
 *
 * @param {boolean} [apply=false]
 * @return {string} a human-readable report
 */
function normaliseSheetDates(apply) {
  crescEditorOnly_('normaliseSheetDates');
  var TARGETS = [
    { sheet: 'IP_Admissions',      columns: ['DOA', 'DOD'] },
    { sheet: 'Master_Beds',        columns: ['DOA'] },
    { sheet: 'Patients',           columns: ['DOB', 'Registered_On', 'Registration_Date'] },
    { sheet: 'Accounts_Ledger',    columns: ['Timestamp', 'Date'] },
    { sheet: 'Accounts_Payables',  columns: ['Due_Date', 'Timestamp'] },
    { sheet: 'Accounts_Insurance', columns: ['Date'] }
  ];

  if (typeof cresc_parseDate_ !== 'function') {
    return 'Shared_Dates.gs is not in this project - copy it across first.';
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lines = [apply ? 'APPLYING changes:' : 'DRY RUN - nothing written. Call normaliseSheetDates(true) to apply.'];
  var changed = 0, unreadable = 0;

  TARGETS.forEach(function (t) {
    var sh = ss.getSheetByName(t.sheet);
    if (!sh) { lines.push('  ' + t.sheet + ': not in this workbook, skipped.'); return; }
    var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return;

    var header = sh.getRange(1, 1, 1, lastCol).getValues()[0]
                   .map(function (h) { return String(h || '').trim(); });

    t.columns.forEach(function (name) {
      var col = header.indexOf(name);
      if (col === -1) return;

      var n = lastRow - 1;
      var range = sh.getRange(2, col + 1, n, 1);
      var values = range.getValues();
      var out = [], touched = 0, bad = [];

      for (var i = 0; i < n; i++) {
        var v = values[i][0];
        if (v === '' || v === null || v === undefined) { out.push(['']); continue; }
        if (v instanceof Date && !isNaN(v.getTime())) { out.push([v]); continue; }
        var d = cresc_parseDate_(v);
        // A time-only cell belongs to a time column that happens to be listed
        // here; leaving it alone is right, and rewriting it as 1899 is not.
        if (!d || cresc_isSheetEpoch_(d)) {
          out.push([v]);
          if (!d) { bad.push('row ' + (i + 2) + ': "' + v + '"'); unreadable++; }
          continue;
        }
        out.push([d]);
        touched++;
      }

      if (touched && apply) {
        range.setValues(out);
        range.setNumberFormat('dd-mmm-yyyy');
      }
      changed += touched;
      lines.push('  ' + t.sheet + '.' + name + ': ' + touched + ' text cell(s) ' +
                 (apply ? 'converted to real dates' : 'would be converted') +
                 (bad.length ? ', ' + bad.length + ' unreadable' : ''));
      bad.slice(0, 10).forEach(function (b) { lines.push('      unreadable ' + b); });
      if (bad.length > 10) lines.push('      ... and ' + (bad.length - 10) + ' more');
    });
  });

  if (apply) SpreadsheetApp.flush();
  lines.push('');
  lines.push('Total: ' + changed + ' cell(s) ' + (apply ? 'converted' : 'to convert') +
             ', ' + unreadable + ' left alone because no date could be read from them.');
  var report = lines.join('\n');
  Logger.log(report);
  return report;
}


// ---------------------------------------------------------------------------
// OP_ENCOUNTERS: THE VITALS HEADERS
//
// saveOPEncounter_ writes columns D-J positionally (systolic, diastolic,
// pulse, SpO2, temperature, weight, BMI), and every reader in this project
// reads them by position too — so the data is fine. But the clinic's sheet
// has "Sys_BP" written over all seven headers, which is what any export,
// any data-principal request and anybody reading the sheet sees. This puts
// the right names back. It changes row 1 only, and only where a header in
// D-J is a duplicate "Sys_BP" or blank.
// ---------------------------------------------------------------------------

var OPE_VITALS_HEADERS = ['Sys_BP', 'Dia_BP', 'PR', 'SpO2', 'Temp', 'Weight', 'BMI'];   // D..J

/** Editor. dryRun defaults to TRUE: pass false to write. */
function repairOPEncounterHeaders(dryRun) {
  crescEditorOnly_('repairOPEncounterHeaders');
  dryRun = (dryRun !== false);
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('OP_Encounters');
  if (!sh) return 'OP_Encounters: absent.';
  var cur = sh.getRange(1, 4, 1, OPE_VITALS_HEADERS.length).getValues()[0].map(function (x) { return String(x || '').trim(); });
  var changes = [];
  cur.forEach(function (h, i) {
    var want = OPE_VITALS_HEADERS[i];
    if (h === want) return;
    var damaged = !h || (h === 'Sys_BP' && i > 0);
    if (!damaged) { changes.push('  column ' + String.fromCharCode(68 + i) + ': "' + h + '" left alone (not a damaged header)'); return; }
    changes.push('  column ' + String.fromCharCode(68 + i) + ': "' + h + '" -> "' + want + '"');
    if (!dryRun) sh.getRange(1, 4 + i).setValue(want);
  });
  if (!changes.length) return 'OP_Encounters vitals headers are correct.';
  return (dryRun ? 'DRY RUN — nothing written. Run repairOPEncounterHeaders(false) to apply.\n' : 'Written:\n') +
         changes.join('\n');
}


// ---------------------------------------------------------------------------
// THE DATES, TIMES AND HEADERS FOUND WRONG IN THE LIVE WORKBOOK (Sep 2026)
//
// The code that wrote each of these is fixed; this repairs what it already
// wrote. It previews by default and changes nothing until asked to apply.
//
//   1. Appointments: columns H and I carry the fee and the booking time and
//      have no headers, so Accounts (which reads by header) saw neither.
//   2. Appointments.Timestamp and Pharmacy_Inventory.Timestamp: rows written
//      as ISO TEXT ("2026-05-17T04:53:02.393Z" — UTC, five and a half hours
//      early) instead of as dates.
//   3. Any cell holding a stringified date — "Sat Dec 30 1899 11:45:00
//      GMT+0521 (India Standard Time)", "Sat Apr 01 2028 00:00:00 GMT+0530"
//      — is rewritten as the date or time it always meant. Signed discharge
//      snapshots and the audit trails are never touched: those are evidence,
//      and are cleaned where they are displayed instead.
//   4. LAB_AUDIT_LOG: rows written under a header they did not match (the
//      action sitting in UserID, and so on) are moved into their columns.
//      Nothing in them is changed or removed.
//   5. IP_Discharge_Drafts: the old eight-column header is brought up to the
//      ten columns saveDischargeDraft writes (AccountsIPChargesLogic.gs).
//
// Run from the editor:            repairSheetData()      preview
//                                 repairSheetData(true)  apply
// or Admin Dashboard → Operations → Backups → "Check sheet data".
// ---------------------------------------------------------------------------

/** EDITOR. Preview (or apply) the sheet data repair. */
function repairSheetData(apply) {
  crescEditorOnly_('repairSheetData');
  var r = cresc_repairSheetData_(!!apply);
  Logger.log(r.report);
  return r.report;
}

/**
 * FRONTEND ENTRY (Operations → Backups). The same repair, for an
 * administrator who is not in the script editor. Take a backup first; the
 * screen offers one.
 */
function crescRepairSheetData(sessionToken, apply) {
  var lock = LockService.getScriptLock();
  try {
    var actor = crescRequire_(sessionToken, 'admin.config');
    lock.waitLock(30000);
    var r = cresc_repairSheetData_(!!apply);
    if (apply) {
      try {
        logAudit_({ username: actor.username, role: actor.role, doctorId: actor.doctorId },
                  'SHEET_DATA_REPAIRED', 'Workbook', '', { changed: r.changed });
      } catch (e) {}
    }
    return { success: true, changed: r.changed, report: r.report, applied: !!apply };
  } catch (err) {
    return { success: false, message: cresc_reason_(err) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/** Sheets whose text is evidence or payload, never rewritten by step 3. */
var CRESC_REPAIR_SKIP = ['DS_Snapshots', 'DS_Working', 'DS_Workflow_Log', 'Audit_Log',
                         'Audit_Event_Ledger', 'LAB_AUDIT_LOG', 'Sessions', 'Consent_Register',
                         'Breach_Register', 'Document_Grants', 'Backup_Log'];

var CRESC_STAMP_RX = /\b(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun) (?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d{1,2} \d{4} \d{2}:\d{2}:\d{2} GMT[+-]\d{4}/;

function cresc_repairSheetData_(apply) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lines = [apply ? 'APPLYING the sheet data repair:' : 'PREVIEW — nothing has been changed. Apply to make these changes:'];
  var changed = 0;
  var note = function (n, what) { if (n) { changed += n; lines.push('  • ' + what); } };

  // ---- 1. Appointments headers --------------------------------------------
  var ap = ss.getSheetByName('Appointments');
  if (ap && ap.getLastColumn() >= 9) {
    var head = ap.getRange(1, 1, 1, 9).getValues()[0];
    var fixes = [];
    if (!String(head[7] || '').trim()) fixes.push([8, 'Fee']);
    if (!String(head[8] || '').trim()) fixes.push([9, 'Timestamp']);
    if (fixes.length && apply) fixes.forEach(function (f) { ap.getRange(1, f[0]).setValue(f[1]).setFontWeight('bold'); });
    note(fixes.length, 'Appointments: header' + (fixes.length > 1 ? 's' : '') + ' ' +
         fixes.map(function (f) { return '"' + f[1] + '" (column ' + String.fromCharCode(64 + f[0]) + ')'; }).join(' and ') +
         (apply ? ' written' : ' to be written') + '.');
  }

  // ---- 2. ISO text in timestamp columns ------------------------------------
  [['Appointments', 9, 'Timestamp'], ['Pharmacy_Inventory', 1, 'Timestamp']].forEach(function (t) {
    var sh = ss.getSheetByName(t[0]);
    if (!sh || sh.getLastRow() < 2 || sh.getLastColumn() < t[1]) return;
    var rng = sh.getRange(2, t[1], sh.getLastRow() - 1, 1);
    var vals = rng.getValues(), n = 0;
    var out = vals.map(function (r) {
      var v = r[0];
      if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(v.trim())) {
        var d = new Date(v.trim());
        if (!isNaN(d.getTime())) { n++; return [d]; }
      }
      return [v];
    });
    if (n && apply) { rng.setValues(out); rng.setNumberFormat('dd-mmm-yyyy hh:mm'); }
    note(n, t[0] + '.' + t[2] + ': ' + n + ' ISO text value(s) ' + (apply ? 'converted' : 'to convert') + ' to real dates.');
  });

  // ---- 3. Stringified dates anywhere ---------------------------------------
  ss.getSheets().forEach(function (sh) {
    var name = sh.getName();
    if (CRESC_REPAIR_SKIP.indexOf(name) !== -1) return;
    var lr = sh.getLastRow(), lc = sh.getLastColumn();
    if (lr < 2 || lc < 1) return;
    var hdr = sh.getRange(1, 1, 1, lc).getValues()[0].map(function (h) { return String(h || '').trim(); });
    var data = sh.getRange(2, 1, lr - 1, lc).getValues();
    var hits = 0, cols = {};
    for (var i = 0; i < data.length; i++) {
      for (var j = 0; j < lc; j++) {
        var v = data[i][j];
        if (typeof v !== 'string' || v.indexOf('GMT') === -1 || !CRESC_STAMP_RX.test(v)) continue;
        if (/^\s*[\{\[]/.test(v)) continue;                      // JSON: somebody's payload
        var isExpiry = /expir/i.test(hdr[j]);
        var cleaned = cresc_cleanStampedTime_(v);
        var d = cresc_parseDate_(cleaned);
        if (apply) {
          var cell = sh.getRange(i + 2, j + 1);
          if (isExpiry && d && !cresc_isSheetEpoch_(d)) {
            cell.setValue(new Date(d.getFullYear(), d.getMonth(), 1)).setNumberFormat('mmm yyyy');
          } else {
            cell.setNumberFormat('@').setValue(cleaned);
          }
        }
        hits++;
        cols[hdr[j] || ('column ' + (j + 1))] = true;
      }
    }
    note(hits, name + ': ' + hits + ' stringified date(s) in ' + Object.keys(cols).join(', ') +
               (apply ? ' rewritten' : ' to rewrite') + ' as readable dates and times.');
  });

  // ---- 4. LAB_AUDIT_LOG column alignment ------------------------------------
  var la = ss.getSheetByName('LAB_AUDIT_LOG');
  if (la && la.getLastRow() > 1) {
    var lh = la.getRange(1, 1, 1, la.getLastColumn()).getValues()[0].map(function (h) { return String(h || '').trim(); });
    var cU = lh.indexOf('UserID'), cN = lh.indexOf('UserName'), cA = lh.indexOf('Action'),
        cT = lh.indexOf('EntityType'), cI = lh.indexOf('EntityID'), cO = lh.indexOf('OldValue'), cV = lh.indexOf('NewValue');
    if (cU === 2 && cA === 4 && cT === 5 && cI === 6 && cO === 7 && cV === 8) {
      var w = la.getLastColumn();
      var rows = la.getRange(2, 1, la.getLastRow() - 1, w).getValues();
      var moved = 0;
      var out2 = rows.map(function (r) {
        // Written as [AuditID, Timestamp, Action, EntityType, EntityID, Old, New, PerformedBy]
        // under the ten-column header: the action is where UserID belongs.
        var looksOld = /^[A-Z][A-Z_]*[A-Z]$/.test(String(r[2] || '')) && /^[A-Z_]+$/.test(String(r[3] || '')) &&
                       !/^[A-Z][A-Z_]*[A-Z]$/.test(String(r[4] || ''));
        if (!looksOld) return r;
        moved++;
        var n = r.slice();
        n[2] = r[7] || '';            // PerformedBy -> UserID
        n[3] = '';                    // UserName was never recorded
        n[4] = r[2]; n[5] = r[3]; n[6] = r[4]; n[7] = r[5]; n[8] = r[6];
        return n;
      });
      if (moved && apply) la.getRange(2, 1, rows.length, w).setValues(out2);
      note(moved, 'LAB_AUDIT_LOG: ' + moved + ' row(s) ' + (apply ? 'moved' : 'to move') +
                  ' into the columns their header names (nothing in them changed).');
    }
  }

  // ---- 5. IP_Discharge_Drafts header ---------------------------------------
  var dr = ss.getSheetByName('IP_Discharge_Drafts');
  if (dr && typeof IPC_DRAFT_HEADERS !== 'undefined') {
    var dh = dr.getRange(1, 1, 1, Math.max(dr.getLastColumn(), 1)).getValues()[0].map(function (h) { return String(h || '').trim(); });
    if (dh.slice(0, IPC_DRAFT_HEADERS.length).join('|') !== IPC_DRAFT_HEADERS.join('|')) {
      if (apply && typeof ipc_healDraftHeaders_ === 'function') ipc_healDraftHeaders_(dr);
      note(1, 'IP_Discharge_Drafts: header ' + (apply ? 'brought' : 'to bring') +
              ' up to the ten columns the discharge desk writes, so saved ward charges load back.');
    }
  }

  if (apply) SpreadsheetApp.flush();
  if (!changed) lines.push('  Nothing to repair — every check is clean.');
  lines.push('');
  lines.push(apply ? 'Done: ' + changed + ' change(s).' : changed + ' change(s) would be made.');
  return { changed: changed, report: lines.join('\n') };
}
