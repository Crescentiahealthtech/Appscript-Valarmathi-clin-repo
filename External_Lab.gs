// ============================================================================
// External_Lab.gs — Crescentia HealthTech
// Lab values from OUTSIDE labs, entered where the clinician is.
// ----------------------------------------------------------------------------
// A patient admitted with last week's HbA1c from another lab, a CT report from
// a scan centre, a culture sent out because the clinic does not run it: the
// numbers are part of the case, and until now the only place to put them was
// free text in a note, where nothing could chart them, flag them, or find
// them again.
//
// Each report is one or more rows on External_Lab_Results sharing a Report_ID
// (one row per parameter). The shared lab viewer (Lab_Patient_View.gs) and
// the IP Notes Investigations panel show them beside the internal lab's
// results, labelled with the outside lab's name, so nobody mistakes a value
// the clinic did not run for one it did.
//
// Rows are never deleted. A report entered in error is VOIDED with a reason
// and stays retrievable.
// ============================================================================

var EXL_SHEET = 'External_Lab_Results';
var EXL_HEADERS = [
  'Row_ID', 'Report_ID', 'Patient_ID', 'IP_Number', 'Lab_Name', 'Report_Date',
  'Test_Name', 'Parameter', 'Value', 'Unit', 'Ref_Range', 'Flag', 'Notes',
  'Entered_By', 'Entered_At', 'Status', 'Void_Reason'
];

/** Who may enter an outside result: anyone who writes clinical records. */
var EXL_WRITE_PERMS = ['ward.write', 'emr.write', 'lab.result'];

function exl_sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(EXL_SHEET);
  if (!sh) {
    sh = ss.insertSheet(EXL_SHEET);
    sh.getRange(1, 1, 1, EXL_HEADERS.length).setValues([EXL_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function exl_str_(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }

function exl_map_(sh) {
  var h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var m = {};
  h.forEach(function (x, i) { m[exl_str_(x)] = i; });
  return m;
}

/**
 * H / L from a reference range the person typed, when the value is a number.
 * Understands "70-110", "70 – 110", "<5.7", "< 200", ">40", "≤ 5", "up to 40".
 * Returns '' when it cannot tell — an unflagged value is better than a wrong
 * flag.
 */
function exl_flagFor_(value, refRange) {
  var v = parseFloat(String(value).replace(/,/g, ''));
  if (isNaN(v)) return '';
  var r = exl_str_(refRange).replace(/,/g, '').replace(/[–—]/g, '-').toLowerCase();
  if (!r) return '';
  var m = /^(-?\d+(?:\.\d+)?)\s*-\s*(-?\d+(?:\.\d+)?)/.exec(r);
  if (m) {
    var lo = parseFloat(m[1]), hi = parseFloat(m[2]);
    return v < lo ? 'L' : (v > hi ? 'H' : 'N');
  }
  m = /^(?:<|≤|<=|up to|upto|below)\s*(-?\d+(?:\.\d+)?)/.exec(r);
  if (m) return v > parseFloat(m[1]) ? 'H' : 'N';
  m = /^(?:>|≥|>=|above|over)\s*(-?\d+(?:\.\d+)?)/.exec(r);
  if (m) return v < parseFloat(m[1]) ? 'L' : 'N';
  return '';
}

function exl_patientExists_(pid) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Patients');
  if (!sh || sh.getLastRow() < 2) return false;
  var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (exl_str_(ids[i][0]).toUpperCase() === pid) return true;
  }
  return false;
}

// ---------------------------------------------------------------------------
// FRONTEND ENTRY POINTS
// ---------------------------------------------------------------------------

/**
 * Records one outside report.
 *
 * @param {{patientId, ipNumber?, labName, reportDate, testName, notes?,
 *          rows:[{parameter, value, unit?, refRange?, flag?}]}} payload
 */
function saveExternalLabResult(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var actor = crescRequire_(sessionToken, EXL_WRITE_PERMS);
    payload = payload || {};

    var pid = exl_str_(payload.patientId).toUpperCase();
    var ip = exl_str_(payload.ipNumber).toUpperCase();
    if (!pid) return { success: false, message: 'Open a patient first.' };
    if (!exl_patientExists_(pid)) return { success: false, message: 'Patient ' + pid + ' was not found.' };
    if (ip && typeof resolveIPRead_ === 'function') {
      var gate = resolveIPRead_(sessionToken, ip);
      if (!gate.ok) return { success: false, message: gate.message };
    }

    var lab = exl_str_(payload.labName);
    if (lab.length < 2) return { success: false, message: 'Name the lab the report came from.' };
    var when = (typeof cresc_parseDate_ === 'function') ? cresc_parseDate_(payload.reportDate) : new Date(payload.reportDate);
    if (!when || isNaN(when.getTime())) return { success: false, message: 'Enter the date on the report.' };
    if (when.getTime() > Date.now() + 24 * 3600 * 1000) {
      return { success: false, message: 'The report date is in the future.' };
    }
    var test = exl_str_(payload.testName);
    if (!test) return { success: false, message: 'Name the test or panel (for example "Lipid profile").' };

    var rows = (payload.rows || []).map(function (r) {
      return {
        parameter: exl_str_(r && r.parameter),
        value:     exl_str_(r && r.value),
        unit:      exl_str_(r && r.unit),
        refRange:  exl_str_(r && r.refRange),
        flag:      exl_str_(r && r.flag).toUpperCase()
      };
    }).filter(function (r) { return r.parameter || r.value; });
    if (!rows.length) return { success: false, message: 'Enter at least one value.' };
    for (var k = 0; k < rows.length; k++) {
      if (!rows[k].parameter) return { success: false, message: 'Row ' + (k + 1) + ' has a value but no parameter name.' };
      if (!rows[k].value) return { success: false, message: rows[k].parameter + ' has no value.' };
      if (['', 'H', 'L', 'N', 'C'].indexOf(rows[k].flag) === -1) rows[k].flag = '';
      if (!rows[k].flag) rows[k].flag = exl_flagFor_(rows[k].value, rows[k].refRange);
    }

    var sh = exl_sheet_();
    var now = new Date();
    var reportId = 'EXL-' + Utilities.formatDate(now, 'Asia/Kolkata', 'yyMMdd-HHmmss') + '-' +
                   Utilities.getUuid().substring(0, 4).toUpperCase();
    var who = actor.displayName || actor.username;
    var out = rows.map(function (r, i) {
      return [reportId + '-' + (i + 1), reportId, pid, ip, lab, when, test,
              r.parameter, r.value, r.unit, r.refRange, r.flag, exl_str_(payload.notes),
              who, now, 'ACTIVE', ''];
    });
    // Text, so "0.80" stays "0.80" and a value like "1-2" is not read as a date.
    sh.getRange(sh.getLastRow() + 1, 1, out.length, EXL_HEADERS.length)
      .setNumberFormat('@').setValues(out.map(function (row) {
        return row.map(function (v) { return v instanceof Date ? Utilities.formatDate(v, 'Asia/Kolkata', 'yyyy-MM-dd HH:mm') : v; });
      }));
    SpreadsheetApp.flush();

    try {
      logAudit_({ username: actor.username, role: actor.role, doctorId: actor.doctorId },
                'EXTERNAL_LAB_ENTERED', 'Patient', pid,
                { reportId: reportId, lab: lab, test: test, rows: rows.length, ipNumber: ip });
    } catch (e) { /* the record stands without its audit row */ }

    var abnormal = rows.filter(function (r) { return r.flag && r.flag !== 'N'; }).length;
    return { success: true, reportId: reportId,
             message: 'Saved ' + rows.length + ' value(s) from ' + lab +
                      (abnormal ? ' — ' + abnormal + ' outside the reference range.' : '.') };
  } catch (err) {
    return { success: false, message: cresc_reason_(err) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/** Voids an outside report entered in error. It stays on the sheet, marked. */
function voidExternalLabResult(reportId, reason, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var actor = crescRequire_(sessionToken, EXL_WRITE_PERMS);
    var id = exl_str_(reportId);
    var why = exl_str_(reason);
    if (!id) return { success: false, message: 'No report was named.' };
    if (why.length < 3) return { success: false, message: 'Say why this report is being removed.' };

    var sh = exl_sheet_();
    if (sh.getLastRow() < 2) return { success: false, message: 'Report ' + id + ' was not found.' };
    var m = exl_map_(sh);
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    var n = 0;
    for (var i = 0; i < data.length; i++) {
      if (exl_str_(data[i][m.Report_ID]) !== id || exl_str_(data[i][m.Status]) !== 'ACTIVE') continue;
      sh.getRange(i + 2, m.Status + 1).setValue('VOID');
      sh.getRange(i + 2, m.Void_Reason + 1).setValue(why + ' — ' + (actor.displayName || actor.username) +
        ', ' + Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm'));
      n++;
    }
    if (!n) return { success: false, message: 'Report ' + id + ' was not found, or is already removed.' };
    try {
      logAudit_({ username: actor.username, role: actor.role }, 'EXTERNAL_LAB_VOIDED', 'Report', id, { reason: why });
    } catch (e) {}
    return { success: true, message: 'Removed. It is kept on the sheet, marked void, with your reason.' };
  } catch (err) {
    return { success: false, message: cresc_reason_(err) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * Parameter names, units and reference ranges to suggest in the entry form:
 * the internal lab catalogue first (so an outside HbA1c is filed under the
 * same name as an inside one), then anything typed on an earlier outside
 * report.
 */
function getExternalLabFormContext(sessionToken) {
  try {
    crescRequire_(sessionToken, EXL_WRITE_PERMS);
    var seen = {}, params = [], labs = {}, tests = {};
    var add = function (name, unit, ref) {
      var key = exl_str_(name).toUpperCase();
      if (!key || seen[key]) return;
      seen[key] = true;
      params.push({ name: exl_str_(name), unit: exl_str_(unit), ref: exl_str_(ref) });
    };
    try {
      var cat = (typeof lab_catalog_ === 'function') ? lab_catalog_() : null;
      if (cat && cat.success) {
        var refOf = function (t) {
          var lo = t.maleRefLow, hi = t.maleRefHigh;
          return (lo !== null && lo !== undefined && lo !== '' && hi !== null && hi !== undefined && hi !== '')
            ? (lo + ' - ' + hi) : '';
        };
        (cat.panels || []).forEach(function (p) {
          tests[p.testName] = true;
          (p.parameters || []).forEach(function (t) { add(t.testName, t.unit, refOf(t)); });
        });
        (cat.individuals || []).forEach(function (t) { tests[t.testName] = true; add(t.testName, t.unit, refOf(t)); });
      }
    } catch (e) { /* no catalogue: earlier entries still help */ }

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXL_SHEET);
    if (sh && sh.getLastRow() > 1) {
      var m = exl_map_(sh);
      var data = sh.getRange(2, 1, Math.min(sh.getLastRow() - 1, 3000), sh.getLastColumn()).getValues();
      data.forEach(function (r) {
        add(r[m.Parameter], r[m.Unit], r[m.Ref_Range]);
        if (exl_str_(r[m.Lab_Name])) labs[exl_str_(r[m.Lab_Name])] = true;
        if (exl_str_(r[m.Test_Name])) tests[exl_str_(r[m.Test_Name])] = true;
      });
    }
    return { success: true, parameters: params, labs: Object.keys(labs).sort(),
             tests: Object.keys(tests).filter(Boolean).sort() };
  } catch (err) {
    return { success: false, message: cresc_reason_(err), parameters: [], labs: [], tests: [] };
  }
}

// ---------------------------------------------------------------------------
// READERS FOR THE OTHER VIEWERS (no authorisation of their own)
// ---------------------------------------------------------------------------

/**
 * This patient's active outside reports, shaped like lpv_resultsFor_'s orders
 * so one renderer draws both. Newest first.
 *
 * @param {string} pid        upper-case patient id
 * @param {string} [ipNumber] only reports filed against this admission, or
 *                            with no admission at all
 */
function exl_ordersFor_(pid, ipNumber) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXL_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  var m = exl_map_(sh);
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  var by = {}, order = [];
  data.forEach(function (r) {
    if (exl_str_(r[m.Patient_ID]).toUpperCase() !== pid) return;
    if (exl_str_(r[m.Status]) !== 'ACTIVE') return;
    var rowIp = exl_str_(r[m.IP_Number]).toUpperCase();
    if (ipNumber && rowIp && rowIp !== String(ipNumber).toUpperCase()) return;
    var id = exl_str_(r[m.Report_ID]);
    if (!by[id]) {
      var d = (typeof cresc_parseDate_ === 'function') ? cresc_parseDate_(r[m.Report_Date]) : new Date(r[m.Report_Date]);
      by[id] = {
        orderId: id,
        external: true,
        labName: exl_str_(r[m.Lab_Name]),
        date: d ? Utilities.formatDate(d, 'Asia/Kolkata', 'dd-MMM-yyyy') : exl_str_(r[m.Report_Date]),
        ms: d ? d.getTime() : 0,
        testNames: exl_str_(r[m.Test_Name]),
        doctor: '',
        source: 'EXTERNAL',
        status: 'EXTERNAL',
        verifiedBy: '',
        enteredBy: exl_str_(r[m.Entered_By]),
        notes: exl_str_(r[m.Notes]),
        abnormal: 0,
        results: []
      };
      order.push(id);
    }
    var flag = exl_str_(r[m.Flag]).toUpperCase();
    if (flag && flag !== 'N') by[id].abnormal++;
    by[id].results.push({
      parameterName: exl_str_(r[m.Parameter]),
      value: exl_str_(r[m.Value]),
      unit: exl_str_(r[m.Unit]),
      flag: flag,
      critical: flag === 'C',
      refRangeText: exl_str_(r[m.Ref_Range])
    });
  });
  return order.map(function (id) { return by[id]; })
              .sort(function (a, b) { return b.ms - a.ms; });
}
