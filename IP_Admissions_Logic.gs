// =========================================================================
// IP_Admissions_Logic.gs  —  CRESCENTIA HEALTHTECH / CresRx
// Full rewrite. Replaces the previous file entirely.
//
// IP_Admissions schema (13 cols, FROZEN):
// [0]IP_Number [1]Patient_ID [2]Patient_Name [3]Age_Sex [4]DOA [5]TOA
// [6]Admission_Type [7]Ward_Bed [8]Bed [9]Consultant [10]Diagnosis
// [11]Status [12]DOD
//
// NOTE on cols 7/8: historic rows wrote "B - B201" into Ward_Bed while
// leaving Bed stale. ipa_resolveWardBed_() normalises this on read and
// treats the combined form as authoritative. New writes keep them clean.
// =========================================================================

var IPA_CFG = {
  SHEET: 'IP_Admissions',
  BEDS: 'Master_Beds',
  PATIENTS: 'Patients',
  HEADERS: ['IP_Number', 'Patient_ID', 'Patient_Name', 'Age_Sex', 'DOA', 'TOA',
            'Admission_Type', 'Ward_Bed', 'Bed', 'Consultant', 'Diagnosis',
            'Status', 'DOD'],
  LOCK_MS: 10000
};

var IPA_COL = {
  IP: 0, PATIENT_ID: 1, NAME: 2, AGE_SEX: 3, DOA: 4, TOA: 5,
  TYPE: 6, WARD: 7, BED: 8, CONSULTANT: 9, DIAGNOSIS: 10, STATUS: 11, DOD: 12
};

// ---- helpers --------------------------------------------------------------

function ipa_ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function ipa_str_(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }

/** Canonical form for any patient ID. Fixes the lmtvs0001 / LMTVS0001 split. */
function ipa_pid_(v) { return ipa_str_(v).toUpperCase(); }

function ipa_sheet_() {
  var ss = ipa_ss_();
  var sh = ss.getSheetByName(IPA_CFG.SHEET);
  if (!sh) {
    sh = ss.insertSheet(IPA_CFG.SHEET);
    sh.appendRow(IPA_CFG.HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function ipa_fmt_(v, pattern) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), pattern);
  }
  return ipa_str_(v);
}

function ipa_isLive_(status) {
  var s = ipa_str_(status).toUpperCase();
  return s === 'ACTIVE' || s === 'ADMITTED';
}

/**
 * Reconcile the historically-divergent Ward_Bed / Bed columns.
 * "B - B201" + stale "A104"  ->  { ward: 'B',  bed: 'B201' }
 * "A"        + "A102"        ->  { ward: 'A',  bed: 'A102' }
 */
function ipa_resolveWardBed_(wardRaw, bedRaw) {
  var w = ipa_str_(wardRaw);
  var b = ipa_str_(bedRaw);

  var sep = w.indexOf(' - ');
  if (sep > -1) {
    return { ward: w.substring(0, sep).trim(), bed: w.substring(sep + 3).trim(), repaired: true };
  }
  return { ward: w, bed: b, repaired: false };
}

/** Next sequential IP number for the current month. Call inside the lock. */
function ipa_nextIpNumber_(sheet) {
  var prefix = 'IP' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMM') + '-';
  var lastRow = sheet.getLastRow();
  var maxSeq = 0;

  if (lastRow >= 2) {
    var existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < existing.length; i++) {
      var val = ipa_str_(existing[i][0]);
      if (val.indexOf(prefix) !== 0) continue;
      var seq = parseInt(val.substring(prefix.length), 10);
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  }
  return prefix + ('0000' + (maxSeq + 1)).slice(-4);
}

// ---- READ: ledger ---------------------------------------------------------

/**
 * Full admissions ledger, newest first.
 * Contract: { success, message, data:[] }. data is ALWAYS an array.
 */
function getIPLedgerData() {
  try {
    var sheet = ipa_ss_().getSheetByName(IPA_CFG.SHEET);
    if (!sheet) return { success: true, message: 'Ledger not yet created.', data: [] };

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) {
      return { success: true, message: 'No admissions recorded.', data: [] };
    }

    var rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var ledger = [];

    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (!r[IPA_COL.IP]) continue;

      var wb = ipa_resolveWardBed_(r[IPA_COL.WARD], r[IPA_COL.BED]);

      ledger.push({
        rowIndex:    i + 2,
        ipNumber:    ipa_str_(r[IPA_COL.IP]),
        patientId:   ipa_pid_(r[IPA_COL.PATIENT_ID]),
        patientName: ipa_str_(r[IPA_COL.NAME]),
        ageSex:      ipa_str_(r[IPA_COL.AGE_SEX]),
        doa:         ipa_fmt_(r[IPA_COL.DOA], 'dd MMM yyyy'),
        toa:         ipa_fmt_(r[IPA_COL.TOA], 'hh:mm a'),
        type:        ipa_str_(r[IPA_COL.TYPE]),
        ward:        wb.ward,
        bed:         wb.bed,
        dataWarning: wb.repaired ? 'Ward/Bed columns disagree on this row.' : '',
        consultant:  ipa_str_(r[IPA_COL.CONSULTANT]),
        diagnosis:   ipa_str_(r[IPA_COL.DIAGNOSIS]),
        status:      ipa_str_(r[IPA_COL.STATUS]).toUpperCase() || 'UNKNOWN',
        dod:         lastCol > IPA_COL.DOD ? ipa_fmt_(r[IPA_COL.DOD], 'dd MMM yyyy') : ''
      });
    }

    ledger.reverse();
    return { success: true, message: ledger.length + ' record(s).', data: ledger };

  } catch (e) {
    return { success: false, message: 'Ledger read failed: ' + e.message, data: [] };
  }
}

// ---- READ: patient lookup -------------------------------------------------

function fetchPatientForAdmit(patientId) {
  try {
    var sheet = ipa_ss_().getSheetByName(IPA_CFG.PATIENTS);
    if (!sheet) return { success: false, message: 'Patients sheet not found.' };

    var target = ipa_pid_(patientId);
    if (!target) return { success: false, message: 'Enter a Patient ID.' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (ipa_pid_(data[i][0]) !== target) continue;

      // Warn if this patient is already occupying a bed.
      var live = ipa_liveAdmissionFor_(target);

      return {
        success: true,
        message: live ? 'Already admitted under ' + live + '.' : 'Patient found.',
        activeAdmission: live || '',
        data: {
          patientId: ipa_pid_(data[i][0]),
          name: ipa_str_(data[i][2]),
          age:  ipa_str_(data[i][3]),
          sex:  ipa_str_(data[i][4])
        }
      };
    }
    return { success: false, message: 'No patient with ID ' + patientId + '.' };

  } catch (e) {
    return { success: false, message: 'Lookup failed: ' + e.message };
  }
}

/** Returns the IP number of a patient's live admission, or '' if none. */
function ipa_liveAdmissionFor_(patientId) {
  var sheet = ipa_ss_().getSheetByName(IPA_CFG.SHEET);
  if (!sheet) return '';
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return '';

  var target = ipa_pid_(patientId);
  var rows = sheet.getRange(2, 1, lastRow - 1, IPA_CFG.HEADERS.length).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (ipa_pid_(rows[i][IPA_COL.PATIENT_ID]) === target &&
        ipa_isLive_(rows[i][IPA_COL.STATUS])) {
      return ipa_str_(rows[i][IPA_COL.IP]);
    }
  }
  return '';
}

// ---- READ: wards and beds -------------------------------------------------

function getWardList() {
  try {
    var sheet = ipa_ss_().getSheetByName(IPA_CFG.BEDS);
    if (!sheet) return { success: true, message: 'No bed master.', data: [] };

    var data = sheet.getDataRange().getValues();
    var seen = {}, wards = [];
    for (var i = 1; i < data.length; i++) {
      var w = ipa_str_(data[i][1]);
      if (!w || seen[w]) continue;
      seen[w] = true;
      wards.push(w);
    }
    wards.sort();
    return { success: true, message: wards.length + ' ward(s).', data: wards };

  } catch (e) {
    return { success: false, message: 'Ward lookup failed: ' + e.message, data: [] };
  }
}

function getAvailableBedsByWard(ward) {
  try {
    var sheet = ipa_ss_().getSheetByName(IPA_CFG.BEDS);
    if (!sheet) return { success: true, message: 'No bed master.', data: [] };

    var target = ipa_str_(ward).toUpperCase();
    if (!target) return { success: true, message: 'Select a ward.', data: [] };

    var data = sheet.getDataRange().getValues();
    var beds = [];
    for (var i = 1; i < data.length; i++) {
      if (ipa_str_(data[i][1]).toUpperCase() !== target) continue;
      if (ipa_str_(data[i][2]).toUpperCase() !== 'AVAILABLE') continue;
      beds.push(ipa_str_(data[i][0]));
    }
    beds.sort();
    return { success: true, message: beds.length + ' bed(s) available.', data: beds };

  } catch (e) {
    return { success: false, message: 'Bed lookup failed: ' + e.message, data: [] };
  }
}

// ---- WRITE: new admission -------------------------------------------------

function saveNewAdmissionLedger(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(IPA_CFG.LOCK_MS);

    var patientId = ipa_pid_(payload && payload.patientId);
    var bed       = ipa_str_(payload && payload.bed);
    var ward      = ipa_str_(payload && payload.ward);

    if (!patientId) return { success: false, message: 'Patient ID is required.' };
    if (!ward)      return { success: false, message: 'A ward must be selected.' };
    if (!bed)       return { success: false, message: 'A bed must be selected.' };

    var existingIp = ipa_liveAdmissionFor_(patientId);
    if (existingIp) {
      return { success: false,
               message: 'Patient ' + patientId + ' already has an active admission (' + existingIp + '). Discharge it first.' };
    }

    if (!ipa_bedIsAvailable_(bed)) {
      return { success: false, message: 'Bed ' + bed + ' is no longer available. Refresh and pick another.' };
    }

    var sheet = ipa_sheet_();
    var newIpNumber = ipa_nextIpNumber_(sheet);

    var doaRaw = ipa_str_(payload.doa);
    var doaVal = doaRaw ? new Date(doaRaw) : new Date();
    if (isNaN(doaVal.getTime())) doaVal = new Date();

    sheet.appendRow([
      String(newIpNumber),
      String(patientId),
      ipa_str_(payload.patientName),
      ipa_str_(payload.ageSex),
      doaVal,
      ipa_str_(payload.toa),
      ipa_str_(payload.type),
      String(ward),          // ward ONLY — never "Ward - Bed"
      String(bed),
      ipa_str_(payload.consultant),
      ipa_str_(payload.diagnosis),
      'ACTIVE',
      ''
    ]);

    ipa_occupyBed_(bed, patientId, ipa_str_(payload.patientName), doaVal, newIpNumber);

    SpreadsheetApp.flush();
    return { success: true, message: 'Admitted. IP Number ' + newIpNumber, ipNumber: newIpNumber };

  } catch (e) {
    return { success: false, message: 'Admission failed: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ---- WRITE: bed transfer --------------------------------------------------

function processBedTransfer(ipNumber, oldBedId, newWard, newBedId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(IPA_CFG.LOCK_MS);

    var ip     = ipa_str_(ipNumber);
    var newBed = ipa_str_(newBedId);
    var ward   = ipa_str_(newWard);

    if (!ip)     return { success: false, message: 'IP Number is required.' };
    if (!ward)   return { success: false, message: 'A destination ward is required.' };
    if (!newBed) return { success: false, message: 'A destination bed is required.' };

    if (!ipa_bedIsAvailable_(newBed)) {
      return { success: false, message: 'Bed ' + newBed + ' is no longer available.' };
    }

    var sheet = ipa_sheet_();
    var data = sheet.getDataRange().getValues();
    var patientId = '', patientName = '', doa = '', found = false;

    for (var i = 1; i < data.length; i++) {
      if (ipa_str_(data[i][IPA_COL.IP]) !== ip) continue;
      if (!ipa_isLive_(data[i][IPA_COL.STATUS])) {
        return { success: false, message: 'Admission ' + ip + ' is not active.' };
      }
      // Clean write: ward and bed to their own columns.
      sheet.getRange(i + 1, IPA_COL.WARD + 1).setValue(String(ward));
      sheet.getRange(i + 1, IPA_COL.BED + 1).setValue(String(newBed));

      patientId   = ipa_pid_(data[i][IPA_COL.PATIENT_ID]);
      patientName = ipa_str_(data[i][IPA_COL.NAME]);
      doa         = data[i][IPA_COL.DOA];
      found = true;
      break;
    }
    if (!found) return { success: false, message: 'Admission ' + ip + ' not found.' };

    ipa_releaseBed_(ipa_str_(oldBedId));
    ipa_occupyBed_(newBed, patientId, patientName, doa, ip);

    SpreadsheetApp.flush();
    return { success: true, message: 'Transferred to ' + ward + ' / ' + newBed + '.' };

  } catch (e) {
    return { success: false, message: 'Transfer failed: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ---- WRITE: discharge -----------------------------------------------------

function processPatientDischarge(ipNumber, bedId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(IPA_CFG.LOCK_MS);

    var ip = ipa_str_(ipNumber);
    if (!ip) return { success: false, message: 'IP Number is required.' };

    var sheet = ipa_sheet_();
    var data = sheet.getDataRange().getValues();
    var found = false;
    var resolvedBed = ipa_str_(bedId);

    for (var i = 1; i < data.length; i++) {
      if (ipa_str_(data[i][IPA_COL.IP]) !== ip) continue;
      if (!ipa_isLive_(data[i][IPA_COL.STATUS])) {
        return { success: false, message: 'Admission ' + ip + ' is already discharged.' };
      }
      var wb = ipa_resolveWardBed_(data[i][IPA_COL.WARD], data[i][IPA_COL.BED]);
      if (wb.bed) resolvedBed = wb.bed;

      sheet.getRange(i + 1, IPA_COL.STATUS + 1).setValue('DISCHARGED');
      sheet.getRange(i + 1, IPA_COL.DOD + 1).setValue(new Date());
      found = true;
      break;
    }
    if (!found) return { success: false, message: 'Admission ' + ip + ' not found.' };

    ipa_releaseBed_(resolvedBed);

    SpreadsheetApp.flush();
    return { success: true, message: 'Discharged ' + ip + '.' };

  } catch (e) {
    return { success: false, message: 'Discharge failed: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ---- bed state (private; caller holds the lock) ---------------------------

function ipa_bedIsAvailable_(bedId) {
  var sheet = ipa_ss_().getSheetByName(IPA_CFG.BEDS);
  if (!sheet) return true; // no bed master configured — don't block admissions
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (ipa_str_(data[i][0]) !== ipa_str_(bedId)) continue;
    return ipa_str_(data[i][2]).toUpperCase() === 'AVAILABLE';
  }
  return true; // bed not in master — allow, but it won't be tracked
}

function ipa_occupyBed_(bedId, patientId, patientName, doa, ipNumber) {
  var sheet = ipa_ss_().getSheetByName(IPA_CFG.BEDS);
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (ipa_str_(data[i][0]) !== ipa_str_(bedId)) continue;
    sheet.getRange(i + 1, 3).setValue('Occupied');
    sheet.getRange(i + 1, 4).setValue(String(patientId));
    sheet.getRange(i + 1, 5).setValue(String(patientName));
    sheet.getRange(i + 1, 6).setValue(doa);
    sheet.getRange(i + 1, 7).setValue(String(ipNumber));
    return;
  }
}

function ipa_releaseBed_(bedId) {
  if (!bedId) return;
  var sheet = ipa_ss_().getSheetByName(IPA_CFG.BEDS);
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (ipa_str_(data[i][0]) !== ipa_str_(bedId)) continue;
    sheet.getRange(i + 1, 3).setValue('Cleaning');
    sheet.getRange(i + 1, 4, 1, 4).clearContent();
    return;
  }
}

// =========================================================================
// ONE-SHOT MIGRATIONS — run manually from the editor. Not called by the UI.
// =========================================================================

/**
 * DRY RUN. Reports what repairWardBedColumns() would change. Writes nothing.
 */
function previewWardBedRepair() {
  var sheet = ipa_ss_().getSheetByName(IPA_CFG.SHEET);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Nothing to repair.'); return; }

  var rows = sheet.getRange(2, 1, lastRow - 1, IPA_CFG.HEADERS.length).getValues();
  var changes = [];
  for (var i = 0; i < rows.length; i++) {
    var wb = ipa_resolveWardBed_(rows[i][IPA_COL.WARD], rows[i][IPA_COL.BED]);
    var pidRaw = ipa_str_(rows[i][IPA_COL.PATIENT_ID]);
    var pidFix = ipa_pid_(pidRaw);
    if (!wb.repaired && pidRaw === pidFix) continue;
    changes.push({
      row: i + 2,
      ip: ipa_str_(rows[i][IPA_COL.IP]),
      wardBefore: ipa_str_(rows[i][IPA_COL.WARD]), wardAfter: wb.ward,
      bedBefore: ipa_str_(rows[i][IPA_COL.BED]),   bedAfter: wb.bed,
      pidBefore: pidRaw, pidAfter: pidFix
    });
  }
  Logger.log(JSON.stringify({ rowsAffected: changes.length, changes: changes }, null, 2));
}

/**
 * DESTRUCTIVE. Splits "Ward - Bed" into separate columns and upper-cases
 * patient IDs. Run previewWardBedRepair() first and read the output.
 */
function repairWardBedColumns() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(IPA_CFG.LOCK_MS);
    var sheet = ipa_sheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, message: 'Nothing to repair.' };

    var rows = sheet.getRange(2, 1, lastRow - 1, IPA_CFG.HEADERS.length).getValues();
    var fixed = 0;
    for (var i = 0; i < rows.length; i++) {
      var wb = ipa_resolveWardBed_(rows[i][IPA_COL.WARD], rows[i][IPA_COL.BED]);
      var pidFix = ipa_pid_(rows[i][IPA_COL.PATIENT_ID]);
      var touched = false;

      if (wb.repaired) {
        sheet.getRange(i + 2, IPA_COL.WARD + 1).setValue(wb.ward);
        sheet.getRange(i + 2, IPA_COL.BED + 1).setValue(wb.bed);
        touched = true;
      }
      if (ipa_str_(rows[i][IPA_COL.PATIENT_ID]) !== pidFix) {
        sheet.getRange(i + 2, IPA_COL.PATIENT_ID + 1).setValue(pidFix);
        touched = true;
      }
      if (touched) fixed++;
    }
    SpreadsheetApp.flush();
    return { success: true, message: fixed + ' row(s) repaired.' };

  } catch (e) {
    return { success: false, message: 'Repair failed: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}