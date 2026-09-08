// =========================================================================
// IP_Admissions_Logic.gs  —  CRESCENTIA HEALTHTECH / CresRx
// Hardened admissions ledger. All writes are lock-guarded and return
// the standard { success, message } contract.
//
// IP_Admissions schema (13 cols, FROZEN — do not reorder):
// [0]IP Number [1]Patient ID [2]Patient Name [3]Age/Sex [4]DOA [5]TOA
// [6]Type [7]Ward [8]Bed [9]Consultant [10]Diagnosis [11]Status [12]DOD
// =========================================================================

var IPA_CFG = {
  SHEET: 'IP_Admissions',
  BEDS: 'Master_Beds',
  PATIENTS: 'Patients',
  HEADERS: ['IP Number', 'Patient ID', 'Patient Name', 'Age/Sex', 'DOA', 'TOA',
            'Type', 'Ward', 'Bed', 'Consultant', 'Diagnosis', 'Status', 'DOD'],
  LOCK_MS: 10000
};

// Column indices, by name. Never hardcode a number outside this object.
var IPA_COL = {
  IP: 0, PATIENT_ID: 1, NAME: 2, AGE_SEX: 3, DOA: 4, TOA: 5,
  TYPE: 6, WARD: 7, BED: 8, CONSULTANT: 9, DIAGNOSIS: 10, STATUS: 11, DOD: 12
};

// ---- helpers --------------------------------------------------------------

function ipa_ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function ipa_str_(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }

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
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), pattern);
  }
  return ipa_str_(v);
}

// ACTIVE and ADMITTED are treated as the same live state.
function ipa_isLive_(status) {
  var s = ipa_str_(status).toUpperCase();
  return s === 'ACTIVE' || s === 'ADMITTED';
}

/**
 * Next sequential IP number for the current month: IP2609-0001.
 * MUST be called from inside an acquired script lock.
 */
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

  var next = maxSeq + 1;
  return prefix + ('0000' + next).slice(-4);
}

// ---- READ: ledger ---------------------------------------------------------

/**
 * Returns the full admissions ledger, newest first.
 * Contract: { success, message, data: [...] } — data is ALWAYS an array.
 */
function getIPLedgerData() {
  try {
    var ss = ipa_ss_();
    var sheet = ss.getSheetByName(IPA_CFG.SHEET);

    if (!sheet) {
      return { success: true, message: 'Ledger not yet created.', data: [] };
    }

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

      ledger.push({
        rowIndex:    i + 2,
        ipNumber:    ipa_str_(r[IPA_COL.IP]),
        patientId:   ipa_str_(r[IPA_COL.PATIENT_ID]),
        patientName: ipa_str_(r[IPA_COL.NAME]),
        ageSex:      ipa_str_(r[IPA_COL.AGE_SEX]),
        doa:         ipa_fmt_(r[IPA_COL.DOA], 'dd MMM yyyy'),
        toa:         ipa_fmt_(r[IPA_COL.TOA], 'hh:mm a'),
        type:        ipa_str_(r[IPA_COL.TYPE]),
        ward:        ipa_str_(r[IPA_COL.WARD]),
        bed:         ipa_str_(r[IPA_COL.BED]),
        consultant:  ipa_str_(r[IPA_COL.CONSULTANT]),
        diagnosis:   ipa_str_(r[IPA_COL.DIAGNOSIS]),
        status:      ipa_str_(r[IPA_COL.STATUS]) || 'UNKNOWN',
        dod:         lastCol > IPA_COL.DOD ? ipa_fmt_(r[IPA_COL.DOD], 'dd MMM yyyy') : ''
      });
    }

    ledger.reverse();
    return { success: true, message: ledger.length + ' record(s).', data: ledger };

  } catch (e) {
    return { success: false, message: 'Ledger read failed: ' + e.message, data: [] };
  }
}

// ---- READ: patient lookup for the admit modal -----------------------------

function fetchPatientForAdmit(patientId) {
  try {
    var sheet = ipa_ss_().getSheetByName(IPA_CFG.PATIENTS);
    if (!sheet) return { success: false, message: 'Patients sheet not found.' };

    var target = ipa_str_(patientId).toUpperCase();
    if (!target) return { success: false, message: 'Enter a Patient ID.' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (ipa_str_(data[i][0]).toUpperCase() === target) {
        return {
          success: true,
          message: 'Patient found.',
          data: {
            patientId: ipa_str_(data[i][0]),
            name: ipa_str_(data[i][2]),
            age:  ipa_str_(data[i][3]),
            sex:  ipa_str_(data[i][4])
          }
        };
      }
    }
    return { success: false, message: 'No patient with ID ' + patientId + '.' };

  } catch (e) {
    return { success: false, message: 'Lookup failed: ' + e.message };
  }
}

// ---- READ: available beds -------------------------------------------------

function getAvailableBedsByWard(ward) {
  try {
    var sheet = ipa_ss_().getSheetByName(IPA_CFG.BEDS);
    if (!sheet) return { success: true, message: 'No bed master.', data: [] };

    var target = ipa_str_(ward).toUpperCase();
    var data = sheet.getDataRange().getValues();
    var beds = [];

    for (var i = 1; i < data.length; i++) {
      var rowWard = ipa_str_(data[i][1]).toUpperCase();
      var rowStatus = ipa_str_(data[i][2]).toUpperCase();
      if (rowWard === target && rowStatus === 'AVAILABLE') {
        beds.push(ipa_str_(data[i][0]));
      }
    }
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

    var patientId = ipa_str_(payload && payload.patientId);
    var bed       = ipa_str_(payload && payload.bed);
    if (!patientId) return { success: false, message: 'Patient ID is required.' };
    if (!bed)       return { success: false, message: 'A bed must be selected.' };

    var sheet = ipa_sheet_();

    // Reject a second live admission for the same patient.
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var existing = sheet.getRange(2, 1, lastRow - 1, IPA_CFG.HEADERS.length).getValues();
      for (var i = 0; i < existing.length; i++) {
        if (ipa_str_(existing[i][IPA_COL.PATIENT_ID]).toUpperCase() === patientId.toUpperCase() &&
            ipa_isLive_(existing[i][IPA_COL.STATUS])) {
          return { success: false,
                   message: 'Patient ' + patientId + ' already has an active admission (' +
                            ipa_str_(existing[i][IPA_COL.IP]) + ').' };
        }
      }
    }

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
      ipa_str_(payload.ward),
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

    var ip = ipa_str_(ipNumber);
    var newBed = ipa_str_(newBedId);
    if (!ip)     return { success: false, message: 'IP Number is required.' };
    if (!newBed) return { success: false, message: 'A destination bed is required.' };

    var sheet = ipa_sheet_();
    var data = sheet.getDataRange().getValues();
    var patientId = '', patientName = '', doa = '';
    var found = false;

    for (var i = 1; i < data.length; i++) {
      if (ipa_str_(data[i][IPA_COL.IP]) !== ip) continue;
      if (!ipa_isLive_(data[i][IPA_COL.STATUS])) {
        return { success: false, message: 'Admission ' + ip + ' is not active.' };
      }
      sheet.getRange(i + 1, IPA_COL.WARD + 1).setValue(ipa_str_(newWard));
      sheet.getRange(i + 1, IPA_COL.BED + 1).setValue(String(newBed));
      patientId   = ipa_str_(data[i][IPA_COL.PATIENT_ID]);
      patientName = ipa_str_(data[i][IPA_COL.NAME]);
      doa         = data[i][IPA_COL.DOA];
      found = true;
      break;
    }
    if (!found) return { success: false, message: 'Admission ' + ip + ' not found.' };

    ipa_releaseBed_(ipa_str_(oldBedId));
    ipa_occupyBed_(newBed, patientId, patientName, doa, ip);

    SpreadsheetApp.flush();
    return { success: true, message: 'Transferred to ' + newWard + ' / ' + newBed + '.' };

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

    for (var i = 1; i < data.length; i++) {
      if (ipa_str_(data[i][IPA_COL.IP]) !== ip) continue;
      if (!ipa_isLive_(data[i][IPA_COL.STATUS])) {
        return { success: false, message: 'Admission ' + ip + ' is already discharged.' };
      }
      sheet.getRange(i + 1, IPA_COL.STATUS + 1).setValue('DISCHARGED');
      sheet.getRange(i + 1, IPA_COL.DOD + 1).setValue(new Date());
      found = true;
      break;
    }
    if (!found) return { success: false, message: 'Admission ' + ip + ' not found.' };

    ipa_releaseBed_(ipa_str_(bedId));

    SpreadsheetApp.flush();
    return { success: true, message: 'Discharged ' + ip + '.' };

  } catch (e) {
    return { success: false, message: 'Discharge failed: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ---- bed state (private; callers already hold the lock) -------------------

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