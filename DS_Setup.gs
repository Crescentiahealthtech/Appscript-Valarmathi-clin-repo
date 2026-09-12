// ============================================================================
// DS_Setup.gs  —  Crescentia HealthTech / CresRx
// IP Discharge Summary Engine · Phase 2 · schema installer
// ----------------------------------------------------------------------------
// NAMING — READ THIS BEFORE ADDING A FUNCTION TO THIS MODULE
//   Apps Script is one flat global namespace and a duplicate function name is
//   silently overwritten, last file in the project order wins. The prefix
//   `ds_` is ALREADY TAKEN by two unrelated modules:
//       Doctor_Session_Store.gs : ds_sessionSheet_, ds_touchSession_,
//                                 ds_slideExpiry_, DS_SESSION_HOURS,
//                                 DS_CACHE_SECONDS
//       Doctor_Schedule_Engine.gs: ds_scheduleSheet_, ds_exceptionSheet_,
//                                 ds_generateSlots_, ds_bookingsByTime_,
//                                 ds_parseDateKey_, DS_WEEKDAYS
//   So in this module:
//       ds_*    public API only, and only names verified against the above
//       dsx_*   every private helper
//       DSX_*   every constant
//       DS_*    sheet names (no collision risk — sheets are not identifiers)
//   A helper named ds_sheet_ or ds_headerMap_ here would have taken out the
//   doctor scheduler with no error message at all.
//
// WHAT THIS FILE DOES
//   setupDischargeSummaryModule() creates the five DS_* sheets. It is
//   idempotent: an existing sheet has its headers verified and only MISSING
//   columns appended at the END. Nothing is ever reordered, renamed or
//   deleted, and no data row is ever written — including seed phrases.
//   No column is added to IP_Admissions or to any existing clinical sheet.
// ============================================================================

var DSX_SHEETS = {
  SUMMARIES: 'DS_Summaries',
  WORKING:   'DS_Working',
  SNAPSHOTS: 'DS_Snapshots',
  LOG:       'DS_Workflow_Log',
  PHRASES:   'DS_Phrase_Library'
};

/** One row per admission. The queue reads ONLY this sheet, so it holds no payload. */
var DSX_H_SUMMARIES = [
  'Summary_ID', 'Tenant_ID', 'IP_Number', 'Patient_ID', 'Discharge_Type', 'Status',
  'Row_Version', 'Current_Snapshot_No', 'Last_Signed_Snapshot_No',
  'Initiated_At', 'Initiated_By', 'Planned_Discharge_At', 'Clinical_Discharge_At',
  'Prepared_By', 'Submitted_At', 'Submitted_By', 'Returned_Count',
  'Signed_At', 'Signed_By', 'Signer_Reg_No', 'Signed_Hash',
  'Pdf_Status', 'Pdf_File_ID', 'Print_Count', 'Last_Printed_At',
  'Cancel_Reason', 'Updated_At', 'Updated_By'
];

/** One mutable row per summary — the live draft. */
var DSX_H_WORKING = [
  'Summary_ID', 'Tenant_ID', 'Base_Snapshot_No', 'Payload_Chars',
  'Payload_1', 'Payload_2', 'Payload_3', 'Payload_4',
  'Payload_5', 'Payload_6', 'Payload_7', 'Payload_8',
  'Updated_At', 'Updated_By'
];

/** Append-only frozen copies. Never updated in place. */
var DSX_H_SNAPSHOTS = [
  'Snapshot_Key', 'Tenant_ID', 'Summary_ID', 'Snapshot_No', 'Snapshot_Type',
  'Content_Hash', 'Prev_Signed_Hash', 'Payload_Chars',
  'Payload_1', 'Payload_2', 'Payload_3', 'Payload_4',
  'Payload_5', 'Payload_6', 'Payload_7', 'Payload_8',
  'Created_At', 'Created_By', 'Verify_Token_Hash'
];

/** Append-only workflow history. */
var DSX_H_LOG = [
  'Event_ID', 'Tenant_ID', 'Summary_ID', 'Timestamp',
  'Actor_Username', 'Actor_Role', 'Action', 'From_Status', 'To_Status',
  'Snapshot_No', 'Content_Hash', 'Comment', 'Meta_JSON'
];

/**
 * Clinic-maintained bilingual advice phrases. STARTS EMPTY and stays empty
 * until a human at the clinic writes a line. Tamil clinical text is authored
 * and reviewed by the clinic — this module never machine-translates and never
 * seeds a phrase.
 */
var DSX_H_PHRASES = [
  'Phrase_ID', 'Tenant_ID', 'Category', 'Specialty',
  'Text_EN', 'Text_TA', 'Active', 'Reviewed_By', 'Created_At', 'Created_By'
];

var DSX_PHRASE_CATEGORIES =
  ['DIET', 'ACTIVITY', 'WOUND_CARE', 'RED_FLAG', 'FOLLOW_UP', 'GENERAL'];

var DSX_SHEET_SPEC = [
  { name: DSX_SHEETS.SUMMARIES, headers: DSX_H_SUMMARIES },
  { name: DSX_SHEETS.WORKING,   headers: DSX_H_WORKING },
  { name: DSX_SHEETS.SNAPSHOTS, headers: DSX_H_SNAPSHOTS },
  { name: DSX_SHEETS.LOG,       headers: DSX_H_LOG },
  { name: DSX_SHEETS.PHRASES,   headers: DSX_H_PHRASES }
];

/**
 * Admin-run once per deployment, and safe to run again after every upgrade.
 *
 * @return {{success:boolean, message:string,
 *           data:{created:Array, extended:Array, verified:Array, proposed:Array}}}
 */
function setupDischargeSummaryModule() {
  var lock = LockService.getScriptLock();
  var created = [], extended = [], verified = [];

  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    DSX_SHEET_SPEC.forEach(function (spec) {
      var sh = ss.getSheetByName(spec.name);

      if (!sh) {
        sh = ss.insertSheet(spec.name);
        sh.getRange(1, 1, 1, spec.headers.length).setValues([spec.headers]);
        sh.setFrozenRows(1);
        sh.getRange(1, 1, 1, spec.headers.length)
          .setFontWeight('bold').setBackground('#e8eaf6');
        created.push(spec.name);
        return;
      }

      // Existing sheet: verify, and append ONLY what is missing, at the end.
      var lastCol = Math.max(1, sh.getLastColumn());
      var present = sh.getRange(1, 1, 1, lastCol).getValues()[0]
                      .map(function (h) { return String(h || '').trim(); });

      var missing = spec.headers.filter(function (h) {
        return present.indexOf(h) === -1;
      });

      if (missing.length) {
        sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
        sh.getRange(1, lastCol + 1, 1, missing.length)
          .setFontWeight('bold').setBackground('#e8eaf6');
        extended.push(spec.name + ' (+' + missing.join(', ') + ')');
      } else {
        verified.push(spec.name);
      }
      sh.setFrozenRows(1);
    });

    SpreadsheetApp.flush();

    // The Doctors master is NOT modified here. Gaps are reported for a human
    // decision — see docs/discharge/DISCOVERY.md §13.
    var proposed = dsx_inspectDoctorsMaster_();

    var parts = [];
    if (created.length)  parts.push('created ' + created.length);
    if (extended.length) parts.push('extended ' + extended.length);
    if (verified.length) parts.push('verified ' + verified.length);

    return {
      success: true,
      message: 'Discharge Summary schema ready — ' + (parts.join(', ') || 'nothing to do') + '.',
      data: { created: created, extended: extended, verified: verified, proposed: proposed }
    };

  } catch (e) {
    return {
      success: false,
      message: 'Discharge Summary setup failed: ' + e.message,
      data: { created: created, extended: extended, verified: verified, proposed: [] }
    };
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/**
 * Read-only. Reports what the Doctors master lacks for signing, WITHOUT
 * touching it. Discovery concluded that no column needs to be added:
 * Signature_Line already carries the qualification. What is missing is DATA.
 */
function dsx_inspectDoctorsMaster_() {
  var out = [];
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Doctors');
    if (!sh) return ['Doctors sheet not found — run setupDoctorsSheet() first.'];

    var data = sh.getDataRange().getDisplayValues();
    var head = data[0].map(function (h) { return String(h || '').trim(); });
    var iId = head.indexOf('Doctor_ID'),
        iNm = head.indexOf('Display_Name'),
        iRg = head.indexOf('Reg_No'),
        iSg = head.indexOf('Signature_Line'),
        iSt = head.indexOf('Status');

    if (iRg === -1) out.push('Proposed, not applied: Doctors.Reg_No column is absent.');
    if (iSg === -1) out.push('Proposed, not applied: Doctors.Signature_Line column is absent.');
    if (iRg === -1 || iSg === -1) return out;

    for (var i = 1; i < data.length; i++) {
      var id = String(data[i][iId] || '').trim();
      if (!id) continue;
      if (iSt > -1 && String(data[i][iSt] || '').trim().toUpperCase() !== 'ACTIVE') continue;
      var name = String(data[i][iNm] || '').trim() || id;
      if (!String(data[i][iRg] || '').trim()) {
        out.push('Data gap: ' + name + ' (' + id + ') has no Reg_No — cannot sign a discharge summary.');
      }
      if (!String(data[i][iSg] || '').trim()) {
        out.push('Data gap: ' + name + ' (' + id + ') has no Signature_Line — nothing to print under the signature.');
      }
    }
  } catch (e) {
    out.push('Could not inspect the Doctors master: ' + e.message);
  }
  return out;
}

/**
 * READ-ONLY self test. Writes nothing, anywhere. Checks that every sheet is
 * present with every header, and round-trips a generated NON-CLINICAL string
 * through the payload packer.
 */
function ds_selfTest() {
  var lines = [], ok = true;

  DSX_SHEET_SPEC.forEach(function (spec) {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(spec.name);
    if (!sh) { ok = false; lines.push('FAIL  ' + spec.name + ' — sheet missing'); return; }

    var lastCol = Math.max(1, sh.getLastColumn());
    var present = sh.getRange(1, 1, 1, lastCol).getValues()[0]
                    .map(function (h) { return String(h || '').trim(); });
    var missing = spec.headers.filter(function (h) { return present.indexOf(h) === -1; });

    if (missing.length) { ok = false; lines.push('FAIL  ' + spec.name + ' — missing: ' + missing.join(', ')); }
    else lines.push('pass  ' + spec.name + ' (' + spec.headers.length + ' headers)');
  });

  // Pack/unpack round trip on synthetic text. No patient data is involved.
  try {
    var filler = '';
    while (filler.length < 120000) {
      filler += 'LOREM-' + filler.length + '-IPSUM-0123456789-abcdefghijklmnopqrstuvwxyz;';
    }
    var obj = { schemaVersion: 1, probe: filler, nested: { b: 2, a: 1 } };
    var json = JSON.stringify(obj);
    var chunks = dsx_packPayload_(json);
    var rowObj = { Payload_Chars: json.length };
    for (var i = 0; i < 8; i++) rowObj['Payload_' + (i + 1)] = chunks[i];
    var back = dsx_unpackPayload_(rowObj);

    if (JSON.stringify(back) !== json) { ok = false; lines.push('FAIL  payload round trip — content differs'); }
    else lines.push('pass  payload round trip (' + json.length + ' chars over ' +
                    chunks.filter(String).length + ' chunk(s))');

    // Canonical JSON must be order-independent.
    var c1 = dsx_canonicalJson_({ b: 1, a: { d: 4, c: 3 } });
    var c2 = dsx_canonicalJson_({ a: { c: 3, d: 4 }, b: 1 });
    if (c1 !== c2) { ok = false; lines.push('FAIL  canonical JSON is not order-independent'); }
    else lines.push('pass  canonical JSON is deterministic');

    var h = dsx_sha256Hex_(c1);
    if (!/^[0-9a-f]{64}$/.test(h)) { ok = false; lines.push('FAIL  sha256 hex is malformed: ' + h); }
    else lines.push('pass  sha256 hex (' + h.substring(0, 12) + '…)');

  } catch (e) {
    ok = false;
    lines.push('FAIL  payload helpers threw: ' + e.message);
  }

  // Oversize payloads must be refused loudly, not silently truncated.
  try {
    var huge = new Array(8 * 45000 + 50).join('x');
    dsx_packPayload_(JSON.stringify(huge));
    ok = false;
    lines.push('FAIL  oversize payload was accepted');
  } catch (e2) {
    if (String(e2.message).indexOf('PAYLOAD_TOO_LARGE') > -1) lines.push('pass  oversize payload refused');
    else { ok = false; lines.push('FAIL  oversize payload threw the wrong error: ' + e2.message); }
  }

  var report = (ok ? 'ds_selfTest: PASS' : 'ds_selfTest: FAIL') + '\n' + lines.join('\n');
  Logger.log(report);
  return { success: ok, message: report, data: { lines: lines } };
}