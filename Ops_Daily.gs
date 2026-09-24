// ============================================================================
// Ops_Daily.gs — Crescentia HealthTech / CresRx
// The owner's end-of-day summary, the nightly backup, and the jobs that run
// them.
// ----------------------------------------------------------------------------
//   ownerDailySummary   9 pm   one email: patients seen, money in, money owed,
//                              beds, stock that needs ordering, tomorrow's list,
//                              and anything odd at sign-in
//   opsNightlyBackup    1 am   a full copy of the spreadsheet into a private
//                              Drive folder, checked against the original,
//                              with old copies pruned
//   remindersDaily      6 pm   Patient_Reminders.gs
//
// opsInstallTriggers() — once, from the Apps Script editor — installs all
// three. Admin Dashboard -> Operations shows whether they are installed, and
// lets an administrator see the summary, run a backup, and read the restore
// checklist without the editor.
//
// Script properties this reads (all optional):
//   OWNER_EMAIL        where the summary goes. Default: the Email Address of
//                      every Users row marked Super_Admin.
//   BACKUP_FOLDER_ID   the Drive folder for backups. Created (private) on the
//                      first run if absent, and remembered.
//   BACKUP_KEEP_DAYS   daily copies kept (default 30). The first copy of
//                      each month is kept for a year regardless.
// ============================================================================

var OPS_TRIGGERS = [
  { fn: 'ownerDailySummary', hour: 21, what: 'owner end-of-day summary email' },
  { fn: 'opsNightlyBackup',  hour: 1,  what: 'full spreadsheet backup to a private Drive folder' },
  { fn: 'remindersDaily',    hour: 18, what: 'tomorrow\'s patient reminders (when a provider is set up)' }
];

var OPS_BACKUP = {
  LOG: 'Backup_Log',
  HEADERS: ['Backup_At', 'File_ID', 'File_Name', 'Status', 'Verified', 'Detail', 'Pruned', 'By'],
  PREFIX: 'CresRx backup ',
  FOLDER_NAME: 'CresRx Backups (restricted)',
  // Compared between the original and the copy after every backup.
  CHECK_SHEETS: ['Patients', 'Users', 'Appointments', 'OP_Encounters', 'IP_Admissions',
                 'IP_CaseSheets_DB', 'LAB_ORDERS', 'LAB_RESULTS', 'LAB_BILLING',
                 'Pharmacy_Inventory', 'Pharmacy_Invoices', 'Hospital_Invoices',
                 'Audit_Log']
};

function ops_tz_() { return Session.getScriptTimeZone() || 'Asia/Kolkata'; }
function ops_prop_(k) {
  try { return PropertiesService.getScriptProperties().getProperty(k) || ''; } catch (e) { return ''; }
}
function ops_money_(n) {
  return '₹' + (Math.round((Number(n) || 0) * 100) / 100).toLocaleString('en-IN', { maximumFractionDigits: 2 });
}

// ---------------------------------------------------------------------------
// END-OF-DAY SUMMARY
// ---------------------------------------------------------------------------

/** Who receives the summary. Never anybody outside the owner accounts. */
function ops_ownerEmails_() {
  var set = ops_prop_('OWNER_EMAIL');
  if (set) return set.split(/[,;\s]+/).filter(function (x) { return /@/.test(x); });
  var out = [];
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    var v = sh.getDataRange().getValues(), h = v[0].map(function (x) { return String(x).trim(); });
    var sa = h.indexOf(typeof CRESC_SUPERADMIN_HEADER !== 'undefined' ? CRESC_SUPERADMIN_HEADER : 'Super_Admin');
    var em = h.indexOf('Email Address');
    if (sa !== -1 && em !== -1) {
      for (var i = 1; i < v.length; i++) {
        if (/^(YES|TRUE|1|Y)$/i.test(String(v[i][sa]).trim()) && /@/.test(String(v[i][em]))) out.push(String(v[i][em]).trim());
      }
    }
  } catch (e) {}
  return out;
}

/** Sign-in trouble today: failures, lock-outs, MFA failures. */
function ops_signInTrouble_(todayKey) {
  var out = { failed: 0, locked: 0, mfaFailed: 0, users: {} };
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Audit_Log');
    if (!sh || sh.getLastRow() < 2) return out;
    var n = Math.min(3000, sh.getLastRow() - 1);
    var cols = sh.getLastColumn();
    var h = sh.getRange(1, 1, 1, cols).getValues()[0].map(String);
    var ti = h.indexOf('Timestamp'), ei = h.indexOf('Event'), ui = h.indexOf('Actor_Username');
    var rows = sh.getRange(sh.getLastRow() - n + 1, 1, n, cols).getValues();
    for (var i = rows.length - 1; i >= 0; i--) {
      var t = rows[i][ti];
      if (!(t instanceof Date)) continue;
      var k = Utilities.formatDate(t, ops_tz_(), 'yyyy-MM-dd');
      if (k < todayKey) break;
      if (k !== todayKey) continue;
      var ev = String(rows[i][ei]);
      if (ev === 'LOGIN_FAILED' || ev === 'LOGIN_UNKNOWN_USER') out.failed++;
      else if (ev === 'LOGIN_LOCKED') { out.locked++; out.users[String(rows[i][ui])] = true; }
      else if (ev === 'LOGIN_MFA_FAILED') out.mfaFailed++;
    }
  } catch (e) {}
  out.lockedUsers = Object.keys(out.users);
  delete out.users;
  return out;
}

/** Admissions and discharges today, by header. */
function ops_ipToday_(ss, todayKey, memo) {
  var out = { admitted: 0, discharged: 0 };
  var ip = _dashSheet_(ss, 'IP_Admissions', memo);
  if (ip.length < 2) return out;
  var a = _dashHeaderIndex_(ip[0], ['doa', 'admission_date', 'admitted_at', 'admission date']);
  var d = _dashHeaderIndex_(ip[0], ['discharge_date', 'dod', 'discharged_at', 'discharge date']);
  for (var i = 1; i < ip.length; i++) {
    var ad = a > -1 ? _dashToDate_(ip[i][a]) : null;
    if (ad && Utilities.formatDate(ad, ops_tz_(), 'yyyy-MM-dd') === todayKey) out.admitted++;
    var dd = d > -1 ? _dashToDate_(ip[i][d]) : null;
    if (dd && Utilities.formatDate(dd, ops_tz_(), 'yyyy-MM-dd') === todayKey) out.discharged++;
  }
  return out;
}

/** New registrations today. */
function ops_newPatients_(ss, todayKey, memo) {
  var p = _dashSheet_(ss, 'Patients', memo);
  if (p.length < 2) return 0;
  var r = _dashHeaderIndex_(p[0], ['registration_date', 'registered_at', 'registration date']);
  if (r < 0) return 0;
  var n = 0;
  for (var i = 1; i < p.length; i++) {
    var d = _dashToDate_(p[i][r]);
    if (d && Utilities.formatDate(d, ops_tz_(), 'yyyy-MM-dd') === todayKey) n++;
  }
  return n;
}

/** Everything the summary says, as data. No permission check. */
function ops_buildSummary_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ops_tz_();
  var now = new Date();
  var todayKey = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  var memo = {};
  var s = {
    date: Utilities.formatDate(now, tz, 'EEEE, dd MMM yyyy'),
    generatedAt: Utilities.formatDate(now, tz, 'dd-MMM-yyyy HH:mm'),
    clinic: (typeof cresc_clinic_ === 'function' ? cresc_clinic_().name : '') || 'CresRx'
  };
  try { s.ops = _dashOps_(ss, tz, todayKey, memo); } catch (e) { s.ops = null; }
  try { s.revenue = _dashRevenue_(ss, tz, todayKey, memo); } catch (e) { s.revenue = null; }
  try { s.census = _dashCensus_(ss, memo); } catch (e) { s.census = null; }
  try { s.ip = ops_ipToday_(ss, todayKey, memo); } catch (e) { s.ip = null; }
  try { s.newPatients = ops_newPatients_(ss, todayKey, memo); } catch (e) { s.newPatients = null; }
  try {
    var st = stk_analyse_();
    s.stock = { counts: st.counts, values: st.values,
                low: st.low.slice(0, 8), expiring: st.expiring.filter(function (x) { return x.urgent; }).slice(0, 8) };
  } catch (e) { s.stock = null; }
  try {
    var t = new Date(); t.setHours(0, 0, 0, 0); t.setDate(t.getDate() + 1);
    var items = rem_collect_(t);
    s.tomorrow = { reminders: items.length,
                   appointments: items.filter(function (x) { return x.kind === 'APPT'; }).length,
                   followUps: items.filter(function (x) { return x.kind === 'FOLLOWUP'; }).length,
                   vaccines: items.filter(function (x) { return x.kind === 'VACCINE'; }).length,
                   noConsent: items.filter(function (x) { return x.state === 'NO_CONSENT'; }).length,
                   whatsappAuto: !!rem_provider_(), emailAuto: rem_autoEmail_(),
                   withEmail: items.filter(function (x) { return x.state === 'READY' && x.email; }).length };
  } catch (e) { s.tomorrow = null; }
  s.signIn = ops_signInTrouble_(todayKey);
  try { s.backup = ops_lastBackup_(); } catch (e) { s.backup = null; }
  return s;
}

function ops_esc_(v) {
  return String(v === null || v === undefined ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** The summary as an email body. Figures only — no patient is named. */
function ops_summaryHtml_(s) {
  var row = function (k, v, warn) {
    return '<tr><td style="padding:4px 12px 4px 0;color:#555">' + ops_esc_(k) + '</td>' +
           '<td style="padding:4px 0;font-weight:600;' + (warn ? 'color:#b45309' : '') + '">' + ops_esc_(v) + '</td></tr>';
  };
  var block = function (title, rows) {
    return '<h3 style="font:600 15px Arial;margin:18px 0 6px;color:#1e293b">' + ops_esc_(title) + '</h3>' +
           '<table style="font:14px Arial;border-collapse:collapse">' + rows.join('') + '</table>';
  };
  var html = ['<div style="font:14px Arial;color:#111;max-width:620px">',
    '<h2 style="font:700 18px Arial;margin:0 0 4px">' + ops_esc_(s.clinic) + ' — end of day</h2>',
    '<div style="color:#666">' + ops_esc_(s.date) + '</div>'];

  var o = s.ops || {}, r = s.revenue || {}, c = s.census || {}, ip = s.ip || {};
  html.push(block('Patients', [
    row('Appointments today', (o.apptToday || 0) + ' (' + (o.apptCompleted || 0) + ' seen, ' + (o.apptPending || 0) + ' not closed)'),
    row('New registrations', s.newPatients === null ? '—' : s.newPatients),
    row('Admitted / discharged today', (ip.admitted || 0) + ' / ' + (ip.discharged || 0)),
    row('In-patients now', c.currentInpatients || 0),
    row('Beds in use', (c.occupied || 0) + ' of ' + (c.totalBeds || 0)),
    row('Lab orders still open', o.labPending || 0, (o.labPending || 0) > 0)
  ]));
  if (s.revenue) {
    html.push(block('Money received today', [
      row('Consultations', ops_money_(r.consultToday)),
      row('Hospital bills', ops_money_(r.hospitalToday)),
      row('Pharmacy', ops_money_(r.pharmacyToday)),
      row('Laboratory', ops_money_(r.labToday)),
      row('Total', ops_money_(r.totalToday)),
      row('Owed to the clinic (all open balances)', ops_money_(r.pendingCredit) + '  — hospital ' + ops_money_(r.creditHospital) +
          ', pharmacy ' + ops_money_(r.creditPharmacy) + ', lab ' + ops_money_(r.creditLab), r.pendingCredit > 0)
    ]));
  }
  if (s.stock) {
    var sc = s.stock.counts;
    var rows = [
      row('Running low', sc.low + (sc.outOfStock ? ' (' + sc.outOfStock + ' out of stock)' : ''), sc.low > 0),
      row('Expiring within 30 days', sc.urgent + ' batch(es)', sc.urgent > 0),
      row('Already expired, still on the shelf', sc.expired + ' batch(es), ' + ops_money_(s.stock.values.expired) + ' at cost', sc.expired > 0)
    ];
    s.stock.low.forEach(function (x) { rows.push(row('  · ' + x.brand, x.reason)); });
    html.push(block('Pharmacy stock', rows));
  }
  if (s.tomorrow) {
    var t = s.tomorrow;
    html.push(block('Tomorrow', [
      row('Reminders due', t.reminders + ' (' + t.appointments + ' appointments, ' + t.followUps + ' follow-ups, ' + t.vaccines + ' vaccinations)'),
      row('Cannot be messaged (no consent)', t.noConsent, t.noConsent > 0),
      row('Sent by', (t.emailAuto ? 'email automatically at 6 pm (' + t.withEmail + ' with an address); ' : '') +
                     (t.whatsappAuto ? 'WhatsApp automatically at 6 pm' : 'WhatsApp from the desk — Operations → Reminders'))
    ]));
  }
  var si = s.signIn || {};
  html.push(block('Sign-in', [
    row('Failed attempts', si.failed || 0, (si.failed || 0) > 10),
    row('Accounts locked', (si.locked || 0) + (si.lockedUsers && si.lockedUsers.length ? ' (' + si.lockedUsers.join(', ') + ')' : ''), (si.locked || 0) > 0),
    row('Two-step code failures', si.mfaFailed || 0, (si.mfaFailed || 0) > 0)
  ]));
  var b = s.backup;
  html.push(block('Backup', [
    b ? row('Last backup', b.at + ' — ' + b.status + (b.verified ? ', verified' : ''), b.status !== 'OK' || b.ageHours > 30)
      : row('Last backup', 'none on record', true)
  ]));
  html.push('<p style="color:#888;font-size:12px;margin-top:20px">Sent by CresRx at ' + ops_esc_(s.generatedAt) +
            '. Figures only; no patient is named in this email.</p></div>');
  return html.join('');
}

/** TIME-DRIVEN, 9 pm. Emails the summary to the owner. */
function ownerDailySummary(e) {
  crescTriggerOnly_(e, 'ownerDailySummary');
  return ops_sendSummary_();
}

function ops_sendSummary_() {
  var to = ops_ownerEmails_();
  if (!to.length) {
    Logger.log('Owner summary: no recipient. Set the OWNER_EMAIL script property, or an Email Address on a Super_Admin user.');
    return { success: false, message: 'No owner email is set. Add OWNER_EMAIL in Script Properties, or an email on the owner\'s Users row.' };
  }
  var s = ops_buildSummary_();
  var subject = s.clinic + ' — day summary ' + s.date +
    (s.revenue ? ' · ' + ops_money_(s.revenue.totalToday) : '');
  GmailApp.sendEmail(to.join(','), subject, 'This summary is best read as HTML.', { htmlBody: ops_summaryHtml_(s), name: 'CresRx' });
  return { success: true, message: 'Summary sent to ' + to.length + ' address(es).' };
}

/** FRONTEND ENTRY. Today's summary on screen, and optionally emailed now. */
function getOwnerSummary(sessionToken, sendNow) {
  try {
    crescRequire_(sessionToken, 'accounts.read');
    var s = ops_buildSummary_();
    var out = { success: true, summary: s, html: ops_summaryHtml_(s), recipients: ops_ownerEmails_().length };
    if (sendNow) {
      crescRequire_(sessionToken, 'admin.config');
      var r = ops_sendSummary_();
      out.sent = r.success; out.message = r.message;
    }
    return out;
  } catch (err) {
    return { success: false, message: cresc_reason_(err) };
  }
}

// ---------------------------------------------------------------------------
// NIGHTLY BACKUP
// ---------------------------------------------------------------------------

function ops_backupLog_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(OPS_BACKUP.LOG);
  if (!sh) {
    sh = ss.insertSheet(OPS_BACKUP.LOG);
    sh.appendRow(OPS_BACKUP.HEADERS);
    sh.getRange(1, 1, 1, OPS_BACKUP.HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** The restricted folder: private, owned by the account the job runs as. */
function ops_backupFolder_() {
  var id = ops_prop_('BACKUP_FOLDER_ID');
  var folder = null;
  if (id) { try { folder = DriveApp.getFolderById(id); } catch (e) { folder = null; } }
  if (!folder) {
    folder = DriveApp.createFolder(OPS_BACKUP.FOLDER_NAME);
    PropertiesService.getScriptProperties().setProperty('BACKUP_FOLDER_ID', folder.getId());
  }
  // A backup is every patient record in the clinic: nobody but the owner.
  try { folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); } catch (e) {}
  return folder;
}

/** Take everybody but the owner off a file. @return {number} removed */
function ops_restrict_(file) {
  var n = 0;
  try { file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); } catch (e) {}
  var owner = '';
  try { owner = file.getOwner().getEmail(); } catch (e) {}
  try { file.getEditors().forEach(function (u) { if (u.getEmail() !== owner) { file.removeEditor(u); n++; } }); } catch (e) {}
  try { file.getViewers().forEach(function (u) { if (u.getEmail() !== owner) { file.removeViewer(u); n++; } }); } catch (e) {}
  return n;
}

/** Row counts of the checked sheets in one spreadsheet. */
function ops_counts_(ss) {
  var out = {};
  var names = OPS_BACKUP.CHECK_SHEETS.slice();
  if (typeof DPDP_CFG !== 'undefined' && DPDP_CFG.CONSENT) names.push(DPDP_CFG.CONSENT);
  names.forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (sh) out[n] = sh.getLastRow();
  });
  out.__sheets = ss.getSheets().length;
  return out;
}

/** Old copies out: keep BACKUP_KEEP_DAYS, and each month's first for a year. */
function ops_prune_(folder) {
  var keepDays = parseInt(ops_prop_('BACKUP_KEEP_DAYS'), 10) || 30;
  var now = Date.now(), DAY = 86400000, pruned = 0;
  var files = [], it = folder.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    if (f.getName().indexOf(OPS_BACKUP.PREFIX) === 0) files.push(f);
  }
  files.sort(function (a, b) { return a.getDateCreated() - b.getDateCreated(); });
  var monthFirst = {};
  files.forEach(function (f) {
    var m = Utilities.formatDate(f.getDateCreated(), ops_tz_(), 'yyyy-MM');
    if (!monthFirst[m]) monthFirst[m] = f.getId();
  });
  files.forEach(function (f) {
    var age = (now - f.getDateCreated().getTime()) / DAY;
    if (age <= keepDays) return;
    if (monthFirst[Utilities.formatDate(f.getDateCreated(), ops_tz_(), 'yyyy-MM')] === f.getId() && age <= 366) return;
    try { f.setTrashed(true); pruned++; } catch (e) {}
  });
  return pruned;
}

/** One backup, verified, logged. @return {{status, fileId, verified, detail}} */
function ops_backup_(by) {
  var log = ops_backupLog_();
  var stamp = Utilities.formatDate(new Date(), ops_tz_(), 'yyyy-MM-dd HHmm');
  var name = OPS_BACKUP.PREFIX + stamp;
  try {
    // Not under the script lock: a large copy takes long enough that every
    // desk saving at that moment would time out waiting. Rows appended
    // during the copy are why the check below accepts "more", never "fewer".
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    SpreadsheetApp.flush();
    var before = ops_counts_(ss);
    var folder = ops_backupFolder_();
    var copy = DriveApp.getFileById(ss.getId()).makeCopy(name, folder);

    var removed = ops_restrict_(copy);
    var after = ops_counts_(SpreadsheetApp.openById(copy.getId()));
    var diffs = [];
    Object.keys(before).forEach(function (k) {
      if (!(after[k] >= before[k])) diffs.push(k + ' ' + before[k] + '→' + (after[k] === undefined ? 'missing' : after[k]));
    });
    var pruned = ops_prune_(folder);
    var verified = !diffs.length;
    var detail = verified
      ? before.__sheets + ' sheets; row counts match on ' + (Object.keys(before).length - 1) + ' key sheets'
      : 'MISMATCH: ' + diffs.join(', ');
    if (removed) detail += '; ' + removed + ' inherited share(s) removed';
    log.appendRow([new Date(), copy.getId(), name, verified ? 'OK' : 'CHECK', verified ? 'YES' : 'NO', detail, pruned, by || 'SYSTEM']);
    return { status: verified ? 'OK' : 'CHECK', fileId: copy.getId(), verified: verified, detail: detail, pruned: pruned };
  } catch (e) {
    try { log.appendRow([new Date(), '', name, 'FAILED', 'NO', String(e.message).slice(0, 300), 0, by || 'SYSTEM']); } catch (x) {}
    return { status: 'FAILED', fileId: '', verified: false, detail: e.message };
  }
}

function ops_lastBackup_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OPS_BACKUP.LOG);
  if (!sh || sh.getLastRow() < 2) return null;
  var r = sh.getRange(sh.getLastRow(), 1, 1, OPS_BACKUP.HEADERS.length).getValues()[0];
  var at = r[0] instanceof Date ? r[0] : new Date(r[0]);
  return { at: Utilities.formatDate(at, ops_tz_(), 'dd-MMM HH:mm'), status: String(r[3]),
           verified: String(r[4]) === 'YES', detail: String(r[5]),
           ageHours: Math.round((Date.now() - at.getTime()) / 3600000) };
}

/** TIME-DRIVEN, 1 am. */
function opsNightlyBackup(e) {
  crescTriggerOnly_(e, 'opsNightlyBackup');
  var r = ops_backup_('SYSTEM');
  Logger.log('Backup: ' + r.status + ' — ' + r.detail);
  return r;
}

/** The restore checklist, the same text as docs/RESTORE_CHECKLIST.md. */
var OPS_RESTORE_CHECKLIST = [
  'Decide what you are restoring: ONE sheet (a bad edit, a deleted row) or EVERYTHING (the spreadsheet is lost or corrupted). One sheet is almost always the answer.',
  'Stop the damage first: switch the web app to "Only myself" (Deploy → Manage deployments) so nobody writes while you restore. Tell the desks.',
  'Pick the backup: Admin Dashboard → Operations → Backups lists them. Choose the newest one from BEFORE the problem. Check its "Verified" column says YES.',
  'ONE SHEET: open the backup, right-click the sheet tab → Copy to → the live spreadsheet. In the live file, rename the damaged sheet to "<name>_damaged", rename the copy to the exact original name, and drag it to the same position. Formulas and other modules find sheets by name.',
  'Rows added after the backup: compare the "_damaged" sheet with the restored one and copy across any rows added since the backup time. Appointments, invoices and results added today are the usual ones.',
  'EVERYTHING: make a copy of the backup (File → Make a copy), then point the deployment at it — open Extensions → Apps Script in the COPY, Deploy → New deployment, and update the link staff use. Script properties (API keys, clinic name, backup folder) do NOT travel with a copy: re-enter them from Project Settings of the old project.',
  'Reinstall the scheduled jobs in the restored project: run opsInstallTriggers() and dpdpInstallTriggers() from the editor. Triggers do not copy.',
  'Check: sign in as a test user, open a patient, open today\'s appointments, the lab desk and the pharmacy. Run verifyDeployment() from the editor.',
  'Sessions: everyone signs in again — the session cache does not carry over. Passwords and two-step sign-in do (they are in the Users sheet).',
  'Record it: add a row to Backup_Log saying who restored what, from which backup, and why. Under the DPDP Act a lost or corrupted patient record may be a personal data breach — check with the grievance officer whether it must be reported.',
  'Re-open the web app to its normal audience, and delete the "_damaged" sheet only once you are sure nothing in it is still needed.'
];

/** FRONTEND ENTRY. Backups, schedule, and the checklist. */
function getBackupStatus(sessionToken) {
  try {
    crescRequire_(sessionToken, 'admin.config');
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OPS_BACKUP.LOG);
    var rows = [];
    if (sh && sh.getLastRow() > 1) {
      var n = Math.min(20, sh.getLastRow() - 1);
      rows = sh.getRange(sh.getLastRow() - n + 1, 1, n, OPS_BACKUP.HEADERS.length).getValues().reverse().map(function (r) {
        var at = r[0] instanceof Date ? r[0] : new Date(r[0]);
        return { at: Utilities.formatDate(at, ops_tz_(), 'dd-MMM-yyyy HH:mm'), fileId: String(r[1]),
                 name: String(r[2]), status: String(r[3]), verified: String(r[4]) === 'YES',
                 detail: String(r[5]), pruned: r[6], by: String(r[7]),
                 url: r[1] ? 'https://docs.google.com/spreadsheets/d/' + r[1] + '/edit' : '' };
      });
    }
    var folderId = ops_prop_('BACKUP_FOLDER_ID');
    return { success: true, rows: rows, last: ops_lastBackup_(),
             folderUrl: folderId ? 'https://drive.google.com/drive/folders/' + folderId : '',
             keepDays: parseInt(ops_prop_('BACKUP_KEEP_DAYS'), 10) || 30,
             triggers: ops_triggerState_(), checklist: OPS_RESTORE_CHECKLIST };
  } catch (err) {
    return { success: false, message: cresc_reason_(err) };
  }
}

/** FRONTEND ENTRY. A backup now (before a risky change, say). */
function runBackupNow(sessionToken) {
  try {
    var actor = crescRequire_(sessionToken, 'admin.config');
    var r = ops_backup_(actor.username);
    return { success: r.status !== 'FAILED', status: r.status, detail: r.detail,
             message: r.status === 'OK' ? 'Backup taken and verified.' : (r.status === 'CHECK' ? 'Backup taken, but the check found differences: ' + r.detail : 'Backup failed: ' + r.detail) };
  } catch (err) {
    return { success: false, message: cresc_reason_(err) };
  }
}

// ---------------------------------------------------------------------------
// THE SCHEDULE
// ---------------------------------------------------------------------------

function ops_triggerState_() {
  var have = {};
  try { ScriptApp.getProjectTriggers().forEach(function (t) { have[t.getHandlerFunction()] = true; }); } catch (e) {}
  return OPS_TRIGGERS.map(function (t) {
    return { fn: t.fn, hour: t.hour, what: t.what, installed: !!have[t.fn] };
  });
}

/** ADMIN, from the script editor. Installs (or reinstalls) the three jobs. */
function opsInstallTriggers() {
  crescEditorOnly_('opsInstallTriggers');
  var names = OPS_TRIGGERS.map(function (t) { return t.fn; });
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (names.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  OPS_TRIGGERS.forEach(function (t) {
    ScriptApp.newTrigger(t.fn).timeBased().atHour(t.hour).everyDays(1).inTimezone(ops_tz_()).create();
  });
  var msg = 'Installed:\n' + OPS_TRIGGERS.map(function (t) {
    return '  ' + t.fn + '  daily ~' + (t.hour < 10 ? '0' : '') + t.hour + ':00  ' + t.what;
  }).join('\n') + '\n\nThey run as ' + (function () {
    try { return Session.getEffectiveUser().getEmail(); } catch (e) { return 'the installing account'; }
  })() + '; backups are stored in that account\'s Drive.';
  Logger.log(msg);
  return msg;
}
