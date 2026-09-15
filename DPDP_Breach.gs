// ============================================================================
// DPDP_Breach.gs — Crescentia HealthTech
// Section 8(6): noticing, deciding, and telling.
// ----------------------------------------------------------------------------
// WHAT THE ACT REQUIRES
//
// s.8(6): on a personal data breach, the Data Fiduciary must notify the Data
// Protection Board AND every affected Data Principal. There is no materiality
// threshold in the section and no "we judged it minor" exemption — the
// judgement a clinic makes is about WHAT HAPPENED, not about whether to say.
//
// WHAT WAS MISSING (finding H5, HIGH)
//
// Audit_Log and Auth_Audit existed and were append-only, which is the hard
// part and was already done. Nothing read them. No procedure said who decides
// something is a breach, who notifies, in what time, using what words. So the
// clinic could answer "did anyone take the register" only by scrolling a
// spreadsheet nobody scrolls.
//
// WHAT THIS FILE PROVIDES
//
//   dpdpAnomalyScan()      reads the audit log for the four patterns that
//                          actually precede a disclosure here
//   Breach_Register        one row per incident, INCLUDING the ones judged
//                          not notifiable — the record of having considered
//                          it is worth as much as the conclusion, and in an
//                          inquiry it is the only thing that shows the clinic
//                          was looking
//   dpdpBreachNotice()     the s.8(6) wording for the Board and for the
//                          patient, filled in from the row
//
// WHAT IT CANNOT DO
//
// Apps Script does not give a web app the caller's IP address, so "signed in
// from where" is out of reach: every pattern below is about WHO, WHAT and
// WHEN. Said plainly rather than left as a gap somebody assumes is covered.
// ============================================================================

var DPDP_BREACH_CFG = {
  SHEET: 'Breach_Register',

  /** Failed sign-ins against one account inside the window before it is worth
   *  a human's attention. The lockout already stops five in fifteen minutes;
   *  this is the slower, quieter version that a lockout never trips. */
  FAILED_LOGINS_PER_ACCOUNT: 8,

  /** Failed sign-ins across ALL accounts in the window — someone working
   *  through patient IDs trips this without ever tripping the per-account
   *  threshold. */
  FAILED_LOGINS_TOTAL: 25,

  /** Patient records one person opens, prints or exports in a day before it
   *  stops looking like a clinic day and starts looking like a copy. */
  READS_PER_ACTOR_PER_DAY: 60,

  /** One shared document opened this many times is not one patient reading
   *  their report. */
  OPENS_PER_DOCUMENT: 6,

  /** The window the weekly review looks back over. */
  REVIEW_DAYS: 7
};

/** Audit events that mean somebody LOOKED at personal data. */
var DPDP_READ_EVENTS = [
  'DOCUMENT_OPENED', 'DPDP_DATA_EXPORTED', 'CLINICAL_DOCUMENT_READ',
  'PATIENT_SEARCHED', 'PATIENT_RECORD_READ'
];

function dpdp_breachSheet_() {
  return dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), DPDP_BREACH_CFG.SHEET, [
    'Breach_ID', 'Detected_At', 'Detected_By', 'Source', 'Category', 'Summary',
    'Affected_Count', 'Affected_Patients', 'Severity', 'Notifiable',
    'Decision_Rationale', 'Board_Notified_At', 'Principals_Notified_At',
    'Status', 'Closed_At', 'Closed_By'
  ]);
}

// ---------------------------------------------------------------------------
// SECTION A — READING THE AUDIT LOG
// ---------------------------------------------------------------------------

/**
 * The four patterns, over the last `days`. PRIVATE: the weekly trigger runs
 * as the owner with no session, so the public wrapper below carries the guard
 * and this one carries the work.
 *
 * Each finding says what was seen and what to do about it, because a weekly
 * email that says "3 anomalies" and nothing else gets filed unread.
 */
function dpdp_anomalyScan_(days) {
  var window = (days && days > 0) ? days : DPDP_BREACH_CFG.REVIEW_DAYS;
  var since = Date.now() - window * 86400000;
  var findings = [];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Audit_Log');
  if (!sh || sh.getLastRow() < 2) {
    return { success: true, windowDays: window, findings: [],
             message: 'No audit log to review.' };
  }

  var values = dc_sheetValues_(sh);
  var head = {};
  (values[0] || []).forEach(function (h, i) { head[dpdp_str_(h)] = i; });

  var failedByAccount = {}, failedTotal = 0;
  var readsByActor = {}, opensByDoc = {};
  var legacyRefusals = 0, offHours = {};
  var rows = 0;

  for (var i = 1; i < values.length; i++) {
    var when = (typeof cresc_parseDate_ === 'function')
      ? cresc_parseDate_(values[i][head['Timestamp']])
      : new Date(values[i][head['Timestamp']]);
    if (!when || isNaN(when.getTime()) || when.getTime() < since) continue;
    rows++;

    var event = dpdp_str_(values[i][head['Event']]).toUpperCase();
    var actor = dpdp_str_(values[i][head['Actor_Username']]) || '(blank)';
    var entityId = dpdp_str_(values[i][head['Entity_ID']]);

    if (event === 'LOGIN_FAILED' || event === 'LOGIN_UNKNOWN_USER') {
      failedByAccount[actor] = (failedByAccount[actor] || 0) + 1;
      failedTotal++;
    }
    if (event === 'LOGIN_LEGACY_CREDENTIAL') legacyRefusals++;

    if (DPDP_READ_EVENTS.indexOf(event) !== -1) {
      var day = Utilities.formatDate(when, DPDP_CFG.TZ, 'yyyy-MM-dd');
      var key = actor + '|' + day;
      readsByActor[key] = (readsByActor[key] || 0) + 1;
    }
    if (event === 'DOCUMENT_OPENED') {
      opensByDoc[entityId] = (opensByDoc[entityId] || 0) + 1;
    }

    // 22:00-05:59. A ward runs at night and this is not by itself wrong; a
    // burst of RECORD reads at 3am by someone who is not on nights is.
    var hour = parseInt(Utilities.formatDate(when, DPDP_CFG.TZ, 'H'), 10);
    if ((hour >= 22 || hour < 6) && DPDP_READ_EVENTS.indexOf(event) !== -1) {
      offHours[actor] = (offHours[actor] || 0) + 1;
    }
  }

  function add(severity, title, detail, action) {
    findings.push({ severity: severity, title: title, detail: detail, action: action });
  }

  Object.keys(failedByAccount).forEach(function (who) {
    if (failedByAccount[who] < DPDP_BREACH_CFG.FAILED_LOGINS_PER_ACCOUNT) return;
    add('HIGH', 'Repeated failed sign-ins for "' + who + '"',
        failedByAccount[who] + ' failed attempts in ' + window + ' days.',
        'If this is not the account holder forgetting their password, reset it ' +
        '(crescAdminResetPassword) and check Audit_Log for a success that follows ' +
        'the failures — that is the row that matters.');
  });

  if (failedTotal >= DPDP_BREACH_CFG.FAILED_LOGINS_TOTAL) {
    add('HIGH', 'Failed sign-ins across many accounts',
        failedTotal + ' failed attempts in ' + window + ' days, spread over ' +
        Object.keys(failedByAccount).length + ' identifiers.',
        'This is the shape of someone walking the patient ID range. Patient IDs ' +
        'are sequential and printed on every barcode, so the list is public; the ' +
        'password is what stops it. Confirm no plain-text credentials remain ' +
        '(crescCredentialStatus).');
  }

  Object.keys(readsByActor).forEach(function (key) {
    if (readsByActor[key] < DPDP_BREACH_CFG.READS_PER_ACTOR_PER_DAY) return;
    var parts = key.split('|');
    add('MEDIUM', 'Unusual volume of record access by "' + parts[0] + '"',
        readsByActor[key] + ' patient records read, printed or exported on ' + parts[1] + '.',
        'Ask them what they were doing. A day of chart audit looks exactly like ' +
        'this and is a perfectly good answer — but it should be an answer ' +
        'somebody gave, not an assumption.');
  });

  Object.keys(opensByDoc).forEach(function (doc) {
    if (opensByDoc[doc] < DPDP_BREACH_CFG.OPENS_PER_DOCUMENT) return;
    add('MEDIUM', 'One shared document opened ' + opensByDoc[doc] + ' times',
        'Document link ' + doc + '.',
        'A report a patient reads once or twice does not get opened this often. ' +
        'The likeliest explanation is a forwarded message. Withdraw the link ' +
        '(dpdpRevokeDocumentLink) and consider whether the patient should be told.');
  });

  if (legacyRefusals) {
    add('HIGH', 'Accounts still on plain-text passwords',
        legacyRefusals + ' sign-in attempt(s) were refused because the stored ' +
        'credential is still in clear text.',
        'Those passwords are readable by everyone with access to the spreadsheet. ' +
        'Run crescMigrateCredentials() — until then those people cannot sign in.');
  }

  Object.keys(offHours).forEach(function (who) {
    if (offHours[who] < 10) return;
    add('LOW', 'Record access outside working hours by "' + who + '"',
        offHours[who] + ' reads between 22:00 and 06:00.',
        'Normal on a ward with night duty. Worth one question if this account ' +
        'does not do nights.');
  });

  var order = { HIGH: 0, MEDIUM: 1, LOW: 2 };
  findings.sort(function (a, b) { return order[a.severity] - order[b.severity]; });

  return { success: true, windowDays: window, rowsExamined: rows,
           findings: findings,
           message: findings.length
             ? findings.length + ' thing(s) worth a look in the last ' + window + ' days.'
             : 'Nothing unusual in the last ' + window + ' days (' + rows + ' audit rows).' };
}

/** FRONTEND ENTRY. The same scan, for whoever answers for data protection. */
function dpdpAnomalyScan(days, sessionToken) {
  try {
    crescRequire_(sessionToken, ['dpdp.manage', 'admin.audit']);
    return dpdp_anomalyScan_(days);
  } catch (err) {
    return { success: false, findings: [],
             message: String(err.message || err).replace('FORBIDDEN: ', '') };
  }
}

// ---------------------------------------------------------------------------
// SECTION B — THE REGISTER
// ---------------------------------------------------------------------------

/**
 * FRONTEND ENTRY. Records an incident. Raise one for anything that MIGHT be a
 * breach — the decision comes next, in dpdpAssessBreach(), and a register
 * that only holds confirmed breaches is a register that proves nothing about
 * the ones you decided were fine.
 *
 * @param {{summary, category, source, severity, affectedCount,
 *          affectedPatients}} payload
 */
function dpdpRaiseBreach(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    payload = payload || {};
    var actor = crescRequire_(sessionToken, ['dpdp.manage', 'admin.config', 'admin.audit']);

    var summary = dpdp_str_(payload.summary);
    if (summary.length < 20) {
      return { success: false,
               message: 'Describe what happened in a sentence or two. A breach ' +
                        'register entry that says "issue" tells an inquiry nothing.' };
    }

    var now = dpdp_now_();
    var id = 'BRC-' + Utilities.formatDate(now, DPDP_CFG.TZ, 'yyMMdd-HHmmss') + '-' +
             Utilities.getUuid().substring(0, 4).toUpperCase();

    dpdp_breachSheet_().appendRow([
      id, now, actor.username, dpdp_str_(payload.source) || 'REPORTED',
      dpdp_str_(payload.category) || 'UNCLASSIFIED', summary,
      parseInt(payload.affectedCount, 10) || 0,
      dpdp_str_(payload.affectedPatients),
      dpdp_str_(payload.severity).toUpperCase() || 'UNASSESSED',
      '', '', '', '', 'OPEN', '', ''
    ]);
    dc_invalidate_(DPDP_BREACH_CFG.SHEET);
    SpreadsheetApp.flush();

    try {
      logAudit_({ username: actor.username, role: actor.role },
                'DPDP_BREACH_RAISED', 'Breach', id,
                { category: dpdp_str_(payload.category) });
    } catch (e) {}

    return { success: true, breachId: id,
             message: 'Incident ' + id + ' recorded. Now decide whether it is ' +
                      'notifiable — section 8(6) — and record why.' };
  } catch (err) {
    return { success: false, message: String(err.message || err).replace('FORBIDDEN: ', '') };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * FRONTEND ENTRY. Records the decision, and refuses to take one without a
 * reason.
 *
 * "Not notifiable" is a legitimate conclusion and a dangerous one to reach
 * silently: the thing that makes it defensible two years later is the
 * sentence explaining it, written at the time by a named person.
 */
function dpdpAssessBreach(breachId, notifiable, rationale, sessionToken) {
  try {
    var actor = crescRequire_(sessionToken, ['dpdp.manage', 'admin.config']);
    var why = dpdp_str_(rationale);
    if (why.length < 20) {
      return { success: false,
               message: 'Say why. Both answers need a reason — "not notifiable" ' +
                        'most of all, because that is the one somebody will ask ' +
                        'you to justify.' };
    }

    var sh = dpdp_breachSheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    var want = dpdp_str_(breachId).toUpperCase();

    for (var i = 1; i < (data ? data.length : 0); i++) {
      if (dpdp_str_(data[i][m['Breach_ID']]).toUpperCase() !== want) continue;
      sh.getRange(i + 1, m['Notifiable'] + 1).setValue(notifiable ? 'YES' : 'NO');
      sh.getRange(i + 1, m['Decision_Rationale'] + 1).setValue(
        why + '  [' + actor.username + ', ' + dpdp_fmt_(dpdp_now_()) + ']');
      dc_invalidate_(DPDP_BREACH_CFG.SHEET);
      SpreadsheetApp.flush();
      try {
        logAudit_({ username: actor.username, role: actor.role },
                  'DPDP_BREACH_ASSESSED', 'Breach', want, { notifiable: !!notifiable });
      } catch (e) {}
      return { success: true,
               message: notifiable
                 ? 'Recorded as NOTIFIABLE. Section 8(6) requires the Data ' +
                   'Protection Board AND every affected data principal to be told. ' +
                   'dpdpBreachNotice("' + want + '") gives you both texts.'
                 : 'Recorded as not notifiable, with your reason.' };
    }
    return { success: false, message: 'No incident with the id ' + breachId + '.' };
  } catch (err) {
    return { success: false, message: String(err.message || err).replace('FORBIDDEN: ', '') };
  }
}

/** FRONTEND ENTRY. Stamps that a notification actually went out. */
function dpdpRecordBreachNotification(breachId, who, sessionToken) {
  try {
    var actor = crescRequire_(sessionToken, ['dpdp.manage', 'admin.config']);
    var target = dpdp_str_(who).toUpperCase();
    if (target !== 'BOARD' && target !== 'PRINCIPALS') {
      return { success: false, message: 'Say which: BOARD or PRINCIPALS.' };
    }
    var sh = dpdp_breachSheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    var want = dpdp_str_(breachId).toUpperCase();
    var col = target === 'BOARD' ? 'Board_Notified_At' : 'Principals_Notified_At';

    for (var i = 1; i < (data ? data.length : 0); i++) {
      if (dpdp_str_(data[i][m['Breach_ID']]).toUpperCase() !== want) continue;
      sh.getRange(i + 1, m[col] + 1).setValue(dpdp_now_());
      dc_invalidate_(DPDP_BREACH_CFG.SHEET);
      SpreadsheetApp.flush();
      try {
        logAudit_({ username: actor.username, role: actor.role },
                  'DPDP_BREACH_NOTIFIED', 'Breach', want, { who: target });
      } catch (e) {}
      return { success: true, message: target === 'BOARD'
        ? 'Board notification recorded.'
        : 'Data principal notification recorded.' };
    }
    return { success: false, message: 'No incident with the id ' + breachId + '.' };
  } catch (err) {
    return { success: false, message: String(err.message || err).replace('FORBIDDEN: ', '') };
  }
}

/** FRONTEND ENTRY. Closes an incident, with what was done. */
function dpdpCloseBreach(breachId, outcome, sessionToken) {
  try {
    var actor = crescRequire_(sessionToken, ['dpdp.manage', 'admin.config']);
    var text = dpdp_str_(outcome);
    if (text.length < 20) {
      return { success: false, message: 'Say what was done and what changed so it ' +
                                        'does not happen again.' };
    }
    var sh = dpdp_breachSheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    var want = dpdp_str_(breachId).toUpperCase();

    for (var i = 1; i < (data ? data.length : 0); i++) {
      if (dpdp_str_(data[i][m['Breach_ID']]).toUpperCase() !== want) continue;
      if (!dpdp_str_(data[i][m['Notifiable']])) {
        return { success: false,
                 message: 'This incident has not been assessed yet. Record whether ' +
                          'it is notifiable, and why, before closing it.' };
      }
      sh.getRange(i + 1, m['Status'] + 1).setValue('CLOSED');
      sh.getRange(i + 1, m['Closed_At'] + 1).setValue(dpdp_now_());
      sh.getRange(i + 1, m['Closed_By'] + 1).setValue(actor.username);
      sh.getRange(i + 1, m['Decision_Rationale'] + 1).setValue(
        dpdp_str_(data[i][m['Decision_Rationale']]) + '\n\nCLOSED: ' + text);
      dc_invalidate_(DPDP_BREACH_CFG.SHEET);
      SpreadsheetApp.flush();
      return { success: true, message: 'Incident ' + want + ' closed.' };
    }
    return { success: false, message: 'No incident with the id ' + breachId + '.' };
  } catch (err) {
    return { success: false, message: String(err.message || err).replace('FORBIDDEN: ', '') };
  }
}

/** FRONTEND ENTRY. The register, newest first, unassessed ones marked. */
function dpdpListBreaches(sessionToken) {
  try {
    crescRequire_(sessionToken, ['dpdp.manage', 'admin.audit', 'admin.config']);
    var sh = dpdp_breachSheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    var rows = [];
    for (var i = 1; i < (data ? data.length : 0); i++) {
      rows.push({
        breachId: dpdp_str_(data[i][m['Breach_ID']]),
        detectedAt: dpdp_str_(data[i][m['Detected_At']]),
        detectedBy: dpdp_str_(data[i][m['Detected_By']]),
        source: dpdp_str_(data[i][m['Source']]),
        category: dpdp_str_(data[i][m['Category']]),
        summary: dpdp_str_(data[i][m['Summary']]),
        affectedCount: parseInt(data[i][m['Affected_Count']], 10) || 0,
        severity: dpdp_str_(data[i][m['Severity']]),
        notifiable: dpdp_str_(data[i][m['Notifiable']]),
        assessed: !!dpdp_str_(data[i][m['Notifiable']]),
        boardNotifiedAt: dpdp_str_(data[i][m['Board_Notified_At']]),
        principalsNotifiedAt: dpdp_str_(data[i][m['Principals_Notified_At']]),
        status: dpdp_str_(data[i][m['Status']])
      });
    }
    rows.reverse();
    var unassessed = rows.filter(function (r) { return !r.assessed; }).length;
    return { success: true, rows: rows, unassessed: unassessed, message: '' };
  } catch (err) {
    return { success: false, rows: [],
             message: String(err.message || err).replace('FORBIDDEN: ', '') };
  }
}

/**
 * FRONTEND ENTRY. The two notifications s.8(6) requires, filled in from the
 * row and written in the two different registers they need.
 *
 * The Board's version is factual and complete. The patient's version leads
 * with what it means for THEM, because a notification that opens with
 * "pursuant to section 8(6)" is one the recipient stops reading before the
 * part that tells them what to do.
 */
function dpdpBreachNotice(breachId, sessionToken) {
  try {
    crescRequire_(sessionToken, ['dpdp.manage', 'admin.config']);
    var sh = dpdp_breachSheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    var want = dpdp_str_(breachId).toUpperCase();
    var clinic = (typeof cresc_clinic_ === 'function')
      ? cresc_clinic_() : { name: 'the clinic', address: '', phone: '', email: '' };
    var officer = dpdp_officer_();

    for (var i = 1; i < (data ? data.length : 0); i++) {
      if (dpdp_str_(data[i][m['Breach_ID']]).toUpperCase() !== want) continue;

      var summary = dpdp_str_(data[i][m['Summary']]);
      var detected = dpdp_str_(data[i][m['Detected_At']]);
      var count = parseInt(data[i][m['Affected_Count']], 10) || 0;
      var contact = officer
        ? officer.name + (officer.email ? ', ' + officer.email : '') +
          (officer.phone ? ', ' + officer.phone : '')
        : '[NO GRIEVANCE OFFICER NAMED — section 13 requires one before you send this]';

      var board = [
        'To: Data Protection Board of India',
        'Subject: Intimation of a personal data breach under section 8(6), ' +
          'Digital Personal Data Protection Act, 2023',
        '',
        'Data Fiduciary: ' + clinic.name,
        'Address: ' + (clinic.address || '[address]'),
        'Reference: ' + want,
        '',
        '1. Nature of the breach',
        '   ' + summary,
        '',
        '2. When it was detected',
        '   ' + detected,
        '',
        '3. Categories of personal data involved',
        '   [Name, patient identifier, contact details, clinical records — ' +
          'delete what does not apply and add what does.]',
        '',
        '4. Number of data principals affected',
        '   ' + (count || '[number]'),
        '',
        '5. Likely consequences',
        '   [State them plainly. Health data disclosed to an unauthorised person ' +
          'carries a risk of distress, stigma and discrimination; say so if it ' +
          'applies.]',
        '',
        '6. Measures taken and proposed',
        '   [What was done immediately, what has changed so it cannot recur.]',
        '',
        '7. Intimation to affected data principals',
        '   [When and how they were told.]',
        '',
        'Contact for this intimation: ' + contact
      ].join('\n');

      var principal = [
        'Dear [patient name],',
        '',
        'We are writing to tell you about something that went wrong with your ' +
        'personal information at ' + clinic.name + ', because you have a right ' +
        'to know and because there may be something you want to do about it.',
        '',
        'WHAT HAPPENED',
        summary,
        '',
        'WHEN',
        'We found out on ' + detected + '.',
        '',
        'WHAT INFORMATION WAS INVOLVED',
        '[Say exactly what — name, phone number, your test results. Do not ' +
        'write "certain data".]',
        '',
        'WHAT WE HAVE DONE',
        '[What was stopped, closed or changed.]',
        '',
        'WHAT YOU CAN DO',
        '[Anything practical: change your portal password, be alert to calls ' +
        'claiming to be from us. If there is nothing to do, say that too.]',
        '',
        'If you have any question about this, or you are unhappy with how we ' +
        'have handled it, please contact ' + contact + '. You may also complain ' +
        'to the Data Protection Board of India.',
        '',
        'We are sorry.',
        '',
        clinic.name,
        (clinic.phone || '') + (clinic.email ? '  ' + clinic.email : '')
      ].join('\n');

      return { success: true, breachId: want, board: board, principal: principal,
               message: 'Fill in every square bracket before sending either of these.' };
    }
    return { success: false, message: 'No incident with the id ' + breachId + '.' };
  } catch (err) {
    return { success: false, message: String(err.message || err).replace('FORBIDDEN: ', '') };
  }
}
