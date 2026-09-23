// ============================================================================
// DPDP_Triggers.gs — Crescentia HealthTech
// The part of compliance that has to happen on a day nobody remembers it.
// ----------------------------------------------------------------------------
// Three of the obligations in this repository are only met if something runs
// on its own:
//
//   s.8(7)  a shared document link that expires only expires if something
//           expires it
//   s.8(6)  a breach nobody is looking for is found by the person it was
//           taken from
//   s.8(7)  sessions and logs that accumulate for ever are data kept longer
//           than necessary, however good the intention was
//
// dpdpInstallTriggers() sets up the three time-driven triggers that do them,
// and is safe to re-run: it removes its own before adding them, so running it
// twice does not give you six.
//
// WHAT RUNS AS WHOM. A time-driven trigger runs as the account that installed
// it, with no session and no browser. Nothing in here calls a guarded
// endpoint for that reason — the private forms (dpdp_anomalyScan_) exist so
// the trigger can do the work without a token it has no way to hold.
//
//   dpdpInstallTriggers()   once, from the script editor
//   dpdpTriggerStatus()     what is installed now
//   dpdpRemoveTriggers()    take them off again
// ============================================================================

/** Functions this file owns. Anything else on the trigger list is left alone. */
var DPDP_TRIGGER_FUNCTIONS = ['dpdpDailyMaintenance', 'dpdpWeeklyReview',
                              'dpdpMonthlyRetentionReport'];

/**
 * ADMIN, from the script editor. Installs the three jobs.
 *
 * Times are chosen for a clinic: the daily sweep at 2am when nobody is
 * dispatching documents, the weekly review early on Monday so it is in the
 * officer's inbox before the week starts, and the retention report on the
 * first of the month.
 */
function dpdpInstallTriggers() {
  crescEditorOnly_('dpdpInstallTriggers');
  var removed = dpdpRemoveTriggers();

  ScriptApp.newTrigger('dpdpDailyMaintenance')
    .timeBased().atHour(2).everyDays(1).inTimezone(DPDP_CFG.TZ).create();

  ScriptApp.newTrigger('dpdpWeeklyReview')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7)
    .inTimezone(DPDP_CFG.TZ).create();

  ScriptApp.newTrigger('dpdpMonthlyRetentionReport')
    .timeBased().onMonthDay(1).atHour(6).inTimezone(DPDP_CFG.TZ).create();

  var msg = [
    'DPDP triggers installed' + (removed ? ' (' + removed + ' old one(s) removed)' : '') + ':',
    '  dpdpDailyMaintenance        daily, 02:00 — expire document links and any',
    '                              legacy public Drive shares, purge dead sessions',
    '  dpdpWeeklyReview            Monday 07:00 — read the audit log for the four',
    '                              patterns that precede a disclosure, and email',
    '                              the grievance officer',
    '  dpdpMonthlyRetentionReport  1st, 06:00 — what is past its retention period',
    '',
    'They run as ' + (function () {
      try { return Session.getEffectiveUser().getEmail(); } catch (e) { return 'the installing account'; }
    })() + '. If that account is removed from the clinic, reinstall them as somebody else.',
    '',
    'The weekly review needs somebody to send to: run',
    '  dpdpSetGrievanceOfficer("Dr. …", "officer@clinic.in", "+91 …")'
  ].join('\n');
  Logger.log(msg);
  return msg;
}

/** ADMIN. Removes the triggers this file installs, and reports how many. */
function dpdpRemoveTriggers() {
  crescEditorOnly_('dpdpRemoveTriggers');
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (DPDP_TRIGGER_FUNCTIONS.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t); n++;
    }
  });
  return n;
}

/** ADMIN. What is installed, so "we added a trigger" can be checked. */
function dpdpTriggerStatus() {
  crescEditorOnly_('dpdpTriggerStatus', ['dpdp.manage', 'admin.config', 'admin.audit']);
  var mine = [], others = [];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var line = t.getHandlerFunction() + '  (' + t.getEventType() + ')';
    if (DPDP_TRIGGER_FUNCTIONS.indexOf(t.getHandlerFunction()) !== -1) mine.push(line);
    else others.push(line);
  });
  var report = [
    'DPDP triggers: ' + (mine.length ? '\n  ' + mine.join('\n  ') : 'NONE INSTALLED.'),
    '',
    'Other project triggers: ' + (others.length ? '\n  ' + others.join('\n  ') : 'none.'),
    '',
    mine.length === DPDP_TRIGGER_FUNCTIONS.length
      ? 'All three are in place.'
      : 'Run dpdpInstallTriggers().'
  ].join('\n');
  Logger.log(report);
  return report;
}

// ---------------------------------------------------------------------------
// THE JOBS
// ---------------------------------------------------------------------------

/**
 * Daily. Closes what should have closed.
 *
 * Deliberately deletes nothing that is a record. It expires links, it revokes
 * public Drive shares that predate this change, and it removes sessions that
 * are long dead — none of which is clinical data, and all of which is data
 * kept past its purpose if it is left.
 */
function dpdpDailyMaintenance(e) {
  crescTriggerOnly_(e, 'dpdpDailyMaintenance');
  var out = ['DPDP daily maintenance — ' + dpdp_fmt_(dpdp_now_())];

  try { out.push('  grants:   ' + dpdpExpireDocumentGrants_()); }
  catch (e) { out.push('  grants:   FAILED ' + e.message); }

  // The legacy sweep: files that were published with ANYONE_WITH_LINK before
  // document grants existed. It is a no-op once the backlog is cleared.
  try { out.push('  legacy:   ' + dpdpExpireSharedLinks_(false)); }
  catch (e) { out.push('  legacy:   FAILED ' + e.message); }

  try {
    if (typeof purgeExpiredSessions_ === 'function') {
      out.push('  sessions: ' + purgeExpiredSessions_());
    }
  } catch (e) { out.push('  sessions: FAILED ' + e.message); }

  var report = out.join('\n');
  Logger.log(report);
  return report;
}

/**
 * Weekly. Reads the audit log and tells a person.
 *
 * An anomaly report that goes nowhere is the same as no anomaly report, so
 * this emails the grievance officer — and if none is named it says so loudly
 * in the log, because an unnamed officer is itself a s.13 finding.
 */
function dpdpWeeklyReview(e) {
  crescTriggerOnly_(e, 'dpdpWeeklyReview');
  var scan;
  try { scan = dpdp_anomalyScan_(DPDP_BREACH_CFG.REVIEW_DAYS); }
  catch (e) { scan = { success: false, findings: [], message: 'Scan failed: ' + e.message }; }

  var lines = ['Weekly data-protection review — ' + dpdp_fmt_(dpdp_now_()), '',
               scan.message, ''];
  (scan.findings || []).forEach(function (f, i) {
    lines.push((i + 1) + '. [' + f.severity + '] ' + f.title);
    lines.push('   ' + f.detail);
    lines.push('   WHAT TO DO: ' + f.action);
    lines.push('');
  });

  // Overdue rights requests belong in the same email: both questions are
  // "is anything waiting on us that has a legal clock attached".
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DPDP_CFG.REQUESTS);
    if (sh) {
      var v = dc_sheetValues_(sh), m = dc_headerMap_(sh), late = 0, open = 0;
      for (var i = 1; i < (v ? v.length : 0); i++) {
        if (dpdp_str_(v[i][m['Status']]).toUpperCase() !== 'OPEN') continue;
        open++;
        var due = (typeof cresc_parseDate_ === 'function')
          ? cresc_parseDate_(v[i][m['Due_By']]) : new Date(v[i][m['Due_By']]);
        if (due && !isNaN(due.getTime()) && due.getTime() < Date.now()) late++;
      }
      lines.push('Data principal requests: ' + open + ' open, ' + late + ' OVERDUE.');
      if (late) lines.push('  Sections 11-13 each carry a deadline. Answer them.');
    }
  } catch (e) { /* the review must not fail over one section */ }

  // Unassessed incidents: raised, never decided. That is the state s.8(6)
  // cannot be defended from.
  try {
    var bs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DPDP_BREACH_CFG.SHEET);
    if (bs) {
      var bv = dc_sheetValues_(bs), bm = dc_headerMap_(bs), un = 0;
      for (var j = 1; j < (bv ? bv.length : 0); j++) {
        if (!dpdp_str_(bv[j][bm['Notifiable']])) un++;
      }
      if (un) {
        lines.push('');
        lines.push(un + ' incident(s) in the breach register have never been assessed.');
        lines.push('  Record whether each is notifiable, and why. Both answers need a reason.');
      }
    }
  } catch (e) {}

  var report = lines.join('\n');
  Logger.log(report);

  var officer = dpdp_officer_();
  if (officer && officer.email) {
    try {
      MailApp.sendEmail({
        to: officer.email,
        subject: 'Weekly data-protection review — ' +
                 ((scan.findings || []).length ? (scan.findings.length + ' to look at')
                                               : 'nothing unusual'),
        body: report + '\n\n--\nSent by DPDP_Triggers.gs in the CresRx script. ' +
              'To stop these, run dpdpRemoveTriggers().'
      });
    } catch (e) {
      Logger.log('Could not email the grievance officer: ' + e.message);
    }
  } else {
    Logger.log('NO GRIEVANCE OFFICER EMAIL IS SET, so this review was written to the ' +
               'log and read by nobody. Section 13 requires a named officer: run ' +
               'dpdpSetGrievanceOfficer("name", "email", "phone").');
  }
  return report;
}

/** Monthly. What is past its period — reported, never deleted. */
function dpdpMonthlyRetentionReport(e) {
  crescTriggerOnly_(e, 'dpdpMonthlyRetentionReport');
  var report;
  try { report = dpdpRetentionReport_(); }
  catch (e) { report = 'Retention report failed: ' + e.message; }

  var officer = dpdp_officer_();
  if (officer && officer.email) {
    try {
      MailApp.sendEmail({
        to: officer.email,
        subject: 'Monthly retention report — ' +
                 Utilities.formatDate(new Date(), DPDP_CFG.TZ, 'MMMM yyyy'),
        body: report + '\n\n--\nNothing has been deleted. Erasure is a decision ' +
              'for a person: see docs/RETENTION_SCHEDULE.md.'
      });
    } catch (e) { Logger.log('Could not email the retention report: ' + e.message); }
  }
  return report;
}
