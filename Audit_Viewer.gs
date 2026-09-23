// ============================================================================
// Audit_Viewer.gs — Crescentia HealthTech / CresRx
// One place to read who did what, filtered by user, patient or action.
// ----------------------------------------------------------------------------
// The application keeps three audit trails, written by different modules at
// different times:
//
//   Audit_Log            clinical, sign-in, consent, IP, EMR    (logAudit_)
//   Audit_Event_Ledger   money: bills, refunds, settlements     (acc_audit_)
//   LAB_AUDIT_LOG        lab orders, results, verification      (labAudit_)
//
// The only screen onto any of them was the Sign-in Audit, which reads the
// sign-in rows of the first and nothing else. "Who changed this patient's
// bill", "who verified this report", "what did this user do on Tuesday" meant
// opening three sheets and filtering each by hand.
//
// This reads the tail of each (the sheets are append-only and the question is
// nearly always about the recent past), maps them onto one row shape, and
// merges newest first. Export is done in the browser from what is returned,
// so the file holds exactly what the screen showed.
// ============================================================================

/** The trails, and how each names its columns. */
var AUDV_SOURCES = [
  { key: 'app', label: 'Clinical & sign-in', sheet: 'Audit_Log',
    ts: ['Timestamp'], user: ['Actor_Username'], role: ['Actor_Role'],
    action: ['Event'], etype: ['Entity_Type'], eid: ['Entity_ID'],
    detail: ['Details_JSON'] },
  { key: 'finance', label: 'Billing & accounts', sheet: 'Audit_Event_Ledger',
    ts: ['Timestamp'], user: ['User_Name'], role: [],
    action: ['Action_Type'], etype: ['Module'], eid: ['Reference_ID'],
    detail: ['Old_Value', 'New_Value', 'Reason_Remarks'] },
  { key: 'lab', label: 'Laboratory', sheet: 'LAB_AUDIT_LOG',
    ts: ['Timestamp'], user: ['PerformedBy', 'UserName'], role: [],
    action: ['Action'], etype: ['EntityType'], eid: ['EntityID'],
    detail: ['OldValue', 'NewValue'] },
  { key: 'schedule', label: 'Doctor schedule', sheet: 'Audit_Logs',
    ts: ['Timestamp'], user: ['User'], role: [],
    action: ['Action'], etype: ['Module'], eid: ['Target Dates'],
    detail: [] }
];

/** Rows read per sheet at most, however wide the window. */
var AUDV_SCAN_CAP = 25000;

/** A timestamp cell as a Date, or null. Sheets hold both Dates and strings. */
function audv_when_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  var s = String(v || '').trim();
  if (!s) return null;
  var m = /^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?/.exec(s);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0));
  m = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:[ ,]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/.exec(s);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

/** Column index for the first of several header names present, or -1. */
function audv_col_(map, names) {
  for (var i = 0; i < names.length; i++) {
    if (map.hasOwnProperty(names[i])) return map[names[i]];
  }
  return -1;
}

/**
 * A details cell as one readable line. Details_JSON is flattened to
 * "key: value · key: value" so the export is legible in a spreadsheet.
 */
function audv_detail_(parts) {
  var out = [];
  parts.forEach(function (p) {
    var s = String(p.v === null || p.v === undefined ? '' : p.v).trim();
    if (!s || s === '{}') return;
    if (p.json && /^[\{\[]/.test(s)) {
      try {
        var o = JSON.parse(s);
        if (o && typeof o === 'object' && !Array.isArray(o)) {
          var bits = [];
          Object.keys(o).forEach(function (k) {
            if (k === 'at') return;      // duplicates the Timestamp column
            var val = o[k];
            if (val === '' || val === null || val === undefined) return;
            bits.push(k + ': ' + (typeof val === 'object' ? JSON.stringify(val) : String(val)));
          });
          s = bits.join(' · ');
        }
      } catch (e) {}
    }
    if (s) out.push(p.label ? p.label + ': ' + s : s);
  });
  var line = out.join(' · ');
  return line.length > 600 ? line.slice(0, 597) + '…' : line;
}

/**
 * FRONTEND ENTRY. The audit trails, merged, newest first.
 *
 * @param {string} token
 * @param {{days?:number, from?:string, to?:string, user?:string,
 *          patient?:string, action?:string, sources?:Array<string>,
 *          limit?:number}} [opts]
 *   user     part of a username, any case
 *   patient  a patient ID, IP number, bill or report number — matched
 *            against the record the row is about and its details
 *   action   part of the event name ("DISCHARGE", "REFUND", "LOGIN")
 *   from/to  yyyy-MM-dd, inclusive; override days
 */
function crescGetAuditLog(token, opts) {
  try {
    crescRequire_(token, 'admin.audit');
    var o = opts || {};
    var limit = Math.min(Math.max(parseInt(o.limit, 10) || 500, 1), 3000);
    var tz = Session.getScriptTimeZone();

    var until = null, since;
    if (o.from) {
      since = audv_when_(String(o.from) + ' 00:00:00');
      if (o.to) until = audv_when_(String(o.to) + ' 23:59:59');
    }
    if (!since) {
      var days = Math.min(Math.max(parseInt(o.days, 10) || 7, 1), 366);
      since = new Date(Date.now() - days * 86400000);
    }
    var wantUser = String(o.user || '').trim().toUpperCase();
    var wantPt = String(o.patient || '').trim().toUpperCase();
    var wantAct = String(o.action || '').trim().toUpperCase();
    var wantSrc = Array.isArray(o.sources) && o.sources.length ? o.sources : null;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rows = [], perSource = {}, actions = {}, truncated = false;

    AUDV_SOURCES.forEach(function (src) {
      if (wantSrc && wantSrc.indexOf(src.key) === -1) return;
      var sh = ss.getSheetByName(src.sheet);
      perSource[src.key] = { label: src.label, rows: 0, present: !!sh };
      if (!sh || sh.getLastRow() < 2) return;

      var cols = sh.getLastColumn();
      var map = (src.key === 'lab' && typeof labHeaderMap_ === 'function')
        ? labHeaderMap_(sh) : {};
      if (src.key !== 'lab') {
        sh.getRange(1, 1, 1, cols).getValues()[0].forEach(function (h, i) {
          if (h !== '' && h !== null) map[String(h).trim()] = i;
        });
      }
      var c = {
        ts: audv_col_(map, src.ts), user: audv_col_(map, src.user), role: audv_col_(map, src.role),
        action: audv_col_(map, src.action), etype: audv_col_(map, src.etype), eid: audv_col_(map, src.eid)
      };
      if (c.ts < 0) return;
      var detailCols = src.detail.map(function (n) {
        return { i: audv_col_(map, [n]), label: src.detail.length > 1 ? n.replace(/_/g, ' ') : '',
                 json: /JSON/.test(n) };
      }).filter(function (d) { return d.i >= 0; });

      var cursor = sh.getLastRow(), scanned = 0, block = 600, taken = 0;
      // Out-of-order rows are tolerated: a few late writes do not end the scan,
      // a run of 50 rows older than the window does.
      var olderRun = 0;
      while (cursor > 1 && scanned < AUDV_SCAN_CAP && taken < limit) {
        var from = Math.max(2, cursor - block + 1);
        var data = sh.getRange(from, 1, cursor - from + 1, cols).getValues();
        for (var i = data.length - 1; i >= 0 && taken < limit; i--) {
          scanned++;
          var r = data[i];
          var when = audv_when_(r[c.ts]);
          if (!when) continue;
          if (when < since) { if (++olderRun >= 50) { cursor = 1; break; } continue; }
          olderRun = 0;
          if (until && when > until) continue;

          var user = c.user >= 0 ? String(r[c.user] || '') : '';
          var action = c.action >= 0 ? String(r[c.action] || '') : '';
          if (wantUser && user.toUpperCase().indexOf(wantUser) === -1) continue;
          if (wantAct && action.toUpperCase().indexOf(wantAct) === -1) continue;

          var eid = c.eid >= 0 ? String(r[c.eid] || '') : '';
          var detail = audv_detail_(detailCols.map(function (d) {
            return { v: r[d.i], label: d.label, json: d.json };
          }));
          if (wantPt && (eid + ' ' + detail).toUpperCase().indexOf(wantPt) === -1) continue;

          actions[action] = (actions[action] || 0) + 1;
          taken++;
          rows.push({
            ms: when.getTime(),
            at: Utilities.formatDate(when, tz, 'dd-MMM-yyyy HH:mm:ss'),
            source: src.key,
            user: user,
            role: c.role >= 0 ? String(r[c.role] || '') : '',
            action: action,
            entityType: c.etype >= 0 ? String(r[c.etype] || '') : '',
            entityId: eid,
            detail: detail
          });
        }
        if (cursor > 1) cursor = from - 1;
      }
      if (taken >= limit) truncated = true;
      perSource[src.key].rows = taken;
    });

    rows.sort(function (a, b) { return b.ms - a.ms; });
    if (rows.length > limit) { rows = rows.slice(0, limit); truncated = true; }

    return {
      success: true,
      rows: rows,
      count: rows.length,
      truncated: truncated,
      limit: limit,
      sources: perSource,
      actions: Object.keys(actions).sort(),
      since: Utilities.formatDate(since, tz, 'dd-MMM-yyyy'),
      message: ''
    };
  } catch (err) {
    return { success: false, rows: [], count: 0,
             message: String((err && err.message) || err).replace('FORBIDDEN: ', '') };
  }
}
