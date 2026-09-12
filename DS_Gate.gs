// ============================================================================
// DS_Gate.gs  —  Crescentia HealthTech / CresRx
// IP Discharge Summary Engine · Phase 8 · the discharge gate
// ----------------------------------------------------------------------------
// WHY A GATE
//   Discovery §1 found TWO places an admission becomes DISCHARGED:
//     processPatientDischarge()  IP_Admissions_Logic.gs:599  (the ward action)
//     settleDischarge()          AccountsIPChargesLogic.gs:299 (the settlement)
//   Gating only one of them would be decorative — the other button still lets a
//   patient leave with no summary. Both are gated, through this one function.
//
// THE CONTRACT WITH EXISTING BEHAVIOUR
//   With DS_BILLING_GATE = OFF, dsx_gateCheck_ returns allow:true before
//   reading anything, so both discharge paths behave exactly as they did.
//   That is the rollback switch: one Script Property, no data change.
//
//   Default is WARN: a discharge with no signed summary is refused ONCE with
//   code DS_NOT_SIGNED, and proceeds when the caller supplies a reason, which
//   is written to both audit trails.
//
//   BLOCK refuses outright — except for DEATH and ABSCONDED, which always
//   behave as WARN. A death certificate cannot wait on a typed summary, and an
//   absconded patient is already gone; blocking there would only teach staff to
//   route around the system.
// ============================================================================

/**
 * @param {string} ipNumber
 * @param {string} overrideReason  supplied on the caller's second attempt
 * @return {{allow:boolean, code:string, message:string, status:string,
 *           dischargeType:string, summaryId:string, overridden:boolean}}
 */
function dsx_gateCheck_(ipNumber, overrideReason) {
  var pass = function (extra) {
    var o = { allow: true, code: '', message: '', status: '', dischargeType: '',
              summaryId: '', overridden: false };
    if (extra) Object.keys(extra).forEach(function (k) { o[k] = extra[k]; });
    return o;
  };

  var cfg;
  try { cfg = dsx_config_(); } catch (e) { return pass(); }
  var mode = cfg.billingGate;
  if (mode === 'OFF') return pass();

  var header = null;
  try { header = dsx_getHeader_(dsx_summaryIdFor_(ipNumber)); }
  catch (e) { return pass(); }   // the module is not installed: never block a discharge

  var status = header ? dsx_upper_(header.Status) : 'NOT STARTED';
  var type = header ? (dsx_upper_(header.Discharge_Type) || 'NORMAL') : '';
  var summaryId = header ? dsx_str_(header.Summary_ID) : '';

  if (status === DSX_STATUS.SIGNED) {
    return pass({ status: status, dischargeType: type, summaryId: summaryId });
  }

  var reason = dsx_str_(overrideReason);
  if (reason) {
    return { allow: true, code: '', message: '', status: status, dischargeType: type,
             summaryId: summaryId, overridden: true, reason: reason };
  }

  var effective = mode;
  if (mode === 'BLOCK' && (type === 'DEATH' || type === 'ABSCONDED')) effective = 'WARN';

  var human = status === 'NOT STARTED'
    ? 'No discharge summary has been started for ' + dsx_ip_(ipNumber) + '.'
    : 'The discharge summary for ' + dsx_ip_(ipNumber) + ' is ' +
      status.replace(/_/g, ' ').toLowerCase() + ', not signed.';

  if (effective === 'BLOCK') {
    return { allow: false, code: 'DS_NOT_SIGNED', status: status, dischargeType: type,
             summaryId: summaryId, overridden: false,
             message: human + ' A doctor must verify and sign it before this patient ' +
                      'can be discharged.' };
  }

  return { allow: false, code: 'DS_NOT_SIGNED', status: status, dischargeType: type,
           summaryId: summaryId, overridden: false,
           message: human + ' Proceed with a reason?' };
}

/**
 * Writes the override to both audit trails: the clinical one, because it is a
 * clinical governance event, and the accounts one, because the accountant who
 * clicked through is the person an auditor will ask.
 */
function dsx_logGateOverride_(ipNumber, gate, user, path) {
  try {
    dsx_audit_({ username: user, role: '' }, 'DS_GATE_OVERRIDE', gate.summaryId || dsx_ip_(ipNumber), {
      ipNumber: dsx_ip_(ipNumber),
      summaryStatus: gate.status,
      dischargeType: gate.dischargeType,
      reason: gate.reason,
      path: path
    });
  } catch (e) {}
  try {
    if (typeof acc_audit_ === 'function') {
      acc_audit_(user || 'UNKNOWN', 'DS_GATE_OVERRIDE', 'DISCHARGE_SUMMARY',
                 gate.summaryId || dsx_ip_(ipNumber),
                 'Summary ' + gate.status, 'Discharged anyway', gate.reason);
    }
  } catch (e) {}
}

/**
 * Logged after a discharge completes. When the billing discharge date differs
 * from the clinical one, that is a soft note in the workflow log — never an
 * edit to signed content.
 */
function dsx_logDischargeCompleted_(ipNumber, user, path) {
  try {
    var header = dsx_getHeader_(dsx_summaryIdFor_(ipNumber));
    if (!header) return;

    var billingAt = new Date();
    var clinicalAt = dsx_toDate_(header.Clinical_Discharge_At);
    var meta = { path: path, billingDischargeAt: billingAt.toISOString() };

    if (clinicalAt) {
      var a = Utilities.formatDate(clinicalAt, dsx_tz_(), 'yyyy-MM-dd');
      var b = Utilities.formatDate(billingAt, dsx_tz_(), 'yyyy-MM-dd');
      if (a !== b) {
        meta.note = 'Billing discharge (' + b + ') is on a different date from the ' +
                    'clinical discharge (' + a + ').';
      }
    }
    dsx_logEvent_(dsx_str_(header.Summary_ID), { username: user, role: '' },
                  'DS_DISCHARGE_COMPLETED', dsx_upper_(header.Status),
                  dsx_upper_(header.Status), 0, '', meta.note || '', meta);
  } catch (e) { /* never fail a completed discharge over a log line */ }
}

/**
 * Read-only status for a discharge screen, so the client can show the gate
 * state before the user clicks anything.
 */
function ds_getGateStatus(token, ipNumber) {
  try {
    dsx_requireRole_(token, ['view', 'viewSigned']);
    var cfg = dsx_config_();
    var gate = dsx_gateCheck_(ipNumber, '');
    return dsx_ok_('', {
      mode: cfg.billingGate,
      allow: gate.allow,
      status: gate.status,
      dischargeType: gate.dischargeType,
      summaryId: gate.summaryId,
      message: gate.message
    });
  } catch (e) {
    return dsx_fromError_(e);
  }
}

// ---------------------------------------------------------------------------
// DASHBOARD
// ---------------------------------------------------------------------------

/**
 * Discharge KPIs. Reads DS_Summaries only — no payloads — so it is cheap
 * enough to sit on a dashboard that refreshes.
 */
function ds_getDashboardKpis(token) {
  try {
    var actor = dsx_requireRole_(token, ['view', 'viewSigned']);
    var sh = dsx_sheet_(DSX_SHEETS.SUMMARIES);
    var out = {
      pendingSignature: 0, returned: 0, inPreparation: 0,
      signedToday: 0, medianTatTodayMin: 0, medianTat7dMin: 0,
      dischargedWithoutSignedSummary: 0
    };
    if (!sh || sh.getLastRow() < 2) return dsx_ok_('', out);

    var map = dsx_headerMap_(sh);
    var width = Math.max(1, sh.getLastColumn());
    var values = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();

    var today = Utilities.formatDate(new Date(), dsx_tz_(), 'yyyy-MM-dd');
    var weekAgo = Date.now() - 7 * 86400000;
    var tatToday = [], tat7d = [];

    for (var i = 0; i < values.length; i++) {
      var status = dsx_upper_(values[i][map['Status']]);
      if (status === DSX_STATUS.PENDING_SIGNATURE) out.pendingSignature++;
      if (status === DSX_STATUS.RETURNED) out.returned++;
      if (status === DSX_STATUS.IN_PREPARATION) out.inPreparation++;

      var signedAt = dsx_toDate_(values[i][map['Signed_At']]);
      var initAt = dsx_toDate_(values[i][map['Initiated_At']]);
      if (!signedAt || !initAt) continue;

      var mins = Math.round((signedAt.getTime() - initAt.getTime()) / 60000);
      if (Utilities.formatDate(signedAt, dsx_tz_(), 'yyyy-MM-dd') === today) {
        out.signedToday++;
        tatToday.push(mins);
      }
      if (signedAt.getTime() >= weekAgo) tat7d.push(mins);
    }

    out.medianTatTodayMin = dsx_median_(tatToday);
    out.medianTat7dMin = dsx_median_(tat7d);
    out.dischargedWithoutSignedSummary = dsx_countGateOverrides_();

    return dsx_ok_('', out);
  } catch (e) {
    return dsx_fromError_(e);
  }
}

function dsx_median_(arr) {
  if (!arr.length) return 0;
  var a = arr.slice().sort(function (x, y) { return x - y; });
  var mid = Math.floor(a.length / 2);
  return a.length % 2 ? a[mid] : Math.round((a[mid - 1] + a[mid]) / 2);
}

/** Gate overrides in the workflow log over the last 7 days. */
function dsx_countGateOverrides_() {
  try {
    var sh = dsx_sheet_(DSX_SHEETS.LOG);
    if (!sh || sh.getLastRow() < 2) return 0;
    var map = dsx_headerMap_(sh);
    var rows = dsx_findRowsByKey_(sh, 'Action', 'DS_GATE_OVERRIDE');
    var weekAgo = Date.now() - 7 * 86400000;
    var n = 0;
    rows.forEach(function (r) {
      var at = dsx_toDate_(sh.getRange(r, map['Timestamp'] + 1).getValue());
      if (at && at.getTime() >= weekAgo) n++;
    });
    return n;
  } catch (e) { return 0; }
}