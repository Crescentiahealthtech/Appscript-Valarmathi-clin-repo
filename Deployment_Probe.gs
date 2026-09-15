// ============================================================================
// Deployment_Probe.gs — Crescentia HealthTech
// Measuring the one setting the readiness check could only nag about.
// ----------------------------------------------------------------------------
// THE PROBLEM
//
// dpdpReadinessCheck() reports, and has always reported:
//
//     HIGH · Deployment
//     The manifest asks for access "ANYONE" and executeAs USER_DEPLOYING …
//     now for signed-in callers only, AND ONLY IF THIS DEPLOYMENT WAS
//     PUBLISHED AFTER THE MANIFEST CHANGED.
//     Fix: … open the /exec URL in a private window: if it answers without
//     asking you to sign in, the old ANYONE_ANONYMOUS deployment is still
//     live.
//
// Both halves of that are true, and it is an unsatisfactory finding, because
// appsscript.json describes what the NEXT version would be published with. An
// /exec URL deployed before the manifest changed keeps its old access setting
// for ever, and there is no Apps Script API that lets a script read its own
// live deployment's access mode. So the check reads a file that may not
// describe reality, reports HIGH either way, and asks a human to go and look
// in a private window — which nobody does twice.
//
// It stays HIGH on a clinic that fixed it months ago, which is the exact way
// a readiness report stops being read.
//
// WHAT CAN ACTUALLY BE MEASURED
//
// Session.getActiveUser().getEmail() returns '' when the caller is not
// identified to the script, and an email when they are. That is a direct
// consequence of the live deployment's access mode:
//
//   ANYONE_ANONYMOUS  a visitor who has not signed in reaches doGet, and the
//                     active user is EMPTY.
//   ANYONE            Google requires a sign-in before doGet runs, so the
//                     active user is an email address.
//
// So every page load is EVIDENCE about the live deployment. This file records
// one row per load — a date, whether the caller was identified, and nothing
// else — and turns the finding from a reminder into a fact:
//
//     "The last 340 page loads over 21 days were all signed in, so this
//      deployment requires a Google account."
//
// or
//
//     "12 of the last 340 page loads reached the app with NO identified user.
//      The live deployment is still ANYONE_ANONYMOUS. Most recent: 14-Sep."
//
// WHAT IT DELIBERATELY DOES NOT STORE
//
// Not the email. Not the IP. Not the query string, which carries document
// grant keys. A counter per day and a flag, because the question is about the
// deployment and not about the person: storing who visited the sign-in page
// in order to prove the sign-in page is protected would be its own s.8
// problem. A day with even one anonymous load is what matters, and a count
// says that without naming anybody.
// ============================================================================

var DEP_PROBE = {
  SHEET: 'Deployment_Probe',

  /** How much history the report reads. */
  WINDOW_DAYS: 60,

  /**
   * Written at most once per day per bucket, through the cache — so a clinic
   * with 400 page loads a day costs 2 sheet writes, not 400.
   */
  CACHE_S: 21600      // 6 hours
};

function dep_probeSheet_() {
  return dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), DEP_PROBE.SHEET, [
    'Date', 'Identified_Loads', 'Anonymous_Loads', 'Last_Seen_At', 'Notes'
  ]);
}

/**
 * Records one page load's identification state. NEVER THROWS — this runs on
 * the critical path of doGet, and an unwritable probe must not stop the
 * clinic's application from loading.
 *
 * Called from doGet AFTER the special routes, so the verification pages and
 * the document links — which are anonymous BY DESIGN and prove nothing about
 * the deployment's access mode — are not counted.
 */
function depProbeRecord_() {
  try {
    var identified = false;
    try {
      identified = !!String(Session.getActiveUser().getEmail() || '').trim();
    } catch (e) {
      // A throw here is itself the anonymous case on some configurations.
      identified = false;
    }

    var tz = 'Asia/Kolkata';
    var now = new Date();
    var day = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    var bucket = day + (identified ? ':ID' : ':ANON');

    var cache = CacheService.getScriptCache();
    var seen = cache.get('DEPPROBE_' + bucket);
    var count = seen ? (parseInt(seen, 10) || 0) + 1 : 1;
    cache.put('DEPPROBE_' + bucket, String(count), DEP_PROBE.CACHE_S);

    // An ANONYMOUS load is written IMMEDIATELY, every time. It is the finding,
    // and buffering it behind a cache window would lose it if the cache were
    // evicted before the flush. Identified loads are the uninteresting case
    // and are flushed periodically.
    if (!identified || count === 1 || count % 25 === 0) {
      dep_probeFlush_(day, identified, count);
    }
  } catch (e) {
    try { Logger.log('depProbeRecord_: ' + e.message); } catch (e2) {}
  }
}

/** Writes today's counters into the register, replacing the day's row. */
function dep_probeFlush_(day, identified, count) {
  var sh = dep_probeSheet_();
  var m = dc_headerMap_(sh);
  var data = dc_sheetValues_(sh);
  var col = identified ? 'Identified_Loads' : 'Anonymous_Loads';

  for (var i = 1; i < (data ? data.length : 0); i++) {
    if (String(data[i][m['Date']] || '').indexOf(day) !== 0 &&
        dc_dateKey_(data[i][m['Date']]) !== day) continue;
    var have = parseInt(data[i][m[col]], 10) || 0;
    // The larger of the two: a cache eviction resets the counter, and a
    // register that went backwards would under-report anonymous loads, which
    // is the wrong way for this to be wrong.
    sh.getRange(i + 1, m[col] + 1).setValue(Math.max(have, count));
    sh.getRange(i + 1, m['Last_Seen_At'] + 1).setValue(new Date());
    dc_invalidate_(DEP_PROBE.SHEET);
    return;
  }

  var row = new Array(sh.getLastColumn()).fill('');
  row[m['Date']] = day;
  row[m[col]] = count;
  row[m['Last_Seen_At']] = new Date();
  sh.appendRow(row);
  dc_invalidate_(DEP_PROBE.SHEET);
}

/**
 * What the probe has observed. No authorisation of its own — called by
 * dpdpReadinessCheck, which has already established the caller.
 *
 * @return {{days:number, identified:number, anonymous:number,
 *           lastAnonymous:string, verdict:string}}
 */
function dep_probeSummary_() {
  var out = { days: 0, identified: 0, anonymous: 0, lastAnonymous: '',
              verdict: 'NO_DATA' };
  try {
    var sh = dep_probeSheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    if (!data || data.length < 2) return out;

    var cutoff = Date.now() - DEP_PROBE.WINDOW_DAYS * 86400000;
    for (var i = 1; i < data.length; i++) {
      var d = (typeof cresc_parseDate_ === 'function')
        ? cresc_parseDate_(data[i][m['Date']]) : new Date(data[i][m['Date']]);
      if (!d || isNaN(d.getTime()) || d.getTime() < cutoff) continue;
      out.days++;
      var id = parseInt(data[i][m['Identified_Loads']], 10) || 0;
      var an = parseInt(data[i][m['Anonymous_Loads']], 10) || 0;
      out.identified += id;
      out.anonymous += an;
      if (an > 0) {
        var when = Utilities.formatDate(d, 'Asia/Kolkata', 'dd-MMM-yyyy');
        if (!out.lastAnonymous || when > out.lastAnonymous) out.lastAnonymous = when;
      }
    }

    if (!out.identified && !out.anonymous) out.verdict = 'NO_DATA';
    else if (out.anonymous > 0) out.verdict = 'ANONYMOUS';
    // One signed-in load proves sign-in is POSSIBLE, not that it is required.
    // A week of them with none anonymous is the evidence worth reporting.
    else if (out.days >= 7) out.verdict = 'SIGNED_IN';
    else out.verdict = 'LOOKS_OK_EARLY';
    return out;
  } catch (e) { return out; }
}

/**
 * The Deployment finding, from what has actually been observed.
 *
 * Replaces the unconditional HIGH reminder. The severity now follows the
 * evidence, which is the whole point: a clinic that redeployed and has a
 * month of signed-in loads should not keep reading the same HIGH about a
 * setting it already fixed, and one that never redeployed should read
 * something stronger than "go and check in a private window".
 *
 * @return {{severity:string, text:string, fix:string}}
 */
function depDeploymentFinding() {
  var p = dep_probeSummary_();
  var total = p.identified + p.anonymous;

  if (p.verdict === 'ANONYMOUS') {
    return {
      severity: 'CRITICAL',
      text: p.anonymous + ' of the last ' + total + ' page load(s) over ' +
            p.days + ' day(s) reached this application with NO identified ' +
            'user — most recently on ' + p.lastAnonymous + '. The LIVE ' +
            'deployment is still ANYONE_ANONYMOUS, whatever the manifest ' +
            'says, so anyone with the /exec URL can load the page and call ' +
            'every google.script.run endpoint with the owner\'s full ' +
            'spreadsheet access.',
      fix: 'Deploy > Manage deployments > edit the ACTIVE deployment > ' +
           'New version. The manifest already asks for access "ANYONE"; it ' +
           'takes effect only on a new version. Then treat the period since ' +
           'the first anonymous load as potentially breached and assess it ' +
           'under s.8(6) — dpdpRaiseBreach() opens the register entry.'
    };
  }

  if (p.verdict === 'SIGNED_IN') {
    return {
      severity: 'LOW',
      text: 'All ' + total + ' page load(s) over the last ' + p.days +
            ' day(s) were from an identified Google account, so the live ' +
            'deployment requires a sign-in. executeAs is still ' +
            'USER_DEPLOYING, which means every endpoint runs with the ' +
            'owner\'s spreadsheet access — that is what RBAC.gs guards, and ' +
            'crescRbacCoverage() is where to check it is still complete.',
      fix: 'Nothing to change. Re-read crescRbacCoverage() after adding any ' +
           'new frontend endpoint, since the deployment\'s access is now the ' +
           'outer door and the permission check is the only inner one.'
    };
  }

  if (p.verdict === 'LOOKS_OK_EARLY') {
    return {
      severity: 'MEDIUM',
      text: 'Every one of the ' + total + ' page load(s) so far was signed ' +
            'in, but that is only ' + p.days + ' day(s) of evidence — not ' +
            'yet enough to say the live deployment requires it.',
      fix: 'Leave it a week and re-run this check. In the meantime, opening ' +
           'the /exec URL in a private window answers it immediately: if it ' +
           'loads without asking you to sign in, the old anonymous ' +
           'deployment is still live.'
    };
  }

  return {
    severity: 'HIGH',
    text: 'Nothing is known about the live deployment yet — this check ' +
          'measures whether callers arrive identified, and no page load has ' +
          'been observed since the probe was installed. The manifest asks ' +
          'for access "ANYONE" and executeAs USER_DEPLOYING, but a manifest ' +
          'only takes effect on a NEW version, so an /exec URL published ' +
          'earlier keeps its old setting.',
    fix: 'Open the application once and re-run this check — it will then ' +
         'report what the live deployment actually does. Meanwhile: Deploy > ' +
         'Manage deployments > edit > New version, then open the /exec URL ' +
         'in a private window. If it answers without asking you to sign in, ' +
         'the old ANYONE_ANONYMOUS deployment is still live and any period ' +
         'it was anonymous is potentially a breach (s.8(6)).'
  };
}

/** ADMIN, from the editor. What the probe has seen, in words. */
function RUN_deploymentEvidence() {
  var f = depDeploymentFinding();
  var out = f.severity + ' · Deployment\n\n' + f.text + '\n\nWhat to do:\n' + f.fix;
  Logger.log(out);
  return out;
}
