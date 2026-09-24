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
// WHAT CAN ACTUALLY BE MEASURED — AND WHAT THIS FILE GOT WRONG
//
// The first version of this file reasoned:
//
//   ANYONE_ANONYMOUS  a visitor who has not signed in reaches doGet, and the
//                     active user is EMPTY.
//   ANYONE            Google requires a sign-in before doGet runs, so the
//                     active user is an email address.
//
// The first line is true. THE SECOND IS NOT, and the clinic this runs in is
// exactly the case where it fails.
//
// Session.getActiveUser().getEmail() returns '' whenever the script is not
// permitted to learn who the caller is, and a web app deployed to EXECUTE AS
// ME is the documented case where it is not: the script is authorised by the
// owner, not by the visitor, so the visitor's identity is withheld unless
// they are the owner or share the owner's Google Workspace domain. This
// project is deployed as "executeAs": "USER_DEPLOYING" (appsscript.json) from
// a consumer Gmail account, so EVERY visitor but the owner reads as empty
// whether they signed in or not.
//
// The consequence was a permanent, unfixable CRITICAL. A clinic that had set
// access to "Anyone with a Google account" and redeployed was still told, on
// every page load:
//
//     "17 of the last 21 page loads reached this application with NO
//      identified user … the LIVE deployment is still ANYONE_ANONYMOUS"
//
// — a breach allegation, produced by the deployment being configured
// correctly. That is worse than no check: it is a register saying the clinic
// is leaking patient data when it is not, and it buries the findings next to
// it that are real.
//
// SO THE PROBE NOW RECORDS ONE MORE BIT, AND CLAIMS LESS.
//
// Alongside "was the caller identified", each load records whether the
// identified caller was somebody OTHER than the script's effective user (the
// owner). That single boolean is what separates the two worlds:
//
//   • at least one load identified as a NON-OWNER  →  this deployment can
//     resolve visitor identities, so a load that resolved to nobody really
//     was an unidentified caller. The anonymous finding is then evidence.
//
//   • never  →  the deployment withholds identities from the script, which
//     is what executeAs USER_DEPLOYING does outside a shared Workspace
//     domain. An empty active user is then NOT EVIDENCE OF ANYTHING, and the
//     honest answer is that this cannot be measured from inside the script.
//
// What it stores is still a date, three counters and a last-seen time. No
// email, no IP, no query string.
//
// FOR THE CASE IT CANNOT MEASURE
//
// There is one authoritative test and a person has to run it: open the /exec
// URL in a private window and see whether Google asks for a sign-in.
// dpdpConfirmDeploymentAccess() records that answer with the date and who
// checked, and the finding then reports what was checked and when, instead
// of asking again for ever.
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
  // Other_User_Loads is the column added when the inference was corrected:
  // loads identified as somebody who is NOT the script's effective user. It
  // is what tells a deployment that can resolve visitors apart from one that
  // withholds their identity. dc_ensureSheet_ adds it to an existing
  // register, and the rows written before it simply read 0.
  return dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), DEP_PROBE.SHEET, [
    'Date', 'Identified_Loads', 'Anonymous_Loads', 'Other_User_Loads',
    'Last_Seen_At', 'Notes'
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
    var active = '', effective = '';
    try { active = String(Session.getActiveUser().getEmail() || '').trim(); }
    catch (e) { active = ''; }          // a throw here is the unidentified case
    try { effective = String(Session.getEffectiveUser().getEmail() || '').trim(); }
    catch (e) { effective = ''; }

    var identified = !!active;
    // The bit the whole inference turns on: an identified caller who is NOT
    // the account the script runs as. Only its TRUTH is kept, never the
    // address — the register answers a question about the deployment, and
    // recording who visited the sign-in page in order to prove the sign-in
    // page is protected would be its own section 8 problem.
    var otherUser = identified && effective &&
                    active.toLowerCase() !== effective.toLowerCase();

    var tz = 'Asia/Kolkata';
    var now = new Date();
    var day = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    var kind = identified ? (otherUser ? 'OTHER' : 'ID') : 'ANON';
    var bucket = day + ':' + kind;

    var cache = CacheService.getScriptCache();
    var seen = cache.get('DEPPROBE_' + bucket);
    var count = seen ? (parseInt(seen, 10) || 0) + 1 : 1;
    cache.put('DEPPROBE_' + bucket, String(count), DEP_PROBE.CACHE_S);

    // An ANONYMOUS load is written IMMEDIATELY, every time. It is the finding,
    // and buffering it behind a cache window would lose it if the cache were
    // evicted before the flush. Identified loads are the uninteresting case
    // and are flushed periodically.
    if (!identified || count === 1 || count % 25 === 0) {
      dep_probeFlush_(day, kind, count);
    }
  } catch (e) {
    try { Logger.log('depProbeRecord_: ' + e.message); } catch (e2) {}
  }
}

/** Writes today's counters into the register, replacing the day's row. */
function dep_probeFlush_(day, kind, count) {
  var sh = dep_probeSheet_();
  var m = dc_headerMap_(sh);
  var data = dc_sheetValues_(sh);
  var col = (kind === 'ANON') ? 'Anonymous_Loads'
          : (kind === 'OTHER') ? 'Other_User_Loads' : 'Identified_Loads';
  if (m[col] === undefined) col = (kind === 'ANON') ? 'Anonymous_Loads' : 'Identified_Loads';

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
  var out = { days: 0, identified: 0, anonymous: 0, otherUser: 0,
              lastAnonymous: '', lastAnonymousMs: 0, verdict: 'NO_DATA' };
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
      var ot = (m['Other_User_Loads'] === undefined) ? 0
             : (parseInt(data[i][m['Other_User_Loads']], 10) || 0);
      out.identified += id + ot;      // a non-owner load is an identified one
      out.otherUser += ot;
      out.anonymous += an;
      if (an > 0 && d.getTime() > out.lastAnonymousMs) {
        out.lastAnonymousMs = d.getTime();
        out.lastAnonymous = Utilities.formatDate(d, 'Asia/Kolkata', 'dd-MMM-yyyy');
      }
    }

    if (!out.identified && !out.anonymous) { out.verdict = 'NO_DATA'; return out; }

    if (out.anonymous > 0) {
      // THE DISTINCTION THIS FILE EXISTS TO DRAW.
      //
      // A load that resolved to nobody means "not signed in" only if this
      // deployment resolves anybody at all besides the owner. Under
      // executeAs USER_DEPLOYING on a consumer account it never does, and
      // then an empty active user says nothing whatsoever about the access
      // mode — see the note at the top of this file.
      out.verdict = (out.otherUser > 0) ? 'ANONYMOUS' : 'NOT_MEASURABLE';
      return out;
    }

    // One signed-in load proves sign-in is POSSIBLE, not that it is required.
    // A week of them with none anonymous is the evidence worth reporting.
    out.verdict = (out.days >= 7) ? 'SIGNED_IN' : 'LOOKS_OK_EARLY';
    return out;
  } catch (e) { return out; }
}

// ---------------------------------------------------------------------------
// THE ANSWER A PERSON HAS TO GO AND GET
//
// No Apps Script API reports a deployment's own access mode, and the probe
// above can only measure it on a deployment that resolves visitor identities.
// For every other clinic there is exactly one authoritative test, it takes
// fifteen seconds, and a person has to run it: open the /exec URL in a
// private window.
//
// So the answer is RECORDED rather than asked for again on every report.
// What is stored is the answer, the day it was checked and who checked it —
// an attestation, which is the ordinary instrument for a control that cannot
// be measured automatically. It is re-asked after DEP_ATTEST_DAYS, and the
// finding always prints the date so nobody has to trust an old one blindly.
// ---------------------------------------------------------------------------

var DEP_ATTEST_KEY = 'DEP_ACCESS_ATTESTATION';
var DEP_ATTEST_DAYS = 180;

/** The stored attestation, or null. */
function dep_attestation_() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(DEP_ATTEST_KEY);
    if (!raw) return null;
    var o = JSON.parse(raw);
    if (!o || !o.mode || !o.at) return null;
    var at = new Date(o.at);
    if (isNaN(at.getTime())) return null;
    o.ageDays = Math.floor((Date.now() - at.getTime()) / 86400000);
    o.stale = o.ageDays > DEP_ATTEST_DAYS;
    o.when = Utilities.formatDate(at, 'Asia/Kolkata', 'dd-MMM-yyyy');
    return o;
  } catch (e) { return null; }
}

/**
 * FRONTEND ENTRY. Records what the private-window check actually showed.
 *
 * @param {string} mode  'SIGN_IN_REQUIRED' or 'ANYONE_ANONYMOUS'
 * @param {string} sessionToken
 * @return {{success:boolean, message:string}}
 */
function dpdpConfirmDeploymentAccess(mode, sessionToken) {
  try {
    var actor = crescRequire_(sessionToken, ['dpdp.manage', 'admin.config']);
    var m = String(mode || '').toUpperCase();
    if (m !== 'SIGN_IN_REQUIRED' && m !== 'ANYONE_ANONYMOUS') {
      return { success: false, message: 'Answer either "it asked me to sign in" or ' +
                                        '"it opened without asking".' };
    }
    PropertiesService.getScriptProperties().setProperty(DEP_ATTEST_KEY, JSON.stringify({
      mode: m,
      at: new Date().toISOString(),
      by: String((actor && (actor.displayName || actor.username)) || '')
    }));

    // An admission that the live deployment is open is a section 8(6) matter,
    // not a setting. It opens a breach entry rather than sitting in a
    // property where only this report would ever see it.
    if (m === 'ANYONE_ANONYMOUS' && typeof dpdpRaiseBreach === 'function') {
      try {
        dpdpRaiseBreach({
          category: 'UNAUTHORISED_ACCESS',
          source: 'DEPLOYMENT_CHECK',
          severity: 'UNASSESSED',
          summary: 'The live web app deployment was confirmed by ' +
                   ((actor && (actor.displayName || actor.username)) || 'an administrator') +
                   ' to open without asking for a Google sign-in, so anyone holding ' +
                   'the /exec URL could reach every endpoint with the owner\u2019s ' +
                   'spreadsheet access. Any period it was live in that state is ' +
                   'potentially a personal data breach and needs assessing under s.8(6).'
        }, sessionToken);
      } catch (e) { /* the attestation is recorded either way */ }
    }

    return { success: true, message: (m === 'SIGN_IN_REQUIRED')
      ? 'Recorded: the live deployment asks for a Google sign-in. This will be ' +
        're-checked in ' + DEP_ATTEST_DAYS + ' days.'
      : 'Recorded: the live deployment opens without a sign-in. Publish a new ' +
        'version now, then record the check again.' };
  } catch (e) {
    return { success: false, message: String(e.message || e).replace('FORBIDDEN: ', '') };
  }
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
  crescEditorOnly_('depDeploymentFinding', ['dpdp.manage', 'admin.config', 'admin.audit']);
  var p = dep_probeSummary_();
  var att = dep_attestation_();
  var total = p.identified + p.anonymous;

  // ---- an ADMITTED open deployment outranks everything below -------------
  if (att && att.mode === 'ANYONE_ANONYMOUS' && !att.stale) {
    return {
      severity: 'CRITICAL',
      text: 'The live deployment was checked on ' + att.when +
            (att.by ? ' by ' + att.by : '') + ' and opened WITHOUT asking for a ' +
            'Google sign-in. Anyone holding the /exec URL can load the page and ' +
            'call every google.script.run endpoint with the owner\u2019s full ' +
            'spreadsheet access.',
      fix: 'Deploy > Manage deployments > edit the ACTIVE deployment > New ' +
           'version, with access set to "Anyone with a Google account". The ' +
           'manifest already asks for it; it takes effect only on a new version. ' +
           'Then check the /exec URL in a private window again and record the ' +
           'result, and assess the period it was open under s.8(6).'
    };
  }

  // ---- measured: this deployment DOES resolve visitors, and some were not -
  if (p.verdict === 'ANONYMOUS') {
    return {
      severity: 'CRITICAL',
      text: p.anonymous + ' of the last ' + total + ' page load(s) over ' +
            p.days + ' day(s) reached this application with NO identified ' +
            'user — most recently on ' + p.lastAnonymous + '. This deployment ' +
            'does resolve who its visitors are (' + p.otherUser + ' load(s) ' +
            'identified as somebody other than the owner), so those were ' +
            'genuinely unidentified callers and the live deployment is still ' +
            'ANYONE_ANONYMOUS.',
      fix: 'Deploy > Manage deployments > edit the ACTIVE deployment > ' +
           'New version. The manifest already asks for access "ANYONE"; it ' +
           'takes effect only on a new version. Then treat the period since ' +
           'the first anonymous load as potentially breached and assess it ' +
           'under s.8(6) — dpdpRaiseBreach() opens the register entry.'
    };
  }

  // ---- the case this file used to report as a breach ---------------------
  if (p.verdict === 'NOT_MEASURABLE') {
    if (att && att.mode === 'SIGN_IN_REQUIRED' && !att.stale) {
      return {
        severity: 'LOW',
        text: 'This deployment runs as "execute as me", so Apps Script does not ' +
              'tell the script who its visitors are and page loads cannot show ' +
              'whether a sign-in was required — ' + p.anonymous + ' of ' + total +
              ' load(s) resolved to nobody, which is the expected reading either ' +
              'way. The private-window check was done on ' + att.when +
              (att.by ? ' by ' + att.by : '') + ' and the deployment DID ask for ' +
              'a Google sign-in. executeAs is still USER_DEPLOYING, so every ' +
              'endpoint runs with the owner\u2019s spreadsheet access — that is ' +
              'what RBAC.gs guards.',
        fix: 'Nothing to change. Re-check after any new deployment, and in any ' +
             'case within ' + Math.max(0, DEP_ATTEST_DAYS - att.ageDays) + ' day(s), ' +
             'when this attestation lapses. Run crescRbacCoverage() after adding ' +
             'any new frontend endpoint.'
      };
    }
    return {
      severity: 'MEDIUM',
      text: 'Whether the live deployment requires a Google sign-in CANNOT BE ' +
            'MEASURED from inside this script. It runs as "execute as me" ' +
            '(executeAs USER_DEPLOYING), and Apps Script withholds a visitor\u2019s ' +
            'identity from a script authorised by its owner unless they share a ' +
            'Google Workspace domain — so all ' + p.anonymous + ' of the ' + total +
            ' load(s) over ' + p.days + ' day(s) that resolved to nobody read ' +
            'exactly the same whether those people signed in or not. ' +
            'This finding previously called that a breach; it was not evidence ' +
            'of one.' +
            (att && att.stale
              ? ' The last check, on ' + att.when + ', is more than ' +
                DEP_ATTEST_DAYS + ' days old.'
              : ''),
      fix: 'Open the /exec URL in a private window. If Google asks you to sign ' +
           'in, the deployment is correct — record that on the Data protection ' +
           'screen (Readiness > "Record the deployment check") and this finding ' +
           'settles. If it opens straight into the app, publish a new version ' +
           'with access "Anyone with a Google account" and assess the period it ' +
           'was open under s.8(6).'
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

  if (att && att.mode === 'SIGN_IN_REQUIRED' && !att.stale) {
    return {
      severity: 'LOW',
      text: 'No page load has been observed since the probe was installed, so ' +
            'there is nothing measured — but the private-window check was done ' +
            'on ' + att.when + (att.by ? ' by ' + att.by : '') +
            ' and the deployment asked for a Google sign-in.',
      fix: 'Nothing to change. Re-check after any new deployment.'
    };
  }

  return {
    severity: 'MEDIUM',
    text: 'Nothing is known about the live deployment yet — no page load has ' +
          'been observed since the probe was installed, and no private-window ' +
          'check has been recorded. The manifest asks for access "ANYONE" and ' +
          'executeAs USER_DEPLOYING, but a manifest only takes effect on a NEW ' +
          'version, so an /exec URL published earlier keeps its old setting.',
    fix: 'Open the /exec URL in a private window and record what happened on ' +
         'the Data protection screen (Readiness > "Record the deployment ' +
         'check"). If it answers without asking you to sign in, publish a new ' +
         'version and assess the period it was open under s.8(6).'
  };
}

/** ADMIN, from the editor. What the probe has seen, in words. */
function RUN_deploymentEvidence() {
  crescEditorOnly_('RUN_deploymentEvidence');
  var f = depDeploymentFinding();
  var out = f.severity + ' · Deployment\n\n' + f.text + '\n\nWhat to do:\n' + f.fix;
  Logger.log(out);
  return out;
}
