// ============================================================================
// RUN_Setup.gs — Crescentia HealthTech
// The deployment runbook, as functions you can actually press Run on.
// ----------------------------------------------------------------------------
// WHY THIS FILE EXISTS
//
// Two things about the Apps Script editor make a runbook written as
// "now call dpdpSetup()" harder to follow than it looks:
//
//   1. The function dropdown next to Run only lists functions in the file
//      that is CURRENTLY OPEN. So a thirteen-step runbook spread across six
//      files means opening six files and hunting through each one.
//
//   2. The dropdown CANNOT PASS ARGUMENTS. Every function it runs is called
//      with none. That is a real hazard here, because
//
//          crescMigrateCredentials(true)   is the DRY RUN
//          crescMigrateCredentials()       RESETS EVERY PASSWORD IN THE CLINIC
//
//      and the dropdown can only ever do the second one. The same trap is on
//      dpdpExpireSharedLinks(). Read that twice before pressing anything.
//
// So this file is the runbook itself: one function per step, in order, taking
// no arguments, named so the dropdown reads as a list of instructions. Open
// this file, work down the list.
//
// NOTHING NEW HAPPENS HERE. Every function below is a one-line call to
// something that already existed; this file adds no logic, so there is nothing
// in it that can behave differently from what the runbook describes.
//
// STEPS MARKED **CHANGES DATA** CANNOT BE UNDONE. They are named so you have
// to mean it.
//
// WHY EVERY FUNCTION STARTS WITH crescEditorOnly_()
//
// These used to end in `_`, which keeps a function away from google.script.run
// — and also out of the editor's Run dropdown, which made the whole runbook
// impossible to select. So the names are plain again, and each one's first
// line is crescEditorOnly_(): it lets the call through only when the person
// running it is the account the script belongs to (the editor), and refuses a
// browser. Without that line this file would publish "reset every password in
// the clinic" and "mint an administrator session" as endpoints anyone with the
// web app URL could call by name.
// ============================================================================


// ---------------------------------------------------------------------------
// PART 1 — CHECK WHAT ARRIVED  (read-only, run these first, in any order)
// ---------------------------------------------------------------------------

/**
 * 0. GRANT THE PERMISSIONS THE APP NEEDS.  Run this first, and again after
 *    any update that adds a new Google service.
 *
 * The web app runs as YOU (executeAs: USER_DEPLOYING), so it can only use a
 * Google service you have approved for this script. A new service in the code
 * is not approved until somebody runs a function in the editor and accepts
 * the permission screen. Until then the web app fails with
 * "You do not have permission to call UrlFetchApp.fetch" — which is exactly
 * what Google sign-in reports when this step has been skipped.
 *
 * Press Run, choose your account, and on "Google hasn't verified this app"
 * choose Advanced → Go to (project) → Allow. That screen appears because
 * this is your own script, not a published app; it is expected.
 *
 * Each service is touched once, harmlessly, so the report says which work:
 *   UrlFetchApp  Google sign-in (checks the token with Google), WhatsApp API
 *   DriveApp     nightly backups
 *   GmailApp     password-reset email, reminders by email, the day summary
 *   ScriptApp    the scheduled jobs
 */
function RUN_00_authorizeServices() {
  crescEditorOnly_('RUN_00_authorizeServices');
  var out = ['PERMISSIONS', ''];
  var check = function (label, fn) {
    try { out.push('PASS  ' + label + ' — ' + fn()); }
    catch (e) { out.push('FAIL  ' + label + ' — ' + e.message); }
  };
  check('UrlFetchApp (Google sign-in)', function () {
    return 'reached Google (HTTP ' + UrlFetchApp.fetch('https://www.google.com/generate_204',
      { muteHttpExceptions: true }).getResponseCode() + ')';
  });
  check('DriveApp (backups)', function () { DriveApp.getRootFolder().getName(); return 'can read your Drive'; });
  check('GmailApp (emails)', function () { GmailApp.getAliases(); return 'can send as ' + Session.getEffectiveUser().getEmail(); });
  check('ScriptApp (scheduled jobs)', function () { return ScriptApp.getProjectTriggers().length + ' trigger(s) installed'; });
  out.push('');
  if (out.join('\n').indexOf('FAIL') === -1) {
    out.push('Every line PASS: Google sign-in and the other services now work in the');
    out.push('web app. No new deployment is needed for permissions alone.');
  } else {
    out.push('A FAIL above means that permission is not in this project\'s manifest, so');
    out.push('Google never offers it. Fix it once:');
    out.push('  1. Project Settings (gear) -> tick "Show appsscript.json manifest file".');
    out.push('  2. Open appsscript.json and replace it with the repository\'s copy, whose');
    out.push('     "oauthScopes" list includes');
    out.push('     https://www.googleapis.com/auth/script.external_request');
    out.push('  3. Save, run RUN_00_authorizeServices again, and Allow on the new screen.');
    out.push('  4. Deploy -> Manage deployments -> Edit -> Version: New version -> Deploy.');
  }
  var report = out.join('\n');
  Logger.log(report);
  return report;
}

/** 1. Did every .gs file get pasted in? Lists anything missing. */
function RUN_01_checkFilesArrived() {
  crescEditorOnly_('RUN_01_checkFilesArrived');
  return verifyDeployment();
}

/** 2. Is the role matrix internally consistent? Typos in it grant nothing. */
function RUN_02_checkRoleMatrix() {
  crescEditorOnly_('RUN_02_checkRoleMatrix');
  return crescRbacSelfTest();
}

/**
 * 3. What do passwords look like today? Prints COUNTS, never a password.
 *    Expect: "N in PLAIN TEXT" before step 7, "0 in PLAIN TEXT" after.
 */
function RUN_03_checkPasswordStorage() {
  crescEditorOnly_('RUN_03_checkPasswordStorage');
  return crescCredentialStatus();
}

/**
 * 4. What does hashing cost on this runtime? Read the milliseconds.
 *    If 10,000 iterations takes more than ~1.5 seconds, run
 *    RUN_04b_makeSignInFaster() below — a sign-in people call slow is one
 *    they stop signing out of, which is worse than a lower iteration count.
 */
function RUN_04_timePasswordHashing() {
  crescEditorOnly_('RUN_04_timePasswordHashing');
  return crescPwdBenchmark();
}

/** 4b. OPTIONAL. Halves the hashing cost. Only if step 4 said it is slow. */
function RUN_04b_makeSignInFaster() {
  crescEditorOnly_('RUN_04b_makeSignInFaster');
  PropertiesService.getScriptProperties().setProperty('CRESC_PWD_ITERATIONS', '5000');
  return 'Hashing set to 5,000 iterations. Passwords already stored keep the ' +
         'count they were written with, so nobody is locked out. Re-run ' +
         'RUN_04_timePasswordHashing() to see the new cost.';
}


// ---------------------------------------------------------------------------
// PART 2 — SET UP  (run in this order)
// ---------------------------------------------------------------------------

/** 5. Creates every DPDP register. Safe to re-run; it never overwrites. */
function RUN_05_createRegisters() {
  crescEditorOnly_('RUN_05_createRegisters');
  var out = [dpdpSetup()];
  // The antenatal and immunisation registers. Idempotent, and the patient
  // portal's pregnancy and vaccination panels stay hidden until these exist
  // and somebody records something in them — see Maternal_Child.gs.
  try { out.push(mcSetup()); }
  catch (e) { out.push('Maternal_Child.gs is not in this project: ' + e.message); }
  return out.join('\n\n');
}

/**
 * 6. Names the person a patient complains to — section 13 requires one.
 *
 *    EDIT THE THREE VALUES BELOW before you run this. They appear on the
 *    notice the patient is given, so they must be a real person who will
 *    actually answer.
 */
function RUN_06_nameGrievanceOfficer() {
  crescEditorOnly_('RUN_06_nameGrievanceOfficer');
  return dpdpSetGrievanceOfficer(
    'Dr. [full name]',            // who
    '[officer]@[clinic].in',      // email — the weekly review is sent here
    '+91 [phone]'                 // phone
  );
}

/**
 * 7. **CHANGES DATA — READ THIS FIRST.**
 *
 *    Issues a fresh random password to every staff and patient account whose
 *    password is still stored in plain text, and forces each of them to
 *    change it at first sign-in.
 *
 *    THE LOG IT PRINTS IS THE ONLY COPY. Copy it somewhere safe before you
 *    close the tab. Nothing — not this script, not the spreadsheet, not
 *    Google — can show you those passwords again.
 *
 *    Until you run this, NOBODY CAN SIGN IN, because a plain-text password is
 *    refused. That is deliberate: it is what a forced reset means. So do not
 *    run the deployment at the start of a clinic session.
 *
 *    Run RUN_07a_migrateCredentials_DRYRUN() first to see who is affected.
 */
function RUN_07a_migrateCredentials_DRYRUN() {
  crescEditorOnly_('RUN_07a_migrateCredentials_DRYRUN');
  return crescMigrateCredentials(true);
}

/** 7b. **CHANGES DATA.** The real thing. See the warning on 7a. */
function RUN_07b_migrateCredentials_FOR_REAL() {
  crescEditorOnly_('RUN_07b_migrateCredentials_FOR_REAL');
  return crescMigrateCredentials();
}

/**
 * 8. Installs the three scheduled jobs: the nightly expiry sweep, the Monday
 *    audit review emailed to the grievance officer, and the monthly retention
 *    report. Without these, nothing expires and nobody reads the audit log.
 *
 *    They run as WHOEVER PRESSES RUN HERE. If that person leaves the clinic,
 *    somebody else has to run this again.
 */
function RUN_08_installScheduledJobs() {
  crescEditorOnly_('RUN_08_installScheduledJobs');
  return dpdpInstallTriggers();
}

/** 9. Confirms all three jobs exist. Run it after step 8, and after any change. */
function RUN_09_checkScheduledJobs() {
  crescEditorOnly_('RUN_09_checkScheduledJobs');
  return dpdpTriggerStatus();
}


// ---------------------------------------------------------------------------
// PART 3 — CLEAR THE BACKLOG  (dry run first, every time)
// ---------------------------------------------------------------------------

/**
 * 10. How many patient documents are still readable by anyone with the link,
 *     from before this change? Reports only.
 */
function RUN_10a_revokeOldPublicLinks_DRYRUN() {
  crescEditorOnly_('RUN_10a_revokeOldPublicLinks_DRYRUN');
  return dpdpExpireSharedLinks(true);
}

/** 10b. **CHANGES DATA.** Makes those files private. This is the fix for H1. */
function RUN_10b_revokeOldPublicLinks_FOR_REAL() {
  crescEditorOnly_('RUN_10b_revokeOldPublicLinks_FOR_REAL');
  return dpdpExpireSharedLinks(false);
}

/**
 * 11. Education, occupation and marital status were collected and never read
 *     by anything. Section 6(1) allows collection only for a stated purpose,
 *     so there is no reason to keep them. Reports only.
 */
function RUN_11a_eraseUnusedFields_DRYRUN() {
  crescEditorOnly_('RUN_11a_eraseUnusedFields_DRYRUN');
  return dpdpEraseUnusedFields(false);
}

/** 11b. **CHANGES DATA.** Clears those three columns for every patient. */
function RUN_11b_eraseUnusedFields_FOR_REAL() {
  crescEditorOnly_('RUN_11b_eraseUnusedFields_FOR_REAL');
  return dpdpEraseUnusedFields(true);
}


/**
 * 11c. The OP_Encounters vitals headers (D-J) read "Sys_BP" seven times on
 *      the clinic's sheet. The data underneath is right; this renames the
 *      headers. Reports only.
 */
function RUN_11c_fixOPEncounterHeaders_DRYRUN() {
  crescEditorOnly_('RUN_11c_fixOPEncounterHeaders_DRYRUN');
  return repairOPEncounterHeaders(true);
}

/** 11d. **CHANGES DATA** (row 1 of OP_Encounters only). */
function RUN_11d_fixOPEncounterHeaders_FOR_REAL() {
  crescEditorOnly_('RUN_11d_fixOPEncounterHeaders_FOR_REAL');
  return repairOPEncounterHeaders(false);
}


// ---------------------------------------------------------------------------
// PART 4 — WHERE YOU STAND  (read-only, run these any time)
// ---------------------------------------------------------------------------

/**
 * 12. The technical half of docs/DPDP_READINESS.md, live against this
 *     deployment. Run it after every step above, and once a month afterwards.
 */
function RUN_12_readinessCheck() {
  crescEditorOnly_('RUN_12_readinessCheck');
  var res = dpdpReadinessCheck();
  return res && res.report;
}

/** 13. What is past its retention period. Reports, and deletes NOTHING. */
function RUN_13_retentionReport() {
  crescEditorOnly_('RUN_13_retentionReport');
  return dpdpRetentionReport();
}

/** 14. Reads the audit log for the four patterns that precede a disclosure. */
function RUN_14_auditReviewThisWeek() {
  crescEditorOnly_('RUN_14_auditReviewThisWeek');
  var res = dpdp_anomalyScan_(7);
  var out = [res.message, ''];
  (res.findings || []).forEach(function (f, i) {
    out.push((i + 1) + '. [' + f.severity + '] ' + f.title);
    out.push('   ' + f.detail);
    out.push('   WHAT TO DO: ' + f.action);
    out.push('');
  });
  var report = out.join('\n');
  Logger.log(report);
  return report;
}


// ---------------------------------------------------------------------------
// PART 5 — THINGS YOU WILL NEED ONE DAY
// ---------------------------------------------------------------------------

/**
 * Resets ONE person's password, when they are locked out and everybody else
 * is fine. Edit the username below, then run.
 *
 * WHY IT LOOKS ODD. crescAdminResetPassword() is a normal guarded endpoint —
 * it asks who is calling, because from the application it must. The editor has
 * no session, so this mints one for the person who is already the owner of
 * this script and this spreadsheet. That grants nothing they do not already
 * have; it is a formality to satisfy the same door everyone else comes
 * through, not a way round it.
 */
function RUN_90_resetOnePassword() {
  crescEditorOnly_('RUN_90_resetOnePassword');
  var USERNAME = '[the user id to reset]';

  var token = issueSession_({ username: 'SCRIPT_OWNER', role: 'admin',
                              doctorId: '', name: 'Script owner (editor)' });
  var res = crescAdminResetPassword(USERNAME, token);
  try { revokeSession_(token); } catch (e) {}      // the session ends with the job

  Logger.log(res.message);
  return res.message;
}

/**
 * Says whether dictated clinical audio may leave the building (finding M3).
 * Change 'FORBIDDEN' to 'ALLOWED' if the clinic decides the other way, then
 * say which in the Consent Notice — docs/CONSENT_NOTICE.md drafts both.
 *
 * This can also be set from the application: Privacy > Posture.
 */
function RUN_91_setVoicePolicy() {
  crescEditorOnly_('RUN_91_setVoicePolicy');
  var POLICY = 'FORBIDDEN';        // or 'ALLOWED'

  var token = issueSession_({ username: 'SCRIPT_OWNER', role: 'admin',
                              doctorId: '', name: 'Script owner (editor)' });
  var res = dpdpSetVoicePolicy(POLICY, token);
  try { revokeSession_(token); } catch (e) {}

  Logger.log(res.message);
  return res.message;
}

/**
 * Takes the three scheduled jobs off again. Only needed if you are moving the
 * script to another account, or stopping the weekly email.
 */
function RUN_99_removeScheduledJobs() {
  crescEditorOnly_('RUN_99_removeScheduledJobs');
  return dpdpRemoveTriggers() + ' scheduled job(s) removed.';
}
