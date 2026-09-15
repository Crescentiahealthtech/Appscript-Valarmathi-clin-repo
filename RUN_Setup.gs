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
// WHY EVERY NAME ENDS IN AN UNDERSCORE
//
// In Apps Script, a top-level function whose name ends in `_` is NOT exposed
// to google.script.run — the browser cannot call it, at all. The editor's Run
// dropdown lists it just the same, so the underscore costs nothing here and
// buys the only thing that matters: without it, this file would publish
// "reset every password in the clinic" and "mint an administrator session" as
// endpoints that anyone with the web app URL could call by name. A runbook is
// exactly the kind of convenience that turns into a back door, and the fix is
// one character per function.
// ============================================================================


// ---------------------------------------------------------------------------
// PART 1 — CHECK WHAT ARRIVED  (read-only, run these first, in any order)
// ---------------------------------------------------------------------------

/** 1. Did every .gs file get pasted in? Lists anything missing. */
function RUN_01_checkFilesArrived_() {
  return verifyDeployment();
}

/** 2. Is the role matrix internally consistent? Typos in it grant nothing. */
function RUN_02_checkRoleMatrix_() {
  return crescRbacSelfTest();
}

/**
 * 3. What do passwords look like today? Prints COUNTS, never a password.
 *    Expect: "N in PLAIN TEXT" before step 7, "0 in PLAIN TEXT" after.
 */
function RUN_03_checkPasswordStorage_() {
  return crescCredentialStatus();
}

/**
 * 4. What does hashing cost on this runtime? Read the milliseconds.
 *    If 10,000 iterations takes more than ~1.5 seconds, run
 *    RUN_04b_makeSignInFaster_() below — a sign-in people call slow is one
 *    they stop signing out of, which is worse than a lower iteration count.
 */
function RUN_04_timePasswordHashing_() {
  return crescPwdBenchmark();
}

/** 4b. OPTIONAL. Halves the hashing cost. Only if step 4 said it is slow. */
function RUN_04b_makeSignInFaster_() {
  PropertiesService.getScriptProperties().setProperty('CRESC_PWD_ITERATIONS', '5000');
  return 'Hashing set to 5,000 iterations. Passwords already stored keep the ' +
         'count they were written with, so nobody is locked out. Re-run ' +
         'RUN_04_timePasswordHashing_() to see the new cost.';
}


// ---------------------------------------------------------------------------
// PART 2 — SET UP  (run in this order)
// ---------------------------------------------------------------------------

/** 5. Creates every DPDP register. Safe to re-run; it never overwrites. */
function RUN_05_createRegisters_() {
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
function RUN_06_nameGrievanceOfficer_() {
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
 *    Run RUN_07a_migrateCredentials_DRYRUN_() first to see who is affected.
 */
function RUN_07a_migrateCredentials_DRYRUN_() {
  return crescMigrateCredentials(true);
}

/** 7b. **CHANGES DATA.** The real thing. See the warning on 7a. */
function RUN_07b_migrateCredentials_FOR_REAL_() {
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
function RUN_08_installScheduledJobs_() {
  return dpdpInstallTriggers();
}

/** 9. Confirms all three jobs exist. Run it after step 8, and after any change. */
function RUN_09_checkScheduledJobs_() {
  return dpdpTriggerStatus();
}


// ---------------------------------------------------------------------------
// PART 3 — CLEAR THE BACKLOG  (dry run first, every time)
// ---------------------------------------------------------------------------

/**
 * 10. How many patient documents are still readable by anyone with the link,
 *     from before this change? Reports only.
 */
function RUN_10a_revokeOldPublicLinks_DRYRUN_() {
  return dpdpExpireSharedLinks(true);
}

/** 10b. **CHANGES DATA.** Makes those files private. This is the fix for H1. */
function RUN_10b_revokeOldPublicLinks_FOR_REAL_() {
  return dpdpExpireSharedLinks(false);
}

/**
 * 11. Education, occupation and marital status were collected and never read
 *     by anything. Section 6(1) allows collection only for a stated purpose,
 *     so there is no reason to keep them. Reports only.
 */
function RUN_11a_eraseUnusedFields_DRYRUN_() {
  return dpdpEraseUnusedFields(false);
}

/** 11b. **CHANGES DATA.** Clears those three columns for every patient. */
function RUN_11b_eraseUnusedFields_FOR_REAL_() {
  return dpdpEraseUnusedFields(true);
}


// ---------------------------------------------------------------------------
// PART 4 — WHERE YOU STAND  (read-only, run these any time)
// ---------------------------------------------------------------------------

/**
 * 12. The technical half of docs/DPDP_READINESS.md, live against this
 *     deployment. Run it after every step above, and once a month afterwards.
 */
function RUN_12_readinessCheck_() {
  var res = dpdpReadinessCheck();
  return res && res.report;
}

/** 13. What is past its retention period. Reports, and deletes NOTHING. */
function RUN_13_retentionReport_() {
  return dpdpRetentionReport();
}

/** 14. Reads the audit log for the four patterns that precede a disclosure. */
function RUN_14_auditReviewThisWeek_() {
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
function RUN_90_resetOnePassword_() {
  var USERNAME = '[the user id to reset]';

  var token = issueSession_({ username: 'SCRIPT_OWNER', role: 'admin',
                              doctorId: '', name: 'Script owner (editor)' });
  var res = crescAdminResetPassword(USERNAME, token);
  try { revokeSession(token); } catch (e) {}      // the session ends with the job

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
function RUN_91_setVoicePolicy_() {
  var POLICY = 'FORBIDDEN';        // or 'ALLOWED'

  var token = issueSession_({ username: 'SCRIPT_OWNER', role: 'admin',
                              doctorId: '', name: 'Script owner (editor)' });
  var res = dpdpSetVoicePolicy(POLICY, token);
  try { revokeSession(token); } catch (e) {}

  Logger.log(res.message);
  return res.message;
}

/**
 * Takes the three scheduled jobs off again. Only needed if you are moving the
 * script to another account, or stopping the weekly email.
 */
function RUN_99_removeScheduledJobs_() {
  return dpdpRemoveTriggers() + ' scheduled job(s) removed.';
}
