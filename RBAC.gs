// ============================================================================
// RBAC.gs — Crescentia HealthTech / CresRx
// One permission vocabulary, one role matrix, one server-side guard.
// ----------------------------------------------------------------------------
// WHAT "RBAC" MEANT IN THIS PROJECT BEFORE THIS FILE
//
// Real role-based access control is three things. The application had one of
// them, and it was the one that does not protect anything.
//
//   1. AUTHENTICATION — who are you.
//      verifyLogin() in AuthLogin.gs. Works.
//
//   2. AUTHORISATION — may you do this. THE SERVER DECIDES.
//      Almost entirely absent. Of roughly 349 functions reachable from
//      google.script.run, about 20 files carry any session check at all, and
//      most of those carry one or two. DS_Workflow.gs is the exception and
//      the model: dsx_requireRole_() is the first statement of every public
//      function in it. Nothing else copied that.
//
//   3. AFFORDANCE — should you SEE the button.
//      applyRBAC() in ScriptsMV.html adds `d-none` to eleven element ids in
//      the browser. That is all it does.
//
// Only (3) existed, and (3) is decoration. Deleting a class in devtools
// restores every hidden control, and the server answers the call either way.
//
// WHY THAT IS WORSE HERE THAN IT WOULD BE ELSEWHERE
//
// appsscript.json deploys this web app as:
//
//     "executeAs": "USER_DEPLOYING",  "access": "ANYONE_ANONYMOUS"
//
// so anyone who has the URL can load the page WITHOUT SIGNING IN, open the
// console, and call any of those functions directly — and each one runs with
// the deploying account's full access to the spreadsheet. Not the caller's
// access. The owner's. `updatePharmacyStock`, `getAccountsDashboard`,
// `processPatientDischarge` and the rest are, today, an unauthenticated API
// over the clinic's entire record.
//
// Hiding a nav link does not touch that. Only a check that runs on the server,
// before the work, does.
//
// HOW THIS FILE IS MEANT TO BE USED
//
// Add ONE line as the FIRST statement of a public function:
//
//     function updatePharmacyStock(payload) {
//       var actor = crescRequire_(payload.token, 'pharmacy.stock_edit');
//       ...
//     }
//
// `actor` carries the username, role and display name, so the function also
// stops having to be told who the caller is — several already take a `by` or
// `user` argument from the client, which the client is free to make up.
//
// crescRbacCoverage() lists every public function that still has no guard, so
// the migration is finishable and measurable rather than open-ended.
// ============================================================================


// ---------------------------------------------------------------------------
// SECTION A — THE VOCABULARY
// ---------------------------------------------------------------------------

/**
 * Every permission the application can check, and what it means in the
 * clinic. One flat list, deliberately: a nested tree reads better and is
 * harder to audit, and auditing is the whole point.
 *
 * The naming is `area.verb`. `area` matches the module; `verb` is the thing
 * a person does, not the function that does it — several functions usually
 * share one permission.
 */
var CRESC_PERMS = {
  // --- patients ---------------------------------------------------------
  'patient.read':            'Open a patient record',
  'patient.write':           'Edit patient demographics',
  'patient.register':        'Register a new patient',

  // --- appointments -----------------------------------------------------
  'appointment.read':        'See the appointment ledger',
  'appointment.write':       'Book or reschedule',
  'appointment.cancel':      'Cancel or delete an appointment',
  'schedule.write':          'Publish doctor availability',

  // --- clinical ---------------------------------------------------------
  'emr.read':                'Read clinical notes and the timeline',
  'emr.write':               'Write a consultation or case sheet',
  'rx.write':                'Prescribe',
  'ward.read':               'See the ward, beds and admissions',
  'ward.write':              'Write ward notes and orders',
  'ward.admit':              'Admit a patient / assign a bed',
  'ward.discharge':          'Close an admission',

  // --- laboratory -------------------------------------------------------
  'lab.read':                'See the lab queue and results',
  'lab.order':               'Raise a lab order',
  'lab.collect':             'Collect and accession a sample',
  'lab.result':              'Enter a result',
  'lab.verify':              'Verify and release a report',
  'lab.ack_critical':        'Acknowledge a critical value',
  'lab.catalog':             'Edit the test catalogue and prices',

  // THE DESK'S OWN TILL. Lab bills used to sit behind billing.write, the same
  // key as the pharmacy's and reception's — so a lab technician could raise
  // and settle pharmacy bills and a pharmacist could collect on lab bills.
  // Each desk now has its own key, and billing.* means the hospital desk.
  'lab.bill':                'Raise, collect on and reprint lab bills',

  // --- pharmacy ---------------------------------------------------------
  'pharmacy.read':           'See stock and the dispensing queue',
  'pharmacy.dispense':       'Dispense against a prescription',
  'pharmacy.bill':           'Settle, search and reprint pharmacy bills',
  'pharmacy.register':       'Print the Schedule H / H1 / X drug register',
  'pharmacy.stock_add':      'Receive stock',
  'pharmacy.stock_edit':     'Adjust a batch',
  'pharmacy.stock_discard':  'Write stock off as expired, damaged or lost',
  'pharmacy.return':         'Accept a return / issue a refund',

  // --- money ------------------------------------------------------------
  'billing.read':            'See a hospital bill',
  'billing.write':           'Raise or take payment on a hospital bill',
  'billing.cancel':          'Cancel or void a bill',
  'accounts.read':           'See the Finance Hub and the ledger',
  'accounts.write':          'Record an expense or a transfer',
  'accounts.settle':         'Settle a discharge bill',
  'accounts.lock_period':    'Lock a financial period',
  'accounts.payables':       'Pay a vendor',
  'accounts.tax':            'Tax and compliance',

  // --- reference and overview -------------------------------------------
  // Neither of these touches a patient record, but an endpoint nobody has to
  // sign in for is still an endpoint. They are separated from patient.read so
  // that "may look up a drug dose" never has to mean "may open a patient".
  'reference.read':          'Look up drug, dose and test reference data',
  // THE COMMAND CENTER. The clinic-wide overview — every desk's takings,
  // footfall, occupancy — is for the people who run the clinic. It is held
  // by administrators and doctors; anyone else is given it individually from
  // Staff Accounts → Access, and lands on their own desk instead.
  'dashboard.read':          'Open the Command Center (clinic overview)',

  // --- administration ---------------------------------------------------
  'admin.users':             'Create and disable staff logins',
  // THE SECOND KEY, and the only permission an administrator does not hold.
  //
  // admin.users covers the accounts a clinic manager legitimately opens and
  // closes all year: a nurse, a receptionist, a pharmacist, a lab
  // technician. It does NOT cover the two roles that can undo that
  // boundary. A doctor's account signs prescriptions and is attributable in
  // law; an administrator's account can create more administrators. An
  // account that can mint either of those is an account that can grant
  // itself anything, so it is a separate key held by a separate role.
  'admin.users.elevated':    'Create, disable or reset a doctor or an administrator',
  'admin.audit':             'Read the audit log',
  'admin.config':            'Change clinic configuration',

  // --- data protection (DPDP Act, 2023) ---------------------------------
  // Held by whoever answers a data principal. Kept separate from admin.config
  // because the grievance officer is a named person under s.13, and the
  // clinic may want them to reach the registers without holding the keys to
  // everything else.
  'dpdp.manage':             'Run the data-protection registers and answer a data principal',

  // --- the patient portal -----------------------------------------------
  'portal.self':             'Read your own record in the patient portal'
};


/**
 * ENDPOINTS THAT ANSWER WITHOUT A SESSION, AND WHY.
 *
 * An allowlist exists so that "this one is fine" has to be written down and
 * defended once, instead of being re-argued every time somebody reads the
 * file — and so that tools/rbac.js counts them as a decision rather than as
 * a gap. Anything not on this list and not guarded is a finding.
 *
 * The test each entry has to pass: it returns NO personal data, or it is the
 * step that creates the session in the first place. Nothing is here because
 * adding a token was inconvenient.
 */
var CRESC_PUBLIC_BY_DESIGN = {
  'verifyLogin':        'The sign-in itself. There is no session to check yet; it is rate-limited and audited instead (Auth_Audit.gs).',
  'verifyGoogleLogin':  'The same, for Google sign-in.',
  'verifyMFA':          'The second factor of a sign-in still in progress.',
  'crescChangePassword': 'Changing a password after a forced reset, before a session exists. It proves the old password itself.',
  'getClinicProfile':   'The clinic letterhead — name, address, phone, GSTIN. Already printed on every document that leaves the building, and the invoice renderers need it before sign-in.',
  'crescGetBundle':     'UI markup only: the same HTML the shell used to inline. No patient data passes through it.',
  'getDPDPNotice':      'Section 5 requires the notice to be given AT OR BEFORE collection. A notice you have to sign in to read is not a notice.',
  'crescRequestPasswordReset': 'Forgot password: the person asking cannot sign in, by definition. Mails a temporary password only to the address already on the record, answers every outcome the same way, and is capped per account and per deployment (Auth_Reset.gs).',
  'doGet':              'The web app itself. Serves the sign-in page and the anonymous-by-design document and verification links; every data route inside it checks its own token.',
  'include':            'HTML partials for templates — the same markup the page already ships. No data passes through it.',
  'fetchIPRecordsLedger': 'Legacy stub: returns a fixed "please reload" message and reads nothing.',
  'fetchFullIPRecord':  'Legacy stub: returns a fixed "please reload" message and reads nothing.',
  'onEdit':             'Spreadsheet simple trigger: a hand edit ends the cached copy of a reference list (Master_Cache.gs). From a browser it receives no Range and does nothing.',
  'dpdpSubmitPublicRequest': 'Section 11-13: a data principal must be able to ask without holding a staff login. It writes to a queue that is verified before anything is answered, and is rate-limited per browser.'
};


// ---------------------------------------------------------------------------
// SECTION B — THE MATRIX
// ---------------------------------------------------------------------------

/**
 * Role -> permissions. THIS IS THE ONLY PLACE A ROLE NAME DECIDES ANYTHING.
 *
 * Wherever else in the project you find `if (role === 'admin')`, it is either
 * a bug or something that belongs here.
 *
 * Roles are matched lower-cased, and the aliases the Users sheet has
 * accumulated ('accounts' and 'accountant', 'pharmacy' and 'pharmacist',
 * 'reception' and 'receptionist') are resolved by CRESC_ROLE_ALIAS below
 * rather than by duplicating the lists.
 */
var CRESC_ROLE_MATRIX = {

  // The clinic's administrator. '*' is expanded by crescPermsFor_() to every
  // permission EXCEPT those in CRESC_ELEVATED_PERMS, so a capability added to
  // CRESC_PERMS is granted here automatically — new capability, closed by
  // default — while the key that creates more administrators is not.
  //
  // THERE IS NO SEPARATE 'superadmin' ROLE, and that is deliberate. This
  // project has about twenty-five `role === 'admin'` comparisons scattered
  // through the clinical, billing and appointment modules — the appointment
  // status writers, the barcode action map, the ward's acts-on-behalf list,
  // the template scopes. A new role string would have been refused by every
  // one of them, so the system owner would have signed in and found they
  // could do LESS than the administrators they are meant to be above.
  // Elevation is therefore a property of the ACCOUNT, not a different role:
  // see crescIsElevated_ below. The owner is an administrator, with one
  // extra key.
  admin: ['*'],

  doctor: [
    'patient.read', 'patient.write',
    'appointment.read', 'appointment.write', 'appointment.cancel',
    'emr.read', 'emr.write', 'rx.write',
    'ward.read', 'ward.write', 'ward.admit', 'ward.discharge',
    'lab.read', 'lab.order', 'lab.verify', 'lab.ack_critical',
    'pharmacy.read',
    'billing.read',
    'reference.read', 'dashboard.read'
  ],

  // dashboard.read — the Command Center — is held by admin and doctor only.
  // Every other role lands on its own desk; see CRESC_UI_PERMS and
  // crescLandingTab() in ScriptsMV.html.
  nurse: [
    'patient.read',
    'appointment.read',
    'emr.read', 'emr.write',
    'ward.read', 'ward.write', 'ward.admit',
    // A nurse takes the call from the lab and is the person who reaches the
    // clinician, so acknowledging is theirs. Verifying a result is not.
    'lab.read', 'lab.order', 'lab.collect', 'lab.ack_critical',
    'pharmacy.read',
    'reference.read'
  ],

  receptionist: [
    'patient.read', 'patient.write', 'patient.register',
    'appointment.read', 'appointment.write', 'appointment.cancel',
    'billing.read', 'billing.write',
    'ward.read'
  ],

  pharmacist: [
    'patient.read',
    'pharmacy.read', 'pharmacy.dispense', 'pharmacy.bill', 'pharmacy.register',
    'pharmacy.stock_add', 'pharmacy.stock_edit', 'pharmacy.stock_discard',
    'pharmacy.return',
    'reference.read'
  ],

  lab: [
    'patient.read',
    'lab.read', 'lab.order', 'lab.collect', 'lab.result', 'lab.verify',
    'lab.ack_critical', 'lab.catalog', 'lab.bill',
    'reference.read'
  ],

  accountant: [
    'patient.read',
    'billing.read', 'billing.write', 'billing.cancel',
    'accounts.read', 'accounts.write', 'accounts.settle',
    'accounts.lock_period', 'accounts.payables', 'accounts.tax',
    'ward.read'
  ],

  // A patient reaches exactly one thing: their own record. Every portal
  // function must ALSO check that the record asked for is theirs — the
  // permission says "the portal", not "any patient in it".
  patient: ['portal.self']
};

/** Spellings the Users sheet has used for the same role. */
var CRESC_ROLE_ALIAS = {
  'accounts':     'accountant',
  'account':      'accountant',
  'finance':      'accountant',
  'pharmacy':     'pharmacist',
  'reception':    'receptionist',
  'frontdesk':    'receptionist',
  'front-desk':   'receptionist',
  'labtech':      'lab',
  'lab_tech':     'lab',
  'technician':   'lab',
  'administrator':'admin',
  'sysadmin':     'admin',
  // A Users row spelled with any of these means an administrator whose
  // Super_Admin flag should be set; see crescIsElevated_.
  'superadmin':   'admin',
  'super-admin':  'admin',
  'super_admin':  'admin',
  'owner':        'admin'
};


// ---------------------------------------------------------------------------
// SECTION C — THE GUARD
// ---------------------------------------------------------------------------

/** '' for anything unusable, so nothing below has to test for null. */
function crescStr_(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }

/** The canonical role name for whatever the Users sheet spelled. */
function crescRole_(raw) {
  var r = crescStr_(raw).toLowerCase();
  return CRESC_ROLE_ALIAS[r] || r;
}

/**
 * Permissions no wildcard reaches — only a role that names them, or '**'.
 *
 * Kept as a list rather than as an absence from admin's grant so that the
 * rule survives the next permission somebody adds: a new key is ordinary and
 * flows to admin through '*' unless it is deliberately put here.
 */
var CRESC_ELEVATED_PERMS = ['admin.users.elevated'];

/** Every permission a role holds, with admin's '*' expanded. */
function crescPermsFor_(role) {
  var list = CRESC_ROLE_MATRIX[crescRole_(role)];
  if (!list) return [];
  if (list.length === 1 && list[0] === '*') {
    return Object.keys(CRESC_PERMS).filter(function (p) {
      return CRESC_ELEVATED_PERMS.indexOf(p) === -1;
    });
  }
  return list.slice();
}

/** Does this role hold this permission? The whole matrix in one line. */
function crescCan_(role, permission) {
  return crescPermsFor_(role).indexOf(crescStr_(permission)) !== -1;
}


// ---------------------------------------------------------------------------
// SECTION B2 — PER-PERSON ACCESS, AND THE ACCOUNT'S STATE RIGHT NOW
// ---------------------------------------------------------------------------
//
// A role is a starting point, not a person. Two receptionists are not always
// trusted with the same things, and a nurse covering the lab for a month
// needs the lab for a month. Before this section the only way to give one
// person one more thing was to change their whole role, and the only way to
// take one thing away was to take the role away.
//
// So each staff account carries two lists on its Users row:
//
//   Access_Grant   permissions this person has ON TOP of their role
//   Access_Revoke  permissions of their role this person does NOT have
//
// effective = (role's permissions + Access_Grant) − Access_Revoke
//
// Both are edited from Staff Accounts → Access, every change is audited, and
// both are plain comma-separated keys an owner can read on the sheet.
//
// THE SAME READ ALSO ANSWERS "IS THIS ACCOUNT STILL ALLOWED IN?" A session is
// a cached copy of who someone was when they signed in. Switching an account
// off, or changing its role, used to take effect only when that session
// expired — up to eight hours later, and never, for a session in steady use.
// dc_validateSession_ now asks crescAccountState_() on every call, so a
// switched-off account is refused on its very next click and a role change
// applies on the next one.
//
// COST. The Users sheet is read once and kept in the script cache for
// CRESC_ACL_TTL_S, shared by every user; each call pays one cache read, and
// one execution reads it at most once. Every writer in the application that
// changes an account calls cresc_aclBust_(), and a hand edit to the Users
// sheet busts it through onEdit (Master_Cache.gs).
//
// IF THE SHEET CANNOT BE READ AT ALL (a transient Sheets fault), accounts are
// judged on their role alone — exactly the behaviour before this section —
// rather than every member of staff being signed out at once. An account
// that is READ and found missing or switched off is always refused.
// ---------------------------------------------------------------------------

var CRESC_GRANT_HEADER  = 'Access_Grant';
var CRESC_REVOKE_HEADER = 'Access_Revoke';
var CRESC_ACL_CACHE_KEY = 'CRESC_ACL_V1';
var CRESC_ACL_TTL_S     = 600;

/**
 * Session usernames minted by the script itself for editor-run jobs
 * (RUN_Setup.gs, DPDP_Consent_Backfill.gs, Doctor_Schedule_Engine.gs). They
 * have no Users row by design; every one of them is issued only from a
 * function guarded by crescEditorOnly_.
 */
var CRESC_SYSTEM_USERS = ['SCRIPT_OWNER', 'SYSTEM_SEED'];

/**
 * Keys only the system owner may hand to someone whose role lacks them.
 * Each of these reaches past one person's own work: creating logins,
 * changing clinic-wide configuration, reading everyone's audit trail,
 * answering for the clinic under the DPDP Act, or freezing the books.
 */
var CRESC_OWNER_GRANT_ONLY = ['admin.users', 'admin.config', 'admin.audit',
                              'dpdp.manage', 'accounts.lock_period'];

/**
 * Keys that are never granted individually. The owner's key is a property of
 * the account (the Super_Admin flag), not a permission to hand around; the
 * portal key belongs to patients and means nothing on a staff login.
 */
var CRESC_NEVER_GRANT = ['admin.users.elevated', 'portal.self'];

/** This execution's copy of the access map. Globals do not outlive a call. */
var CRESC_ACL_MEMO = null;

/** A comma/space/semicolon separated list of permission keys, cleaned. */
function cresc_permList_(v) {
  var out = [], seen = {};
  crescStr_(v).split(/[\s,;]+/).forEach(function (p) {
    p = p.trim().toLowerCase();
    if (!p || seen[p] || !CRESC_PERMS.hasOwnProperty(p)) return;
    seen[p] = true;
    out.push(p);
  });
  return out;
}

/**
 * The access map: { users: { USERNAME: {n, r, rr, a, o, g, x} }, at }.
 *   n  the username as typed on the sheet
 *   r  canonical role      rr the role as spelled on the sheet, lower-cased
 *   a  active              o  system owner (Super_Admin)
 *   g  granted keys        x  revoked keys
 * null only when the Users sheet could not be read.
 */
function cresc_aclMap_() {
  if (CRESC_ACL_MEMO) return CRESC_ACL_MEMO;
  var cache = null;
  try {
    cache = CacheService.getScriptCache();
    var raw = cache.get(CRESC_ACL_CACHE_KEY);
    if (raw) { CRESC_ACL_MEMO = JSON.parse(raw); return CRESC_ACL_MEMO; }
  } catch (e) { /* a cache fault is a slower read, not a failure */ }

  var built = cresc_aclBuild_();
  if (built) {
    try { if (cache) cache.put(CRESC_ACL_CACHE_KEY, JSON.stringify(built), CRESC_ACL_TTL_S); }
    catch (e) {}
    CRESC_ACL_MEMO = built;
  }
  return built;
}

/** One read of the Users sheet into the access map. */
function cresc_aclBuild_() {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    if (!sh) return null;
    var users = {};
    if (sh.getLastRow() < 2) return { users: users, at: Date.now() };

    var ownerCol = cresc_superAdminColumn_(sh);
    var data = sh.getDataRange().getValues();
    if (ownerCol !== -1) cresc_seedFirstOwner_(sh, data, ownerCol);

    var head = data[0].map(function (h) { return crescStr_(h); });
    var gCol = head.indexOf(CRESC_GRANT_HEADER);
    var xCol = head.indexOf(CRESC_REVOKE_HEADER);

    for (var i = 1; i < data.length; i++) {
      var name = crescStr_(data[i][0]).toUpperCase();
      if (!name || users[name]) continue;
      var owner = ownerCol !== -1 &&
        (/^(YES|TRUE|1|Y)$/i.test(crescStr_(data[i][ownerCol])) || cresc_seededRow_ === i + 1);
      users[name] = {
        n:  crescStr_(data[i][0]),
        r:  crescRole_(data[i][2]),
        rr: crescStr_(data[i][2]).toLowerCase(),
        // The same test AuthLogin applies at the door: blank is active,
        // anything other than "active" is not.
        a:  !crescStr_(data[i][3]) || crescStr_(data[i][3]).toLowerCase() === 'active',
        o:  !!owner,
        g:  gCol === -1 ? [] : cresc_permList_(data[i][gCol]),
        x:  xCol === -1 ? [] : cresc_permList_(data[i][xCol])
      };
    }
    return { users: users, at: Date.now() };
  } catch (e) {
    return null;
  }
}

/** Forget the cached access map. Call after ANY change to an account. */
function cresc_aclBust_() {
  CRESC_ACL_MEMO = null;
  try { CacheService.getScriptCache().remove(CRESC_ACL_CACHE_KEY); } catch (e) {}
}

/** The map entry for a username, or null. */
function cresc_aclEntry_(username) {
  var map = cresc_aclMap_();
  if (!map || !map.users) return null;
  return map.users[crescStr_(username).toUpperCase()] || null;
}

/**
 * IS THIS SESSION'S ACCOUNT STILL ALLOWED IN, AND AS WHAT?
 *
 * @param {{username:string, role:string}} sess
 * @return {{ok:boolean, role?:string, reason?:string}}
 *   role is the account's CURRENT role as spelled on the Users sheet, for
 *   the caller to put on the session in place of the one it signed in with.
 */
function crescAccountState_(sess) {
  if (!sess) return { ok: false, reason: 'No session.' };
  var role = crescRole_(sess.role);
  // A patient's login lives on the Patients sheet, not Users.
  if (role === 'patient') return { ok: true, role: crescStr_(sess.role) };

  var user = crescStr_(sess.username).toUpperCase();
  var map = cresc_aclMap_();
  if (!map || !map.users) return { ok: true, role: crescStr_(sess.role) };   // unreadable: role only

  var entry = map.users[user];
  if (!entry && CRESC_SYSTEM_USERS.indexOf(user) !== -1) {
    return { ok: true, role: crescStr_(sess.role) };
  }
  if (!entry) {
    // Before refusing, read the sheet once more: a row typed in by hand (the
    // onEdit bust does not fire for every kind of edit) must not lock its new
    // owner out for the life of the cached copy.
    cresc_aclBust_();
    map = cresc_aclMap_();
    entry = (map && map.users) ? map.users[user] : null;
    if (!map || !map.users) return { ok: true, role: crescStr_(sess.role) };
  }
  if (!entry) return { ok: false, reason: 'This account no longer exists.' };
  if (!entry.a) return { ok: false, reason: 'This account has been switched off.' };
  return { ok: true, role: entry.rr || entry.r };
}

/**
 * Everything one person may do: their role, plus what they were granted,
 * minus what was taken away, plus the owner's key when the account is the
 * system owner.
 *
 * @return {{perms:Array<string>, elevated:boolean, granted:Array<string>,
 *           revoked:Array<string>}}
 */
function crescEffectivePerms_(username, role) {
  var canonical = crescRole_(role);
  var perms = crescPermsFor_(canonical);
  var out = { perms: perms, elevated: false, granted: [], revoked: [] };
  if (!perms.length || canonical === 'patient') return out;

  var entry = cresc_aclEntry_(username);
  if (!entry) return out;

  (entry.g || []).forEach(function (p) {
    if (CRESC_NEVER_GRANT.indexOf(p) !== -1) return;
    if (perms.indexOf(p) === -1) { perms.push(p); out.granted.push(p); }
  });
  (entry.x || []).forEach(function (p) {
    var at = perms.indexOf(p);
    if (at !== -1) { perms.splice(at, 1); out.revoked.push(p); }
  });

  if (canonical === 'admin' && entry.o) {
    out.elevated = true;
    CRESC_ELEVATED_PERMS.forEach(function (p) {
      if (perms.indexOf(p) === -1) perms.push(p);
    });
  }
  return out;
}

/** The Users column that marks an administrator as the system owner. */
var CRESC_SUPERADMIN_HEADER = 'Super_Admin';

/**
 * IS THIS ACCOUNT THE SYSTEM OWNER?
 *
 * YES / TRUE / 1 / Y in the Super_Admin column of the Users sheet, on an
 * account whose role is already admin. Anything else, including a missing
 * column, is no — the flag has to be set deliberately and nothing grants it
 * by accident.
 *
 * BOOTSTRAP. With the column absent or empty on every row, nobody is
 * elevated and the elevated endpoints refuse everybody. That is the safe
 * direction, but it is also a clinic locked out of adding its first doctor,
 * so crescEnsureSuperAdminColumn_() marks the FIRST administrator on the
 * sheet when no one is marked at all. One owner, chosen by the order the
 * accounts were created in, and visible on the sheet afterwards.
 *
 * @param {string} username
 * @param {string} role  the canonical role, already through crescRole_
 * @return {boolean}
 */
function crescIsElevated_(username, role) {
  if (crescRole_(role) !== 'admin') return false;
  var want = crescStr_(username).toUpperCase();
  if (!want) return false;
  // Answered from the shared access map (SECTION B2) rather than a read of
  // the Users sheet on every call an administrator makes. The map is built
  // by the same Super_Admin column read — including the first-owner seed —
  // and is busted whenever an account changes.
  var entry = cresc_aclEntry_(want);
  return !!(entry && entry.o && entry.r === 'admin');
}

/** The Super_Admin column index, adding the column if it is missing. */
function cresc_superAdminColumn_(sh) {
  var m = dc_headerMap_(sh);
  if (m[CRESC_SUPERADMIN_HEADER] === undefined) {
    if (typeof dc_ensureColumn_ !== 'function') return -1;
    dc_ensureColumn_(sh, CRESC_SUPERADMIN_HEADER);
    m = dc_headerMap_(sh);
  }
  var col = m[CRESC_SUPERADMIN_HEADER];
  return (col === undefined) ? -1 : col;
}

/** The row this execution seeded, so the caller's stale copy can agree. */
var cresc_seededRow_ = -1;

/**
 * Marks the first administrator as owner when NOBODY is marked.
 *
 * Idempotent, and it only ever fires on a sheet with no owner at all: once a
 * clinic has one, this cannot appoint another.
 */
function cresc_seedFirstOwner_(sh, data, col) {
  var firstAdminRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (!crescStr_(data[i][0])) continue;
    if (/^(YES|TRUE|1|Y)$/i.test(crescStr_(data[i][col]))) return;   // already owned
    if (firstAdminRow === -1 && crescRole_(data[i][2]) === 'admin' &&
        crescStr_(data[i][3]).toUpperCase() !== 'INACTIVE') {
      firstAdminRow = i + 1;
    }
  }
  if (firstAdminRow === -1) return;
  try {
    sh.getRange(firstAdminRow, col + 1).setValue('YES');
    cresc_seededRow_ = firstAdminRow;
    if (typeof dc_invalidate_ === 'function') dc_invalidate_('Users');
  } catch (e) { /* advisory */ }
}

/**
 * The column, ensured, for callers that only need it to exist — the staff
 * account list reads it directly afterwards.
 */
function crescEnsureSuperAdminColumn_(sh) {
  var col = cresc_superAdminColumn_(sh);
  if (col === -1) return;
  cresc_seedFirstOwner_(sh, sh.getDataRange().getValues(), col);
}

/**
 * THE AMBIENT ACTOR — who this execution has already proved itself to be.
 *
 * WHY IT EXISTS. Guarding every endpoint broke something that was not
 * obvious until it was done: server functions call each other. The discharge
 * settlement calls the credit-bill settler; the case sheet calls
 * createLabRequest(); the ward note calls saveIPNote(); the archive link
 * builder calls getLabBillHtml(). The browser passed a token to the OUTER
 * call and there is nothing to pass to the inner one, so a guard on the
 * inner function would refuse the clinic's own code.
 *
 * The fix is not to leave the inner function open — it is reachable from the
 * browser too. It is to remember, for the length of ONE execution, the actor
 * the outer guard already resolved, and let the inner guard check ITS OWN
 * permission against that same person. A nurse whose ward note raises a lab
 * order still has to hold lab.order; she simply does not have to re-prove
 * who she is.
 *
 * WHY THIS IS SAFE. Apps Script gives every google.script.run call a fresh
 * script context: globals are re-initialised per execution and never shared
 * between two callers or two requests. So this variable cannot leak one
 * user's identity into another user's call — it has the lifetime of a single
 * server call and dies with it. It is only ever SET by a successful
 * crescRequire_ / crescActor_, so an unauthenticated outer call leaves it
 * null and every inner guard still fails closed.
 */
var CRESC_CURRENT_ACTOR = null;

/**
 * Records whoever a session check just accepted as the ambient actor, if
 * nobody has been recorded yet in this execution.
 *
 * crescRequire_ is not the only door. The doctor, IP, portal and billing
 * screens validate through dc_validateSession_ directly (resolveScope_,
 * resolveWriteDoctor_, pp_me_, hb_actor_), and before this an inner helper
 * guarded by crescRequire_(undefined) could not see who they had let in. Now
 * any successful validation is visible to the helpers it calls — and a direct
 * google.script.run call to a helper still finds nobody and is refused.
 *
 * No elevated-permission lookup here: that costs a Users read and only
 * crescActor_ needs it, which overwrites this with the full actor.
 */
function crescNoteSession_(sess) {
  if (CRESC_CURRENT_ACTOR || !sess) return;
  var role = crescRole_(sess.role);
  // The person's own access, not just their role's: an inner guard reached
  // through a ward note or a billing desk must refuse what was revoked from
  // this person exactly as the outer one would.
  var eff = crescEffectivePerms_(sess.username, role);
  CRESC_CURRENT_ACTOR = {
    username:    crescStr_(sess.username),
    role:        role,
    rawRole:     crescStr_(sess.role),
    elevated:    eff.elevated,
    doctorId:    crescStr_(sess.doctorId),
    displayName: crescStr_(sess.displayName) || crescStr_(sess.name) || crescStr_(sess.username),
    permissions: eff.perms
  };
}

/**
 * Resolves a session token to an actor, or null.
 *
 * Never throws: callers that want a hard stop use crescRequire_(). This is
 * for the handful of reads that legitimately degrade — a dashboard that
 * shows fewer cards to a role rather than refusing to load.
 *
 * A call with no usable token inherits the ambient actor if — and only if —
 * an outer call in the same execution has already been authorised. See
 * CRESC_CURRENT_ACTOR above.
 *
 * @param {string} token
 * @return {{username,role,doctorId,displayName,permissions}|null}
 */
function crescActor_(token) {
  if (!crescStr_(token) && CRESC_CURRENT_ACTOR) return CRESC_CURRENT_ACTOR;

  var sess = null;
  try {
    sess = (typeof dc_validateSession_ === 'function')
      ? dc_validateSession_(token)
      : validateSession_(token);
  } catch (e) { sess = null; }
  if (!sess) return CRESC_CURRENT_ACTOR || null;

  // The role as the Users sheet says it is NOW (dc_validateSession_ has
  // already put it on the session), then this person's own grants and
  // revocations, then — for the system owner — the one extra key.
  var role = crescRole_(sess.role);
  var eff = crescEffectivePerms_(sess.username, role);
  var perms = eff.perms;
  var elevated = eff.elevated;

  var actor = {
    username:    crescStr_(sess.username),
    role:        role,
    rawRole:     crescStr_(sess.role),
    elevated:    elevated,
    doctorId:    crescStr_(sess.doctorId),
    displayName: (typeof dc_sessionName_ === 'function')
                   ? dc_sessionName_(sess) : crescStr_(sess.username),
    permissions: perms
  };
  CRESC_CURRENT_ACTOR = actor;
  return actor;
}

/**
 * THE GUARD. First statement of every public function that touches data.
 *
 * Fails closed at three separate points — no session, unknown role, missing
 * permission — and the message says which, because "Access denied" sends the
 * user to the wrong person for help. An expired session is the common case
 * and is worded as what it is, so nobody files it as a permissions bug.
 *
 * @param {string} token       the caller's sessionToken
 * @param {string|Array<string>} permission  one permission, or any-of a list
 * @return {{username,role,doctorId,displayName,permissions}} the actor
 * @throws {Error} prefixed FORBIDDEN:
 */
function crescRequire_(token, permission) {
  var actor = crescActor_(token);

  if (!actor) {
    throw new Error('FORBIDDEN: your session has expired. Please sign in again.');
  }
  if (!actor.permissions.length) {
    throw new Error('FORBIDDEN: the role "' + (actor.rawRole || 'unset') +
                    '" on your account is not one this system knows. ' +
                    'Ask an administrator to correct it on the Users sheet.');
  }

  var need = (permission === undefined || permission === null) ? []
           : (typeof permission === 'string' ? [permission] : permission);

  // A permission that is not in the vocabulary is a typo at the call site,
  // and a typo must not silently pass. It is a server fault, not the user's,
  // so it is not dressed up as FORBIDDEN.
  for (var i = 0; i < need.length; i++) {
    if (!CRESC_PERMS.hasOwnProperty(need[i])) {
      throw new Error('RBAC misconfiguration: "' + need[i] + '" is not a known ' +
                      'permission. See CRESC_PERMS in RBAC.gs.');
    }
  }

  if (need.length) {
    var held = false;
    for (var j = 0; j < need.length; j++) {
      if (actor.permissions.indexOf(need[j]) !== -1) { held = true; break; }
    }
    if (!held) {
      throw new Error('FORBIDDEN: your role (' + actor.role + ') cannot ' +
                      (CRESC_PERMS[need[0]] || need[0]).toLowerCase() + '.');
    }
  }

  return actor;
}

/**
 * FOR FUNCTIONS MEANT TO BE RUN FROM THE SCRIPT EDITOR, NOT THE BROWSER.
 *
 * Every top-level function without a trailing underscore is reachable through
 * google.script.run, and this deployment runs as the owner. So enrolMFA(),
 * diagnoseMFA(), crescMigrateCredentials(), setClinicProfile() and the other
 * runbook tools were endpoints anyone with the URL could call by name — MFA
 * secrets and the live TOTP code for any account included.
 *
 * They take no session token (the editor's Run button cannot pass one), so
 * this lets a call through in exactly two cases:
 *
 *   1. It is the script owner, in the editor: the signed-in user IS the
 *      account the script runs as.
 *   2. It is an inner call from an endpoint that already checked a signed-in
 *      user holding one of `perms` (default admin.config), via the ambient
 *      actor.
 *
 * A web-app visitor is neither: Apps Script reports no active user, or a
 * different one, for anyone but the owner.
 *
 * @param {string} what      the function's name, for the refusal
 * @param {string|string[]} [perms]
 */
function crescEditorOnly_(what, perms) {
  var need = perms ? (typeof perms === 'string' ? [perms] : perms) : ['admin.config'];
  if (CRESC_CURRENT_ACTOR && CRESC_CURRENT_ACTOR.permissions) {
    for (var i = 0; i < need.length; i++) {
      if (CRESC_CURRENT_ACTOR.permissions.indexOf(need[i]) !== -1) return;
    }
  }
  var active = '', effective = '';
  try { active = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase(); } catch (e) {}
  try { effective = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase(); } catch (e) {}
  if (active && active === effective) return;
  throw new Error('FORBIDDEN: ' + what + '() can only be run from the Apps Script editor.');
}

/**
 * FOR TIME-DRIVEN TRIGGER HANDLERS.
 *
 * A trigger can only call a public function, so the nightly and monthly jobs
 * are reachable through google.script.run as well. Apps Script passes every
 * trigger run an event object carrying the trigger's unique id; this checks
 * that id against the project's installed triggers. A browser can pass an
 * object too, but it cannot know a real trigger id. The owner running the job
 * by hand from the editor is let through by crescEditorOnly_.
 *
 * @param {Object} e     the handler's first argument
 * @param {string} what  the handler's name, for the refusal
 */
function crescTriggerOnly_(e, what) {
  var uid = (e && typeof e === 'object') ? String(e.triggerUid || '') : '';
  if (uid) {
    try {
      var ts = ScriptApp.getProjectTriggers();
      for (var i = 0; i < ts.length; i++) {
        if (String(ts[i].getUniqueId()) === uid) return;
      }
    } catch (err) {}
  }
  crescEditorOnly_(what);
}

/**
 * WHO IS DOING THIS — for a "…By" column.
 *
 * Session.getActiveUser() was used for this in the lab and pharmacy modules,
 * and for this web app — deployed to run as its owner — it names the OWNER,
 * or nobody, whoever is actually signed in. Every lab order, sample, result,
 * bill and every pharmacy return in the clinic's history was therefore
 * recorded as the work of one account. The session knows who it is; this
 * reads it from the actor the endpoint's own guard already resolved.
 *
 * @param {string} [fallback]  for a scheduled job with no signed-in person
 * @return {string} the signed-in username
 */
function cresc_actorName_(fallback) {
  var a = CRESC_CURRENT_ACTOR;
  if (a && a.username) return a.username;
  return fallback || 'SYSTEM';
}

/** The signed-in person's display name, for a document someone reads. */
function cresc_actorDisplay_(fallback) {
  var a = CRESC_CURRENT_ACTOR;
  if (a) return a.displayName || a.username || fallback || 'SYSTEM';
  return fallback || 'SYSTEM';
}

/**
 * A CAUGHT ERROR, SAID IN A SENTENCE THE USER CAN ACT ON.
 *
 * It lives here, beside crescRequire_, because crescRequire_ is what
 * produces most of them. "FORBIDDEN: your session has expired. Please sign
 * in again." is already the right words; the prefix is not, and the prefix
 * is the first thing a receptionist reads.
 *
 * Every frontend entry point that catches its own guard runs its message
 * through this, so a refusal reads the same wherever it surfaces.
 *
 * @param {Error|string} error
 * @return {string}
 */
function cresc_reason_(error) {
  var msg = (error && error.message) ? String(error.message) : String(error || '');
  msg = msg.replace(/^FORBIDDEN:\s*/, '');
  return msg || 'The request could not be completed.';
}

/**
 * A patient may only ever reach their own record.
 *
 * The portal's session username IS the patient id, so this is one comparison
 * — but it has to happen on every portal read, and forgetting it turns
 * 'portal.self' into 'portal.everyone'. Staff roles pass through: they hold
 * patient.read and are audited instead.
 */
function crescRequireOwnRecord_(token, patientId) {
  var actor = crescRequire_(token, ['portal.self', 'patient.read']);
  if (actor.role !== 'patient') return actor;

  var mine = crescStr_(actor.username).toUpperCase();
  var want = crescStr_(patientId).toUpperCase();
  if (!want || want !== mine) {
    throw new Error('FORBIDDEN: this record does not belong to your account.');
  }
  return actor;
}


// ---------------------------------------------------------------------------
// SECTION D — WHAT THE CLIENT IS ALLOWED TO KNOW
// ---------------------------------------------------------------------------

/**
 * FRONTEND ENTRY. The signed-in user's permissions, for drawing buttons.
 *
 * ADVISORY ONLY, and worth being blunt about: this exists so the interface
 * does not offer a control that will be refused, not to enforce anything. A
 * caller who edits the reply, or never calls this at all, gains nothing —
 * every function that matters calls crescRequire_() for itself.
 *
 * @param {string} token
 * @return {{success:boolean, role:string, permissions:Object, message:string}}
 */
function crescGetMyPermissions(token) {
  try {
    var actor = crescActor_(token);
    if (!actor) {
      return { success: false, role: '', permissions: {},
               message: 'Your session has expired. Please sign in again.' };
    }
    // An object, not an array: the client tests `p['lab.verify']`, and a
    // missing key is falsy without anyone having to remember indexOf.
    var map = {};
    actor.permissions.forEach(function (p) { map[p] = true; });
    return { success: true, role: actor.role, permissions: map,
             displayName: actor.displayName, message: '' };
  } catch (err) {
    return { success: false, role: '', permissions: {}, message: err.message };
  }
}


// ---------------------------------------------------------------------------
// SECTION E — KEEPING THE MATRIX HONEST
// ---------------------------------------------------------------------------

/**
 * Run from the script editor after editing the matrix.
 *
 * Catches the three ways a permission matrix rots: a role granted something
 * that does not exist (a typo — silently grants nothing), a permission no
 * role holds (dead, or an oversight), and a role nobody can be.
 */
function crescRbacSelfTest() {
  crescEditorOnly_('crescRbacSelfTest');
  var problems = [];
  var granted = {};

  Object.keys(CRESC_ROLE_MATRIX).forEach(function (role) {
    var list = CRESC_ROLE_MATRIX[role];
    if (list.length === 1 && list[0] === '*') {
      Object.keys(CRESC_PERMS).forEach(function (p) { granted[p] = true; });
      return;
    }
    list.forEach(function (p) {
      if (!CRESC_PERMS.hasOwnProperty(p)) {
        problems.push('Role "' + role + '" is granted "' + p + '", which is not a ' +
                      'known permission — it grants nothing.');
      }
      granted[p] = true;
    });
  });

  Object.keys(CRESC_PERMS).forEach(function (p) {
    if (!granted[p]) problems.push('Permission "' + p + '" is held by no role at all.');
  });

  Object.keys(CRESC_ROLE_ALIAS).forEach(function (alias) {
    var target = CRESC_ROLE_ALIAS[alias];
    if (!CRESC_ROLE_MATRIX[target]) {
      problems.push('Alias "' + alias + '" points at "' + target + '", which has no matrix entry.');
    }
  });

  // The Access screen's groups (RBAC_Access.gs) must list every key a person
  // can be given, once — a key missing from them is one no screen can grant.
  if (typeof CRESC_ACCESS_GROUPS !== 'undefined') {
    var listed = {};
    CRESC_ACCESS_GROUPS.forEach(function (g) {
      g.perms.forEach(function (p) {
        if (!CRESC_PERMS.hasOwnProperty(p)) problems.push('Access group "' + g.key + '" lists unknown "' + p + '".');
        if (listed[p]) problems.push('"' + p + '" is listed in two access groups.');
        listed[p] = true;
      });
    });
    Object.keys(CRESC_PERMS).forEach(function (p) {
      if (!listed[p] && CRESC_NEVER_GRANT.indexOf(p) === -1) {
        problems.push('"' + p + '" is in no access group, so Staff Accounts → Access cannot show it.');
      }
    });
  }

  // Every role actually present on the Users sheet must resolve to something,
  // and every key typed into Access_Grant / Access_Revoke must be a real one.
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    if (sheet && sheet.getLastRow() > 1) {
      var all = sheet.getDataRange().getValues();
      var head = all[0].map(function (h) { return crescStr_(h); });
      var seen = {};
      for (var r = 1; r < all.length; r++) {
        var raw = crescStr_(all[r][2]);
        if (raw && !seen[raw]) {
          seen[raw] = true;
          if (!CRESC_ROLE_MATRIX[crescRole_(raw)]) {
            problems.push('The Users sheet contains role "' + raw + '", which resolves ' +
                          'to nothing — those accounts can do nothing at all.');
          }
        }
        [CRESC_GRANT_HEADER, CRESC_REVOKE_HEADER].forEach(function (h) {
          var c = head.indexOf(h);
          if (c === -1) return;
          crescStr_(all[r][c]).split(/[\s,;]+/).forEach(function (k) {
            k = k.trim().toLowerCase();
            if (k && !CRESC_PERMS.hasOwnProperty(k)) {
              problems.push('Users row ' + (r + 1) + ' (' + crescStr_(all[r][0]) + ') has "' + k +
                            '" in ' + h + ', which is not a permission — it is ignored.');
            }
          });
        });
      }
    }
  } catch (e) { problems.push('Could not read the Users sheet: ' + e.message); }

  var report = problems.length
    ? 'crescRbacSelfTest: ' + problems.length + ' problem(s)\n' + problems.join('\n')
    : 'crescRbacSelfTest: ' + Object.keys(CRESC_PERMS).length + ' permissions across ' +
      Object.keys(CRESC_ROLE_MATRIX).length + ' roles, all consistent.';
  Logger.log(report);
  return { success: problems.length === 0, message: report };
}

/**
 * How much of the server is actually guarded.
 *
 * Reports every function reachable from google.script.run that does NOT call
 * crescRequire_, dsx_requireRole_ or a session validator. Run it, fix the top
 * of the list, run it again — the number only goes down, and it is the honest
 * measure of how far "make it real RBAC" has got.
 *
 * Reads the project's own source through the Apps Script API when the
 * advanced service is enabled; without it, it reports what it can and says so.
 */
function crescRbacCoverage() {
  crescEditorOnly_('crescRbacCoverage', ['admin.config', 'dpdp.manage']);
  var GUARDS = /crescRequire_|crescActor_|crescRequireOwnRecord_|dsx_requireRole_|dc_validateSession_|validateSession_/;
  // Private (trailing underscore) functions are not reachable from the client
  // and are guarded by whatever public function called them.
  var PUBLIC_FN = /^function\s+([A-Za-z_$][\w$]*?)\s*\(/gm;

  var files;
  try {
    var id = ScriptApp.getScriptId();
    var res = UrlFetchApp.fetch(
      'https://script.googleapis.com/v1/projects/' + id + '/content',
      { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) {
      var hint = 'crescRbacCoverage needs the Apps Script API turned on for this ' +
                 'account (script.google.com/home/usersettings) and the ' +
                 'https://www.googleapis.com/auth/script.projects scope in ' +
                 'appsscript.json. The API answered ' + res.getResponseCode() + '.';
      Logger.log(hint);
      return { success: false, message: hint };
    }
    files = JSON.parse(res.getContentText()).files || [];
  } catch (e) {
    Logger.log('crescRbacCoverage could not read the project: ' + e.message);
    return { success: false, message: e.message };
  }

  var guarded = [], open = [];
  files.forEach(function (f) {
    if (f.type !== 'SERVER_JS') return;
    var src = String(f.source || '');
    var bodies = src.split(/^function\s+/m);
    var m;
    PUBLIC_FN.lastIndex = 0;
    while ((m = PUBLIC_FN.exec(src)) !== null) {
      var name = m[1];
      if (name.charAt(name.length - 1) === '_') continue;   // private by convention
      // The body is everything up to the next top-level `function`.
      var start = m.index;
      var next = src.indexOf('\nfunction ', start + 1);
      var body = src.substring(start, next === -1 ? src.length : next);
      (GUARDS.test(body) ? guarded : open).push(f.name + '.gs  ' + name);
    }
  });

  open.sort();
  var total = guarded.length + open.length;
  var pct = total ? Math.round(guarded.length / total * 100) : 0;
  var report = 'crescRbacCoverage: ' + guarded.length + ' of ' + total +
               ' public server functions are guarded (' + pct + '%).\n\n' +
               'UNGUARDED — reachable by anyone with the web app URL:\n' +
               open.join('\n');
  Logger.log(report);
  return { success: open.length === 0, guarded: guarded.length, open: open.length,
           unguarded: open, message: report };
}
