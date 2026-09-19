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

  // --- pharmacy ---------------------------------------------------------
  'pharmacy.read':           'See stock and the dispensing queue',
  'pharmacy.dispense':       'Dispense against a prescription',
  'pharmacy.stock_add':      'Receive stock',
  'pharmacy.stock_edit':     'Adjust a batch',
  'pharmacy.stock_discard':  'Write stock off as expired, damaged or lost',
  'pharmacy.return':         'Accept a return / issue a refund',

  // --- money ------------------------------------------------------------
  'billing.read':            'See a bill',
  'billing.write':           'Raise or take payment on a bill',
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
  'dashboard.read':          'See the clinic overview dashboard',

  // --- administration ---------------------------------------------------
  'admin.users':             'Create and disable staff logins',
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

  // The system owner. '*' is expanded by crescPermsFor_(), so a permission
  // added to CRESC_PERMS is granted to admin automatically and to nobody else
  // — new capability, closed by default.
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

  nurse: [
    'patient.read',
    'appointment.read',
    'emr.read', 'emr.write',
    'ward.read', 'ward.write', 'ward.admit',
    // A nurse takes the call from the lab and is the person who reaches the
    // clinician, so acknowledging is theirs. Verifying a result is not.
    'lab.read', 'lab.order', 'lab.collect', 'lab.ack_critical',
    'pharmacy.read',
    'reference.read', 'dashboard.read'
  ],

  receptionist: [
    'patient.read', 'patient.write', 'patient.register',
    'appointment.read', 'appointment.write', 'appointment.cancel',
    'billing.read', 'billing.write',
    'ward.read',
    'dashboard.read'
  ],

  pharmacist: [
    'patient.read',
    'pharmacy.read', 'pharmacy.dispense', 'pharmacy.stock_add',
    'pharmacy.stock_edit', 'pharmacy.stock_discard', 'pharmacy.return',
    'billing.read', 'billing.write',
    'reference.read', 'dashboard.read'
  ],

  lab: [
    'patient.read',
    'lab.read', 'lab.order', 'lab.collect', 'lab.result', 'lab.verify',
    'lab.ack_critical', 'lab.catalog',
    'billing.read', 'billing.write',
    'reference.read', 'dashboard.read'
  ],

  accountant: [
    'patient.read',
    'billing.read', 'billing.write',
    'accounts.read', 'accounts.write', 'accounts.settle',
    'accounts.lock_period', 'accounts.payables', 'accounts.tax',
    'ward.read',
    'dashboard.read'
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
  'sysadmin':     'admin'
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

/** Every permission a role holds, with admin's '*' expanded. */
function crescPermsFor_(role) {
  var list = CRESC_ROLE_MATRIX[crescRole_(role)];
  if (!list) return [];
  if (list.length === 1 && list[0] === '*') return Object.keys(CRESC_PERMS);
  return list.slice();
}

/** Does this role hold this permission? The whole matrix in one line. */
function crescCan_(role, permission) {
  return crescPermsFor_(role).indexOf(crescStr_(permission)) !== -1;
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

  var role = crescRole_(sess.role);
  var actor = {
    username:    crescStr_(sess.username),
    role:        role,
    rawRole:     crescStr_(sess.role),
    doctorId:    crescStr_(sess.doctorId),
    displayName: (typeof dc_sessionName_ === 'function')
                   ? dc_sessionName_(sess) : crescStr_(sess.username),
    permissions: crescPermsFor_(role)
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

  // Every role actually present on the Users sheet must resolve to something.
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    if (sheet && sheet.getLastRow() > 1) {
      var rows = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues();
      var seen = {};
      rows.forEach(function (r) {
        var raw = crescStr_(r[0]);
        if (!raw || seen[raw]) return;
        seen[raw] = true;
        if (!CRESC_ROLE_MATRIX[crescRole_(raw)]) {
          problems.push('The Users sheet contains role "' + raw + '", which resolves ' +
                        'to nothing — those accounts can do nothing at all.');
        }
      });
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
