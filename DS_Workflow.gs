// ============================================================================
// DS_Workflow.gs  —  Crescentia HealthTech / CresRx
// IP Discharge Summary Engine · Phase 3 · RBAC, state machine, public API
// ----------------------------------------------------------------------------
// Prefix rules: see DS_Setup.gs. Public API is `ds_*` (contract names, each
// verified not to collide with Doctor_Session_Store.gs or
// Doctor_Schedule_Engine.gs); everything private is `dsx_*`.
//
// THE RULE THIS MODULE EXISTS TO BREAK WITH
//   The IP Admissions module has no backend role validation at all —
//   processPatientDischarge() takes no session token. Nothing here copies that
//   pattern. dsx_requireRole_() is the FIRST statement of every public
//   function, the permission matrix is enforced server-side, and the
//   permissions object returned to the client is for drawing buttons only.
//
// LOCK DISCIPLINE
//   Assembly, diffing, rendering and PDF generation run OUTSIDE the lock. The
//   project shares one global script lock with every other module; holding it
//   through a 30-day assembly would stall the pharmacy counter.
// ============================================================================

var DSX_LOCK_MS = 15000;
var DSX_PRESENCE_TTL = 120;     // seconds
var DSX_SIGN_FAIL_WINDOW = 900; // 15 minutes, in seconds
var DSX_SIGN_FAIL_MAX = 5;

// ---------------------------------------------------------------------------
// SECTION A — AUTHORISATION
// ---------------------------------------------------------------------------

/**
 * Every action the module can perform. The matrix below is the only place
 * that decides who may do what; nothing else in this file tests a role name
 * directly except the signer-identity rule, which needs the doctor profile.
 */
var DSX_ACTIONS = [
  'view', 'viewSigned', 'viewMedsOnly', 'initiate', 'generate', 'edit',
  'submit', 'return', 'sign', 'amend', 'cancel', 'printDraft', 'printFinal'
];

/**
 * Role -> actions. `preparer` is not a login role; it is resolved at runtime
 * from DS_PREPARER_ROLES (default "nurse,doctor") and unioned into whatever
 * the role already has.
 */
var DSX_ROLE_MATRIX = {
  doctor: ['view', 'viewSigned', 'initiate', 'generate', 'edit', 'submit', 'return',
           'sign', 'amend', 'cancel', 'printDraft', 'printFinal'],
  admin: ['view', 'viewSigned', 'cancel', 'printFinal'],
  receptionist: ['viewSigned', 'printFinal'],
  reception: ['viewSigned', 'printFinal'],
  accountant: ['viewSigned', 'printFinal'],
  accounts: ['viewSigned', 'printFinal'],
  pharmacist: ['viewMedsOnly'],
  pharmacy: ['viewMedsOnly']
};

/** What a configured preparer role gains. */
var DSX_PREPARER_ACTIONS =
  ['view', 'viewSigned', 'generate', 'edit', 'submit', 'printDraft', 'printFinal'];

/**
 * Resolves the caller and asserts they hold at least one of `required`.
 *
 * @param {string} token
 * @param {Array<string>|string} required  action name(s); [] means "any known role"
 * @return {{username, role, doctorId, displayName, actions, cfg}}
 * @throws FORBIDDEN
 */
function dsx_requireRole_(token, required) {
  var sess = null;
  try { sess = dc_validateSession_(token); } catch (e) { sess = null; }
  if (!sess) throw new Error('FORBIDDEN: your session has expired. Please sign in again.');

  var role = dsx_str_(sess.role).toLowerCase();
  if (!role || role === 'patient') {
    throw new Error('FORBIDDEN: this area is not available to your account.');
  }

  var cfg = dsx_config_();
  var actions = (DSX_ROLE_MATRIX[role] || []).slice();
  if (cfg.preparerRoles.indexOf(role) !== -1) {
    DSX_PREPARER_ACTIONS.forEach(function (a) {
      if (actions.indexOf(a) === -1) actions.push(a);
    });
  }
  if (!actions.length) {
    throw new Error('FORBIDDEN: your role (' + role + ') has no discharge-summary access.');
  }

  var need = (required === undefined || required === null) ? []
           : (typeof required === 'string' ? [required] : required);
  if (need.length) {
    var held = need.some(function (a) { return actions.indexOf(a) !== -1; });
    if (!held) {
      throw new Error('FORBIDDEN: your role (' + role + ') cannot ' + need[0] + ' a discharge summary.');
    }
  }

  return {
    username: dsx_str_(sess.username),
    role: role,
    doctorId: dsx_str_(sess.doctorId),
    displayName: (typeof dc_sessionName_ === 'function') ? dc_sessionName_(sess) : dsx_str_(sess.username),
    actions: actions,
    cfg: cfg
  };
}

/** The permissions object the client uses to draw buttons. Advisory only. */
function dsx_permissions_(actor, header) {
  var status = header ? dsx_upper_(header.Status) : '';
  var can = function (a) { return actor.actions.indexOf(a) !== -1; };
  var live = status && status !== DSX_STATUS.SIGNED && status !== DSX_STATUS.CANCELLED;
  var editable = live && status !== DSX_STATUS.PENDING_SIGNATURE;

  return {
    view: can('view') || can('viewSigned') || can('viewMedsOnly'),
    viewSignedOnly: !can('view') && can('viewSigned'),
    viewMedsOnly: !can('view') && !can('viewSigned') && can('viewMedsOnly'),
    initiate: can('initiate'),
    generate: can('generate') && live && status !== DSX_STATUS.AMENDMENT_IN_PROGRESS,
    edit: can('edit') &&
          (editable || (can('sign') && status === DSX_STATUS.PENDING_SIGNATURE) ||
           status === DSX_STATUS.AMENDMENT_IN_PROGRESS),
    submit: can('submit') &&
            (status === DSX_STATUS.GENERATED || status === DSX_STATUS.IN_PREPARATION),
    return: can('return') && status === DSX_STATUS.PENDING_SIGNATURE,
    sign: can('sign') && dsx_canSignFrom_(status, actor.cfg),
    amend: can('amend') && status === DSX_STATUS.SIGNED,
    discardAmendment: can('amend') && status === DSX_STATUS.AMENDMENT_IN_PROGRESS,
    cancel: can('cancel') && live,
    printDraft: can('printDraft') && status !== DSX_STATUS.SIGNED,
    printFinal: can('printFinal') && status === DSX_STATUS.SIGNED
  };
}

function dsx_canSignFrom_(status, cfg) {
  if (status === DSX_STATUS.PENDING_SIGNATURE) return true;
  if (status === DSX_STATUS.AMENDMENT_IN_PROGRESS) return true;
  // Solo-clinic fast path: off by default, per-tenant config.
  if (!cfg.requirePreparerReview) {
    return status === DSX_STATUS.GENERATED || status === DSX_STATUS.IN_PREPARATION;
  }
  return false;
}

// ---------------------------------------------------------------------------
// SECTION B — THE STATE MACHINE
// ---------------------------------------------------------------------------

/**
 * One transition table. Anything not listed here is INVALID_STATE — there is
 * no second place in this file where a status is assigned.
 *
 * `from: null` means "no summary row exists yet".
 */
var DSX_TRANSITIONS = [
  { action: 'INITIATE',   from: [null, DSX_STATUS.CANCELLED],                    to: DSX_STATUS.GENERATED },
  { action: 'PREPARE',    from: [DSX_STATUS.GENERATED, DSX_STATUS.RETURNED],     to: DSX_STATUS.IN_PREPARATION },
  { action: 'EDIT',       from: [DSX_STATUS.GENERATED, DSX_STATUS.IN_PREPARATION,
                                 DSX_STATUS.RETURNED, DSX_STATUS.PENDING_SIGNATURE,
                                 DSX_STATUS.AMENDMENT_IN_PROGRESS],              to: null /* unchanged */ },
  { action: 'REGENERATE', from: [DSX_STATUS.GENERATED, DSX_STATUS.IN_PREPARATION,
                                 DSX_STATUS.RETURNED],                           to: null },
  { action: 'SUBMIT',     from: [DSX_STATUS.GENERATED, DSX_STATUS.IN_PREPARATION], to: DSX_STATUS.PENDING_SIGNATURE },
  { action: 'RETURN',     from: [DSX_STATUS.PENDING_SIGNATURE],                  to: DSX_STATUS.RETURNED },
  { action: 'SIGN',       from: [DSX_STATUS.PENDING_SIGNATURE],                  to: DSX_STATUS.SIGNED },
  // Solo-clinic fast path, gated on DS_REQUIRE_PREPARER_REVIEW = false.
  { action: 'SIGN_FAST',  from: [DSX_STATUS.GENERATED, DSX_STATUS.IN_PREPARATION], to: DSX_STATUS.SIGNED },
  { action: 'SIGN_AMEND', from: [DSX_STATUS.AMENDMENT_IN_PROGRESS],              to: DSX_STATUS.SIGNED },
  { action: 'AMEND_START',from: [DSX_STATUS.SIGNED],                             to: DSX_STATUS.AMENDMENT_IN_PROGRESS },
  { action: 'AMEND_DISCARD', from: [DSX_STATUS.AMENDMENT_IN_PROGRESS],           to: DSX_STATUS.SIGNED },
  { action: 'CANCEL',     from: [DSX_STATUS.GENERATED, DSX_STATUS.IN_PREPARATION,
                                 DSX_STATUS.PENDING_SIGNATURE, DSX_STATUS.RETURNED], to: DSX_STATUS.CANCELLED }
];

/**
 * @return {{ok:boolean, to:string|null, message:string}}
 *         `to` null means "status unchanged, this action does not move it".
 */
function dsx_transition_(action, fromStatus) {
  var from = fromStatus ? dsx_upper_(fromStatus) : null;
  for (var i = 0; i < DSX_TRANSITIONS.length; i++) {
    var t = DSX_TRANSITIONS[i];
    if (t.action !== action) continue;
    if (t.from.indexOf(from) === -1) continue;
    return { ok: true, to: t.to, message: '' };
  }
  return {
    ok: false, to: null,
    message: 'Cannot ' + action.toLowerCase().replace(/_/g, ' ') +
             ' a summary that is ' + (from || 'not yet created') + '.'
  };
}

// ---------------------------------------------------------------------------
// SECTION C — OPTIMISTIC CONCURRENCY
// ---------------------------------------------------------------------------

/**
 * Re-reads the header inside the lock and refuses the write if someone else
 * moved first. The script lock prevents corruption; this prevents LOST
 * UPDATES, which is the failure a nurse and a doctor editing together actually
 * hit.
 *
 * @throws Error carrying a JSON body the caller converts to VERSION_CONFLICT
 */
function dsx_checkVersion_(header, expectedRowVersion) {
  var current = dsx_int_(header.Row_Version);
  var expected = dsx_int_(expectedRowVersion);
  if (expected === current) return current;

  var who = dsx_str_(header.Updated_By) || 'another user';
  var when = dsx_fmt_(header.Updated_At, 'dd-MMM hh:mm a') || 'just now';
  var e = new Error('VERSION_CONFLICT: Updated by ' + who + ' at ' + when + '. Reload to continue.');
  e.dsConflict = {
    currentRowVersion: current,
    updatedBy: who,
    updatedAt: when
  };
  throw e;
}

/** Stamps the header after a successful mutation. Caller holds the lock. */
function dsx_bumpVersion_(sheet, header, actor, extra) {
  var fields = extra || {};
  fields.Row_Version = dsx_int_(header.Row_Version) + 1;
  fields.Updated_At = new Date();
  fields.Updated_By = dsx_str_(actor.username);
  dsx_writeRow_(sheet, header._row, fields);
  return fields.Row_Version;
}

// ---------------------------------------------------------------------------
// SECTION D — EDITING PRESENCE (CacheService only; no sheet write, no lock)
// ---------------------------------------------------------------------------

function dsx_presenceKey_(summaryId) { return 'DSPRES_' + dsx_upper_(summaryId); }

/**
 * Advisory only. Presence tells a nurse that the doctor has the summary open;
 * the Row_Version check is the real guard, and nothing here can block a write.
 */
function ds_heartbeatEditing(token, summaryId) {
  try {
    var actor = dsx_requireRole_(token, ['view', 'edit']);
    var cache = CacheService.getScriptCache();
    var key = dsx_presenceKey_(summaryId);

    var list = [];
    try { list = JSON.parse(cache.get(key) || '[]'); } catch (e) { list = []; }

    var now = Date.now();
    list = list.filter(function (p) {
      return p && p.username !== actor.username && (now - dsx_int_(p.ms)) < DSX_PRESENCE_TTL * 1000;
    });
    list.push({ username: actor.username, role: actor.role,
                name: actor.displayName, ms: now, at: dsx_nowIso_() });

    cache.put(key, JSON.stringify(list), DSX_PRESENCE_TTL);
    return dsx_ok_('', { editors: dsx_presenceOthers_(list, actor.username) });
  } catch (e) {
    return dsx_fromError_(e);
  }
}

function dsx_presenceList_(summaryId) {
  try {
    var raw = CacheService.getScriptCache().get(dsx_presenceKey_(summaryId));
    var list = raw ? JSON.parse(raw) : [];
    var now = Date.now();
    return list.filter(function (p) {
      return p && (now - dsx_int_(p.ms)) < DSX_PRESENCE_TTL * 1000;
    });
  } catch (e) { return []; }
}

function dsx_presenceOthers_(list, username) {
  return (list || []).filter(function (p) { return p.username !== username; })
                     .map(function (p) { return { username: p.username, role: p.role, name: p.name }; });
}

// ---------------------------------------------------------------------------
// SECTION E — READ ENDPOINTS
// ---------------------------------------------------------------------------

/**
 * The Discharge Desk. Reads DS_Summaries only (no payloads) plus the active
 * admissions that have no summary yet, so the desk loads fast.
 *
 * @param {Object} filter {status, search, limit}
 */
function ds_getQueue(token, filter) {
  try {
    var actor = dsx_requireRole_(token, ['view', 'viewSigned', 'viewMedsOnly']);
    filter = filter || {};

    var sh = dsx_summariesSheet_();
    var map = dsx_headerMap_(sh);
    var lastRow = sh.getLastRow();
    var rows = [];

    if (lastRow >= 2) {
      var width = Math.max(1, sh.getLastColumn());
      var values = sh.getRange(2, 1, lastRow - 1, width).getValues();
      var tenant = dsx_upper_(dsx_tenant_());
      var now = Date.now();

      for (var i = 0; i < values.length; i++) {
        var v = values[i];
        var id = dsx_str_(v[map['Summary_ID']]);
        if (!id) continue;
        var rowTenant = dsx_upper_(v[map['Tenant_ID']]);
        if (rowTenant && rowTenant !== tenant) continue;

        var status = dsx_upper_(v[map['Status']]);
        var initiatedAt = dsx_toDate_(v[map['Initiated_At']]);
        var signedAt = dsx_toDate_(v[map['Signed_At']]);
        var endMs = signedAt ? signedAt.getTime() : now;
        var tatMin = initiatedAt ? Math.round((endMs - initiatedAt.getTime()) / 60000) : 0;

        rows.push({
          summaryId: id,
          ipNumber: dsx_str_(v[map['IP_Number']]),
          patientId: dsx_str_(v[map['Patient_ID']]),
          dischargeType: dsx_upper_(v[map['Discharge_Type']]) || 'NORMAL',
          status: status,
          rowVersion: dsx_int_(v[map['Row_Version']]),
          initiatedAt: dsx_fmt_(initiatedAt, 'dd-MMM hh:mm a'),
          plannedAt: dsx_fmt_(v[map['Planned_Discharge_At']], 'dd-MMM hh:mm a'),
          preparedBy: dsx_str_(v[map['Prepared_By']]),
          signedAt: dsx_fmt_(signedAt, 'dd-MMM hh:mm a'),
          signedBy: dsx_str_(v[map['Signed_By']]),
          returnedCount: dsx_int_(v[map['Returned_Count']]),
          pdfStatus: dsx_upper_(v[map['Pdf_Status']]),
          lastActor: dsx_str_(v[map['Updated_By']]),
          lastActorAt: dsx_fmt_(v[map['Updated_At']], 'dd-MMM hh:mm a'),
          tatMinutes: tatMin,
          tatLevel: dsx_tatLevel_(status, tatMin, actor.cfg)
        });
      }
    }

    // Role trimming: anyone who may only view signed summaries sees only those.
    var canSeeDrafts = actor.actions.indexOf('view') !== -1;
    if (!canSeeDrafts) {
      rows = rows.filter(function (r) { return r.status === DSX_STATUS.SIGNED; });
    }

    var counts = {};
    Object.keys(DSX_STATUS).forEach(function (k) { counts[DSX_STATUS[k]] = 0; });
    rows.forEach(function (r) { if (counts[r.status] !== undefined) counts[r.status]++; });

    var out = rows;
    var wanted = dsx_upper_(filter.status);
    if (wanted && wanted !== 'ALL') {
      out = out.filter(function (r) { return r.status === wanted; });
    }
    // Patient names come from the admissions sheet, which the queue does not
    // otherwise read. Decorate BEFORE searching, or a search by patient name
    // would drop every row before its name had been loaded.
    dsx_decorateWithAdmission_(out);

    var q = dsx_upper_(filter.search);
    if (q) {
      out = out.filter(function (r) {
        return r.ipNumber.toUpperCase().indexOf(q) > -1 ||
               r.patientId.toUpperCase().indexOf(q) > -1 ||
               dsx_upper_(r.patientName).indexOf(q) > -1;
      });
    }

    out.sort(function (a, b) { return b.tatMinutes - a.tatMinutes; });
    var limit = dsx_int_(filter.limit) || 200;
    if (out.length > limit) out = out.slice(0, limit);

    var awaiting = canSeeDrafts ? dsx_admissionsWithoutSummary_(rows) : [];

    return dsx_ok_('', {
      rows: out,
      counts: counts,
      awaitingInitiation: awaiting,
      permissions: dsx_permissions_(actor, null),
      config: { tatAmberMin: actor.cfg.tatAmberMin, tatRedMin: actor.cfg.tatRedMin }
    });
  } catch (e) {
    return dsx_fromError_(e);
  }
}

function dsx_tatLevel_(status, minutes, cfg) {
  if (status === DSX_STATUS.SIGNED || status === DSX_STATUS.CANCELLED) return 'done';
  if (minutes >= cfg.tatRedMin) return 'red';
  if (minutes >= cfg.tatAmberMin) return 'amber';
  return 'ok';
}

/** Fills patientName / ward / bed / consultant from IP_Admissions. */
function dsx_decorateWithAdmission_(rows) {
  if (!rows.length) return;
  var sh = dsx_ss_().getSheetByName('IP_Admissions');
  if (!sh || sh.getLastRow() < 2) return;

  var map = dsx_headerMap_(sh);
  var width = Math.max(1, sh.getLastColumn());
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();

  var byIp = {};
  for (var i = 0; i < values.length; i++) {
    byIp[dsx_ip_(values[i][map['IP_Number']])] = values[i];
  }
  rows.forEach(function (r) {
    var v = byIp[dsx_ip_(r.ipNumber)];
    if (!v) { r.patientName = r.patientName || ''; return; }
    var wb = ipa_resolveWardBed_(v[map['Ward_Bed']], v[map['Bed']]);
    r.patientName = dsx_str_(v[map['Patient_Name']]);
    r.ageSex = dsx_str_(v[map['Age_Sex']]);
    r.ward = wb.ward;
    r.bed = wb.bed;
    r.consultant = dsx_str_(v[map['Consultant']]);
    r.admissionStatus = dsx_upper_(v[map['Status']]);
  });
}

/** Active admissions with no discharge summary row yet. */
function dsx_admissionsWithoutSummary_(existingRows) {
  var have = {};
  existingRows.forEach(function (r) {
    if (r.status !== DSX_STATUS.CANCELLED) have[dsx_ip_(r.ipNumber)] = true;
  });

  var sh = dsx_ss_().getSheetByName('IP_Admissions');
  if (!sh || sh.getLastRow() < 2) return [];

  var map = dsx_headerMap_(sh);
  var width = Math.max(1, sh.getLastColumn());
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();
  var out = [];

  for (var i = 0; i < values.length; i++) {
    var v = values[i];
    var ip = dsx_ip_(v[map['IP_Number']]);
    if (!ip || have[ip]) continue;
    if (!ipa_isLive_(v[map['Status']])) continue;
    var wb = ipa_resolveWardBed_(v[map['Ward_Bed']], v[map['Bed']]);
    out.push({
      ipNumber: dsx_str_(v[map['IP_Number']]),
      patientId: dsx_str_(v[map['Patient_ID']]),
      patientName: dsx_str_(v[map['Patient_Name']]),
      ageSex: dsx_str_(v[map['Age_Sex']]),
      ward: wb.ward, bed: wb.bed,
      consultant: dsx_str_(v[map['Consultant']]),
      doa: dsx_fmt_(v[map['DOA']], 'dd-MMM-yyyy'),
      diagnosis: dsx_str_(v[map['Diagnosis']])
    });
  }
  return out;
}

/** Status badges for the IP Admissions and IP Notes screens. */
function ds_getStatusMap(token, ipNumbers) {
  try {
    dsx_requireRole_(token, ['view', 'viewSigned', 'viewMedsOnly']);
    var wanted = {};
    (ipNumbers || []).forEach(function (ip) { wanted[dsx_ip_(ip)] = true; });

    var sh = dsx_summariesSheet_();
    var map = dsx_headerMap_(sh);
    var out = {};
    if (sh.getLastRow() < 2) return dsx_ok_('', out);

    var width = Math.max(1, sh.getLastColumn());
    var values = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();
    for (var i = 0; i < values.length; i++) {
      var ip = dsx_ip_(values[i][map['IP_Number']]);
      if (!ip) continue;
      if (Object.keys(wanted).length && !wanted[ip]) continue;
      out[ip] = {
        summaryId: dsx_str_(values[i][map['Summary_ID']]),
        status: dsx_upper_(values[i][map['Status']]),
        dischargeType: dsx_upper_(values[i][map['Discharge_Type']]) || 'NORMAL',
        rowVersion: dsx_int_(values[i][map['Row_Version']])
      };
    }
    return dsx_ok_('', out);
  } catch (e) {
    return dsx_fromError_(e);
  }
}

/** Everything the editor needs in one call. */
function ds_getSummary(token, summaryId) {
  try {
    var actor = dsx_requireRole_(token, ['view', 'viewSigned', 'viewMedsOnly']);
    var header = dsx_getHeader_(summaryId);
    if (!header) return dsx_err_('VALIDATION_FAILED', 'No discharge summary found for ' + summaryId + '.');

    var status = dsx_upper_(header.Status);
    var canSeeDrafts = actor.actions.indexOf('view') !== -1;
    if (!canSeeDrafts && status !== DSX_STATUS.SIGNED) {
      return dsx_err_('FORBIDDEN', 'This summary has not been signed yet.');
    }

    var working = dsx_getWorking_(summaryId);
    var payload = working ? working.payload : null;

    if (payload) {
      payload = dsx_trimPayloadForRole_(payload, actor);
    }

    var readiness = { hard: [], soft: [] };
    if (payload && typeof dsx_readiness_ === 'function') {
      try { readiness = dsx_readiness_(payload, null, header); } catch (e) { /* advisory */ }
    }

    return dsx_ok_('', {
      header: dsx_headerForClient_(header),
      payload: payload,
      readiness: readiness,
      permissions: dsx_permissions_(actor, header),
      events: dsx_recentEvents_(summaryId, 50),
      editors: dsx_presenceOthers_(dsx_presenceList_(summaryId), actor.username),
      snapshots: dsx_listSnapshots_(summaryId).map(function (s) {
        return { snapshotNo: s.snapshotNo, type: s.type,
                 at: dsx_fmt_(s.createdAt, 'dd-MMM hh:mm a'), by: s.createdBy,
                 shortHash: dsx_shortHash_(s.contentHash) };
      })
    });
  } catch (e) {
    return dsx_fromError_(e);
  }
}

function dsx_headerForClient_(h) {
  return {
    summaryId: dsx_str_(h.Summary_ID),
    ipNumber: dsx_str_(h.IP_Number),
    patientId: dsx_str_(h.Patient_ID),
    dischargeType: dsx_upper_(h.Discharge_Type) || 'NORMAL',
    status: dsx_upper_(h.Status),
    rowVersion: dsx_int_(h.Row_Version),
    currentSnapshotNo: dsx_int_(h.Current_Snapshot_No),
    lastSignedSnapshotNo: dsx_int_(h.Last_Signed_Snapshot_No),
    initiatedAt: dsx_fmt_(h.Initiated_At, 'dd-MMM-yyyy hh:mm a'),
    initiatedBy: dsx_str_(h.Initiated_By),
    plannedDischargeAt: dsx_fmt_(h.Planned_Discharge_At, 'dd-MMM-yyyy hh:mm a'),
    clinicalDischargeAt: dsx_fmt_(h.Clinical_Discharge_At, 'dd-MMM-yyyy hh:mm a'),
    preparedBy: dsx_str_(h.Prepared_By),
    submittedAt: dsx_fmt_(h.Submitted_At, 'dd-MMM hh:mm a'),
    submittedBy: dsx_str_(h.Submitted_By),
    returnedCount: dsx_int_(h.Returned_Count),
    signedAt: dsx_fmt_(h.Signed_At, 'dd-MMM-yyyy hh:mm a'),
    signedBy: dsx_str_(h.Signed_By),
    signerRegNo: dsx_str_(h.Signer_Reg_No),
    shortHash: dsx_shortHash_(h.Signed_Hash),
    pdfStatus: dsx_upper_(h.Pdf_Status),
    printCount: dsx_int_(h.Print_Count),
    cancelReason: dsx_str_(h.Cancel_Reason)
  };
}

/** A pharmacist sees the discharge medications of a signed summary. Nothing else. */
function dsx_trimPayloadForRole_(payload, actor) {
  if (actor.actions.indexOf('view') !== -1 || actor.actions.indexOf('viewSigned') !== -1) {
    return payload;
  }
  if (actor.actions.indexOf('viewMedsOnly') === -1) return null;

  var trimmed = {
    schemaVersion: payload.schemaVersion,
    dischargeType: payload.dischargeType,
    meta: payload.meta,
    sections: {}
  };
  ['PATIENT_BANNER', 'DISCHARGE_MEDICATIONS'].forEach(function (k) {
    if (payload.sections[k]) trimmed.sections[k] = payload.sections[k];
  });
  return trimmed;
}

// ---------------------------------------------------------------------------
// SECTION F — MUTATIONS
// ---------------------------------------------------------------------------

/**
 * Creates (or reopens) the summary for an active admission and runs assembly.
 */
function ds_initiateDischarge(token, ipNumber, dischargeType, plannedDischargeAtIso) {
  var lock = LockService.getScriptLock();
  try {
    var actor = dsx_requireRole_(token, 'initiate');
    var ip = dsx_ip_(ipNumber);
    if (!ip) return dsx_err_('VALIDATION_FAILED', 'An IP number is required.');

    var type = dsx_upper_(dischargeType) || 'NORMAL';
    if (DSX_DISCHARGE_TYPES.indexOf(type) === -1) {
      return dsx_err_('VALIDATION_FAILED', 'Unknown discharge type "' + dischargeType + '".');
    }

    var adm = dsx_admissionRow_(ip);
    if (!adm) return dsx_err_('VALIDATION_FAILED', 'Admission ' + ip + ' was not found.');
    if (!ipa_isLive_(adm.Status)) {
      return dsx_err_('INVALID_STATE', 'Admission ' + ip + ' is already ' +
                      dsx_upper_(adm.Status) + '. A summary can only be initiated on a live admission.');
    }

    var summaryId = dsx_summaryIdFor_(ip);
    var existing = dsx_getHeader_(summaryId);
    var from = existing ? dsx_upper_(existing.Status) : null;

    var t = dsx_transition_('INITIATE', from);
    if (!t.ok) {
      return dsx_err_('INVALID_STATE',
        'A discharge summary already exists for ' + ip + ' and is ' + from +
        '. Open it from the Discharge Desk.');
    }

    // ---- assembly runs OUTSIDE the lock -----------------------------------
    if (typeof dsx_assemble_ !== 'function') {
      return dsx_err_('NOT_IMPLEMENTED',
        'The clinical assembly engine is not installed. Add DS_Assembly.gs.');
    }
    var built = dsx_assemble_(ip, type, actor);   // {payload, warnings, timings}

    // ---- commit ------------------------------------------------------------
    lock.waitLock(DSX_LOCK_MS);
    dsx_resetHeaderCache_();

    var sh = dsx_summariesSheet_();
    var header = dsx_getHeader_(summaryId);
    var fromNow = header ? dsx_upper_(header.Status) : null;
    if (!dsx_transition_('INITIATE', fromNow).ok) {
      return dsx_err_('INVALID_STATE', 'Another user initiated this discharge a moment ago. Reload.');
    }

    var now = new Date();
    var snapshotNo = header ? dsx_int_(header.Current_Snapshot_No) + 1 : 1;
    var canonical = dsx_canonicalJson_(built.payload);
    var hash = dsx_sha256Hex_(canonical);

    var fields = {
      Summary_ID: summaryId,
      Tenant_ID: dsx_tenant_(),
      IP_Number: ip,
      Patient_ID: dsx_pid_(adm.Patient_ID),
      Discharge_Type: type,
      Status: DSX_STATUS.GENERATED,
      Current_Snapshot_No: snapshotNo,
      Planned_Discharge_At: dsx_toDate_(plannedDischargeAtIso) || '',
      Updated_At: now,
      Updated_By: actor.username,
      Cancel_Reason: ''
    };

    var rowVersion;
    if (header) {
      fields.Row_Version = dsx_int_(header.Row_Version) + 1;
      rowVersion = fields.Row_Version;
      dsx_writeRow_(sh, header._row, fields);
    } else {
      fields.Row_Version = 1;
      fields.Initiated_At = now;
      fields.Initiated_By = actor.username;
      fields.Last_Signed_Snapshot_No = 0;
      fields.Returned_Count = 0;
      fields.Print_Count = 0;
      fields.Pdf_Status = '';
      rowVersion = 1;
      dsx_appendRow_(sh, fields);
    }

    dsx_appendSnapshot_(summaryId, snapshotNo, 'BASELINE', built.payload, hash, '', actor.username, '');
    dsx_putWorking_(summaryId, built.payload, snapshotNo, actor.username);

    dsx_logEvent_(summaryId, actor, 'DS_INITIATE', fromNow, DSX_STATUS.GENERATED,
                  snapshotNo, hash, '', { dischargeType: type, warnings: built.warnings });
    dsx_audit_(actor, 'DS_INITIATE', summaryId, { ipNumber: ip, dischargeType: type });

    SpreadsheetApp.flush();

    return dsx_ok_('Discharge summary generated for ' + ip + '.', {
      summaryId: summaryId,
      rowVersion: rowVersion,
      warnings: built.warnings,
      timings: built.timings
    });

  } catch (e) {
    return dsx_conflictOr_(e);
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/**
 * Saves section patches. Only `content`, `reviewed` and AI acceptance may
 * change; read-only sections are refused outright.
 *
 * @param {Object} sectionPatches {SECTION_KEY: {content, reviewed}}
 */
function ds_saveWorking(token, summaryId, expectedRowVersion, sectionPatches) {
  var lock = LockService.getScriptLock();
  try {
    var actor = dsx_requireRole_(token, 'edit');
    if (!sectionPatches || !Object.keys(sectionPatches).length) {
      return dsx_err_('VALIDATION_FAILED', 'Nothing to save.');
    }

    lock.waitLock(DSX_LOCK_MS);
    dsx_resetHeaderCache_();

    var header = dsx_getHeader_(summaryId);
    if (!header) return dsx_err_('VALIDATION_FAILED', 'No discharge summary found for ' + summaryId + '.');
    dsx_checkVersion_(header, expectedRowVersion);

    var status = dsx_upper_(header.Status);
    var t = dsx_transition_('EDIT', status);
    if (!t.ok) return dsx_err_('INVALID_STATE', t.message);

    // A doctor may edit while PENDING_SIGNATURE; a preparer may not.
    if (status === DSX_STATUS.PENDING_SIGNATURE && actor.actions.indexOf('sign') === -1) {
      return dsx_err_('INVALID_STATE',
        'This summary is with the doctor for signature. Ask them to return it for correction.');
    }
    if (status === DSX_STATUS.AMENDMENT_IN_PROGRESS && actor.actions.indexOf('amend') === -1) {
      return dsx_err_('INVALID_STATE', 'Only a doctor may edit an amendment in progress.');
    }

    var working = dsx_getWorking_(summaryId);
    if (!working || !working.payload) {
      return dsx_err_('VALIDATION_FAILED', 'The working draft is missing. Regenerate the summary.');
    }

    var payload = working.payload;
    var applied = [], rejected = [];
    var nowIso = dsx_nowIso_();

    Object.keys(sectionPatches).forEach(function (key) {
      var k = dsx_upper_(key);
      if (DSX_READONLY_SECTIONS.indexOf(k) !== -1) {
        rejected.push(k + ' is read-only.');
        return;
      }
      var sec = payload.sections[k];
      if (!sec) { rejected.push(k + ' is not a section of this summary.'); return; }

      var patch = sectionPatches[key] || {};
      var touched = false;

      if (patch.content !== undefined) {
        sec.content = patch.content;
        sec.edited = true;
        sec.editedBy = actor.username;
        sec.editedAt = nowIso;
        sec.origin = (sec.origin === 'AI') ? 'AI' : 'MANUAL';
        sec.sourceChangedSinceEdit = false;
        touched = true;
      }
      if (patch.reviewed !== undefined) { sec.reviewed = !!patch.reviewed; touched = true; }
      if (patch.aiAccepted === true) {
        sec.aiAcceptedBy = actor.username;
        sec.aiAcceptedAt = nowIso;
        touched = true;
      }
      if (touched) applied.push(k);
    });

    if (!applied.length) {
      return dsx_err_('VALIDATION_FAILED', rejected.join(' ') || 'Nothing to save.');
    }

    var check = dsx_validatePayload_(payload);
    if (!check.ok) return dsx_err_('VALIDATION_FAILED', check.errors.join(' '));

    dsx_putWorking_(summaryId, payload, working.baseSnapshotNo, actor.username);

    // The first save by a preparer moves GENERATED | RETURNED -> IN_PREPARATION.
    var extra = {};
    var toStatus = status;
    var moved = dsx_transition_('PREPARE', status);
    if (moved.ok) {
      toStatus = moved.to;
      extra.Status = toStatus;
      if (!dsx_str_(header.Prepared_By)) extra.Prepared_By = actor.username;
    }

    var rowVersion = dsx_bumpVersion_(dsx_summariesSheet_(), header, actor, extra);
    dsx_logEvent_(summaryId, actor, 'DS_SAVE', status, toStatus,
                  dsx_int_(header.Current_Snapshot_No), '', '', { sections: applied });
    SpreadsheetApp.flush();

    return dsx_ok_('Saved.', {
      rowVersion: rowVersion, status: toStatus,
      applied: applied, rejected: rejected
    });

  } catch (e) {
    return dsx_conflictOr_(e);
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/** Hands the summary to the doctor. All HARD readiness items must pass. */
function ds_submit(token, summaryId, expectedRowVersion, softAcks) {
  var lock = LockService.getScriptLock();
  try {
    var actor = dsx_requireRole_(token, 'submit');

    var header0 = dsx_getHeader_(summaryId);
    if (!header0) return dsx_err_('VALIDATION_FAILED', 'No discharge summary found for ' + summaryId + '.');
    var t0 = dsx_transition_('SUBMIT', dsx_upper_(header0.Status));
    if (!t0.ok) return dsx_err_('INVALID_STATE', t0.message);

    var working0 = dsx_getWorking_(summaryId);
    if (!working0 || !working0.payload) {
      return dsx_err_('VALIDATION_FAILED', 'The working draft is missing. Regenerate the summary.');
    }

    // Readiness is recomputed on the server, outside the lock. Whatever the
    // client believed is irrelevant.
    var readiness = { hard: [], soft: [] };
    if (typeof dsx_readiness_ === 'function') {
      readiness = dsx_readiness_(working0.payload, null, header0);
    }
    if (readiness.hard.length) {
      return dsx_err_('VALIDATION_FAILED',
        'This summary is not ready for signature: ' +
        readiness.hard.map(function (h) { return h.label; }).join('; ') + '.',
        { readiness: readiness });
    }

    lock.waitLock(DSX_LOCK_MS);
    dsx_resetHeaderCache_();

    var header = dsx_getHeader_(summaryId);
    dsx_checkVersion_(header, expectedRowVersion);
    var status = dsx_upper_(header.Status);
    var t = dsx_transition_('SUBMIT', status);
    if (!t.ok) return dsx_err_('INVALID_STATE', t.message);

    var working = dsx_getWorking_(summaryId);
    var snapshotNo = dsx_int_(header.Current_Snapshot_No) + 1;
    var hash = dsx_sha256Hex_(dsx_canonicalJson_(working.payload));

    dsx_appendSnapshot_(summaryId, snapshotNo, 'SUBMITTED', working.payload, hash, '', actor.username, '');
    dsx_putWorking_(summaryId, working.payload, working.baseSnapshotNo, actor.username);

    var rowVersion = dsx_bumpVersion_(dsx_summariesSheet_(), header, actor, {
      Status: t.to,
      Current_Snapshot_No: snapshotNo,
      Submitted_At: new Date(),
      Submitted_By: actor.username,
      Prepared_By: dsx_str_(header.Prepared_By) || actor.username
    });

    dsx_logEvent_(summaryId, actor, 'DS_SUBMIT', status, t.to, snapshotNo, hash, '',
                  { softAcks: softAcks || [], soft: readiness.soft });
    dsx_audit_(actor, 'DS_SUBMIT', summaryId, { ipNumber: dsx_str_(header.IP_Number) });
    SpreadsheetApp.flush();

    return dsx_ok_('Submitted for signature.', {
      rowVersion: rowVersion, status: t.to, snapshotNo: snapshotNo,
      shortHash: dsx_shortHash_(hash)
    });

  } catch (e) {
    return dsx_conflictOr_(e);
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/** Doctor sends it back. The comment is mandatory; section comments are pinned. */
function ds_return(token, summaryId, expectedRowVersion, comment, sectionComments) {
  var lock = LockService.getScriptLock();
  try {
    var actor = dsx_requireRole_(token, 'return');
    var text = dsx_str_(comment);
    if (!text) {
      return dsx_err_('VALIDATION_FAILED',
        'Say what needs correcting — the preparer cannot act on an empty comment.');
    }

    lock.waitLock(DSX_LOCK_MS);
    dsx_resetHeaderCache_();

    var header = dsx_getHeader_(summaryId);
    if (!header) return dsx_err_('VALIDATION_FAILED', 'No discharge summary found for ' + summaryId + '.');
    dsx_checkVersion_(header, expectedRowVersion);

    var status = dsx_upper_(header.Status);
    var t = dsx_transition_('RETURN', status);
    if (!t.ok) return dsx_err_('INVALID_STATE', t.message);

    var rowVersion = dsx_bumpVersion_(dsx_summariesSheet_(), header, actor, {
      Status: t.to,
      Returned_Count: dsx_int_(header.Returned_Count) + 1
    });

    dsx_logEvent_(summaryId, actor, 'DS_RETURN', status, t.to,
                  dsx_int_(header.Current_Snapshot_No), '', text,
                  { sectionComments: sectionComments || {} });
    dsx_audit_(actor, 'DS_RETURN', summaryId, { ipNumber: dsx_str_(header.IP_Number), comment: text });
    SpreadsheetApp.flush();

    return dsx_ok_('Returned for correction.', { rowVersion: rowVersion, status: t.to });

  } catch (e) {
    return dsx_conflictOr_(e);
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/** Pre-signature cancellation. Doctor or admin, reason mandatory. */
function ds_cancel(token, summaryId, expectedRowVersion, reason) {
  var lock = LockService.getScriptLock();
  try {
    var actor = dsx_requireRole_(token, 'cancel');
    var text = dsx_str_(reason);
    if (!text) return dsx_err_('VALIDATION_FAILED', 'A cancellation reason is required.');

    lock.waitLock(DSX_LOCK_MS);
    dsx_resetHeaderCache_();

    var header = dsx_getHeader_(summaryId);
    if (!header) return dsx_err_('VALIDATION_FAILED', 'No discharge summary found for ' + summaryId + '.');
    dsx_checkVersion_(header, expectedRowVersion);

    var status = dsx_upper_(header.Status);
    var t = dsx_transition_('CANCEL', status);
    if (!t.ok) return dsx_err_('INVALID_STATE', t.message);

    var rowVersion = dsx_bumpVersion_(dsx_summariesSheet_(), header, actor, {
      Status: t.to, Cancel_Reason: text
    });

    dsx_logEvent_(summaryId, actor, 'DS_CANCEL', status, t.to, 0, '', text, {});
    dsx_audit_(actor, 'DS_CANCEL', summaryId, { ipNumber: dsx_str_(header.IP_Number), reason: text });
    SpreadsheetApp.flush();

    return dsx_ok_('Discharge summary cancelled.', { rowVersion: rowVersion, status: t.to });

  } catch (e) {
    return dsx_conflictOr_(e);
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/** Opens an amendment on a signed summary. Reason mandatory. */
function ds_startAmendment(token, summaryId, expectedRowVersion, reason) {
  var lock = LockService.getScriptLock();
  try {
    var actor = dsx_requireRole_(token, 'amend');
    var text = dsx_str_(reason);
    if (!text) return dsx_err_('VALIDATION_FAILED', 'An amendment reason is required and will be printed.');

    lock.waitLock(DSX_LOCK_MS);
    dsx_resetHeaderCache_();

    var header = dsx_getHeader_(summaryId);
    if (!header) return dsx_err_('VALIDATION_FAILED', 'No discharge summary found for ' + summaryId + '.');
    dsx_checkVersion_(header, expectedRowVersion);

    var status = dsx_upper_(header.Status);
    var t = dsx_transition_('AMEND_START', status);
    if (!t.ok) return dsx_err_('INVALID_STATE', t.message);

    var lastSigned = dsx_int_(header.Last_Signed_Snapshot_No);
    var snap = dsx_getSnapshotPayload_(summaryId, lastSigned);
    if (!snap || !snap.payload) {
      return dsx_err_('VALIDATION_FAILED',
        'The last signed version could not be read, so an amendment cannot start from it.');
    }

    dsx_putWorking_(summaryId, snap.payload, lastSigned, actor.username);
    var rowVersion = dsx_bumpVersion_(dsx_summariesSheet_(), header, actor, { Status: t.to });

    dsx_logEvent_(summaryId, actor, 'DS_AMEND_START', status, t.to, lastSigned, '', text, {});
    dsx_audit_(actor, 'DS_AMEND', summaryId, { ipNumber: dsx_str_(header.IP_Number), reason: text });
    SpreadsheetApp.flush();

    return dsx_ok_('Amendment started from signed version ' + lastSigned + '.',
                   { rowVersion: rowVersion, status: t.to });

  } catch (e) {
    return dsx_conflictOr_(e);
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/** Abandons an amendment; the working copy resets to the last signed snapshot. */
function ds_discardAmendment(token, summaryId, expectedRowVersion) {
  var lock = LockService.getScriptLock();
  try {
    var actor = dsx_requireRole_(token, 'amend');

    lock.waitLock(DSX_LOCK_MS);
    dsx_resetHeaderCache_();

    var header = dsx_getHeader_(summaryId);
    if (!header) return dsx_err_('VALIDATION_FAILED', 'No discharge summary found for ' + summaryId + '.');
    dsx_checkVersion_(header, expectedRowVersion);

    var status = dsx_upper_(header.Status);
    var t = dsx_transition_('AMEND_DISCARD', status);
    if (!t.ok) return dsx_err_('INVALID_STATE', t.message);

    var lastSigned = dsx_int_(header.Last_Signed_Snapshot_No);
    var snap = dsx_getSnapshotPayload_(summaryId, lastSigned);
    if (!snap || !snap.payload) {
      return dsx_err_('VALIDATION_FAILED', 'The last signed version could not be read.');
    }

    dsx_putWorking_(summaryId, snap.payload, lastSigned, actor.username);
    var rowVersion = dsx_bumpVersion_(dsx_summariesSheet_(), header, actor, { Status: t.to });

    dsx_logEvent_(summaryId, actor, 'DS_AMEND_DISCARD', status, t.to, lastSigned, '', '', {});
    SpreadsheetApp.flush();

    return dsx_ok_('Amendment discarded; the signed version stands.',
                   { rowVersion: rowVersion, status: t.to });

  } catch (e) {
    return dsx_conflictOr_(e);
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

// ---------------------------------------------------------------------------
// SECTION G — SIGNING
// ---------------------------------------------------------------------------

/**
 * Verify and electronically sign.
 *
 * @param {Object} credential   {method:"TOTP"|"PASSWORD", value:"..."}
 * @param {Object} attestations {finalDiagnosis:true, allergies:true, ...}
 * @param {string} onBehalfReason  required when the signer is not the
 *                                 consultant of record
 */
function ds_sign(token, summaryId, expectedRowVersion, credential, attestations, onBehalfReason) {
  var lock = LockService.getScriptLock();
  try {
    var actor = dsx_requireRole_(token, 'sign');

    // ---- 1. signer identity ------------------------------------------------
    if (!actor.doctorId) {
      return dsx_err_('FORBIDDEN',
        'Your login is not linked to a doctor profile. Ask the administrator to set ' +
        'Linked_Username in the Doctors sheet before you sign.');
    }
    var profile = dsx_doctorProfile_(actor.doctorId);
    if (!profile) return dsx_err_('FORBIDDEN', 'Doctor profile ' + actor.doctorId + ' was not found.');
    if (!profile.regNo) {
      return dsx_err_('VALIDATION_FAILED',
        'Your registration number is not on file. A discharge summary must carry the signing ' +
        "doctor's registration number — ask the administrator to fill Reg_No for " +
        profile.name + ' in the Doctors sheet.');
    }

    // ---- 2. sign lockout ---------------------------------------------------
    if (dsx_signLocked_(actor.username)) {
      return dsx_err_('FORBIDDEN',
        'Too many failed signing attempts. Try again in a few minutes.');
    }

    var header0 = dsx_getHeader_(summaryId);
    if (!header0) return dsx_err_('VALIDATION_FAILED', 'No discharge summary found for ' + summaryId + '.');

    var status0 = dsx_upper_(header0.Status);
    var signAction = (status0 === DSX_STATUS.AMENDMENT_IN_PROGRESS) ? 'SIGN_AMEND'
                   : (status0 === DSX_STATUS.PENDING_SIGNATURE) ? 'SIGN' : 'SIGN_FAST';
    if (signAction === 'SIGN_FAST' && actor.cfg.requirePreparerReview) {
      return dsx_err_('INVALID_STATE',
        'This summary is ' + status0 + '. It must be submitted for signature first.');
    }
    var t0 = dsx_transition_(signAction, status0);
    if (!t0.ok) return dsx_err_('INVALID_STATE', t0.message);

    // ---- 3. consultant-of-record rule -------------------------------------
    var cor = dsx_consultantsOfRecord_(dsx_str_(header0.IP_Number));
    var isConsultant = cor.ids.indexOf(dsx_upper_(actor.doctorId)) !== -1;
    var reason = dsx_str_(onBehalfReason);

    if (!isConsultant) {
      if (actor.cfg.signerPolicy !== 'CONSULTANT_OR_ANY_DOCTOR_WITH_REASON') {
        return dsx_err_('FORBIDDEN',
          'Only the consultant of record may sign this discharge summary.');
      }
      if (!reason) {
        return dsx_err_('VALIDATION_FAILED',
          'You are not the consultant of record for this admission. Give a reason for ' +
          'signing on their behalf — it is logged and printed on the summary.');
      }
    } else {
      reason = '';
    }

    // ---- 4. readiness, recomputed server-side -----------------------------
    var working0 = dsx_getWorking_(summaryId);
    if (!working0 || !working0.payload) {
      return dsx_err_('VALIDATION_FAILED', 'The working draft is missing. Regenerate the summary.');
    }
    var readiness = { hard: [], soft: [] };
    if (typeof dsx_readiness_ === 'function') {
      readiness = dsx_readiness_(working0.payload, null, header0);
    }
    if (readiness.hard.length) {
      return dsx_err_('VALIDATION_FAILED',
        'This summary cannot be signed yet: ' +
        readiness.hard.map(function (h) { return h.label; }).join('; ') + '.',
        { readiness: readiness });
    }
    var unaccepted = dsx_unacceptedAiSections_(working0.payload);
    if (unaccepted.length) {
      return dsx_err_('VALIDATION_FAILED',
        'AI-drafted text must be accepted or rejected before signing: ' + unaccepted.join(', ') + '.');
    }

    // ---- 5. credential -----------------------------------------------------
    var cred = dsx_verifyCredential_(actor, credential);
    if (!cred.ok) {
      var left = dsx_recordSignFailure_(actor.username);
      return dsx_err_('FORBIDDEN', cred.message +
        (left > 0 ? ' ' + left + ' attempt(s) remaining.' : ' Signing is now locked for 15 minutes.'));
    }
    dsx_clearSignFailures_(actor.username);

    // ---- 6. hash, computed outside the lock -------------------------------
    var signedAt = new Date();
    var signedAtIso = signedAt.toISOString();
    var prevSignedHash = dsx_str_(header0.Signed_Hash);
    var snapshotNo = dsx_int_(header0.Current_Snapshot_No) + 1;

    var payload = working0.payload;
    payload.sections = payload.sections || {};
    payload.signature = {
      signedBy: actor.username,
      signerName: profile.name,
      signerQualification: profile.signatureLine,
      signerRegNo: profile.regNo,
      signedAt: signedAtIso,
      onBehalfReason: reason,
      preparedBy: dsx_str_(header0.Prepared_By),
      attestations: attestations || {},
      snapshotNo: snapshotNo
    };

    var canonical = dsx_canonicalJson_(payload);
    var basis = canonical + '|' + dsx_upper_(summaryId) + '|' + snapshotNo + '|' +
                actor.username + '|' + signedAtIso + '|' + prevSignedHash;
    var contentHash = dsx_sha256Hex_(basis);

    var verifyToken = dsx_hmacSha256B64Url_(dsx_upper_(summaryId) + ':' + snapshotNo);
    var verifyTokenHash = dsx_sha256Hex_(verifyToken);

    // ---- 7. commit ---------------------------------------------------------
    lock.waitLock(DSX_LOCK_MS);
    dsx_resetHeaderCache_();

    var header = dsx_getHeader_(summaryId);
    dsx_checkVersion_(header, expectedRowVersion);
    var status = dsx_upper_(header.Status);
    var t = dsx_transition_(signAction, status);
    if (!t.ok) return dsx_err_('INVALID_STATE', t.message);

    dsx_appendSnapshot_(summaryId, snapshotNo, 'SIGNED', payload,
                        contentHash, prevSignedHash, actor.username, verifyTokenHash);
    dsx_putWorking_(summaryId, payload, snapshotNo, actor.username);

    var rowVersion = dsx_bumpVersion_(dsx_summariesSheet_(), header, actor, {
      Status: DSX_STATUS.SIGNED,
      Current_Snapshot_No: snapshotNo,
      Last_Signed_Snapshot_No: snapshotNo,
      Clinical_Discharge_At: dsx_toDate_(header.Clinical_Discharge_At) || signedAt,
      Signed_At: signedAt,
      Signed_By: actor.username,
      Signer_Reg_No: profile.regNo,
      Signed_Hash: contentHash,
      Pdf_Status: 'PENDING'
    });

    dsx_logEvent_(summaryId, actor, (signAction === 'SIGN_AMEND') ? 'DS_SIGN_AMENDMENT' : 'DS_SIGN',
                  status, DSX_STATUS.SIGNED, snapshotNo, contentHash,
                  reason ? ('Signed on behalf: ' + reason) : '',
                  {
                    signerName: profile.name,
                    signerRegNo: profile.regNo,
                    credentialMethod: cred.method,
                    onBehalfReason: reason,
                    withoutPreparerReview: (signAction === 'SIGN_FAST'),
                    softAcknowledged: readiness.soft.length,
                    consultantOfRecord: isConsultant,
                    consultantIds: cor.ids,
                    // Empty unless IP_Admissions.Consultant names somebody the
                    // Doctors master does not know. When it is set, a doctor may
                    // have been asked for a reason only because of stale data.
                    unresolvedConsultantName: cor.unresolvedConsultant
                  });
    dsx_audit_(actor, 'DS_SIGN', summaryId, {
      ipNumber: dsx_str_(header.IP_Number),
      snapshotNo: snapshotNo,
      shortHash: dsx_shortHash_(contentHash),
      credential: cred.method,
      onBehalfReason: reason
    });

    SpreadsheetApp.flush();
    // The PDF is deliberately NOT generated here — the lock is released first
    // (Phase 7 does it outside, and a PDF failure must never roll back a
    // signature).

    return dsx_ok_('Signed by ' + profile.name + '.', {
      rowVersion: rowVersion,
      status: DSX_STATUS.SIGNED,
      snapshotNo: snapshotNo,
      shortHash: dsx_shortHash_(contentHash),
      signedAt: dsx_fmt_(signedAt, 'dd-MMM-yyyy hh:mm a'),
      signerName: profile.name,
      signerRegNo: profile.regNo
    });

  } catch (e) {
    return dsx_conflictOr_(e);
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/** Doctor master fields needed to sign and to print. */
function dsx_doctorProfile_(doctorId) {
  var sh = dsx_ss_().getSheetByName('Doctors');
  if (!sh) return null;
  var row = dsx_findRowByKey_(sh, 'Doctor_ID', dsx_upper_(doctorId));
  if (!row) return null;
  var o = dsx_readRow_(sh, row);
  if (dsx_upper_(o.Status) !== 'ACTIVE') return null;
  return {
    doctorId: dsx_str_(o.Doctor_ID),
    name: dsx_str_(o.Display_Name),
    specialty: dsx_str_(o.Specialty),
    regNo: dsx_str_(o.Reg_No),
    signatureLine: dsx_str_(o.Signature_Line),
    username: dsx_str_(o.Linked_Username)
  };
}

/**
 * Every doctor who counts as a consultant of record for this admission, and
 * may therefore sign without giving an on-behalf reason.
 *
 * IP_Admissions carries TWO consultant fields and they disagree on live rows:
 * `Consultant` is a display string ("Dr. Logavignesh") and `Primary_Doctor_ID`
 * is an ID (DOC001). Confirmed with the CEO: BOTH are legitimate — whichever
 * of them a doctor matches, they are signing their own patient's summary, not
 * somebody else's.
 *
 * Resolving the display name can fail (the string may name nobody in the
 * Doctors master). That is reported in `unresolved` rather than swallowed, so
 * the signing log records why a doctor was asked for a reason.
 *
 * @return {{ids:Array<string>, unresolvedConsultant:string}}
 */
function dsx_consultantsOfRecord_(ipNumber) {
  var out = { ids: [], unresolvedConsultant: '' };
  var adm = dsx_admissionRow_(ipNumber);
  if (!adm) return out;

  var add = function (id) {
    var v = dsx_upper_(id);
    if (v && out.ids.indexOf(v) === -1) out.ids.push(v);
  };

  add(adm.Primary_Doctor_ID);

  var consultantName = dsx_str_(adm.Consultant);
  if (consultantName) {
    var resolved = '';
    try {
      if (typeof ipc_doctorIdByName_ === 'function') {
        resolved = dsx_str_(ipc_doctorIdByName_(consultantName));
      }
    } catch (e) { resolved = ''; }
    if (resolved) add(resolved);
    else out.unresolvedConsultant = consultantName;
  }

  return out;
}

/**
 * Signing re-authentication. TOTP when the user has a secret enrolled;
 * password re-entry otherwise, and only while DS_SIGN_ALLOW_PASSWORD is true.
 *
 * Discovery §5: passwords in the Users sheet are CLEAR TEXT today. The
 * password path is therefore a weak credential by construction; it is
 * switchable per tenant so a hospital can insist on TOTP.
 */
function dsx_verifyCredential_(actor, credential) {
  var method = dsx_upper_(credential && credential.method);
  var value = dsx_str_(credential && credential.value);
  if (!value) return { ok: false, message: 'Enter your signing credential.' };

  var user = dsx_userRow_(actor.username);
  if (!user) return { ok: false, message: 'Your account could not be verified.' };

  var hasTotp = !!user.mfaSecret;

  if (method === 'TOTP' || (!method && hasTotp)) {
    if (!hasTotp) {
      return { ok: false, message: 'You have no authenticator enrolled. Sign with your password instead.' };
    }
    // processTOTP returns an ENVELOPE ({success, message, code}), never the
    // boolean true. Comparing it with === true made every TOTP signature fail,
    // so a clinic that had enrolled MFA could not sign a discharge summary at
    // all — and the message blamed the doctor's code for it.
    var res = null;
    try {
      if (typeof processTOTP === 'function') res = processTOTP(user.mfaSecret, value);
    } catch (e) { res = { success: false, message: 'Verification error: ' + e.message }; }

    if (res && res.success === true) return { ok: true, method: 'TOTP' };
    return { ok: false,
             message: (res && res.message) || 'That authenticator code is not valid.' };
  }

  if (!actor.cfg.signAllowPassword) {
    return { ok: false, message: 'This clinic requires an authenticator code to sign. Enrol MFA first.' };
  }
  if (hasTotp) {
    // Someone with an authenticator must use it; falling back would make the
    // stronger factor optional and therefore pointless.
    return { ok: false, message: 'Use your authenticator code to sign.' };
  }
  var match = dsx_timingSafeEqual_(value, user.password);
  return match ? { ok: true, method: 'PASSWORD' }
               : { ok: false, message: 'That password is not correct.' };
}

function dsx_userRow_(username) {
  var sh = dsx_ss_().getSheetByName('Users');
  if (!sh) return null;
  var row = dsx_findRowByKey_(sh, 'Username', username);
  if (!row) return null;
  var o = dsx_readRow_(sh, row);
  if (dsx_upper_(o.Status) !== 'ACTIVE') return null;
  return {
    username: dsx_str_(o.Username),
    password: dsx_str_(o.Password),
    role: dsx_str_(o.Role).toLowerCase(),
    mfaSecret: dsx_str_(o.MFA_Secret)
  };
}

function dsx_signFailKey_(username) { return 'DSSIGNFAIL_' + dsx_upper_(username); }

function dsx_signLocked_(username) {
  try {
    return dsx_int_(CacheService.getScriptCache().get(dsx_signFailKey_(username))) >= DSX_SIGN_FAIL_MAX;
  } catch (e) { return false; }
}

/** @return {number} attempts remaining before lockout */
function dsx_recordSignFailure_(username) {
  try {
    var cache = CacheService.getScriptCache();
    var key = dsx_signFailKey_(username);
    var n = dsx_int_(cache.get(key)) + 1;
    cache.put(key, String(n), DSX_SIGN_FAIL_WINDOW);
    if (n >= DSX_SIGN_FAIL_MAX) {
      dsx_audit_({ username: username, role: '' }, 'DS_SIGN_LOCKOUT', '',
                 { username: username, failures: n });
    }
    return Math.max(0, DSX_SIGN_FAIL_MAX - n);
  } catch (e) { return DSX_SIGN_FAIL_MAX; }
}

function dsx_clearSignFailures_(username) {
  try { CacheService.getScriptCache().remove(dsx_signFailKey_(username)); } catch (e) {}
}

function dsx_unacceptedAiSections_(payload) {
  var out = [];
  var sections = (payload && payload.sections) || {};
  Object.keys(sections).forEach(function (k) {
    var s = sections[k];
    if (s && dsx_upper_(s.origin) === 'AI' && !dsx_str_(s.aiAcceptedBy)) out.push(k);
  });
  return out;
}

// ---------------------------------------------------------------------------
// SECTION H — DIFF AND SOURCE LOOK-THROUGH
// ---------------------------------------------------------------------------

/**
 * Per-section old/new content between two refs. The client renders the diff —
 * the server never ships HTML here.
 *
 * @param {string} fromRef  "BASELINE:n" | "SUBMITTED:n" | "SIGNED:n" | "WORKING"
 *                          or bare "BASELINE" / "SUBMITTED" / "SIGNED" for the latest
 */
function ds_getDiff(token, summaryId, fromRef, toRef) {
  try {
    var actor = dsx_requireRole_(token, ['view', 'viewSigned']);
    var header = dsx_getHeader_(summaryId);
    if (!header) return dsx_err_('VALIDATION_FAILED', 'No discharge summary found for ' + summaryId + '.');

    var a = dsx_resolveRef_(summaryId, fromRef);
    var b = dsx_resolveRef_(summaryId, toRef);
    if (!a.payload) return dsx_err_('VALIDATION_FAILED', 'Cannot read version "' + fromRef + '".');
    if (!b.payload) return dsx_err_('VALIDATION_FAILED', 'Cannot read version "' + toRef + '".');

    var keys = {};
    Object.keys(a.payload.sections || {}).forEach(function (k) { keys[k] = true; });
    Object.keys(b.payload.sections || {}).forEach(function (k) { keys[k] = true; });

    var sections = [];
    Object.keys(keys).forEach(function (k) {
      var sa = (a.payload.sections || {})[k];
      var sb = (b.payload.sections || {})[k];
      var ja = sa ? dsx_canonicalJson_(sa.content) : '';
      var jb = sb ? dsx_canonicalJson_(sb.content) : '';
      sections.push({
        key: k,
        title: (sb && sb.title) || (sa && sa.title) || k,
        format: (sb && sb.format) || (sa && sa.format) || 'TEXT',
        changed: ja !== jb,
        old: sa ? sa.content : null,
        new: sb ? sb.content : null,
        editedBy: sb ? dsx_str_(sb.editedBy) : '',
        editedAt: sb ? dsx_str_(sb.editedAt) : ''
      });
    });

    sections.sort(function (x, y) { return (y.changed ? 1 : 0) - (x.changed ? 1 : 0); });

    return dsx_ok_('', {
      from: a.label, to: b.label,
      changedCount: sections.filter(function (s) { return s.changed; }).length,
      sections: sections
    });
  } catch (e) {
    return dsx_fromError_(e);
  }
}

function dsx_resolveRef_(summaryId, ref) {
  var r = dsx_upper_(ref) || 'WORKING';
  if (r === 'WORKING') {
    var w = dsx_getWorking_(summaryId);
    return { label: 'WORKING', payload: w ? w.payload : null };
  }
  var parts = r.split(':');
  var type = parts[0];
  var no = parts.length > 1 ? dsx_int_(parts[1]) : 0;

  if (!no) {
    var latest = dsx_latestSnapshotOfType_(summaryId, type);
    if (!latest) return { label: r, payload: null };
    no = latest.snapshotNo;
  }
  var snap = dsx_getSnapshotPayload_(summaryId, no);
  return { label: type + ':' + no, payload: snap ? snap.payload : null };
}

/**
 * The original note or result behind a source chip. Read-only, role-checked,
 * and it never returns a row from a different admission.
 *
 * @param {Object} ref {type:"NOTE"|"LAB"|"CASESHEET"|"ADMISSION"|"PATIENT", id, summaryId}
 */
function ds_getSourceItem(token, ref) {
  try {
    var actor = dsx_requireRole_(token, ['view']);
    ref = ref || {};
    var type = dsx_upper_(ref.type);
    var id = dsx_str_(ref.id);
    if (!type || !id) return dsx_err_('VALIDATION_FAILED', 'A source reference is required.');

    var header = dsx_getHeader_(ref.summaryId);
    if (!header) return dsx_err_('VALIDATION_FAILED', 'No discharge summary found.');
    var ip = dsx_ip_(header.IP_Number);

    if (typeof dsx_readSourceItem_ !== 'function') {
      return dsx_err_('NOT_IMPLEMENTED', 'Source look-through needs DS_Assembly.gs.');
    }
    var item = dsx_readSourceItem_(ip, type, id);
    if (!item) return dsx_err_('VALIDATION_FAILED', 'That source record could not be found for this admission.');
    return dsx_ok_('', item);
  } catch (e) {
    return dsx_fromError_(e);
  }
}

// ---------------------------------------------------------------------------
// SECTION I — error plumbing
// ---------------------------------------------------------------------------

/** Converts a thrown version clash into the documented VERSION_CONFLICT envelope. */
function dsx_conflictOr_(e) {
  if (e && e.dsConflict) {
    return {
      success: false,
      code: 'VERSION_CONFLICT',
      message: String(e.message).replace(/^VERSION_CONFLICT:\s*/, ''),
      data: e.dsConflict
    };
  }
  if (e && /Could not obtain lock|Timeout/i.test(String(e.message))) {
    return dsx_err_('LOCK_TIMEOUT', 'The system is busy. Try again in a moment.');
  }
  return dsx_fromError_(e);
}

// ---------------------------------------------------------------------------
// SECTION J — in-memory test of the transition table and matrix
// ---------------------------------------------------------------------------

/** Writes nothing. Validates the state machine and the permission matrix. */
function ds_testTransitions() {
  var lines = [], pass = 0, fail = 0;
  var check = function (label, got, want) {
    var ok = (got === want);
    ok ? pass++ : fail++;
    lines.push((ok ? 'pass  ' : 'FAIL  ') + label + '  got=' + got + ' want=' + want);
  };

  check('initiate from nothing', dsx_transition_('INITIATE', null).to, DSX_STATUS.GENERATED);
  check('initiate from CANCELLED', dsx_transition_('INITIATE', DSX_STATUS.CANCELLED).to, DSX_STATUS.GENERATED);
  check('initiate from SIGNED is refused', dsx_transition_('INITIATE', DSX_STATUS.SIGNED).ok, false);
  check('first save GENERATED -> IN_PREPARATION', dsx_transition_('PREPARE', DSX_STATUS.GENERATED).to, DSX_STATUS.IN_PREPARATION);
  check('first save RETURNED -> IN_PREPARATION', dsx_transition_('PREPARE', DSX_STATUS.RETURNED).to, DSX_STATUS.IN_PREPARATION);
  check('submit from IN_PREPARATION', dsx_transition_('SUBMIT', DSX_STATUS.IN_PREPARATION).to, DSX_STATUS.PENDING_SIGNATURE);
  check('submit from SIGNED is refused', dsx_transition_('SUBMIT', DSX_STATUS.SIGNED).ok, false);
  check('return from PENDING_SIGNATURE', dsx_transition_('RETURN', DSX_STATUS.PENDING_SIGNATURE).to, DSX_STATUS.RETURNED);
  check('return from GENERATED is refused', dsx_transition_('RETURN', DSX_STATUS.GENERATED).ok, false);
  check('sign from PENDING_SIGNATURE', dsx_transition_('SIGN', DSX_STATUS.PENDING_SIGNATURE).to, DSX_STATUS.SIGNED);
  check('sign from GENERATED is refused on the normal path', dsx_transition_('SIGN', DSX_STATUS.GENERATED).ok, false);
  check('fast-path sign from GENERATED exists', dsx_transition_('SIGN_FAST', DSX_STATUS.GENERATED).to, DSX_STATUS.SIGNED);
  check('amend from SIGNED', dsx_transition_('AMEND_START', DSX_STATUS.SIGNED).to, DSX_STATUS.AMENDMENT_IN_PROGRESS);
  check('sign amendment', dsx_transition_('SIGN_AMEND', DSX_STATUS.AMENDMENT_IN_PROGRESS).to, DSX_STATUS.SIGNED);
  check('discard amendment', dsx_transition_('AMEND_DISCARD', DSX_STATUS.AMENDMENT_IN_PROGRESS).to, DSX_STATUS.SIGNED);
  check('cancel from PENDING_SIGNATURE', dsx_transition_('CANCEL', DSX_STATUS.PENDING_SIGNATURE).to, DSX_STATUS.CANCELLED);
  check('cancel a SIGNED summary is refused', dsx_transition_('CANCEL', DSX_STATUS.SIGNED).ok, false);
  check('edit a SIGNED summary is refused', dsx_transition_('EDIT', DSX_STATUS.SIGNED).ok, false);
  check('regenerate a PENDING_SIGNATURE summary is refused', dsx_transition_('REGENERATE', DSX_STATUS.PENDING_SIGNATURE).ok, false);

  // Permission matrix, with the default preparer roles.
  var cfg = { preparerRoles: ['nurse', 'doctor'], requirePreparerReview: true };
  var actorFor = function (role) {
    var actions = (DSX_ROLE_MATRIX[role] || []).slice();
    if (cfg.preparerRoles.indexOf(role) !== -1) {
      DSX_PREPARER_ACTIONS.forEach(function (a) { if (actions.indexOf(a) === -1) actions.push(a); });
    }
    return { role: role, actions: actions, cfg: cfg };
  };
  var has = function (role, action) { return actorFor(role).actions.indexOf(action) !== -1; };

  check('doctor may sign', has('doctor', 'sign'), true);
  check('nurse may not sign', has('nurse', 'sign'), false);
  check('nurse may edit', has('nurse', 'edit'), true);
  check('nurse may submit', has('nurse', 'submit'), true);
  check('nurse may not return', has('nurse', 'return'), false);
  check('nurse may not cancel', has('nurse', 'cancel'), false);
  check('admin may cancel', has('admin', 'cancel'), true);
  check('admin may not edit', has('admin', 'edit'), false);
  check('admin may not sign', has('admin', 'sign'), false);
  check('receptionist sees signed only', has('receptionist', 'view'), false);
  check('receptionist may print final', has('receptionist', 'printFinal'), true);
  check('pharmacist gets meds only', has('pharmacist', 'viewMedsOnly'), true);
  check('pharmacist cannot view drafts', has('pharmacist', 'view'), false);
  check('unknown role gets nothing', (DSX_ROLE_MATRIX['radiographer'] || []).length, 0);

  var report = 'ds_testTransitions: ' + pass + ' passed, ' + fail + ' failed\n' + lines.join('\n');
  Logger.log(report);
  return { success: fail === 0, message: report, data: { pass: pass, fail: fail, lines: lines } };
}
