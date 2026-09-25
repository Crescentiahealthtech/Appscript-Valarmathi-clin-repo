// ============================================================================
// RBAC_Access.gs — Crescentia HealthTech / CresRx
// Giving one person one more thing, or taking one thing away, from the app.
// ----------------------------------------------------------------------------
// RBAC.gs decides; this file is the screen's server. Staff Accounts → Access
// (Admin_Dashboard.html) lists every staff account with its role and its own
// grants and revocations, and saves changes through crescSaveUserAccess().
//
// THE RULES, all enforced here and not in the browser:
//
//   * admin.users to open the screen or change anything on it.
//   * Doctors and administrators are the owner's (admin.users.elevated),
//     exactly as creating, disabling or resetting them already is.
//   * Nobody edits their own access — an administrator who could widen their
//     own account needs no owner, and one who could narrow it could lock the
//     clinic out of its own user management.
//   * The keys in CRESC_OWNER_GRANT_ONLY are handed out by the owner alone.
//   * CRESC_NEVER_GRANT is never granted by anyone.
//   * Only what differs from the role is stored: a grant of something the
//     role already has, or a revocation of something it lacks, is dropped.
//
// Every save writes one STAFF_ACCESS_CHANGED row to Audit_Log with the
// before and after, which the Audit Log viewer already lists and exports.
// ============================================================================

/**
 * The permissions, grouped the way a clinic thinks about them — by the part
 * of the building they open. Every key in CRESC_PERMS except the ones in
 * CRESC_NEVER_GRANT must appear exactly once; crescRbacSelfTest checks.
 */
var CRESC_ACCESS_GROUPS = [
  { key: 'command',  label: 'Command Center',   icon: 'fa-chart-line',
    perms: ['dashboard.read'] },
  { key: 'patients', label: 'Patients',         icon: 'fa-hospital-user',
    perms: ['patient.read', 'patient.write', 'patient.register'] },
  { key: 'appts',    label: 'Appointments',     icon: 'fa-calendar-check',
    perms: ['appointment.read', 'appointment.write', 'appointment.cancel', 'schedule.write'] },
  { key: 'emr',      label: 'Clinical (EMR)',   icon: 'fa-laptop-medical',
    perms: ['emr.read', 'emr.write', 'rx.write', 'reference.read'] },
  { key: 'ward',     label: 'Ward (IP)',        icon: 'fa-bed',
    perms: ['ward.read', 'ward.write', 'ward.admit', 'ward.discharge'] },
  { key: 'lab',      label: 'Laboratory',       icon: 'fa-flask-vial',
    perms: ['lab.read', 'lab.order', 'lab.collect', 'lab.result', 'lab.verify',
            'lab.ack_critical', 'lab.catalog', 'lab.bill'] },
  { key: 'pharmacy', label: 'Pharmacy',         icon: 'fa-pills',
    perms: ['pharmacy.read', 'pharmacy.dispense', 'pharmacy.bill', 'pharmacy.register',
            'pharmacy.stock_add', 'pharmacy.stock_edit', 'pharmacy.stock_discard',
            'pharmacy.return'] },
  { key: 'billing',  label: 'Hospital billing', icon: 'fa-file-invoice-dollar',
    perms: ['billing.read', 'billing.write', 'billing.cancel'] },
  { key: 'accounts', label: 'Accounts',         icon: 'fa-scale-balanced',
    perms: ['accounts.read', 'accounts.write', 'accounts.settle', 'accounts.payables',
            'accounts.tax', 'accounts.lock_period'] },
  { key: 'admin',    label: 'Administration',   icon: 'fa-user-shield',
    perms: ['admin.users', 'admin.audit', 'admin.config', 'dpdp.manage'] }
];

/** Readable role names, for the screen and for refusals. */
var CRESC_ROLE_LABELS = {
  admin: 'Administrator', doctor: 'Doctor', nurse: 'Nurse',
  receptionist: 'Receptionist', pharmacist: 'Pharmacist',
  lab: 'Lab technician', accountant: 'Accountant'
};

/**
 * Roles this screen may move an account between. Doctor is not one of them:
 * a doctor needs a Doctors row (signature, registration number, schedule),
 * which Add Staff creates. Administrator is offered to the owner only.
 */
function cresc_assignableRoles_(elevated) {
  var out = ['nurse', 'receptionist', 'pharmacist', 'lab', 'accountant'];
  if (elevated) out.push('admin');
  return out;
}

/** The Users row for a username: {row, data, head}, or null. */
function cresc_findUserRow_(sh, username) {
  var want = crescStr_(username).toUpperCase();
  if (!want) return null;
  var data = sh.getDataRange().getValues();
  var head = data[0].map(function (h) { return crescStr_(h); });
  for (var i = 1; i < data.length; i++) {
    if (crescStr_(data[i][0]).toUpperCase() === want) {
      return { row: i + 1, data: data[i], head: head };
    }
  }
  return null;
}

/**
 * FRONTEND ENTRY. Everything the Access screen draws.
 *
 * @param {string} sessionToken
 * @return {{success:boolean, message:string, elevated:boolean, you:string,
 *           groups:Array, roleDefaults:Object, roleLabels:Object,
 *           assignableRoles:Array<string>, ownerOnly:Array<string>,
 *           accounts:Array}}
 */
function crescGetAccessOverview(sessionToken) {
  try {
    var actor = crescRequire_(sessionToken, 'admin.users');
    var elevated = actor.permissions.indexOf('admin.users.elevated') !== -1;

    // A screen that grants access reads the sheet as it is now, not as it was
    // cached up to ten minutes ago.
    cresc_aclBust_();
    var map = cresc_aclMap_();
    if (!map || !map.users) {
      return { success: false, message: 'The Users sheet could not be read. Try again in a moment.' };
    }

    var groups = CRESC_ACCESS_GROUPS.map(function (g) {
      return {
        key: g.key, label: g.label, icon: g.icon,
        perms: g.perms.filter(function (p) { return CRESC_PERMS.hasOwnProperty(p); })
          .map(function (p) {
            return { key: p, label: CRESC_PERMS[p],
                     ownerOnly: CRESC_OWNER_GRANT_ONLY.indexOf(p) !== -1 };
          })
      };
    });

    var roleDefaults = {};
    crescStaffRoles_().forEach(function (r) { roleDefaults[r] = crescPermsFor_(r); });

    var me = crescStr_(actor.username).toUpperCase();
    var accounts = Object.keys(map.users).map(function (k) {
      var e = map.users[k];
      var eff = crescEffectivePerms_(e.n, e.r);
      return {
        username: e.n,
        role: e.r,
        roleLabel: CRESC_ROLE_LABELS[e.r] || e.rr || e.r || '(no role)',
        knownRole: !!CRESC_ROLE_MATRIX[e.r],
        active: !!e.a,
        owner: !!e.o,
        granted: eff.granted,
        revoked: eff.revoked,
        effective: eff.perms,
        isSelf: k === me,
        manageable: k !== me && (elevated || !cresc_roleIsElevated_(e.r))
      };
    }).filter(function (a) { return a.role !== 'patient'; });

    accounts.sort(function (a, b) {
      if (a.active !== b.active) return a.active ? -1 : 1;
      return a.username.localeCompare(b.username);
    });

    return {
      success: true, message: '',
      elevated: elevated,
      you: actor.username,
      groups: groups,
      roleDefaults: roleDefaults,
      roleLabels: CRESC_ROLE_LABELS,
      assignableRoles: cresc_assignableRoles_(elevated),
      ownerOnly: CRESC_OWNER_GRANT_ONLY.slice(),
      accounts: accounts
    };
  } catch (err) {
    return { success: false, message: cresc_reason_(err) };
  }
}

/**
 * FRONTEND ENTRY. Saves one person's role and their own access.
 *
 * `change.grant` and `change.revoke` are the COMPLETE lists after the edit,
 * not a delta: the screen sends what it shows, and the server works out what
 * actually differs from the role, so a stale screen cannot leave a stray key
 * behind.
 *
 * @param {string} sessionToken
 * @param {string} username
 * @param {{role?:string, grant:Array<string>, revoke:Array<string>, note?:string}} change
 */
function crescSaveUserAccess(sessionToken, username, change) {
  var lock = LockService.getScriptLock();
  try {
    var actor = crescRequire_(sessionToken, 'admin.users');
    var elevated = actor.permissions.indexOf('admin.users.elevated') !== -1;
    change = change || {};

    var want = crescStr_(username);
    if (!want) return { success: false, message: 'Name the account.' };
    if (want.toUpperCase() === crescStr_(actor.username).toUpperCase()) {
      return { success: false,
               message: 'You cannot change your own access. Ask another administrator, ' +
                        'or the system owner.' };
    }

    lock.waitLock(20000);
    var sh = cresc_usersSheet_();
    var hit = cresc_findUserRow_(sh, want);
    if (!hit) return { success: false, message: 'No staff account with the ID ' + want + '.' };

    var oldRole = crescRole_(hit.data[2]);
    if (oldRole === 'patient') {
      return { success: false, message: 'Patient portal logins are not managed here.' };
    }
    if (cresc_roleIsElevated_(oldRole) && !elevated) {
      return cresc_elevationRefusal_('change the access of', oldRole);
    }

    // ---- the role -------------------------------------------------------
    var newRole = oldRole;
    var askedRole = crescStr_(change.role) ? crescRole_(change.role) : '';
    if (askedRole && askedRole !== oldRole) {
      if (oldRole === 'doctor') {
        return { success: false,
                 message: 'A doctor\'s role is not changed here: their prescriptions and ' +
                          'notes are signed as a doctor. Switch the account off and add the ' +
                          'person again in the role they now hold.' };
      }
      if (cresc_assignableRoles_(elevated).indexOf(askedRole) === -1) {
        return askedRole === 'admin' || askedRole === 'doctor'
          ? cresc_elevationRefusal_('make somebody', askedRole)
          : { success: false, message: '"' + change.role + '" is not a role this screen can assign.' };
      }
      newRole = askedRole;
    }

    // ---- the lists --------------------------------------------------------
    var base = crescPermsFor_(newRole);
    var grant = cresc_permList_((change.grant || []).join(','))
      .filter(function (p) { return base.indexOf(p) === -1; });
    var revoke = cresc_permList_((change.revoke || []).join(','))
      .filter(function (p) { return base.indexOf(p) !== -1; });

    var never = grant.filter(function (p) { return CRESC_NEVER_GRANT.indexOf(p) !== -1; });
    if (never.length) {
      return { success: false,
               message: '"' + CRESC_PERMS[never[0]] + '" cannot be granted to a person. ' +
                        (never[0] === 'admin.users.elevated'
                          ? 'The system owner is marked on the Users sheet (Super_Admin).'
                          : '') };
    }

    // An owner-only key the account ALREADY holds may stay; the rule is about
    // who can hand it out, not about taking back what the owner gave.
    var current = cresc_permList_(hit.head.indexOf(CRESC_GRANT_HEADER) === -1 ? ''
                                  : hit.data[hit.head.indexOf(CRESC_GRANT_HEADER)]);
    if (!elevated) {
      var needsOwner = grant.filter(function (p) {
        return CRESC_OWNER_GRANT_ONLY.indexOf(p) !== -1 && current.indexOf(p) === -1;
      });
      if (needsOwner.length) {
        return { success: false, code: 'ELEVATION_REQUIRED',
                 message: 'Only the system owner may give somebody "' +
                          CRESC_PERMS[needsOwner[0]].toLowerCase() + '".' };
      }
    }

    // ---- write ----------------------------------------------------------
    var gCol = dc_ensureColumn_(sh, CRESC_GRANT_HEADER);
    var xCol = dc_ensureColumn_(sh, CRESC_REVOKE_HEADER);
    var before = {
      role: crescStr_(hit.data[2]),
      grant: current,
      revoke: cresc_permList_(hit.head.indexOf(CRESC_REVOKE_HEADER) === -1 ? ''
                              : hit.data[hit.head.indexOf(CRESC_REVOKE_HEADER)])
    };

    if (newRole !== oldRole) sh.getRange(hit.row, 3).setValue(newRole);
    sh.getRange(hit.row, gCol + 1).setNumberFormat('@').setValue(grant.join(', '));
    sh.getRange(hit.row, xCol + 1).setNumberFormat('@').setValue(revoke.join(', '));
    SpreadsheetApp.flush();
    dc_invalidate_('Users');
    cresc_aclBust_();

    // A new role is a new person as far as every screen is concerned: their
    // menus, their landing page, what the desk offers them. Their open
    // sessions end so the next sign-in draws all of it fresh.
    var ended = 0;
    if (newRole !== oldRole) {
      try { ended = ds_revokeUserSessions_(want, ''); } catch (e) {}
    }

    var after = { role: newRole, grant: grant, revoke: revoke };
    try {
      logAudit_({ username: actor.username, role: actor.role, doctorId: actor.doctorId },
                'STAFF_ACCESS_CHANGED', 'User', want.toUpperCase(),
                { before: before, after: after, note: crescStr_(change.note).slice(0, 300),
                  sessionsEnded: ended });
    } catch (e) {}

    var eff = crescEffectivePerms_(want, newRole);
    var summary = [];
    if (newRole !== oldRole) {
      summary.push('role is now ' + (CRESC_ROLE_LABELS[newRole] || newRole) +
                   ' (they will be asked to sign in again)');
    }
    if (grant.length) summary.push(grant.length + ' extra permission(s)');
    if (revoke.length) summary.push(revoke.length + ' withdrawn');
    return {
      success: true,
      message: 'Access for ' + want + ' saved' +
               (summary.length ? ': ' + summary.join(', ') + '.' : ' — exactly their role\'s defaults.'),
      account: { username: want, role: newRole, granted: eff.granted,
                 revoked: eff.revoked, effective: eff.perms }
    };
  } catch (err) {
    return { success: false, message: cresc_reason_(err) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
