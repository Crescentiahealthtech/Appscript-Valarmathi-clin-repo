/* Per-person access, exercised against a pretend spreadsheet.

   RBAC.gs SECTION B2 and RBAC_Access.gs decide who may do what when a role
   is not the whole story: a person's own grants and revocations, an account
   switched off while signed in, a role changed while signed in, and the
   rules about who may change whose access. This loads the real RBAC.gs,
   RBAC_Access.gs and Doctor_Session_Store.gs into one V8 context (Apps
   Script's model: one global scope), gives them a fake Users sheet and a
   fake cache, and walks each rule.

       node tools/access.js                 (part of ./tools/check.sh)  */
const fs = require('fs'), path = require('path'), vm = require('vm'), crypto = require('crypto');
const ROOT = process.env.REPO || path.resolve(__dirname, '..');

function makeSheet(name, rows) {
  const data = rows.map(r => r.slice());
  const width = () => Math.max(0, ...data.map(r => r.length));
  const cell = (r, c) => (data[r - 1] && data[r - 1][c - 1] !== undefined) ? data[r - 1][c - 1] : '';
  const set = (r, c, v) => { while (data.length < r) data.push([]); const row = data[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v; };
  const range = (r, c, nr, nc) => ({
    getValue: () => cell(r, c),
    getValues: () => Array.from({ length: nr || 1 }, (_, i) => Array.from({ length: nc || 1 }, (_, j) => cell(r + i, c + j))),
    getDisplayValues() { return this.getValues().map(row => row.map(v => v instanceof Date ? v.toISOString() : String(v))); },
    setValue(v) { set(r, c, v); return this; },
    setValues(vs) { vs.forEach((row, i) => row.forEach((v, j) => set(r + i, c + j, v))); return this; },
    setNumberFormat() { return this; }, setFontWeight() { return this; }, setBackground() { return this; }
  });
  return {
    data, getName: () => name,
    getLastRow: () => data.length,
    getLastColumn: () => width(),
    getRange: (r, c, nr, nc) => range(r, c, nr, nc),
    getDataRange: () => ({
      getValues: () => data.map(r => { const o = r.slice(); while (o.length < width()) o.push(''); return o; }),
      getDisplayValues: () => data.map(r => { const o = r.map(v => v instanceof Date ? v.toISOString() : String(v)); while (o.length < width()) o.push(''); return o; })
    }),
    appendRow(r) { data.push(r.slice()); },
    setFrozenRows() {}
  };
}

const cache = new Map();
const audit = [];
const env = {
  console, JSON, Date, Math,
  Utilities: { getUuid: () => crypto.randomUUID(), formatDate: () => '2026-09-25 10:00' },
  CacheService: { getScriptCache: () => ({
    get: k => cache.has(k) ? cache.get(k) : null,
    put: (k, v) => cache.set(k, String(v)),
    remove: k => cache.delete(k)
  }) },
  LockService: { getScriptLock: () => ({ waitLock() {}, tryLock() { return true; }, releaseLock() {} }) },
  Logger: { log() {} },
  Session: { getScriptTimeZone: () => 'Asia/Kolkata',
             getActiveUser: () => ({ getEmail: () => '' }), getEffectiveUser: () => ({ getEmail: () => 'owner@x' }) }
};
vm.createContext(env);

// What these files call in other files, reduced to what they do here.
vm.runInContext(`
  function dc_str_(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }
  function dc_upper_(v) { return dc_str_(v).toUpperCase(); }
  function dc_headerMap_(sh) {
    var h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0], m = {};
    h.forEach(function (x, i) { x = String(x || '').trim(); if (x && m[x] === undefined) m[x] = i; });
    return m;
  }
  function dc_ensureColumn_(sh, name) {
    var m = dc_headerMap_(sh); if (m[name] !== undefined) return m[name];
    var c = sh.getLastColumn() + 1; sh.getRange(1, c).setValue(name); return c - 1;
  }
  function dc_invalidate_(n) { if (n === 'Users' && typeof cresc_aclBust_ === 'function') cresc_aclBust_(); }
  function dc_ensureSheet_(ss, name, headers) {
    var sh = ss.getSheetByName(name); if (sh) return sh;
    return ss.insertSheet(name, headers);
  }
  function dc_sessionName_(s) { return dc_str_(s.displayName) || dc_str_(s.name) || dc_str_(s.username); }
  function validateSession_(t) { var raw = CacheService.getScriptCache().get('SESS_' + t); return raw ? JSON.parse(raw) : null; }
  function cresc_usersSheet_() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users'); }
  var CRESC_ELEVATED_ROLES = ['doctor', 'admin'];
  function cresc_roleIsElevated_(r) { return CRESC_ELEVATED_ROLES.indexOf(String(r || '').toLowerCase()) !== -1; }
  function crescStaffRoles_() { return Object.keys(CRESC_ROLE_MATRIX).filter(function (r) { return r !== 'patient'; }); }
  function cresc_elevationRefusal_(what, role) { return { success: false, code: 'ELEVATION_REQUIRED', message: 'owner only: ' + what + ' ' + role }; }
  function logAudit_(a, ev, t, id, d) { __audit.push({ by: a.username, ev: ev, id: id, d: d }); }
`, Object.assign(env, { __audit: audit }));
for (const f of ['RBAC.gs', 'RBAC_Access.gs', 'Doctor_Session_Store.gs']) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), env, { filename: f });
}

let SS;
env.SpreadsheetApp = { getActiveSpreadsheet: () => SS, flush() {} };
function fresh() {
  cache.clear(); audit.length = 0;
  const sheets = {
    Users: makeSheet('Users', [
      ['Username', 'Password', 'Role', 'Status', 'Email Address', 'MFA_Secret', 'Must_Change', 'Password_Updated_At', 'Super_Admin', 'Access_Grant', 'Access_Revoke'],
      ['owner',  'x', 'admin',        'ACTIVE',   '', '', '', '', 'YES', '', ''],
      ['admin2', 'x', 'admin',        'ACTIVE',   '', '', '', '', '',    '', ''],
      ['doc1',   'x', 'doctor',       'ACTIVE',   '', '', '', '', '',    '', ''],
      ['nurse1', 'x', 'nurse',        'ACTIVE',   '', '', '', '', '',    'lab.bill, admin.users.elevated, not.a.key', 'emr.write'],
      ['recep1', 'x', 'Receptionist', 'ACTIVE',   '', '', '', '', '',    '', ''],
      ['pharm1', 'x', 'pharmacy',     'ACTIVE',   '', '', '', '', '',    '', ''],
      ['lab1',   'x', 'lab',          'ACTIVE',   '', '', '', '', '',    '', ''],
      ['gone1',  'x', 'nurse',        'INACTIVE', '', '', '', '', '',    '', '']
    ]),
    Sessions: makeSheet('Sessions', [['Token', 'Username', 'Role', 'Doctor_ID', 'Display_Name', 'Issued_At', 'Last_Seen', 'Expires_At', 'Status']])
  };
  SS = {
    sheets,
    getSheetByName: n => sheets[n] || null,
    insertSheet: (n, h) => (sheets[n] = makeSheet(n, [h || []]))
  };
}
/** One Apps Script execution: globals start fresh every call. */
function run(code) {
  vm.runInContext('CRESC_CURRENT_ACTOR = null; CRESC_ACL_MEMO = null;', env);
  return vm.runInContext(code, env);
}
function login(username, role) {
  const token = 'T-' + username + '-' + crypto.randomUUID().slice(0, 6);
  cache.set('SESS_' + token, JSON.stringify({ username, role, doctorId: '', name: username }));
  env.SS_sessions = SS.sheets.Sessions;
  SS.sheets.Sessions.appendRow([token, username, role, '', username, new Date(), new Date(),
                                new Date(Date.now() + 8 * 3600e3), 'ACTIVE']);
  return token;
}
const perms = token => run(`(function(){ var a = crescActor_(${JSON.stringify(token)}); return a ? a.permissions : null; })()`);

let pass = 0; const fails = [];
function check(label, got, want) {
  if (JSON.stringify(got) === JSON.stringify(want)) pass++;
  else fails.push(label + ': got ' + JSON.stringify(got) + ', wanted ' + JSON.stringify(want));
}
const has = (list, p) => !!(list && list.indexOf(p) !== -1);

// ---- 1. The role matrix -----------------------------------------------------
fresh();
check('Command Center: admin', run(`crescCan_('admin','dashboard.read')`), true);
check('Command Center: doctor', run(`crescCan_('doctor','dashboard.read')`), true);
['nurse', 'receptionist', 'pharmacist', 'lab', 'accountant'].forEach(r =>
  check('Command Center withheld from ' + r, run(`crescCan_('${r}','dashboard.read')`), false));
check('a pharmacist cannot take hospital payments', run(`crescCan_('pharmacist','billing.write')`), false);
check('a pharmacist bills the pharmacy', run(`crescCan_('pharmacist','pharmacy.bill')`), true);
check('a lab tech cannot bill the pharmacy', run(`crescCan_('lab','pharmacy.bill')`), false);
check('a lab tech bills the lab', run(`crescCan_('lab','lab.bill')`), true);
check('only admin and accounts void bills', ['admin', 'accountant', 'receptionist', 'lab', 'pharmacist']
  .map(r => run(`crescCan_('${r}','billing.cancel')`)), [true, true, false, false, false]);
check('the owner key is not in admin\'s wildcard', run(`crescCan_('admin','admin.users.elevated')`), false);

// ---- 2. A person's own grants and revocations -------------------------------
fresh();
let t = login('nurse1', 'nurse');
let p = perms(t);
check('a grant adds to the role', has(p, 'lab.bill'), true);
check('a revocation removes from the role', has(p, 'emr.write'), false);
check('the owner key can never be granted', has(p, 'admin.users.elevated'), false);
check('an unknown key grants nothing', has(p, 'not.a.key'), false);
check('the rest of the role is untouched', has(p, 'ward.write'), true);

// ---- 3. The account as it is now --------------------------------------------
fresh();
t = login('recep1', 'receptionist');
check('an active account is let in', !!perms(t), true);
SS.sheets.Users.data[5][3] = 'INACTIVE';
run('cresc_aclBust_()');
check('switched off while signed in: refused on the next call', perms(t), null);
check('...and its sessions are ended',
  SS.sheets.Sessions.data.filter(r => r[1] === 'recep1').map(r => r[8]), ['REVOKED']);
check('...and a fresh cached session cannot revive it', run(`dc_validateSession_(${JSON.stringify(t)})`), null);

fresh();
t = login('recep1', 'receptionist');
SS.sheets.Users.data[5][2] = 'accountant';
run('cresc_aclBust_()');
p = perms(t);
check('a role changed while signed in applies on the next call', has(p, 'accounts.read'), true);
check('...and the old role\'s keys go', has(p, 'patient.register'), false);

fresh();
check('a login with no Users row is refused', perms(login('ghost', 'admin')), null);
check('the editor\'s own job session still works', has(perms(login('SCRIPT_OWNER', 'admin')), 'admin.config'), true);
check('a patient is judged on the portal key alone', perms(login('LMTVS0001', 'patient')), ['portal.self']);
check('a row added by hand is found without waiting for the cache', (() => {
  const tok = login('late1', 'nurse');
  perms(login('recep1', 'receptionist'));                 // warms the cache without late1
  SS.sheets.Users.appendRow(['late1', 'x', 'nurse', 'ACTIVE', '', '', '', '', '', '', '']);
  return has(perms(tok), 'emr.read');
})(), true);

// ---- 4. The owner ------------------------------------------------------------
fresh();
check('the owner holds the owner key', has(perms(login('owner', 'admin')), 'admin.users.elevated'), true);
check('another admin does not', has(perms(login('admin2', 'admin')), 'admin.users.elevated'), false);

// ---- 5. Who may change whose access ------------------------------------------
fresh();
const owner = login('owner', 'admin'), admin2 = login('admin2', 'admin');
const save = (tok, user, change) => run(`crescSaveUserAccess(${JSON.stringify(tok)}, ${JSON.stringify(user)}, ${JSON.stringify(change)})`);

check('nobody changes their own access', save(admin2, 'admin2', { grant: ['dpdp.manage'] }).success, false);
check('an admin cannot change another admin', save(admin2, 'owner', { revoke: ['admin.users'] }).code, 'ELEVATION_REQUIRED');
check('an admin cannot change a doctor', save(admin2, 'doc1', { grant: ['lab.bill'] }).code, 'ELEVATION_REQUIRED');
check('an admin cannot hand out an owner-only key', save(admin2, 'recep1', { grant: ['admin.audit'] }).code, 'ELEVATION_REQUIRED');
check('an admin cannot make somebody an admin', save(admin2, 'recep1', { role: 'admin', grant: [], revoke: [] }).code, 'ELEVATION_REQUIRED');

let r = save(admin2, 'recep1', { grant: ['lab.bill', 'lab.read', 'billing.read'], revoke: ['appointment.cancel', 'lab.read'] });
check('an admin grants and revokes ordinary keys', r.success, true);
check('only the difference from the role is stored', [SS.sheets.Users.data[5][9], SS.sheets.Users.data[5][10]],
      ['lab.bill, lab.read', 'appointment.cancel']);
check('the change is audited with before and after', audit.filter(a => a.ev === 'STAFF_ACCESS_CHANGED').length, 1);
p = perms(login('recep1', 'receptionist'));
check('the grant is in force at once', [has(p, 'lab.bill'), has(p, 'appointment.cancel')], [true, false]);

check('the owner may hand out an owner-only key', save(owner, 'recep1', { grant: ['admin.audit'], revoke: [] }).success, true);
check('the owner key itself is never grantable', save(owner, 'recep1', { grant: ['admin.users.elevated'] }).success, false);

const nurseTok = login('nurse1', 'nurse');
r = save(admin2, 'nurse1', { role: 'lab', grant: [], revoke: [] });
check('a role change is saved', [r.success, SS.sheets.Users.data[4][2]], [true, 'lab']);
check('...and signs the person out', perms(nurseTok), null);
check('a doctor\'s role is not changed here', save(owner, 'doc1', { role: 'nurse' }).success, false);

// ---- 6. The overview -----------------------------------------------------------
fresh();
const ov = run(`crescGetAccessOverview(${JSON.stringify(login('admin2', 'admin'))})`);
check('the overview lists every staff account', ov.accounts.length, 8);
check('the overview marks what this admin may change',
  ['owner', 'doc1', 'recep1', 'admin2'].map(u => ov.accounts.find(a => a.username === u).manageable), [false, false, true, false]);
check('every grantable key is in exactly one access group', (() => {
  const listed = [].concat(...ov.groups.map(g => g.perms.map(x => x.key)));
  const all = run(`Object.keys(CRESC_PERMS).filter(function (p) { return CRESC_NEVER_GRANT.indexOf(p) === -1; })`);
  return listed.length === new Set(listed).size && all.every(k => listed.indexOf(k) !== -1);
})(), true);

if (fails.length) {
  fails.forEach(f => console.log('FAIL  ' + f));
  console.log(pass + ' passed, ' + fails.length + ' failed.');
  process.exit(1);
}
console.log(pass + ' access checks passed (role matrix, grants and revocations, accounts switched off or re-roled mid-session, owner-only keys, who may change whose access).');
