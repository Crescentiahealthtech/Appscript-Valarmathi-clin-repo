/* The sign-in flows, exercised against a pretend spreadsheet.

   credtest.js checks the password maths. This checks the paths around it,
   which is where the two worst sign-in bugs in this project lived:

     * verifyLogin issued a session BEFORE the second factor, so a password
       alone was a full sign-in for anyone who read the reply;
     * a password reset overwrote the real password the moment anyone asked,
       so knowing a colleague's user id was enough to lock them out.

   It loads the real AuthLogin.gs, Auth_Credentials.gs and Auth_Reset.gs into
   one V8 context (Apps Script's model: one global scope), gives them a fake
   Users and Patients sheet, and walks each flow.

       node tools/authflow.js                 (part of ./tools/check.sh)  */
const fs = require('fs'), path = require('path'), vm = require('vm'), crypto = require('crypto');
const ROOT = process.env.REPO || path.resolve(__dirname, '..');

// --- a spreadsheet, as far as these files use one ---------------------------
function makeSheet(rows) {
  const data = rows.map(r => r.slice());
  const width = () => Math.max(...data.map(r => r.length));
  const cell = (r, c) => (data[r - 1] && data[r - 1][c - 1] !== undefined) ? data[r - 1][c - 1] : '';
  const set = (r, c, v) => { while (data.length < r) data.push([]); const row = data[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v; };
  const range = (r, c, nr, nc) => ({
    getValue: () => cell(r, c),
    getDisplayValue: () => { const v = cell(r, c); return v instanceof Date ? v.toISOString() : String(v); },
    getValues: () => Array.from({ length: nr || 1 }, (_, i) => Array.from({ length: nc || 1 }, (_, j) => cell(r + i, c + j))),
    setValue(v) { set(r, c, v); return this; },
    setValues(vs) { vs.forEach((row, i) => row.forEach((v, j) => set(r + i, c + j, v))); return this; },
    setNumberFormat() { return this; }, setFontWeight() { return this; }, setBackground() { return this; }
  });
  return {
    data,
    getLastRow: () => data.length,
    getLastColumn: () => width(),
    getRange: (r, c, nr, nc) => range(r, c, nr, nc),
    getDataRange: () => ({ getValues: () => data.map(r => { const o = r.slice(); while (o.length < width()) o.push(''); return o; }),
                           getDisplayValues: () => data.map(r => r.map(String)) }),
    appendRow(r) { data.push(r.slice()); }
  };
}

const cacheStore = new Map();
const mail = [];
const audit = [];
const env = {
  console,
  Utilities: {
    computeHmacSha256Signature(v, k) {
      const b = x => Array.isArray(x) ? Buffer.from(x.map(n => n & 0xff)) : Buffer.from(String(x), 'utf8');
      return Array.from(crypto.createHmac('sha256', b(k)).update(b(v)).digest());
    },
    computeHmacSignature(_a, v, k) {
      const b = x => Buffer.from(x.map(n => n & 0xff));
      return Array.from(crypto.createHmac('sha1', b(k)).update(b(v)).digest()).map(n => (n << 24) >> 24);
    },
    MacAlgorithm: { HMAC_SHA_1: 'sha1' },
    computeDigest: (_a, v) => Array.from(crypto.createHash('sha256').update(String(v)).digest()),
    DigestAlgorithm: { SHA_256: 'SHA_256' },
    base64Encode: b => Buffer.from(b.map(n => n & 0xff)).toString('base64'),
    getUuid: () => crypto.randomUUID(),
    newBlob: s => ({ getBytes: () => Array.from(Buffer.from(String(s), 'utf8')) }),
    formatDate: () => '2026-09-23 10:00'
  },
  PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty() {} }) },
  CacheService: { getScriptCache: () => ({
    get: k => cacheStore.has(k) ? cacheStore.get(k) : null,
    put: (k, v) => cacheStore.set(k, String(v)),
    remove: k => cacheStore.delete(k)
  }) },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  Logger: { log() {} },
  Session: { getScriptTimeZone: () => 'Asia/Kolkata' },
  GmailApp: { sendEmail: (to, subject, body) => mail.push({ to, subject, body }) },
  UrlFetchApp: { fetch() { throw new Error('no network in tests'); } }
};
vm.createContext(env);
for (const f of ['Auth_Credentials.gs', 'Auth_Reset.gs', 'AuthLogin.gs', 'Auth_MFA_Admin.gs', 'DS_QR_Lib.gs']) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), env, { filename: f });
}
// What the flows call in other files, reduced to what they do here.
vm.runInContext(`
  function ensureColumn_(sheet, name) {
    const h = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const i = h.indexOf(name); if (i !== -1) return i + 1;
    sheet.getRange(1, h.length + 1).setValue(name); return h.length + 1;
  }
  var CRESC_AUTH_EVENTS = new Proxy({}, { get: (t, k) => String(k) });
  function crescAuthGuard_() { return null; }
  function crescAuthFailed_() { return ''; }
  function crescAuthPassed_() {}
  function crescAuthAudit_(e, u) { __audit.push(e + ' ' + u); }
  function resolveDoctorByUsername_() { return null; }
  function issueSession_(o) { const t = 'SESS-' + Utilities.getUuid(); CacheService.getScriptCache().put('SESS_' + t, JSON.stringify(o)); return t; }
  function cresc_clinic_() { return { name: 'Test Clinic' }; }
  // Sessions for the in-app MFA set-up: token -> who is asking.
  var __actors = {
    'T-NURSE': { username: 'nurse1', role: 'nurse', permissions: ['emr.read'] },
    'T-ADMIN': { username: 'admin1', role: 'admin', permissions: ['admin.users'] },
    'T-OWNER': { username: 'owner', role: 'admin', permissions: ['admin.users', 'admin.users.elevated'] }
  };
  function crescRequire_(token, perm) {
    var a = __actors[token];
    if (!a) throw new Error('FORBIDDEN: your session has expired. Please sign in again.');
    if (perm && a.permissions.indexOf(perm) === -1) throw new Error('FORBIDDEN: you may not do that.');
    return a;
  }
  function crescRole_(r) { return String(r || '').toLowerCase(); }
`, Object.assign(env, { __audit: audit }));

let SS;
env.SpreadsheetApp = { getActiveSpreadsheet: () => SS, flush() {} };
function fresh(mfaSecret) {
  const enc = p => env.crescPwdEncode_(p);
  SS = {
    sheets: {
      Users: makeSheet([
        ['Username', 'Password', 'Role', 'Status', 'Email', 'MFA_Secret'],
        ['nurse1', enc('Old-password-123'), 'nurse', 'Active', 'nurse1@example.com', mfaSecret || ''],
        ['drA', enc('Doctor-pass-321'), 'doctor', 'Active', '', '']
      ]),
      Patients: makeSheet([
        Array.from({ length: 17 }, (_, i) => ['Patient_ID', 'Password', 'Name'][i] || 'C' + i),
        Object.assign(Array(17).fill(''), { 0: 'LMTVS0001', 1: enc('Portal-pass-456'), 2: 'Test', 16: 'pt@example.com' })
      ])
    },
    getSheetByName(n) { return this.sheets[n] || null; }
  };
  cacheStore.clear(); mail.length = 0;
}
const tempFrom = () => { const m = /Temporary password:\s+(\S+)/.exec((mail[mail.length - 1] || {}).body || ''); return m && m[1]; };

let failed = 0, ran = 0;
function check(what, ok, detail) { ran++; if (!ok) { failed++; console.log('  FAIL  ' + what + (detail ? '\n        ' + JSON.stringify(detail) : '')); } }

// --- 1. a reset does not lock anybody out -----------------------------------
fresh();
let r = env.crescRequestPasswordReset({ username: 'nurse1' });
const temp = tempFrom();
check('reset request answers success', r.success === true, r);
check('reset mails a temporary password', !!temp, mail);
r = env.verifyLogin({ username: 'nurse1', password: 'Old-password-123' });
check('the OLD password still signs in after a reset request', r.success === true && !!r.sessionToken, r);
r = env.verifyLogin({ username: 'nurse1', password: temp });
check('signing in with the old password cancelled the temporary one', r.success === false && r.code !== 'MUST_CHANGE', r);

// --- 2. the temporary password leads to a forced change ---------------------
fresh();
env.crescRequestPasswordReset({ username: 'nurse1' });
const t2 = tempFrom();
r = env.verifyLogin({ username: 'nurse1', password: t2 });
check('the temporary password asks for a new one', r.success === false && r.code === 'MUST_CHANGE', r);
check('...and issues no session', !r.sessionToken, r);
r = env.crescChangePassword({ username: 'nurse1', currentPassword: t2, newPassword: 'Brand-new-pass-789' });
check('the temporary password can set a new password', r.success === true, r);
r = env.verifyLogin({ username: 'nurse1', password: 'Brand-new-pass-789' });
check('the new password signs in', r.success === true && !!r.sessionToken, r);
r = env.verifyLogin({ username: 'nurse1', password: 'Old-password-123' });
check('the old password stops once the temporary one was used', r.success === false, r);
r = env.verifyLogin({ username: 'nurse1', password: t2 });
check('the temporary password works once only', r.success === false && r.code !== 'MUST_CHANGE', r);

// --- 3. patients get the same treatment --------------------------------------
fresh();
env.crescRequestPasswordReset({ username: 'LMTVS0001' });
const t3 = tempFrom();
r = env.verifyLogin({ username: 'LMTVS0001', password: 'Portal-pass-456' });
check('a patient\'s old password survives a reset request', r.success === true, r);
env.crescRequestPasswordReset({ username: 'LMTVS0001' });
const t4 = tempFrom();
r = env.verifyLogin({ username: 'LMTVS0001', password: t4 });
check('a patient\'s temporary password asks for a new one', r.code === 'MUST_CHANGE', r);
check('an older temporary password was replaced by the newer', t3 !== t4);

// --- 4. MFA: no session until the code is right ------------------------------
const SECRET = 'JBSWY3DPEHPK3PXPJBSWY3DPEHPK3PXP';
fresh(SECRET);
r = env.verifyLogin({ username: 'nurse1', password: 'Old-password-123' });
check('a password alone does NOT return a session when MFA is enrolled', r.success === true && r.mfaRequired === true && !r.sessionToken, r);
const ticket = r.mfaTicket;
let m = env.verifyMFA('nurse1', '000000', ticket);
check('a wrong code returns no session', !m.sessionToken, m);
m = env.verifyMFA('nurse1', '123456', 'not-a-ticket');
check('a code without the password step is refused', m.success === false && m.code === 'TICKET_EXPIRED', m);
const code = env.generateTOTPAlgorithm_(env.base32ToBytes_(SECRET), Math.floor(Date.now() / 30000));
m = env.verifyMFA('nurse1', code, ticket);
check('the right code with the ticket returns a session', m.success === true && !!m.sessionToken, m);
m = env.verifyMFA('nurse1', code, ticket);
check('a ticket works once', !m.sessionToken, m);
fresh();
r = env.verifyLogin({ username: 'nurse1', password: 'Old-password-123' });
check('no MFA enrolled: the password signs straight in', r.success === true && !!r.sessionToken && !r.mfaRequired, r);

// --- 5. Google sign-in does not take the browser's word ----------------------
r = env.verifyGoogleLogin('nurse1@example.com');
check('an email address alone is not a Google sign-in', r.success === false && !r.sessionToken, r);

// --- 6. MFA set up in the app: saved only once the phone proves it ----------
const codeFor = sec => env.generateTOTPAlgorithm_(env.base32ToBytes_(sec.replace(/\s/g, '')), Math.floor(Date.now() / 30000));
const users = () => SS.sheets.Users.data;
fresh();
let b = env.mfaBeginEnrolment({}, 'T-NURSE');
check('your own set-up asks for your password', b.success === false && b.code === 'BAD_PASSWORD', b);
b = env.mfaBeginEnrolment({ currentPassword: 'Old-password-123' }, 'T-NURSE');
check('your own set-up starts with the right password', b.success === true && !!b.ticket && /^otpauth:\/\/totp\//.test(b.otpauthUrl), b);
check('the QR is drawn on the server', /^data:image\/gif;base64,/.test(b.qr || ''), (b.qr || '').slice(0, 40));
check('nothing is saved before the phone confirms', !users()[1][5]);
let c = env.mfaConfirmEnrolment(b.ticket, '000000', 'T-NURSE');
check('a wrong code saves nothing', c.success === false && !users()[1][5], c);
c = env.mfaConfirmEnrolment(b.ticket, codeFor(b.secret), 'T-ADMIN');
check('another session cannot finish your set-up', c.success === false && !users()[1][5], c);
c = env.mfaConfirmEnrolment(b.ticket, codeFor(b.secret), 'T-NURSE');
check('the phone\'s code saves the secret', c.success === true && users()[1][5] === b.secret.replace(/\s/g, ''), c);
r = env.verifyLogin({ username: 'nurse1', password: 'Old-password-123' });
check('...and the next sign-in asks for the code', r.mfaRequired === true && !r.sessionToken, r);
let x = env.mfaConfirmEnrolment(b.ticket, codeFor(b.secret), 'T-NURSE');
check('the enrolment ticket is single use', x.success === false && x.code === 'EXPIRED', x);
b = env.mfaBeginEnrolment({ username: 'drA' }, 'T-NURSE');
check('a nurse cannot set up somebody else', b.success === false, b);
b = env.mfaBeginEnrolment({ username: 'drA' }, 'T-ADMIN');
check('an ordinary admin cannot set up a doctor', b.success === false && b.code === 'ELEVATION_REQUIRED', b);
b = env.mfaBeginEnrolment({ username: 'drA' }, 'T-OWNER');
check('the owner can set up a doctor, no password asked', b.success === true, b);
b = env.mfaBeginEnrolment({ username: 'nurse1' }, 'T-ADMIN');
check('an admin can move a nurse to a new phone', b.success === true && b.replacing === true, b);
x = env.mfaRemove('nurse1', '', 'T-ADMIN');
check('removing needs a reason', x.success === false && !!users()[1][5], x);
x = env.mfaRemove('nurse1', 'lost phone', 'T-NURSE');
check('a nurse cannot remove it', x.success === false && !!users()[1][5], x);
x = env.mfaRemove('nurse1', 'lost phone', 'T-ADMIN');
check('an admin can remove it for a nurse', x.success === true && !users()[1][5], x);
r = env.verifyLogin({ username: 'nurse1', password: 'Old-password-123' });
check('...after which the password signs straight in', r.success === true && !!r.sessionToken, r);
check('set-up and removal are audited', audit.some(a => a === 'MFA_ENROLLED nurse1') && audit.some(a => a === 'MFA_REMOVED nurse1'), audit.slice(-4));

if (failed) { console.log(`${failed} of ${ran} sign-in checks FAILED.`); process.exit(1); }
console.log(`${ran} sign-in checks passed (reset keeps the old password, temp works once, MFA gates the session, Google needs a token, in-app MFA set-up saves only a confirmed secret).`);
