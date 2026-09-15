/* The password code, exercised.

   Auth_Credentials.gs is the one file in this project where a bug is silent
   and catastrophic at the same time: a hash that does not round-trip locks out
   the whole clinic, and a comparison that returns true too easily lets anybody
   in. Neither shows up in a syntax check, and there is no test runner here
   because the code is pasted into an Apps Script editor.

   So this loads the real source, stubs the four Utilities calls it uses with
   Node's crypto — the same HMAC-SHA-256, the same base64 — and runs the
   assertions that matter. It tests the ALGORITHM, not Apps Script: if this
   passes, a failure in the editor is a platform difference, and that is a much
   smaller haystack.

       node tools/credtest.js                 (part of ./tools/check.sh)  */
const fs = require('fs'), path = require('path'), crypto = require('crypto');
const ROOT = path.join(__dirname, '..');

// --- the platform, as far as this file is concerned -------------------------
const Utilities = {
  computeHmacSha256Signature(value, key) {
    const v = Buffer.isBuffer(value) ? value : Buffer.from(toBytes(value));
    const k = Buffer.isBuffer(key) ? key : Buffer.from(toBytes(key));
    // Apps Script returns a signed Byte[]; the values round-trip through
    // base64 either way, and the test only cares that it is deterministic.
    return Array.from(crypto.createHmac('sha256', k).update(v).digest());
  },
  computeDigest(_alg, value) {
    return Array.from(crypto.createHash('sha256').update(String(value)).digest());
  },
  DigestAlgorithm: { SHA_256: 'SHA_256' },
  base64Encode(bytes) { return Buffer.from(toBytes(bytes)).toString('base64'); },
  getUuid() { return crypto.randomUUID(); },
  newBlob(s) { return { getBytes: () => Array.from(Buffer.from(String(s), 'utf8')) }; },
  formatDate() { return '2026-09-15'; }
};
function toBytes(v) {
  if (Array.isArray(v)) return v.map(b => b & 0xff);
  if (Buffer.isBuffer(v)) return Array.from(v);
  return Array.from(Buffer.from(String(v), 'utf8'));
}
const PropertiesService = { getScriptProperties: () => ({ getProperty: () => null }) };
const Logger = { log() {} };
const Session = { getScriptTimeZone: () => 'Asia/Kolkata' };

// --- load the real file -----------------------------------------------------
// Only the pure part: everything below SECTION B touches sheets, which this
// harness deliberately does not pretend to have.
const src = fs.readFileSync(path.join(ROOT, 'Auth_Credentials.gs'), 'utf8');
const pure = src.slice(0, src.indexOf('// SECTION B'));
const load = new Function('Utilities', 'PropertiesService', 'Logger', 'Session',
  pure + `
  return { crescPwdEncode_, crescPwdVerify_, crescPwdIsHashed_, crescPwdEquals_,
           crescRandomPassword_, crescPwdPolicy_, cresc_pwdDigest_,
           CRESC_PWD_MIN_LENGTH, CRESC_PWD_ALPHABET };`);
const C = load(Utilities, PropertiesService, Logger, Session);

// --- the assertions ---------------------------------------------------------
let failed = 0, ran = 0;
function check(what, ok, detail) {
  ran++;
  if (ok) return;
  failed++;
  console.log('  FAIL  ' + what + (detail ? '\n        ' + detail : ''));
}

// A password must verify against its own digest, and nothing else's.
const stored = C.crescPwdEncode_('correct horse battery staple');
check('a password verifies against its own stored digest',
      C.crescPwdVerify_('correct horse battery staple', stored).ok === true);
check('a wrong password does not verify',
      C.crescPwdVerify_('correct horse battery stapl', stored).ok === false);
check('an empty password does not verify',
      C.crescPwdVerify_('', stored).ok === false);
check('case matters', C.crescPwdVerify_('Correct horse battery staple', stored).ok === false);

// The stored form must be recognisable, and must not contain the password.
check('the stored form is marked as hashed', C.crescPwdIsHashed_(stored) === true);
check('the stored form does not contain the password',
      stored.indexOf('correct') === -1);
check('the stored form has all five fields', stored.split('$').length === 5,
      'got: ' + stored.split('$').length);

// Two users with the SAME password must not share a digest — that is what the
// salt is for, and a shared digest tells an attacker which accounts to try first.
const a = C.crescPwdEncode_('same password'), b = C.crescPwdEncode_('same password');
check('the same password salts differently each time', a !== b);
check('both still verify',
      C.crescPwdVerify_('same password', a).ok && C.crescPwdVerify_('same password', b).ok);

// A plain-text cell is LEGACY, never ok. This is the whole migration story:
// accepting it would keep every already-readable password working.
const legacy = C.crescPwdVerify_('Mei2001', 'Mei2001');
check('a plain-text cell is reported legacy', legacy.legacy === true);
check('a plain-text cell never verifies, even with the right password',
      legacy.ok === false);
check('an empty cell is neither legacy nor ok',
      C.crescPwdVerify_('x', '').legacy === false && C.crescPwdVerify_('x', '').ok === false);

// A malformed stored value must fail closed rather than throw.
['pbkdf2$sha256$', 'pbkdf2$sha256$abc$salt$hash', 'pbkdf2$sha256$0$s$h'].forEach(bad => {
  let threw = false, res = null;
  try { res = C.crescPwdVerify_('anything', bad); } catch (e) { threw = true; }
  check('a malformed stored value "' + bad + '" fails closed without throwing',
        !threw && res && res.ok === false);
});

// Constant-time comparison must still be a CORRECT comparison.
check('equal strings compare equal', C.crescPwdEquals_('abcdef', 'abcdef') === true);
check('different strings compare unequal', C.crescPwdEquals_('abcdef', 'abcdeg') === false);
check('different lengths compare unequal', C.crescPwdEquals_('abc', 'abcd') === false);
check('a prefix does not pass', C.crescPwdEquals_('abc', 'abcdef') === false);
check('empty against empty is equal', C.crescPwdEquals_('', '') === true);
check('empty against non-empty is not', C.crescPwdEquals_('', 'a') === false);

// The iteration count is part of the stored form, so an old digest keeps
// verifying after the count is raised.
const at2000 = 'pbkdf2$sha256$2000$fixedsalt$' + C.cresc_pwdDigest_('hello there', 'fixedsalt', 2000);
check('a digest stored at a different iteration count still verifies',
      C.crescPwdVerify_('hello there', at2000).ok === true);

// Generated passwords: right length, no look-alike characters, not repeated.
const gen = C.crescRandomPassword_(12);
check('a generated password is the length asked for', gen.length === 12, 'got ' + gen.length);
check('a generated password avoids O 0 I l 1 S 5 B 8',
      !/[O0Il1S5B8]/.test(gen), 'got "' + gen + '"');
const many = new Set();
for (let i = 0; i < 200; i++) many.add(C.crescRandomPassword_(10));
check('200 generated passwords are 200 different passwords', many.size === 200,
      'got ' + many.size);

// The policy: length first, then the two patterns this clinic was bitten by.
check('a short password is refused',
      C.crescPwdPolicy_('short', 'user', '').ok === false);
check('a long enough one is accepted',
      C.crescPwdPolicy_('a reasonable pass phrase', 'user', '').ok === true);
check('a password containing the user id is refused',
      C.crescPwdPolicy_('nurse1-nurse1-nurse1', 'nurse1', '').ok === false);
check('the old derived portal pattern is refused',
      C.crescPwdPolicy_('Mei2001', 'LMTVS0001', 'Meivasagam').ok === false);
check('one repeated character is refused',
      C.crescPwdPolicy_('aaaaaaaaaaaa', 'user', '').ok === false);
check('the minimum length is the one the policy advertises',
      C.crescPwdPolicy_('x'.repeat(C.CRESC_PWD_MIN_LENGTH - 1), 'u', '').ok === false &&
      C.crescPwdPolicy_('xyzabcdefghij'.slice(0, C.CRESC_PWD_MIN_LENGTH), 'u', '').ok === true);

console.log(failed
  ? `\n${failed} of ${ran} credential checks FAILED.`
  : `${ran} credential checks passed (hash round-trip, salting, legacy refusal, ` +
    `constant-time compare, generation, policy).`);
process.exit(failed ? 1 : 0);
