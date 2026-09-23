/* Which client-callable .gs functions carry any session or permission check.
   The local equivalent of crescRbacCoverage() in RBAC.gs, which needs the
   Apps Script API turned on and therefore cannot run here.

   The number this prints is the one docs/DPDP_READINESS.md quotes, so a
   guard added to an endpoint shows up as progress rather than as a claim.

   ---------------------------------------------------------------------------
   WHY THIS FILE WAS REWRITTEN (and why the old number was wrong)

   The first version was wrong in both directions, and the two errors did not
   cancel out — they hid different things.

   TOO FEW, because of how it found what the browser calls. It read the first
   700 characters after `google.script.run` and took the last `.name(` in that
   window. A call written as

       google.script.run
         .withSuccessHandler(res => { ...forty lines of UI... })
         .withFailureHandler(err => { ... })
         .registerPatient(data);

   puts the server function PAST the window, so registerPatient — an
   unauthenticated write to the patient master — did not appear in the list at
   all. This version walks the chain properly: after `google.script.run` it
   reads `.identifier`, skips the balanced parentheses of any
   with{Success,Failure}Handler / withUserObject, and the first identifier that
   is not one of those IS the server call. 136 endpoints were reachable, not
   62.

   TOO MANY, because of how it found a guard. It looked for the guard call
   inside the function's own body, so every endpoint that delegates its check
   to a gate helper — hb_actor_() in Hospital_Billing.gs, resolveIPRead_() /
   resolveIPWrite_() in IP_Clinical_Access.gs — was reported as unguarded
   although the session is validated before a row is read. This version
   follows the call graph: a function counts as guarded if it, or anything it
   calls, reaches one of the roots below. A hand-maintained list of helper
   names would have needed editing every time somebody wrote a new gate.

   WHAT IT STILL CANNOT SEE. It is a static reader, not an interpreter. A
   function that calls a guarded function but ignores its answer looks guarded
   here; so does one that guards a branch it does not take. It tells you where
   to look, and the code has to be read.  */
const fs = require('fs'), path = require('path');
const ROOT = path.join(__dirname, '..');

/** The calls that actually resolve a session. Everything else is guarded by
 *  reaching one of these, however many helpers deep. */
const GUARD_ROOTS = [
  'crescRequire_',           // RBAC.gs — the permission matrix
  'crescActor_',             // RBAC.gs — resolve, may return null
  'crescRequireOwnRecord_',  // RBAC.gs — a patient, their own record only
  'crescEditorOnly_',        // RBAC.gs — the script owner, or an admin's inner call
  'crescTriggerOnly_',       // RBAC.gs — an installed trigger, or the owner
  'dsx_requireRole_',        // DS_Workflow.gs — the discharge desk
  'dc_validateSession_',     // Doctor_Session_Store.gs — the durable session
  'validateSession_'         // Doctors_Engine.gs — the cache-only original
];

/**
 * Endpoints that are deliberately reachable without a session, each with the
 * reason. Read from RBAC.gs rather than kept here, so the list a reviewer
 * finds in the source is the list the count uses.
 */
function publicByDesign() {
  const src = fs.readFileSync(path.join(ROOT, 'RBAC.gs'), 'utf8');
  const block = /var\s+CRESC_PUBLIC_BY_DESIGN\s*=\s*\{([\s\S]*?)\n\};/.exec(src);
  if (!block) return {};
  const out = {};
  for (const m of block[1].matchAll(/'([A-Za-z_$][\w$]*)'\s*:\s*'((?:[^'\\]|\\.)*)'/g)) {
    out[m[1]] = m[2];
  }
  return out;
}

// --- what the browser actually calls ---------------------------------------
/**
 * The terminal call of every google.script.run chain in one file.
 * Walks the chain instead of guessing from a fixed window: see the note above.
 */
/** Whitespace and comments between the links of a chain. A `// still load`
 *  after a handler used to end the walk, and the endpoint after it —
 *  setupLabDatabase(), which rebuilds the lab database — went unreported. */
function skipGap(src, p) {
  for (;;) {
    while (p < src.length && /\s/.test(src[p])) p++;
    if (src[p] === '/' && src[p + 1] === '/') {
      const nl = src.indexOf('\n', p);
      if (nl === -1) return src.length;
      p = nl + 1; continue;
    }
    if (src[p] === '/' && src[p + 1] === '*') {
      const end = src.indexOf('*/', p);
      if (end === -1) return src.length;
      p = end + 2; continue;
    }
    return p;
  }
}

function terminalCalls(src) {
  const found = new Set();
  const NEEDLE = 'google.script.run';
  let at = 0;
  while ((at = src.indexOf(NEEDLE, at)) !== -1) {
    let p = at + NEEDLE.length;
    for (;;) {
      p = skipGap(src, p);
      if (src[p] !== '.') break;
      p++;
      p = skipGap(src, p);
      const m = /^[A-Za-z_$][\w$]*/.exec(src.slice(p));
      if (!m) break;
      const name = m[0];
      p += name.length;
      p = skipGap(src, p);
      const isPlumbing = /^with(SuccessHandler|FailureHandler|UserObject)$/.test(name);
      if (src[p] !== '(') { if (!isPlumbing) found.add(name); break; }
      if (!isPlumbing) { found.add(name); break; }
      // Skip the handler's argument, quotes and nesting included.
      let depth = 0, quote = null;
      for (; p < src.length; p++) {
        const c = src[p];
        if (quote) { if (c === '\\') { p++; continue; } if (c === quote) quote = null; continue; }
        // A comment inside a handler is not code: an apostrophe in
        // `/* the caller's refresh */` used to open a quote that never closed,
        // and every endpoint after it in the file went unseen.
        if (c === '/' && src[p + 1] === '/') { const n = src.indexOf('\n', p); p = (n === -1 ? src.length : n) - 1; continue; }
        if (c === '/' && src[p + 1] === '*') { const e = src.indexOf('*/', p + 2); p = (e === -1 ? src.length : e + 1); continue; }
        if (c === "'" || c === '"' || c === '`') { quote = c; continue; }
        if (c === '(') depth++;
        else if (c === ')') { depth--; if (depth === 0) { p++; break; } }
      }
    }
    at += NEEDLE.length;
  }
  return found;
}

const called = new Set();
for (const f of fs.readdirSync(ROOT)) {
  if (!f.endsWith('.html')) continue;
  const src = fs.readFileSync(path.join(ROOT, f), 'utf8');
  terminalCalls(src).forEach(n => called.add(n));
  // The dynamic form: server[fn](...) with fn from a variable — the lab and
  // pharmacy desks both use it, so their endpoints are named as strings.
  for (const m of src.matchAll(/run\(\s*['"]([A-Za-z_$][\w$]*)['"]/g)) called.add(m[1]);
}

// --- every function in the project, and what it calls -----------------------
const fns = new Map();   // name -> { file, calls:Set }
for (const f of fs.readdirSync(ROOT)) {
  if (!f.endsWith('.gs')) continue;
  const src = fs.readFileSync(path.join(ROOT, f), 'utf8');
  const re = /^function\s+([A-Za-z_$][\w$]*)\s*\(/gm;
  let m;
  while ((m = re.exec(src)) !== null) {
    const name = m[1];
    const next = src.indexOf('\nfunction ', m.index + 1);
    const body = src.substring(m.index, next === -1 ? src.length : next)
      // A guard named in a comment is not a guard.
      .replace(/\/\*[\s\S]*?\*\//g, ' ').replace(/(^|[^:])\/\/[^\n]*/g, '$1 ');
    const calls = new Set();
    for (const c of body.matchAll(/([A-Za-z_$][\w$]*)\s*\(/g)) {
      if (c[1] !== name) calls.add(c[1]);
    }
    fns.set(name, { file: f, calls });
  }
}

/** Does this function reach a guard root, through however many helpers? */
const memo = new Map();
function guarded(name, seen) {
  if (memo.has(name)) return memo.get(name);
  seen = seen || new Set();
  if (seen.has(name)) return false;          // recursion: no verdict from here
  seen.add(name);
  const fn = fns.get(name);
  if (!fn) return false;
  let ok = false;
  for (const c of fn.calls) {
    if (GUARD_ROOTS.indexOf(c) !== -1) { ok = true; break; }
    if (guarded(c, seen)) { ok = true; break; }
  }
  if (seen.size === 1) memo.set(name, ok);
  return ok;
}

const PUBLIC = publicByDesign();
const guardedList = [], open = [], declaredPublic = [];
for (const [name, fn] of fns) {
  if (name.endsWith('_')) continue;                    // private by convention
  if (PUBLIC[name]) { declaredPublic.push({ name, file: fn.file }); continue; }
  (guarded(name) ? guardedList : open).push({ name, file: fn.file, called: called.has(name) });
}

// A public function that mints its own session passes every check above — it
// reaches crescRequire_ — while guarding nothing: the caller supplies no
// credential and gets an admin's. Only the sign-in endpoints may do this.
const MINTS_OK = ['verifyLogin', 'verifyGoogleLogin', 'verifyMFA'];
const minting = [];
for (const [name, fn] of fns) {
  if (name.endsWith('_') || MINTS_OK.indexOf(name) !== -1) continue;
  // Editor-only jobs may mint one: crescEditorOnly_ refuses the browser first.
  if (fn.calls.has('issueSession_') && !fn.calls.has('crescEditorOnly_')) minting.push(fn.file + ': ' + name);
}

const openCalled = open.filter(o => o.called);
const total = guardedList.length + open.length + declaredPublic.length;
console.log(`${total} public .gs functions`);
console.log(`  ${guardedList.length} carry a session or permission check`);
console.log(`  ${declaredPublic.length} are public by design (CRESC_PUBLIC_BY_DESIGN in RBAC.gs)`);
console.log(`  ${open.length} carry none — ${openCalled.length} of those are reachable from the browser`);

if (openCalled.length) {
  console.log('\n=== unguarded AND client-callable (fix these) ===');
  const byFile = {};
  openCalled.forEach(o => { (byFile[o.file] = byFile[o.file] || []).push(o.name); });
  Object.keys(byFile).sort().forEach(f => console.log(' ' + f + ': ' + byFile[f].sort().join(', ')));
  console.log('\nAdd one line at the top of each:');
  console.log("  var actor = crescRequire_(sessionToken, 'area.verb');");
  console.log('See RBAC.gs for the vocabulary and docs/DPDP_READINESS.md finding C2.');
} else {
  console.log('\nEvery endpoint the browser can reach carries a check, or is');
  console.log('declared public by design with its reason. That is finding C2 closed.');
}

if (minting.length) {
  console.log('\n=== public functions that mint a session for their caller (fix these) ===');
  minting.sort().forEach(m => console.log(' ' + m));
  console.log('Start it with crescEditorOnly_(), or end the name in _, so google.script.run cannot use it.');
}

// Unreachable from the browser is not the same as safe: the deployment is
// what decides that, and a function nobody calls today is one refactor from
// being called tomorrow.
const openUncalled = open.filter(o => !o.called);
if (openUncalled.length) {
  console.log('\n(' + openUncalled.length + ' unguarded public functions are not called from any ' +
              '.html in this project.\n Apps Script still exposes every one of them to ' +
              'google.script.run — they are\n reachable by anyone who knows the name, so ' +
              'they are a smaller hole, not none.)');
}
