/* Which client-callable .gs functions carry any session or permission check.
   The local equivalent of crescRbacCoverage() in RBAC.gs, which needs the
   Apps Script API turned on and therefore cannot run here.

   The number this prints is the one docs/DPDP_READINESS.md quotes, so a
   guard added to an endpoint shows up as progress rather than as a claim.  */
const fs = require('fs'), path = require('path');
const ROOT = path.join(__dirname, '..');

// Anything that resolves a session or asserts a permission counts as a guard.
const GUARDS = /crescRequire_|crescActor_|crescRequireOwnRecord_|dsx_requireRole_|dc_validateSession_|validateSession_|resolveScope_|resolveWriteDoctor_/;

// --- what the browser actually calls ---------------------------------------
const called = new Set();
for (const f of fs.readdirSync(ROOT)) {
  if (!f.endsWith('.html')) continue;
  const src = fs.readFileSync(path.join(ROOT, f), 'utf8');
  // The terminal call is the last one in a google.script.run chain.
  for (const chain of src.split('google.script.run').slice(1)) {
    const head = chain.slice(0, 700);
    const calls = (head.match(/\.\s*([A-Za-z_$][\w$]*)\s*\(/g) || [])
      .map(r => r.replace(/[.\s(]/g, ''))
      .filter(n => !/^with(SuccessHandler|FailureHandler|UserObject)$/.test(n));
    if (calls.length) called.add(calls[calls.length - 1]);
  }
  // The dynamic form: server[fn](...) with fn from a variable — the lab and
  // pharmacy desks both use it, so their endpoints are named as strings.
  for (const m of src.matchAll(/run\(\s*['"]([A-Za-z_$][\w$]*)['"]/g)) called.add(m[1]);
}

const guarded = [], open = [];
for (const f of fs.readdirSync(ROOT)) {
  if (!f.endsWith('.gs')) continue;
  const src = fs.readFileSync(path.join(ROOT, f), 'utf8');
  const re = /^function\s+([A-Za-z_$][\w$]*)\s*\(/gm;
  let m;
  while ((m = re.exec(src)) !== null) {
    const name = m[1];
    if (name.endsWith('_')) continue;            // private by convention
    const next = src.indexOf('\nfunction ', m.index + 1);
    const body = src.substring(m.index, next === -1 ? src.length : next);
    (GUARDS.test(body) ? guarded : open).push({ name, file: f, called: called.has(name) });
  }
}

const openCalled = open.filter(o => o.called);
console.log(`${guarded.length + open.length} public .gs functions`);
console.log(`  ${guarded.length} carry a session or permission check`);
console.log(`  ${open.length} carry none — ${openCalled.length} of those are reachable from the browser`);

if (openCalled.length) {
  console.log('\n=== unguarded AND client-callable (fix these) ===');
  const byFile = {};
  openCalled.forEach(o => { (byFile[o.file] = byFile[o.file] || []).push(o.name); });
  Object.keys(byFile).sort().forEach(f => console.log(' ' + f + ': ' + byFile[f].sort().join(', ')));
  console.log('\nAdd one line at the top of each:');
  console.log("  var actor = crescRequire_(sessionToken, 'area.verb');");
  console.log('See RBAC.gs for the vocabulary and docs/DPDP_READINESS.md finding C2.');
}
