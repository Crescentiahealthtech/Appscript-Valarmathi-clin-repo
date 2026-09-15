/* Does every call actually PASS the session token, in the right slot?

   tools/rbac.js answers "does the server check". This answers the other half:
   "does the browser tell it who is asking". They fail in opposite directions
   and neither one catches the other's failure.

   A guard reads a named parameter — crescRequire_(sessionToken, …) — and
   JavaScript fills parameters BY POSITION. So

       function getWardBedBoard(ward, sessionToken)
       google.script.run.getWardBedBoard(ward)                  // undefined -> refused
       google.script.run.getIPRecordPrintHtml(id, opts, token)  // token lands in `opts`

   are both broken, and both look completely fine in a diff. The first shows up
   as "your session has expired" on a screen where the user is plainly signed
   in; the second silently sends a token where a settings object was expected
   and refuses the call for the wrong reason. Adding a parameter to a .gs
   function without updating every call site produces one or the other, and
   this project has 200-odd call sites.

   So: for every google.script.run call whose server function has a parameter
   named sessionToken or token, check that the argument in THAT POSITION looks
   like a token.

   What it cannot see: the dynamic run('name', [args]) helpers in the two lab
   desks, which append the token inside the helper. Those are checked by
   reading the helper once, not per call.  */
const fs = require('fs'), path = require('path');
const ROOT = path.join(__dirname, '..');

/** Anything that plausibly evaluates to the session token. */
const LOOKS_LIKE_TOKEN = /[Tt]oken|[Tt]ok\s*\(|leaving/;

// --- every server signature -------------------------------------------------
const sig = {};
for (const f of fs.readdirSync(ROOT)) {
  if (!f.endsWith('.gs')) continue;
  const src = fs.readFileSync(path.join(ROOT, f), 'utf8');
  for (const m of src.matchAll(/^function\s+([A-Za-z_$][\w$]*)\s*\(([^)]*)\)/gm)) {
    sig[m[1]] = {
      file: f,
      params: m[2].split(',').map(s => s.trim().split(/[\s=]/)[0]).filter(Boolean)
    };
  }
}

function skipGap(src, p) {
  for (;;) {
    while (p < src.length && /\s/.test(src[p])) p++;
    if (src[p] === '/' && src[p + 1] === '/') {
      const n = src.indexOf('\n', p); if (n === -1) return src.length; p = n + 1; continue;
    }
    if (src[p] === '/' && src[p + 1] === '*') {
      const e = src.indexOf('*/', p); if (e === -1) return src.length; p = e + 2; continue;
    }
    return p;
  }
}

/** Split an argument list on top-level commas, quotes and nesting respected. */
function splitArgs(s) {
  const out = []; let depth = 0, quote = null, cur = '';
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (quote) { cur += c; if (c === '\\') { cur += s[++i] || ''; continue; } if (c === quote) quote = null; continue; }
    if (c === "'" || c === '"' || c === '`') { quote = c; cur += c; continue; }
    if ('([{'.includes(c)) depth++;
    if (')]}'.includes(c)) depth--;
    if (c === ',' && depth === 0) { out.push(cur.trim()); cur = ''; continue; }
    cur += c;
  }
  if (cur.trim()) out.push(cur.trim());
  return out;
}

const missing = [], misplaced = [], stray = [];
let checked = 0;

for (const f of fs.readdirSync(ROOT)) {
  if (!f.endsWith('.html')) continue;
  const src = fs.readFileSync(path.join(ROOT, f), 'utf8');
  const NEEDLE = 'google.script.run';
  let at = 0;
  while ((at = src.indexOf(NEEDLE, at)) !== -1) {
    let p = at + NEEDLE.length;
    for (;;) {
      p = skipGap(src, p);
      if (src[p] !== '.') break;
      p++; p = skipGap(src, p);
      const m = /^[A-Za-z_$][\w$]*/.exec(src.slice(p));
      if (!m) break;
      const name = m[0];
      p += name.length;
      p = skipGap(src, p);
      if (src[p] !== '(') break;
      const open = p;
      let depth = 0, quote = null;
      for (; p < src.length; p++) {
        const c = src[p];
        if (quote) { if (c === '\\') { p++; continue; } if (c === quote) quote = null; continue; }
        if (c === "'" || c === '"' || c === '`') { quote = c; continue; }
        if (c === '(') depth++;
        else if (c === ')') { depth--; if (depth === 0) break; }
      }
      const close = p;
      const plumbing = /^with(SuccessHandler|FailureHandler|UserObject)$/.test(name);

      if (!plumbing) {
        const s = sig[name];
        if (s) {
          const args = splitArgs(src.slice(open + 1, close));
          const slot = s.params.findIndex(x => /^(sessionToken|token)$/.test(x));
          const line = src.slice(0, open).split('\n').length;
          const where = f + ':' + line;
          if (slot !== -1) {
            checked++;
            const given = args[slot] || '';
            if (!given) missing.push(`${where}  ${name}() — ${s.params.length} parameter(s), ` +
                                     `token expected at #${slot + 1}, ${args.length} passed`);
            else if (!LOOKS_LIKE_TOKEN.test(given))
              misplaced.push(`${where}  ${name}(${s.params.join(', ')}) — argument #${slot + 1} is "${given.split('\n')[0].slice(0, 50)}"`);
          } else if (args.some(a => LOOKS_LIKE_TOKEN.test(a) && /session/i.test(a))) {
            stray.push(`${where}  ${name}(${s.params.join(', ') || '—'}) is passed a token it has no parameter for`);
          }
        }
        break;
      }
      p = close + 1;
    }
    at += NEEDLE.length;
  }
}

console.log('=== calls whose server function wants a token but gets none ===');
console.log(missing.length ? missing.map(x => '  ' + x).join('\n') : '  (none)');
console.log('\n=== calls where the token is in the wrong argument position ===');
console.log(misplaced.length ? misplaced.map(x => '  ' + x).join('\n') : '  (none)');
if (stray.length) {
  console.log('\n=== a token passed to a function that has no parameter for it ===');
  console.log(stray.map(x => '  ' + x).join('\n'));
  console.log('  (harmless at runtime — JavaScript drops the extra argument — but it');
  console.log('   means somebody believed that call was guarded when it is not.)');
}
console.log(`\n${checked} call site(s) checked against their server signature.`);
