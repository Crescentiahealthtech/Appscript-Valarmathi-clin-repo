/* Syntax-check every .gs file and every <script> block in every .html file.
   Apps Script runs V8, so node's parser is the same parser. */
const fs = require('fs'), path = require('path'), vm = require('vm');
const ROOT = process.env.REPO || path.resolve(__dirname, '..');
let bad = 0, checked = 0;

function check(label, code) {
  checked++;
  try { new vm.Script(code, { filename: label }); }
  catch (e) {
    bad++;
    console.log('SYNTAX  ' + label + '\n        ' + e.message.split('\n')[0]);
    const m = /(\d+)$/.exec(label);
  }
}

for (const f of fs.readdirSync(ROOT).sort()) {
  const p = path.join(ROOT, f);
  if (!fs.statSync(p).isFile()) continue;
  const src = fs.readFileSync(p, 'utf8');
  if (f.endsWith('.gs')) { check(f, src); continue; }
  if (!f.endsWith('.html')) continue;
  // Apps Script scriptlets are server-side; a file using them can't be parsed as JS.
  const re = /<script\b[^>]*>([\s\S]*?)<\/script>/gi;
  let m, i = 0;
  while ((m = re.exec(src))) {
    const tag = src.slice(m.index, m.index + m[0].indexOf('>') + 1);
    if (/\bsrc\s*=/.test(tag)) continue;            // external
    if (/type\s*=\s*["'](?!text\/javascript|module)/.test(tag)) continue; // templates
    const body = m[1];
    if (/<\?[=!]?/.test(body)) { i++; continue; }   // Apps Script scriptlet
    const line = src.slice(0, m.index).split('\n').length;
    check(f + ' <script> @line ' + line, body);
    i++;
  }
}
console.log('\n' + checked + ' blocks checked, ' + bad + ' with syntax errors.');
process.exit(bad ? 1 : 0);
