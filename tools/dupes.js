/* Top-level names defined in more than one .gs file.

   Apps Script loads every .gs file into ONE global scope. Two `function x()`
   in different files is not an error there: the one loaded last wins and the
   other is dead code, with no warning anywhere. Two `const x` IS an error,
   and it stops every function in the project from loading. Either way the
   place to find out is here, not in the editor.

       node tools/dupes.js                    (part of ./tools/check.sh)  */
const fs = require('fs'), path = require('path');
const ROOT = process.env.REPO || path.resolve(__dirname, '..');
const seen = new Map();
for (const f of fs.readdirSync(ROOT).filter(f => f.endsWith('.gs'))) {
  const src = fs.readFileSync(path.join(ROOT, f), 'utf8');
  for (const m of src.matchAll(/^(?:function\s+([A-Za-z_$][\w$]*)|(?:var|let|const)\s+([A-Za-z_$][\w$]*))/gm)) {
    const name = m[1] || m[2];
    if (!seen.has(name)) seen.set(name, []);
    seen.get(name).push(f);
  }
}
const dupes = [...seen].filter(([, files]) => files.length > 1);
dupes.forEach(([name, files]) => console.log('  DUPLICATE  ' + name + '  in ' + files.join(', ')));
if (dupes.length) { console.log(dupes.length + ' name(s) defined more than once.'); process.exit(1); }
console.log('No top-level name is defined in more than one .gs file.');
