/* Names used but defined nowhere — the bug the syntax check cannot see.

   Apps Script has one global scope and no module system, so a function that
   calls a name defined in no file is syntactically perfect and throws a
   ReferenceError only when that line runs. Two of these were live when this
   check was written: the Pharmacy Dashboard called pdashFetchDash(), which
   does not exist, so the screen had never loaded by itself; and a guard
   rewrite left an audit call naming a variable it had just removed.

   This runs ESLint's no-undef over every .gs file and every inline <script>,
   with every top-level name the project defines (and the Apps Script and
   browser globals) as known. It needs ESLint; without it, it says so and
   passes, so the check never blocks a machine that lacks it.

       node tools/undef.js                 (part of ./tools/check.sh)  */
const fs = require('fs'), path = require('path'), cp = require('child_process');
const ROOT = process.env.REPO || path.resolve(__dirname, '..');

let ESLint = null;
for (const where of [() => require('eslint'),
                     () => require(path.join(cp.execSync('npm root -g', { stdio: ['ignore', 'pipe', 'ignore'] }).toString().trim(), 'eslint'))]) {
  try { ESLint = where().ESLint; if (ESLint) break; } catch (e) { /* try the next */ }
}
if (!ESLint) { console.log('ESLint is not installed here; skipped (npm i -g eslint to run it).'); process.exit(0); }

const APPS_SCRIPT = ['SpreadsheetApp', 'CacheService', 'LockService', 'Utilities', 'Session', 'Logger',
  'HtmlService', 'ScriptApp', 'PropertiesService', 'DriveApp', 'GmailApp', 'MailApp', 'UrlFetchApp',
  'ContentService', 'Charts', 'CalendarApp', 'DocumentApp', 'XmlService', 'console'];
const BROWSER = ['window', 'document', 'navigator', 'location', 'localStorage', 'sessionStorage', 'setTimeout',
  'clearTimeout', 'setInterval', 'clearInterval', 'requestAnimationFrame', 'cancelAnimationFrame', 'fetch', 'URL',
  'Blob', 'FileReader', 'Event', 'CustomEvent', 'MutationObserver', 'IntersectionObserver', 'ResizeObserver',
  'getComputedStyle', 'alert', 'confirm', 'prompt', 'atob', 'btoa', 'Image', 'Audio', 'HTMLElement', 'Element',
  'Node', 'NodeList', 'KeyboardEvent', 'MouseEvent', 'performance', 'history', 'screen', 'matchMedia', 'crypto',
  'TextEncoder', 'TextDecoder', 'AbortController', 'FormData', 'DOMParser', 'print', 'open', 'close', 'focus',
  'blur', 'scrollTo', 'innerWidth', 'innerHeight', 'devicePixelRatio', 'speechSynthesis', 'SpeechSynthesisUtterance',
  'webkitSpeechRecognition', 'SpeechRecognition', 'MediaRecorder', 'AudioContext', 'webkitAudioContext',
  'Notification', 'queueMicrotask', 'structuredClone', 'HTMLCanvasElement', 'ImageData', 'createImageBitmap',
  'self', 'top', 'parent', 'name', 'event', 'CSS', 'visualViewport', 'DOMException', 'PointerEvent', 'TouchEvent',
  'InputEvent', 'HTMLInputElement', 'HTMLSelectElement', 'HTMLTextAreaElement', 'MediaStream', 'DocumentFragment',
  'getSelection', 'scrollY', 'scrollX', 'pageYOffset', 'pageXOffset', 'frameElement', 'console',
  // loaded from CDNs by Index.html, or by the page before a partial runs
  'google', 'bootstrap', 'Chart', 'firebase', 'jsQR', 'BarcodeDetector', 'auth'];
// Vendored libraries whose UMD wrappers probe for module loaders by design.
const VENDORED = ['DS_QR_Lib.gs', 'Barcode_QR_Lib.html'];

const known = {};
const gs = fs.readdirSync(ROOT).filter(f => f.endsWith('.gs'));
const html = fs.readdirSync(ROOT).filter(f => f.endsWith('.html'));
const scripts = [];
for (const f of gs) {
  const s = fs.readFileSync(path.join(ROOT, f), 'utf8');
  for (const m of s.matchAll(/^(?:async\s+)?function\s+([A-Za-z_$][\w$]*)/gm)) known[m[1]] = 'readonly';
  for (const m of s.matchAll(/^(?:var|let|const)\s+([A-Za-z_$][\w$]*)/gm)) known[m[1]] = 'writable';
}
const browserKnown = {};
for (const f of html) {
  const s = fs.readFileSync(path.join(ROOT, f), 'utf8');
  let i = 0;
  for (const m of s.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)) {
    const code = m[1].replace(/<\?!?=[\s\S]*?\?>/g, '0');
    scripts.push({ file: f, n: i++, code });
    for (const d of code.matchAll(/^\s{0,2}(?:async\s+)?function\s+([A-Za-z_$][\w$]*)/gm)) browserKnown[d[1]] = 'readonly';
    for (const d of code.matchAll(/^\s{0,2}(?:var|let|const)\s+([A-Za-z_$][\w$]*)/gm)) browserKnown[d[1]] = 'writable';
    for (const d of code.matchAll(/window\.([A-Za-z_$][\w$]*)\s*=/g)) browserKnown[d[1]] = 'writable';
  }
}
const asGlobals = list => list.reduce((o, n) => (o[n] = 'readonly', o), {});

(async function () {
  const lint = async (globals, files) => {
    const eslint = new ESLint({
      cwd: ROOT,
      overrideConfigFile: true,
      overrideConfig: [{ languageOptions: { ecmaVersion: 2022, sourceType: 'script', globals },
                         rules: { 'no-undef': 'error' } }]
    });
    const out = [];
    for (const f of files) {
      if (VENDORED.indexOf(f.file) !== -1) continue;
      // A virtual path inside the project: ESLint ignores anything outside cwd.
      const [res] = await eslint.lintText(f.code, { filePath: path.join(ROOT, '__undef_check__.js') });
      res.messages.filter(m => m.ruleId === 'no-undef').forEach(m => out.push(f.file + (f.n !== undefined ? ' <script ' + (f.n + 1) + '>' : '') +
                                         ':' + m.line + '  ' + m.message));
    }
    return out;
  };
  const server = await lint(Object.assign({}, asGlobals(APPS_SCRIPT), known),
                            gs.map(f => ({ file: f, code: fs.readFileSync(path.join(ROOT, f), 'utf8') })));
  const client = await lint(Object.assign({}, asGlobals(BROWSER), browserKnown), scripts);
  const all = server.concat(client);
  all.forEach(l => console.log(l));
  console.log(all.length ? all.length + ' name(s) used but defined nowhere.'
                         : 'Every name used in ' + gs.length + ' .gs files and ' + scripts.length +
                           ' inline scripts is defined somewhere in the project.');
  process.exit(all.length ? 1 : 0);
})();
