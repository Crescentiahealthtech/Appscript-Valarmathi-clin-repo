/* ===========================================================================
 * Text inside a <script> body that the HTML PARSER reacts to.
 *
 * Check 1 (validate.js) finds each script body with a regular expression and
 * syntax-checks it. The browser does not use a regular expression: it uses
 * the HTML tokeniser, and in a few cases the two disagree about where a
 * script ENDS. Where they disagree, check 1 passes on a string the browser
 * never runs, and the browser runs a truncated one — which is a SyntaxError,
 * which means the whole block is discarded, which means a screen that
 * renders perfectly and answers nothing.
 *
 * That is not hypothetical: it is what "Pharmacy Billing buttons dead" was,
 * reported as "missing ) after argument list".
 *
 * Three sequences matter inside a script body:
 *
 *   </script    ends the element wherever it appears — inside a string, a
 *               template literal or a comment. Always fatal. Write <\/script.
 *
 *   <!--        enters "script data escaped" state. Harmless on its own,
 *               but see the next one. Write <\!-- ; it is the same string
 *               to JavaScript and invisible to the parser.
 *
 *   <script     while in escaped state, enters "double escaped" state, and
 *               it then takes TWO </script> to leave. The real one is
 *               swallowed and the JS runs on into the following markup.
 *
 * Exits non-zero when it finds one, so check.sh can fail on it.
 * =========================================================================== */
const fs = require('fs'), path = require('path');
const ROOT = process.env.REPO || path.resolve(__dirname, '..');

var found = 0;
fs.readdirSync(ROOT).filter(f => f.endsWith('.html')).forEach(function (f) {
  const txt = fs.readFileSync(path.join(ROOT, f), 'utf8');
  const re = /<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script\s*>/gi;
  let m;
  while ((m = re.exec(txt))) {
    const body = m[1];
    const line0 = txt.slice(0, m.index).split('\n').length;
    const at = i => line0 + body.slice(0, i).split('\n').length - 1;
    const say = (i, what, why) => {
      found++;
      console.log('  ' + f + ':' + at(i) + '  ' + what + ' — ' + why);
      console.log('      ' + body.slice(Math.max(0, i - 50), i + 50).replace(/\n/g, '\\n'));
    };

    // </script anywhere is fatal.
    let x = /<\/script/gi, mm;
    while ((mm = x.exec(body))) say(mm.index, '</script', 'ends the element here; the rest of the JS is dropped');

    // <!-- is only dangerous once a <script follows it.
    const open = body.search(/<!--/);
    if (open !== -1) {
      const after = body.slice(open);
      const close = after.search(/-->/);
      const nested = after.search(/<script/i);
      if (nested !== -1 && (close === -1 || nested < close)) {
        say(open, '<!-- followed by <script',
            'puts the parser in double-escaped state; the closing </script> is swallowed');
      } else if (close === -1) {
        say(open, '<!-- with no -->',
            'leaves the parser in escaped state to the end of the body');
      }
    }
  }
});

console.log(found
  ? '\n' + found + ' hazard(s). Each one changes the script the browser runs.'
  : 'No script body contains text the HTML parser would act on.');
process.exit(found ? 1 : 0);
