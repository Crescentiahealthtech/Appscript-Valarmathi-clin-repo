// ============================================================================
// NOT APPS SCRIPT. Do not paste this file into the Apps Script editor.
//
// This is a Node.js script. It reads the project's .gs and .html files as text
// and reports problems in them; it is not part of the application. Pasted into
// the editor it would fail on require()/__dirname, and dep.js would additionally
// clobber Deployment_Check.gs's DEP_MAP. See tools/README.md.
//
// Run it from the repository root:  node tools/audit.js
// Or run every check at once:       ./tools/check.sh
// ============================================================================
/* Static audit: inline onclick handlers that name a function nothing defines,
   and getElementById ids that no markup declares. */
const fs=require('fs'), path=require('path');
const ROOT = process.env.REPO || require('path').resolve(__dirname, '..');
const files=fs.readdirSync(ROOT).filter(f=>f.endsWith('.html')||f.endsWith('.gs'));
let html='', js='';
const src={};
for(const f of files){ const s=fs.readFileSync(path.join(ROOT,f),'utf8'); src[f]=s;
  if(f.endsWith('.html')) html+=s; js+=s; }

// ---- 1. Every function name reachable as a global ------------------------
const defined=new Set();
for(const re of [
  /\bfunction\s+([A-Za-z_$][\w$]*)\s*\(/g,
  /\bwindow\.([A-Za-z_$][\w$]*)\s*=/g,
  /\b(?:var|let|const)\s+([A-Za-z_$][\w$]*)\s*=\s*(?:function|\([^)]*\)\s*=>|async)/g,
  /\b(?:var|let|const)\s+([A-Za-z_$][\w$]*)\s*=\s*\(function/g,
]) { let m; while((m=re.exec(js))) defined.add(m[1]); }
// Browser + Bootstrap + Apps Script + FontAwesome globals we don't define.
for(const g of ['alert','confirm','prompt','print','open','close','event','this','setTimeout',
  'console','parseInt','parseFloat','Number','String','Boolean','Array','Object','JSON','Math',
  'Date','RegExp','Promise','Map','Set','isNaN','encodeURIComponent','decodeURIComponent',
  'bootstrap','google','window','document','history','location','navigator','return','if','for',
  'switch','while','typeof','new','delete','void','eval','require','Utilities','SpreadsheetApp',
  'Session','Logger','LockService','PropertiesService','HtmlService','DriveApp','MailApp',
  'UrlFetchApp','CacheService','ScriptApp','GmailApp','Charts','XmlService','ContentService'])
  defined.add(g);

// ---- 2. Inline handlers ---------------------------------------------------
const bad=[];
for(const f of files.filter(f=>f.endsWith('.html'))){
  const s=src[f];
  const re=/\son(?:click|change|input|submit|blur|focus|keyup|keydown|mousedown)\s*=\s*(["'])([\s\S]*?)\1/gi;
  let m;
  while((m=re.exec(s))){
    const code=m[2];
    const cre=/([A-Za-z_$][\w$]*)\s*\(/g; let c;
    while((c=cre.exec(code))){
      const name=c[1];
      // Skip method calls (obj.method(...)) and keywords.
      const before=code.slice(0,c.index).trimEnd();
      if(before.endsWith('.')) continue;
      if(defined.has(name)) continue;
      const line=s.slice(0,m.index).split('\n').length;
      bad.push(`${f}:${line}  ->  ${name}()   in  ${code.slice(0,70).replace(/\s+/g,' ')}`);
    }
  }
}
console.log('=== inline handlers naming an undefined function ===');
console.log(bad.length?[...new Set(bad)].join('\n'):'(none)');

// ---- 3. getElementById ids with no matching id= in any markup -------------
const declared=new Set();
{ const re=/\bid\s*=\s*(["'])([^"']+)\1/g; let m; while((m=re.exec(html))) declared.add(m[2]);
  // ids built in template strings: id="foo-${x}" etc. -> record the prefix
  const re2=/\bid\s*=\s*\\?["']([A-Za-z0-9_-]+)/g; while((m=re2.exec(js))) declared.add(m[1]);
  const re3=/\.id\s*=\s*['"]([^'"]+)['"]/g; while((m=re3.exec(js))) declared.add(m[1]);
}
const missing=new Set();
{ const re=/getElementById\(\s*(['"])([A-Za-z0-9_-]+)\1\s*\)/g; let m;
  while((m=re.exec(js))) if(!declared.has(m[2])) missing.add(m[2]); }
console.log('\n=== getElementById ids never declared in markup ===');
console.log(missing.size?[...missing].join('\n'):'(none)');
