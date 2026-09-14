// ============================================================================
// NOT APPS SCRIPT. Do not paste this file into the Apps Script editor.
//
// This is a Node.js script. It reads the project's .gs and .html files as text
// and reports problems in them; it is not part of the application. Pasted into
// the editor it would fail on require()/__dirname, and dep.js would additionally
// clobber Deployment_Check.gs's DEP_MAP. See tools/README.md.
//
// Run it from the repository root:  node tools/dep.js
// Or run every check at once:       ./tools/check.sh
// ============================================================================
/* Does every name in DEP_MAP actually exist in the file it claims? */
const fs=require('fs'), path=require('path');
const ROOT = process.env.REPO || require('path').resolve(__dirname, '..');
const dc=fs.readFileSync(path.join(ROOT,'Deployment_Check.gs'),'utf8');
const m=dc.match(/var DEP_MAP = \{[\s\S]*?\n\};/);
const DEP_MAP=eval('('+m[0].replace(/^var DEP_MAP = /,'').replace(/;$/,'')+')');
let bad=0, wrongFile=0;
const where={};
for(const f of fs.readdirSync(ROOT).filter(x=>x.endsWith('.gs'))){
  const s=fs.readFileSync(path.join(ROOT,f),'utf8');
  const re=/^\s*function\s+([A-Za-z_$][\w$]*)\s*\(/gm; let x;
  while((x=re.exec(s))) (where[x[1]]=where[x[1]]||[]).push(f);
}
for(const [file,fns] of Object.entries(DEP_MAP)){
  for(const fn of fns){
    if(!where[fn]){ console.log('MISSING  '+fn+'  (DEP_MAP says '+file+')'); bad++; continue; }
    if(!where[fn].includes(file)){ console.log('WRONG FILE  '+fn+'  DEP_MAP says '+file+', actually in '+where[fn].join(',')); wrongFile++; }
  }
}
console.log('\n'+bad+' names that exist nowhere, '+wrongFile+' attributed to the wrong file.');
