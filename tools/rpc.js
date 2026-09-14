// ============================================================================
// NOT APPS SCRIPT. Do not paste this file into the Apps Script editor.
//
// This is a Node.js script. It reads the project's .gs and .html files as text
// and reports problems in them; it is not part of the application. Pasted into
// the editor it would fail on require()/__dirname, and dep.js would additionally
// clobber Deployment_Check.gs's DEP_MAP. See tools/README.md.
//
// Run it from the repository root:  node tools/rpc.js
// Or run every check at once:       ./tools/check.sh
// ============================================================================
/* Properly walk each google.script.run chain to the terminal call. */
const fs=require('fs'), path=require('path');
const ROOT = process.env.REPO || require('path').resolve(__dirname, '..');
const all=fs.readdirSync(ROOT);
let gs='';
for(const f of all) if(f.endsWith('.gs')) gs+='\n'+fs.readFileSync(path.join(ROOT,f),'utf8');
const server=new Set();
{ const re=/^\s*function\s+([A-Za-z_$][\w$]*)\s*\(/gm; let m;
  while((m=re.exec(gs))) server.add(m[1]); }

function skipParens(s,i){ // s[i] === '(' -> index just past the match
  let d=0, str=null;
  for(;i<s.length;i++){
    const c=s[i];
    if(str){ if(c==='\\'){i++;continue;} if(c===str)str=null; continue; }
    if(c==='"'||c==="'"||c==='`'){str=c;continue;}
    if(c==='(')d++;
    else if(c===')'){d--; if(!d) return i+1;}
  }
  return -1;
}
const missing=new Map(), found=new Set();
for(const f of all.filter(f=>f.endsWith('.html'))){
  const s=fs.readFileSync(path.join(ROOT,f),'utf8');
  let i=0;
  while((i=s.indexOf('google.script.run',i))!==-1){
    let j=i+'google.script.run'.length;
    for(;;){
      // skip whitespace and one '.'
      while(j<s.length && /[\s\r\n]/.test(s[j])) j++;
      if(s[j]!=='.') break;
      j++;
      while(j<s.length && /[\s\r\n]/.test(s[j])) j++;
      const m=/^([A-Za-z_$][\w$]*)/.exec(s.slice(j));
      if(!m) break;
      const name=m[1];
      j+=name.length;
      while(j<s.length && /[\s\r\n]/.test(s[j])) j++;
      if(s[j]!=='(') break;
      const end=skipParens(s,j);
      if(end<0) break;
      if(/^with(Success|Failure|User)Handler$/.test(name)) { j=end; continue; }
      // terminal call
      found.add(name);
      if(!server.has(name)){
        const line=s.slice(0,j).split('\n').length;
        if(!missing.has(name)) missing.set(name,[]);
        missing.get(name).push(f+':'+line);
      }
      break;
    }
    i+= 'google.script.run'.length;
  }
}
console.log('=== terminal google.script.run calls with NO .gs function ===');
if(!missing.size) console.log('(none)');
for(const [n,w] of missing) console.log('  '+n.padEnd(32)+w.slice(0,3).join(', '));
console.log('\n'+found.size+' distinct server methods called from the client.');
