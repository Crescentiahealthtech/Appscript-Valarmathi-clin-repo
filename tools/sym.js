const fs=require('fs'), path=require('path');
const ROOT = process.env.REPO || require('path').resolve(__dirname, '..');
const auth=fs.readFileSync(path.join(ROOT,'Auth.html'),'utf8');
const m=auth.match(/window\.CRESC_SYMBOL_HOME = \{([\s\S]*?)\n  \};/);
if(!m){console.log('map not found');process.exit(0)}
const where={};
for(const f of fs.readdirSync(ROOT).filter(x=>x.endsWith('.gs'))){
  const s=fs.readFileSync(path.join(ROOT,f),'utf8');
  const re=/^\s*function\s+([A-Za-z_$][\w$]*)\s*\(/gm; let x;
  while((x=re.exec(s))) (where[x[1]]=where[x[1]]||[]).push(f);
}
let bad=0;
const re=/([A-Za-z_$][\w$]*)\s*:\s*'([^']+)'/g; let e;
while((e=re.exec(m[1]))){
  const [_,fn,file]=e;
  if(!where[fn]){ console.log('MISSING   '+fn+' (map says '+file+')'); bad++; }
  else if(!where[fn].includes(file)){ console.log('WRONG FILE '+fn+' map says '+file+', actually '+where[fn].join(',')); bad++; }
}
console.log(bad?('\n'+bad+' wrong entries'):'CRESC_SYMBOL_HOME: every entry checks out.');
