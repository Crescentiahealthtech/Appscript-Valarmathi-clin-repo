/* Exact google.script.run chains that attach NO withFailureHandler. Such a
   call fails silently: the spinner it started never stops. */
const fs=require('fs'), path=require('path');
const ROOT = process.env.REPO || require('path').resolve(__dirname, '..');
function skipParens(s,i){let d=0,str=null;for(;i<s.length;i++){const c=s[i];
  if(str){if(c==='\\'){i++;continue;}if(c===str)str=null;continue;}
  if(c==='"'||c==="'"||c==='`'){str=c;continue;}
  if(c==='(')d++;else if(c===')'){d--;if(!d)return i+1;}}return -1;}
for(const f of fs.readdirSync(ROOT).filter(f=>f.endsWith('.html'))){
  const s=fs.readFileSync(path.join(ROOT,f),'utf8');
  let i=0;
  while((i=s.indexOf('google.script.run',i))!==-1){
    let j=i+'google.script.run'.length, hasFail=false, terminal=null;
    for(;;){
      while(j<s.length&&/\s/.test(s[j]))j++;
      if(s[j]!=='.')break; j++;
      while(j<s.length&&/\s/.test(s[j]))j++;
      const m=/^([A-Za-z_$][\w$]*)/.exec(s.slice(j)); if(!m)break;
      const name=m[1]; j+=name.length;
      while(j<s.length&&/\s/.test(s[j]))j++;
      if(s[j]!=='(')break;
      const end=skipParens(s,j); if(end<0)break;
      if(name==='withFailureHandler'){hasFail=true;j=end;continue;}
      if(/^with(Success|User)Handler$/.test(name)){j=end;continue;}
      terminal=name; break;
    }
    if(terminal && !hasFail)
      console.log(f+':'+s.slice(0,i).split('\n').length+'  '+terminal+'()');
    i+='google.script.run'.length;
  }
}
