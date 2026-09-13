/* Classes used in markup/JS that no stylesheet in the project defines and
   Bootstrap 5 does not supply. */
const fs=require('fs'), path=require('path');
const ROOT = process.env.REPO || require('path').resolve(__dirname, '..');
const files=fs.readdirSync(ROOT).filter(f=>f.endsWith('.html'));
let all='', css='';
for(const f of files){ const s=fs.readFileSync(path.join(ROOT,f),'utf8'); all+='\n'+s;
  const re=/<style\b[^>]*>([\s\S]*?)<\/style>/gi; let m;
  while((m=re.exec(s))) css+='\n'+m[1]; }

const defined=new Set();
{ const re=/\.(-?[_a-zA-Z][\w-]*)/g; let m; while((m=re.exec(css))) defined.add(m[1]); }

// Bootstrap 5 utility/component prefixes we don't ship rules for.
const BS=/^(container|row|col|g|gx|gy|d|flex|justify|align|order|offset|m|mt|mb|ms|me|mx|my|p|pt|pb|ps|pe|px|py|w|h|mw|mh|vw|vh|min|max|text|fs|fw|lh|font|text-|bg|border|rounded|shadow|opacity|overflow|position|top|bottom|start|end|translate|float|clearfix|link|visually|stretched|vr|hstack|vstack|btn|badge|alert|card|nav|navbar|dropdown|modal|offcanvas|table|form|input|spinner|progress|list|accordion|breadcrumb|pagination|placeholder|popover|tooltip|toast|carousel|close|collapse|fade|show|active|disabled|ratio|sticky|fixed|user|pe-none|pe-auto|gap|is|was|valid|invalid|small|mark|blockquote|figure|img|initialism|display|lead|fa|fas|far|fab|fa-|ri|material)(-|$)/;
const used=new Map();
{ const re=/class\s*=\s*(["'`])([\s\S]*?)\1/g; let m;
  while((m=re.exec(all))){
    for(let c of m[2].split(/\s+/)){
      c=c.trim(); if(!c||c.includes('${')||c.includes('<')||c.includes('+')) continue;
      if(BS.test(c)||defined.has(c)) continue;
      used.set(c,(used.get(c)||0)+1);
    }
  }
}
// classList.add('x') / className = 'x'
{ const re=/classList\.(?:add|toggle|remove)\(\s*'([^']+)'/g; let m;
  while((m=re.exec(all))){ const c=m[1]; if(BS.test(c)||defined.has(c)) continue; used.set(c,(used.get(c)||0)+1); } }

const rows=[...used.entries()].filter(([c,n])=>n>=2).sort((a,b)=>b[1]-a[1]);
console.log('classes used '+'>=2 times with no rule anywhere (count, class):');
rows.forEach(([c,n])=>console.log(String(n).padStart(4)+'  '+c));
console.log('\n'+rows.length+' distinct.');
