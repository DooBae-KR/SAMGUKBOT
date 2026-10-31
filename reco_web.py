"""추천덱 HTML 페이지(/reco): 내 추천덱(장수 카드·전법·초월 돌·턴 순서·평가 버튼) + 추천덱 랭킹 + 시너지 점수."""
import html, json

TPL = r'''<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>추천덱</title>
<style>
:root{--bg:#0f1a18;--card:#16262a;--card2:#0e1718;--line:#2b4448;--txt:#e8efe9;--mute:#8fa8a3;--gold:#f0c75e;--jade:#5fcfb1}
@media (prefers-color-scheme: light){:root{--bg:#eef3f1;--card:#fff;--card2:#f4f8f6;--line:#cdd9d5;--txt:#1b2a27;--mute:#5d716c;--gold:#9a6b00;--jade:#0d8a6a}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--txt);font-family:'Noto Sans KR','Apple SD Gothic Neo','Malgun Gothic',sans-serif;font-size:14px}
.wrap{max-width:1040px;margin:0 auto;padding:14px 14px 40px}h1{font-size:20px;margin:6px 0 2px}.sub{color:var(--mute);font-size:12px;margin-bottom:10px}
.bar{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin:10px 0}.tab,.sbtn{background:var(--card);border:1px solid var(--line);color:var(--txt);border-radius:8px;padding:7px 14px;cursor:pointer;font-size:13px;text-decoration:none}
.tab.on,.sbtn.on{background:var(--jade);color:#06231c;border-color:var(--jade);font-weight:700}.sp{flex:1}
.unit{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:12px;margin-bottom:14px}
.uh{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:10px}.badge{background:var(--gold);color:#2b2000;border-radius:14px;padding:2px 10px;font-weight:700;font-size:12px}
.uh .nm{font-size:16px;font-weight:700}.uh .tr{color:var(--mute);font-size:12px}.votes{margin-left:auto;display:flex;gap:6px;flex-wrap:wrap}
.vg{display:flex;align-items:center;gap:3px;background:var(--card2);border:1px solid var(--line);border-radius:8px;padding:2px 4px;font-size:11px;color:var(--mute)}
.vb{background:transparent;border:1px solid var(--line);color:var(--txt);border-radius:6px;padding:3px 7px;cursor:pointer;font-size:12px}.vb.u.on{background:#2e7d4f;border-color:#2e7d4f;color:#fff}.vb.d.on{background:#a64040;border-color:#a64040;color:#fff}
.hs{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}@media(max-width:760px){.hs{gap:5px;grid-template-columns:repeat(3,minmax(0,1fr))}.unit{padding:8px}.h,.hi,.tac>div{min-width:0}.h{flex-direction:column;gap:5px;padding:5px}.pic{width:100%;height:auto;aspect-ratio:3/4}.pic .lb{font-size:11px;padding:12px 0 2px}
 .hn{font-size:13px;flex-wrap:wrap;gap:3px}.nat{font-size:10px;padding:0 4px}.hi>.m,.hi>.sk{display:none}.stats{gap:2px}.stats span{font-size:9.5px;padding:0 3px}.stats span:nth-child(-n+3){display:none}
 .tac{gap:4px;margin-top:3px}.tac img{width:20px;height:28px}.tac .t{font-size:11px;line-height:1.2;word-break:keep-all}.tac .k{display:none}.eqp{font-size:10px}.src{font-size:10px}
 .gems{margin-bottom:0}.rk{transform:scale(.85)}}
.h{background:var(--card2);border:1px solid var(--line);border-radius:10px;padding:8px;display:flex;gap:10px}
.pic{position:relative;width:96px;height:150px;flex:none;border-radius:6px;overflow:hidden;background:#0a1110}.pic img.p{width:100%;height:100%;object-fit:cover}
.pic .lb{position:absolute;left:0;right:0;bottom:0;background:linear-gradient(transparent,rgba(0,0,0,.85));color:#fff;text-align:center;font-size:12px;padding:16px 0 3px}
.gems{display:flex;justify-content:center;gap:1px;margin-bottom:1px}.gems i{width:13px;height:17px;background:var(--st) center/contain no-repeat;filter:grayscale(1) brightness(.5);display:block}.gems i.on{filter:none}
.hi{flex:1;min-width:0}.hn{font-size:16px;font-weight:700;display:flex;align-items:center;gap:6px}.nat{font-size:11px;padding:0 6px;border-radius:5px;color:#10201a}
.rk{margin-left:auto;background:var(--jade);color:#06231c;border-radius:8px;width:22px;height:22px;display:inline-flex;align-items:center;justify-content:center;font-size:12px}
.m{color:var(--mute);font-size:11px}.sk{color:var(--gold);font-size:12px}.stats{display:flex;gap:4px;flex-wrap:wrap;margin:3px 0}.stats span{background:var(--card);border:1px solid var(--line);border-radius:5px;padding:0 5px;font-size:11px}.stats b{color:var(--jade)}.stats em{font-style:normal;color:#4cd37a}
.tac{display:flex;gap:6px;align-items:center;margin-top:5px}.tac img{width:26px;height:38px;border-radius:3px;object-fit:cover;background:#0a1110}.tac .t{font-size:13px;font-weight:600}.tac .k{font-size:11px}
.dot{color:#f0966c}.burst{color:#ea5a42}.norm{color:var(--mute)}.heal{color:#5ecfb1}.sup{color:#7fa6d1}.src{color:#f0966c;font-size:11px}.empty{color:var(--mute);font-size:12px;border:1px dashed var(--line);border-radius:6px;padding:6px;margin-top:5px}
.flow{margin-top:10px;font-size:12px;line-height:1.7}.flow div{display:flex;gap:8px}.flow b{width:64px;flex:none;color:var(--gold)}.flow span{color:var(--txt)}
table{width:100%;border-collapse:collapse;font-size:13px}th{color:var(--mute);font-size:11px;text-align:left;padding:6px;border-bottom:1px solid var(--line)}td{padding:7px 6px;border-bottom:1px solid var(--line)}
.pin{background:var(--card2);border:1px solid var(--line);color:var(--txt);border-radius:6px;padding:3px 8px;cursor:pointer;font-size:12px}.pin.on{background:var(--gold);color:#2b2000;border-color:var(--gold);font-weight:700}
h2{font-size:15px;color:var(--gold);margin:18px 0 8px}.chip{display:inline-block;background:var(--card2);border:1px solid var(--line);border-radius:12px;padding:0 8px;margin:1px;font-size:12px}
.fm{display:inline-flex;align-items:center;gap:4px;margin-right:6px}.fg{display:inline-grid;grid-template-columns:repeat(3,7px);gap:1px}.fg i{width:7px;height:7px;border-radius:2px;background:#cdd6d2}.fg i.on{background:#d9a94a}
.eqp{font-size:11.5px;color:#6b7c76;margin-top:3px}.eqp b{color:#8a6a1c}
.trow{margin-top:6px}.trow .rl{display:none}.tc{background:var(--card2);border:1px solid var(--line);border-radius:8px;padding:5px 6px;min-width:0}.tc .tac{margin-top:0}
.tinfo{margin-top:6px}.tinfo summary{cursor:pointer;font-size:13px;font-weight:700;color:var(--gold);padding:4px 0}.tinfo .tc{white-space:normal;line-height:1.4}
.bs{border:1px solid var(--line);background:var(--card2);border-radius:8px;padding:5px 7px;display:flex;align-items:center;gap:5px;min-height:44px;cursor:pointer;position:relative;font-size:13px;min-width:0}.bs.empty{color:var(--mute);justify-content:center}.bs.dis{opacity:.4;cursor:default}.bs.sel{border-color:var(--jade);box-shadow:0 0 0 2px var(--jade)}.bs.t span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:12px}.bs .x{margin-left:auto;font-style:normal;color:var(--mute);font-size:11px;padding:2px 4px}.bs b{font-size:13px}
.kw{display:block;font-size:10px;color:var(--gold)}.pdet{position:sticky;top:0;background:var(--card);border:1px solid var(--jade);border-radius:10px;padding:8px;margin:6px 0}.pdet .dsc{font-size:12.5px;line-height:1.5;margin:4px 0}.oc.pd{border-color:var(--gold);box-shadow:0 0 0 2px var(--gold)}.picker{position:sticky;top:0;z-index:5;background:var(--bg);border:2px solid var(--jade);border-radius:12px;padding:8px;margin-bottom:12px;max-height:46vh;overflow:auto}
.mk .nm{font-size:15px}.strat{background:#f0c75e;color:#2b2000;border-radius:10px;padding:1px 8px;font-size:11px;font-weight:700}.mu{display:grid;grid-template-columns:22px repeat(3,minmax(0,1fr));gap:6px;align-items:start;margin-top:5px}.badge.s{padding:0 6px;text-align:center;height:20px}.mh{display:flex;gap:5px;font-size:12px;min-width:0}.mh .th{flex:none}.mh.no .th{opacity:.45}.mt{font-size:11px;color:var(--jade);line-height:1.3}.mt.no{color:var(--mute);text-decoration:line-through}.x1{color:#c0504d;font-size:10px}.vb:disabled{opacity:.5;cursor:default}
@media(max-width:760px){.mu{grid-template-columns:18px repeat(3,minmax(0,1fr));gap:3px}.mh{flex-direction:column;gap:2px}.bs{padding:3px;font-size:11px;flex-direction:column;align-items:flex-start;gap:2px}.bs .x{position:absolute;top:0;right:0}.picker{max-height:50vh}}
.fr{width:100%;height:78vh;border:1px solid var(--line);border-radius:12px;background:#fff}.vs{width:2px;height:24px;background:var(--line);margin:0 2px}.tab.late{background:#f7f2e4}.tab.mt,.tab.lt{font-weight:700;background:var(--card2)}.tab.lt{background:#efe6c8}.mo{font-weight:700}
.ocs{display:flex;flex-wrap:wrap;gap:6px}.oc{display:flex;align-items:center;gap:6px;text-align:left}.oc .nm{font-weight:600}.oc .m{display:block;font-size:11px;color:var(--mute)}.th{width:33px;height:42px;border-radius:5px;object-fit:cover}.ts{width:30px;height:30px;border-radius:5px;object-fit:cover}.sts{display:flex;gap:3px;margin-top:3px}.sg{width:14px;height:18px;background:var(--st) center/contain no-repeat;filter:grayscale(1) brightness(.45);cursor:pointer;display:block}.sg.on{filter:none}.gh{flex:0 0 100%;font-size:12px;color:var(--gold);margin:6px 0 0;border-bottom:1px solid var(--line)}.oc{background:var(--card2);border:1px solid var(--line);color:var(--mute);border-radius:8px;padding:4px 9px;cursor:pointer;font-size:13px;opacity:.7}.oc.on{opacity:1;color:var(--txt);border-color:#2e9f5f;background:rgba(46,159,95,.15);font-weight:600}.sq{background:var(--card);border:1px solid var(--line);color:var(--txt);border-radius:8px;padding:7px 10px;font-size:13px;min-width:160px}\n.rw{cursor:pointer}.rw:hover{background:var(--card2)}.dt td{background:var(--card2)}.dr{display:flex;gap:10px;margin:6px 0}.dr>b{width:70px;flex:none;color:var(--gold);font-size:12px}.dh{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;flex:1}.dc{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:6px;line-height:2}.ownt{color:#2e9f5f}.tk{display:inline-flex;align-items:center;gap:3px;border:1px solid var(--line);border-radius:6px;padding:0 5px;font-size:12px;background:var(--card2)}.tk.own{border-color:#2e9f5f;color:#2e9f5f;font-weight:700}.tk img{width:16px;height:23px;object-fit:cover;border-radius:2px}
@media(max-width:760px){.dh{grid-template-columns:1fr}}
.err{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:20px;text-align:center}
</style></head><body><div class="wrap" id="app"></div>
<script>
const D=__DATA__;const ERR=__ERR__;const $=s=>document.querySelector(s);
const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const NAT={'촉':'#50aa6e','위':'#648cdc','오':'#dc645a','군':'#aaaaaa'};const KC={'한방딜':'burst','도트딜':'dot','일반딜':'norm','회복':'heal','지원':'sup'};
let MODE='pvp';try{MODE=sessionStorage.getItem('rmode')||'pvp'}catch(e){}
const FR={dogam:'장수 도감',skdex:'전법 도감',chart:'상성 차트',mx:'상성표',fill:'부족한 것 메우기',guide:'상황 가이드'};
let EQ='';let TAB='rank';try{TAB=sessionStorage.getItem('rtab')||'rank';if(TAB==='mine')TAB='rank'}catch(e){}
async function post(o){try{const r=await fetch('/api/reco',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({t:D.t,season:D.season,...o})});const j=await r.json();if(!j.ok)alert(j.error||'실패');else location.reload()}catch(e){alert('네트워크 오류')}}
function US(){return MODE==='dev'&&D.dev?D.dev.units:D.units}
function setMode(m){MODE=m;try{sessionStorage.setItem('rmode',m)}catch(e){}draw()}
function vote(i,scope,kind){const u=US()[i];if(scope==='h')post({op:'vote',scope:'h',key:u.hk,kind,heroes:u.heroes.filter(h=>h.n).map(h=>h.n)});else post({op:'vote',scope:'t',key:u.tk,kind,pairs:u.heroes.flatMap(h=>h.n?h.tac.map(t=>[h.n,t.n]):[]),tset:u.tset})}
function tg(k){const e=document.getElementById('dt'+k),a=document.getElementById('ar'+k);const o=e.style.display==='none';e.style.display=o?'':'none';a.textContent=o?'▾':'▸'}
let TRS=null,OWNS=null,OQ='',OF='all',DIRTY=false;
async function setTr(name,k){TRS=TRS||Object.assign({},D.own.tr||{});const n=(TRS[name]||0)===k?k-1:k;
  try{const r=await fetch('/api/reco',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({t:D.t,season:D.season,op:'tr',name,n})});const j=await r.json();if(!j.ok){alert(j.error||'실패');return}}catch(e){alert('네트워크 오류');return}
  TRS[name]=n;DIRTY=true;const y=scrollY;draw();scrollTo(0,y)}
async function toggleOwn(kind,name){OWNS=OWNS||{h:new Set(D.own.h),s:new Set(D.own.s)};const on=!OWNS[kind].has(name);
  try{const r=await fetch('/api/reco',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({t:D.t,season:D.season,op:'own',kind,name,on})});const j=await r.json();if(!j.ok){alert(j.error||'실패');return}}catch(e){alert('네트워크 오류');return}
  on?OWNS[kind].add(name):OWNS[kind].delete(name);DIRTY=true;const y=scrollY;draw();scrollTo(0,y)}
function setQ(v){OQ=v;const e=document.getElementById('ol');if(e)e.innerHTML=ownList()}
function setF(f){OF=f;draw()}
function fm(s){const ns=(s||'').split(/[\/·,]/).map(x=>({'안영진':'안형진','추행진':'추형진','기영진':'기형진'}[x.trim()]||x.trim())).filter(x=>D.forms[x]);
  if(!ns.length)return s?esc(s):'';return ns.map(n=>`<span class="fm" title="${esc(n)}"><span class="fg">${D.forms[n].flat().map(v=>`<i class="${v?'on':''}"></i>`).join('')}</span>${esc(n)}</span>`).join(' ')}
function pin(name,on){post({op:'pin',deck:name,on})}
function setTab(t){if(DIRTY&&t==='build'){DIRTY=false;try{sessionStorage.setItem('rtab','build')}catch(e){}location.reload();return}TAB=t;try{sessionStorage.setItem('rtab',t)}catch(e){}draw()}
function vg(label,v,i,scope){return `<span class="vg">${label}<button class="vb u ${v.m==='u'?'on':''}" onclick="vote(${i},'${scope}','u')">👍 ${v.u}</button><button class="vb d ${v.m==='d'?'on':''}" onclick="vote(${i},'${scope}','d')">👎 ${v.d}</button></span>`}
function tcell(t){return `<div class="tac">${t.img?`<img src="data:image/jpeg;base64,${t.img}">`:'<img>'}<div><div class="t">${esc(t.n)}</div><div class="k"><span class="${KC[t.kind]||''}">${esc(t.kind)}</span> <span class="m">${esc(t.type)}·${esc(t.trait)}</span></div></div></div>`}
function hero(h){
  if(!h.n)return `<div class="h"><div class="empty">${esc(h.src||'보유 장수 부족')}</div></div>`;
  const gem=h.tr>0?`<div class="gems" style="--st:url(${D.stone})">${[1,2,3,4,5].map(k=>`<i class="${h.tr>=k?'on':''}"></i>`).join('')}</div>`:'';
  const L=['무력','지력','통솔','선공'];
  const st=h.fin?L.map((l,k)=>`<span>${l} <b>${h.fin[k]}</b>${h.add[k]?`<em> (+${h.add[k]})</em>`:''}</span>`).join(''):'';
  const tac=h.tac.length?h.tac.map(t=>`<div class="tac">${t.img?`<img src="data:image/jpeg;base64,${t.img}">`:'<img>'}<div><div class="t">${esc(t.n)}</div><div class="k"><span class="${KC[t.kind]||''}">${esc(t.kind)}</span> <span class="m">${esc(t.type)}·${esc(t.trait)} · ${esc(t.why)}</span></div></div></div>`).join(''):`<div class="empty">전법 부족 (보유 전법 없음)</div>`;
  return `<div class="h"><div class="pic">${h.img?`<img class="p" src="data:image/jpeg;base64,${h.img}">`:''}<div class="lb">${gem}50 ${esc(h.n)}</div></div><div class="hi">
  <div class="hn">${esc(h.n)}${h.nation?`<span class="nat" style="background:${NAT[h.nation]||'#999'}">${esc(h.nation)}</span>`:''}${h.rank?`<span class="rk" title="행동 순서">${h.rank}</span>`:''}</div>
  <div class="m">${esc([h.pos,h.troop].filter(Boolean).join(' · '))}</div><div class="sk">${esc(h.sk)}${h.sktype?` (${esc(h.sktype)})`:''}</div>
  <div class="stats">${st}</div>${h.eq?`<div class="eqp" title="특기 대체: ${esc(h.eq.alt)}">🛡 특기 <b>${esc(h.eq.trait)}</b> · 이펙트 <b>${esc(h.eq.effects.join(', '))}</b></div>`:''}${h.src?`<div class="src">← ${esc(h.src)}</div>`:''}</div></div>`}
function unit(u,i){
  const fl=u.flow,s=fl.start;
  const lines=[`<div><b>전투 시작</b><span>지휘 ${esc(s['지휘'].join(' · ')||'-')} → 패시브 ${esc(s['패시브'].join(' · ')||'-')}</span></div>`].concat(fl.turn.map((t,k)=>`<div><b>턴 ${k+1}번째</b><span>${esc(t[0])}${t[1]!=null?`(${t[1]})`:''} 액티브 ${esc(t[2].join(' · ')||'-')} → 평타${t[3].length?` → 추격 ${esc(t[3].join(' · '))}`:''}</span></div>`));
  if(u.status.g.length)lines.push(`<div><b>이상상태</b><span class="m">부여 ${esc(u.status.g.join(', '))}${u.status.miss.length?` · 부족 ${esc(u.status.miss.join(', '))}`:''}</span></div>`);
  return `<div class="unit"><div class="uh"><span class="badge">${u.i}부대</span>${u.name?`<span class="m">${esc(u.tier)}</span><span class="nm">${esc(u.name)}</span><span class="tr">${fm(u.form)}</span>`:'<span class="nm">자유 편성</span>'}
  ${u.name?`<div class="votes">${vg('장수 조합',u.hv,i,'h')}${u.tset.length?vg('전법 조합',u.tv,i,'t'):''}</div>`:''}</div>
  <div class="hs">${u.heroes.map(hero).join('')}</div>${[0,1].map(k=>`<div class="hs trow"><div class="rl">전법${k+1}</div>${u.heroes.map(h=>`<div class="tc">${h.n&&h.tac[k]?tcell(h.tac[k]):'<span class="m">-</span>'}</div>`).join('')}</div>`).join('')}${tinfo(u)}<div class="flow">${lines.join('')}</div></div>`}
function tinfo(u){if(!u.name||!u.heroes.some(h=>h.ti&&(h.ti['스탯']||[]).length))return'';
  const LB={'스탯':'스탯 투자','병법':'병법','장비':'장비 스탯'};
  return `<details class="tinfo"><summary>📋 티어덱 정보 — 스탯·병법·장비${u.note?` <span class="m">${esc(u.note)}</span>`:''}</summary>`+['스탯','병법','장비'].map(k=>`<div class="hs trow"><div class="rl">${LB[k]}</div>${u.heroes.map(h=>`<div class="tc"><span class="m"><b>${LB[k]}</b> ${esc(((h.ti||{})[k]||[]).join(' · ')||'-')}</span></div>`).join('')}</div>`).join('')+`</details>`}
function mine(){const dv=MODE==='dev';return `<div class="bar"><button class="sbtn mo ${!dv?'on':''}" onclick="setMode('pvp')">PvP 추천</button><button class="sbtn mo ${dv?'on':''}" onclick="setMode('dev')">개척 추천</button><span class="m">${dv?'개척덱 기준':'PvP 티어덱 기준'} · 내 보유 장수·전법으로 5부대</span></div>`+US().map(unit).join('')+`<div class="m">티어덱을 완성하지 못한 부대는 같은 역할의 보유 장수로 채웁니다. 장수 👎는 그 덱을 빼고, 전법 👎는 전법만 다시 고릅니다. 모든 서버의 평가가 합산됩니다.</div>`}
function rank(){const R=D.rank;
  const own=new Set(D.ownsk||[]),nk=s=>s.replace(/[\s·\-]/g,'');
  const rhero=h=>`<div class="h"><div class="pic">${h.img?`<img class="p" src="data:image/jpeg;base64,${h.img}">`:''}<div class="lb">50 ${esc(h.n)}</div></div><div class="hi">
    <div class="hn">${esc(h.n)}${h.nation?`<span class="nat" style="background:${NAT[h.nation]||'#999'}">${esc(h.nation)}</span>`:''}</div>
    <div class="m">${esc([h.pos,h.troop].filter(Boolean).join(' · '))}${h.alts.length?` · 대체 ${esc(h.alts.join('/'))}`:''}</div><div class="sk">${esc(h.sk)}${h.sktype?` (${esc(h.sktype)})`:''}</div>
    ${h.tac.map(t=>`<div class="tac">${t.img?`<img src="data:image/jpeg;base64,${t.img}">`:'<img>'}<div><div class="t ${own.has(nk(t.n))?'ownt':''}">${esc(t.n)}${own.has(nk(t.n))?' ✓':''}</div><div class="k"><span class="${KC[t.kind]||''}">${esc(t.kind)}</span> <span class="m">${esc([t.type,t.trait].filter(Boolean).join('·'))}</span></div>${t.alt.length?`<div class="m">대체: ${esc(t.alt.join(' / '))}</div>`:''}</div></div>`).join('')}
    ${h.tac.length<2?'<div class="empty">추천 전법 정보 없음</div>':''}</div></div>`;
  const detail=d=>{const cb=R.tacs.filter(t=>t.deck===d.name);const rated=cb.length?`<div class="flow"><div><b>평가된 전법 조합</b><span>${cb.map(t=>`${t.tset.map(x=>`<span class="chip">${esc(x)}</span>`).join('')} <span class="m">👍 ${t.u} 👎 ${t.d}</span>`).join('<br>')}</span></div></div>`:'';
    return `<div class="m" style="margin-bottom:6px">${fm(d.form)}${d.note?' · '+esc(d.note):''} · <span style="color:#2e9f5f;font-weight:700">✓</span> = 내가 보유한 전법 · 장수마다 추천 전법 2개</div><div class="hs">${d.cards.map(rhero).join('')}</div>${rated}`};
  const sorted=R.decks.map((d,k)=>({d,k})).sort((a,b)=>tierNum(a.d.tier)-tierNum(b.d.tier)||a.d.idx-b.d.idx);let lastT=null;
  const stt=d=>{const st=deckStat(d);const miss=[...st.mh.map(m=>`<div>🔸 장수 <b>${esc(m.want)}</b> 없음 → 대체 후보 ${m.sub.length?m.sub.map(esc).join(', '):'없음'}</div>`),...st.mt.map(m=>`<div>🔹 ${esc(m.hero)}의 전법 <b>${esc(m.want)}</b> 없음 → 대체 후보 ${m.sub.length?m.sub.map(esc).join(', '):'없음'}</div>`)].join('');return{st,miss}};
  const dk=`<table><tr><th>티어</th><th>덱 (눌러서 전법·부족한 것 보기)</th><th>내 상황</th><th>👍</th><th>👎</th><th>내 추천</th></tr>${sorted.map(({d,k})=>{const{st,miss}=stt(d);const g=d.tier!==lastT;lastT=d.tier;
    const badge=st.full?'<span class="ownt"><b>✓ 바로 편성 가능</b></span>':`<span class="m">장수 ${st.hh}/${st.th} · 전법 ${st.ht}/${st.tt}</span>`;
    return `<tr class="rw" onclick="tg(${k})"><td><b>${g?esc(d.tier):''}</b></td><td><b>${esc(d.name)}</b> <span class="m" id="ar${k}">▸</span><div class="m">${esc(d.heroes.join(' · '))}</div></td><td>${badge}</td><td>${d.u}</td><td>${d.d}</td><td><button class="pin ${d.pin?'on':''}" onclick="event.stopPropagation();pin('${esc(d.name).replace(/'/g,"\\'")}',${!d.pin})">${d.pin?'📌 적용 중':'적용'}</button></td></tr><tr class="dt" id="dt${k}" style="display:none"><td></td><td colspan="5">${miss?`<div class="flow"><div><b>부족한 것 메우기</b><span>${miss}</span></div></div>`:''}${detail(d)}</td></tr>`}).join('')}</table>`;
  const tc=R.tacs.length?`<table><tr><th>덱</th><th>전법 조합</th><th>👍</th><th>👎</th></tr>${R.tacs.map(t=>`<tr><td>${esc(t.deck)}</td><td>${t.tset.map(x=>`<span class="chip">${esc(x)}</span>`).join('')}</td><td>${t.u}</td><td>${t.d}</td></tr>`).join('')}</table>`:'<div class="m">아직 평가된 전법 조합이 없습니다.</div>';
  const KN={ht:'장수↔전법',tt:'전법↔전법',hh:'장수↔장수'};
  const sy=R.syn.length?`<table><tr><th>종류</th><th>궁합</th><th>점수</th><th>메모</th></tr>${R.syn.map(s=>`<tr><td>${KN[s.kind]||s.kind}</td><td><b>${esc(s.a)}</b> ↔ <b>${esc(s.b)}</b></td><td>${s.score>0?'+':''}${s.score}</td><td class="m">${esc(s.note)}${s.src==='vote'?' (평가로 학습)':''}</td></tr>`).join('')}</table>`:'<div class="m">저장된 시너지가 없습니다.</div>';
  return `<h2>티어덱 — 티어 순서 · 내 보유 기준 부족한 것</h2><div class="m" style="margin-bottom:6px">티어별로 묶었고, <b>내 상황</b>은 내가 가진 장수·전법으로 이 덱을 얼마나 채울 수 있는지입니다. 덱을 누르면 부족한 장수·전법과 <b>대체 후보</b>가 나옵니다. 5부대를 한 번에 짜려면 <b>편성기</b>의 <b>✨ 내 상황 맞춤 자동 편성</b>을 쓰세요.</div>${dk}
  <h2>전법 조합 랭킹</h2>${tc}<h2>시너지 점수 TOP — 장수·전법 궁합(추천 점수에 더해집니다)</h2>${sy}`}
function ownList(){TRS=TRS||Object.assign({},D.own.tr||{});OWNS=OWNS||{h:new Set(D.own.h),s:new Set(D.own.s)};const q=OQ.replace(/\s/g,'');
  const keep=(n,o,ss,kwx)=>(ss||1)<=D.season&&(!q||n.replace(/\s/g,'').includes(q)||(kwx||'').replace(/\s/g,'').includes(q))&&(OF==='all'||(OF==='own')===o);
  const GR={'전설':0,'영웅':1},ko=(a,b)=>a.n.localeCompare(b.n,'ko');
  const hs=D.all.h.filter(h=>keep(h.n,OWNS.h.has(h.n),h.season)).sort(ko);
  const ss=D.all.s.filter(s=>keep(s.n,OWNS.s.has(s.n),s.season,(s.kw||[]).join('')+(s.trait||'')+(s.desc||''))).sort((a,b)=>((GR[a.grade]??9)-(GR[b.grade]??9))||ko(a,b));
  const im=(b,c)=>b?`<img class="${c}" src="data:image/jpeg;base64,${b}">`:'';
  const stones=h=>{const n=TRS[h.n]||0;return `<span class="sts" title="초월 단계" style="--st:url(${D.stone})">${[1,2,3,4,5].map(k=>`<i class="sg ${k<=n?'on':''}" onclick="event.stopPropagation();setTr('${esc(h.n)}',${k})"></i>`).join('')}</span>`};
  const card=h=>`<div class="oc hc ${OWNS.h.has(h.n)?'on':''}" onclick="toggleOwn('h','${esc(h.n)}')">${im(h.img,'th')}<span><span class="nm">${esc(h.n)}</span>${OWNS.h.has(h.n)?stones(h):''}</span></div>`;
  const NO={'촉':0,'위':1,'오':2,'군':3},NN={'촉':'촉','위':'위','오':'오','군':'군웅'};let hc='';
  for(const nt of ['촉','위','오','군']){const g=hs.filter(h=>h.nation===nt);if(g.length)hc+=`<div class="gh">${NN[nt]} (${g.length}명 · 보유 ${g.filter(h=>OWNS.h.has(h.n)).length})</div>`+g.map(card).join('')}
  const oth=hs.filter(h=>!(h.nation in NO));if(oth.length)hc+=`<div class="gh">기타</div>`+oth.map(card).join('');
  let sc='',lg=null;for(const s of ss){if(s.grade!==lg){lg=s.grade;sc+=`<div class="gh">${esc(lg||'기타')}</div>`}
   sc+=`<button class="oc sc ${OWNS.s.has(s.n)?'on':''}" onclick="toggleOwn('s','${esc(s.n).replace(/'/g,"\\'")}')">${im(s.img,'ts')}<span class="nm">${esc(s.n)}</span><span class="m">${esc(s.type)}·${esc(s.trait)}</span></button>`}
  return `<h2>장수 (보유 ${OWNS.h.size})</h2><div class="ocs">${hc||'<span class="m">없음</span>'}</div><h2>전법 (보유 ${OWNS.s.size})</h2><div class="ocs">${sc||'<span class="m">없음</span>'}</div>`}
function own(){TRS=TRS||Object.assign({},D.own.tr||{});OWNS=OWNS||{h:new Set(D.own.h),s:new Set(D.own.s)};
  return `<div class="m" style="margin-bottom:8px">눌러서 보유 여부를 바꿉니다(자동 저장). 보유 장수의 <b>붉은 돌</b>을 눌러 초월 단계(0~5)를 정합니다. 스탯 조율은 <b>장수 도감</b> 탭에서 합니다. 바꾼 뒤 <b>내 추천덱</b> 탭으로 돌아가면 다시 편성합니다.</div>
  <div class="bar"><input id="oq" class="sq" placeholder="이름·키워드 검색 (예: 위협, 치유, 홍수)" value="${esc(OQ)}" oninput="setQ(this.value)"><button class="sbtn ${OF==='all'?'on':''}" onclick="setF('all')">전체</button><button class="sbtn ${OF==='own'?'on':''}" onclick="setF('own')">보유</button><button class="sbtn ${OF==='no'?'on':''}" onclick="setF('no')">미보유</button></div><div id="ol">${ownList()}</div>`}
function eqBody(){const E=D.equip,PC={'필수':'#c0392b','권장':'#0f8a63','상황별':'#8a6a1c','하위':'#8a9a94'};const q=EQ.replace(/\s/g,'');
  const tbl=(rows,lab)=>`<table><tr><th>${lab}</th><th>우선순위</th><th>효과</th><th>평가</th></tr>${rows.filter(r=>!q||(r[0]+r[2]+r[3]).replace(/\s/g,'').includes(q)).map(r=>`<tr><td><b>${esc(r[0])}</b></td><td><span class="chip" style="border-color:${PC[r[1]]};color:${PC[r[1]]}">${esc(r[1])}</span></td><td>${esc(r[2])}</td><td class="m">${esc(r[3])}</td></tr>`).join('')}</table>`;
  return `<h2>장비 특기</h2>${tbl(E.traits,'특기')}<h2>이펙트</h2>${tbl(E.effects,'이펙트')}`}
function equip(){return `<div class="m" style="margin-bottom:8px">장비 특기 15종·이펙트 19종. 추천 부대 카드의 <b>🛡 특기·이펙트</b>는 장수 역할(병기·책략 딜러/탱커/서포터)과 쓰는 전법 종류로 고른 추천입니다. 평가는 중국 서버 기준이라 한국 서버와 다를 수 있습니다. (출처: cheonha-deck.xyz)</div>
  <div class="bar"><input class="sq" placeholder="특기·이펙트 검색" value="${esc(EQ)}" oninput="EQ=this.value;document.getElementById('el').innerHTML=eqBody()"></div>
  <div id="el">${eqBody()}</div>`}
function frame(v){return `<iframe class="fr" src="/tier?t=${encodeURIComponent(D.t)}&s=${D.season}&embed=1#${v}"></iframe>`}
let MENU=true,LATE=true;try{MENU=sessionStorage.getItem('rmenu')!=='0';LATE=sessionStorage.getItem('rlate')!=='0'}catch(e){}
function toggleMenu(){MENU=!MENU;try{sessionStorage.setItem('rmenu',MENU?'1':'0')}catch(e){}draw()}
function toggleLate(){LATE=!LATE;try{sessionStorage.setItem('rlate',LATE?'1':'0')}catch(e){}draw()}

// ---------- 편성기 · 공유마켓 · 지략가 ----------
const HM=Object.fromEntries(D.all.h.map(h=>[h.n,h])),SM=Object.fromEntries(D.all.s.map(x=>[x.n,x])),SNK=Object.fromEntries(D.all.s.map(x=>[x.n.replace(/[\s·\-]/g,''),x.n]));
const T0=()=>({h:['','',''],t:[['',''],['',''],['','']]});
let BLD=(D.draft||[]).map(u=>({h:[...u.h],t:[0,1,2].map(k=>{const a=[...((u.t||[])[k]||[])];while(a.length<2)a.push('');return a})}));while(BLD.length<5)BLD.push(T0());
let PK=null,PQ='',MS='like',MF=false,MT='',MN='',SVT=null;
const ownH=()=>OWNS?OWNS.h:new Set(D.own.h),ownS=()=>OWNS?OWNS.s:new Set(D.own.s);
const usedH=()=>new Set(BLD.flatMap(u=>u.h).filter(Boolean)),usedS=()=>new Set(BLD.flatMap(u=>u.t.flat()).filter(Boolean));
const imj=(b,c)=>b?`<img class="${c}" src="data:image/jpeg;base64,${b}">`:'';
function saveDraft(){clearTimeout(SVT);SVT=setTimeout(()=>{fetch('/api/reco',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({t:D.t,season:D.season,op:'draft',units:BLD.some(u=>u.h.some(Boolean))?BLD.map(u=>({h:u.h,t:u.t.map(a=>a.filter(Boolean))})):null})}).catch(()=>{})},700)}
async function mpost(o){try{const r=await fetch('/api/reco',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({t:D.t,season:D.season,...o})});const j=await r.json();if(!j.ok){alert(j.error||'실패');return false}if(j.mk)D.mk=j.mk;return true}catch(e){alert('네트워크 오류');return false}}
function pick(u,k,kind,j){PD=null;PK={u,k,kind,j};PQ='';draw()}
function setPQ(v){PQ=v;const e=document.getElementById('pl');if(e)e.innerHTML=pickBody()}
function assign(name){PD=null;const p=PK;if(!p)return;const u=BLD[p.u];if(p.kind==='h')u.h[p.k]=name;else u.t[p.k][p.j]=name;PK=null;saveDraft();const y=scrollY;draw();scrollTo(0,y)}
function clr(u,k,kind,j){if(kind==='h'){BLD[u].h[k]='';BLD[u].t[k]=['','']}else BLD[u].t[k][j]='';PK=null;saveDraft();draw()}
function clearAll(){if(!confirm('5부대를 모두 비울까요?'))return;BLD=[0,1,2,3,4].map(T0);PK=null;saveDraft();draw()}
function fillTier(u,name){const d=D.rank.decks.find(x=>x.name===name);if(!d)return;const o=BLD[u];BLD[u]=T0();const uh=usedH(),us=usedS(),oh=ownH(),os=ownS();
  d.slots.forEach((slot,k)=>{const h=slot.split('/').find(x=>oh.has(x)&&!uh.has(x)&&HM[x]&&(HM[x].season||1)<=D.season);if(!h)return;BLD[u].h[k]=h;uh.add(h);
    ((d.tc||[])[k]||[]).slice(0,2).forEach((cell,j)=>{const t=cell.split(/[>\/]/).map(x=>SNK[x.trim().replace(/[\s·\-]/g,'')]).find(x=>x&&os.has(x)&&!us.has(x)&&(SM[x].season||1)<=D.season);if(t){BLD[u].t[k][j]=t;us.add(t)}})});
  saveDraft();draw()}
function slotH(u,k){const n=BLD[u].h[k],on=PK&&PK.u===u&&PK.k===k&&PK.kind==='h';
  return n?`<div class="bs ${on?'sel':''}" onclick="pick(${u},${k},'h')">${imj((HM[n]||{}).img,'th')}<b>${esc(n)}</b><i class="x" onclick="event.stopPropagation();clr(${u},${k},'h')">✕</i></div>`:`<div class="bs empty ${on?'sel':''}" onclick="pick(${u},${k},'h')">＋ 장수</div>`}
function slotT(u,k,j){const n=BLD[u].t[k][j],on=PK&&PK.u===u&&PK.k===k&&PK.kind==='t'&&PK.j===j,hh=BLD[u].h[k];
  if(!hh)return`<div class="bs empty dis">전법${j+1}</div>`;
  return n?`<div class="bs t ${on?'sel':''}" onclick="pick(${u},${k},'t',${j})">${imj((SM[n]||{}).img,'ts')}<span>${esc(n)}</span><i class="x" onclick="event.stopPropagation();clr(${u},${k},'t',${j})">✕</i></div>`:`<div class="bs empty ${on?'sel':''}" onclick="pick(${u},${k},'t',${j})">＋ 전법${j+1}</div>`}

const CAN_HOVER=window.matchMedia&&matchMedia('(hover:hover) and (pointer:fine)').matches;let PD=null;
const tipOf=x=>`${x.n} · ${x.grade}·${x.type}·${x.trait}${(x.kw||[]).length?'\n키워드: '+x.kw.join(', '):''}\n${x.desc||''}`;
function tacClick(n){if(CAN_HOVER){assign(n);return}PD=PD===n?null:n;const e=document.getElementById('pl'),d=document.getElementById('pdet');if(e)e.innerHTML=pickBody();if(d)d.innerHTML=pdetBody()}
function pdetBody(){if(!PD||!PK||PK.kind!=='t')return'';const x=SM[PD];if(!x)return'';return `<div class="pdet"><div class="nm">${imj(x.img,'ts')} <b>${esc(x.n)}</b> <span class="m">${esc(x.grade)}·${esc(x.type)}·${esc(x.trait)}</span></div>${(x.kw||[]).length?`<div class="kw" style="margin:2px 0">키워드: ${x.kw.map(esc).join(' · ')}</div>`:''}<div class="dsc">${esc(x.desc||'설명 없음')}</div><div class="bar"><button class="sbtn on" onclick="const n=PD;PD=null;assign(n)">✔ 적용</button><button class="sbtn" onclick="PD=null;document.getElementById('pdet').innerHTML='';document.getElementById('pl').innerHTML=pickBody()">취소</button></div></div>`}
function pickBody(){const q=PQ.replace(/\s/g,''),GR={'전설':0,'영웅':1};let body='';
  const grp=(l,key,order,lab)=>{const o=[...order];l.forEach(x=>{const k=key(x);if(!o.includes(k))o.push(k)});return o.map(k=>{const g=l.filter(x=>key(x)===k);return g.length?`<div class="gh">${esc(lab(k))} (${g.length})</div>`+g.map(x=>x._c).join(''):''}).join('')};
  if(PK.kind==='h'){const uh=usedH(),nu=new Set(D.nu||[]);
    const l=D.all.h.filter(h=>ownH().has(h.n)&&!uh.has(h.n)&&!nu.has(h.n)&&(h.season||1)<=D.season&&(!q||(h.n+h.nation+h.troop+h.pos).replace(/\s/g,'').includes(q))).sort((a,b)=>a.n.localeCompare(b.n,'ko'))
      .map(h=>({...h,_c:`<button class="oc hc on" onclick="assign('${esc(h.n)}')">${imj(h.img,'th')}<span class="nm">${esc(h.n)}</span><span class="m">${esc(h.troop)}</span></button>`}));
    body=grp(l,h=>h.nation||'기타',['촉','위','오','군'],k=>({'군':'군웅'}[k]||k))||'<span class="m">쓸 수 있는 보유 장수가 없습니다 (이미 다른 부대에서 썼거나 미보유)</span>'}
  else{const us=usedS();
    const l=D.all.s.filter(x=>ownS().has(x.n)&&!us.has(x.n)&&(x.season||1)<=D.season&&(!q||(x.n+x.trait+x.type+x.grade+(x.kw||[]).join('')+(x.desc||'')).replace(/\s/g,'').includes(q))).sort((a,b)=>((GR[a.grade]??9)-(GR[b.grade]??9))||a.n.localeCompare(b.n,'ko'))
      .map(x=>({...x,_c:`<button class="oc sc on ${PD===x.n?'pd':''}" title="${esc(tipOf(x))}" onclick="tacClick('${esc(x.n).replace(/'/g,"\\'")}')">${imj(x.img,'ts')}<span class="nm">${esc(x.n)}</span><span class="m">${esc(x.grade)}·${esc(x.type)}</span>${(x.kw||[]).length?`<span class="kw">${(x.kw||[]).map(esc).join(' · ')}</span>`:''}</button>`}));
    body=grp(l,x=>x.trait||'기타',['책략','병기','방어','치유','보조','문무'],k=>k)||'<span class="m">쓸 수 있는 보유 전법이 없습니다 (이미 다른 부대에서 썼거나 미보유)</span>'}
  return body}
function picker(){if(!PK)return'';return `<div class="picker"><div class="bar"><b>${PK.u+1}부대 ${PK.kind==='h'?'장수':'전법'} 고르기</b><input id="pq" class="sq" placeholder="${PK.kind==='h'?'이름·나라·병종 검색':'이름·계열·키워드 검색 (예: 위협, 치유, 홍수)'}" value="${esc(PQ)}" oninput="setPQ(this.value)"><button class="sbtn" onclick="PK=null;draw()">닫기</button></div><div id="pdet">${pdetBody()}</div><div class="ocs" id="pl">${pickBody()}</div></div>`}

// ---------- 내 상황 맞춤: 부족한 것 메우기 + 티어순 자동 편성 ----------
const tierNum=t=>{const m=String(t||'').match(/T?(\d+(?:\.\d+)?)([+-]?)/);return m?parseFloat(m[1])+(m[2]==='+'?-.1:m[2]==='-'?.1:0):99};
const NUSE=()=>new Set(D.nu||[]);
const cellNames=c=>String(c||'').split(/[>\/]/).map(x=>SNK[x.trim().replace(/[\s·\-]/g,'')]).filter(Boolean);
const simH=(a,b)=>{if(!a||!b)return 0;const sa=new Set((a.style||'').split(/\s+/)),sb=(b.style||'').split(/\s+/);return 2*sb.filter(x=>x&&sa.has(x)).length+(a.pos===b.pos)+(a.troop===b.troop)+(a.nation===b.nation)*.5};
function subH(want,uh){const w=HM[want];return D.all.h.filter(h=>ownH().has(h.n)&&!uh.has(h.n)&&!NUSE().has(h.n)&&(h.season||1)<=D.season).map(h=>[simH(w,h),h]).sort((a,b)=>b[0]-a[0]||a[1].n.localeCompare(b[1].n,'ko'))}
function subT(want,us){const w=SM[want]||{};return D.all.s.filter(x=>ownS().has(x.n)&&!us.has(x.n)&&(x.season||1)<=D.season).map(x=>[(x.trait===w.trait)*2+(x.type===w.type),x]).filter(a=>a[0]>0).sort((a,b)=>b[0]-a[0]||a[1].n.localeCompare(b[1].n,'ko'))}
function deckStat(d,uh,us){uh=uh||new Set();us=us||new Set();const oh=ownH(),os=ownS();let hh=0,ht=0,th=0,tt=0;const mh=[],mt=[];
  d.slots.forEach((slot,k)=>{const alts=slot.split('/');const own=alts.find(x=>oh.has(x)&&!uh.has(x));th++;
    if(own){hh++;((d.tc||[])[k]||[]).slice(0,2).forEach(c=>{const names=cellNames(c);if(!names.length)return;tt++;if(names.some(x=>os.has(x)&&!us.has(x)))ht++;else mt.push({hero:own,want:names[0],sub:subT(names[0],us).slice(0,3).map(a=>a[1].n)})})}
    else mh.push({want:alts[0],sub:subH(alts[0],uh).slice(0,3).map(a=>a[1].n)})});
  return{hh,th,ht,tt,mh,mt,full:hh===th&&ht===tt}}
function autoCompose(){const ds=D.rank.decks.slice().sort((a,b)=>tierNum(a.tier)-tierNum(b.tier)||a.idx-b.idx),uh=new Set(),us=new Set(),oh=ownH(),os=ownS(),res=[],used=new Set();
  for(const minc of [3,2,1])for(const d of ds){if(res.length>=5)break;if(used.has(d.name))continue;
    const have=d.slots.filter(sl=>sl.split('/').some(x=>oh.has(x)&&!uh.has(x)&&!NUSE().has(x)&&HM[x]&&(HM[x].season||1)<=D.season)).length;
    if(have<minc)continue;used.add(d.name);const u=T0();
    d.slots.forEach((slot,k)=>{const h=slot.split('/').find(x=>oh.has(x)&&!uh.has(x)&&!NUSE().has(x)&&HM[x]&&(HM[x].season||1)<=D.season);if(h){u.h[k]=h;uh.add(h)}});
    d.slots.forEach((slot,k)=>{if(u.h[k])return;const c=subH(slot.split('/')[0],uh)[0];if(c){u.h[k]=c[1].n;uh.add(c[1].n)}});
    d.slots.forEach((slot,k)=>{if(!u.h[k])return;const own=slot.split('/').includes(u.h[k]);((d.tc||[])[k]||[]).slice(0,2).forEach((cell,j)=>{const names=cellNames(cell);let t=own?names.find(x=>os.has(x)&&!us.has(x)&&(SM[x].season||1)<=D.season):null;if(!t&&names.length){const c=subT(names[0],us)[0];if(c)t=c[1].n}if(t){u.t[k][j]=t;us.add(t)}})});
    u.src=d.name;res.push(u)}
  while(res.length<5)res.push(T0());BLD=res.slice(0,5);saveDraft();draw()}
function build(){const dk=D.rank.decks.slice().sort((a,b)=>a.idx-b.idx);
  const units=BLD.map((u,i)=>`<div class="unit"><div class="uh"><span class="badge">${i+1}부대</span>${BLD[i].src?`<span class="m">${esc((D.rank.decks.find(x=>x.name===BLD[i].src)||{}).tier||'')} <b>${esc(BLD[i].src)}</b> 기준</span>`:''}<select class="sq" onchange="if(this.value){fillTier(${i},this.value);}this.value=''"><option value="">티어덱 불러오기…</option>${dk.map(d=>`<option value="${esc(d.name)}">${esc(d.tier)} ${esc(d.name)}</option>`).join('')}</select></div>
    <div class="hs">${[0,1,2].map(k=>slotH(i,k)).join('')}</div>${[0,1].map(j=>`<div class="hs trow">${[0,1,2].map(k=>slotT(i,k,j)).join('')}</div>`).join('')}</div>`).join('');
  const n=BLD.flatMap(u=>u.h).filter(Boolean).length;
  return `<div class="m" style="margin-bottom:8px">내가 <b>보유한</b> 장수·전법으로 5부대를 직접 짭니다. <b>한 번 쓴 장수·전법은 다른 부대에서 쓸 수 없습니다</b>(목록에서 사라짐). 편성은 자동 저장됩니다. 보유 체크는 <b>내 장수·전법</b> 탭에서 합니다.</div>
  ${picker()}<div class="bar"><button class="sbtn on" onclick="if(confirm('내 보유 장수·전법에 맞춰 티어 순서대로 5부대를 자동 편성합니다. 지금 편성은 덮어씁니다.'))autoCompose()">✨ 내 상황 맞춤 자동 편성 (티어순)</button><span class="m">부족한 자리는 비슷한 보유 장수·같은 계열 전법으로 메웁니다</span></div>${units}<div class="unit"><div class="uh"><b>마켓에 올리기</b><span class="m">장수 ${n}명 · 올린 글은 모든 서버가 보고 추천할 수 있습니다 (최대 ${D.mk.max}개)</span></div>
  <div class="bar"><input class="sq" style="flex:1;min-width:140px" maxlength="24" placeholder="제목 (예: 도트딜 5부대)" value="${esc(MT)}" oninput="MT=this.value"><input class="sq" style="flex:2;min-width:180px" maxlength="120" placeholder="한 줄 설명 (선택)" value="${esc(MN)}" oninput="MN=this.value"></div>
  <div class="bar"><button class="sbtn on" onclick="postMk()">🛒 마켓에 올리기</button><button class="sbtn" onclick="clearAll()">모두 비우기</button></div></div>`}
async function postMk(){if(!BLD.some(u=>u.h.some(Boolean))){alert('장수를 한 명 이상 넣어 주세요');return}if(!MT.trim()){alert('제목을 입력해 주세요');return}
  if(await mpost({op:'mk_post',title:MT,note:MN,units:BLD.map(u=>({h:u.h,t:u.t.map(a=>a.filter(Boolean))}))})){MT='';MN='';alert('마켓에 올렸습니다');setTab('market')}}
async function likeMk(id){if(await mpost({op:'mk_like',id})){const y=scrollY;draw();scrollTo(0,y)}}
async function delMk(id){if(!confirm('이 편성을 삭제할까요? (받은 추천도 사라집니다)'))return;if(await mpost({op:'mk_del',id}))draw()}
function loadMk(id){const p=D.mk.posts.find(x=>x.id===id);if(!p)return;if(!confirm('내 편성기에 가져옵니다. 보유하지 않은 장수·전법은 빈 칸이 되며 지금 편성은 덮어씁니다. 계속할까요?'))return;
  const oh=ownH(),os=ownS(),uh=new Set(),us=new Set();BLD=p.units.map(u=>{const o=T0();[0,1,2].forEach(k=>{const h=u.h[k];if(h&&oh.has(h)&&!uh.has(h)){o.h[k]=h;uh.add(h);(u.t[k]||[]).slice(0,2).forEach((t,j)=>{if(os.has(t)&&!us.has(t)){o.t[k][j]=t;us.add(t)}})}});return o});while(BLD.length<5)BLD.push(T0());saveDraft();setTab('build')}
function mkStat(p){const oh=ownH(),os=ownS();let a=0,b=0;p.units.forEach(u=>u.h.forEach((h,k)=>{if(h){b++;if(oh.has(h))a++;(u.t[k]||[]).forEach(t=>{b++;if(os.has(t))a++})}}));return[a,b]}
function market(){const os=ownS(),oh=ownH();let l=D.mk.posts.filter(p=>p.season===D.season);if(MF)l=l.filter(p=>{const[a,b]=mkStat(p);return a===b});l=l.slice().sort((a,b)=>MS==='like'?(b.likes-a.likes||b.ts-a.ts):(b.ts-a.ts));
  const card=p=>{const[a,b]=mkStat(p);return `<div class="unit mk"><div class="uh"><span class="nm">${esc(p.title)}</span><span class="m">by <b>${esc(p.by)}</b>${p.strat?' <span class="strat">🧠 지략가</span>':''}</span>
    <div class="votes"><span class="m">내 보유 ${a}/${b}</span><button class="vb u ${p.liked?'on':''}" ${p.mine?'disabled title="내 글"':''} onclick="likeMk('${p.id}')">👍 ${p.likes}</button><button class="sbtn" onclick="loadMk('${p.id}')">내 편성기로 가져오기</button>${p.mine?`<button class="vb d" onclick="delMk('${p.id}')">삭제</button>`:''}</div></div>
    ${p.note?`<div class="m" style="margin-bottom:6px">${esc(p.note)}</div>`:''}
    ${p.units.map((u,i)=>`<div class="mu"><span class="badge s">${i+1}</span>${[0,1,2].map(k=>{const h=u.h[k];if(!h)return'<div class="mh"></div>';return `<div class="mh ${oh.has(h)?'':'no'}">${imj((HM[h]||{}).img,'th')}<div><b>${esc(h)}</b>${oh.has(h)?'':' <span class="x1">미보유</span>'}${(u.t[k]||[]).map(t=>`<div class="mt ${os.has(t)?'':'no'}">${os.has(t)?'✓':'✗'} ${esc(t)}</div>`).join('')}</div></div>`}).join('')}</div>`).join('')}</div>`};
  return `<div class="m" style="margin-bottom:8px">다른 동맹원들이 올린 5부대 편성입니다(모든 서버 공유). 👍 추천을 많이 받은 유저는 <b>지략가 TOP3</b>가 됩니다. <b>✓</b>는 내가 보유한 것, <b>✗</b>는 없는 것입니다.</div>
  <div class="bar"><button class="sbtn ${MS==='like'?'on':''}" onclick="MS='like';draw()">추천순</button><button class="sbtn ${MS==='new'?'on':''}" onclick="MS='new';draw()">최신순</button><button class="sbtn ${MF?'on':''}" onclick="MF=!MF;draw()">내가 전부 보유한 것만</button><span class="sp"></span><button class="sbtn" onclick="setTab('build')">＋ 내 편성 올리기</button></div>
  ${l.map(card).join('')||'<div class="m">아직 올라온 편성이 없습니다. 편성기에서 첫 편성을 올려 보세요!</div>'}`}
function strat(){const r=D.mk.rank;const md=['🥇','🥈','🥉'];
  return `<div class="m" style="margin-bottom:8px">올린 편성이 받은 <b>👍 추천 합계</b> 순위입니다. <b>TOP3는 🧠 지략가</b>가 되어 마켓에 배지가 붙고, 디스코드 서버에서 <b>지략가 역할</b>을 받습니다(순위가 바뀌면 자동 회수·부여).</div>
  <table><tr><th>#</th><th>유저</th><th>받은 👍</th><th>올린 편성</th><th></th></tr>${r.map((x,i)=>`<tr ${x.me?'style="background:var(--card2)"':''}><td>${md[i]||i+1}</td><td><b>${esc(x.by)}</b>${x.me?' <span class="m">(나)</span>':''}</td><td><b>${x.likes}</b></td><td>${x.posts}</td><td>${x.strat?'<span class="strat">🧠 지략가</span>':''}</td></tr>`).join('')||'<tr><td colspan="5" class="m">아직 순위가 없습니다</td></tr>'}</table>`}
function draw(){
  if(ERR){$('#app').innerHTML=`<div class="err"><h1>추천덱</h1><p>${esc(ERR)}</p></div>`;return}
  const q=s=>`?t=${encodeURIComponent(D.t)}&s=${s}`;
  const tb=(k,l,c)=>`<button class="tab ${c||''} ${TAB===k?'on':''}" onclick="setTab('${k}')">${l}</button>`;
  const NAME={rank:'티어덱',build:'편성기',market:'공유마켓',strat:'지략가',own:'내 장수·전법',equip:'장비',...FR};
  const main=tb('rank','티어덱')+tb('build','편성기')+tb('market','🛒 공유마켓')+tb('strat','🧠 지략가')+tb('own','내 장수·전법')+tb('dogam','장수 도감')+tb('skdex','전법 도감');
  const late=tb('equip','장비','late')+tb('chart','상성 차트','late')+tb('mx','상성표','late')+tb('fill','부족한 것 메우기','late')+tb('guide','상황 가이드','late');
  $('#app').innerHTML=`<h1>🃏 ${esc(D.who)} · 티어덱 · 편성 공유마켓 · 시즌${D.season}</h1><div class="sub">티어덱을 보고, 내 보유 장수·전법으로 5부대를 짜서 공유하세요 · 한 번 쓴 장수·전법은 다른 부대에서 쓸 수 없음</div>
  <div class="bar"><button class="tab mt" onclick="toggleMenu()">☰ 메뉴 ${MENU?'접기 ▴':'펼치기 ▾'}</button>${MENU?'':`<span class="m">현재: <b>${esc(NAME[TAB]||'')}</b></span>`}<span class="sp"></span>
  <a class="sbtn ${D.season===1?'on':''}" href="${q(1)}">시즌1</a><a class="sbtn ${D.season===2?'on':''}" href="${q(2)}">시즌2</a></div>
  ${MENU?`<div class="bar">${main}<span class="vs"></span><button class="tab lt" onclick="toggleLate()">참고 ${LATE?'접기 ▴':'펼치기 ▾'}</button>${LATE?late:''}</div>`:''}
  ${TAB==='build'?build():TAB==='market'?market():TAB==='strat'?strat():TAB==='rank'?rank():TAB==='own'?own():TAB==='equip'?equip():FR[TAB]?frame(TAB):rank()}`}
draw();
</script></body></html>'''


def render(data, err=None):
    """data(dict) 또는 err(문장)로 완성된 HTML 문자열을 만든다."""
    j = json.dumps(data or {"units": [], "rank": {"decks": [], "tacs": [], "syn": []}, "t": "", "season": 1, "who": "", "stone": ""}, ensure_ascii=False).replace("</", "<\\/")
    return TPL.replace("__DATA__", j).replace("__ERR__", json.dumps(err, ensure_ascii=False))
