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
.hs{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}@media(max-width:760px){.hs{grid-template-columns:1fr}}
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
.ocs{display:flex;flex-wrap:wrap;gap:6px}.oc{display:flex;align-items:center;gap:6px;text-align:left}.oc .nm{font-weight:600}.oc .m{display:block;font-size:11px;color:var(--mute)}.th{width:33px;height:42px;border-radius:5px;object-fit:cover}.ts{width:30px;height:30px;border-radius:5px;object-fit:cover}.sts{display:flex;gap:3px;margin-top:3px}.sg{width:11px;height:14px;border-radius:3px 3px 50% 50%/3px 3px 60% 60%;background:#3a2a2a;border:1px solid #5a3a3a;cursor:pointer}.sg.on{background:radial-gradient(circle at 35% 30%,#ff9a8a,#d42a2a 60%,#7a1010);border-color:#ff6a5a}.gh{flex:0 0 100%;font-size:12px;color:var(--gold);margin:6px 0 0;border-bottom:1px solid var(--line)}.oc{background:var(--card2);border:1px solid var(--line);color:var(--mute);border-radius:8px;padding:4px 9px;cursor:pointer;font-size:13px;opacity:.7}.oc.on{opacity:1;color:var(--txt);border-color:#2e9f5f;background:rgba(46,159,95,.15);font-weight:600}.sq{background:var(--card);border:1px solid var(--line);color:var(--txt);border-radius:8px;padding:7px 10px;font-size:13px;min-width:160px}\n.rw{cursor:pointer}.rw:hover{background:var(--card2)}.dt td{background:var(--card2)}.dr{display:flex;gap:10px;margin:6px 0}.dr>b{width:70px;flex:none;color:var(--gold);font-size:12px}.dh{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;flex:1}.dc{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:6px;line-height:2}.ownt{color:#2e9f5f}.tk{display:inline-flex;align-items:center;gap:3px;border:1px solid var(--line);border-radius:6px;padding:0 5px;font-size:12px;background:var(--card2)}.tk.own{border-color:#2e9f5f;color:#2e9f5f;font-weight:700}.tk img{width:16px;height:23px;object-fit:cover;border-radius:2px}
@media(max-width:760px){.dh{grid-template-columns:1fr}}
.err{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:20px;text-align:center}
</style></head><body><div class="wrap" id="app"></div>
<script>
const D=__DATA__;const ERR=__ERR__;const $=s=>document.querySelector(s);
const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const NAT={'촉':'#50aa6e','위':'#648cdc','오':'#dc645a','군':'#aaaaaa'};const KC={'한방딜':'burst','도트딜':'dot','일반딜':'norm','회복':'heal','지원':'sup'};
let TAB='mine';try{TAB=sessionStorage.getItem('rtab')||'mine'}catch(e){}
async function post(o){try{const r=await fetch('/api/reco',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({t:D.t,season:D.season,...o})});const j=await r.json();if(!j.ok)alert(j.error||'실패');else location.reload()}catch(e){alert('네트워크 오류')}}
function vote(i,scope,kind){const u=D.units[i];if(scope==='h')post({op:'vote',scope:'h',key:u.hk,kind,heroes:u.heroes.filter(h=>h.n).map(h=>h.n)});else post({op:'vote',scope:'t',key:u.tk,kind,pairs:u.heroes.flatMap(h=>h.n?h.tac.map(t=>[h.n,t.n]):[]),tset:u.tset})}
function tg(k){const e=document.getElementById('dt'+k),a=document.getElementById('ar'+k);const o=e.style.display==='none';e.style.display=o?'':'none';a.textContent=o?'▾':'▸'}
let TRS=null,OWNS=null,OQ='',OF='all',DIRTY=false;
async function setTr(name,k){TRS=TRS||Object.assign({},D.own.tr||{});const n=(TRS[name]||0)===k?k-1:k;
  try{const r=await fetch('/api/reco',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({t:D.t,season:D.season,op:'tr',name,n})});const j=await r.json();if(!j.ok){alert(j.error||'실패');return}}catch(e){alert('네트워크 오류');return}
  TRS[name]=n;DIRTY=true;const y=scrollY;draw();scrollTo(0,y)}
async function toggleOwn(kind,name){OWNS=OWNS||{h:new Set(D.own.h),s:new Set(D.own.s)};const on=!OWNS[kind].has(name);
  try{const r=await fetch('/api/reco',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({t:D.t,season:D.season,op:'own',kind,name,on})});const j=await r.json();if(!j.ok){alert(j.error||'실패');return}}catch(e){alert('네트워크 오류');return}
  on?OWNS[kind].add(name):OWNS[kind].delete(name);DIRTY=true;const y=scrollY;draw();scrollTo(0,y)}
function setQ(v){OQ=v;const y=scrollY;draw();const e=document.getElementById('oq');if(e){e.focus();e.setSelectionRange(v.length,v.length)}scrollTo(0,y)}
function setF(f){OF=f;draw()}
function pin(name,on){post({op:'pin',deck:name,on})}
function setTab(t){if(DIRTY&&t==='mine'){DIRTY=false;try{sessionStorage.setItem('rtab','mine')}catch(e){}location.reload();return}TAB=t;try{sessionStorage.setItem('rtab',t)}catch(e){}draw()}
function vg(label,v,i,scope){return `<span class="vg">${label}<button class="vb u ${v.m==='u'?'on':''}" onclick="vote(${i},'${scope}','u')">👍 ${v.u}</button><button class="vb d ${v.m==='d'?'on':''}" onclick="vote(${i},'${scope}','d')">👎 ${v.d}</button></span>`}
function hero(h){
  if(!h.n)return `<div class="h"><div class="empty">${esc(h.src||'보유 장수 부족')}</div></div>`;
  const gem=h.tr>0?`<div class="gems" style="--st:url(${D.stone})">${[1,2,3,4,5].map(k=>`<i class="${h.tr>=k?'on':''}"></i>`).join('')}</div>`:'';
  const L=['무력','지력','통솔','선공'];
  const st=h.fin?L.map((l,k)=>`<span>${l} <b>${h.fin[k]}</b>${h.add[k]?`<em> (+${h.add[k]})</em>`:''}</span>`).join(''):'';
  const tac=h.tac.length?h.tac.map(t=>`<div class="tac">${t.img?`<img src="data:image/jpeg;base64,${t.img}">`:'<img>'}<div><div class="t">${esc(t.n)}</div><div class="k"><span class="${KC[t.kind]||''}">${esc(t.kind)}</span> <span class="m">${esc(t.type)}·${esc(t.trait)} · ${esc(t.why)}</span></div></div></div>`).join(''):`<div class="empty">전법 부족 (보유 전법 없음)</div>`;
  return `<div class="h"><div class="pic">${h.img?`<img class="p" src="data:image/jpeg;base64,${h.img}">`:''}<div class="lb">${gem}50 ${esc(h.n)}</div></div><div class="hi">
  <div class="hn">${esc(h.n)}${h.nation?`<span class="nat" style="background:${NAT[h.nation]||'#999'}">${esc(h.nation)}</span>`:''}${h.rank?`<span class="rk" title="행동 순서">${h.rank}</span>`:''}</div>
  <div class="m">${esc([h.pos,h.troop].filter(Boolean).join(' · '))}</div><div class="sk">${esc(h.sk)}${h.sktype?` (${esc(h.sktype)})`:''}</div>
  <div class="stats">${st}</div>${h.src?`<div class="src">← ${esc(h.src)}</div>`:''}${tac}</div></div>`}
function unit(u,i){
  const fl=u.flow,s=fl.start;
  const lines=[`<div><b>전투 시작</b><span>지휘 ${esc(s['지휘'].join(' · ')||'-')} → 패시브 ${esc(s['패시브'].join(' · ')||'-')}</span></div>`].concat(fl.turn.map((t,k)=>`<div><b>턴 ${k+1}번째</b><span>${esc(t[0])}${t[1]!=null?`(${t[1]})`:''} 액티브 ${esc(t[2].join(' · ')||'-')} → 평타${t[3].length?` → 추격 ${esc(t[3].join(' · '))}`:''}</span></div>`));
  if(u.status.g.length)lines.push(`<div><b>이상상태</b><span class="m">부여 ${esc(u.status.g.join(', '))}${u.status.miss.length?` · 부족 ${esc(u.status.miss.join(', '))}`:''}</span></div>`);
  return `<div class="unit"><div class="uh"><span class="badge">${u.i}부대</span>${u.name?`<span class="m">${esc(u.tier)}</span><span class="nm">${esc(u.name)}</span><span class="tr">${esc(u.form)}</span>`:'<span class="nm">자유 편성</span>'}
  ${u.name?`<div class="votes">${vg('장수 조합',u.hv,i,'h')}${u.tset.length?vg('전법 조합',u.tv,i,'t'):''}</div>`:''}</div>
  <div class="hs">${u.heroes.map(hero).join('')}</div><div class="flow">${lines.join('')}</div></div>`}
function mine(){return D.units.map(unit).join('')+`<div class="m">티어덱을 완성하지 못한 부대는 같은 역할의 보유 장수로 채웁니다. 장수 👎는 그 덱을 빼고, 전법 👎는 전법만 다시 고릅니다. 모든 서버의 평가가 합산됩니다.</div>`}
function rank(){const R=D.rank;
  const own=new Set(D.ownsk||[]),nk=s=>s.replace(/[\s·\-]/g,'');
  const rhero=h=>`<div class="h"><div class="pic">${h.img?`<img class="p" src="data:image/jpeg;base64,${h.img}">`:''}<div class="lb">50 ${esc(h.n)}</div></div><div class="hi">
    <div class="hn">${esc(h.n)}${h.nation?`<span class="nat" style="background:${NAT[h.nation]||'#999'}">${esc(h.nation)}</span>`:''}</div>
    <div class="m">${esc([h.pos,h.troop].filter(Boolean).join(' · '))}${h.alts.length?` · 대체 ${esc(h.alts.join('/'))}`:''}</div><div class="sk">${esc(h.sk)}${h.sktype?` (${esc(h.sktype)})`:''}</div>
    ${h.tac.map(t=>`<div class="tac">${t.img?`<img src="data:image/jpeg;base64,${t.img}">`:'<img>'}<div><div class="t ${own.has(nk(t.n))?'ownt':''}">${esc(t.n)}${own.has(nk(t.n))?' ✓':''}</div><div class="k"><span class="${KC[t.kind]||''}">${esc(t.kind)}</span> <span class="m">${esc([t.type,t.trait].filter(Boolean).join('·'))}</span></div>${t.alt.length?`<div class="m">대체: ${esc(t.alt.join(' / '))}</div>`:''}</div></div>`).join('')}
    ${h.tac.length<2?'<div class="empty">추천 전법 정보 없음</div>':''}</div></div>`;
  const detail=d=>{const cb=R.tacs.filter(t=>t.deck===d.name);const rated=cb.length?`<div class="flow"><div><b>평가된 전법 조합</b><span>${cb.map(t=>`${t.tset.map(x=>`<span class="chip">${esc(x)}</span>`).join('')} <span class="m">👍 ${t.u} 👎 ${t.d}</span>`).join('<br>')}</span></div></div>`:'';
    return `<div class="m" style="margin-bottom:6px">${esc(d.form)}${d.note?' · '+esc(d.note):''} · <span style="color:#2e9f5f;font-weight:700">✓</span> = 내가 보유한 전법 · 장수마다 추천 전법 2개</div><div class="hs">${d.cards.map(rhero).join('')}</div>${rated}`};
  const dk=`<table><tr><th>#</th><th>티어</th><th>덱 (눌러서 전법 보기)</th><th>장수</th><th>👍</th><th>👎</th><th>점수</th><th>내 추천</th></tr>${R.decks.map((d,k)=>`<tr class="rw" onclick="tg(${k})"><td>${k+1}</td><td>${esc(d.tier)}</td><td><b>${esc(d.name)}</b> <span class="m" id="ar${k}">▸</span></td><td class="m">${esc(d.heroes.join(' · '))}</td><td>${d.u}</td><td>${d.d}</td><td><b>${d.net>0?'+':''}${d.net}</b></td><td><button class="pin ${d.pin?'on':''}" onclick="event.stopPropagation();pin('${esc(d.name).replace(/'/g,"\\'")}',${!d.pin})">${d.pin?'📌 적용 중':'적용'}</button></td></tr><tr class="dt" id="dt${k}" style="display:none"><td></td><td colspan="7">${detail(d)}</td></tr>`).join('')}</table>`;
  const tc=R.tacs.length?`<table><tr><th>덱</th><th>전법 조합</th><th>👍</th><th>👎</th></tr>${R.tacs.map(t=>`<tr><td>${esc(t.deck)}</td><td>${t.tset.map(x=>`<span class="chip">${esc(x)}</span>`).join('')}</td><td>${t.u}</td><td>${t.d}</td></tr>`).join('')}</table>`:'<div class="m">아직 평가된 전법 조합이 없습니다.</div>';
  const KN={ht:'장수↔전법',tt:'전법↔전법',hh:'장수↔장수'};
  const sy=R.syn.length?`<table><tr><th>종류</th><th>궁합</th><th>점수</th><th>메모</th></tr>${R.syn.map(s=>`<tr><td>${KN[s.kind]||s.kind}</td><td><b>${esc(s.a)}</b> ↔ <b>${esc(s.b)}</b></td><td>${s.score>0?'+':''}${s.score}</td><td class="m">${esc(s.note)}${s.src==='vote'?' (평가로 학습)':''}</td></tr>`).join('')}</table>`:'<div class="m">저장된 시너지가 없습니다.</div>';
  return `<h2>덱 랭킹 — 장수 조합 평가 점수(모든 서버 합산)</h2><div class="m" style="margin-bottom:6px">점수가 높을수록 추천 순위가 올라갑니다. <b>적용</b>을 누르면 그 덱을 내 추천에서 항상 먼저 편성합니다(보유 장수가 있을 때).</div>${dk}
  <h2>전법 조합 랭킹</h2>${tc}<h2>시너지 점수 TOP — 장수·전법 궁합(추천 점수에 더해집니다)</h2>${sy}`}
function own(){TRS=TRS||Object.assign({},D.own.tr||{});OWNS=OWNS||{h:new Set(D.own.h),s:new Set(D.own.s)};const q=OQ.replace(/\s/g,'');
  const keep=(n,o,ss)=>(ss||1)<=D.season&&(!q||n.replace(/\s/g,'').includes(q))&&(OF==='all'||(OF==='own')===o);
  const GR={'전설':0,'영웅':1},ko=(a,b)=>a.n.localeCompare(b.n,'ko');
  const hs=D.all.h.filter(h=>keep(h.n,OWNS.h.has(h.n),h.season)).sort(ko);
  const ss=D.all.s.filter(s=>keep(s.n,OWNS.s.has(s.n),s.season)).sort((a,b)=>((GR[a.grade]??9)-(GR[b.grade]??9))||ko(a,b));
  const im=(b,c)=>b?`<img class="${c}" src="data:image/jpeg;base64,${b}">`:'';
  const stones=h=>{const n=TRS[h.n]||0;return `<span class="sts" title="초월 단계">${[1,2,3,4,5].map(k=>`<i class="sg ${k<=n?'on':''}" onclick="event.stopPropagation();setTr('${esc(h.n)}',${k})"></i>`).join('')}</span>`};
  const hc=hs.map(h=>`<div class="oc hc ${OWNS.h.has(h.n)?'on':''}" onclick="toggleOwn('h','${esc(h.n)}')">${im(h.img,'th')}<span><span class="nm">${esc(h.n)}</span><span class="m">${esc(h.nation)} ${esc(h.troop)}${(TRS[h.n]||0)?' · 초월 '+TRS[h.n]:''}</span>${OWNS.h.has(h.n)?stones(h):''}</span></div>`).join('');
  let sc='',lg=null;for(const s of ss){if(s.grade!==lg){lg=s.grade;sc+=`<div class="gh">${esc(lg||'기타')}</div>`}
   sc+=`<button class="oc sc ${OWNS.s.has(s.n)?'on':''}" onclick="toggleOwn('s','${esc(s.n).replace(/'/g,"\\'")}')">${im(s.img,'ts')}<span class="nm">${esc(s.n)}</span><span class="m">${esc(s.type)}·${esc(s.trait)}</span></button>`}
  return `<div class="m" style="margin-bottom:8px">눌러서 보유 여부를 바꿉니다(자동 저장). 보유 장수의 <b>붉은 돌</b>을 눌러 초월 단계(0~5)를 정합니다. 스탯 조율은 <b>/도감</b> 페이지에서 합니다. 바꾼 뒤 <b>내 추천덱</b> 탭으로 돌아가면 다시 편성합니다.</div>
  <div class="bar"><input id="oq" class="sq" placeholder="이름 검색" value="${esc(OQ)}" oninput="setQ(this.value)"><button class="sbtn ${OF==='all'?'on':''}" onclick="setF('all')">전체</button><button class="sbtn ${OF==='own'?'on':''}" onclick="setF('own')">보유</button><button class="sbtn ${OF==='no'?'on':''}" onclick="setF('no')">미보유</button></div>
  <h2>장수 (보유 ${OWNS.h.size})</h2><div class="ocs">${hc||'<span class="m">없음</span>'}</div><h2>전법 (보유 ${OWNS.s.size})</h2><div class="ocs">${sc||'<span class="m">없음</span>'}</div>`}
function draw(){
  if(ERR){$('#app').innerHTML=`<div class="err"><h1>추천덱</h1><p>${esc(ERR)}</p></div>`;return}
  const q=s=>`?t=${encodeURIComponent(D.t)}&s=${s}`;
  $('#app').innerHTML=`<h1>🃏 ${esc(D.who)} 추천덱 · 시즌${D.season}</h1><div class="sub">장수 카드·전법·초월 반영 선공 순서 · 한 번 쓴 장수·전법은 다시 쓰지 않음</div>
  <div class="bar"><button class="tab ${TAB==='mine'?'on':''}" onclick="setTab('mine')">내 추천덱</button><button class="tab ${TAB==='rank'?'on':''}" onclick="setTab('rank')">추천덱 랭킹</button><button class="tab ${TAB==='own'?'on':''}" onclick="setTab('own')">내 장수·전법</button><span class="sp"></span>
  <a class="sbtn ${D.season===1?'on':''}" href="${q(1)}">시즌1</a><a class="sbtn ${D.season===2?'on':''}" href="${q(2)}">시즌2</a></div>${TAB==='mine'?mine():TAB==='rank'?rank():own()}`}
draw();
</script></body></html>'''


def render(data, err=None):
    """data(dict) 또는 err(문장)로 완성된 HTML 문자열을 만든다."""
    j = json.dumps(data or {"units": [], "rank": {"decks": [], "tacs": [], "syn": []}, "t": "", "season": 1, "who": "", "stone": ""}, ensure_ascii=False).replace("</", "<\\/")
    return TPL.replace("__DATA__", j).replace("__ERR__", json.dumps(err, ensure_ascii=False))
