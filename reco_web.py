"""추천덱 HTML 페이지(/reco): 내 추천덱(장수 카드·전법·돌·턴 순서·평가 버튼) + 추천덱 랭킹 + 시너지 점수."""
import html, json

TPL = r'''<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>덱 검색기</title>
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
.hs{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}@media(max-width:760px){.hcols{grid-template-columns:repeat(3,minmax(0,1fr))!important;gap:6px}.hblk,.hcol,.hst{display:contents!important}.sumw{display:block!important}.stw{display:none!important}.stw.op{display:block!important}.r1{grid-row:1}.r2{grid-row:2}.r3{grid-row:3}.r4{grid-row:4}.r5{grid-row:5;grid-column:1/-1}.c0{grid-column:1}.c1{grid-column:2}.c2{grid-column:3}.sumw{display:block}.sbs{display:block!important}.sbs0{display:block!important}.stw{display:none}.stw.op{display:block}.sbx{display:block!important}.sbn{display:inline!important;font-size:14px}.hs{gap:5px;grid-template-columns:repeat(3,minmax(0,1fr))}.unit{padding:8px}.h,.hi,.tac>div{min-width:0}.h{flex-direction:column;gap:5px;padding:5px}.pic{width:100%;height:auto;aspect-ratio:3/4}.pic .lb{font-size:11px;padding:12px 0 2px}
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
#tip{position:fixed;z-index:99;display:none;max-width:320px;pointer-events:none;background:linear-gradient(160deg,#16322d,#0e1d1a);color:#e8efe9;border:1px solid var(--jade);border-left:4px solid var(--gold);border-radius:10px;padding:9px 12px;font-size:12.5px;line-height:1.5;box-shadow:0 8px 24px rgba(0,0,0,.35)}#tip b{font-size:14px;color:var(--gold)}#tip .tm{color:#9fc3bb;font-size:11px}#tip .tk{color:var(--jade);font-size:11px;margin-top:2px}#tip .td{margin-top:4px;color:#dfe9e4}
@media(hover:hover){.oc.sc,.oc.hc{transition:transform .12s,box-shadow .12s,border-color .12s}.oc.sc:hover,.oc.hc:hover{transform:translateY(-2px);border-color:var(--jade);box-shadow:0 6px 16px rgba(0,0,0,.22)}.tac[data-tip]:hover{background:var(--card2);border-radius:8px}.bs:hover{border-color:var(--jade)}}
.gh.rec{color:var(--jade);font-weight:700;border-bottom-color:var(--jade)}.kw.rc{color:var(--jade)}.pdet{border-left:4px solid var(--gold);background:linear-gradient(160deg,var(--card),var(--card2))}.pdet .nm b{font-size:15px}
.bs img{pointer-events:none;-webkit-user-drag:none}.dgh{position:fixed;z-index:120;pointer-events:none;background:var(--jade);color:#06231c;font-weight:700;border-radius:8px;padding:5px 10px;font-size:13px;box-shadow:0 6px 18px rgba(0,0,0,.35);opacity:.95}.bs.dsrc{opacity:.35}.bs.dz{border-color:var(--gold);box-shadow:0 0 0 2px var(--gold);background:rgba(240,199,94,.15)}body.dragging{user-select:none;-webkit-user-select:none;cursor:grabbing}.bs[data-tu],.bs[data-du]{-webkit-touch-callout:none;-webkit-user-select:none;user-select:none}.dh{font-size:11px}
.shm{position:fixed;inset:0;z-index:130;background:rgba(0,0,0,.7);overflow:auto;padding:14px}.shb{max-width:760px;margin:0 auto;background:var(--card);border-radius:12px;padding:12px}.shb img{width:100%;border-radius:8px;margin-top:8px;display:block}
.kw{display:block;font-size:10px;color:var(--gold)}.picker{display:flex;flex-direction:column;overflow:hidden!important;max-height:62vh}.picker>.bar{flex:none;margin:0 0 6px}#pdet{flex:none}#pl{flex:1;min-height:0;overflow-y:auto;-webkit-overflow-scrolling:touch;padding-right:2px}.pdet.empty{text-align:center;padding:10px}.pdet{background:var(--card);border:1px solid var(--jade);border-radius:10px;padding:8px;margin:6px 0}.pdet .dsc{font-size:12.5px;line-height:1.5;margin:4px 0}.oc.pd{border-color:var(--gold);box-shadow:0 0 0 2px var(--gold)}.picker{position:sticky;top:0;z-index:5;background:var(--bg);border:2px solid var(--jade);border-radius:12px;padding:8px;margin-bottom:12px;}
.hchips{display:flex;flex-wrap:wrap;gap:4px;margin:4px 0}.hchip{position:relative;display:inline-block}.hchip .th{width:34px;height:42px;border-radius:6px;object-fit:cover;display:block}.hchip .ct{position:absolute;right:0;bottom:0;background:#0d7a66;color:#fff;font-style:normal;font-size:10px;font-weight:700;padding:0 4px;border-radius:6px 0 6px 0}.ans{border-top:1px dashed var(--line);margin-top:10px;padding-top:8px}.ans.ad{background:rgba(13,122,102,.07);border-radius:8px;padding:8px}.ansb{border:2px solid #0d7a66}.fmr{display:flex;align-items:center;gap:8px;margin:6px 0}.sbs{width:100%;box-sizing:border-box}.uh .dh{flex-basis:100%}.bs.ro{cursor:default}.bs.no{opacity:.55;border-style:dashed}.sbs.ro{cursor:default;pointer-events:none}.sbs.ro .sm{color:var(--mute)}.sbs.ro .sn em{font-style:normal;color:#0d7a66;font-size:10px;margin-left:1px}.mkv{margin-top:8px;background:var(--card2)}.bwrap{max-width:820px;margin:0 auto}.hcols{grid-template-columns:repeat(3,minmax(0,1fr))!important;gap:6px!important}.hblk,.hcol,.hst{display:contents!important}.r1{grid-row:1}.r2{grid-row:2}.r3{grid-row:3}.r4{grid-row:4}.r5{grid-row:5;grid-column:1/-1}.c0{grid-column:1}.c1{grid-column:2}.c2{grid-column:3}.sumw{display:block!important}.sbs,.sbs0{display:block!important}.stw{display:none!important}.stw.op{display:block!important}.sbx{display:block!important}.sbn{display:inline!important;font-size:14px}
.rn{margin:4px 0}.bimg{max-width:100%;max-height:360px;border-radius:8px;border:1px solid var(--line);display:block;margin:6px 0}.bprev{display:flex;align-items:flex-end;gap:8px;margin:6px 0}.bprev .bimg{max-height:120px;margin:0}.bdb{white-space:pre-wrap;word-break:break-word;margin:6px 0;line-height:1.55}.bdb.cl{display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden}.bcs{border-top:1px dashed var(--line);margin-top:8px;padding-top:6px}.bc{padding:4px 0;border-bottom:1px solid var(--line);font-size:14px}.bc:last-of-type{border:0}.vb.d.on{background:#c0443a;color:#fff}.ttl{display:flex;align-items:flex-start;justify-content:space-between;gap:8px}.ttl h1{margin:0 0 6px}.dbtn{flex:none;white-space:nowrap}.hcols{display:grid;grid-template-columns:repeat(auto-fit,minmax(330px,1fr));gap:8px}.hblk{display:flex;gap:8px;align-items:flex-start;min-width:0}.hcol{flex:0 0 130px;display:flex;flex-direction:column;gap:6px;min-width:0}.hst{flex:1;min-width:0}.sumw{display:none}.sbs,.sbs0{display:none}.sbs{border:1px solid var(--line);background:var(--card);border-radius:10px;padding:4px 6px;text-align:left;cursor:pointer;font:inherit}.sbs.op{border-color:#0d7a66;background:var(--card2)}.sbs .sn{display:grid;grid-template-columns:1fr 1fr;gap:0 6px;font-weight:700;font-size:12px;color:#9a6b00}.sbs .sn i{font-style:normal;color:var(--mute);font-weight:600;margin-right:2px;font-size:11px}.sbs .sm{display:block;font-size:11px;color:#0d7a66;margin-top:2px;font-weight:700}.sbn{display:none}.sbx{border:1px solid var(--line);border-radius:10px;padding:6px 8px 8px;background:var(--card)}.sbh{display:flex;align-items:center;gap:6px;flex-wrap:wrap;font-size:12px;margin-bottom:4px}.sbt{display:flex;align-items:center;gap:6px;margin-bottom:5px}.sbt .sbtn{padding:4px 10px}.avl{flex:1;text-align:center;background:var(--card2);border:1px solid var(--line);border-radius:8px;padding:4px 6px;font-size:13px;color:var(--mute)}.avl b{color:#0d7a66;font-size:15px;margin-left:4px}.avl b.full{color:#b3541e}.avl i{font-style:normal;font-size:11px;margin-left:2px}.sbr{position:relative;display:flex;align-items:center;gap:6px;margin-top:4px;padding:3px 4px 3px 6px;background:var(--card2);border:1px solid var(--line);border-radius:8px}.sic{width:24px;height:24px;border-radius:6px;background:#0d7a66;color:#fff;display:flex;align-items:center;justify-content:center;font-size:13px;flex:none}.sl{font-weight:700;font-size:14px;min-width:34px}.sv{flex:1;font-weight:800;font-size:15px;color:#9a6b00;white-space:nowrap}.sv em{font-style:normal;font-size:12px;color:#0d7a66;margin-left:3px;font-weight:700}.rcm{background:#e8873a;color:#fff;font-size:10px;font-weight:700;padding:1px 6px;border-radius:9px;margin-left:-2px}.stb{border:1px solid var(--line);background:#fff;color:#0d7a66;border-radius:6px;width:32px;height:30px;font-size:17px;line-height:1;cursor:pointer;padding:0;font-weight:700}.stb:hover:not(:disabled){background:#0d7a66;color:#fff}.stb:disabled{opacity:.35;cursor:default}
.mk .nm{font-size:15px}.strat{background:#f0c75e;color:#2b2000;border-radius:10px;padding:1px 8px;font-size:11px;font-weight:700}.mu{display:grid;grid-template-columns:22px repeat(3,minmax(0,1fr));gap:6px;align-items:start;margin-top:5px}.badge.s{padding:0 6px;text-align:center;height:20px}.mh{display:flex;gap:5px;font-size:12px;min-width:0}.mh .th{flex:none}.mh.no .th{opacity:.45}.mt{font-size:11px;color:var(--jade);line-height:1.3}.mt.no{color:var(--mute);text-decoration:line-through}.x1{color:#c0504d;font-size:10px}.vb:disabled{opacity:.5;cursor:default}
@media(max-width:760px){.mu{grid-template-columns:18px repeat(3,minmax(0,1fr));gap:3px}.mh{flex-direction:column;gap:2px}.bs{padding:3px;font-size:11px;flex-direction:column;align-items:flex-start;gap:2px}.bs .x{position:absolute;top:0;right:0}.picker{max-height:66vh}}
.fr{width:100%;height:78vh;border:1px solid var(--line);border-radius:12px;background:#fff}.vs{width:2px;height:24px;background:var(--line);margin:0 2px}.tab.late{background:#f7f2e4}.tab.mt,.tab.lt{font-weight:700;background:var(--card2)}.tab.lt{background:#efe6c8}.mo{font-weight:700}
.ocs{display:flex;flex-wrap:wrap;gap:6px}.oc{display:flex;align-items:center;gap:6px;text-align:left}.oc .nm{font-weight:600}.oc .m{display:block;font-size:11px;color:var(--mute)}.th{width:33px;height:42px;border-radius:5px;object-fit:cover}.ts{width:30px;height:30px;border-radius:5px;object-fit:cover}.sts{display:flex;gap:3px;margin-top:3px}.sg{width:14px;height:18px;background:var(--st) center/contain no-repeat;filter:grayscale(1) brightness(.45);cursor:pointer;display:block}.sg.on{filter:none}.gh{flex:0 0 100%;font-size:12px;color:var(--gold);margin:6px 0 0;border-bottom:1px solid var(--line)}.oc{background:var(--card2);border:1px solid var(--line);color:var(--mute);border-radius:8px;padding:4px 9px;cursor:pointer;font-size:13px;opacity:.7}.oc.on{opacity:1;color:var(--txt);border-color:#2e9f5f;background:rgba(46,159,95,.15);font-weight:600}.sq{background:var(--card);border:1px solid var(--line);color:var(--txt);border-radius:8px;padding:7px 10px;font-size:13px;min-width:160px}\n.rw{cursor:pointer}.rw:hover{background:var(--card2)}.dt td{background:var(--card2)}.dr{display:flex;gap:10px;margin:6px 0}.dr>b{width:70px;flex:none;color:var(--gold);font-size:12px}.dh{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;flex:1}.dc{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:6px;line-height:2}.ownt{color:#2e9f5f}.tk{display:inline-flex;align-items:center;gap:3px;border:1px solid var(--line);border-radius:6px;padding:0 5px;font-size:12px;background:var(--card2)}.tk.own{border-color:#2e9f5f;color:#2e9f5f;font-weight:700}.tk img{width:16px;height:23px;object-fit:cover;border-radius:2px}
@media(max-width:760px){.dh{grid-template-columns:1fr}}
.err{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:20px;text-align:center}
.eqt{display:flex;flex-direction:column;gap:6px}.eqr{display:flex;gap:6px;align-items:stretch}.eql{flex:0 0 52px;color:#fff;border-radius:8px;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:4px 0}.eql b{font-size:22px;line-height:1}.eql span{font-size:10px;opacity:.9}.eqc{flex:1;display:flex;flex-wrap:wrap;gap:6px}.eqi{flex:1 1 150px;max-width:100%;border:1px solid;border-radius:8px;padding:6px 8px;background:var(--card,#fff);display:flex;flex-direction:column;gap:2px;font-size:12px}.eqi b{font-size:14px}.eqi small{color:#8a9a94;font-size:10.5px}
.tg{display:inline-flex;align-items:center;gap:1px;font-size:10px;color:var(--gold,#e0b65a);margin-left:3px}.tg:before{content:"";width:9px;height:12px;background:var(--st) center/contain no-repeat}
</style></head><body><div class="wrap" id="app"></div>
<script>
const D=__DATA__;const ERR=__ERR__;const $=s=>document.querySelector(s);
const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const NAT={'촉':'#50aa6e','위':'#648cdc','오':'#dc645a','군':'#aaaaaa'};const KC={'한방딜':'burst','도트딜':'dot','일반딜':'norm','회복':'heal','지원':'sup'};
let MODE='pvp';try{MODE=sessionStorage.getItem('rmode')||'pvp'}catch(e){}
const FR={dogam:'장수 도감',skdex:'전법 도감',chart:'상성 차트',mx:'상성표',fill:'부족한 것 메우기',guide:'상황 가이드'};
let EQ='';let TAB='rank';try{TAB=sessionStorage.getItem('rtab')||'rank';if(TAB==='mine'||TAB==='own')TAB='rank'}catch(e){}
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
function showDon(){const m=document.createElement('div');m.className='shm';m.innerHTML=`<div class="shb" style="max-width:340px;text-align:center"><div class="bar"><b>☕ 제작자 커피 후원</b><span class="sp"></span><button class="sbtn" onclick="this.closest('.shm').remove()">닫기</button></div>`+donate()+`</div>`;m.onclick=e=>{if(e.target===m)m.remove()};document.body.appendChild(m)}
function openDon(){const l=(D.don||{}).link;if(!l)return;const mob=/Android|iPhone|iPad|iPod/i.test(navigator.userAgent);location.href=l;if(!mob)setTimeout(()=>alert('토스 앱이 설치된 휴대폰에서 송금 페이지가 열립니다.\n휴대폰 카메라로 QR을 스캔해 주세요.'),600)}
function donate(){const d=D.don||{};return `<div><div style="margin:8px 0;line-height:1.6">${esc(d.text||'봇 제작자의 커피타임을 만들어주시면 더욱 좋은 기능으로 찾아뵙겠습니다!')}</div>${d.img?`<img alt="후원 QR" src="data:image/${esc(d.ext||'png')};base64,${d.img}" style="max-width:260px;width:80%;border-radius:12px;border:1px solid var(--line);cursor:pointer" ondblclick="openDon()" title="더블클릭하면 송금 페이지로 이동">`:'<div class="m">QR 이미지를 준비 중입니다.</div>'}${d.link?`<div style="margin-top:10px"><button class="sbtn on" ondblclick="openDon()" onclick="this.nextElementSibling.style.display='block'">💸 송금하기 (더블클릭)</button><div class="m" style="display:none;margin-top:4px">QR 또는 이 버튼을 더블클릭하면 송금 페이지로 이동합니다.</div></div>`:''}<div class="m" style="margin-top:8px">후원은 자율이며, 모든 기능은 후원 없이 이용할 수 있습니다. 디스코드 <b>/후원</b> 명령이나 <b>#후원</b> 채널에서도 볼 수 있습니다.</div></div>`}
window.addEventListener('message',e=>{if(e.data&&e.data.dexSaved)DIRTY=true});
function setTab(t){if(DIRTY&&(t==='build'||t==='market')){DIRTY=false;try{sessionStorage.setItem('rtab',t)}catch(e){}location.reload();return}TAB=t;try{sessionStorage.setItem('rtab',t)}catch(e){}draw()}
function vg(label,v,i,scope){return `<span class="vg">${label}<button class="vb u ${v.m==='u'?'on':''}" onclick="vote(${i},'${scope}','u')">👍 ${v.u}</button><button class="vb d ${v.m==='d'?'on':''}" onclick="vote(${i},'${scope}','d')">👎 ${v.d}</button></span>`}
function tcell(t){return `<div class="tac" data-tip="${esc(tipHtml(t.n))}">${t.img?`<img src="data:image/jpeg;base64,${t.img}">`:'<img>'}<div><div class="t">${esc(t.n)}</div><div class="k"><span class="${KC[t.kind]||''}">${esc(t.kind)}</span> <span class="m">${esc(t.type)}·${esc(t.trait)}</span></div></div></div>`}
function hero(h){
  if(!h.n)return `<div class="h"><div class="empty">${esc(h.src||'보유 장수 부족')}</div></div>`;
  const gem=h.tr>0?`<div class="gems" style="--st:url(${D.stone})">${[1,2,3,4,5].map(k=>`<i class="${h.tr>=k?'on':''}"></i>`).join('')}</div>`:'';
  const L=['무력','지력','통솔','선공'];
  const st=h.fin?L.map((l,k)=>`<span>${l} <b>${h.fin[k]}</b>${h.add[k]?`<em> (+${h.add[k]})</em>`:''}</span>`).join(''):'';
  const tac=h.tac.length?h.tac.map(t=>`<div class="tac">${t.img?`<img src="data:image/jpeg;base64,${t.img}">`:'<img>'}<div><div class="t">${esc(t.n)}</div><div class="k"><span class="${KC[t.kind]||''}">${esc(t.kind)}</span> <span class="m">${esc(t.type)}·${esc(t.trait)} · ${esc(t.why)}</span></div></div></div>`).join(''):`<div class="empty">전법 부족 (보유 전법 없음)</div>`;
  return `<div class="h"><div class="pic">${h.img?`<img class="p" src="data:image/jpeg;base64,${h.img}">`:''}<div class="lb">${gem}50 ${esc(h.n)}</div></div><div class="hi">
  <div class="hn">${esc(h.n)}${h.nation?`<span class="nat" style="background:${NAT[h.nation]||'#999'}">${esc(h.nation)}</span>`:''}${h.rank?`<span class="rk" data-tip="행동 순서">${h.rank}</span>`:''}</div>
  <div class="m">${esc([h.pos,h.troop].filter(Boolean).join(' · '))}</div><div class="sk">${esc(h.sk)}${h.sktype?` (${esc(h.sktype)})`:''}</div>
  <div class="stats">${st}</div>${h.eq?`<div class="eqp" data-tip="특기 대체: ${esc(h.eq.alt)}">🛡 특기 <b>${esc(h.eq.trait)}</b> · 이펙트 <b>${esc(h.eq.effects.join(', '))}</b></div>`:''}${h.src?`<div class="src">← ${esc(h.src)}</div>`:''}</div></div>`}
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
    ${h.tac.map(t=>`<div class="tac" data-tip="${esc(tipHtml(t.n))}">${t.img?`<img src="data:image/jpeg;base64,${t.img}">`:'<img>'}<div><div class="t ${own.has(nk(t.n))?'ownt':''}">${esc(t.n)}${own.has(nk(t.n))?' ✓':''}</div><div class="k"><span class="${KC[t.kind]||''}">${esc(t.kind)}</span> <span class="m">${esc([t.type,t.trait].filter(Boolean).join('·'))}</span></div>${t.alt.length?`<div class="m">대체: ${esc(t.alt.join(' / '))}</div>`:''}</div></div>`).join('')}
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
  const stones=h=>{const n=TRS[h.n]||0;return `<span class="sts" data-tip="돌 단계 (돌을 눌러 변경)" style="--st:url(${D.stone})">${[1,2,3,4,5].map(k=>`<i class="sg ${k<=n?'on':''}" onclick="event.stopPropagation();setTr('${esc(h.n)}',${k})"></i>`).join('')}</span>`};
  const card=h=>`<div class="oc hc ${OWNS.h.has(h.n)?'on':''}" onclick="toggleOwn('h','${esc(h.n)}')">${im(h.img,'th')}<span><span class="nm">${esc(h.n)}</span>${OWNS.h.has(h.n)?stones(h):''}</span></div>`;
  const NO={'촉':0,'위':1,'오':2,'군':3},NN={'촉':'촉','위':'위','오':'오','군':'군웅'};let hc='';
  for(const nt of ['촉','위','오','군']){const g=hs.filter(h=>h.nation===nt);if(g.length)hc+=`<div class="gh">${NN[nt]} (${g.length}명 · 보유 ${g.filter(h=>OWNS.h.has(h.n)).length})</div>`+g.map(card).join('')}
  const oth=hs.filter(h=>!(h.nation in NO));if(oth.length)hc+=`<div class="gh">기타</div>`+oth.map(card).join('');
  let sc='',lg=null;for(const s of ss){if(s.grade!==lg){lg=s.grade;sc+=`<div class="gh">${esc(lg||'기타')}</div>`}
   sc+=`<button class="oc sc ${OWNS.s.has(s.n)?'on':''}" onclick="toggleOwn('s','${esc(s.n).replace(/'/g,"\\'")}')">${im(s.img,'ts')}<span class="nm">${esc(s.n)}</span><span class="m">${esc(s.type)}·${esc(s.trait)}</span></button>`}
  return `<h2>장수 (보유 ${OWNS.h.size})</h2><div class="ocs">${hc||'<span class="m">없음</span>'}</div><h2>전법 (보유 ${OWNS.s.size})</h2><div class="ocs">${sc||'<span class="m">없음</span>'}</div>`}
function own(){TRS=TRS||Object.assign({},D.own.tr||{});OWNS=OWNS||{h:new Set(D.own.h),s:new Set(D.own.s)};
  return `<div class="m" style="margin-bottom:8px">눌러서 보유 여부를 바꿉니다(자동 저장). 보유 장수의 <b>붉은 돌</b>을 눌러 돌 단계(0~5)를 정합니다. 스탯 조율은 <b>장수 도감</b> 탭에서 합니다. 바꾼 뒤 <b>내 추천덱</b> 탭으로 돌아가면 다시 편성합니다.</div>
  <div class="bar"><input id="oq" class="sq" placeholder="이름·키워드 검색 (예: 위협, 치유, 홍수)" value="${esc(OQ)}" oninput="setQ(this.value)"><button class="sbtn ${OF==='all'?'on':''}" onclick="setF('all')">전체</button><button class="sbtn ${OF==='own'?'on':''}" onclick="setF('own')">보유</button><button class="sbtn ${OF==='no'?'on':''}" onclick="setF('no')">미보유</button></div><div id="ol">${ownList()}</div>`}
function eqBody(){const E=D.equip,TC=[['필수','S','#c0392b'],['권장','A','#0f8a63'],['상황별','B','#8a6a1c'],['하위','C','#8a9a94']];const q=EQ.replace(/\s/g,'');
  const tbl=rows=>{const f=rows.filter(r=>!q||(r[0]+r[2]+r[3]).replace(/\s/g,'').includes(q));return `<div class="eqt">${TC.map(([k,L,c])=>{const g=f.filter(r=>r[1]===k);return g.length?`<div class="eqr"><div class="eql" style="background:${c}"><b>${L}</b><span>${k}</span></div><div class="eqc">${g.map(r=>`<div class="eqi" style="border-color:${c}55"><b>${esc(r[0])}</b><span>${esc(r[2])}</span><small>${esc(r[3])}</small></div>`).join('')}</div></div>`:''}).join('')||'<span class="m">검색 결과 없음</span>'}</div>`};
  return `<h2>장비 특기</h2>${tbl(E.traits)}<h2>이펙트</h2>${tbl(E.effects)}`}
function equip(){return `<div class="m" style="margin-bottom:8px">장비 특기 15종·이펙트 19종. 추천 부대 카드의 <b>🛡 특기·이펙트</b>는 장수 역할(병기·책략 딜러/탱커/서포터)과 쓰는 전법 종류로 고른 추천입니다. 평가는 중국 서버 기준이라 한국 서버와 다를 수 있습니다. (출처: cheonha-deck.xyz)</div>
  <div class="bar"><input class="sq" placeholder="특기·이펙트 검색" value="${esc(EQ)}" oninput="EQ=this.value;document.getElementById('el').innerHTML=eqBody()"></div>
  <div id="el">${eqBody()}</div>`}
function frame(v){return `<iframe class="fr" src="/tier?t=${encodeURIComponent(D.t)}&s=${D.season}&embed=1#${v}"></iframe>`}
let MENU=true,LATE=true;try{MENU=sessionStorage.getItem('rmenu')!=='0';LATE=sessionStorage.getItem('rlate')!=='0'}catch(e){}
function toggleMenu(){MENU=!MENU;try{sessionStorage.setItem('rmenu',MENU?'1':'0')}catch(e){}draw()}
function toggleLate(){LATE=!LATE;try{sessionStorage.setItem('rlate',LATE?'1':'0')}catch(e){}draw()}

// ---------- 편성기 · 공유마켓 · 지략가 ----------
const HM=Object.fromEntries(D.all.h.map(h=>[h.n,h])),SM=Object.fromEntries(D.all.s.map(x=>[x.n,x])),SNK=Object.fromEntries(D.all.s.map(x=>[x.n.replace(/[\s·\-]/g,''),x.n]));
const Z4=()=>[0,0,0,0];
const T0=()=>({h:['','',''],t:[['',''],['',''],['','']],f:'',a:[Z4(),Z4(),Z4()]});
const mkU=u=>({h:[...u.h],t:[0,1,2].map(k=>{const a=[...((u.t||[])[k]||[])];while(a.length<2)a.push('');return a}),f:u.f||'',a:[0,1,2].map(k=>{const x=((u.a||[])[k]||[]).slice(0,4).map(v=>+v||0);while(x.length<4)x.push(0);return x}),n:u.n});
const ser=u=>({h:u.h,t:u.t.map(a=>a.filter(Boolean)),f:u.f||'',a:u.a.map((x,k)=>u.h[k]?x:Z4())});
let BLD=(D.draft||[]).map(mkU);while(BLD.length<5)BLD.push(T0());
let PK=null,PQ='',MS='like',MF=false,MT='',MN='',SVT=null;
let ANS=null,MY=null,MKT='req',EXP=new Set(),RT='',RN='',AN='';
const ownH=()=>ANS?ANS.h:(OWNS?OWNS.h:new Set(D.own.h)),ownS=()=>ANS?ANS.s:(OWNS?OWNS.s:new Set(D.own.s));
const usedH=()=>new Set(BLD.flatMap(u=>u.h).filter(Boolean)),usedS=()=>new Set(BLD.flatMap(u=>u.t.flat()).filter(Boolean));
const imj=(b,c)=>b?`<img class="${c}" src="data:image/jpeg;base64,${b}">`:'';
function saveDraft(){if(ANS)return;clearTimeout(SVT);SVT=setTimeout(()=>{fetch('/api/reco',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({t:D.t,season:D.season,op:'draft',units:BLD.some(u=>u.h.some(Boolean))?BLD.map(ser):null})}).catch(()=>{})},700)}
async function mpost(o){try{const r=await fetch('/api/reco',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({t:D.t,season:D.season,...o})});const j=await r.json();if(!j.ok){alert(j.error||'실패');return false}if(j.mk)D.mk=j.mk;return true}catch(e){alert('네트워크 오류');return false}}
function pick(u,k,kind,j){PD=null;PK={u,k,kind,j};PQ='';draw()}
function setPQ(v){PQ=v;const e=document.getElementById('pl');if(e)e.innerHTML=pickBody()}
function assign(name){PD=null;const p=PK;if(!p)return;const u=BLD[p.u];if(p.kind==='h'){u.h[p.k]=name;u.a[p.k]=Z4();fixA(u)}else u.t[p.k][p.j]=name;PK=null;saveDraft();const y=scrollY;draw();scrollTo(0,y)}
function clr(u,k,kind,j){if(kind==='h'){BLD[u].h[k]='';BLD[u].t[k]=['',''];BLD[u].a[k]=Z4()}else BLD[u].t[k][j]='';PK=null;saveDraft();draw()}
function clearAll(){if(!confirm('5부대를 모두 비울까요?'))return;BLD=[0,1,2,3,4].map(T0);PK=null;saveDraft();draw()}
function fillTier(u,name){const d=D.rank.decks.find(x=>x.name===name);if(!d)return;const o=BLD[u];BLD[u]=T0();BLD[u].src=name;BLD[u].f=(d.form||'').split(/[\/·,]/).map(x=>({'안영진':'안형진','추행진':'추형진','기영진':'기형진'}[x.trim()]||x.trim())).find(x=>D.forms[x])||'';const uh=usedH(),us=usedS(),oh=ownH(),os=ownS();
  d.slots.forEach((slot,k)=>{const h=slot.split('/').find(x=>oh.has(x)&&!uh.has(x)&&HM[x]&&(HM[x].season||1)<=D.season);if(!h)return;BLD[u].h[k]=h;uh.add(h);
    ((d.tc||[])[k]||[]).slice(0,2).forEach((cell,j)=>{const t=cell.split(/[>\/]/).map(x=>SNK[x.trim().replace(/[\s·\-]/g,'')]).find(x=>x&&os.has(x)&&!us.has(x)&&(SM[x].season||1)<=D.season);if(t){BLD[u].t[k][j]=t;us.add(t)}})});
  recA(u,true);fixA(BLD[u]);saveDraft();draw()}
function slotH(u,k){const n=BLD[u].h[k],on=PK&&PK.u===u&&PK.k===k&&PK.kind==='h';
  return n?`<div class="bs ${on?'sel':''}" data-du="${u}" data-dk="${k}" onclick="pick(${u},${k},'h')">${imj((HM[n]||{}).img,'th')}<b>${esc(n)}</b><i class="x" onclick="event.stopPropagation();clr(${u},${k},'h')">✕</i></div>`:`<div class="bs empty ${on?'sel':''}" data-du="${u}" data-dk="${k}" onclick="pick(${u},${k},'h')">＋ 장수</div>`}
function slotT(u,k,j){const n=BLD[u].t[k][j],on=PK&&PK.u===u&&PK.k===k&&PK.kind==='t'&&PK.j===j,hh=BLD[u].h[k];
  if(!hh)return`<div class="bs empty dis">전법${j+1}</div>`;
  return n?`<div class="bs t ${on?'sel':''}" data-tu="${u}" data-tk="${k}" data-tj="${j}" onclick="pick(${u},${k},'t',${j})">${imj((SM[n]||{}).img,'ts')}<span>${esc(n)}</span>${tgm(tsOf(n))}<i class="x" onclick="event.stopPropagation();clr(${u},${k},'t',${j})">✕</i></div>`:`<div class="bs empty ${on?'sel':''}" data-tu="${u}" data-tk="${k}" data-tj="${j}" onclick="pick(${u},${k},'t',${j})">＋ 전법${j+1}</div>`}

const CAN_HOVER=window.matchMedia&&matchMedia('(hover:hover) and (pointer:fine)').matches;let PD=null;
const tipOf=x=>`${x.n} · ${x.grade}·${x.type}·${x.trait}${(x.kw||[]).length?'\n키워드: '+x.kw.join(', '):''}\n${x.desc||''}`;
function tacClick(n){if(CAN_HOVER){assign(n);return}PD=PD===n?null:n;const e=document.getElementById('pl'),d=document.getElementById('pdet');if(e)e.innerHTML=pickBody();if(d)d.innerHTML=pdetBody()}
function hov(n){if(!CAN_HOVER)return;const d=document.getElementById('pdet');if(d)d.innerHTML=pdetBody(n,true)}
function pdetBody(nm,hv){nm=nm||PD;if(!PK||PK.kind!=='t')return'';if(!nm)return`<div class="pdet empty"><span class="m">${CAN_HOVER?'전법에 마우스를 올리면 설명이 여기에 고정 표시됩니다':'전법을 누르면 설명이 여기에 표시됩니다'}</span></div>`;const x=SM[nm];if(!x)return'';return `<div class="pdet"><div class="nm">${imj(x.img,'ts')} <b>${esc(x.n)}</b> <span class="m">${esc(x.grade)}·${esc(x.type)}·${esc(x.trait)}</span></div>${(x.kw||[]).length?`<div class="kw" style="margin:2px 0">키워드: ${x.kw.map(esc).join(' · ')}</div>`:''}<div class="dsc">${esc(x.desc||'설명 없음')}</div>${hv?'<div class="m">클릭하면 적용됩니다</div>':`<div class="bar"><button class="sbtn on" onclick="const n=PD;PD=null;assign(n)">✔ 적용</button><button class="sbtn" onclick="PD=null;document.getElementById('pdet').innerHTML=pdetBody();document.getElementById('pl').innerHTML=pickBody()">취소</button></div>`}</div>`}


// 테마 툴팁(PC 호버)
(function(){const t=document.createElement('div');t.id='tip';document.body.appendChild(t);let on=null;
  const mv=e=>{const w=t.offsetWidth,h=t.offsetHeight;let x=e.clientX+14,y=e.clientY+16;if(x+w>innerWidth-8)x=e.clientX-w-14;if(y+h>innerHeight-8)y=e.clientY-h-12;t.style.left=Math.max(6,x)+'px';t.style.top=Math.max(6,y)+'px'};
  document.addEventListener('mouseover',e=>{if(!matchMedia('(hover:hover)').matches)return;const el=e.target.closest&&e.target.closest('[data-tip]');if(!el){if(on){t.style.display='none';on=null}return}if(on===el)return;on=el;t.innerHTML=el.getAttribute('data-tip');t.style.display='block';mv(e)});
  document.addEventListener('mousemove',e=>{if(on)mv(e)});document.addEventListener('scroll',()=>{t.style.display='none';on=null},true)})();
function tipHtml(n){const x=SM[n];if(!x)return'';return `<b>${esc(x.n)}</b> <span class="tm">${esc(x.grade)}·${esc(x.type)}·${esc(x.trait)}</span>${(x.kw||[]).length?`<div class="tk">${x.kw.map(esc).join(' · ')}</div>`:''}<div class="td">${esc(x.desc||'')}</div>`}
function recTac(u,k){const h=BLD[u].h[k];if(!h)return[];const out=[],seen=new Set(),us=usedS(),os=ownS();
  const ds=D.rank.decks.slice().sort((a,b)=>(BLD[u].src===a.name?-1:0)-(BLD[u].src===b.name?-1:0)||tierNum(a.tier)-tierNum(b.tier)||a.idx-b.idx);
  for(const d of ds){const si=d.slots.findIndex(sl=>sl.split('/').includes(h));if(si<0)continue;
    for(const c of ((d.tc||[])[si]||[]).slice(0,2))for(const n of cellNames(c)){if(seen.has(n)||us.has(n)||!os.has(n)||(SM[n].season||1)>D.season)continue;seen.add(n);out.push({n,deck:d.name,tier:d.tier})}}
  return out.slice(0,8)}
function pickBody(){const q=PQ.replace(/\s/g,''),GR={'전설':0,'영웅':1};let body='';
  const grp=(l,key,order,lab)=>{const o=[...order];l.forEach(x=>{const k=key(x);if(!o.includes(k))o.push(k)});return o.map(k=>{const g=l.filter(x=>key(x)===k);return g.length?`<div class="gh">${esc(lab(k))} (${g.length})</div>`+g.map(x=>x._c).join(''):''}).join('')};
  if(PK.kind==='h'){const uh=usedH(),nu=new Set(D.nu||[]);
    const l=D.all.h.filter(h=>ownH().has(h.n)&&!uh.has(h.n)&&!nu.has(h.n)&&(h.season||1)<=D.season&&(!q||(h.n+h.nation+h.troop+h.pos).replace(/\s/g,'').includes(q))).sort((a,b)=>a.n.localeCompare(b.n,'ko'))
      .map(h=>({...h,_c:`<button class="oc hc on" onclick="assign('${esc(h.n)}')">${imj(h.img,'th')}<span class="nm">${esc(h.n)}</span><span class="m">${esc(h.troop)}</span></button>`}));
    body=grp(l,h=>h.nation||'기타',['촉','위','오','군'],k=>({'군':'군웅'}[k]||k))||'<span class="m">쓸 수 있는 보유 장수가 없습니다 (이미 다른 부대에서 썼거나 미보유)</span>'}
  else{const us=usedS();
    const l=D.all.s.filter(x=>ownS().has(x.n)&&!us.has(x.n)&&(x.season||1)<=D.season&&(!q||(x.n+x.trait+x.type+x.grade+(x.kw||[]).join('')+(x.desc||'')).replace(/\s/g,'').includes(q))).sort((a,b)=>((GR[a.grade]??9)-(GR[b.grade]??9))||a.n.localeCompare(b.n,'ko'))
      .map(x=>({...x,_c:`<button class="oc sc on ${PD===x.n?'pd':''}" onmouseenter="hov('${esc(x.n).replace(/'/g,"\\'")}')" onclick="tacClick('${esc(x.n).replace(/'/g,"\\'")}')">${imj(x.img,'ts')}<span class="nm">${esc(x.n)}</span><span class="m">${esc(x.grade)}·${esc(x.type)}</span>${(x.kw||[]).length?`<span class="kw">${(x.kw||[]).map(esc).join(' · ')}</span>`:''}</button>`}));
    const rc=recTac(PK.u,PK.k).filter(r=>l.some(x=>x.n===r.n)),recH=rc.length&&!q?`<div class="gh rec">⭐ 추천 전법 — ${esc(BLD[PK.u].h[PK.k])} (티어덱 기준 · ${rc.length})</div>`+rc.map(r=>l.find(x=>x.n===r.n)._c.replace('</button>',`<span class="kw rc">${esc(r.tier)} ${esc(r.deck)}</span></button>`)).join(''):'';
    body=recH+(grp(l,x=>x.trait||'기타',['책략','병기','방어','치유','보조','문무'],k=>k)||'')||'<span class="m">쓸 수 있는 보유 전법이 없습니다 (이미 다른 부대에서 썼거나 미보유)</span>'}
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
  while(res.length<5)res.push(T0());BLD=res.slice(0,5);BLD.forEach((u,i)=>{if(u.src)recA(i,true)});BLD.forEach(fixA);saveDraft();draw()}

const SL=['무력','지력','통솔','선공'];
const trOf=h=>ANS?((ANS.tr||{})[h]||0):(((TRS||D.own.tr)||{})[h]||0);
const capOf=h=>50+10*trOf(h);
function defA(h){if(ANS)return Z4();const x=((D.own.tra||{})[h]||[]).slice(0,4).map(v=>+v||0);while(x.length<4)x.push(0);return clampA(x,h)}
function clampA(x,h){let c=capOf(h);return x.map(v=>{v=Math.max(0,Math.min(v,c));c-=v;return v})}
function fixA(u){[0,1,2].forEach(k=>{const h=u.h[k];if(!h)u.a[k]=Z4();else if(!u.a[k].some(Boolean))u.a[k]=defA(h);else u.a[k]=clampA(u.a[k],h)})}
function setA(i,k,q,v,rel){const h=BLD[i].h[k];if(!h)return;const a=BLD[i].a[k],oth=a.reduce((p,c,j)=>j===q?p:p+c,0);let n=rel?a[q]+v:(parseInt(v)||0);n=Math.max(0,Math.min(n,capOf(h)-oth));a[q]=n;saveDraft();const y=scrollY;draw();scrollTo(0,y)}
const stones=n=>n>0?`<span class="sts" style="--st:url(${D.stone})">${[1,2,3,4,5].map(k=>`<i class="sg ${k<=n?'on':''}"></i>`).join('')}</span>`:'';
const tsOf=t=>ANS?((ANS.trs||{})[t]||0):(((D.own||{}).trs||{})[t]||0);
const tgm=n=>n>0?`<b class="tg" style="--st:url(${D.stone})">${n}</b>`:'';
const SI=['⚔','🔥','🛡','➤'];
function statBox(i,k){const h=BLD[i].h[k];if(!h)return'<div></div>';const a=BLD[i].a[k],n=trOf(h),cap=capOf(h),used=a.reduce((p,c)=>p+c,0),bs=(HM[h]||{}).bs;
  const dk=BLD[i].src&&D.rank.decks.find(x=>x.name===BLD[i].src),rec=((dk&&dk.ts||[])[k]||[]).join(' ');
  return `<div class="sbx ${SOP===i+'-'+k?'op':''}"><div class="sbh"><b class="sbn">${esc(h)}</b><span class="m">돌</span>${stones(n)||'<span class="m">0단계</span>'}</div>
  <div class="sbt"><button class="sbtn" onclick="resetA(${i},${k})">↺ 초기화</button><span class="avl">사용 가능 <b class="${used>=cap?'full':''}">${cap-used}</b><i>/${cap}</i></span></div>
  ${SL.map((l,q)=>`<div class="sbr"><span class="sic">${SI[q]}</span><span class="sl">${l}</span>${rec.includes(l)?'<span class="rcm">추천</span>':''}<span class="sv">${bs?bs[q]+a[q]:a[q]}${a[q]&&bs?`<em>(+${a[q]})</em>`:''}</span><button class="stb" ${a[q]?'':'disabled'} onclick="setA(${i},${k},${q},-1,1)">−</button><button class="stb" ${used>=cap?'disabled':''} onclick="setA(${i},${k},${q},1,1)">＋</button></div>`).join('')}</div>`}
function boostOf(i,k){return (BLD[i].t[k]||[]).filter(Boolean).reduce((p,t)=>{const m=((SM[t]||{}).desc||'').match(/자신의 선공이\s*[\d.]+\s*→\s*([\d.]+)\s*포인트\s*(?:증가|상승)/);return p+(m?parseFloat(m[1]):(t==='신속전개'?30:0))},0)}
function recA(i,quiet){const u=BLD[i],dk=u.src&&D.rank.decks.find(x=>x.name===u.src);const notes=[];
  const ks=[0,1,2].filter(k=>u.h[k]);
  const info=ks.map(k=>{const h=u.h[k],bs=(HM[h]||{}).bs,txt=((dk&&dk.ts||[])[k]||[]).join(' '),uniq=[...new Set([...txt.matchAll(/무력|지력|통솔|선공/g)].map(m=>m[0]))];
    let pr=uniq.find(x=>x!=='선공')||(uniq.length?'선공':null);
    if(!pr&&bs){const m=Math.max(bs[0],bs[1],bs[2]);pr=SL[[0,1,2].find(q=>bs[q]===m)]}
    return {k,h,bs,txt,uniq,pr,cap:capOf(h)}});
  info.forEach(o=>{const a=Z4();if(o.pr){const q=SL.indexOf(o.pr);a[q]=o.cap;if(q!==3&&o.uniq.includes('선공')){const f=o.txt.match(/선공\s*\+\s*(\d+)/);const m=/보다\s*높/.test(o.txt)?0:Math.min(o.cap,f?+f[1]:Math.round(o.cap*.3));a[q]-=m;a[3]=m}}u.a[o.k]=a});
  const cons=[];
  info.forEach(o=>{const m=o.txt.match(/선공\s*([가-힣]+)\s*보다\s*높/);if(!m)return;const x=info.find(y=>y.h===m[1]||m[1].includes(y.h)||y.h.includes(m[1]));if(x&&x!==o)cons.push([o,x])});
  const rear=info.find(y=>(y.h==='황충'||y.h==='관우'))&&info.find(y=>y.h==='유비')&&info.find(y=>y.h==='장비')?info.find(y=>y.h==='황충'||y.h==='관우'):null;
  if(rear){['장비','유비'].forEach(n=>{const o=info.find(y=>y.h===n);if(o&&!cons.some(c=>c[0]===o&&c[1]===rear))cons.push([o,rear])})}
  cons.forEach(([o,x])=>{const ai=u.a[o.k],ax=u.a[x.k],bo=o.bs?o.bs[3]:null,bx=x.bs?x.bs[3]:null,bt=boostOf(i,x.k);
    let need=(bo!=null&&bx!=null)?(bx+ax[3]+bt)-(bo+ai[3])+1:20+bt;const tb=bt?` (${x.h} 전법 선공 +${bt} 반영)`:'';
    if(need<=0){notes.push(`${o.h} 선공이 이미 ${x.h}보다 높음${tb}`);return}
    const src=o.pr&&o.pr!=='선공'?SL.indexOf(o.pr):-1;let mv=0;if(src>=0){mv=Math.min(need,ai[src]);ai[src]-=mv;ai[3]+=mv;need-=mv}
    if(need>0&&ax[3]>0){const d=Math.min(need,ax[3]);ax[3]-=d;need-=d}
    notes.push(need>0?`${o.h} 선공을 ${x.h}보다 높이기엔 포인트 부족${tb}`:`${o.h} 선공 ${ai[3]} 투자 → ${x.h}보다 먼저 행동${tb}`)});
  u.rn=notes;if(!quiet){saveDraft();const y=scrollY;draw();scrollTo(0,y)}}
let SOP=null;
function togS(i,k){const key=i+'-'+k;SOP=SOP===key?null:key;const y=scrollY;draw();scrollTo(0,y)}
function statSum(i,k){const h=BLD[i].h[k];if(!h)return'<div class="sbs0"></div>';const a=BLD[i].a[k],bs=(HM[h]||{}).bs,cap=capOf(h),left=cap-a.reduce((p,c)=>p+c,0),op=SOP===i+'-'+k;
  return `<button class="sbs ${op?'op':''}" onclick="togS(${i},${k})"><span class="sn">${SL.map((l,q)=>`<span><i>${l[0]}</i>${bs?bs[q]+a[q]:a[q]}</span>`).join('')}</span><span class="sm">스탯 조절 · 남은 ${left} ${op?'▴':'▾'}</span></button>`}
function resetA(i,k){BLD[i].a[k]=Z4();saveDraft();const y=scrollY;draw();scrollTo(0,y)}
const aTxt=a=>SL.map((l,q)=>a&&a[q]?l[0]+a[q]:'').filter(Boolean).join(' ');
function setF(i,v){BLD[i].f=v;saveDraft();draw()}
function resetUnit(i){if(!BLD[i].h.some(Boolean)||confirm(`${i+1}부대를 초기화할까요?`)){BLD[i]=T0();PK=null;saveDraft();draw()}}
// ---- 끌어서 옮기기(마우스 즉시 / 터치 길게 누르기) ----
let DR=null,SUP=0;
function drSrc(el){if(el.dataset.tu!=null)return{kind:'t',u:+el.dataset.tu,k:+el.dataset.tk,j:+el.dataset.tj};if(el.dataset.du!=null)return{kind:'h',u:+el.dataset.du,k:+el.dataset.dk}}
function drBegin(){if(!DR||DR.on)return;const s=DR.src;if(s.kind==='t'&&!BLD[s.u].t[s.k][s.j])return;if(s.kind==='h'&&!BLD[s.u].h[s.k])return;DR.on=true;PK=null;
  const g=document.createElement('div');g.className='dgh';g.textContent=s.kind==='t'?BLD[s.u].t[s.k][s.j]:BLD[s.u].h[s.k];document.body.appendChild(g);DR.g=g;DR.el.classList.add('dsrc');document.body.classList.add('dragging');if(navigator.vibrate)try{navigator.vibrate(15)}catch(e){}}
function drTarget(x,y,kind){kind=kind||(DR&&DR.src.kind);const e=document.elementFromPoint(x,y);if(!e)return null;const el=e.closest('.bs[data-tu],.bs[data-du]');if(!el)return null;const t=drSrc(el);return t&&t.kind===kind?{el,t}:null}
document.addEventListener('pointerdown',e=>{const el=e.target.closest&&e.target.closest('.bs[data-tu],.bs[data-du]');if(!el||e.target.closest('.x')||e.button>0)return;const src=drSrc(el);
  DR={el,src,x:e.clientX,y:e.clientY,touch:e.pointerType==='touch',on:false,timer:null};if(DR.touch)DR.timer=setTimeout(drBegin,280)});
document.addEventListener('pointermove',e=>{if(!DR)return;const dx=e.clientX-DR.x,dy=e.clientY-DR.y,mv=Math.hypot(dx,dy);
  if(!DR.on){if(DR.touch){if(mv>10){clearTimeout(DR.timer);DR=null}}else if(mv>6)drBegin();return}
  DR.g.style.left=e.clientX+12+'px';DR.g.style.top=e.clientY+12+'px';document.querySelectorAll('.dz').forEach(x=>x.classList.remove('dz'));const t=drTarget(e.clientX,e.clientY);if(t&&t.el!==DR.el)t.el.classList.add('dz')});
document.addEventListener('touchmove',e=>{if(DR&&DR.on)e.preventDefault()},{passive:false});
function drEnd(e,cancel){if(!DR)return;clearTimeout(DR.timer);const d=DR;DR=null;if(!d.on)return;d.g.remove();document.body.classList.remove('dragging');document.querySelectorAll('.dz,.dsrc').forEach(x=>x.classList.remove('dz','dsrc'));SUP=Date.now();if(cancel)return;
  const t=drTarget(e.clientX,e.clientY,d.src.kind);if(!t)return;const a=d.src,b=t.t;if(a.u===b.u&&a.k===b.k&&(a.kind==='h'||a.j===b.j))return;
  if(a.kind==='t'){const x=BLD[a.u].t[a.k][a.j];BLD[a.u].t[a.k][a.j]=BLD[b.u].t[b.k][b.j];BLD[b.u].t[b.k][b.j]=x}
  else{const h=BLD[a.u].h[a.k],tt=BLD[a.u].t[a.k],ss=BLD[a.u].a[a.k];BLD[a.u].h[a.k]=BLD[b.u].h[b.k];BLD[a.u].t[a.k]=BLD[b.u].t[b.k];BLD[a.u].a[a.k]=BLD[b.u].a[b.k];BLD[b.u].h[b.k]=h;BLD[b.u].t[b.k]=tt;BLD[b.u].a[b.k]=ss}
  saveDraft();draw()}
document.addEventListener('pointerup',e=>drEnd(e,false));document.addEventListener('pointercancel',e=>drEnd(e,true));
document.addEventListener('click',e=>{if(Date.now()-SUP<350){e.stopPropagation();e.preventDefault()}},true);
document.addEventListener('dragstart',e=>{if(e.target.closest&&e.target.closest('.bs'))e.preventDefault()});document.addEventListener('contextmenu',e=>{if(DR&&DR.on)e.preventDefault()});
// ---- 공유하기: 5부대를 이미지로 ----
function loadImg(b){return new Promise(r=>{if(!b)return r(null);const i=new Image();i.onload=()=>r(i);i.onerror=()=>r(null);i.src='data:image/jpeg;base64,'+b})}
async function shareImg(){if(!BLD.some(u=>u.h.some(Boolean))){alert('편성이 비어 있습니다');return}
  const W=1080,RH=176,H=118+5*RH+54,c=document.createElement('canvas');c.width=W;c.height=H;const x=c.getContext('2d');const F="'Noto Sans KR','Apple SD Gothic Neo','Malgun Gothic',sans-serif";
  x.fillStyle='#0f1a18';x.fillRect(0,0,W,H);const g=x.createLinearGradient(0,0,W,0);g.addColorStop(0,'#0d4c43');g.addColorStop(1,'#1f8f7a');x.fillStyle=g;x.fillRect(0,0,W,96);
  x.fillStyle='#fff';x.font=`800 38px ${F}`;x.fillText((MT.trim()||'내 5부대 편성'),30,50);x.font=`500 20px ${F}`;x.fillStyle='rgba(255,255,255,.85)';x.fillText(`${D.who} · 시즌${D.season} · 삼국지 천하결전 동맹 봇`,30,80);
  for(let i=0;i<5;i++){const y=108+i*RH;x.fillStyle='#16262a';rr(x,14,y,W-28,RH-10,16);x.fill();x.strokeStyle='#2b4448';x.lineWidth=2;x.stroke();
    x.fillStyle='#f0c75e';rr(x,28,y+12,74,30,15);x.fill();x.fillStyle='#2b2000';x.font=`800 18px ${F}`;x.textAlign='center';x.fillText(`${i+1}부대`,65,y+33);x.textAlign='left';
    {const lb=[BLD[i].src?BLD[i].src+' 기준':'',BLD[i].f?'진형 '+BLD[i].f:''].filter(Boolean).join(' · ');if(lb){x.fillStyle='#8fa8a3';x.font=`500 16px ${F}`;x.fillText(lb,112,y+33)}}
    for(let k=0;k<3;k++){const n=BLD[i].h[k],ox=28+k*350;if(!n){x.fillStyle='#223338';rr(x,ox,y+50,330,100,10);x.fill();continue}
      const im=await loadImg((HM[n]||{}).img2||(HM[n]||{}).img);if(im){x.save();rr(x,ox,y+48,84,106,10);x.clip();x.drawImage(im,ox,y+48,84,106);x.restore()}
      x.fillStyle='#e8efe9';x.font=`800 24px ${F}`;x.fillText(n,ox+96,y+70);const nw=x.measureText(n).width;x.fillStyle='#8fa8a3';x.font=`500 15px ${F}`;x.fillText([(HM[n]||{}).nation,(HM[n]||{}).troop,trOf(n)?'돌'+trOf(n):'',aTxt(BLD[i].a[k])].filter(Boolean).join('·'),ox+96+nw+8,y+70);
      for(let j=0;j<2;j++){const t=BLD[i].t[k][j];if(!t)continue;const ti=await loadImg((SM[t]||{}).img2||(SM[t]||{}).img),yy=y+108+j*34;if(ti){x.save();rr(x,ox+96,yy-17,28,28,6);x.clip();x.drawImage(ti,ox+96,yy-17,28,28);x.restore()}
        x.fillStyle='#5fcfb1';x.font=`600 18px ${F}`;x.fillText(t,ox+132,yy+4)}}}
  x.fillStyle='#8fa8a3';x.font=`500 16px ${F}`;x.fillText('한 번 쓴 장수·전법은 다른 부대에서 쓸 수 없음 · 무소유',30,H-20);
  const url=c.toDataURL('image/png');const m=document.createElement('div');m.className='shm';m.innerHTML=`<div class="shb"><div class="bar"><b>📷 내 5부대 편성 이미지</b><span class="sp"></span><a class="sbtn on" download="내_5부대_편성.png" href="${url}">⬇ 저장</a><button class="sbtn" onclick="this.closest('.shm').remove()">닫기</button></div><div class="m">모바일은 이미지를 길게 눌러 저장하세요</div><img src="${url}" alt="편성"></div>`;m.onclick=e=>{if(e.target===m)m.remove()};document.body.appendChild(m)}
function rr(x,a,b,w,h,r){x.beginPath();x.moveTo(a+r,b);x.arcTo(a+w,b,a+w,b+h,r);x.arcTo(a+w,b+h,a,b+h,r);x.arcTo(a,b+h,a,b,r);x.arcTo(a,b,a+w,b,r);x.closePath()}
function build(){const dk=D.rank.decks.slice().sort((a,b)=>a.idx-b.idx);
  const units=BLD.map((u,i)=>`<div class="unit"><div class="uh"><span class="badge">${i+1}부대</span>${BLD[i].src?`<span class="m">${esc((D.rank.decks.find(x=>x.name===BLD[i].src)||{}).tier||'')} <b>${esc(BLD[i].src)}</b> 기준</span>`:''}<select class="sq" onchange="if(this.value){fillTier(${i},this.value);}this.value=''"><option value="">티어덱 불러오기…</option>${dk.map(d=>`<option value="${esc(d.name)}">${esc(d.tier)} ${esc(d.name)}</option>`).join('')}</select><button class="sbtn" onclick="resetUnit(${i})">↺ 부대 초기화</button><button class="sbtn on" onclick="recA(${i})">✨ 추천 속성치</button><span class="m dh">⠿ 장수·전법을 끌어서 옮기기(모바일은 길게 누르기)</span></div>
    <div class="hs fmr"><span class="m">진형</span><select class="sq" onchange="setF(${i},this.value)"><option value="">진형 선택…</option>${Object.keys(D.forms).map(n=>`<option ${u.f===n?'selected':''}>${esc(n)}</option>`).join('')}</select>${u.f?fm(u.f):''}</div>
    ${(u.rn||[]).length?`<div class="m rn">💡 ${u.rn.map(esc).join(' · ')}</div>`:''}<div class="hcols">${[0,1,2].map(k=>`<div class="hblk"><div class="hcol"><div class="gi r1 c${k}">${slotH(i,k)}</div><div class="gi r3 c${k}">${slotT(i,k,0)}</div><div class="gi r4 c${k}">${slotT(i,k,1)}</div></div><div class="hst"><div class="gi r2 c${k} sumw">${statSum(i,k)}</div><div class="gi r5 stw ${SOP===i+'-'+k?'op':''}">${statBox(i,k)}</div></div></div>`).join('')}</div></div>`).join('');
  const n=BLD.flatMap(u=>u.h).filter(Boolean).length;
  return `<div class="m" style="margin-bottom:8px">내가 <b>보유한</b> 장수·전법으로 5부대를 직접 짭니다. <b>한 번 쓴 장수·전법은 다른 부대에서 쓸 수 없습니다</b>(목록에서 사라짐). 편성은 자동 저장됩니다. 보유 체크는 <b>내 장수·전법</b> 탭에서 합니다.</div>
  ${ANS?`<div class="unit ansb"><div class="uh"><b>✍ 의뢰 답변 작성 중</b><span class="m">‘${esc(ANS.title)}’ · 의뢰자 <b>${esc(ANS.by)}</b>의 보유 장수 ${ANS.h.size}명·전법 ${ANS.s.size}개로만 편성할 수 있습니다</span></div></div>`:''}${picker()}<div class="bar"><button class="sbtn on" onclick="if(confirm('내 보유 장수·전법에 맞춰 티어 순서대로 5부대를 자동 편성합니다. 지금 편성은 덮어씁니다.'))autoCompose()">✨ 내 상황 맞춤 자동 편성 (티어순)</button><span class="m">부족한 자리는 비슷한 보유 장수·같은 계열 전법으로 메웁니다</span></div>${units}${ANS?`<div class="unit"><div class="uh"><b>답변 제출</b><span class="m">장수 ${n}명 · 같은 의뢰에는 한 번만 답변할 수 있고, 다시 제출하면 내 답변이 교체됩니다</span></div>
  <div class="bar"><input class="sq" style="flex:1;min-width:180px" maxlength="120" placeholder="한 줄 설명 (선택: 이 편성의 포인트)" value="${esc(AN)}" oninput="AN=this.value"></div>
  <div class="bar"><button class="sbtn on" onclick="submitAns()">✍ 답변 올리기</button><button class="sbtn" onclick="shareImg()">📷 이미지</button><button class="sbtn" onclick="clearAll()">⟲ 전체 초기화</button><button class="sbtn" onclick="endAns()">취소(내 편성기로 복귀)</button></div></div>`:`<div class="unit"><div class="uh"><b>완성 편성 공유하기</b><span class="m">장수 ${n}명 · 올린 글은 모든 서버가 보고 추천할 수 있습니다 (최대 ${D.mk.max}개)</span></div>
  <div class="bar"><input class="sq" style="flex:1;min-width:140px" maxlength="24" placeholder="제목 (예: 도트딜 5부대)" value="${esc(MT)}" oninput="MT=this.value"><input class="sq" style="flex:2;min-width:180px" maxlength="120" placeholder="한 줄 설명 (선택)" value="${esc(MN)}" oninput="MN=this.value"></div>
  <div class="bar"><button class="sbtn on" onclick="postMk()">🛒 완성 편성으로 공유하기</button><button class="sbtn" onclick="shareImg()">📷 공유하기(이미지)</button><button class="sbtn" onclick="clearAll()">⟲ 전체 초기화</button></div></div>`}`}
async function postMk(){if(!BLD.some(u=>u.h.some(Boolean))){alert('장수를 한 명 이상 넣어 주세요');return}if(!MT.trim()){alert('제목을 입력해 주세요');return}
  if(await mpost({op:'mk_post',title:MT,note:MN,units:BLD.map(ser)})){MT='';MN='';alert('완성 편성 공유에 올렸습니다');setTab('market')}}
async function disMk(id){if(await mpost({op:'mk_dis',id})){const y=scrollY;draw();scrollTo(0,y)}}
async function likeMk(id){if(await mpost({op:'mk_like',id})){const y=scrollY;draw();scrollTo(0,y)}}
async function delMk(id){if(!confirm('이 편성을 삭제할까요? (받은 추천도 사라집니다)'))return;if(await mpost({op:'mk_del',id}))draw()}
function mkStat(p){const oh=ownH(),os=ownS();let a=0,b=0;p.units.forEach(u=>u.h.forEach((h,k)=>{if(h){b++;if(oh.has(h))a++;(u.t[k]||[]).forEach(t=>{b++;if(os.has(t))a++})}}));return[a,b]}
function unitsHtml(units,own){const oh=ownH(),os=ownS();return units.map((u,i)=>{
  const cells=[0,1,2].map(k=>{const h=u.h[k];if(!h)return `<div class="gi r1 c${k}"><div class="bs empty">-</div></div>`;
    const hv=own?true:oh.has(h),a=(u.a||[])[k]||[0,0,0,0],bs=(HM[h]||{}).bs,n=(u.n||[])[k]||0;
    const hero=`<div class="gi r1 c${k}"><div class="bs ro ${hv?'':'no'}">${imj((HM[h]||{}).img,'th')}<b>${esc(h)}</b>${hv?'':'<i class="x1">미보유</i>'}</div></div>`;
    const sum=`<div class="gi r2 c${k}"><div class="sbs ro"><span class="sn">${SL.map((l,q)=>`<span><i>${l[0]}</i>${bs?bs[q]+(a[q]||0):(a[q]||0)}${a[q]&&bs?`<em>+${a[q]}</em>`:''}</span>`).join('')}</span><span class="sm">${stones(n)||'돌 0'}</span></div></div>`;
    const tac=[0,1].map(j=>{const t=((u.t||[])[k]||[])[j];return `<div class="gi r${3+j} c${k}">${t?`<div class="bs t ro ${own||os.has(t)?'':'no'}" data-tip="${esc(tipHtml(t))}">${imj((SM[t]||{}).img,'ts')}<span>${esc(t)}</span>${tgm((((u.tn||[])[k]||[])[j])|0)}${own||os.has(t)?'':'<i class="x1">✗</i>'}</div>`:'<div class="bs empty">-</div>'}</div>`}).join('');
    return hero+sum+tac}).join('');
  return `<div class="unit mkv"><div class="uh"><span class="badge">${i+1}부대</span>${u.f?`<span class="m">진형</span>${fm(u.f)}`:''}</div><div class="hcols">${cells}</div></div>`}).join('')}
function loadFrom(units){if(!confirm('내 편성기에 가져옵니다. 보유하지 않은 장수·전법은 빈 칸이 되며 지금 편성은 덮어씁니다. 계속할까요?'))return;
  const oh=ownH(),os=ownS(),uh=new Set(),us=new Set();BLD=units.map(u=>{const o=T0();o.f=u.f||'';[0,1,2].forEach(k=>{const h=u.h[k];if(h&&oh.has(h)&&!uh.has(h)){o.h[k]=h;o.a[k]=((u.a||[])[k]||Z4()).slice();uh.add(h);(u.t[k]||[]).slice(0,2).forEach((t,j)=>{if(os.has(t)&&!us.has(t)){o.t[k][j]=t;us.add(t)}})}});return o});while(BLD.length<5)BLD.push(T0());BLD.forEach(fixA);saveDraft();setTab('build')}
function loadMk(id){const p=D.mk.posts.find(x=>x.id===id);if(p)loadFrom(p.units)}
function loadAns(rid,aid){const r=D.mk.reqs.find(x=>x.id===rid),a=r&&r.ans.find(x=>x.id===aid);if(a)loadFrom(a.units)}
const sb=p=>p.strat?' <span class="strat">🧠 지략가</span>':'';
function mkTabs(){return `<div class="bar"><button class="sbtn ${MKT==='req'?'on':''}" onclick="MKT='req';draw()">🙋 편성 의뢰 <span class="m">(${D.mk.reqs.filter(r=>r.season===D.season).length})</span></button><button class="sbtn ${MKT==='done'?'on':''}" onclick="MKT='done';draw()">🛒 완성 편성 공유 <span class="m">(${D.mk.posts.filter(p=>p.season===D.season).length})</span></button></div>`}
async function postReq(){if(await mpost({op:'rq_post',title:RT,note:RN})){RT='';RN='';alert('의뢰를 올렸습니다. 다른 동맹원들이 내 보유 장수·전법으로 5부대를 짜 줍니다');draw()}}
async function delReq(id){if(!confirm('이 의뢰와 달린 답변을 모두 삭제할까요?'))return;if(await mpost({op:'rq_del',id}))draw()}
async function delAns(id){if(!confirm('내 답변을 삭제할까요?'))return;if(await mpost({op:'ans_del',id}))draw()}
async function adopt(rid,aid){if(await mpost({op:'rq_adopt',rid,aid})){const y=scrollY;draw();scrollTo(0,y)}}
function togReq(id){EXP.has(id)?EXP.delete(id):EXP.add(id);const y=scrollY;draw();scrollTo(0,y)}
function startAns(rid){const r=D.mk.reqs.find(x=>x.id===rid);if(!r)return;MY=BLD;const old=r.ans.find(a=>a.mine);
  ANS={rid,title:r.title,by:r.by,h:new Set(r.own.h),s:new Set(r.own.s),tr:r.own.tr||{},trs:r.own.trs||{}};BLD=(old?old.units:[]).map(mkU);while(BLD.length<5)BLD.push(T0());BLD.forEach(fixA);AN=old?old.note:'';PK=null;setTab('build')}
function endAns(){ANS=null;BLD=MY||BLD;MY=null;PK=null;draw()}
async function submitAns(){if(!BLD.some(u=>u.h.some(Boolean))){alert('장수를 한 명 이상 넣어 주세요');return}
  if(await mpost({op:'ans_post',rid:ANS.rid,note:AN,units:BLD.map(ser)})){alert('답변을 올렸습니다. 의뢰자가 추천·채택하면 지략가 점수가 올라갑니다');AN='';ANS=null;BLD=MY||BLD;MY=null;MKT='req';setTab('market')}}
function reqCard(r){const open=EXP.has(r.id),n=r.ans.length,hs=r.own.h.map(h=>`<span class="hchip" data-tip="${esc(h)}${(r.own.tr||{})[h]?' · 돌 '+r.own.tr[h]:''}">${imj((HM[h]||{}).img,'th')}${(r.own.tr||{})[h]?`<i class="ct">${r.own.tr[h]}</i>`:''}</span>`).join('');
  return `<div class="unit mk ${r.done?'done':''}"><div class="uh"><span class="nm">${r.done?'✅ ':''}${esc(r.title)}</span><span class="m">의뢰 <b>${esc(r.by)}</b>${sb(r)}</span>
    <div class="votes"><span class="m">보유 장수 ${r.own.h.length} · 전법 ${r.own.s.length} · 답변 ${n}</span>${r.mine?`<button class="vb d" onclick="delReq('${r.id}')">삭제</button>`:''}${r.done?'':`<button class="sbtn on" onclick="startAns('${r.id}')">✍ ${r.ans.some(a=>a.mine)?'내 답변 수정':(r.mine?'내가 직접 편성':'편성해주기')}</button>`}<button class="sbtn" onclick="togReq('${r.id}')">${open?'접기 ▴':`답변 보기 ▾`}</button></div></div>
    ${r.note?`<div class="m" style="margin-bottom:6px">${esc(r.note)}</div>`:''}<div class="hchips">${hs}</div>
    ${open?(n?r.ans.map(a=>`<div class="ans ${a.adopted?'ad':''}"><div class="uh"><span class="m">${a.adopted?'<b class="ownt">✅ 채택된 편성</b> · ':''}${a.mine?'<b>내 답변</b>':`편성 <b>${esc(a.by)}</b>`}${sb(a)}${a.note?' · '+esc(a.note):''}</span>
      <div class="votes"><button class="vb u ${a.liked?'on':''}" ${a.mine?'disabled':''} onclick="likeMk('${a.id}')">👍 ${a.likes}</button><button class="vb d ${a.disliked?'on':''}" ${a.mine?'disabled':''} onclick="disMk('${a.id}')">👎 ${a.dis||0}</button>${r.mine?`<button class="sbtn ${a.adopted?'on':''}" onclick="adopt('${r.id}','${a.id}')">${a.adopted?'채택 취소':'✔ 채택'}</button><button class="sbtn" onclick="loadAns('${r.id}','${a.id}')">내 편성기로 가져오기</button>`:''}${a.mine&&!a.adopted?`<button class="vb d" onclick="delAns('${a.id}')">삭제</button>`:''}</div></div>${unitsHtml(a.units,true)}</div>`).join(''):'<div class="m" style="margin-top:8px">아직 답변이 없습니다.</div>'):''}</div>`}
let BT='',BB='',BX=new Set(),BC={},BI='',BCI={};
function shrink(file){return new Promise((ok,no)=>{const r=new FileReader();r.onerror=()=>no();r.onload=()=>{const im=new Image();im.onerror=()=>no();im.onload=()=>{const M=800,k=Math.min(1,M/Math.max(im.width,im.height)),c=document.createElement('canvas');c.width=Math.round(im.width*k);c.height=Math.round(im.height*k);const x=c.getContext('2d');x.fillStyle='#fff';x.fillRect(0,0,c.width,c.height);x.drawImage(im,0,0,c.width,c.height);let q=.78,u=c.toDataURL('image/jpeg',q);while(u.length*0.75>110000&&q>.3){q-=.08;u=c.toDataURL('image/jpeg',q)}u.length*0.75>115000?no():ok(u.split(',')[1])};im.src=r.result};r.readAsDataURL(file)})}
async function pickImg(inp,pid){const f=inp.files&&inp.files[0];if(!f)return;try{const b=await shrink(f);if(pid)BCI[pid]=b;else BI=b;const y=scrollY;draw();scrollTo(0,y)}catch(e){alert('이미지를 불러오지 못했거나 너무 큽니다. 다른 이미지를 선택해 주세요.')}}
function noImg(pid){if(pid)delete BCI[pid];else BI='';draw()}
const imgTag=(id,u)=>id?`<a href="/bimg/${id}" target="_blank"><img class="bimg" loading="lazy" src="/bimg/${id}"></a>`:''
const prevTag=(b,pid)=>b?`<div class="bprev"><img class="bimg" src="data:image/jpeg;base64,${b}"><button class="sbtn" onclick="noImg(${pid?`'${pid}'`:''})">✕ 이미지 빼기</button></div>`:''
const ago=t=>{const s=Math.max(0,Date.now()/1000-t);return s<60?'방금':s<3600?Math.floor(s/60)+'분 전':s<86400?Math.floor(s/3600)+'시간 전':Math.floor(s/86400)+'일 전'};
async function postBd(){if(await mpost({op:'bd_post',title:BT,body:BB,img:BI})){BT='';BB='';BI='';draw()}}
async function delBd(id){if(!confirm('이 글과 댓글을 삭제할까요?'))return;if(await mpost({op:'bd_del',id}))draw()}
async function comBd(pid){const v=(BC[pid]||'').trim();if(!v&&!BCI[pid])return;if(await mpost({op:'bd_com',pid,body:v,img:BCI[pid]||''})){BC[pid]='';delete BCI[pid];const y=scrollY;draw();scrollTo(0,y)}}
async function delCom(id){if(await mpost({op:'bd_cdel',id})){const y=scrollY;draw();scrollTo(0,y)}}
function togBd(id){BX.has(id)?BX.delete(id):BX.add(id);const y=scrollY;draw();scrollTo(0,y)}
function board(){const l=D.mk.board||[];
  return `<div class="m" style="margin-bottom:8px">자유롭게 이야기하는 게시판입니다. 작성자는 <b>익명</b>(형용사+명사)으로 표시되고, <b>같은 서버 동맹원</b>에게는 <b>[서버명 동맹이름] 닉네임</b>으로 보입니다. 모든 서버가 함께 봅니다.</div>
  <div class="unit"><div class="uh"><b>✏ 글쓰기</b><span class="m">하루 10개 · 댓글 50개</span></div>
  <div class="bar"><input class="sq" style="flex:1;min-width:200px" maxlength="30" placeholder="제목" value="${esc(BT)}" oninput="BT=this.value"></div>
  <textarea class="sq" style="width:100%;min-height:80px;box-sizing:border-box;font:inherit" maxlength="1000" placeholder="내용 (최대 1000자, 욕설·비방·개인정보는 삼가 주세요)" oninput="BB=this.value">${esc(BB)}</textarea>
  ${prevTag(BI,'')}<div class="bar"><label class="sbtn" style="cursor:pointer">🖼 이미지 추가<input type="file" accept="image/*" style="display:none" onchange="pickImg(this,'')"></label><button class="sbtn on" onclick="postBd()">올리기</button><span class="m">이미지는 자동으로 줄여서 올라가며 하루 5장까지</span></div></div>
  ${l.map(p=>{const open=!BX.has(p.id);return `<div class="unit"><div class="uh"><span class="nm">${esc(p.title)}</span><span class="m"><b>${esc(p.by)}</b>${sb(p)} · ${ago(p.ts)}</span>
    <div class="votes"><button class="vb u ${p.liked?'on':''}" ${p.mine?'disabled':''} onclick="likeMk('${p.id}')">👍 ${p.likes}</button><button class="vb d ${p.disliked?'on':''}" ${p.mine?'disabled':''} onclick="disMk('${p.id}')">👎 ${p.dis}</button><button class="sbtn" onclick="togBd('${p.id}')">💬 ${p.com.length} ${open?'▴':'▾'}</button>${p.mine?`<button class="vb d" onclick="delBd('${p.id}')">삭제</button>`:''}</div></div>
    ${p.body?`<div class="bdb">${esc(p.body)}</div>`:''}${imgTag(p.img)}
    ${open?`<div class="bcs">${p.com.map(c=>`<div class="bc"><b>${esc(c.by)}</b> <span class="m">${ago(c.ts)}</span>${c.mine?` <a class="m" style="cursor:pointer" onclick="delCom('${c.id}')">삭제</a>`:''}${c.body?`<div>${esc(c.body)}</div>`:''}${imgTag(c.img)}</div>`).join('')||'<div class="m">첫 댓글을 남겨 보세요.</div>'}
      <div class="bar"><input class="sq" style="flex:1;min-width:160px" maxlength="200" placeholder="댓글" value="${esc(BC[p.id]||'')}" oninput="BC['${p.id}']=this.value" onkeydown="if(event.key==='Enter'&&!event.isComposing){comBd('${p.id}')}"><label class="sbtn" style="cursor:pointer">🖼<input type="file" accept="image/*" style="display:none" onchange="pickImg(this,'${p.id}')"></label><button class="sbtn on" onclick="comBd('${p.id}')">등록</button></div>${prevTag(BCI[p.id],p.id)}</div>`:''}</div>`}).join('')||'<div class="m">아직 글이 없습니다. 첫 글을 남겨 보세요!</div>'}`}
function market(){const body=MKT==='req'?reqView():doneView();return `<div class="m" style="margin-bottom:8px">덱을 짜기 어려우면 <b>내 보유 장수·전법을 올려 의뢰</b>하세요. 다른 동맹원이 그 보유로만 5부대를 짜 주고, 의뢰자가 <b>👍·채택</b>하면 짜 준 사람이 <b>🧠 지략가</b> 점수를 얻습니다(모든 서버 공유).</div>${mkTabs()}${body}`}
function reqView(){const l=D.mk.reqs.filter(r=>r.season===D.season);const mine=l.some(r=>r.mine&&!r.done);
  return `<div class="unit"><div class="uh"><b>🙋 내 보유로 편성 의뢰하기</b><span class="m">내 보유 장수 ${D.own.h.length}명 · 전법 ${D.own.s.length}개(돌 포함)가 자동 첨부됩니다 · 시즌${D.season}</span></div>
  <div class="bar"><input class="sq" style="flex:1;min-width:140px" maxlength="24" placeholder="제목 (예: PvP 5부대 짜 주세요)" value="${esc(RT)}" oninput="RT=this.value"><input class="sq" style="flex:2;min-width:180px" maxlength="160" placeholder="요청 사항 (선택: 도트딜 위주, 병기 싫음 등)" value="${esc(RN)}" oninput="RN=this.value"><button class="sbtn on" onclick="postReq()">🙋 의뢰 올리기</button></div>
  <div class="m">보유 목록·돌을 바꾸고 싶으면 <b>장수 도감</b>·<b>전법 도감</b> 탭에서 먼저 수정하세요. 의뢰를 올린 시점의 보유가 첨부됩니다.</div></div>
  ${l.map(reqCard).join('')||'<div class="m">아직 올라온 의뢰가 없습니다.</div>'}`}
function doneView(){let l=D.mk.posts.filter(p=>p.season===D.season);if(MF)l=l.filter(p=>{const[a,b]=mkStat(p);return a===b});l=l.slice().sort((a,b)=>MS==='like'?((b.likes-(b.dis||0))-(a.likes-(a.dis||0))||b.ts-a.ts):(b.ts-a.ts));
  const card=p=>{const[a,b]=mkStat(p);return `<div class="unit mk"><div class="uh"><span class="nm">${esc(p.title)}</span><span class="m">by <b>${esc(p.by)}</b>${sb(p)}</span>
    <div class="votes"><span class="m">내 보유 ${a}/${b}</span><button class="vb u ${p.liked?'on':''}" ${p.mine?'disabled title="내 글"':''} onclick="likeMk('${p.id}')">👍 ${p.likes}</button><button class="vb d ${p.disliked?'on':''}" ${p.mine?'disabled':''} onclick="disMk('${p.id}')">👎 ${p.dis||0}</button><button class="sbtn" onclick="loadMk('${p.id}')">내 편성기로 가져오기</button>${p.mine?`<button class="vb d" onclick="delMk('${p.id}')">삭제</button>`:''}</div></div>
    ${p.note?`<div class="m" style="margin-bottom:6px">${esc(p.note)}</div>`:''}${unitsHtml(p.units,false)}</div>`};
  return `<div class="m" style="margin-bottom:8px">직접 짠 5부대를 공유하는 곳입니다. <b>✓</b>는 내가 보유한 것, <b>✗</b>는 없는 것. 편성기의 <b>완성 편성 공유하기</b>로 올립니다.</div>
  <div class="bar"><button class="sbtn ${MS==='like'?'on':''}" onclick="MS='like';draw()">추천순</button><button class="sbtn ${MS==='new'?'on':''}" onclick="MS='new';draw()">최신순</button><button class="sbtn ${MF?'on':''}" onclick="MF=!MF;draw()">내가 전부 보유한 것만</button><span class="sp"></span><button class="sbtn" onclick="setTab('build')">＋ 내 편성 올리기</button></div>
  ${l.map(card).join('')||'<div class="m">아직 올라온 편성이 없습니다.</div>'}`}
function strat(){const r=D.mk.rank;const md=['🥇','🥈','🥉'];
  return `<div class="m" style="margin-bottom:8px"><b>점수 = 받은 👍 + 채택×${D.mk.adopt_bonus}</b> 순위입니다(의뢰에 짜 준 편성·공유한 편성 모두 포함). <b>TOP3는 🧠 지략가</b>가 되어 마켓에 배지가 붙고, 디스코드 서버에서 <b>지략가 역할</b>을 받습니다(순위가 바뀌면 자동 회수·부여).</div>
  <table><tr><th>#</th><th>유저</th><th>점수</th><th>받은 👍</th><th>채택</th><th>공유 편성</th><th></th></tr>${r.map((x,i)=>`<tr ${x.me?'style="background:var(--card2)"':''}><td>${md[i]||i+1}</td><td><b>${esc(x.by)}</b>${x.me?' <span class="m">(나)</span>':''}</td><td><b>${x.score}</b></td><td>${x.likes}</td><td>${x.adopt}</td><td>${x.posts}</td><td>${x.strat?'<span class="strat">🧠 지략가</span>':''}</td></tr>`).join('')||'<tr><td colspan="7" class="m">아직 순위가 없습니다</td></tr>'}</table>`}
function draw(){
  if(ERR){$('#app').innerHTML=`<div class="err"><h1>추천덱</h1><p>${esc(ERR)}</p></div>`;return}
  const q=s=>`?t=${encodeURIComponent(D.t)}&s=${s}`;
  const tb=(k,l,c)=>`<button class="tab ${c||''} ${TAB===k?'on':''}" onclick="setTab('${k}')">${l}</button>`;
  const NAME={board:'자유게시판',rank:'티어덱',build:'편성기',market:'편성대요청',strat:'지략가',equip:'장비',...FR};
  const main=tb('rank','티어덱')+tb('build','편성기')+tb('market','🙋 편성대요청')+tb('board','💬 자유게시판')+tb('strat','🧠 지략가')+tb('dogam','장수 도감')+tb('skdex','전법 도감');
  const late=tb('equip','장비','late')+tb('chart','상성 차트','late')+tb('mx','상성표','late')+tb('fill','부족한 것 메우기','late')+tb('guide','상황 가이드','late');
  $('#app').innerHTML=`<div class="ttl"><h1>🃏 ${esc([D.ally,D.who].filter(Boolean).join(' '))} · 덱 검색기 · 시즌${D.season}</h1><button class="sbtn on dbtn" onclick="showDon()">☕ 후원</button></div><div class="sub">티어덱을 보고, 내 보유 장수·전법으로 5부대를 짜거나 편성대요청으로 도움을 받으세요 · 한 번 쓴 장수·전법은 다른 부대에서 쓸 수 없음</div>
  <div class="bar"><button class="tab mt" onclick="toggleMenu()">☰ 메뉴 ${MENU?'접기 ▴':'펼치기 ▾'}</button>${MENU?'':`<span class="m">현재: <b>${esc(NAME[TAB]||'')}</b></span>`}<span class="sp"></span>
  <a class="sbtn ${D.season===1?'on':''}" href="${q(1)}">시즌1</a><a class="sbtn ${D.season===2?'on':''}" href="${q(2)}">시즌2</a></div>
  ${MENU?`<div class="bar">${main}<span class="vs"></span><button class="tab lt" onclick="toggleLate()">참고 ${LATE?'접기 ▴':'펼치기 ▾'}</button>${LATE?late:''}</div>`:''}
  ${TAB==='board'?board():TAB==='build'?`<div class="bwrap">${build()}</div>`:TAB==='market'?market():TAB==='strat'?strat():TAB==='rank'?rank():TAB==='equip'?equip():FR[TAB]?frame(TAB):rank()}`}
draw();
</script></body></html>'''


def render(data, err=None):
    """data(dict) 또는 err(문장)로 완성된 HTML 문자열을 만든다."""
    j = json.dumps(data or {"units": [], "rank": {"decks": [], "tacs": [], "syn": []}, "t": "", "season": 1, "who": "", "stone": ""}, ensure_ascii=False).replace("</", "<\\/")
    return TPL.replace("__DATA__", j).replace("__ERR__", json.dumps(err, ensure_ascii=False))
