#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ppa_dashboard_render.py — PPA 6개 표 data → 자기완결 HTML 대시보드.

정산 대시보드(dashboard_render.py)와 같은 패턴: 서버 없이, 데이터를 JSON으로
HTML에 통째로 넣고 JS로 탭을 처리하는 단일 파일.

탭 구성: 홈(요약 KPI+표 관계 구조) · 관계조회(PK로 FK 체인 전체 보기) · 비교
(레코드 여러 건 나란히 비교) · 표별 탭(발전소~수급매칭) · 검증. 각 표 탭에서
PK/FK 값을 누르면 그 레코드 기준 관계조회로 바로 이동하고, 검증 탭의 오류
항목을 누르면 문제의 레코드로 바로 이동합니다.
"""
import datetime
import json


def render_dashboard(data: dict) -> str:
    payload = json.dumps(data, ensure_ascii=False)
    demo = (
        '<div class="demo">⚠ 데모 데이터 — 화면 확인용. 실제 xlsm/CSV로 생성하면 '
        "동일 구조로 채워집니다.</div>"
        if data.get("is_demo")
        else ""
    )
    return (
        TEMPLATE.replace("/*__DATA__*/", "const DATA = " + payload + ";")
        .replace("{{DEMO}}", demo)
        .replace("{{NOW}}", datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))
    )


TEMPLATE = r"""<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>PPA 계약관리 현황</title>
<style>
:root{--paper:#F5F4EF;--panel:#FFF;--ink:#16262B;--sub:#5C6B6E;--line:#E3E1D8;--thead-bg:#EFEEE7;--row-hover:#FAF9F5;
--teal:#0E7C7B;--teal-d:#0A5A59;--teal-w:#E7F1F0;--amber:#B07817;--amber-w:#FAEEDA;
--purple:#534AB7;--purple-w:#EEEDFE;--pass:#1F7A54;--pass-w:#E7F3EC;--fail:#B23A3A;--fail-w:#FBEDEC;--mute:#C9C6BB;}
:root[data-theme="dark"]{--paper:#14191B;--panel:#1C2224;--ink:#EDEFEE;--sub:#8B9A9C;--line:#2C3436;--thead-bg:#242B2D;--row-hover:#222829;
--teal:#28A6A0;--teal-d:#5CC7C1;--teal-w:#0F2E2C;--amber:#E0A94A;--amber-w:#3A2E14;
--purple:#9C93E8;--purple-w:#241F42;--pass:#4CC08A;--pass-w:#123425;--fail:#E2726B;--fail-w:#3A1616;--mute:#4A5254;}
*{box-sizing:border-box}
body{margin:0;background:var(--paper);color:var(--ink);font-family:-apple-system,BlinkMacSystemFont,"Malgun Gothic","Apple SD Gothic Neo",Pretendard,sans-serif;-webkit-font-smoothing:antialiased;padding-bottom:60px}
.mono{font-family:ui-monospace,SFMono-Regular,Consolas,monospace;font-variant-numeric:tabular-nums;letter-spacing:-.02em}
.wrap{max-width:1200px;margin:0 auto;padding:0 22px}
header{border-bottom:2px solid var(--ink);background:var(--paper);position:sticky;top:0;z-index:20}
.mh{max-width:1200px;margin:0 auto;padding:20px 22px 0;display:flex;justify-content:space-between;align-items:flex-start;gap:20px;flex-wrap:wrap}
.eyebrow{font-size:11.5px;letter-spacing:.16em;text-transform:uppercase;color:var(--teal);font-weight:700;margin:0 0 3px}
h1{font-size:23px;margin:0;letter-spacing:-.02em;font-weight:800}.mh .sup{font-size:12.5px;color:var(--sub);margin-top:3px}
#status{font-size:12px;font-weight:700;padding:4px 11px;border-radius:20px;white-space:nowrap}
#status.ok{color:var(--pass);background:var(--pass-w)}#status.no{color:var(--fail);background:var(--fail-w)}
.themebtn{border:1px solid var(--line);background:var(--panel);border-radius:20px;width:32px;height:32px;font-size:14px;cursor:pointer;display:flex;align-items:center;justify-content:center}
.demo{background:var(--amber-w);color:var(--amber);font-size:12px;font-weight:600;padding:7px 22px;text-align:center;border-bottom:1px solid var(--line)}
.gsearchwrap{max-width:1200px;margin:0 auto;padding:12px 22px 0;position:relative}
.globalresults{position:absolute;left:22px;right:22px;top:100%;margin-top:2px;background:var(--panel);border:1px solid var(--line);border-radius:10px;box-shadow:0 8px 24px rgba(0,0,0,.12);max-height:360px;overflow:auto;z-index:30;display:none}
.globalresults.show{display:block}
.gresrow{display:block;width:100%;text-align:left;padding:9px 16px;border:none;background:none;cursor:pointer;font-size:13px;border-bottom:1px solid var(--line);color:var(--ink)}
.gresrow:last-child{border-bottom:none}.gresrow:hover{background:var(--teal-w)}
.gresrow .tag{font-size:10.5px;font-weight:700;color:var(--teal-d);background:var(--teal-w);padding:1px 7px;border-radius:20px;margin-right:8px}
.tabbar{max-width:1200px;margin:0 auto;padding:14px 22px 0;display:flex;gap:4px;flex-wrap:wrap}
.tab{font-size:13.5px;font-weight:600;color:var(--sub);background:none;border:none;cursor:pointer;padding:9px 14px;border-radius:8px 8px 0 0;border-bottom:2.5px solid transparent}
.tab:hover{color:var(--ink)}.tab.on{color:var(--teal-d);border-bottom-color:var(--teal);background:var(--panel)}
.tab.hl{color:var(--purple)}.tab.hl.on{color:var(--purple);border-bottom-color:var(--purple)}
section{margin-top:22px}
#view{animation:fadein .15s ease}
@keyframes fadein{from{opacity:0;transform:translateY(2px)}to{opacity:1;transform:none}}
.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:16px}
.kpi{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:13px 15px;display:flex;flex-direction:column;gap:3px}
.kpi.accent{background:var(--teal);border-color:var(--teal)}.kpi.accent .kk,.kpi.accent .ks{color:#CDE8E6}.kpi.accent .kv{color:#fff}
.kpi.warn{background:var(--fail);border-color:var(--fail)}.kpi.warn .kk,.kpi.warn .ks{color:#F9DEDC}.kpi.warn .kv{color:#fff}
.kpi.clickable{cursor:pointer}
.kk{font-size:11.5px;color:var(--sub);font-weight:600}.kv{font-family:ui-monospace,SFMono-Regular,Consolas,monospace;font-size:21px;font-weight:700;letter-spacing:-.03em}
.ks{font-size:11px;color:var(--sub);font-family:ui-monospace,SFMono-Regular,Consolas,monospace}
.panel{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:16px 18px;margin-bottom:16px}
.ph{display:flex;align-items:baseline;gap:10px;margin-bottom:12px;flex-wrap:wrap}.ph h3{font-size:14.5px;margin:0;font-weight:700}.ph .sub{font-size:12px;color:var(--sub);margin-left:auto}
.toolbar{display:flex;gap:10px;align-items:center;margin-bottom:10px;flex-wrap:wrap}
.search{flex:1;min-width:180px;font-size:13px;padding:8px 12px;border:1px solid var(--line);border-radius:8px;background:var(--paper);color:var(--ink)}
.count{font-size:12px;color:var(--sub);font-family:ui-monospace,SFMono-Regular,Consolas,monospace}
.filterbar{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:8px;align-items:center}
.filtersel{font-size:12.5px;padding:7px 10px;border:1px solid var(--line);border-radius:8px;background:var(--paper);color:var(--ink);max-width:200px}
.clearbtn{font-size:12px;font-weight:700;color:var(--fail);background:var(--fail-w);border:none;border-radius:8px;padding:7px 13px;cursor:pointer;white-space:nowrap}
.chiprow{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:10px}
.fchip{display:inline-flex;align-items:center;gap:6px;font-size:12px;font-weight:600;background:var(--purple-w);color:var(--purple);padding:4px 6px 4px 11px;border-radius:20px}
.fchip button{border:none;background:none;color:inherit;cursor:pointer;font-size:13px;line-height:1;padding:2px}
.datef{display:inline-flex;align-items:center;gap:6px;font-size:12.5px;background:var(--paper);border:1px solid var(--line);border-radius:8px;padding:6px 10px}
.dflabel{color:var(--sub);font-weight:600}
.datef input[type=date]{border:none;background:none;font-size:12.5px;color:var(--ink);font-family:inherit;padding:0}
.dfsep{color:var(--sub)}
.colpicker{position:relative;display:inline-block}
.colpicker summary{list-style:none;cursor:pointer;font-size:12.5px;font-weight:600;color:var(--sub);background:var(--paper);border:1px solid var(--line);border-radius:8px;padding:8px 12px;white-space:nowrap}
.colpicker summary::-webkit-details-marker{display:none}
.colpickbody{position:absolute;z-index:15;margin-top:4px;background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:10px 14px;min-width:200px;max-height:280px;overflow:auto;box-shadow:0 8px 20px rgba(0,0,0,.12)}
.colopt{display:flex;align-items:center;gap:8px;font-size:12.5px;padding:5px 2px;white-space:nowrap;cursor:pointer}
.tbl-wrap{max-height:520px;overflow:auto;border:1px solid var(--line);border-radius:10px}
table{width:100%;border-collapse:collapse;background:var(--panel)}
th,td{padding:9px 13px;font-size:13px;border-bottom:1px solid var(--line);text-align:left;white-space:nowrap}
thead th{position:sticky;top:0;background:var(--thead-bg);font-size:11px;letter-spacing:.02em;color:var(--sub);font-weight:700;text-transform:uppercase;cursor:pointer;z-index:1}
thead th .ar{color:var(--teal);margin-left:3px}
tbody tr:hover{background:var(--row-hover)}tbody tr.rowerr{background:var(--fail-w)}tbody tr.rootrow{background:var(--teal-w)}tbody tr.rootrow td{font-weight:700}
td.cellerr{background:var(--fail-w);color:var(--fail);font-weight:700;border-radius:4px}
td.diffcell{background:var(--amber-w);color:var(--amber);font-weight:700;border-radius:4px}
.unpinbtn{border:none;background:rgba(120,120,120,.25);color:inherit;border-radius:999px;width:16px;height:16px;font-size:10px;line-height:1;cursor:pointer;margin-left:6px}
.pkbadge,.fkbadge{font-size:9.5px;font-weight:700;padding:1px 6px;border-radius:999px;margin-left:6px;vertical-align:middle}
.pkbadge{background:var(--teal);color:#fff}.fkbadge{background:var(--purple-w);color:var(--purple)}
.badge{font-size:11px;font-weight:700;padding:3px 10px;border-radius:20px;white-space:nowrap}
.badge.ok{color:var(--pass);background:var(--pass-w)}.badge.no{color:var(--fail);background:var(--fail-w)}
.boolbadge{font-size:11px;font-weight:700;padding:2px 9px;border-radius:20px;white-space:nowrap}
.boolbadge.yes{background:var(--pass-w);color:var(--pass)}.boolbadge.no{background:var(--fail-w);color:var(--fail)}
.idlink{color:var(--teal-d);text-decoration:underline;text-underline-offset:2px;cursor:pointer;font-weight:700}
.idlink:hover{color:var(--teal)}
.chk{display:grid;grid-template-columns:1fr auto auto auto;gap:12px;align-items:center;background:var(--panel);border:1px solid var(--line);border-left-width:4px;border-radius:8px;padding:10px 15px;margin-bottom:7px;cursor:pointer}
.chk:hover{background:var(--teal-w)}.chk.no{border-left-color:var(--fail)}
.subtabbar{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:2px}
.subtab{font-size:12.5px;font-weight:600;color:var(--sub);background:var(--paper);border:1px solid var(--line);cursor:pointer;padding:8px 14px;border-radius:20px}
.subtab.on{color:#fff;background:var(--teal);border-color:var(--teal)}
.candlist{display:flex;flex-direction:column;gap:5px;margin-top:12px;max-height:340px;overflow:auto}
.candrow{text-align:left;font-size:13px;padding:10px 13px;border:1px solid var(--line);border-radius:8px;background:var(--panel);cursor:pointer;color:var(--ink)}
.candrow:hover{background:var(--teal-w);border-color:var(--teal)}
.nocand{font-size:12.5px;color:var(--sub);padding:10px 2px}
.chipcount{font-size:11.5px;font-weight:700;color:var(--teal-d);background:var(--teal-w);padding:3px 10px;border-radius:20px;margin-left:6px;display:inline-block;margin-top:4px}
.lookuphead{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px}
.unsecrow{display:flex;justify-content:space-between;align-items:center;padding:9px 4px;border-bottom:1px solid var(--line);cursor:pointer}
.unsecrow:last-child{border-bottom:none}.unsecrow:hover{background:var(--paper)}
.mixrow{display:flex;align-items:center;gap:10px;padding:6px 0}
.mixlabel{width:76px;font-size:12.5px;color:var(--sub);flex-shrink:0}
.mixbar{flex:1;height:10px;background:var(--paper);border-radius:6px;overflow:hidden}
.mixfill{height:100%;background:var(--teal);border-radius:6px}
.mixval{font-size:12px;width:96px;text-align:right;flex-shrink:0}
.schemarow{display:flex;align-items:center;gap:8px;flex-wrap:wrap;padding:8px 0}
.schematag{font-size:11px;font-weight:700;color:var(--sub);width:56px;flex-shrink:0}
.schemabox{display:flex;flex-direction:column;align-items:center;gap:2px;font-size:12.5px;font-weight:700;background:var(--paper);border:1px solid var(--line);border-radius:10px;padding:8px 16px;cursor:pointer;color:var(--ink)}
.schemabox:hover{border-color:var(--teal);color:var(--teal-d)}
.schemacount{font-size:10.5px;font-weight:600;color:var(--sub)}
.schemaarrow{color:var(--mute);font-size:16px;flex-shrink:0}
footer{margin-top:30px;padding-top:16px;border-top:1px solid var(--line);font-size:11.5px;color:var(--sub);display:flex;justify-content:space-between;flex-wrap:wrap;gap:6px}
@media(max-width:760px){.kpis{grid-template-columns:1fr 1fr}}
@media print{
  header{position:static}
  .tabbar,.gsearchwrap,.toolbar,.filterbar,.chiprow,.candlist,.subtabbar,.clearbtn,.colpicker,#globalResults,.themebtn{display:none!important}
  .tbl-wrap{max-height:none;overflow:visible;border:none}
  body{padding-bottom:0}
  .panel{break-inside:avoid;border:1px solid #999}
  a.idlink{color:inherit;text-decoration:none}
}
</style></head><body>
<header><div class="mh">
  <div><p class="eyebrow">PPA 계약관리</p><h1>데이터 현황 (조회 전용)</h1>
    <div class="sup">서버 없이 스크립트로 생성된 정적 스냅샷 · 편집은 엑셀에서 진행 후 재생성</div></div>
  <div style="display:flex;align-items:center;gap:8px">
    <button id="themeBtn" class="themebtn" onclick="toggleTheme()" title="화면 테마 전환">🌙</button>
    <span id="status" onclick="state.tab='검증';render()" style="cursor:pointer" title="클릭하면 검증 탭으로 이동">—</span>
  </div>
</div>{{DEMO}}
<div class="gsearchwrap"><input id="globalSearch" class="search" placeholder="전체 표에서 검색 (ID, 이름, 담당자 등)…" oninput="onGlobalSearch(this.value)" onfocus="onGlobalSearch(this.value)">
  <div id="globalResults" class="globalresults"></div>
</div>
<div class="tabbar" id="tabbar"></div></header>
<div class="wrap"><div id="view"></div>
  <footer><span>생성 {{NOW}}</span><span id="foot-src"></span></footer></div>
<script>
/*__DATA__*/
const byKey={};DATA.tables.forEach(t=>byKey[t.key]=t);

// PK 값으로 행을 바로 찾기 위한 색인, 그리고 FK 관계 그래프(엣지) — 관계조회 탭에서 사용.
const rowIndex={};
DATA.tables.forEach(t=>{
  rowIndex[t.key]={};
  t.rows.forEach(r=>{
    const pkv=r.cells[t.pk];
    if(pkv!==undefined&&pkv!=='') rowIndex[t.key][String(pkv)]=r;
  });
});
const edges=[];
DATA.tables.forEach(t=>{
  Object.entries(t.fk||{}).forEach(([col,ref])=>edges.push({table:t.key,col,ref}));
});

function readLS(k,fallback){try{const v=localStorage.getItem(k);return v===null?fallback:v;}catch(e){return fallback;}}
function writeLS(k,v){try{localStorage.setItem(k,v);}catch(e){}}

let state={
  tab:'홈', sort:{}, q:{}, filters:{}, dateFilters:{},
  hidden:{}, colPickerOpen:{},
  lookup:null, lookupTable:DATA.tables[0].key, lookupQ:'',
  recent:[], pinned:[],
  theme: readLS('ppa_theme','light'),
};

function esc(s){return String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function jsq(s){return String(s??'').replace(/\\/g,'\\\\').replace(/'/g,"\\'");}

const NUMCOL=/\(MW\)|\(원\/kWh\)/;
function fmtVal(col,val){
  if(NUMCOL.test(col)&&val!==''){
    const n=Number(val);
    if(!isNaN(n)) return n.toLocaleString('ko-KR',{maximumFractionDigits:4});
  }
  return val;
}
function boolLabel(col,val){
  const bad=val==='TRUE';
  if(col.indexOf('미확보')>=0) return bad?'미확보':'확보완료';
  return bad?'예':'아니오';
}
function cellHtml(t,col,rawVal){
  const val=rawVal??'';
  if(val==='') return '';
  if(val==='TRUE'||val==='FALSE'){
    const bad=val==='TRUE';
    return `<span class="boolbadge ${bad?'no':'yes'}">${esc(boolLabel(col,val))}</span>`;
  }
  const isPk=col===t.pk;
  const refKey=(t.fk||{})[col];
  const display=fmtVal(col,val);
  if(isPk||refKey){
    const target=refKey||t.key;
    return `<a class="idlink" onclick="jumpTo('${jsq(target)}','${jsq(val)}')">${esc(display)}</a>`;
  }
  return esc(display);
}

// ── 컬럼 표시/숨김 (표별로 기억, localStorage에 저장 — 파일을 다시 열어도 유지되지만
// 다른 PC/브라우저에서 열면 초기화되는 "이번 브라우저 한정" 편의 기능입니다) ──
function isColHidden(t,col){return (state.hidden[t.key]||new Set()).has(col);}
function visibleColumns(t){return t.columns.filter(c=>!isColHidden(t,c));}
function toggleCol(tk,col){
  state.hidden[tk]=state.hidden[tk]||new Set();
  if(state.hidden[tk].has(col)) state.hidden[tk].delete(col); else state.hidden[tk].add(col);
  writeLS('ppa_hidden_'+tk, JSON.stringify([...state.hidden[tk]]));
  render();
}
function loadHidden(){
  DATA.tables.forEach(t=>{
    const raw=readLS('ppa_hidden_'+t.key,null);
    if(raw){ try{ state.hidden[t.key]=new Set(JSON.parse(raw)); }catch(e){} }
  });
}
function colPicker(t){
  const isOpen=!!state.colPickerOpen[t.key];
  const opts=t.columns.filter(c=>c!==t.pk).map(c=>{
    const checked=!isColHidden(t,c)?'checked':'';
    return `<label class="colopt"><input type="checkbox" ${checked} onchange="toggleCol('${jsq(t.key)}','${jsq(c)}')">${esc(c)}</label>`;
  }).join('');
  return `<details class="colpicker" ${isOpen?'open':''} ontoggle="state.colPickerOpen['${jsq(t.key)}']=this.open"><summary>컬럼 선택</summary><div class="colpickbody">${opts}</div></details>`;
}

function uniqueValues(t,col){
  const s=new Set();
  t.rows.forEach(r=>{const v=r.cells[col];if(v!==undefined&&v!=='') s.add(String(v));});
  const arr=[...s];
  const allNum=arr.length>0&&arr.every(v=>!isNaN(Number(v)));
  return allNum?arr.sort((a,b)=>Number(a)-Number(b)):arr.sort((a,b)=>a.localeCompare(b,'ko'));
}
function filterableColumns(t){
  return t.columns.filter(c=>{
    if(c===t.pk) return false;
    const n=uniqueValues(t,c).length;
    return n>=2&&n<=20&&n<t.rows.length;
  });
}
function filterBar(t){
  const cols=filterableColumns(t);
  const cur=state.filters[t.key]||{};
  const activeEntries=Object.entries(cur).filter(([,v])=>v);
  const chips=activeEntries.map(([col,val])=>{
    const label=(val==='TRUE'||val==='FALSE')?boolLabel(col,val):val;
    return `<span class="fchip">${esc(col)}: ${esc(label)}<button onclick="setFilter('${jsq(t.key)}','${jsq(col)}','')">✕</button></span>`;
  }).join('');
  const chipRow=chips?`<div class="chiprow">${chips}</div>`:'';
  if(!cols.length) return chipRow;
  const selects=cols.map(c=>{
    const vals=uniqueValues(t,c);
    const sel=cur[c]||'';
    const opts=[`<option value="">${esc(c)}: 전체</option>`].concat(
      vals.map(v=>`<option value="${esc(v)}" ${v===sel?'selected':''}>${esc(v==='TRUE'||v==='FALSE'?boolLabel(c,v):v)}</option>`));
    return `<select class="filtersel" onchange="setFilter('${jsq(t.key)}','${jsq(c)}',this.value)">${opts.join('')}</select>`;
  }).join('');
  const clearBtn=activeEntries.length>1?`<button class="clearbtn" onclick="clearFilters('${jsq(t.key)}')">전체 초기화</button>`:'';
  return `<div class="filterbar">${selects}${clearBtn}</div>${chipRow}`;
}
function setFilter(k,col,val){state.filters[k]=state.filters[k]||{};if(val)state.filters[k][col]=val;else delete state.filters[k][col];render();}
function clearFilters(k){state.filters[k]={};render();}

// ── 날짜 범위 필터: 값이 YYYY-MM-DD 형태인 컬럼을 자동 감지 (공급기한_구매/판매, 계약일 등) ──
function isDateCol(t,col){
  const sample=t.rows.find(r=>r.cells[col]);
  return !!sample && /^\d{4}-\d{2}-\d{2}$/.test(String(sample.cells[col]));
}
function dateFilterBar(t){
  const cols=t.columns.filter(c=>isDateCol(t,c));
  if(!cols.length) return '';
  const df=state.dateFilters[t.key]||{};
  const parts=cols.map(c=>{
    const range=df[c]||{};
    return `<span class="datef"><span class="dflabel">${esc(c)}</span>
      <input type="date" value="${esc(range.from||'')}" onchange="setDateFilter('${jsq(t.key)}','${jsq(c)}','from',this.value)">
      <span class="dfsep">~</span>
      <input type="date" value="${esc(range.to||'')}" onchange="setDateFilter('${jsq(t.key)}','${jsq(c)}','to',this.value)"></span>`;
  }).join('');
  return `<div class="filterbar">${parts}</div>`;
}
function setDateFilter(k,col,edge,val){
  state.dateFilters[k]=state.dateFilters[k]||{};
  state.dateFilters[k][col]=state.dateFilters[k][col]||{};
  state.dateFilters[k][col][edge]=val;
  render();
}
function applyDateFilter(rows,t){
  const df=state.dateFilters[t.key]||{};
  Object.entries(df).forEach(([col,range])=>{
    if(!range.from&&!range.to) return;
    rows=rows.filter(r=>{
      const v=String(r.cells[col]||'');
      if(!v) return false;
      if(range.from&&v<range.from) return false;
      if(range.to&&v>range.to) return false;
      return true;
    });
  });
  return rows;
}

function tableView(t){
  const q=(state.q[t.key]||'').trim().toLowerCase();
  const sc=state.sort[t.key];
  const filt=state.filters[t.key]||{};
  let rows=t.rows.map((r,idx)=>({...r,_idx:idx}));
  if(q) rows=rows.filter(r=>t.columns.some(c=>String(r.cells[c]??'').toLowerCase().includes(q)));
  Object.entries(filt).forEach(([col,val])=>{rows=rows.filter(r=>String(r.cells[col]??'')===val);});
  rows=applyDateFilter(rows,t);
  if(sc){rows=[...rows].sort((a,b)=>{
    const av=a.cells[sc.key]??'',bv=b.cells[sc.key]??'';
    const an=Number(av),bn=Number(bv);
    let cmp;
    if(av!==''&&bv!==''&&!isNaN(an)&&!isNaN(bn)) cmp=an-bn;
    else cmp=String(av).localeCompare(String(bv),'ko');
    return sc.dir*cmp;});}
  const cols=visibleColumns(t);
  const head=cols.map(c=>{
    const ar=sc&&sc.key===c?`<span class="ar">${sc.dir>0?'▲':'▼'}</span>`:'';
    const pk=c===t.pk?'<span class="pkbadge">PK</span>':'';
    const fk=(t.fk||{})[c]?'<span class="fkbadge">FK</span>':'';
    return `<th onclick="doSort('${jsq(t.key)}','${jsq(c)}')">${esc(c)}${pk}${fk}${ar}</th>`;}).join('');
  const body=rows.map(r=>{
    const err=(r.error_cols||[]).length>0;
    const tds=cols.map(c=>{
      const bad=(r.error_cols||[]).includes(c);
      return `<td class="${bad?'cellerr':''}">${cellHtml(t,c,r.cells[c])}</td>`;}).join('');
    return `<tr class="${err?'rowerr':''}">${tds}</tr>`;}).join('');
  return `<div class="toolbar"><input class="search" placeholder="검색…" value="${esc(q)}" oninput="setQ('${jsq(t.key)}',this.value)">
    ${colPicker(t)}
    <span class="count">${rows.length.toLocaleString('ko-KR')} / ${t.rows.length.toLocaleString('ko-KR')}건</span></div>
    ${filterBar(t)}${dateFilterBar(t)}
    <div class="tbl-wrap"><table><thead><tr>${head}</tr></thead><tbody>${body||`<tr><td colspan="${cols.length}" style="text-align:center;padding:24px;color:var(--sub)">데이터 없음</td></tr>`}</tbody></table></div>`;
}
function doSort(k,c){const s=state.sort[k];state.sort[k]=(s&&s.key===c)?{key:c,dir:-s.dir}:{key:c,dir:1};render();}
function setQ(k,v){state.q[k]=v;render();}
const kpi=(k,v,s,cls,onclick)=>`<div class="kpi${cls?' '+cls:''}${onclick?' clickable':''}"${onclick?` onclick="${onclick}"`:''}><span class="kk">${k}</span><span class="kv">${v}</span><span class="ks">${s||''}</span></div>`;
const panel=(t,s,inner)=>`<div class="panel"><div class="ph"><h3>${t}</h3><span class="sub">${s||''}</span></div>${inner}</div>`;

function tData(t){
  const errRows=t.rows.filter(r=>(r.error_cols||[]).length>0).length;
  const hiddenN=(state.hidden[t.key]||new Set()).size;
  return `<section><div class="kpis">
    ${kpi('전체 행 수',t.rows.length.toLocaleString('ko-KR'),t.label,'accent')}
    ${kpi('오류 있는 행',errRows.toLocaleString('ko-KR'),errRows>0?'검증 탭 참고':'문제 없음',errRows>0?'warn':'',errRows>0?"state.tab='검증';render()":'')}
    ${kpi('컬럼 수',t.columns.length,hiddenN?hiddenN+'개 숨김':'')}
    ${kpi('참조 관계',t.fk_columns.length>0?t.fk_columns.join(', '):'없음')}</div>
    ${panel(t.label+' 전체 목록','PK/FK 값 클릭 시 관계조회 · 클릭으로 정렬 · 빨간 셀은 검증 오류',tableView(t))}</section>`;
}

// ── 관계조회: 표를 넘나들지 않고, PK 하나를 고르면 FK로 이어진 다른 표의
// 행들을 전부 한 화면에 모아 보여줍니다 (발전소↔구매계약↔수급매칭↔전기사용지↔판매계약↔수요기업).
function relatedChain(startTable,startPk){
  const visited={};
  const startKey=String(startPk);
  visited[startTable]=new Set([startKey]);
  const queue=[[startTable,startKey]];
  while(queue.length){
    const [tk,pkv]=queue.shift();
    const row=rowIndex[tk]&&rowIndex[tk][pkv];
    if(!row) continue;
    const t=byKey[tk];
    Object.entries(t.fk||{}).forEach(([col,refKey])=>{
      const refPkv=row.cells[col];
      if(refPkv!==undefined&&refPkv!==''){
        const rk=String(refPkv);
        visited[refKey]=visited[refKey]||new Set();
        if(!visited[refKey].has(rk)){visited[refKey].add(rk);queue.push([refKey,rk]);}
      }
    });
    edges.filter(e=>e.ref===tk).forEach(e=>{
      const childT=byKey[e.table];
      childT.rows.forEach(cr=>{
        if(String(cr.cells[e.col]??'')===pkv){
          const childPk=cr.cells[childT.pk];
          if(childPk!==undefined&&childPk!==''){
            const ck=String(childPk);
            visited[e.table]=visited[e.table]||new Set();
            if(!visited[e.table].has(ck)){visited[e.table].add(ck);queue.push([e.table,ck]);}
          }
        }
      });
    });
  }
  return visited;
}
function miniTable(t,rows,highlightPk){
  const cols=visibleColumns(t);
  const head=cols.map(c=>{
    const pk=c===t.pk?'<span class="pkbadge">PK</span>':'';
    const fk=(t.fk||{})[c]?'<span class="fkbadge">FK</span>':'';
    return `<th>${esc(c)}${pk}${fk}</th>`;}).join('');
  const body=rows.map(r=>{
    const isRoot=highlightPk!==null&&String(r.cells[t.pk])===String(highlightPk);
    const tds=cols.map(c=>`<td>${cellHtml(t,c,r.cells[c])}</td>`).join('');
    return `<tr class="${isRoot?'rootrow':''}">${tds}</tr>`;}).join('');
  return `<div class="tbl-wrap"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`;
}
function setLookupTable(k){state.lookupTable=k;state.lookupQ='';render();}
function setLookupQ(v){state.lookupQ=v;render();}
function setLookup(table,pk){
  state.lookup={table,pk};
  const t=byKey[table];
  const row=rowIndex[table]&&rowIndex[table][String(pk)];
  const shortLabel=row?t.columns.slice(0,2).map(c=>row.cells[c]).filter(v=>v!==undefined&&v!=='').join(' · '):String(pk);
  state.recent=state.recent.filter(r=>!(r.table===table&&String(r.pk)===String(pk)));
  state.recent.unshift({table,pk,label:t.label+' · '+shortLabel});
  state.recent=state.recent.slice(0,5);
}
function chooseLookup(table,pk){setLookup(table,pk);state.lookupQ='';render();}
function clearLookup(){state.lookup=null;state.lookupQ='';render();}
function jumpTo(table,pk){state.tab='관계조회';state.lookupTable=table;setLookup(table,pk);render();}

function isPinned(table,pk){return state.pinned.some(p=>p.table===table&&String(p.pk)===String(pk));}
function togglePin(table,pk){
  const idx=state.pinned.findIndex(p=>p.table===table&&String(p.pk)===String(pk));
  if(idx>=0) state.pinned.splice(idx,1);
  else{state.pinned.push({table,pk});if(state.pinned.length>3) state.pinned.shift();}
  render();
}
function recentChips(){
  if(!state.recent.length) return '';
  const chips=state.recent.map(r=>`<button class="candrow" onclick="chooseLookup('${jsq(r.table)}','${jsq(r.pk)}')">${esc(r.label)}</button>`).join('');
  return panel('최근 조회','',`<div class="candlist" style="max-height:none">${chips}</div>`);
}

function tLookup(){
  const activeT=byKey[state.lookupTable];
  const pickTabs=DATA.tables.map(t=>
    `<button class="subtab ${t.key===state.lookupTable?'on':''}" onclick="setLookupTable('${jsq(t.key)}')">${esc(t.label)}</button>`).join('');
  const q=(state.lookupQ||'').trim().toLowerCase();
  const matches=q?activeT.rows.filter(r=>activeT.columns.some(c=>String(r.cells[c]??'').toLowerCase().includes(q))):activeT.rows;
  const shown=matches.slice(0,25);
  const candHtml=shown.map(r=>{
    const pkv=r.cells[activeT.pk];
    const label=activeT.columns.slice(0,3).map(c=>r.cells[c]).filter(v=>v!==undefined&&v!=='').join(' · ');
    return `<button class="candrow" onclick="chooseLookup('${jsq(activeT.key)}','${jsq(pkv)}')">${esc(label)}</button>`;
  }).join('')||'<div class="nocand">일치하는 항목이 없습니다.</div>';
  const moreNote=matches.length>25?`<div class="nocand">${(matches.length-25).toLocaleString('ko-KR')}건 더 있음 — 검색어를 추가해 좁혀보세요.</div>`:'';

  const picker=`<div class="panel"><div class="ph"><h3>1. 표 선택</h3></div>
    <div class="subtabbar">${pickTabs}</div>
    <div class="ph" style="margin-top:16px"><h3>2. 검색해서 선택</h3><span class="sub">비워두면 목록이 그대로 보입니다</span></div>
    <input class="search" style="width:100%" placeholder="${esc(activeT.label)} 검색 (ID, 이름 등)…" value="${esc(state.lookupQ||'')}" oninput="setLookupQ(this.value)">
    <div class="candlist">${candHtml}</div>${moreNote}</div>`;

  if(!state.lookup) return `<section>${recentChips()}${picker}</section>`;

  const rootT=byKey[state.lookup.table];
  const rootRow=rowIndex[state.lookup.table]&&rowIndex[state.lookup.table][String(state.lookup.pk)];
  if(!rootRow){
    return `<section><div class="panel lookuphead">
        <div><div class="ph" style="margin-bottom:4px"><h3>조회 중: ${esc(rootT.label)} · ${esc(state.lookup.pk)}</h3></div>
        <span class="badge no">이 ID를 가진 실제 레코드가 없습니다 — 오타나 참조 오류일 수 있습니다 (검증 탭 확인).</span></div>
        <button class="clearbtn" onclick="clearLookup()">다른 항목 조회</button>
      </div></section>`;
  }
  const chain=relatedChain(state.lookup.table,state.lookup.pk);
  const rootLabel=rootT.columns.slice(0,3).map(c=>rootRow.cells[c]).filter(v=>v!==undefined&&v!=='').join(' · ');
  const chainRows=DATA.tables.map(t=>{
    const pkset=chain[t.key];
    const rows=pkset&&pkset.size>0?t.rows.filter(r=>pkset.has(String(r.cells[t.pk]))):[];
    return {t,rows};
  });
  const summary=chainRows.map(({t,rows})=>
    rows.length>0?`<span class="chipcount">${esc(t.label)} ${rows.length}</span>`:'').join('');
  const panels=chainRows.map(({t,rows})=>{
    if(!rows.length) return '';
    const highlight=state.lookup.table===t.key?state.lookup.pk:null;
    return panel(`${esc(t.label)} (${rows.length}건 연결)`,'',miniTable(t,rows,highlight));
  }).join('');
  const pinned=isPinned(state.lookup.table,state.lookup.pk);

  return `<section><div class="panel lookuphead">
      <div><div class="ph" style="margin-bottom:4px"><h3>조회 중: ${esc(rootT.label)} · ${esc(rootLabel)}</h3></div>${summary}</div>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <button class="clearbtn" style="background:var(--teal-w);color:var(--teal-d)" onclick="togglePin('${jsq(state.lookup.table)}','${jsq(state.lookup.pk)}')">${pinned?'비교에서 제거':'비교에 추가 ('+state.pinned.length+'/3)'}</button>
        <button class="clearbtn" onclick="clearLookup()">다른 항목 조회</button>
      </div>
    </div>${panels}</section>`;
}

function tCompare(){
  if(!state.pinned.length){
    return `<section><div class="panel"><div class="ph"><h3>비교할 레코드가 없습니다</h3></div>
      <div class="nocand">관계조회에서 레코드를 연 뒤 "비교에 추가"를 누르면 여기서 최대 3건까지 나란히 비교할 수 있습니다.</div></div></section>`;
  }
  const groups={};
  state.pinned.forEach(p=>{groups[p.table]=groups[p.table]||[];groups[p.table].push(p.pk);});
  const panels=Object.entries(groups).map(([tk,pks])=>{
    const t=byKey[tk];
    const rows=pks.map(pk=>rowIndex[tk]&&rowIndex[tk][String(pk)]).filter(Boolean);
    if(!rows.length) return '';
    const cols=visibleColumns(t);
    const head=`<th>항목</th>`+rows.map(r=>
      `<th>${esc(r.cells[t.pk])}<button class="unpinbtn" onclick="togglePin('${jsq(tk)}','${jsq(r.cells[t.pk])}')">✕</button></th>`).join('');
    const body=cols.map(c=>{
      const vals=rows.map(r=>r.cells[c]);
      const differ=new Set(vals.map(v=>String(v??''))).size>1;
      const tds=rows.map(r=>`<td class="${differ?'diffcell':''}">${cellHtml(t,c,r.cells[c])}</td>`).join('');
      const pk=c===t.pk?'<span class="pkbadge">PK</span>':'';
      const fk=(t.fk||{})[c]?'<span class="fkbadge">FK</span>':'';
      return `<tr><td>${esc(c)}${pk}${fk}</td>${tds}</tr>`;
    }).join('');
    return panel(`${esc(t.label)} 비교 (${rows.length}건)`,'다른 값은 강조 표시됩니다',
      `<div class="tbl-wrap"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`);
  }).join('');
  return `<section><div class="panel lookuphead"><div class="ph" style="margin-bottom:0"><h3>비교 중 (${state.pinned.length}건)</h3></div>
    <button class="clearbtn" onclick="state.pinned=[];render()">전체 해제</button></div>${panels}</section>`;
}

function sumCol(tableKey,col){
  const t=byKey[tableKey];if(!t) return 0;
  return t.rows.reduce((s,r)=>{const n=Number(r.cells[col]);return s+(isNaN(n)?0:n);},0);
}
function countWhere(tableKey,col,val){
  const t=byKey[tableKey];if(!t) return 0;
  return t.rows.filter(r=>String(r.cells[col]??'')===val).length;
}
function groupSum(tableKey,groupCol,sumColName){
  const t=byKey[tableKey];if(!t) return [];
  const m={};
  t.rows.forEach(r=>{
    const g=r.cells[groupCol]||'(미지정)';
    const n=Number(r.cells[sumColName]);
    m[g]=(m[g]||0)+(isNaN(n)?0:n);
  });
  return Object.entries(m).sort((a,b)=>b[1]-a[1]);
}
function jumpToFilter(tableKey,col,val){
  state.tab=tableKey;
  state.filters[tableKey]=state.filters[tableKey]||{};
  state.filters[tableKey][col]=val;
  render();
}
function schemaBox(tableKey){
  const t=byKey[tableKey];
  return `<button class="schemabox" onclick="state.tab='${jsq(tableKey)}';render()">${esc(t.label)}<span class="schemacount mono">${t.rows.length}</span></button>`;
}
function schemaDiagram(){
  const has=k=>!!byKey[k];
  if(!has('T_발전소')||!has('T_구매계약')||!has('T_수급매칭')||!has('T_수요기업')||!has('T_판매계약')||!has('T_전기사용지')) return '';
  return `<div class="schemarow"><span class="schematag">공급측</span>${schemaBox('T_발전소')}<span class="schemaarrow">→</span>${schemaBox('T_구매계약')}<span class="schemaarrow">→</span>${schemaBox('T_수급매칭')}</div>
    <div class="schemarow"><span class="schematag">수요측</span>${schemaBox('T_수요기업')}<span class="schemaarrow">→</span>${schemaBox('T_판매계약')}<span class="schemaarrow">→</span>${schemaBox('T_전기사용지')}<span class="schemaarrow">→</span>${schemaBox('T_수급매칭')}</div>`;
}
function tHome(){
  const hasPlant=byKey['T_발전소'],hasPurch=byKey['T_구매계약'],hasSale=byKey['T_판매계약'];
  const supplyMW=hasPlant?sumCol('T_발전소','설비용량(MW)'):0;
  const purchMW=hasPurch?sumCol('T_구매계약','구매계약용량(MW)'):0;
  const saleMW=hasSale?sumCol('T_판매계약','판매계약용량(MW)'):0;
  const purchUnsecured=hasPurch?countWhere('T_구매계약','수요기업 미확보','TRUE'):0;
  const saleUnsecured=hasSale?countWhere('T_판매계약','공급자원 미확보','TRUE'):0;
  const mix=hasPlant?groupSum('T_발전소','발전원','설비용량(MW)'):[];
  const mixMax=mix.length?mix[0][1]:0;
  const mixBars=mix.map(([g,v])=>`<div class="mixrow"><span class="mixlabel">${esc(g)}</span>
      <div class="mixbar"><div class="mixfill" style="width:${mixMax>0?(v/mixMax*100):0}%"></div></div>
      <span class="mixval mono">${v.toLocaleString('ko-KR',{maximumFractionDigits:1})} MW</span></div>`).join('')||'<div class="nocand">데이터 없음</div>';
  const ok=DATA.validation.total_errors===0;
  const schema=schemaDiagram();

  return `<section><div class="kpis">
    ${kpi('발전소 설비용량 합계',supplyMW.toLocaleString('ko-KR',{maximumFractionDigits:1})+' MW',hasPlant?hasPlant.rows.length+'개 발전소':'','accent')}
    ${kpi('구매계약 총 용량',purchMW.toLocaleString('ko-KR',{maximumFractionDigits:1})+' MW',hasPurch?hasPurch.rows.length+'건':'')}
    ${kpi('판매계약 총 용량',saleMW.toLocaleString('ko-KR',{maximumFractionDigits:1})+' MW',hasSale?hasSale.rows.length+'건':'')}
    ${kpi('검증 오류',DATA.validation.total_errors,ok?'전 표 정상':'클릭해서 확인',ok?'':'warn',"state.tab='검증';render()")}</div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">
      ${panel('발전원별 설비용량 비중','',mixBars)}
      ${panel('미확보 현황','클릭하면 해당 표로 이동 + 필터 적용',`
        <div class="unsecrow" onclick="jumpToFilter('T_구매계약','수요기업 미확보','TRUE')"><span>구매계약 — 수요기업 미확보</span><span class="badge ${purchUnsecured>0?'no':'ok'}">${purchUnsecured}건</span></div>
        <div class="unsecrow" onclick="jumpToFilter('T_판매계약','공급자원 미확보','TRUE')"><span>판매계약 — 공급자원 미확보</span><span class="badge ${saleUnsecured>0?'no':'ok'}">${saleUnsecured}건</span></div>`)}
    </div>
    ${schema?panel('표 관계 구조','발전소부터 수요기업까지 이어지는 참조 관계 — 박스를 누르면 해당 탭으로 이동합니다',schema):''}
    </section>`;
}

function tVerify(){
  const v=DATA.validation;
  const byTable=Object.entries(v.by_table).map(([k,c])=>
    `<tr onclick="state.tab='${jsq(k)}';render()" style="cursor:pointer"><td>${esc(byKey[k]?byKey[k].label:k)}</td><td class="mono">${c}</td></tr>`
  ).join('')||'<tr><td colspan="2" style="text-align:center;color:var(--sub)">오류 없음</td></tr>';
  const byItem=Object.entries(v.by_item).map(([k,c])=>`<tr><td>${esc(k)}</td><td class="mono">${c}</td></tr>`).join('')||'<tr><td colspan="2" style="text-align:center;color:var(--sub)">오류 없음</td></tr>';
  const detail=v.errors.map(e=>{
    const pkv=e.pk_value;
    const onclickAttr=pkv?`jumpTo('${jsq(e.table)}','${jsq(pkv)}')`:`state.tab='${jsq(e.table)}';render()`;
    return `<div class="chk no" onclick="${onclickAttr}">
      <span>${esc(byKey[e.table]?byKey[e.table].label:e.table)} · PK=${esc(pkv||'(공란)')}</span>
      <span class="mono">행 ${e.row_index+1}</span>
      <span>${esc(e.error_item)}</span>
      <span class="badge no">오류 · 클릭해서 확인</span></div>`;
  }).join('')||'<div style="text-align:center;color:var(--sub);padding:20px">오류가 없습니다.</div>';
  return `<section><div class="kpis">
    ${kpi('총 오류 건수',v.total_errors,v.total_errors>0?'표별 세부는 아래':'전 표 정상',v.total_errors>0?'warn':'accent')}</div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">
      ${panel('표별 오류 건수','',`<table><thead><tr><th>표</th><th>건수</th></tr></thead><tbody>${byTable}</tbody></table>`)}
      ${panel('오류항목별 건수','',`<table><thead><tr><th>오류항목</th><th>건수</th></tr></thead><tbody>${byItem}</tbody></table>`)}
    </div>
    ${panel('상세 오류 목록','행을 클릭하면 문제의 레코드로 이동합니다',detail)}</section>`;
}

function onGlobalSearch(v){
  const box=document.getElementById('globalResults');
  const q=(v||'').trim().toLowerCase();
  if(!q){box.classList.remove('show');box.innerHTML='';return;}
  const results=[];
  DATA.tables.forEach(t=>{
    t.rows.forEach(r=>{
      if(t.columns.some(c=>String(r.cells[c]??'').toLowerCase().includes(q))){
        const label=t.columns.slice(0,3).map(c=>r.cells[c]).filter(x=>x!==undefined&&x!=='').join(' · ');
        results.push({table:t.key,tlabel:t.label,pk:r.cells[t.pk],text:label});
      }
    });
  });
  const shown=results.slice(0,20);
  box.innerHTML = shown.length
    ? shown.map(r=>`<button class="gresrow" onclick="closeGlobalSearch();jumpTo('${jsq(r.table)}','${jsq(r.pk)}')"><span class="tag">${esc(r.tlabel)}</span>${esc(r.text)}</button>`).join('') +
      (results.length>20?`<div class="nocand" style="padding:8px 16px">${(results.length-20).toLocaleString('ko-KR')}건 더 있음 — 검색어를 좁혀보세요.</div>`:'')
    : '<div class="nocand" style="padding:10px 16px">일치하는 항목이 없습니다.</div>';
  box.classList.add('show');
}
function closeGlobalSearch(){
  const inp=document.getElementById('globalSearch');if(inp) inp.value='';
  const box=document.getElementById('globalResults');if(box){box.classList.remove('show');box.innerHTML='';}
}
document.addEventListener('click',e=>{
  const wrap=document.querySelector('.gsearchwrap');
  if(wrap&&!wrap.contains(e.target)) closeGlobalSearch();
});

function applyTheme(){
  document.documentElement.setAttribute('data-theme',state.theme);
  const btn=document.getElementById('themeBtn');
  if(btn) btn.textContent=state.theme==='dark'?'☀️':'🌙';
}
function toggleTheme(){
  state.theme=state.theme==='dark'?'light':'dark';
  writeLS('ppa_theme',state.theme);
  applyTheme();
}

function syncHash(){
  if(state.tab==='관계조회'&&state.lookup){
    const h='#lookup='+encodeURIComponent(state.lookup.table)+':'+encodeURIComponent(state.lookup.pk);
    if(location.hash!==h) history.replaceState(null,'',h);
  }else if(location.hash){
    history.replaceState(null,'',location.pathname+location.search);
  }
}
function parseHash(){
  const m=location.hash.match(/^#lookup=([^:]+):(.+)$/);
  if(m){
    const table=decodeURIComponent(m[1]),pk=decodeURIComponent(m[2]);
    if(byKey[table]){state.tab='관계조회';state.lookupTable=table;setLookup(table,pk);}
  }
}

function renderTabs(){
  const tabs=[['홈','홈'],['관계조회','관계조회']];
  if(state.pinned.length) tabs.push(['비교','비교 ('+state.pinned.length+')']);
  DATA.tables.forEach(t=>tabs.push([t.key,t.label]));
  tabs.push(['검증','검증']);
  document.getElementById('tabbar').innerHTML=tabs.map(([k,l])=>
    `<button class="tab${(k==='홈'||k==='관계조회'||k==='비교')?' hl':''}${k===state.tab?' on':''}" data-k="${esc(k)}" onclick="state.tab='${jsq(k)}';render()">${esc(l)}</button>`).join('');
}

function render(){
  renderTabs();
  const view=document.getElementById('view');
  let html;
  if(state.tab==='홈') html=tHome();
  else if(state.tab==='관계조회') html=tLookup();
  else if(state.tab==='비교') html=tCompare();
  else if(state.tab==='검증') html=tVerify();
  else html=tData(byKey[state.tab]);
  view.innerHTML=html;
  view.classList.remove('fadein');void view.offsetWidth;view.classList.add('fadein');
  const st=document.getElementById('status');
  const ok=DATA.validation.total_errors===0;
  st.className=ok?'ok':'no';st.textContent=ok?'전 표 검증 통과':`검증 오류 ${DATA.validation.total_errors}건`;
  syncHash();
}
(function(){
  loadHidden();
  applyTheme();
  parseHash();
  document.getElementById('foot-src').textContent=(DATA.is_demo?'데모 데이터':'실 데이터') +
    ' · ' + DATA.tables.map(t=>t.label+' '+t.rows.length).join(' · ');
  render();
})();
</script></body></html>"""
