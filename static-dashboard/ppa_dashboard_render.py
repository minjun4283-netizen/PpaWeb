#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ppa_dashboard_render.py — PPA 6개 표 data → 자기완결 HTML 대시보드.

정산 대시보드(dashboard_render.py)와 같은 패턴: 서버 없이, 데이터를 JSON으로
HTML에 통째로 넣고 JS로 탭/검색/정렬을 처리하는 단일 파일. 표별 탭 + 관계조회
탭 + 검증 탭. 검증 탭에서 잡힌 오류는 원본 웹앱처럼 그 표의 정확한 셀을
빨갛게 표시합니다.

관계조회 탭: 표를 시트별로 넘겨보는 대신, PK 하나를 고르면 그 값과 FK로
연결된 다른 표의 행들(발전소→구매계약→수급매칭→전기사용지→판매계약→
수요기업)을 한 화면에서 체인으로 보여줍니다. 각 표 탭에서도 PK/FK 값을
누르면 바로 그 레코드를 기준으로 관계조회 탭으로 이동합니다.
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
:root{--paper:#F5F4EF;--panel:#FFF;--ink:#16262B;--sub:#5C6B6E;--line:#E3E1D8;
--teal:#0E7C7B;--teal-d:#0A5A59;--teal-w:#E7F1F0;--amber:#B07817;--amber-w:#FAEEDA;
--purple:#534AB7;--purple-w:#EEEDFE;--pass:#1F7A54;--pass-w:#E7F3EC;--fail:#B23A3A;--fail-w:#FBEDEC;--mute:#C9C6BB;}
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
.demo{background:var(--amber-w);color:var(--amber);font-size:12px;font-weight:600;padding:7px 22px;text-align:center;border-bottom:1px solid var(--line)}
.tabbar{max-width:1200px;margin:0 auto;padding:14px 22px 0;display:flex;gap:4px;flex-wrap:wrap}
.tab{font-size:13.5px;font-weight:600;color:var(--sub);background:none;border:none;cursor:pointer;padding:9px 14px;border-radius:8px 8px 0 0;border-bottom:2.5px solid transparent}
.tab:hover{color:var(--ink)}.tab.on{color:var(--teal-d);border-bottom-color:var(--teal);background:var(--panel)}
.tab.hl{color:var(--purple)}.tab.hl.on{color:var(--purple);border-bottom-color:var(--purple)}
section{margin-top:22px}
.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:16px}
.kpi{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:13px 15px;display:flex;flex-direction:column;gap:3px}
.kpi.accent{background:var(--teal);border-color:var(--teal)}.kpi.accent .kk,.kpi.accent .ks{color:#CDE8E6}.kpi.accent .kv{color:#fff}
.kpi.warn{background:var(--fail);border-color:var(--fail)}.kpi.warn .kk,.kpi.warn .ks{color:#F9DEDC}.kpi.warn .kv{color:#fff}
.kk{font-size:11.5px;color:var(--sub);font-weight:600}.kv{font-family:ui-monospace,SFMono-Regular,Consolas,monospace;font-size:21px;font-weight:700;letter-spacing:-.03em}
.ks{font-size:11px;color:var(--sub);font-family:ui-monospace,SFMono-Regular,Consolas,monospace}
.panel{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:16px 18px;margin-bottom:16px}
.ph{display:flex;align-items:baseline;gap:10px;margin-bottom:12px;flex-wrap:wrap}.ph h3{font-size:14.5px;margin:0;font-weight:700}.ph .sub{font-size:12px;color:var(--sub);margin-left:auto}
.toolbar{display:flex;gap:10px;align-items:center;margin-bottom:10px;flex-wrap:wrap}
.search{flex:1;min-width:180px;font-size:13px;padding:8px 12px;border:1px solid var(--line);border-radius:8px;background:var(--paper);color:var(--ink)}
.count{font-size:12px;color:var(--sub);font-family:ui-monospace,SFMono-Regular,Consolas,monospace}
.filterbar{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px}
.filtersel{font-size:12.5px;padding:7px 10px;border:1px solid var(--line);border-radius:8px;background:var(--paper);color:var(--ink);max-width:200px}
.clearbtn{font-size:12px;font-weight:700;color:var(--fail);background:var(--fail-w);border:none;border-radius:8px;padding:7px 13px;cursor:pointer;white-space:nowrap}
.tbl-wrap{max-height:520px;overflow:auto;border:1px solid var(--line);border-radius:10px}
table{width:100%;border-collapse:collapse;background:var(--panel)}
th,td{padding:9px 13px;font-size:13px;border-bottom:1px solid var(--line);text-align:left;white-space:nowrap}
thead th{position:sticky;top:0;background:#EFEEE7;font-size:11px;letter-spacing:.02em;color:var(--sub);font-weight:700;text-transform:uppercase;cursor:pointer;z-index:1}
thead th .ar{color:var(--teal);margin-left:3px}
tbody tr:hover{background:#FAF9F5}tbody tr.rowerr{background:var(--fail-w)}tbody tr.rootrow{background:var(--teal-w)}tbody tr.rootrow td{font-weight:700}
td.cellerr{background:var(--fail-w);color:var(--fail);font-weight:700;border-radius:4px}
.pkbadge,.fkbadge{font-size:9.5px;font-weight:700;padding:1px 6px;border-radius:999px;margin-left:6px;vertical-align:middle}
.pkbadge{background:var(--teal);color:#fff}.fkbadge{background:var(--purple-w);color:var(--purple)}
.badge{font-size:11px;font-weight:700;padding:3px 10px;border-radius:20px;white-space:nowrap}
.badge.ok{color:var(--pass);background:var(--pass-w)}.badge.no{color:var(--fail);background:var(--fail-w)}
.boolbadge{font-size:11px;font-weight:700;padding:2px 9px;border-radius:20px;white-space:nowrap}
.boolbadge.yes{background:var(--pass-w);color:var(--pass)}.boolbadge.no{background:var(--fail-w);color:var(--fail)}
.idlink{color:var(--teal-d);text-decoration:underline;text-underline-offset:2px;cursor:pointer;font-weight:700}
.idlink:hover{color:var(--teal)}
.chk{display:grid;grid-template-columns:1fr auto auto auto;gap:12px;align-items:center;background:var(--panel);border:1px solid var(--line);border-left-width:4px;border-radius:8px;padding:10px 15px;margin-bottom:7px}
.chk.no{border-left-color:var(--fail)}
.subtabbar{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:2px}
.subtab{font-size:12.5px;font-weight:600;color:var(--sub);background:var(--paper);border:1px solid var(--line);cursor:pointer;padding:8px 14px;border-radius:20px}
.subtab.on{color:#fff;background:var(--teal);border-color:var(--teal)}
.candlist{display:flex;flex-direction:column;gap:5px;margin-top:12px;max-height:340px;overflow:auto}
.candrow{text-align:left;font-size:13px;padding:10px 13px;border:1px solid var(--line);border-radius:8px;background:var(--panel);cursor:pointer}
.candrow:hover{background:var(--teal-w);border-color:var(--teal)}
.nocand{font-size:12.5px;color:var(--sub);padding:10px 2px}
.chipcount{font-size:11.5px;font-weight:700;color:var(--teal-d);background:var(--teal-w);padding:3px 10px;border-radius:20px;margin-left:6px;display:inline-block;margin-top:4px}
.lookuphead{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px}
footer{margin-top:30px;padding-top:16px;border-top:1px solid var(--line);font-size:11.5px;color:var(--sub);display:flex;justify-content:space-between;flex-wrap:wrap;gap:6px}
@media(max-width:760px){.kpis{grid-template-columns:1fr 1fr}}
</style></head><body>
<header><div class="mh">
  <div><p class="eyebrow">PPA 계약관리</p><h1>데이터 현황 (조회 전용)</h1>
    <div class="sup">서버 없이 스크립트로 생성된 정적 스냅샷 · 편집은 엑셀에서 진행 후 재생성</div></div>
  <span id="status">—</span>
</div>{{DEMO}}<div class="tabbar" id="tabbar"></div></header>
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

let state={tab:DATA.tables[0].key,sort:{},q:{},filters:{},
  lookup:null,lookupTable:DATA.tables[0].key,lookupQ:''};

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
  if(!cols.length) return '';
  const cur=state.filters[t.key]||{};
  const selects=cols.map(c=>{
    const vals=uniqueValues(t,c);
    const sel=cur[c]||'';
    const opts=[`<option value="">${esc(c)}: 전체</option>`].concat(
      vals.map(v=>`<option value="${esc(v)}" ${v===sel?'selected':''}>${esc(v==='TRUE'||v==='FALSE'?boolLabel(c,v):v)}</option>`));
    return `<select class="filtersel" onchange="setFilter('${jsq(t.key)}','${jsq(c)}',this.value)">${opts.join('')}</select>`;
  }).join('');
  const activeCount=Object.values(cur).filter(v=>v).length;
  const clearBtn=activeCount?`<button class="clearbtn" onclick="clearFilters('${jsq(t.key)}')">필터 초기화 (${activeCount})</button>`:'';
  return `<div class="filterbar">${selects}${clearBtn}</div>`;
}
function setFilter(k,col,val){state.filters[k]=state.filters[k]||{};if(val)state.filters[k][col]=val;else delete state.filters[k][col];render();}
function clearFilters(k){state.filters[k]={};render();}

function tableView(t){
  const q=(state.q[t.key]||'').trim().toLowerCase();
  const sc=state.sort[t.key];
  const filt=state.filters[t.key]||{};
  let rows=t.rows.map((r,idx)=>({...r,_idx:idx}));
  if(q) rows=rows.filter(r=>t.columns.some(c=>String(r.cells[c]??'').toLowerCase().includes(q)));
  Object.entries(filt).forEach(([col,val])=>{rows=rows.filter(r=>String(r.cells[col]??'')===val);});
  if(sc){rows=[...rows].sort((a,b)=>{
    const av=a.cells[sc.key]??'',bv=b.cells[sc.key]??'';
    const an=Number(av),bn=Number(bv);
    let cmp;
    if(av!==''&&bv!==''&&!isNaN(an)&&!isNaN(bn)) cmp=an-bn;
    else cmp=String(av).localeCompare(String(bv),'ko');
    return sc.dir*cmp;});}
  const head=t.columns.map(c=>{
    const ar=sc&&sc.key===c?`<span class="ar">${sc.dir>0?'▲':'▼'}</span>`:'';
    const pk=c===t.pk?'<span class="pkbadge">PK</span>':'';
    const fk=(t.fk||{})[c]?'<span class="fkbadge">FK</span>':'';
    return `<th onclick="doSort('${jsq(t.key)}','${jsq(c)}')">${esc(c)}${pk}${fk}${ar}</th>`;}).join('');
  const body=rows.map(r=>{
    const err=(r.error_cols||[]).length>0;
    const tds=t.columns.map(c=>{
      const bad=(r.error_cols||[]).includes(c);
      return `<td class="${bad?'cellerr':''}">${cellHtml(t,c,r.cells[c])}</td>`;}).join('');
    return `<tr class="${err?'rowerr':''}">${tds}</tr>`;}).join('');
  return `<div class="toolbar"><input class="search" placeholder="검색…" value="${esc(q)}" oninput="setQ('${jsq(t.key)}',this.value)">
    <span class="count">${rows.length.toLocaleString('ko-KR')} / ${t.rows.length.toLocaleString('ko-KR')}건</span></div>
    ${filterBar(t)}
    <div class="tbl-wrap"><table><thead><tr>${head}</tr></thead><tbody>${body||'<tr><td style="text-align:center;padding:24px;color:var(--sub)">데이터 없음</td></tr>'}</tbody></table></div>`;
}
function doSort(k,c){const s=state.sort[k];state.sort[k]=(s&&s.key===c)?{key:c,dir:-s.dir}:{key:c,dir:1};render();}
function setQ(k,v){state.q[k]=v;render();}
const kpi=(k,v,s,cls)=>`<div class="kpi${cls?' '+cls:''}"><span class="kk">${k}</span><span class="kv">${v}</span><span class="ks">${s||''}</span></div>`;
const panel=(t,s,inner)=>`<div class="panel"><div class="ph"><h3>${t}</h3><span class="sub">${s||''}</span></div>${inner}</div>`;

function tData(t){
  const errRows=t.rows.filter(r=>(r.error_cols||[]).length>0).length;
  return `<section><div class="kpis">
    ${kpi('전체 행 수',t.rows.length.toLocaleString('ko-KR'),t.label,true?'accent':'')}
    ${kpi('오류 있는 행',errRows.toLocaleString('ko-KR'),errRows>0?'검증 탭 참고':'문제 없음',errRows>0?'warn':'')}
    ${kpi('컬럼 수',t.columns.length)}
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
  const head=t.columns.map(c=>{
    const pk=c===t.pk?'<span class="pkbadge">PK</span>':'';
    const fk=(t.fk||{})[c]?'<span class="fkbadge">FK</span>':'';
    return `<th>${esc(c)}${pk}${fk}</th>`;}).join('');
  const body=rows.map(r=>{
    const isRoot=highlightPk!==null&&String(r.cells[t.pk])===String(highlightPk);
    const tds=t.columns.map(c=>`<td>${cellHtml(t,c,r.cells[c])}</td>`).join('');
    return `<tr class="${isRoot?'rootrow':''}">${tds}</tr>`;}).join('');
  return `<div class="tbl-wrap"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`;
}
function setLookupTable(k){state.lookupTable=k;state.lookupQ='';render();}
function setLookupQ(v){state.lookupQ=v;render();}
function chooseLookup(table,pk){state.lookup={table,pk};state.lookupQ='';render();}
function clearLookup(){state.lookup=null;state.lookupQ='';render();}
function jumpTo(table,pk){state.tab='관계조회';state.lookupTable=table;state.lookup={table,pk};render();}

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

  if(!state.lookup) return `<section>${picker}</section>`;

  const chain=relatedChain(state.lookup.table,state.lookup.pk);
  const rootT=byKey[state.lookup.table];
  const rootRow=rowIndex[state.lookup.table]&&rowIndex[state.lookup.table][String(state.lookup.pk)];
  const rootLabel=rootRow?rootT.columns.slice(0,3).map(c=>rootRow.cells[c]).filter(v=>v!==undefined&&v!=='').join(' · '):state.lookup.pk;
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

  return `<section><div class="panel lookuphead">
      <div><div class="ph" style="margin-bottom:4px"><h3>조회 중: ${esc(rootT.label)} · ${esc(rootLabel)}</h3></div>${summary}</div>
      <button class="clearbtn" onclick="clearLookup()">다른 항목 조회</button>
    </div>${panels}</section>`;
}

function tVerify(){
  const v=DATA.validation;
  const byTable=Object.entries(v.by_table).map(([k,c])=>`<tr><td>${byKey[k]?byKey[k].label:k}</td><td class="mono">${c}</td></tr>`).join('')||'<tr><td colspan="2" style="text-align:center;color:var(--sub)">오류 없음</td></tr>';
  const byItem=Object.entries(v.by_item).map(([k,c])=>`<tr><td>${k}</td><td class="mono">${c}</td></tr>`).join('')||'<tr><td colspan="2" style="text-align:center;color:var(--sub)">오류 없음</td></tr>';
  const detail=v.errors.map(e=>`<div class="chk no">
      <span>${byKey[e.table]?byKey[e.table].label:e.table} · PK=${e.pk_value||'(공란)'}</span>
      <span class="mono">행 ${e.row_index+1}</span>
      <span>${e.error_item}</span>
      <span class="badge no">오류</span></div>`).join('')||'<div style="text-align:center;color:var(--sub);padding:20px">오류가 없습니다.</div>';
  return `<section><div class="kpis">
    ${kpi('총 오류 건수',v.total_errors,v.total_errors>0?'표별 세부는 아래':'전 표 정상',v.total_errors>0?'warn':'accent')}</div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">
      ${panel('표별 오류 건수','',`<table><thead><tr><th>표</th><th>건수</th></tr></thead><tbody>${byTable}</tbody></table>`)}
      ${panel('오류항목별 건수','',`<table><thead><tr><th>오류항목</th><th>건수</th></tr></thead><tbody>${byItem}</tbody></table>`)}
    </div>
    ${panel('상세 오류 목록','',detail)}</section>`;
}

function render(){
  const view=document.getElementById('view');
  view.innerHTML = state.tab==='검증' ? tVerify() : (state.tab==='관계조회' ? tLookup() : tData(byKey[state.tab]));
  const st=document.getElementById('status');
  const ok=DATA.validation.total_errors===0;
  st.className=ok?'ok':'no';st.textContent=ok?'전 표 검증 통과':`검증 오류 ${DATA.validation.total_errors}건`;
  document.querySelectorAll('.tab').forEach(b=>b.classList.toggle('on',b.dataset.k===state.tab));
}
(function(){
  const tabs=[['관계조회','관계조회']].concat(DATA.tables.map(t=>[t.key,t.label])).concat([['검증','검증']]);
  document.getElementById('tabbar').innerHTML=tabs.map(([k,l])=>
    `<button class="tab${k==='관계조회'?' hl':''}" data-k="${k}" onclick="state.tab='${k}';render()">${l}</button>`).join('');
  document.getElementById('foot-src').textContent=(DATA.is_demo?'데모 데이터':'실 데이터') +
    ' · ' + DATA.tables.map(t=>t.label+' '+t.rows.length).join(' · ');
  render();
})();
</script></body></html>"""
