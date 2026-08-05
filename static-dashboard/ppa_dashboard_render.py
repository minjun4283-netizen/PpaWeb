#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ppa_dashboard_render.py — PPA 6개 표 data → 자기완결 HTML 대시보드.

정산 대시보드(dashboard_render.py)와 같은 패턴: 서버 없이, 데이터를 JSON으로
HTML에 통째로 넣고 JS로 탭/검색/정렬을 처리하는 단일 파일. 표별 탭 + 검증 탭.
검증 탭에서 잡힌 오류는 원본 웹앱처럼 그 표의 정확한 셀을 빨갛게 표시합니다.
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
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css">
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@500;700&display=swap" rel="stylesheet">
<style>
:root{--paper:#F5F4EF;--panel:#FFF;--ink:#16262B;--sub:#5C6B6E;--line:#E3E1D8;
--teal:#0E7C7B;--teal-d:#0A5A59;--teal-w:#E7F1F0;--amber:#B07817;--amber-w:#FAEEDA;
--purple:#534AB7;--purple-w:#EEEDFE;--pass:#1F7A54;--pass-w:#E7F3EC;--fail:#B23A3A;--fail-w:#FBEDEC;--mute:#C9C6BB;}
*{box-sizing:border-box}
body{margin:0;background:var(--paper);color:var(--ink);font-family:Pretendard,sans-serif;-webkit-font-smoothing:antialiased;padding-bottom:60px}
.mono{font-family:'JetBrains Mono',monospace;font-variant-numeric:tabular-nums;letter-spacing:-.02em}
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
section{margin-top:22px}
.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:16px}
.kpi{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:13px 15px;display:flex;flex-direction:column;gap:3px}
.kpi.accent{background:var(--teal);border-color:var(--teal)}.kpi.accent .kk,.kpi.accent .ks{color:#CDE8E6}.kpi.accent .kv{color:#fff}
.kpi.warn{background:var(--fail);border-color:var(--fail)}.kpi.warn .kk,.kpi.warn .ks{color:#F9DEDC}.kpi.warn .kv{color:#fff}
.kk{font-size:11.5px;color:var(--sub);font-weight:600}.kv{font-family:'JetBrains Mono',monospace;font-size:21px;font-weight:700;letter-spacing:-.03em}
.ks{font-size:11px;color:var(--sub);font-family:'JetBrains Mono',monospace}
.panel{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:16px 18px;margin-bottom:16px}
.ph{display:flex;align-items:baseline;gap:10px;margin-bottom:12px}.ph h3{font-size:14.5px;margin:0;font-weight:700}.ph .sub{font-size:12px;color:var(--sub);margin-left:auto}
.toolbar{display:flex;gap:10px;align-items:center;margin-bottom:10px;flex-wrap:wrap}
.search{flex:1;min-width:180px;font-size:13px;padding:8px 12px;border:1px solid var(--line);border-radius:8px;background:var(--paper);color:var(--ink)}
.count{font-size:12px;color:var(--sub);font-family:'JetBrains Mono',monospace}
.tbl-wrap{max-height:520px;overflow:auto;border:1px solid var(--line);border-radius:10px}
table{width:100%;border-collapse:collapse;background:var(--panel)}
th,td{padding:9px 13px;font-size:13px;border-bottom:1px solid var(--line);text-align:left;white-space:nowrap}
thead th{position:sticky;top:0;background:#EFEEE7;font-size:11px;letter-spacing:.02em;color:var(--sub);font-weight:700;text-transform:uppercase;cursor:pointer;z-index:1}
thead th .ar{color:var(--teal);margin-left:3px}
tbody tr:hover{background:#FAF9F5}tbody tr.rowerr{background:var(--fail-w)}
td.cellerr{background:var(--fail-w);color:var(--fail);font-weight:700;border-radius:4px}
.pkbadge,.fkbadge{font-size:9.5px;font-weight:700;padding:1px 6px;border-radius:999px;margin-left:6px;vertical-align:middle}
.pkbadge{background:var(--teal);color:#fff}.fkbadge{background:var(--purple-w);color:var(--purple)}
.badge{font-size:11px;font-weight:700;padding:3px 10px;border-radius:20px;white-space:nowrap}
.badge.ok{color:var(--pass);background:var(--pass-w)}.badge.no{color:var(--fail);background:var(--fail-w)}
.chk{display:grid;grid-template-columns:1fr auto auto auto;gap:12px;align-items:center;background:var(--panel);border:1px solid var(--line);border-left-width:4px;border-radius:8px;padding:10px 15px;margin-bottom:7px}
.chk.no{border-left-color:var(--fail)}
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
let state={tab:DATA.tables[0].key,sort:{},q:{}};

function tableView(t){
  const q=(state.q[t.key]||'').trim().toLowerCase();
  const sc=state.sort[t.key];
  let rows=t.rows.map((r,idx)=>({...r,_idx:idx}));
  if(q) rows=rows.filter(r=>t.columns.some(c=>String(r.cells[c]??'').toLowerCase().includes(q)));
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
    const fk=t.fk_columns.includes(c)?'<span class="fkbadge">FK</span>':'';
    return `<th onclick="doSort('${t.key}','${c.replace(/'/g,"\\'")}')">${c}${pk}${fk}${ar}</th>`;}).join('');
  const body=rows.map(r=>{
    const err=(r.error_cols||[]).length>0;
    const tds=t.columns.map(c=>{
      const bad=(r.error_cols||[]).includes(c);
      return `<td class="${bad?'cellerr':''}">${r.cells[c]??''}</td>`;}).join('');
    return `<tr class="${err?'rowerr':''}">${tds}</tr>`;}).join('');
  return `<div class="toolbar"><input class="search" placeholder="검색…" value="${q.replace(/"/g,'&quot;')}" oninput="setQ('${t.key}',this.value)">
    <span class="count">${rows.length.toLocaleString('ko-KR')} / ${t.rows.length.toLocaleString('ko-KR')}건</span></div>
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
    ${panel(t.label+' 전체 목록','PK/FK 배지 · 클릭으로 정렬 · 빨간 셀은 검증 오류',tableView(t))}</section>`;
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
  view.innerHTML = state.tab==='검증' ? tVerify() : tData(byKey[state.tab]);
  const st=document.getElementById('status');
  const ok=DATA.validation.total_errors===0;
  st.className=ok?'ok':'no';st.textContent=ok?'전 표 검증 통과':`검증 오류 ${DATA.validation.total_errors}건`;
  document.querySelectorAll('.tab').forEach(b=>b.classList.toggle('on',b.dataset.k===state.tab));
}
(function(){
  const tabs=DATA.tables.map(t=>[t.key,t.label]).concat([['검증','검증']]);
  document.getElementById('tabbar').innerHTML=tabs.map(([k,l])=>`<button class="tab" data-k="${k}" onclick="state.tab='${k}';render()">${l}</button>`).join('');
  document.getElementById('foot-src').textContent=(DATA.is_demo?'데모 데이터':'실 데이터') +
    ' · ' + DATA.tables.map(t=>t.label+' '+t.rows.length).join(' · ');
  render();
})();
</script></body></html>"""
