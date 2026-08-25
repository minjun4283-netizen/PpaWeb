#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ppa_dashboard_render.py — PPA 6개 표 data → 자기완결 HTML 대시보드.

정산 대시보드(dashboard_render.py)와 같은 패턴: 서버 없이, 데이터를 JSON으로
HTML에 통째로 넣고 JS로 전부 처리하는 단일 파일. 외부 CDN/폰트/라이브러리를
전혀 쓰지 않으므로 인터넷이 안 되는 사내망 VDI에서도 그대로 열립니다.

색상은 민트/틸 계열(--teal)을 브랜드 축으로, 배경·글자·상태색(--pass/
--amber/--fail/--info)은 흰 카드 위 WCAG 텍스트 대비 기준으로, 발전원
비중처럼 여러 항목을 동시에 색으로 구분해야 하는 차트는 --s1~--s4(아쿠아·
오렌지·바이올렛·마젠타, 라이트/다크 각각 dataviz 팔레트 6종 검사 통과)를
씁니다. 새 차트 계열색이 필요하면 손으로 고르지 말고
dataviz 스킬의 validate_palette.js로 검증한 값만 CSS 변수로 추가하세요.

화면 구성
  홈       — 한눈에 보기 요약 문장, 요약 KPI, 발전원 비중(계열색 --s1~--s4),
             수급매칭 현황별 비율·용량, 년월별 추이(선/영역 차트로 최근
             12개월 vs 전년 동기 겹쳐보기 — 작년 데이터가 없으면 비교선은
             조용히 숨김, 신규 계약/공급기한 · 건수/용량 토글 + 눈금·격자선,
             점을 누르면 그 달의 일별 막대 추이로 드릴다운 — 브라우저
             뒤로가기나 "← 월별로"로 복귀),
             미확보/만료임박, 검증·변경 요약, 표 관계도 — 각 항목을 클릭하면
             해당 표로 조건이 적용된 채 이동
  관계조회 — PK 하나로 FK 체인 전체(발전소↔구매계약↔수급매칭↔전기사용지↔
             판매계약↔수요기업)를 한 화면에
  비교     — 레코드 최대 3건 나란히 비교 (다른 값 강조)
  표별 탭  — 검색·컬럼필터·기간필터·정렬·페이징·컬럼선택·상세모달·내려받기
  변경     — 직전 생성분 대비 추가/수정/삭제 (ppa_changes.py 스냅샷 비교 결과)
  검증     — PK/FK/조합중복 오류를 표·항목별로, 클릭하면 해당 레코드로 이동

화면 뼈대는 좌측 사이드바(그룹별 아이콘 내비) + 상단바(현재 화면 제목·
전역검색·검증/변경 pill)로 되어 있고, 900px 이하에서는 사이드바가
오프캔버스 드로어로 바뀝니다(햄버거로 열고, 배경 클릭·Esc·항목 선택 시
자동으로 닫힘). 드릴다운 클릭(KPI/현황행/추이막대/표 행 등)으로 화면이
바뀌면 브라우저 뒤로가기로 돌아갈 수 있습니다 — render() 안에서 탭·조회·
필터·모달처럼 "화면이 바뀌었다고 체감되는" 상태만 골라 History API에
쌓고, 검색어·정렬·페이지 이동 같은 잦은 조작은 쌓지 않습니다
(navSnapshot/applyNavSnapshot/popstate, `_JS`의 "뒤로가기" 절 참고).
새로고침해도 관계조회의 `#lookup=표:PK` 딥링크만은 URL 해시로 복원됩니다.

유지보수를 쉽게 하려고 스타일(_CSS)·뼈대(_HTML)·동작(_JS)을 분리해두고
모듈을 읽을 때 한 번만 합칩니다.
"""
import datetime
import json

# ─────────────────────────────────────────────────────────────────────────────
# 스타일
# ─────────────────────────────────────────────────────────────────────────────
_CSS = r"""
:root{
  /* 민트/틸 기반 팔레트 — 배경·글자·상태색은 WCAG 텍스트 대비로,
     차트 계열색(--s1..--s4)은 dataviz 6종 검사(밝기대역·채도하한·CVD
     분리·일반시야하한·대비·고정순서)를 라이트/다크 각각 통과한 값만
     사용합니다(scripts/validate_palette.js 실측). */
  --paper:#F5FAF9;--panel:#FFFFFF;--ink:#142A26;--sub:#5B6B65;--line:#DCE8E4;
  --thead-bg:#EDF5F2;--row-hover:#F1FAF8;--shadow:0 8px 24px rgba(10,20,18,.10);
  --shadow-sm:0 1px 2px rgba(10,20,18,.04),0 2px 10px rgba(10,20,18,.05);
  --teal:#0B8577;--teal-d:#075C52;--teal-w:#E3F5F1;
  --amber:#B45309;--amber-w:#FCEFD9;--purple:#6D28D9;--purple-w:#EEEAFB;
  --pass:#15803D;--pass-w:#E7F5EC;--fail:#B91C1C;--fail-w:#FBEAEA;
  --info:#1D4ED8;--info-w:#E7EEFC;--mute:#C3D0CB;
  --s1:#1baf7a;--s2:#eb6834;--s3:#4a3aa7;--s4:#e87ba4; /* 차트 계열: 아쿠아·오렌지·바이올렛·마젠타 */
}
:root[data-theme="dark"]{
  --paper:#12191A;--panel:#1B2422;--ink:#EAF3F0;--sub:#93A8A1;--line:#29332F;
  --thead-bg:#212B28;--row-hover:#1F2926;--shadow:0 8px 24px rgba(0,0,0,.45);
  --shadow-sm:0 1px 2px rgba(0,0,0,.3),0 2px 12px rgba(0,0,0,.35);
  --teal:#2FBFA8;--teal-d:#6FE0CB;--teal-w:#0E332E;
  --amber:#E3A23D;--amber-w:#3A2C12;--purple:#A78BFA;--purple-w:#241E3D;
  --pass:#4ADE80;--pass-w:#123420;--fail:#F87171;--fail-w:#3A1414;
  --info:#60A5FA;--info-w:#142440;--mute:#48534E;
  --s1:#199e70;--s2:#d95926;--s3:#9085e9;--s4:#d55181;
}
:root{
  --font-sans:"Pretendard","Pretendard Variable",-apple-system,BlinkMacSystemFont,"Segoe UI","Malgun Gothic","Apple SD Gothic Neo",sans-serif;
  --font-mono:ui-monospace,SFMono-Regular,Consolas,monospace;
}
*{box-sizing:border-box}
html{-webkit-text-size-adjust:100%}
body{margin:0;background:var(--paper);color:var(--ink);padding-bottom:60px;
  font-family:var(--font-sans);
  -webkit-font-smoothing:antialiased;text-rendering:optimizeLegibility}
.mono,.kv,.ks,.count{font-family:var(--font-mono);font-variant-numeric:tabular-nums;letter-spacing:-.02em}
.wrap{max-width:1280px;margin:0 auto;padding:0 22px}
button,input,select{font-family:inherit}
/* 전역 포커스 표시 — 브라우저 기본값 대신 브랜드색으로 통일(키보드 탐색
   접근성 + 시각적 일관성). 마우스 클릭 시엔 표시 안 함(:focus-visible). */
input:focus-visible,select:focus-visible,textarea:focus-visible,
button:focus-visible,.search:focus-visible{outline:2px solid var(--teal);outline-offset:1px}
.search:focus,input[type="date"]:focus,input[type="month"]:focus,select:focus{
  border-color:var(--teal);outline:none}

/* 셸: 사이드바 + 상단바 — z-index 스케일: 상단바(sticky)=30, 사이드바
   (데스크톱 sticky·모바일 드로어 공용)=36, 드로어 백드롭=35(사이드바
   바로 아래, 상단바 위), 전역검색 결과=48(드로어 위), 모달/모달백드롭
   =60(기존 값, 항상 최상단) */
.shell{display:flex;min-height:100vh}
.sidebar{width:238px;flex-shrink:0;background:var(--panel);border-right:1px solid var(--line);
  display:flex;flex-direction:column;position:sticky;top:0;height:100vh;z-index:36}
.brand{display:flex;align-items:center;gap:10px;padding:20px 18px 16px;flex-shrink:0}
.brandmark{width:34px;height:34px;border-radius:10px;background:var(--teal);color:#fff;
  display:flex;align-items:center;justify-content:center;font-weight:800;font-size:16px;flex-shrink:0}
.brandtext{min-width:0}
.brandtext b{display:block;font-size:14px;font-weight:800;letter-spacing:-.01em;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.brandtext span{display:block;font-size:10.5px;color:var(--sub);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.navwrap{flex:1;min-height:0;overflow-y:auto;padding:6px 12px 14px}
.navgroup+.navgroup{margin-top:14px}
.navgrouplabel{font-size:10px;font-weight:700;letter-spacing:.09em;text-transform:uppercase;color:var(--mute);padding:8px 10px 5px}
.navitem{display:flex;align-items:center;gap:10px;width:100%;text-align:left;font-size:13.5px;font-weight:600;
  color:var(--sub);background:none;border:none;border-left:3px solid transparent;border-radius:0 9px 9px 0;
  padding:9px 10px 9px 9px;cursor:pointer;margin-bottom:2px}
.navitem:hover{background:var(--paper);color:var(--ink)}
.navitem:focus-visible{outline:2px solid var(--teal);outline-offset:-2px}
.navitem.on{background:var(--teal-w);color:var(--teal-d);border-left-color:var(--teal);font-weight:800}
.navicon{font-size:15px;width:18px;text-align:center;flex-shrink:0}
.navlabel{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.sidefoot{padding:12px 16px 16px;border-top:1px solid var(--line);flex-shrink:0}
.sidethemebtn{display:flex;align-items:center;gap:8px;width:100%;font-size:12.5px;font-weight:600;color:var(--sub);
  background:var(--paper);border:1px solid var(--line);border-radius:9px;padding:9px 12px;cursor:pointer}
.sidethemebtn:hover{color:var(--ink);border-color:var(--teal)}
.sidebarbackdrop{display:none;position:fixed;inset:0;background:rgba(10,15,17,.45);z-index:35}
.menubtn{display:none;border:1px solid var(--line);background:var(--panel);border-radius:9px;width:36px;height:36px;
  font-size:15px;cursor:pointer;align-items:center;justify-content:center;color:var(--ink);flex-shrink:0}

.main{flex:1;min-width:0;display:flex;flex-direction:column}
.topbar{position:sticky;top:0;z-index:30;background:var(--paper);border-bottom:1px solid var(--line);
  display:flex;align-items:center;justify-content:space-between;gap:14px;padding:14px 26px;flex-wrap:wrap}
.topbar-left{display:flex;align-items:center;gap:12px;min-width:0}
.topbar h1{font-size:19px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.topbar-right{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.eyebrow{font-size:11.5px;letter-spacing:.16em;text-transform:uppercase;color:var(--teal);font-weight:700;margin:0 0 3px}
h1{font-size:22px;margin:0;letter-spacing:-.02em;font-weight:800}
.pill{font-size:12px;font-weight:700;padding:4px 11px;border-radius:20px;white-space:nowrap;border:none;cursor:pointer}
.pill.ok{color:var(--pass);background:var(--pass-w)}
.pill.no{color:var(--fail);background:var(--fail-w)}
.pill.chg{color:var(--info);background:var(--info-w)}
.iconbtn{border:1px solid var(--line);background:var(--panel);border-radius:20px;width:32px;height:32px;font-size:14px;cursor:pointer;display:flex;align-items:center;justify-content:center;color:var(--ink)}
.demo{background:var(--amber-w);color:var(--amber);font-size:12px;font-weight:600;padding:7px 22px;text-align:center;border-bottom:1px solid var(--line)}
.gsearchwrap{position:relative;width:250px;max-width:44vw}
.globalresults{position:absolute;left:0;top:100%;margin-top:6px;width:340px;max-width:min(340px,calc(100vw - 32px));
  background:var(--panel);border:1px solid var(--line);border-radius:10px;box-shadow:var(--shadow);max-height:360px;overflow:auto;z-index:48;display:none}
.globalresults.show{display:block}
.gresrow{display:block;width:100%;text-align:left;padding:9px 16px;border:none;background:none;cursor:pointer;font-size:13px;border-bottom:1px solid var(--line);color:var(--ink)}
.gresrow:last-child{border-bottom:none}.gresrow:hover{background:var(--teal-w)}
.gresrow .tag{font-size:10.5px;font-weight:700;color:var(--teal-d);background:var(--teal-w);padding:1px 7px;border-radius:20px;margin-right:8px}
.tabdot{display:inline-block;min-width:17px;padding:0 5px;margin-left:5px;font-size:10.5px;line-height:16px;border-radius:20px;background:var(--fail);color:#fff;text-align:center;font-weight:700}
.tabdot.info{background:var(--info)}

section{margin-top:20px}
#view{animation:fadein .15s ease}
@keyframes fadein{from{opacity:0;transform:translateY(2px)}to{opacity:1;transform:none}}

/* KPI — 엔터프라이즈 SaaS 대시보드 느낌의 타이포그래피 계층: 라벨은
   작고 옅은 대문자 eyebrow, 값은 크고 굵게, 보조문구는 더 작고 옅게. */
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(215px,1fr));gap:14px;margin-bottom:18px}
.kpi{position:relative;background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:16px 17px;display:flex;flex-direction:column;gap:5px;text-align:left;box-shadow:var(--shadow-sm)}
.kpi.accent{background:var(--teal);border-color:var(--teal)}
.kpi.accent .kk,.kpi.accent .ks{color:#CDE8E6}.kpi.accent .kv{color:#fff}
.kpi.warn{background:var(--fail);border-color:var(--fail)}
.kpi.warn .kk,.kpi.warn .ks{color:#F9DEDC}.kpi.warn .kv{color:#fff}
.kpi.warn::after{content:'⚠';position:absolute;top:11px;right:13px;font-size:13px;opacity:.85;line-height:1}
.kpi.clickable{cursor:pointer}.kpi.clickable:hover{border-color:var(--teal);box-shadow:var(--shadow)}
.kk{font-size:10.5px;color:var(--sub);font-weight:700;text-transform:uppercase;letter-spacing:.045em}
.infotip{display:inline-flex;align-items:center;justify-content:center;margin-left:4px;color:var(--mute);
  font-size:11px;cursor:help;text-transform:none;letter-spacing:0}
.infotip:hover,.infotip:focus-visible{color:var(--teal)}
.kv{font-size:23px;font-weight:800;letter-spacing:-.015em}
.ks{font-size:11.5px;color:var(--sub)}

.panel{background:var(--panel);border:1px solid var(--line);border-radius:16px;padding:18px 20px;margin-bottom:18px;box-shadow:var(--shadow-sm)}
.ph{display:flex;align-items:baseline;gap:10px;margin-bottom:14px;flex-wrap:wrap}
.ph h3{font-size:15px;margin:0;font-weight:800;letter-spacing:-.01em}
.ph .sub{font-size:12px;color:var(--sub);margin-left:auto}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px}

/* 툴바 / 필터 */
.toolbar{display:flex;gap:8px;align-items:center;margin-bottom:10px;flex-wrap:wrap}
.search{flex:1;min-width:170px;font-size:13px;padding:8px 12px;border:1px solid var(--line);border-radius:8px;background:var(--paper);color:var(--ink)}
.count{font-size:12px;color:var(--sub);white-space:nowrap}
.btn{font-size:12.5px;font-weight:600;color:var(--sub);background:var(--paper);border:1px solid var(--line);border-radius:8px;padding:8px 12px;cursor:pointer;white-space:nowrap}
.btn:hover{border-color:var(--teal);color:var(--teal-d)}
.btn.on{background:var(--teal);border-color:var(--teal);color:#fff}
.btn.danger{color:var(--fail);background:var(--fail-w);border-color:transparent}
.btn.primary{color:var(--teal-d);background:var(--teal-w);border-color:transparent}
.filterbar{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:8px;align-items:center}
.filtersel{font-size:12.5px;padding:7px 10px;border:1px solid var(--line);border-radius:8px;background:var(--paper);color:var(--ink);max-width:210px}
.chiprow{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:10px}
.fchip{display:inline-flex;align-items:center;gap:6px;font-size:12px;font-weight:600;background:var(--purple-w);color:var(--purple);padding:4px 6px 4px 11px;border-radius:20px}
.fchip button{border:none;background:none;color:inherit;cursor:pointer;font-size:13px;line-height:1;padding:2px}
.datef{display:inline-flex;align-items:center;gap:6px;font-size:12.5px;background:var(--paper);border:1px solid var(--line);border-radius:8px;padding:6px 10px}
.dflabel{color:var(--sub);font-weight:600}
.datef input[type=date]{border:none;background:none;font-size:12.5px;color:var(--ink);font-family:inherit;padding:0}
.dfsep{color:var(--sub)}
.drop{position:relative;display:inline-block}
.drop summary{list-style:none;cursor:pointer;font-size:12.5px;font-weight:600;color:var(--sub);background:var(--paper);border:1px solid var(--line);border-radius:8px;padding:8px 12px;white-space:nowrap}
.drop summary::-webkit-details-marker{display:none}
.drop summary:hover{border-color:var(--teal);color:var(--teal-d)}
.dropbody{position:absolute;z-index:25;margin-top:4px;background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:10px 14px;min-width:210px;max-height:300px;overflow:auto;box-shadow:var(--shadow)}
.dropbody.right{right:0}
.colopt{display:flex;align-items:center;gap:8px;font-size:12.5px;padding:5px 2px;white-space:nowrap;cursor:pointer}
.dlopt{display:block;width:100%;text-align:left;font-size:12.5px;padding:8px 10px;border:none;background:none;border-radius:6px;cursor:pointer;color:var(--ink)}
.dlopt:hover{background:var(--teal-w)}
.dlopt small{display:block;color:var(--sub);font-size:11px;margin-top:1px}

/* 표 */
.tbl-wrap{max-height:560px;overflow:auto;border:1px solid var(--line);border-radius:10px}
table{width:100%;border-collapse:separate;border-spacing:0;background:var(--panel)}
th,td{padding:10px 14px;font-size:13px;border-bottom:1px solid var(--line);text-align:left;white-space:nowrap;background:var(--panel)}
thead th{position:sticky;top:0;background:var(--thead-bg);font-size:10.5px;letter-spacing:.05em;text-transform:uppercase;color:var(--sub);font-weight:700;cursor:pointer;z-index:2}
thead th.nosort{cursor:default}
thead th .ar{color:var(--teal);margin-left:3px}
tbody tr:hover td{background:var(--row-hover)}
tbody tr.clickrow{cursor:pointer}
tbody tr.rowerr td{background:var(--fail-w)}
tbody tr.rowadd td{background:var(--pass-w)}
tbody tr.rowchg td{background:var(--info-w)}
tbody tr.rootrow td{background:var(--teal-w);font-weight:700}
td.cellerr{color:var(--fail);font-weight:700;box-shadow:inset 3px 0 0 var(--fail)}
td.cellchg{color:var(--info);font-weight:700;box-shadow:inset 3px 0 0 var(--info)}
td.diffcell{background:var(--amber-w)!important;color:var(--amber);font-weight:700}
td.num,th.num{text-align:right;font-family:ui-monospace,SFMono-Regular,Consolas,monospace;font-variant-numeric:tabular-nums}
/* 좌측 PK 열 고정 — 컬럼 많은 표(판매계약 등)에서 가로 스크롤 시 기준점 유지 */
.tbl-wrap.stickyfirst thead th:first-child{left:0;z-index:3}
.tbl-wrap.stickyfirst td:first-child,.tbl-wrap.stickyfirst th:first-child{position:sticky;left:0;z-index:1;box-shadow:1px 0 0 var(--line)}
.tbl-wrap.stickyfirst tbody tr:hover td:first-child{background:var(--row-hover)}
.tbl-wrap.stickyfirst tbody tr.rowerr td:first-child{background:var(--fail-w)}
.tbl-wrap.stickyfirst tbody tr.rowadd td:first-child{background:var(--pass-w)}
.tbl-wrap.stickyfirst tbody tr.rowchg td:first-child{background:var(--info-w)}
.emptyrow{text-align:center;padding:26px;color:var(--sub)}

/* 페이징 */
.pager{display:flex;align-items:center;gap:6px;justify-content:center;margin-top:12px;flex-wrap:wrap}
.pgbtn{font-size:12.5px;font-weight:600;min-width:34px;padding:7px 10px;border:1px solid var(--line);background:var(--panel);color:var(--sub);border-radius:8px;cursor:pointer}
.pgbtn:hover:not(:disabled){border-color:var(--teal);color:var(--teal-d)}
.pgbtn.on{background:var(--teal);border-color:var(--teal);color:#fff}
.pgbtn:disabled{opacity:.4;cursor:default}
.pginfo{font-size:12px;color:var(--sub);margin:0 8px}
.pgdots{color:var(--sub);padding:0 2px}

/* 배지 */
.pkbadge,.fkbadge{font-size:9.5px;font-weight:700;padding:1px 6px;border-radius:999px;margin-left:6px;vertical-align:middle}
.pkbadge{background:var(--teal);color:#fff}
.fkbadge{background:var(--purple-w);color:var(--purple)}
.badge{font-size:11px;font-weight:700;padding:3px 10px;border-radius:20px;white-space:nowrap;display:inline-block}
.badge.ok{color:var(--pass);background:var(--pass-w)}
.badge.no{color:var(--fail);background:var(--fail-w)}
.badge.info{color:var(--info);background:var(--info-w)}
.badge.warn{color:var(--amber);background:var(--amber-w)}
.badge.mute{color:var(--sub);background:var(--paper)}
.idlink{color:var(--teal-d);text-decoration:underline;text-underline-offset:2px;cursor:pointer;font-weight:700}
.idlink:hover{color:var(--teal)}
.fkname{color:var(--sub);font-weight:500;margin-left:6px}
.datewarn{margin-left:6px;font-size:10.5px;font-weight:700;padding:1px 7px;border-radius:20px}
.datewarn.exp{background:var(--fail-w);color:var(--fail)}
.datewarn.soon{background:var(--amber-w);color:var(--amber)}

/* 검증 / 변경 목록 */
.chk{display:grid;grid-template-columns:1fr auto auto auto;gap:12px;align-items:center;background:var(--panel);border:1px solid var(--line);border-left-width:4px;border-radius:8px;padding:10px 15px;margin-bottom:7px;cursor:pointer}
.chk:hover{background:var(--teal-w)}
.chk.no{border-left-color:var(--fail)}
.chk.add{border-left-color:var(--pass)}
.chk.chg{border-left-color:var(--info)}
.chk.del{border-left-color:var(--mute)}
.chgval{font-size:12.5px}
.chgold{color:var(--sub);text-decoration:line-through;margin-right:6px}
.chgnew{color:var(--info);font-weight:700}

/* 관계조회 */
.subtabbar{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:2px}
.subtab{font-size:12.5px;font-weight:600;color:var(--sub);background:var(--paper);border:1px solid var(--line);cursor:pointer;padding:8px 14px;border-radius:20px}
.subtab.on{color:#fff;background:var(--teal);border-color:var(--teal)}
.candlist{display:flex;flex-direction:column;gap:5px;margin-top:12px;max-height:340px;overflow:auto}
.candrow{text-align:left;font-size:13px;padding:10px 13px;border:1px solid var(--line);border-radius:8px;background:var(--panel);cursor:pointer;color:var(--ink)}
.candrow:hover{background:var(--teal-w);border-color:var(--teal)}
.nocand{font-size:12.5px;color:var(--sub);padding:10px 2px}
.chipcount{font-size:11.5px;font-weight:700;color:var(--teal-d);background:var(--teal-w);padding:3px 10px;border-radius:20px;margin-right:6px;display:inline-block;margin-top:4px}
.lookuphead,.homehead{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px}
.homehead{margin-bottom:14px}

/* 관계형 탐색 */
.colchip{font-size:12.5px;font-weight:600;color:var(--sub);background:var(--paper);border:1px solid var(--line);
  border-radius:20px;padding:6px 13px;cursor:pointer;display:inline-flex;align-items:center;gap:5px}
.colchip:hover{border-color:var(--teal);color:var(--teal-d)}
.colchip.on{background:var(--teal);border-color:var(--teal);color:#fff}
.colchip.transit{background:var(--paper);border-style:dashed;color:var(--sub);opacity:.8}
.chipdist{font-size:10px;font-weight:700;opacity:.75}
.colgroup{padding:8px 0;border-bottom:1px solid var(--line)}
.colgroup:last-child{border-bottom:none}
.colgrouphead{display:flex;align-items:center;gap:8px;font-size:12.5px;font-weight:700;margin-bottom:7px}

/* 출력 컬럼 순서 — 순수 HTML5 Drag and Drop 칩 목록 */
.excol-draglist{display:flex;flex-wrap:wrap;gap:7px;padding:2px 2px 10px;margin-bottom:10px;border-bottom:1px dashed var(--line)}
.excol-chip{display:inline-flex;align-items:center;gap:6px;font-size:12.5px;font-weight:600;color:var(--ink);
  background:var(--panel);border:1px solid var(--line);border-radius:9px;padding:6px 8px 6px 6px;
  cursor:grab;user-select:none;box-shadow:var(--shadow-sm);transition:opacity .15s,border-color .15s}
.excol-chip:active{cursor:grabbing}
.excol-chip.dragging{opacity:.35}
.excol-chip.dragover{border-color:var(--teal);box-shadow:0 0 0 2px var(--teal-w)}
.excol-handle{color:var(--mute);font-size:11px;letter-spacing:-2px;line-height:1}
.excol-remove{border:none;background:none;color:var(--sub);cursor:pointer;font-size:12px;padding:0 0 0 2px;line-height:1}
.excol-remove:hover{color:var(--fail)}
.excol-hint{font-size:11.5px;color:var(--sub);margin:-4px 0 12px}
.thtable{display:block;font-size:9.5px;font-weight:700;color:var(--teal-d);letter-spacing:.02em;margin-bottom:1px}
td.cellmiss{color:var(--mute);text-align:center}
tbody tr.rowmiss td{background:var(--amber-w)}
.tbl-wrap.stickyfirst tbody tr.rowmiss td:first-child{background:var(--amber-w)}

/* 홈 위젯 — "한눈에 보기" 카드형 레이아웃(핵심 지표 칩 + 현황 아이콘
   스트립 + 확인 필요 알림 줄), 문장을 읽지 않고 훑어만 봐도 파악되도록 */
.insight{background:linear-gradient(135deg,var(--teal-d),var(--teal));border-radius:14px;padding:20px 24px;margin-bottom:16px;color:#fff}
.insight .ieyebrow{font-size:11px;letter-spacing:.14em;text-transform:uppercase;font-weight:700;color:#CDE8E6;margin:0 0 12px}
.insight .dim{color:#CDE8E6;font-weight:600}
.insight .ichiprow{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:12px}
.insight .ichip{display:flex;align-items:center;gap:9px;background:rgba(255,255,255,.14);border-radius:12px;padding:9px 14px}
.insight .ichipicon{font-size:21px;line-height:1}
.insight .ichiptext{display:flex;flex-direction:column;line-height:1.3}
.insight .ichiptext b{font-size:17.5px;font-weight:800;letter-spacing:-.01em;order:1}
.insight .ichiptext .dim{font-size:11.5px;order:2}
.insight .ichiplabel{font-size:10px;color:#CDE8E6;text-transform:uppercase;letter-spacing:.05em;font-weight:700;order:0}
.insight .istatusrow{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:12px}
.insight .ischip{display:inline-flex;align-items:center;gap:5px;background:rgba(255,255,255,.16);border-radius:20px;padding:5px 11px;font-size:13.5px;font-weight:700}
.insight .ischip.zero{opacity:.4}
.insight .ialert{font-size:14.5px;font-weight:700;background:rgba(255,255,255,.16);border-radius:10px;padding:10px 15px;letter-spacing:-.01em}
.insight .ialert.urgent{background:#fff;color:var(--fail)}
.insight .ialert.ok{font-weight:600}
@media(max-width:640px){
  .insight .ichip{padding:7px 11px;gap:7px}
  .insight .ichipicon{font-size:17px}
  .insight .ichiptext b{font-size:15px}
  .insight .ialert{font-size:13px}
}

.unsecrow{display:flex;justify-content:space-between;align-items:center;gap:10px;padding:10px 5px;border-bottom:1px solid var(--line);cursor:pointer;font-size:13px}
.unsecrow:last-child{border-bottom:none}.unsecrow:hover{background:var(--paper)}
/* 라벨이 길어도(예: "발전 설비용량 대비 수요계약 미매칭 잔여용량") 좁은
   화면에서 가로로 넘치지 않고 줄바꿈되도록 — 값·배지 칸은 줄어들지 않게 */
.unsecrow>span:first-child{flex:1 1 auto;min-width:0;white-space:normal;word-break:keep-all}
.unsecrow>span:not(:first-child){flex:0 0 auto}
.unsecrow.isok{opacity:.6}

/* 우선순위 조치 필요 항목 — 카테고리별 그룹핑 */
.actionsummary{display:flex;gap:8px;flex-wrap:wrap;margin:2px 0 14px}
.actioncat+.actioncat{margin-top:18px;padding-top:14px;border-top:1px solid var(--line)}
.actioncathead{display:flex;align-items:center;gap:8px;font-size:12px;font-weight:800;color:var(--sub);text-transform:uppercase;letter-spacing:.04em;margin-bottom:2px}
.actioncatcount{font-size:10.5px;font-weight:700;color:var(--ink);background:var(--paper);border-radius:20px;padding:1px 8px;letter-spacing:0;text-transform:none}
.mixrow{display:flex;align-items:center;gap:8px;padding:6px 0}
.mixdot{width:9px;height:9px;border-radius:3px;flex-shrink:0}
.mixlabel{width:70px;font-size:12.5px;color:var(--sub);flex-shrink:0}
.mixbar{flex:1;height:10px;background:var(--paper);border-radius:6px;overflow:hidden}
.mixfill{height:100%;background:var(--teal);border-radius:6px}
.mixval{font-size:12px;width:110px;text-align:right;flex-shrink:0}
.schemarow{display:flex;align-items:center;gap:8px;flex-wrap:wrap;padding:8px 0}
.schematag{font-size:11px;font-weight:700;color:var(--sub);width:56px;flex-shrink:0}
.schemabox{display:flex;flex-direction:column;align-items:center;gap:2px;font-size:12.5px;font-weight:700;background:var(--paper);border:1px solid var(--line);border-radius:10px;padding:8px 16px;cursor:pointer;color:var(--ink)}
.schemabox:hover{border-color:var(--teal);color:var(--teal-d)}
.schemacount{font-size:10.5px;font-weight:600;color:var(--sub)}
.schemaarrow{color:var(--mute);font-size:16px;flex-shrink:0}

/* 현황별 비율·용량 (스택형 막대 + 범례) */
.statusbar{display:flex;height:14px;border-radius:7px;overflow:hidden;background:var(--paper);margin-bottom:14px}
.statusseg{height:100%}
.statusseg+.statusseg{box-shadow:inset 2px 0 0 var(--panel)}
.statuslegendhead{display:grid;grid-template-columns:18px 1fr 60px 70px 130px;gap:10px;padding:2px 4px 6px;font-size:10.5px;font-weight:700;color:var(--sub);letter-spacing:.02em}
.statusrow{display:grid;grid-template-columns:18px 1fr 60px 70px 130px;gap:10px;align-items:center;padding:8px 4px;border-bottom:1px solid var(--line);cursor:pointer;font-size:13px}
.statusrow:last-of-type{border-bottom:none}
.statusrow:hover{background:var(--paper)}
.statusdot{width:11px;height:11px;border-radius:4px;flex-shrink:0}
.statuslabel{font-weight:600}
.statuspct,.statuscnt,.statuscap{text-align:right;color:var(--sub)}
.statuscnt{font-weight:700;color:var(--ink)}
.statustotal{margin-top:10px;padding-top:10px;border-top:1px solid var(--line);font-size:12px;color:var(--sub);text-align:right}

/* 년월별 추이 */
.trendtools{display:flex;gap:14px;flex-wrap:wrap;align-items:center;margin-bottom:14px}
.trendrange{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:14px;font-size:12px;color:var(--sub)}
.trendrange label{display:flex;align-items:center;gap:6px;font-weight:600}
.trendrangeinput{font-family:inherit;font-size:12.5px;padding:6px 9px;border:1px solid var(--line);border-radius:8px;background:var(--panel);color:var(--ink)}
.trendrangesep{color:var(--mute)}
.trendrangenote{font-size:11px;color:var(--amber)}
.segtoggle{display:inline-flex;background:var(--paper);border:1px solid var(--line);border-radius:8px;padding:2px}
.segtoggle button{font-size:12px;font-weight:600;color:var(--sub);background:none;border:none;border-radius:6px;padding:6px 12px;cursor:pointer;white-space:nowrap}
.segtoggle button.on{background:var(--teal);color:#fff}
.trendbreadcrumb{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:12px;padding-bottom:10px;border-bottom:1px solid var(--line)}
.trendback{font-size:12.5px;font-weight:700;color:var(--teal-d);background:var(--teal-w);border:none;border-radius:8px;padding:7px 12px;cursor:pointer}
.trendback:hover{background:var(--teal);color:#fff}
.trendcrumbnow{font-size:12.5px;font-weight:700;color:var(--ink)}
.trendplot{display:flex;gap:8px}
.trendaxis{position:relative;width:34px;height:150px;flex-shrink:0}
.trendaxistick{position:absolute;right:4px;left:0;text-align:right;transform:translateY(50%);font-size:9.5px;color:var(--sub);white-space:nowrap;line-height:1}
.trendarea{position:relative;flex:1;min-width:0;height:150px}
.trendgrid{position:absolute;inset:0;pointer-events:none}
.trendgridline{position:absolute;left:0;right:0;border-top:1px solid var(--line)}
.trendchart{position:relative;z-index:1;display:flex;align-items:flex-end;gap:3px;height:100%;padding:0 2px}
.trendcol{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:flex-end;height:100%;cursor:pointer;min-width:6px}
.trendbar{width:100%;max-width:26px;background:var(--teal);border-radius:4px 4px 0 0;transition:background .12s}
.trendcol:hover .trendbar{background:var(--teal-d)}
.trendcol.peak .trendbar{background:var(--amber)}
.trendlabel{font-size:9.5px;color:var(--ink);font-weight:700;margin-bottom:3px;white-space:nowrap}
.trendxrow{display:flex;gap:8px}
.trendaxisspacer{width:34px;flex-shrink:0}
.trendticks{flex:1;min-width:0;display:flex;gap:3px;padding:6px 2px 0}
.trendtick{flex:1;text-align:center;font-size:9.5px;color:var(--sub);min-width:6px;white-space:nowrap}
.trendfoot{margin-top:10px;font-size:12px;color:var(--sub);display:flex;justify-content:space-between;flex-wrap:wrap;gap:6px}
.trendsvg{width:100%;height:100%;display:block;overflow:visible}
.trendendlabel{font-size:10.5px;font-weight:700;fill:var(--ink)}
.trendlegend{display:flex;gap:18px;flex-wrap:wrap;font-size:11.5px;color:var(--sub);margin:10px 0 0 42px}
.trendlegend span{display:inline-flex;align-items:center;gap:6px}
.trendlegendline{width:16px;height:0;border-top:2.5px solid var(--teal);display:inline-block}
.trendlegendline.prev{border-top:2.5px dashed var(--mute)}

/* 상세 모달 */
.backdrop{position:fixed;inset:0;background:rgba(10,15,17,.45);z-index:60;display:flex;align-items:center;justify-content:center;padding:20px}
.modal{background:var(--panel);border-radius:14px;box-shadow:var(--shadow);width:min(720px,100%);max-height:88vh;display:flex;flex-direction:column}
.modalhead{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;padding:18px 20px 12px;border-bottom:1px solid var(--line)}
.modalhead h3{margin:0;font-size:16px;font-weight:800}
.modalhead .sub{font-size:12px;color:var(--sub);margin-top:3px}
.modalbody{padding:6px 20px 16px;overflow:auto}
.modalfoot{display:flex;gap:8px;padding:12px 20px 18px;flex-wrap:wrap;border-top:1px solid var(--line)}
.drow{display:grid;grid-template-columns:190px 1fr;gap:12px;padding:9px 2px;border-bottom:1px solid var(--line);font-size:13px;align-items:baseline}
.drow:last-child{border-bottom:none}
.dkey{color:var(--sub);font-weight:600;font-size:12.5px}
.drow.err .dkey{color:var(--fail)}
.drow.chg .dkey{color:var(--info)}

#toast{position:fixed;bottom:24px;left:50%;transform:translateX(-50%) translateY(20px);background:var(--ink);color:var(--paper);font-size:13px;font-weight:600;padding:10px 18px;border-radius:8px;box-shadow:var(--shadow);opacity:0;pointer-events:none;transition:opacity .2s,transform .2s;z-index:80}
#toast.show{opacity:1;transform:translateX(-50%) translateY(0)}

footer{margin-top:30px;padding-top:16px;border-top:1px solid var(--line);font-size:11.5px;color:var(--sub);display:flex;justify-content:space-between;flex-wrap:wrap;gap:6px}
.printhead{display:none}

/* 반응형 — iPad 세로/휴대폰까지 */
@media(max-width:900px){
  .grid2{grid-template-columns:1fr}
  .drow{grid-template-columns:130px 1fr}
  /* 사이드바 → 오프캔버스 드로어 */
  .menubtn{display:flex}
  .sidebar{position:fixed;left:0;top:0;transform:translateX(-100%);transition:transform .2s ease;box-shadow:var(--shadow)}
  .sidebar.open{transform:translateX(0)}
  .sidebarbackdrop.open{display:block}
}
@media(max-width:640px){
  .wrap,.topbar{padding-left:14px;padding-right:14px}
  .kpis{grid-template-columns:1fr 1fr}
  .toolbar .search{min-width:100%}
  .gsearchwrap{width:auto;max-width:none;flex:1}
  .mixlabel{width:60px}.mixval{width:88px}
  .drow{grid-template-columns:1fr;gap:2px}
  .tbl-wrap{max-height:none}
  .statuslegendhead,.statusrow{grid-template-columns:14px 1fr 34px 46px 72px;gap:5px;font-size:11.5px}
  .statuslegendhead{font-size:9.5px}
  .trendtools{gap:8px}
  .segtoggle button{font-size:11px;padding:6px 8px}
  .trendaxis,.trendaxisspacer{width:26px}
  .trendaxistick{font-size:8.5px;right:2px}
}

/* 보고용 인쇄 */
@media print{
  @page{size:A4 landscape;margin:11mm}
  /* 화면 조작용 UI(탐색·필터·버튼·드래그순서·툴팁 아이콘 등)는 종이에서
     의미가 없으므로 전부 숨기고, 데이터와 인사이트만 A4 보고서 형태로
     남깁니다. */
  .sidebar,.sidebarbackdrop,.topbar,.menubtn,.gsearchwrap,.toolbar,.filterbar,.chiprow,.candlist,.subtabbar,
  .pager,.btn,.drop,.iconbtn,.pill,#globalResults,#toast,.backdrop,footer,.segtoggle,.trendrange,
  .excol-draglist,.excol-hint,.infotip,.excol-remove{display:none!important}
  .insight{background:#fff!important;color:#000!important;border:1px solid #999}
  .insight .ieyebrow,.insight .dim,.insight .ichiplabel{color:#333!important}
  .insight span[title]{border-bottom:none!important}
  .insight .ichip,.insight .ischip,.insight .ialert{background:#f3f3f3!important;color:#000!important}
  .insight .ialert.urgent{border:1.5px solid #b91c1c;font-weight:800}
  .trendlabel{color:#000}
  body{background:#fff;color:#000;padding:0}
  .shell{display:block}
  .main{width:100%}
  .wrap{max-width:none;padding:0}
  section{margin-top:0}
  .printhead{display:block;margin-bottom:10px;border-bottom:2px solid #000;padding-bottom:6px}
  .printhead h2{margin:0 0 3px;font-size:15pt}
  .printhead .pmeta{font-size:8.5pt;color:#333}
  .panel{break-inside:auto;border:1px solid #999;box-shadow:none;margin-bottom:10px;padding:8px 10px}
  .ph{break-after:avoid-page}
  .actioncathead{break-after:avoid-page}
  .tbl-wrap{max-height:none;overflow:visible;border:none}
  table{font-size:8pt}
  th,td{padding:3px 5px;background:#fff!important;border-bottom:1px solid #ccc;white-space:normal}
  thead{display:table-header-group}
  tr,.kpi,.unsecrow,.statusrow,.excol-chip{break-inside:avoid}
  thead th{background:#eee!important;color:#000;position:static;text-transform:none}
  td.cellerr{color:#000;font-weight:700;box-shadow:none;outline:1.2px solid #b23a3a}
  td.cellchg{color:#000;font-weight:700;box-shadow:none;outline:1.2px dashed #1e63a8}
  a.idlink{color:#000;text-decoration:none}
  .kpi{border:1px solid #999;box-shadow:none}
  .kpi.accent,.kpi.warn{background:#fff!important}
  .kpi.accent .kk,.kpi.accent .kv,.kpi.accent .ks,
  .kpi.warn .kk,.kpi.warn .kv,.kpi.warn .ks{color:#000!important}
  /* 배지는 배경색 인쇄 옵션이 꺼져 있어도 구분되도록 글자색 테두리를 둡니다 */
  .badge{border:1px solid currentColor;background:transparent!important}
  .actioncatcount{border:1px solid #999;background:transparent!important;color:#000!important}
}
"""

# ─────────────────────────────────────────────────────────────────────────────
# 동작
# ─────────────────────────────────────────────────────────────────────────────
_JS = r"""
/*__DATA__*/

/* ── 색인: PK로 행 찾기, FK 관계 그래프 ─────────────────────────────────── */
const byKey={};DATA.tables.forEach(t=>byKey[t.key]=t);
const CHANGES=DATA.changes||{has_prev:false};
const CHANGELOG=DATA.changelog||[];
const rowIndex={},rowPos={};
DATA.tables.forEach(t=>{
  rowIndex[t.key]={};rowPos[t.key]={};
  t.rows.forEach((r,i)=>{
    const pkv=r.cells[t.pk];
    if(pkv!==undefined&&pkv!==''&&rowIndex[t.key][String(pkv)]===undefined){
      rowIndex[t.key][String(pkv)]=r;rowPos[t.key][String(pkv)]=i;
    }
  });
});
const edges=[];
DATA.tables.forEach(t=>Object.entries(t.fk||{}).forEach(([col,ref])=>edges.push({table:t.key,col,ref})));
const TODAY=(DATA.generated_at||'').slice(0,10);

/* ── 저장소(브라우저 한정 기억) ─────────────────────────────────────────── */
function readLS(k,f){try{const v=localStorage.getItem(k);return v===null?f:v;}catch(e){return f;}}
function writeLS(k,v){try{localStorage.setItem(k,v);}catch(e){}}

let state={
  tab:'홈',sort:{},q:{},filters:{},dateFilters:{},colQ:{},pkFilter:{},
  hidden:{},page:{},pageSize:Number(readLS('ppa_pagesize','50'))||50,
  onlyErr:{},onlyChg:{},
  lookup:null,lookupTable:DATA.tables[0].key,lookupQ:'',lookupDepth:Number(readLS('ppa_depth','2')),
  recent:[],pinned:[],modal:null,
  theme:readLS('ppa_theme','light'),
  homeTrend:{metric:'new',unit:'cnt',drillYm:null,rangeFrom:null,rangeTo:null},
  clog:{q:'',kind:'',table:''},
};
/* dashboard_form.js(실시간 입력 서버가 붙어 있을 때만 로드됨)가 저장 후
   location.reload() 하기 전에 "지금 보고 있던 탭"을 세션스토리지에 남길 수
   있도록 이 살아있는 참조를 노출합니다 - 저장할 때마다 홈으로 튕기지 않고
   보던 화면 그대로 돌아오게 하기 위함입니다(restoreTabFromSession 참고). */
window.PPA_DASHBOARD_STATE=state;

/* ── 문자열 유틸 ────────────────────────────────────────────────────────── */
function esc(s){return String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function jsq(s){return String(s??'').replace(/\\/g,'\\\\').replace(/'/g,"\\'");}
function nf(n,d){return Number(n).toLocaleString('ko-KR',{maximumFractionDigits:d===undefined?1:d});}

/* ── 셀 표현: 숫자 서식, 불린 배지, 상태 색, 기한 경고, FK 이름 ─────────── */
const NUMCOL=/\(MW\)|\(원\/kWh\)|\(년\)|MGA/;
const DEADLINE_COL=/기한/;
const SOON_DAYS=180;
function isNumCol(c){return NUMCOL.test(c);}
function fmtVal(col,val){
  if(isNumCol(col)&&val!==''){const n=Number(val);if(!isNaN(n)) return nf(n,4);}
  return val;
}
function boolLabel(col,val){
  const on=val==='TRUE';
  if(col.indexOf('미확보')>=0) return on?'미확보':'확보완료';
  return on?'예':'아니오';
}
function daysBetween(a,b){return Math.round((new Date(b)-new Date(a))/86400000);}
function deadlineFlag(col,val){
  if(!DEADLINE_COL.test(col)||!/^\d{4}-\d{2}-\d{2}$/.test(val)||!TODAY) return null;
  const d=daysBetween(TODAY,val);
  if(d<0) return {cls:'exp',text:'만료'};
  if(d<=SOON_DAYS) return {cls:'soon',text:'D-'+d};
  return null;
}
/* 수급매칭 "현황" 8개 값 각각의 아이콘·심각도색 - 자유 텍스트 정규식 추측
   대신 실제 8개 값을 그대로 매핑합니다(정확한 문구는 dashboard_form.js의
   ENUM_COLUMNS["현황"]과 반드시 같은 목록으로 유지 - 입력 폼 드롭다운과
   화면 표시가 어긋나지 않도록). terminal:true는 "이 상태면 계약이 완전히
   끝난 것"이라는 뜻으로 용량 집계(sumCapSplit 등)에서 제외할 근거로만
   씁니다 - 배지 색상(cls)과는 별개 개념입니다(예: "착공 전"은 회색이지만
   종료된 건 아님 - 색만 보고 종료 여부를 판단하면 안 됨). */
const STATUS_META={
  '공급 중':{icon:'⚡',cls:'ok',terminal:false},
  '신고 중':{icon:'📝',cls:'info',terminal:false},
  '상업운전 개시':{icon:'🔌',cls:'ok',terminal:false},
  '공사 중':{icon:'🚧',cls:'warn',terminal:false},
  '착공 전':{icon:'📋',cls:'mute',terminal:false},
  '이슈 발생':{icon:'⚠️',cls:'no',terminal:false},
  '미확보':{icon:'❓',cls:'no',terminal:false},
  '공급종료':{icon:'🏁',cls:'mute',terminal:true},
};
function statusMeta(val){return STATUS_META[parseStatus(val).label]||null;}
function statusClass(val){const m=statusMeta(val);return m?m.cls:null;}
function isTerminalStatus(val){const m=statusMeta(val);return !!(m&&m.terminal);}
/* 8개 현황 각각의 현재 건수 - 홈 탭 "한눈에 보기"의 현황 아이콘 줄에 씀 */
function statusCountsAll(){
  const M=byKey['T_수급매칭'];
  const counts={};Object.keys(STATUS_META).forEach(k=>counts[k]=0);
  if(M) M.rows.forEach(r=>{
    const label=parseStatus(r.cells['현황']).label;
    if(counts[label]!==undefined) counts[label]++;
  });
  return counts;
}
/* ID만 봐서는 뭔지 알기 어려운 표는 FK로 참조될 때 이름을 같이 보여줍니다. */
const NAME_COLS={"T_발전소":["발전소명","발전법인명"],"T_수요기업":["기업명"],"T_전기사용지":["전기사용지명"]};
function displayNameFor(tableKey,pkVal){
  const cols=NAME_COLS[tableKey];if(!cols) return null;
  const row=rowIndex[tableKey]&&rowIndex[tableKey][String(pkVal)];if(!row) return null;
  const parts=cols.map(c=>row.cells[c]).filter(v=>v!==undefined&&v!=='');
  return parts.length?parts.join(' · '):null;
}
function cellHtml(t,col,rawVal){
  const val=rawVal??'';
  if(val==='') return '';
  if(val==='TRUE'||val==='FALSE')
    return `<span class="badge ${val==='TRUE'?'no':'ok'}">${esc(boolLabel(col,val))}</span>`;
  const isPk=col===t.pk,refKey=(t.fk||{})[col];
  if(isPk||refKey){
    const target=refKey||t.key;
    const link=`<a class="idlink" onclick="event.stopPropagation();jumpTo('${jsq(target)}','${jsq(val)}')">${esc(val)}</a>`;
    if(refKey){const nm=displayNameFor(refKey,val);if(nm) return link+`<span class="fkname">${esc(nm)}</span>`;}
    return link;
  }
  const flag=deadlineFlag(col,val);
  if(flag) return esc(val)+`<span class="datewarn ${flag.cls}">${esc(flag.text)}</span>`;
  if(col==='현황'){const m=statusMeta(val);if(m) return `<span class="badge ${m.cls}">${m.icon} ${esc(val)}</span>`;}
  return esc(fmtVal(col,val));
}

/* ── 컬럼 표시/숨김 (브라우저에 기억) ───────────────────────────────────── */
function isColHidden(t,col){return (state.hidden[t.key]||new Set()).has(col);}
function visibleColumns(t){return t.columns.filter(c=>!isColHidden(t,c));}
function toggleCol(tk,col){
  state.hidden[tk]=state.hidden[tk]||new Set();
  const s=state.hidden[tk];s.has(col)?s.delete(col):s.add(col);
  writeLS('ppa_hidden_'+tk,JSON.stringify([...s]));render();
}
function showAllCols(tk){state.hidden[tk]=new Set();writeLS('ppa_hidden_'+tk,'[]');render();}
function loadHidden(){
  DATA.tables.forEach(t=>{
    const raw=readLS('ppa_hidden_'+t.key,null);
    if(raw){try{state.hidden[t.key]=new Set(JSON.parse(raw));}catch(e){}}
  });
}

/* ── 필터 ───────────────────────────────────────────────────────────────── */
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
    return n>=2&&n<=25&&n<t.rows.length;
  });
}
function isDateCol(t,col){
  const s=t.rows.find(r=>r.cells[col]);
  return !!s&&/^\d{4}-\d{2}-\d{2}$/.test(String(s.cells[col]));
}
function setFilter(k,col,val){
  state.filters[k]=state.filters[k]||{};
  if(val) state.filters[k][col]=val; else delete state.filters[k][col];
  state.page[k]=1;render();
}
function setColQ(k,col,val){
  state.colQ[k]=state.colQ[k]||{};
  if(val) state.colQ[k][col]=val; else delete state.colQ[k][col];
  state.page[k]=1;render();
}
function setDateFilter(k,col,edge,val){
  state.dateFilters[k]=state.dateFilters[k]||{};
  state.dateFilters[k][col]=state.dateFilters[k][col]||{};
  state.dateFilters[k][col][edge]=val;
  state.page[k]=1;render();
}
function clearAllFilters(k){
  state.filters[k]={};state.colQ[k]={};state.dateFilters[k]={};state.pkFilter[k]=null;
  state.q[k]='';state.onlyErr[k]=false;state.onlyChg[k]=false;
  state.page[k]=1;render();
}
function activeFilterCount(k){
  let n=0;
  n+=Object.values(state.filters[k]||{}).filter(v=>v).length;
  n+=Object.values(state.colQ[k]||{}).filter(v=>v).length;
  Object.values(state.dateFilters[k]||{}).forEach(r=>{if(r&&(r.from||r.to)) n++;});
  if((state.q[k]||'').trim()) n++;
  if(state.onlyErr[k]) n++;
  if(state.onlyChg[k]) n++;
  if(state.pkFilter[k]&&state.pkFilter[k].length) n++;
  return n;
}
function filterDescription(k){
  const parts=[];
  const q=(state.q[k]||'').trim();
  if(q) parts.push('검색 "'+q+'"');
  Object.entries(state.filters[k]||{}).forEach(([c,v])=>{
    if(v) parts.push(c+'='+((v==='TRUE'||v==='FALSE')?boolLabel(c,v):v));});
  Object.entries(state.colQ[k]||{}).forEach(([c,v])=>{if(v) parts.push(c+'~"'+v+'"');});
  Object.entries(state.dateFilters[k]||{}).forEach(([c,r])=>{
    if(r&&(r.from||r.to)) parts.push(c+' '+(r.from||'')+'~'+(r.to||''));});
  if(state.onlyErr[k]) parts.push('검증오류만');
  if(state.onlyChg[k]) parts.push('변경분만');
  if(state.pkFilter[k]&&state.pkFilter[k].length) parts.push('조치 필요 항목 '+state.pkFilter[k].length+'건');
  return parts.length?parts.join(' · '):'전체';
}
/* 홈 탭 "우선순위 조치 필요 항목"처럼 단일 컬럼 값 비교로는 표현할 수 없는
   계산된 조건(용량 잔여·날짜 역전 등)의 결과를, 해당 표에 PK 목록으로
   그대로 필터링해서 보여주기 위한 범용 드릴다운. */
function jumpToPkSet(tk,pks){
  state.tab=tk;state.pkFilter[tk]=pks;state.page[tk]=1;render();
}
function clearPkFilter(k){state.pkFilter[k]=null;state.page[k]=1;render();}
/* jumpToPkSet(...)를 onclick 문자열 안에 그대로 박아 넣기 위한 JS 배열
   리터럴 문자열화 — onclick 속성 전체가 큰따옴표로 감싸여 있으므로 각 PK는
   작은따옴표 문자열로, jsq()로 이스케이프해서 넣습니다(다른 곳의 onclick
   패턴과 동일한 관례). */
function pkArrLiteral(pks){return '['+pks.map(p=>"'"+jsq(p)+"'").join(',')+']';}

/* 여러 컬럼에 걸쳐 검색어를 매칭 — 공백으로 구분된 각 낱말이 대상 텍스트
   어딘가에 전부(AND) 있으면 매치로 봅니다. "영광 풍력"처럼 띄어 쓴 검색어가
   "영광풍력1호"처럼 붙어 있는 값과도 매치되고, 검색어 전체가 한 컬럼 안에
   고스란히 들어있어야 하던 예전 제약(다중 컬럼에 걸친 값은 못 찾던 문제)도
   없앱니다. */
function matchesSearch(hayText,q){
  const terms=String(q||'').trim().split(/\s+/).filter(Boolean);
  if(!terms.length) return true;
  const hay=String(hayText).toLowerCase();
  return terms.every(t=>hay.includes(t.toLowerCase()));
}
/* 검색창 input 핸들러 — 조합 중(한글 등 IME)에는 절대 렌더링하지 않습니다.
   렌더링은 #view.innerHTML을 통째로 새로 그리는데, 이게 조합 중에 일어나면
   입력창 DOM이 재생성되며 브라우저의 IME 조합 세션이 끊겨 한글이 자모로
   쪼개지거나 마지막 글자가 씹히는 등으로 깨져 보입니다. 타이핑이 잠시
   멈췄을 때만(디바운스) 실제로 상태를 반영해, 대량 데이터에서 매 키 입력마다
   표 전체를 다시 그리는 렉도 함께 줄입니다.

   타이머는 반드시 "입력창 하나마다 하나씩"(e.target에 매달아 둠) 둬야
   합니다 — 예전에는 타이머 변수 하나를 모든 검색창이 같이 썼는데, 그러면
   예를 들어 표 검색창에 타이핑한 직후(140ms 안에) 같은 화면의 "컬럼별
   찾기" 검색창에 이어서 타이핑하면 표 검색창 쪽 타이머가 취소돼버려서
   방금 입력한 검색어가 조용히 사라지는 버그가 있었습니다(전역 검색·표
   검색·컬럼별 찾기·탐색 결과 검색이 전부 한 화면에 같이 있을 수 있는
   구조라 실제로 자주 걸림). */
function onSearchType(e,setter){
  if(e.isComposing) return;
  const el=e.target;
  const v=el.value;
  clearTimeout(el._searchTimer);
  el._searchTimer=setTimeout(()=>setter(v),140);
}

/* 필터·검색·정렬을 모두 적용한 행 목록 (페이징 전) — 화면과 내려받기가 공유 */
function filteredRows(t){
  const k=t.key;
  const q=(state.q[k]||'').trim();
  let rows=t.rows.map((r,i)=>({r,i}));
  if(q) rows=rows.filter(({r})=>matchesSearch(t.columns.map(c=>r.cells[c]??'').join(' '),q));
  Object.entries(state.filters[k]||{}).forEach(([col,val])=>{
    if(val) rows=rows.filter(({r})=>String(r.cells[col]??'')===val);});
  Object.entries(state.colQ[k]||{}).forEach(([col,val])=>{
    const s=String(val).trim().toLowerCase();
    if(s) rows=rows.filter(({r})=>String(r.cells[col]??'').toLowerCase().includes(s));});
  Object.entries(state.dateFilters[k]||{}).forEach(([col,range])=>{
    if(!range||(!range.from&&!range.to)) return;
    rows=rows.filter(({r})=>{
      const v=String(r.cells[col]||'');
      if(!v) return false;
      if(range.from&&v<range.from) return false;
      if(range.to&&v>range.to) return false;
      return true;});});
  if(state.onlyErr[k]) rows=rows.filter(({r})=>(r.error_cols||[]).length>0);
  if(state.onlyChg[k]) rows=rows.filter(({r})=>!!r.change);
  if(state.pkFilter[k]&&state.pkFilter[k].length){
    const pkSet=new Set(state.pkFilter[k]);
    rows=rows.filter(({r})=>pkSet.has(String(r.cells[t.pk]??'')));
  }
  const sc=state.sort[k];
  if(sc) rows=[...rows].sort((a,b)=>{
    const av=a.r.cells[sc.key]??'',bv=b.r.cells[sc.key]??'';
    if(av===''&&bv!=='') return 1;
    if(bv===''&&av!=='') return -1;
    const an=Number(av),bn=Number(bv);
    const cmp=(av!==''&&bv!==''&&!isNaN(an)&&!isNaN(bn))?an-bn:String(av).localeCompare(String(bv),'ko');
    return sc.dir*cmp;});
  return rows;
}
function doSort(k,c){
  const s=state.sort[k];
  state.sort[k]=(s&&s.key===c)?{key:c,dir:-s.dir}:{key:c,dir:1};
  state.page[k]=1;render();
}
function setQ(k,v){state.q[k]=v;state.page[k]=1;render();}
function toggleOnlyErr(k){state.onlyErr[k]=!state.onlyErr[k];state.page[k]=1;render();}
function toggleOnlyChg(k){state.onlyChg[k]=!state.onlyChg[k];state.page[k]=1;render();}
function setPage(k,p){state.page[k]=p;render();
  const el=document.getElementById('tbl-'+k);if(el) el.scrollTop=0;}
function setPageSize(v){
  state.pageSize=Number(v)||50;writeLS('ppa_pagesize',String(state.pageSize));
  Object.keys(state.page).forEach(k=>state.page[k]=1);render();
}

/* ── 공통 조각 ──────────────────────────────────────────────────────────── */
/* 용어 설명 아이콘 — 네이티브 title 속성만 쓰는 가벼운 툴팁(외부 라이브러리
   없이 브라우저 기본 동작으로 마우스 오버 시 설명이 뜸). */
const infoTip=(text)=>`<span class="infotip" title="${esc(text)}" tabindex="0">ⓘ</span>`;
const kpi=(k,v,s,cls,onclick)=>
  `<div class="kpi${cls?' '+cls:''}${onclick?' clickable':''}"${onclick?` onclick="${onclick}"`:''}>
     <span class="kk">${k}</span><span class="kv">${v}</span><span class="ks">${s||''}</span></div>`;
const panel=(t,s,inner)=>`<div class="panel"><div class="ph"><h3>${t}</h3><span class="sub">${s||''}</span></div>${inner}</div>`;

function pager(t,total,page,pages){
  if(pages<=1&&total<=state.pageSize)
    return `<div class="pager"><span class="pginfo">${nf(total,0)}건 전체 표시</span>
      <select class="filtersel" onchange="setPageSize(this.value)">${pageSizeOpts()}</select></div>`;
  const btn=(p,label,dis,on)=>
    `<button class="pgbtn${on?' on':''}" ${dis?'disabled':''} onclick="setPage('${jsq(t.key)}',${p})">${label}</button>`;
  const nums=[];
  const win=2;
  let last=0;
  for(let p=1;p<=pages;p++){
    if(p===1||p===pages||(p>=page-win&&p<=page+win)){
      if(last&&p-last>1) nums.push('<span class="pgdots">…</span>');
      nums.push(btn(p,String(p),false,p===page));last=p;
    }
  }
  const from=(page-1)*state.pageSize+1,to=Math.min(page*state.pageSize,total);
  return `<div class="pager">
    ${btn(1,'«',page===1)}${btn(page-1,'‹',page===1)}${nums.join('')}${btn(page+1,'›',page===pages)}${btn(pages,'»',page===pages)}
    <span class="pginfo">${nf(from,0)}–${nf(to,0)} / ${nf(total,0)}건</span>
    <select class="filtersel" onchange="setPageSize(this.value)">${pageSizeOpts()}</select></div>`;
}
function pageSizeOpts(){
  return [25,50,100,200,1000].map(n=>
    `<option value="${n}" ${n===state.pageSize?'selected':''}>${n===1000?'1000 (전체)':n+'개씩'}</option>`).join('');
}

/* ── 표 화면 ────────────────────────────────────────────────────────────── */
function filterBar(t){
  const k=t.key;
  const cols=filterableColumns(t);
  const cur=state.filters[k]||{};
  const selects=cols.map(c=>{
    const sel=cur[c]||'';
    const opts=[`<option value="">${esc(c)}: 전체</option>`].concat(
      uniqueValues(t,c).map(v=>
        `<option value="${esc(v)}" ${v===sel?'selected':''}>${esc((v==='TRUE'||v==='FALSE')?boolLabel(c,v):v)}</option>`));
    return `<select class="filtersel" onchange="setFilter('${jsq(k)}','${jsq(c)}',this.value)">${opts.join('')}</select>`;
  }).join('');
  const dcols=t.columns.filter(c=>isDateCol(t,c));
  const df=state.dateFilters[k]||{};
  const dates=dcols.map(c=>{
    const r=df[c]||{};
    return `<span class="datef"><span class="dflabel">${esc(c)}</span>
      <input type="date" id="df-${esc(k)}-${esc(c)}-from" value="${esc(r.from||'')}" onchange="setDateFilter('${jsq(k)}','${jsq(c)}','from',this.value)">
      <span class="dfsep">~</span>
      <input type="date" id="df-${esc(k)}-${esc(c)}-to" value="${esc(r.to||'')}" onchange="setDateFilter('${jsq(k)}','${jsq(c)}','to',this.value)"></span>`;
  }).join('');
  const chips=[];
  Object.entries(cur).forEach(([col,val])=>{if(!val) return;
    chips.push(`<span class="fchip">${esc(col)}: ${esc((val==='TRUE'||val==='FALSE')?boolLabel(col,val):val)}<button onclick="setFilter('${jsq(k)}','${jsq(col)}','')">✕</button></span>`);});
  Object.entries(state.colQ[k]||{}).forEach(([col,val])=>{if(!val) return;
    chips.push(`<span class="fchip">${esc(col)} ~ "${esc(val)}"<button onclick="setColQ('${jsq(k)}','${jsq(col)}','')">✕</button></span>`);});
  Object.entries(df).forEach(([col,r])=>{if(!r||(!r.from&&!r.to)) return;
    chips.push(`<span class="fchip">${esc(col)} ${esc(r.from||'')}~${esc(r.to||'')}<button onclick="setDateFilter('${jsq(k)}','${jsq(col)}','from','');setDateFilter('${jsq(k)}','${jsq(col)}','to','')">✕</button></span>`);});
  if(state.pkFilter[k]&&state.pkFilter[k].length)
    chips.push(`<span class="fchip">홈의 조치 필요 항목에서 이동 (${nf(state.pkFilter[k].length,0)}건)<button onclick="clearPkFilter('${jsq(k)}')">✕</button></span>`);
  const chipRow=chips.length?`<div class="chiprow">${chips.join('')}</div>`:'';
  const bar=(selects||dates)?`<div class="filterbar">${selects}${dates}</div>`:'';
  return bar+chipRow;
}

function tableView(t){
  const k=t.key;
  const raw=state.q[k]||'';
  const rows=filteredRows(t);
  const total=rows.length;
  const pages=Math.max(1,Math.ceil(total/state.pageSize));
  const page=Math.min(Math.max(1,state.page[k]||1),pages);
  state.page[k]=page;
  const pageRows=rows.slice((page-1)*state.pageSize,page*state.pageSize);
  const cols=visibleColumns(t);
  const sc=state.sort[k];
  const errN=t.rows.filter(r=>(r.error_cols||[]).length>0).length;
  const chgN=t.rows.filter(r=>!!r.change).length;
  const hiddenN=(state.hidden[k]||new Set()).size;
  const nFilters=activeFilterCount(k);

  const head=cols.map(c=>{
    const ar=sc&&sc.key===c?`<span class="ar">${sc.dir>0?'▲':'▼'}</span>`:'';
    const pk=c===t.pk?'<span class="pkbadge">PK</span>':'';
    const fk=(t.fk||{})[c]?'<span class="fkbadge">FK</span>':'';
    return `<th class="${isNumCol(c)?'num':''}" onclick="doSort('${jsq(k)}','${jsq(c)}')">${esc(c)}${pk}${fk}${ar}</th>`;
  }).join('');
  const body=pageRows.map(({r,i})=>{
    const err=(r.error_cols||[]).length>0;
    const cls=err?'rowerr':(r.change==='added'?'rowadd':(r.change==='changed'?'rowchg':''));
    const tds=cols.map(c=>{
      const bad=(r.error_cols||[]).includes(c);
      const chg=(r.changed_cols||[]).includes(c);
      const title=chg&&r.prev?` title="이전값: ${esc(r.prev[c]||'(공란)')}"`:'';
      return `<td class="${bad?'cellerr':(chg?'cellchg':'')}${isNumCol(c)?' num':''}"${title}>${cellHtml(t,c,r.cells[c])}</td>`;
    }).join('');
    return `<tr class="clickrow ${cls}" onclick="openDetail('${jsq(k)}',${i})">${tds}</tr>`;
  }).join('')||`<tr><td class="emptyrow" colspan="${cols.length}">조건에 맞는 데이터가 없습니다.</td></tr>`;

  const colOpts=t.columns.filter(c=>c!==t.pk).map(c=>
    `<label class="colopt"><input type="checkbox" ${isColHidden(t,c)?'':'checked'} onchange="toggleCol('${jsq(k)}','${jsq(c)}')">${esc(c)}</label>`).join('');
  const colSearchOpts=t.columns.map(c=>
    `<label class="colopt" style="display:block"><span class="dkey">${esc(c)}</span>
      <input id="cq-${esc(k)}-${esc(c)}" class="search" style="width:100%;min-width:0;margin-top:3px"
        value="${esc((state.colQ[k]||{})[c]||'')}" placeholder="이 컬럼에서 찾기…"
        oninput="onSearchType(event,v=>setColQ('${jsq(k)}','${jsq(c)}',v))"></label>`).join('');

  return `<div class="toolbar">
      <input id="search-${esc(k)}" class="search" placeholder="${esc(t.label)} 전체 컬럼 검색…" value="${esc(raw)}" oninput="onSearchType(event,v=>setQ('${jsq(k)}',v))">
      ${errN?`<button class="btn ${state.onlyErr[k]?'on':''}" onclick="toggleOnlyErr('${jsq(k)}')">검증오류만 ${errN}</button>`:''}
      ${chgN?`<button class="btn ${state.onlyChg[k]?'on':''}" onclick="toggleOnlyChg('${jsq(k)}')">변경분만 ${chgN}</button>`:''}
      <details class="drop"><summary>컬럼별 찾기</summary><div class="dropbody">${colSearchOpts}</div></details>
      <details class="drop"><summary>컬럼 선택${hiddenN?' ('+hiddenN+' 숨김)':''}</summary>
        <div class="dropbody">${colOpts}<button class="btn" style="width:100%;margin-top:8px" onclick="showAllCols('${jsq(k)}')">전체 표시</button></div></details>
      <details class="drop"><summary>내려받기</summary><div class="dropbody right">${downloadMenu(k)}</div></details>
      ${nFilters?`<button class="btn danger" onclick="clearAllFilters('${jsq(k)}')">조건 초기화 (${nFilters})</button>`:''}
      <span class="count">${nf(total,0)} / ${nf(t.rows.length,0)}건</span>
    </div>
    ${filterBar(t)}
    <div class="tbl-wrap stickyfirst" id="tbl-${esc(k)}"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>
    ${pager(t,total,page,pages)}`;
}

function tData(t){
  const errRows=t.rows.filter(r=>(r.error_cols||[]).length>0).length;
  const chgRows=t.rows.filter(r=>!!r.change).length;
  const s=(CHANGES.summary||{})[t.key]||{};
  return `<section>${printHead(t.label+' 목록',filterDescription(t.key))}<div class="kpis">
    ${kpi('전체 행 수',nf(t.rows.length,0),t.label,'accent')}
    ${kpi('검증 오류 행',nf(errRows,0),errRows>0?'클릭 시 오류만 보기':'문제 없음',errRows>0?'warn':'',errRows>0?`toggleOnlyErr('${jsq(t.key)}')`:'')}
    ${kpi('지난 대비 변경',CHANGES.has_prev?nf(chgRows,0):'—',CHANGES.has_prev?`추가 ${s.added||0} · 수정 ${s.changed||0} · 삭제 ${s.removed||0}`:'이전 스냅샷 없음',
      '',chgRows>0?`toggleOnlyChg('${jsq(t.key)}')`:'')}
    ${kpi('컬럼 수',t.columns.length,(state.hidden[t.key]||new Set()).size?((state.hidden[t.key]||new Set()).size+'개 숨김'):'전체 표시')}
    </div>
    ${panel(t.label+' 전체 목록','행 클릭 시 상세(수정/삭제 가능) · PK/FK 클릭 시 관계조회 · 머리글 클릭 시 정렬',
      `<div style="margin-bottom:10px"><button class="btn primary" onclick="openInlineAdd('${jsq(t.key)}')">+ 새 ${esc(t.label)} 추가</button></div>`
      +tableView(t))}</section>`;
}

/* ── 상세 모달 ──────────────────────────────────────────────────────────── */
function openDetail(tk,i){state.modal={t:tk,i:i};render();}
function openRemovedDetail(tk,idx){state.modal={t:tk,removed:idx};render();}
function closeDetail(){state.modal=null;render();}
function modalHtml(){
  if(!state.modal) return '';
  const t=byKey[state.modal.t];if(!t) return '';
  if(state.modal.removed!==undefined) return removedModalHtml(t,state.modal.removed);
  const r=t.rows[state.modal.i];if(!r) return '';
  const pkv=r.cells[t.pk];
  /* 모달에서는 "이 레코드에 바로 붙어 있는 것"만 — 1단계 직접 연결 */
  const chain=relatedChain(t.key,pkv,1);
  const rel=DATA.tables.map(x=>{
    const n=chain[x.key]?chain[x.key].size:0;
    return n>0&&x.key!==t.key?`<span class="chipcount">${esc(x.label)} ${nf(n,0)}</span>`:'';}).join('');
  const rows=t.columns.map(c=>{
    const bad=(r.error_cols||[]).includes(c);
    const chg=(r.changed_cols||[]).includes(c);
    const prev=chg&&r.prev?`<div class="chgval"><span class="chgold">${esc(r.prev[c]||'(공란)')}</span>→ <span class="chgnew">${esc(r.cells[c]||'(공란)')}</span></div>`:'';
    return `<div class="drow ${bad?'err':''} ${chg?'chg':''}">
      <span class="dkey">${esc(c)}${c===t.pk?'<span class="pkbadge">PK</span>':''}${(t.fk||{})[c]?'<span class="fkbadge">FK</span>':''}</span>
      <span>${cellHtml(t,c,r.cells[c])||'<span style="color:var(--sub)">(공란)</span>'}${prev}</span></div>`;
  }).join('');
  const errs=(r.error_cols||[]).length
    ? `<div class="badge no" style="margin-top:8px">검증 오류: ${esc((r.error_cols||[]).join(', '))}</div>`:'';
  const chgb=r.change==='added'?'<div class="badge ok" style="margin-top:8px">지난 생성분 이후 추가된 항목</div>'
    :(r.change==='changed'?`<div class="badge info" style="margin-top:8px">지난 생성분 대비 ${(r.changed_cols||[]).length}개 항목 변경</div>`:'');
  return `<div class="backdrop" onclick="if(event.target===this)closeDetail()">
    <div class="modal" role="dialog" aria-modal="true">
      <div class="modalhead">
        <div><h3>${esc(t.label)} 상세</h3>
          <div class="sub">${esc(String(pkv||'(PK 공란)'))}</div>${errs}${chgb}
          <div style="margin-top:8px">${rel||'<span class="nocand">연결된 다른 표 데이터 없음</span>'}</div></div>
        <button class="iconbtn" onclick="closeDetail()" aria-label="닫기">✕</button>
      </div>
      <div class="modalbody">${rows}</div>
      <div class="modalfoot">
        <button class="btn primary" onclick="openInlineEdit('${jsq(t.key)}','${jsq(pkv)}')">수정</button>
        <button class="btn danger" onclick="openInlineDelete('${jsq(t.key)}','${jsq(pkv)}')">삭제</button>
        <button class="btn" onclick="closeDetail();jumpTo('${jsq(t.key)}','${jsq(pkv)}')">관계조회로 보기</button>
        <button class="btn" onclick="togglePin('${jsq(t.key)}','${jsq(pkv)}')">${isPinned(t.key,pkv)?'비교에서 제거':'비교에 추가'}</button>
        <button class="btn" onclick="copyText(detailText('${jsq(t.key)}',${state.modal.i}))">내용 복사</button>
        <button class="btn" onclick="closeDetail()">닫기</button>
      </div></div></div>`;
}
function detailText(tk,i){
  const t=byKey[tk],r=t.rows[i];
  return `[${t.label}] ${r.cells[t.pk]||''}\n`+
    t.columns.map(c=>'- '+c+': '+(r.cells[c]||'')).join('\n');
}
/* 삭제된 행 상세 — 지금 DATA에는 없는(사라진) 행이라 t.rows가 아니라
   스냅샷 비교 결과(CHANGES.removed_rows)에서 그대로 읽어옵니다. 스냅샷
   파일에 저장돼 있던 모든 컬럼 값이 여기서 다 보여야 합니다. */
function removedModalHtml(t,idx){
  const cells=((CHANGES.removed_rows||{})[t.key]||[])[idx];
  if(!cells) return '';
  const pkv=cells[t.pk];
  const rows=t.columns.map(c=>`<div class="drow">
      <span class="dkey">${esc(c)}${c===t.pk?'<span class="pkbadge">PK</span>':''}${(t.fk||{})[c]?'<span class="fkbadge">FK</span>':''}</span>
      <span>${cellHtml(t,c,cells[c])||'<span style="color:var(--sub)">(공란)</span>'}</span></div>`).join('');
  return `<div class="backdrop" onclick="if(event.target===this)closeDetail()">
    <div class="modal" role="dialog" aria-modal="true">
      <div class="modalhead">
        <div><h3>${esc(t.label)} 상세</h3>
          <div class="sub">${esc(String(pkv||'(PK 공란)'))}</div>
          <div class="badge mute" style="margin-top:8px">지난 생성분에는 있었지만 지금은 삭제된 항목 — 스냅샷에 저장된 마지막 값입니다</div></div>
        <button class="iconbtn" onclick="closeDetail()" aria-label="닫기">✕</button>
      </div>
      <div class="modalbody">${rows}</div>
      <div class="modalfoot">
        <button class="btn" onclick="copyText(removedDetailText('${jsq(t.key)}',${idx}))">내용 복사</button>
        <button class="btn" onclick="closeDetail()">닫기</button>
      </div></div></div>`;
}
function removedDetailText(tk,idx){
  const t=byKey[tk],cells=((CHANGES.removed_rows||{})[tk]||[])[idx];
  return `[${t.label} · 삭제됨] ${cells[t.pk]||''}\n`+
    t.columns.map(c=>'- '+c+': '+(cells[c]||'')).join('\n');
}

/* ── 관계조회 ───────────────────────────────────────────────────────────────
   FK를 따라 양방향(참조하는 쪽·참조받는 쪽)으로 퍼져나가며 연결된 레코드를
   모읍니다. 다만 이 스키마는 한 전기사용지가 여러 구매계약에서 공급받고
   한 구매계약이 여러 전기사용지에 공급하는 다대다 구조라, 제한 없이 퍼뜨리면
   실데이터에서 사실상 전체가 하나로 이어져 수백 건이 딸려옵니다. 그래서
   "몇 단계까지 따라갈지"를 사용자가 고를 수 있게 하고(기본 2단계), 안전장치로
   총 노드 수 상한도 둡니다. */
const CHAIN_NODE_CAP=4000;
function relatedChain(startTable,startPk,maxDepth){
  const depth=(maxDepth===undefined)?(state.lookupDepth||0):maxDepth;
  const visited={},startKey=String(startPk);
  visited[startTable]=new Set([startKey]);
  let count=1;
  const queue=[[startTable,startKey,0]];
  const seen=k=>visited[k]||(visited[k]=new Set());
  while(queue.length){
    const [tk,pkv,d]=queue.shift();
    if(depth>0&&d>=depth) continue;
    if(count>=CHAIN_NODE_CAP) break;
    const row=rowIndex[tk]&&rowIndex[tk][pkv];
    if(!row) continue;
    const t=byKey[tk];
    /* 이 행이 참조하는 부모(FK → PK) */
    Object.entries(t.fk||{}).forEach(([col,refKey])=>{
      const v=row.cells[col];
      if(v!==undefined&&v!==''){
        const rk=String(v),s=seen(refKey);
        if(!s.has(rk)){s.add(rk);count++;queue.push([refKey,rk,d+1]);}
      }});
    /* 이 행을 참조하는 자식(PK ← FK) */
    edges.filter(e=>e.ref===tk).forEach(e=>{
      const childT=byKey[e.table],s=seen(e.table);
      childT.rows.forEach(cr=>{
        if(String(cr.cells[e.col]??'')===pkv){
          const cpk=cr.cells[childT.pk];
          if(cpk!==undefined&&cpk!==''){
            const ck=String(cpk);
            if(!s.has(ck)){s.add(ck);count++;queue.push([e.table,ck,d+1]);}
          }}});});
  }
  return visited;
}
function setLookupDepth(v){state.lookupDepth=Number(v);writeLS('ppa_depth',String(v));render();}
function miniTable(t,rows,highlightPk){
  const cols=visibleColumns(t);
  const head=cols.map(c=>{
    const pk=c===t.pk?'<span class="pkbadge">PK</span>':'';
    const fk=(t.fk||{})[c]?'<span class="fkbadge">FK</span>':'';
    return `<th class="nosort ${isNumCol(c)?'num':''}">${esc(c)}${pk}${fk}</th>`;}).join('');
  const body=rows.map(({r,i})=>{
    const isRoot=highlightPk!==null&&String(r.cells[t.pk])===String(highlightPk);
    const err=(r.error_cols||[]).length>0;
    const cls=isRoot?'rootrow':(err?'rowerr':(r.change?'rowchg':''));
    const tds=cols.map(c=>{
      const bad=(r.error_cols||[]).includes(c);
      const chg=(r.changed_cols||[]).includes(c);
      return `<td class="${bad?'cellerr':(chg?'cellchg':'')}${isNumCol(c)?' num':''}">${cellHtml(t,c,r.cells[c])}</td>`;}).join('');
    return `<tr class="clickrow ${cls}" onclick="openDetail('${jsq(t.key)}',${i})">${tds}</tr>`;}).join('');
  return `<div class="tbl-wrap stickyfirst"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`;
}
function setLookupTable(k){state.lookupTable=k;state.lookupQ='';render();}
function setLookupQ(v){state.lookupQ=v;render();}
function setLookup(table,pk){
  state.lookup={table,pk};
  const t=byKey[table],row=rowIndex[table]&&rowIndex[table][String(pk)];
  const short=row?t.columns.slice(0,2).map(c=>row.cells[c]).filter(v=>v!==undefined&&v!=='').join(' · '):String(pk);
  state.recent=state.recent.filter(r=>!(r.table===table&&String(r.pk)===String(pk)));
  state.recent.unshift({table,pk,label:t.label+' · '+short});
  state.recent=state.recent.slice(0,5);
}
function chooseLookup(table,pk){setLookup(table,pk);state.lookupQ='';render();}
function clearLookup(){state.lookup=null;state.lookupQ='';render();}
function jumpTo(table,pk){state.tab='관계조회';state.lookupTable=table;setLookup(table,pk);state.modal=null;render();}
function isPinned(table,pk){return state.pinned.some(p=>p.table===table&&String(p.pk)===String(pk));}
function togglePin(table,pk){
  const i=state.pinned.findIndex(p=>p.table===table&&String(p.pk)===String(pk));
  if(i>=0) state.pinned.splice(i,1);
  else{state.pinned.push({table,pk});if(state.pinned.length>3) state.pinned.shift();}
  render();
}
function tLookup(){
  const activeT=byKey[state.lookupTable];
  const pickTabs=DATA.tables.map(t=>
    `<button class="subtab ${t.key===state.lookupTable?'on':''}" onclick="setLookupTable('${jsq(t.key)}')">${esc(t.label)}</button>`).join('');
  const q=(state.lookupQ||'').trim();
  const matches=q?activeT.rows.filter(r=>matchesSearch(activeT.columns.map(c=>r.cells[c]??'').join(' '),q)):activeT.rows;
  const cand=matches.slice(0,25).map(r=>{
    const pkv=r.cells[activeT.pk];
    const label=activeT.columns.slice(0,3).map(c=>r.cells[c]).filter(v=>v!==undefined&&v!=='').join(' · ');
    return `<button class="candrow" onclick="chooseLookup('${jsq(activeT.key)}','${jsq(pkv)}')">${esc(label)}</button>`;
  }).join('')||'<div class="nocand">일치하는 항목이 없습니다.</div>';
  const more=matches.length>25?`<div class="nocand">${nf(matches.length-25,0)}건 더 있음 — 검색어를 추가해 좁혀보세요.</div>`:'';
  const recent=state.recent.length
    ? panel('최근 조회','',`<div class="candlist" style="max-height:none">${state.recent.map(r=>
        `<button class="candrow" onclick="chooseLookup('${jsq(r.table)}','${jsq(r.pk)}')">${esc(r.label)}</button>`).join('')}</div>`):'';
  const picker=`<div class="panel"><div class="ph"><h3>1. 표 선택</h3></div>
    <div class="subtabbar">${pickTabs}</div>
    <div class="ph" style="margin-top:16px"><h3>2. 검색해서 선택</h3><span class="sub">비워두면 목록이 그대로 보입니다</span></div>
    <input id="lookupSearchInput" class="search" style="width:100%" placeholder="${esc(activeT.label)} 검색 (ID, 이름 등)…" value="${esc(state.lookupQ||'')}" oninput="onSearchType(event,setLookupQ)">
    <div class="candlist">${cand}</div>${more}</div>`;

  if(!state.lookup) return `<section>${recent}${picker}</section>`;

  const rootT=byKey[state.lookup.table];
  const rootRow=rowIndex[state.lookup.table]&&rowIndex[state.lookup.table][String(state.lookup.pk)];
  if(!rootRow) return `<section><div class="panel lookuphead">
      <div><div class="ph" style="margin-bottom:4px"><h3>조회 중: ${esc(rootT.label)} · ${esc(state.lookup.pk)}</h3></div>
      <span class="badge no">이 ID를 가진 실제 레코드가 없습니다 — 오타나 참조 오류일 수 있습니다 (검증 탭 확인).</span></div>
      <button class="btn danger" onclick="clearLookup()">다른 항목 조회</button></div></section>`;

  const chain=relatedChain(state.lookup.table,state.lookup.pk);
  const rootLabel=rootT.columns.slice(0,3).map(c=>rootRow.cells[c]).filter(v=>v!==undefined&&v!=='').join(' · ');
  const CAP=50;
  const chainRows=DATA.tables.map(t=>{
    const set=chain[t.key];
    const rows=set&&set.size?t.rows.map((r,i)=>({r,i})).filter(({r})=>set.has(String(r.cells[t.pk]))):[];
    return {t,rows};});
  const summary=chainRows.map(({t,rows})=>rows.length?`<span class="chipcount">${esc(t.label)} ${nf(rows.length,0)}</span>`:'').join('');
  const panels=chainRows.map(({t,rows})=>{
    if(!rows.length) return '';
    const shown=rows.slice(0,CAP);
    const more=rows.length>CAP
      ? `<div class="nocand">많아서 ${CAP}건만 표시했습니다 (전체 ${nf(rows.length,0)}건).
         단계를 줄이거나 <a class="idlink" onclick="state.tab='${jsq(t.key)}';render()">${esc(t.label)} 탭</a>에서 조건을 걸어 확인하세요.</div>`:'';
    return panel(`${esc(t.label)} (${nf(rows.length,0)}건 연결)`,'',
      miniTable(t,shown,state.lookup.table===t.key?state.lookup.pk:null)+more);}).join('');
  const depthSel=`<span class="datef"><span class="dflabel">연결 범위</span>
      <select class="filtersel" style="border:none;background:none;padding:0" onchange="setLookupDepth(this.value)"
        title="FK를 몇 단계까지 따라갈지 — 넓힐수록 간접적으로 이어진 레코드까지 딸려옵니다">
      ${[[1,'직접 연결만'],[2,'2단계'],[3,'3단계'],[0,'전체 (제한 없음)']].map(([v,l])=>
        `<option value="${v}" ${Number(state.lookupDepth)===v?'selected':''}>${l}</option>`).join('')}</select></span>`;
  return `<section>${printHead('관계조회 — '+rootT.label+' '+rootLabel,'연결된 전 표 데이터')}
    <div class="panel lookuphead">
      <div><div class="ph" style="margin-bottom:4px"><h3>조회 중: ${esc(rootT.label)} · ${esc(rootLabel)}</h3></div>${summary}</div>
      <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center">
        ${depthSel}
        <button class="btn primary" onclick="togglePin('${jsq(state.lookup.table)}','${jsq(state.lookup.pk)}')">${isPinned(state.lookup.table,state.lookup.pk)?'비교에서 제거':'비교에 추가 ('+state.pinned.length+'/3)'}</button>
        <button class="btn" onclick="window.print()">인쇄 / PDF</button>
        <button class="btn danger" onclick="clearLookup()">다른 항목 조회</button>
      </div></div>${panels}</section>`;
}

/* ── 비교 ───────────────────────────────────────────────────────────────── */
function tCompare(){
  if(!state.pinned.length)
    return `<section><div class="panel"><div class="ph"><h3>비교할 레코드가 없습니다</h3></div>
      <div class="nocand">표에서 행을 눌러 상세를 열거나 관계조회에서 레코드를 연 뒤 "비교에 추가"를 누르면
      여기서 최대 3건까지 나란히 비교할 수 있습니다.</div></div></section>`;
  const groups={};
  state.pinned.forEach(p=>{(groups[p.table]=groups[p.table]||[]).push(p.pk);});
  const panels=Object.entries(groups).map(([tk,pks])=>{
    const t=byKey[tk];
    const rows=pks.map(pk=>rowIndex[tk]&&rowIndex[tk][String(pk)]).filter(Boolean);
    if(!rows.length) return '';
    const cols=visibleColumns(t);
    const head='<th class="nosort">항목</th>'+rows.map(r=>
      `<th class="nosort">${esc(r.cells[t.pk])}<button class="iconbtn" style="width:18px;height:18px;font-size:10px;display:inline-flex;margin-left:6px" onclick="togglePin('${jsq(tk)}','${jsq(r.cells[t.pk])}')">✕</button></th>`).join('');
    const body=cols.map(c=>{
      const differ=new Set(rows.map(r=>String(r.cells[c]??''))).size>1;
      const tds=rows.map(r=>`<td class="${differ?'diffcell':''}${isNumCol(c)?' num':''}">${cellHtml(t,c,r.cells[c])}</td>`).join('');
      const pk=c===t.pk?'<span class="pkbadge">PK</span>':'';
      const fk=(t.fk||{})[c]?'<span class="fkbadge">FK</span>':'';
      return `<tr><td><span class="dkey">${esc(c)}</span>${pk}${fk}</td>${tds}</tr>`;}).join('');
    return panel(`${esc(t.label)} 비교 (${rows.length}건)`,'값이 다른 항목은 강조 표시됩니다',
      `<div class="tbl-wrap stickyfirst"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`);
  }).join('');
  return `<section>${printHead('레코드 비교','핀 '+state.pinned.length+'건')}
    <div class="panel lookuphead"><div class="ph" style="margin-bottom:0"><h3>비교 중 (${state.pinned.length}건)</h3></div>
      <div style="display:flex;gap:8px"><button class="btn" onclick="window.print()">인쇄 / PDF</button>
      <button class="btn danger" onclick="state.pinned=[];render()">전체 해제</button></div></div>${panels}</section>`;
}

/* ── 관계형 탐색 (기준 표 · 출력 컬럼 직접 선택 + 누락 데이터) ─────────────
   엑셀에서 VLOOKUP을 여러 번 걸어 만들던 "여러 표를 합친 목록"을 화면에서
   바로 만듭니다. 기준 표를 정하고 연결할 표를 고르면 FK 경로를 따라 조인하며,
   연결 상대가 없는 행도 빈칸으로 남겨두기 때문에(LEFT JOIN) "구매계약이 없는
   발전소" 같은 누락 데이터를 그대로 찾아낼 수 있습니다. */

/* 표 사이 연결 그래프 — 부모(참조 대상)·자식(참조하는 쪽) 양방향 */
const adj={};
DATA.tables.forEach(t=>adj[t.key]=[]);
edges.forEach(e=>{
  adj[e.table].push({to:e.ref,dir:'parent',col:e.col});
  adj[e.ref].push({to:e.table,dir:'child',col:e.col});
});
/* 기준 표에서 각 표까지의 최단 경로(거쳐야 하는 표 포함) */
function joinPaths(base){
  const prev={},dist={};dist[base]=0;
  const q=[base];
  while(q.length){
    const cur=q.shift();
    (adj[cur]||[]).forEach(e=>{
      if(dist[e.to]===undefined){
        dist[e.to]=dist[cur]+1;prev[e.to]={from:cur,dir:e.dir,col:e.col};q.push(e.to);
      }});
  }
  return {prev,dist};
}
/* FK 값 → 자식 행들 (조인 속도용 색인) */
const _fkIdx={};
function fkIndex(tableKey,col){
  const k=tableKey+'|'+col;
  if(_fkIdx[k]) return _fkIdx[k];
  const m={};
  byKey[tableKey].rows.forEach(r=>{
    const v=String(r.cells[col]??'');
    if(v!=='')(m[v]=m[v]||[]).push(r);
  });
  return _fkIdx[k]=m;
}
const EXPLORE_CAP=20000;

function initExplore(base){
  const t=byKey[base];
  state.explore={base:base,tables:[],cols:t.columns.slice(0,4).map(c=>base+'|'+c),
    q:'',sort:null,page:1,missing:''};
}
function setExploreBase(k){initExplore(k);render();}
function exploreEffective(){
  /* 사용자가 고른 표 + 거기까지 가는 데 필요한 경유 표 */
  const ex=state.explore,{prev,dist}=joinPaths(ex.base);
  const need=new Set([ex.base]);
  ex.tables.forEach(tk=>{
    if(dist[tk]===undefined) return;
    let cur=tk;
    while(cur!==ex.base&&prev[cur]){need.add(cur);cur=prev[cur].from;}
  });
  return [...need].sort((a,b)=>(dist[a]||0)-(dist[b]||0));
}
function toggleExploreTable(k){
  const ex=state.explore;
  const i=ex.tables.indexOf(k);
  if(i>=0){
    ex.tables.splice(i,1);
    /* 더 이상 쓰이지 않는 표의 출력 컬럼은 같이 정리 */
    const keep=new Set(exploreEffective());
    ex.cols=ex.cols.filter(c=>keep.has(c.split('|')[0]));
    if(ex.missing==='missing:'+k) ex.missing='';
  }else{
    ex.tables.push(k);
    const t=byKey[k];
    /* 새로 붙인 표는 알아보기 쉬운 컬럼 2개를 기본으로 보여줍니다 */
    const pick=(NAME_COLS[k]||[t.columns[1]||t.pk]).slice(0,2);
    [t.pk].concat(pick).forEach(c=>{
      if(c&&!ex.cols.includes(k+'|'+c)) ex.cols.push(k+'|'+c);});
  }
  ex.page=1;render();
}
function toggleExploreCol(tk,col){
  const ex=state.explore,id=tk+'|'+col;
  const i=ex.cols.indexOf(id);
  if(i>=0) ex.cols.splice(i,1); else ex.cols.push(id);
  ex.page=1;render();
}
function exploreColsAll(tk,on){
  const ex=state.explore,t=byKey[tk];
  t.columns.forEach(c=>{
    const id=tk+'|'+c,i=ex.cols.indexOf(id);
    if(on&&i<0) ex.cols.push(id);
    if(!on&&i>=0) ex.cols.splice(i,1);});
  ex.page=1;render();
}
function setExploreQ(v){state.explore.q=v;state.explore.page=1;render();}
function setExploreMissing(v){state.explore.missing=v;state.explore.page=1;render();}
function setExplorePage(p){state.explore.page=p;render();}
function doExploreSort(id){
  const ex=state.explore;
  ex.sort=(ex.sort&&ex.sort.key===id)?{key:id,dir:-ex.sort.dir}:{key:id,dir:1};
  ex.page=1;render();
}
/* 기준 표의 각 행을 시작으로, 선택한 표들을 경로 순서대로 붙여나갑니다.
   상대가 없으면 null로 남겨 "누락"을 볼 수 있게 합니다(LEFT JOIN). */
function buildExplore(){
  const ex=state.explore;
  const eff=exploreEffective();
  const {prev}=joinPaths(ex.base);
  let out=byKey[ex.base].rows.map(r=>{const o={};o[ex.base]=r;return o;});
  let truncated=false;
  eff.filter(tk=>tk!==ex.base).forEach(tk=>{
    const step=prev[tk];if(!step) return;
    const parentT=byKey[step.from],childT=byKey[tk];
    const next=[];
    out.forEach(rec=>{
      const src=rec[step.from];
      let matches=[];
      if(src){
        if(step.dir==='child'){
          /* tk 가 step.from 을 참조 (1:N) */
          matches=(fkIndex(tk,step.col)[String(src.cells[parentT.pk]??'')]||[]);
        }else{
          /* step.from 이 tk 를 참조 (N:1) */
          const v=String(src.cells[step.col]??'');
          const m=v!==''&&rowIndex[tk]?rowIndex[tk][v]:null;
          matches=m?[m]:[];
        }
      }
      if(!matches.length){const o=Object.assign({},rec);o[tk]=null;next.push(o);}
      else matches.forEach(m=>{
        if(next.length>=EXPLORE_CAP){truncated=true;return;}
        const o=Object.assign({},rec);o[tk]=m;next.push(o);});
    });
    out=next;
  });
  /* 누락 필터 */
  const joined=eff.filter(tk=>tk!==ex.base);
  if(ex.missing==='any'&&joined.length) out=out.filter(rec=>joined.some(tk=>!rec[tk]));
  else if(ex.missing==='none'&&joined.length) out=out.filter(rec=>joined.every(tk=>!!rec[tk]));
  else if(ex.missing.indexOf('missing:')===0){
    const tk=ex.missing.slice(8);
    out=out.filter(rec=>!rec[tk]);
  }
  /* 표시 컬럼 — ex.cols 배열의 순서를 그대로 씁니다("3. 출력 컬럼"에서
     끌어서 바꾼 순서, 또는 새로 켠 컬럼이 뒤에 붙는 기본 순서). 더 이상
     "표 순서 → 원래 컬럼 순서"로 강제 정렬하지 않습니다 — 그렇게 하면
     드래그로 바꾼 순서가 매번 원래대로 되돌아가 버립니다. */
  const cols=ex.cols.filter(id=>eff.includes(id.split('|')[0]));
  /* 검색 — 표시 중인 컬럼들을 한 줄로 합친 텍스트에서, 검색어를 공백
     기준으로 쪼갠 낱말이 전부(AND) 어딘가에 있으면 매치. 특정 한 컬럼에
     검색어 전체가 고스란히 들어있어야 했던 예전 방식보다 여러 컬럼에
     걸친 값도 잘 찾고, 검색어에 공백이 있어도(예: "영광 풍력") 값에 공백이
     없어도(예: "영광풍력1호") 매치됩니다. */
  const q=(ex.q||'').trim();
  if(q) out=out.filter(rec=>matchesSearch(
    cols.map(id=>{const [tk,c]=id.split('|');return (rec[tk]&&rec[tk].cells[c])??'';}).join(' '),q));
  /* 정렬 */
  if(ex.sort){
    const [stk,sc]=ex.sort.key.split('|');
    out=[...out].sort((a,b)=>{
      const av=(a[stk]&&a[stk].cells[sc])??'',bv=(b[stk]&&b[stk].cells[sc])??'';
      if(av===''&&bv!=='') return 1;
      if(bv===''&&av!=='') return -1;
      const an=Number(av),bn=Number(bv);
      const cmp=(av!==''&&bv!==''&&!isNaN(an)&&!isNaN(bn))?an-bn:String(av).localeCompare(String(bv),'ko');
      return ex.sort.dir*cmp;});
  }
  return {rows:out,cols,eff,joined,truncated};
}
/* ── 출력 컬럼 드래그 앤 드롭 순서 변경 (순수 HTML5 Drag and Drop API) ────
   외부 라이브러리 없이 draggable 속성 + drag* 이벤트만으로 구현합니다.
   드롭하면 ex.cols 배열 안에서 항목을 옮기고 다시 그리는데, 그 사이에도
   네이티브 드래그 동작(dragend)은 브라우저가 알아서 정리하므로 별도 처리가
   필요 없습니다. */
let _exDragId=null;
function exColDragStart(e,id){
  _exDragId=id;
  e.dataTransfer.effectAllowed='move';
  e.dataTransfer.setData('text/plain',id); /* Firefox는 이게 없으면 드래그 자체가 시작 안 됨 */
  e.currentTarget.classList.add('dragging');
}
function exColDragOver(e){
  e.preventDefault(); /* dragover에서 preventDefault를 해야 그 자리를 드롭 대상으로 허용함 */
  e.dataTransfer.dropEffect='move';
  e.currentTarget.classList.add('dragover');
}
function exColDragLeave(e){e.currentTarget.classList.remove('dragover');}
function exColDrop(e,targetId){
  e.preventDefault();
  e.currentTarget.classList.remove('dragover');
  const ex=state.explore,fromId=_exDragId;
  _exDragId=null;
  if(!fromId||fromId===targetId||!ex) return;
  const cols=ex.cols;
  const from=cols.indexOf(fromId);
  if(from<0) return;
  cols.splice(from,1); /* 옮길 항목을 먼저 빼고 */
  const to=cols.indexOf(targetId); /* 뺀 뒤 기준으로 목표 위치를 다시 찾아 */
  cols.splice(to<0?from:to,0,fromId); /* 그 앞에 끼워 넣습니다(못 찾으면 원위치로 안전 복구) */
  render();
}
function exColDragEnd(e){
  e.currentTarget.classList.remove('dragging');
  _exDragId=null;
}
function exploreSelectedColsHtml(cols){
  if(!cols.length) return '';
  const chips=cols.map(id=>{
    const [tk,c]=id.split('|');
    return `<span class="excol-chip" draggable="true"
      ondragstart="exColDragStart(event,'${jsq(id)}')"
      ondragover="exColDragOver(event)"
      ondragleave="exColDragLeave(event)"
      ondrop="exColDrop(event,'${jsq(id)}')"
      ondragend="exColDragEnd(event)"
      title="끌어서 순서 변경">
      <span class="excol-handle">⠿⠿</span>${esc(byKey[tk].label)}.${esc(c)}
      <button class="excol-remove" onclick="event.stopPropagation();toggleExploreCol('${jsq(tk)}','${jsq(c)}')" title="이 컬럼 빼기">✕</button>
    </span>`;
  }).join('');
  return `<div class="excol-draglist">${chips}</div>
    <div class="excol-hint">칩을 마우스로 끌어서 표에 나오는 컬럼 순서를 바꿀 수 있습니다.</div>`;
}
function exploreExport(){
  const {rows,cols}=buildExplore();
  const head=cols.map(id=>{const [tk,c]=id.split('|');return byKey[tk].label+'.'+c;});
  const body=rows.map(rec=>cols.map(id=>{
    const [tk,c]=id.split('|');
    return String((rec[tk]&&rec[tk].cells[c])??'');}));
  return {head,body};
}
function exploreName(){
  const ex=state.explore;
  const miss=ex.missing.indexOf('missing:')===0?('_'+byKey[ex.missing.slice(8)].label+'없음')
    :(ex.missing==='any'?'_누락있음':(ex.missing==='none'?'_전부연결':''));
  return 'PPA_탐색_'+byKey[ex.base].label+miss;
}
function exploreDesc(){
  const ex=state.explore,{joined}=buildExplore();
  const parts=['기준 '+byKey[ex.base].label];
  if(joined.length) parts.push('연결 '+joined.map(k=>byKey[k].label).join('+'));
  if(ex.missing==='any') parts.push('누락 있는 행만');
  else if(ex.missing==='none') parts.push('전부 연결된 행만');
  else if(ex.missing.indexOf('missing:')===0) parts.push(byKey[ex.missing.slice(8)].label+' 없음');
  if((ex.q||'').trim()) parts.push('검색 "'+ex.q.trim()+'"');
  return parts.join(' · ');
}
function downloadExplore(fmt){
  const {head,body}=exploreExport();
  if(!head.length){toast('출력할 컬럼을 먼저 선택해주세요.');return;}
  const n=exploreName();
  if(fmt==='csv') downloadCsvRows(n,head,body);
  else if(fmt==='md') downloadMdRows(n,head,body,['- 조건: '+exploreDesc()]);
  else downloadXlsxRows(n,head,body);
}
/* 스키마에서 자동으로 뽑은 "빠진 것 찾기" 질문들 */
function missingPresets(){
  return edges.map(e=>({
    base:e.ref,child:e.table,
    label:`${byKey[e.ref].label} 중 ${byKey[e.table].label} 없음`}));
}
function applyMissingPreset(base,child){
  initExplore(base);
  const ex=state.explore;
  ex.tables=[child];
  const ct=byKey[child];
  ex.cols=ex.cols.concat([child+'|'+ct.pk]);
  ex.missing='missing:'+child;
  state.tab='탐색';render();
}

function tExplore(){
  if(!state.explore) initExplore(DATA.tables[0].key);
  const ex=state.explore;
  const {rows,cols,eff,joined,truncated}=buildExplore();
  const {dist}=joinPaths(ex.base);
  const total=rows.length;
  const pages=Math.max(1,Math.ceil(total/state.pageSize));
  const page=Math.min(Math.max(1,ex.page||1),pages);
  ex.page=page;
  const pageRows=rows.slice((page-1)*state.pageSize,page*state.pageSize);

  const baseTabs=DATA.tables.map(t=>
    `<button class="subtab ${t.key===ex.base?'on':''}" onclick="setExploreBase('${jsq(t.key)}')">${esc(t.label)}</button>`).join('');
  const joinChips=DATA.tables.filter(t=>t.key!==ex.base&&dist[t.key]!==undefined).map(t=>{
    const on=ex.tables.includes(t.key);
    const transit=!on&&eff.includes(t.key);
    return `<button class="colchip ${on?'on':(transit?'transit':'')}" onclick="toggleExploreTable('${jsq(t.key)}')"
      title="${transit?'다른 표를 연결하느라 자동으로 거쳐가는 표':(dist[t.key]+'단계 떨어져 있음')}">${esc(t.label)}<span class="chipdist">${dist[t.key]}</span></button>`;
  }).join('');
  const colGroups=eff.map(tk=>{
    const t=byKey[tk];
    const chips=t.columns.map(c=>
      `<button class="colchip ${ex.cols.includes(tk+'|'+c)?'on':''}" onclick="toggleExploreCol('${jsq(tk)}','${jsq(c)}')">${esc(c)}</button>`).join('');
    return `<div class="colgroup"><div class="colgrouphead">${esc(t.label)}${tk===ex.base?' <span class="badge mute">기준</span>':''}
      <button class="btn" style="padding:3px 8px;font-size:11.5px" onclick="exploreColsAll('${jsq(tk)}',true)">전체</button>
      <button class="btn" style="padding:3px 8px;font-size:11.5px" onclick="exploreColsAll('${jsq(tk)}',false)">해제</button></div>
      <div class="chiprow">${chips}</div></div>`;}).join('');
  const missOpts=[['','연결 여부 상관없이 전체'],['any','한 곳이라도 빠진 행만 (누락)'],['none','전부 연결된 행만']]
    .concat(joined.map(k=>['missing:'+k,byKey[k].label+'이(가) 없는 행만']));
  const missSel=`<select class="filtersel" onchange="setExploreMissing(this.value)">${
    missOpts.map(([v,l])=>`<option value="${esc(v)}" ${ex.missing===v?'selected':''}>${esc(l)}</option>`).join('')}</select>`;
  const presets=missingPresets().map(p=>
    `<button class="btn" onclick="applyMissingPreset('${jsq(p.base)}','${jsq(p.child)}')">${esc(p.label)}</button>`).join('');

  const head=cols.map(id=>{
    const [tk,c]=id.split('|');
    const ar=ex.sort&&ex.sort.key===id?`<span class="ar">${ex.sort.dir>0?'▲':'▼'}</span>`:'';
    return `<th class="${isNumCol(c)?'num':''}" onclick="doExploreSort('${jsq(id)}')">
      <span class="thtable">${esc(byKey[tk].label)}</span>${esc(c)}${ar}</th>`;}).join('');
  const body=pageRows.map(rec=>{
    const missAny=joined.some(tk=>!rec[tk]);
    const tds=cols.map(id=>{
      const [tk,c]=id.split('|');
      if(!rec[tk]) return '<td class="cellmiss">—</td>';
      return `<td class="${isNumCol(c)?'num':''}">${cellHtml(byKey[tk],c,rec[tk].cells[c])}</td>`;}).join('');
    const baseIdx=rowPos[ex.base]?rowPos[ex.base][String(rec[ex.base].cells[byKey[ex.base].pk])]:undefined;
    const click=baseIdx!==undefined?` onclick="openDetail('${jsq(ex.base)}',${baseIdx})"`:'';
    return `<tr class="${missAny?'rowmiss':''}${click?' clickrow':''}"${click}>${tds}</tr>`;}).join('')
    ||`<tr><td class="emptyrow" colspan="${Math.max(1,cols.length)}">조건에 맞는 데이터가 없습니다.</td></tr>`;

  const missCount=joined.length?rows.filter(rec=>joined.some(tk=>!rec[tk])).length:0;
  const pagerHtml=(()=>{
    if(pages<=1) return `<div class="pager"><span class="pginfo">${nf(total,0)}건 전체 표시</span>
      <select class="filtersel" onchange="setPageSize(this.value)">${pageSizeOpts()}</select></div>`;
    const btn=(p,l,d,on)=>`<button class="pgbtn${on?' on':''}" ${d?'disabled':''} onclick="setExplorePage(${p})">${l}</button>`;
    const nums=[];let last=0;
    for(let p=1;p<=pages;p++){
      if(p===1||p===pages||(p>=page-2&&p<=page+2)){
        if(last&&p-last>1) nums.push('<span class="pgdots">…</span>');
        nums.push(btn(p,String(p),false,p===page));last=p;}}
    const from=(page-1)*state.pageSize+1,to=Math.min(page*state.pageSize,total);
    return `<div class="pager">${btn(1,'«',page===1)}${btn(page-1,'‹',page===1)}${nums.join('')}${btn(page+1,'›',page===pages)}${btn(pages,'»',page===pages)}
      <span class="pginfo">${nf(from,0)}–${nf(to,0)} / ${nf(total,0)}건</span>
      <select class="filtersel" onchange="setPageSize(this.value)">${pageSizeOpts()}</select></div>`;})();

  return `<section>${printHead('관계형 데이터 탐색',exploreDesc())}
    <div class="kpis">
      ${kpi('결과 행 수',nf(total,0),exploreDesc(),'accent')}
      ${kpi('출력 컬럼',nf(cols.length,0),cols.length?'선택한 컬럼만 표시':'컬럼을 선택하세요')}
      ${kpi('연결된 표',nf(joined.length,0),joined.length?joined.map(k=>byKey[k].label).join(' · '):'기준 표만')}
      ${kpi('누락 있는 행',nf(missCount,0),joined.length?'연결 상대가 없는 행':'연결할 표를 골라주세요',missCount>0?'warn':'',
        missCount>0?"setExploreMissing('any')":'')}
    </div>
    ${panel('빠른 조회 — 자주 찾는 누락 데이터','누르면 아래 설정이 자동으로 맞춰집니다',`<div class="chiprow">${presets}</div>`)}
    ${panel('1. 기준 표','이 표의 각 행이 결과의 기준이 됩니다',`<div class="subtabbar">${baseTabs}</div>`)}
    ${panel('2. 연결할 표','숫자는 기준 표에서 몇 단계 떨어져 있는지 · 회색은 경로상 자동으로 거쳐가는 표',
      joinChips?`<div class="chiprow">${joinChips}</div>`:'<div class="nocand">연결할 수 있는 표가 없습니다.</div>')}
    ${panel('3. 출력 컬럼','보고 싶은 컬럼만 눌러서 켜고 끄세요',exploreSelectedColsHtml(cols)+colGroups)}
    ${panel('4. 결과',truncated?`행이 너무 많아 ${nf(EXPLORE_CAP,0)}건에서 끊었습니다 — 조건을 좁혀주세요`:'',
      `<div class="toolbar">
        <input id="exploreSearch" class="search" placeholder="결과에서 검색… (여러 낱말 가능)" value="${esc(ex.q||'')}" oninput="onSearchType(event,setExploreQ)">
        ${missSel}
        <details class="drop"><summary>내려받기</summary><div class="dropbody right">
          <div class="nocand" style="padding:2px 2px 8px">현재 결과 <b>${nf(total,0)}건</b> · 선택 컬럼 ${cols.length}개</div>
          <button class="dlopt" onclick="downloadExplore('csv')">CSV (.csv)<small>엑셀에서 바로 열림 (UTF-8 BOM)</small></button>
          <button class="dlopt" onclick="downloadExplore('xlsx')">Excel (.xlsx)<small>서식 있는 실제 엑셀 파일</small></button>
          <button class="dlopt" onclick="downloadExplore('md')">Markdown (.md)<small>메일·보고서 붙여넣기용</small></button>
          <button class="dlopt" onclick="window.print()">PDF / 인쇄<small>브라우저 인쇄창에서 "PDF로 저장"</small></button>
        </div></details>
        <button class="btn" onclick="window.print()">인쇄 / PDF</button>
        <span class="count">${nf(total,0)}건</span>
      </div>
      ${cols.length?`<div class="tbl-wrap stickyfirst"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>${pagerHtml}`
        :'<div class="nocand">출력할 컬럼을 하나 이상 선택해주세요.</div>'}`)}
    </section>`;
}

/* ── 홈 (보고용 요약) ───────────────────────────────────────────────────── */
/* 발전원처럼 순수 명목형(nominal) 카테고리를 색으로 구분할 때 쓰는
   4색 계열 — dataviz 팔레트 검증(밝기대역/채도/CVD분리/대비)을 통과한
   고정 순서. 5개 넘게 나오면 그냥 돌려씀(발전원 종류가 실제로 4개를
   넘는 경우는 드묾) */
const CAT_COLORS=['var(--s1)','var(--s2)','var(--s3)','var(--s4)'];
function sumCol(tk,col){const t=byKey[tk];if(!t) return 0;
  return t.rows.reduce((s,r)=>{const n=Number(r.cells[col]);return s+(isNaN(n)?0:n);},0);}
function countWhere(tk,col,val){const t=byKey[tk];if(!t) return 0;
  return t.rows.filter(r=>String(r.cells[col]??'')===val).length;}
function groupSum(tk,gcol,scol){const t=byKey[tk];if(!t) return [];
  const m={};
  t.rows.forEach(r=>{const g=r.cells[gcol]||'(미지정)';const n=Number(r.cells[scol]);m[g]=(m[g]||0)+(isNaN(n)?0:n);});
  return Object.entries(m).sort((a,b)=>b[1]-a[1]);}
function countExpiring(tk,col,days){
  const t=byKey[tk];if(!t||!TODAY) return 0;
  return t.rows.filter(r=>{
    const v=String(r.cells[col]||'');
    if(!/^\d{4}-\d{2}-\d{2}$/.test(v)) return false;
    const d=daysBetween(TODAY,v);
    return d>=0&&d<=days;}).length;}
function jumpToFilter(tk,col,val){
  state.tab=tk;state.filters[tk]=state.filters[tk]||{};state.filters[tk][col]=val;state.page[tk]=1;render();}
function jumpToDateWindow(tk,col,days){
  state.tab=tk;state.dateFilters[tk]=state.dateFilters[tk]||{};
  const end=new Date(TODAY);end.setDate(end.getDate()+days);
  state.dateFilters[tk][col]={from:TODAY,to:end.toISOString().slice(0,10)};
  state.page[tk]=1;render();}
function jumpToDateWindowRange(tk,col,minDays,maxDays){
  state.tab=tk;state.dateFilters[tk]=state.dateFilters[tk]||{};
  const from=new Date(TODAY),to=new Date(TODAY);
  from.setDate(from.getDate()+minDays);to.setDate(to.getDate()+maxDays);
  state.dateFilters[tk][col]={from:from.toISOString().slice(0,10),to:to.toISOString().slice(0,10)};
  state.page[tk]=1;render();}
/* [minDays,maxDays] 구간(오늘 기준 상대일)에 걸리는 건수·용량 — D-30/D-60
   같은 "임박" 구간 집계에 재사용. */
function dateBucket(tk,col,capCol,minDays,maxDays){
  const t=byKey[tk];if(!t||!TODAY) return {n:0,cap:0};
  let n=0,cap=0;
  t.rows.forEach(r=>{
    const v=String(r.cells[col]||'');if(!/^\d{4}-\d{2}-\d{2}$/.test(v)) return;
    const d=daysBetween(TODAY,v);
    if(d>=minDays&&d<=maxDays){n++;const c=Number(r.cells[capCol]);cap+=isNaN(c)?0:c;}
  });
  return {n,cap};
}
/* 시작일이 종료일보다 뒤인("날짜 역전") 행의 PK 목록 — 개수만 필요하면
   .length를 쓰고, 드릴다운(jumpToPkSet)에는 배열 자체를 그대로 씁니다. */
function dateInversionPks(tk,startCol,endCol){
  const t=byKey[tk];if(!t) return [];
  return t.rows.filter(r=>{
    const a=String(r.cells[startCol]||''),b=String(r.cells[endCol]||'');
    if(!/^\d{4}-\d{2}-\d{2}$/.test(a)||!/^\d{4}-\d{2}-\d{2}$/.test(b)) return false;
    return a>b;
  }).map(r=>String(r.cells[t.pk]??''));
}

/* 현황별로 값을 다른 표(FK)에서 끌어와 합산 — 예: 수급매칭 현황별 구매계약 용량 */
function groupSumJoinFK(tk,gcol,fkCol,refTk,refCapCol){
  const t=byKey[tk],ref=byKey[refTk];if(!t||!ref) return [];
  const refMap={};ref.rows.forEach(r=>{refMap[r.cells[ref.pk]]=Number(r.cells[refCapCol])||0;});
  const m={};
  t.rows.forEach(r=>{
    const g=String(r.cells[gcol]||'').trim()||'(미지정)';
    if(!m[g]) m[g]={cnt:0,cap:0};
    m[g].cnt++;m[g].cap+=refMap[r.cells[fkCol]]||0;
  });
  return Object.entries(m).map(([g,v])=>({g,cnt:v.cnt,cap:v.cap})).sort((a,b)=>b.cnt-a.cnt);
}
/* "N. 라벨" 형태의 현황 값에서 정렬용 코드와 사람이 읽는 라벨을 분리 —
   화면에는 라벨만 보여주고, 정렬은 코드 기준 오름차순으로 합니다. 코드가
   없는 값(자유 텍스트로 입력된 경우)은 맨 뒤로, 라벨 가나다순으로 정렬. */
function parseStatus(raw){
  const s=String(raw||'').trim();
  const m=s.match(/^(\d+)\s*\.?\s*(.*)$/);
  if(m&&m[2]) return {code:Number(m[1]),label:m[2].trim()};
  return {code:null,label:s};
}
/* 구매계약/판매계약 표 자체엔 "종료" 상태 컬럼이 없어서, 그 계약에 걸린
   수급매칭 현황을 근거로 판단합니다: 매칭이 하나 이상 있고 그 전부가
   "8. 공급종료"(isTerminalStatus)일 때만 "완전히 종료"로 봅니다. 매칭이
   하나도 없으면 판단할 근거가 없으니 활성으로 취급합니다(계약이 아직
   매칭 전 단계일 수도 있으므로). 색상 분류(statusClass)와는 별개로
   terminal 여부만 따로 봅니다 - "착공 전" 등 다른 회색(mute) 상태를
   종료로 착각하지 않도록. */
function purchaseTerminatedIds(){
  const M=byKey['T_수급매칭'];const out=new Set();
  if(!M) return out;
  const agg={};
  M.rows.forEach(r=>{
    const pk=String(r.cells['구매계약ID']||'');if(!pk) return;
    if(!agg[pk]) agg[pk]={total:0,term:0};
    agg[pk].total++;
    if(isTerminalStatus(r.cells['현황'])) agg[pk].term++;
  });
  Object.entries(agg).forEach(([pk,v])=>{if(v.total>0&&v.term===v.total) out.add(pk);});
  return out;
}
function saleTerminatedIds(){
  const E=byKey['T_전기사용지'],M=byKey['T_수급매칭'];const out=new Set();
  if(!E||!M) return out;
  const elecToSale={};
  E.rows.forEach(r=>{elecToSale[String(r.cells['전기사용지ID']||'')]=String(r.cells['판매계약ID']||'');});
  const agg={};
  M.rows.forEach(r=>{
    const saleId=elecToSale[String(r.cells['전기사용지ID']||'')];if(!saleId) return;
    if(!agg[saleId]) agg[saleId]={total:0,term:0};
    agg[saleId].total++;
    if(isTerminalStatus(r.cells['현황'])) agg[saleId].term++;
  });
  Object.entries(agg).forEach(([id,v])=>{if(v.total>0&&v.term===v.total) out.add(id);});
  return out;
}
/* 종료 건을 뺀 유효 용량과, 뺀 종료 건 자체의 건수·용량을 한 번에 계산 —
   메인 KPI(유효 집계)와 보조 "종료/만료 용량" 카드가 이 하나의 함수를
   같은 기준으로 나눠 쓰게 해서 두 숫자가 항상 서로 앞뒤가 맞습니다. */
/* 지난 기준점(스냅샷) 대비 용량 증감(MW) — added/changed는 현재 표의
   change 마크(prev)로, removed는 CHANGES.removed_rows(삭제 당시 값 그대로
   보존됨)로 계산해 셋을 합칩니다. Executive 요약에서 "늘었는지 줄었는지"를
   3초 안에 보여주기 위한 델타 표시용. */
function capacityDelta(tk,capCol){
  const t=byKey[tk];if(!t||!CHANGES.has_prev) return 0;
  let d=0;
  t.rows.forEach(r=>{
    const n=Number(r.cells[capCol]);if(isNaN(n)) return;
    if(r.change==='added') d+=n;
    else if(r.change==='changed'&&r.prev&&Object.prototype.hasOwnProperty.call(r.prev,capCol)){
      const pn=Number(r.prev[capCol]);d+=n-(isNaN(pn)?0:pn);
    }
  });
  const removed=(CHANGES.removed_rows&&CHANGES.removed_rows[tk])||[];
  removed.forEach(cells=>{const n=Number(cells[capCol]);if(!isNaN(n)) d-=n;});
  return d;
}
function deltaBadge(delta){
  if(!CHANGES.has_prev||Math.abs(delta)<0.005) return '';
  const up=delta>0;
  return ` <span style="color:${up?'var(--pass)':'var(--fail)'};font-weight:700" title="지난 기준점 대비 변화">${up?'▲':'▼'} ${up?'+':''}${nf(delta)}MW</span>`;
}
function sumCapSplit(tk,capCol,excludeIds){
  const t=byKey[tk];if(!t) return {activeN:0,activeMW:0,termN:0,termMW:0};
  let activeN=0,activeMW=0,termN=0,termMW=0;
  t.rows.forEach(r=>{
    const pk=String(r.cells[t.pk]||'');
    const n=Number(r.cells[capCol]);const v=isNaN(n)?0:n;
    if(excludeIds.has(pk)){termN++;termMW+=v;} else {activeN++;activeMW+=v;}
  });
  return {activeN,activeMW,termN,termMW};
}

/* ── 용량 정합성 분석 ────────────────────────────────────────────────────
   "구매계약 총 용량이 왜 발전소 총 설비용량과 다른가" 같은 질문에 숫자
   차이만 보여주는 게 아니라, 그 차이를 발전소 단위로 쪼개 "계약이 아예
   없는 발전소가 몇 개(몇 MW)", "일부만 계약된 발전소의 미계약분이 몇 MW",
   "설비용량을 초과해 계약된(데이터 이상 소지) 발전소가 몇 개"로 원인을
   분해해서 보여줍니다. 판매계약↔전기사용지 관계에도 같은 로직을 재사용. */
function sumByFK(childTk,fkCol,capCol){
  const t=byKey[childTk];const m={};
  if(!t) return m;
  t.rows.forEach(r=>{
    const fk=String(r.cells[fkCol]||'').trim();if(!fk) return;
    const n=Number(r.cells[capCol]);
    m[fk]=(m[fk]||0)+(isNaN(n)?0:n);
  });
  return m;
}
function capacityGap(parentTk,parentCapCol,parentNameCol,childTk,fkCol,childCapCol){
  const p=byKey[parentTk];if(!p) return null;
  const sums=sumByFK(childTk,fkCol,childCapCol);
  const items=p.rows.map((r,i)=>{
    const pk=String(r.cells[p.pk]||'');
    const installed=Number(r.cells[parentCapCol])||0;
    const contracted=sums[pk]||0;
    return {i,pk:r.cells[p.pk],name:parentNameCol?(r.cells[parentNameCol]||''):'',installed,contracted,gap:installed-contracted};
  });
  const totalInstalled=items.reduce((s,x)=>s+x.installed,0);
  const totalContracted=items.reduce((s,x)=>s+x.contracted,0);
  return {
    items,totalInstalled,totalContracted,totalGap:totalInstalled-totalContracted,
    zero:items.filter(x=>x.installed>0&&x.contracted<=0.0001),
    under:items.filter(x=>x.gap>0.0001&&x.contracted>0.0001),
    over:items.filter(x=>x.gap<-0.0001),
  };
}
function capacityGapBlock(title,parentTk,parentCapCol,parentNameCol,childTk,fkCol,childCapCol,childLabel){
  const P=byKey[parentTk],C=byKey[childTk];
  if(!P||!C||!P.rows.length) return '';
  const g=capacityGap(parentTk,parentCapCol,parentNameCol,childTk,fkCol,childCapCol);
  if(!g) return '';
  const gapAbs=Math.abs(g.totalGap);
  const headLine=gapAbs>0.05
    ?`총 ${nf(g.totalInstalled)}MW 중 ${esc(childLabel)} ${nf(g.totalContracted)}MW — 차이 ${nf(gapAbs)}MW ${g.totalGap>0?'미계약':'초과'}`
    :`총 ${nf(g.totalInstalled)}MW — ${esc(childLabel)} 용량과 정확히 일치`;
  const pct=g.totalInstalled>0?Math.min(100,g.totalContracted/g.totalInstalled*100):0;
  const overFlag=g.totalGap<-0.05;
  const barHtml=`<div class="mixbar" style="margin:6px 0 8px" title="${esc(childLabel)} ${nf(g.totalContracted)}MW / 설비 ${nf(g.totalInstalled)}MW">
    <div class="mixfill" style="width:${pct.toFixed(1)}%${overFlag?';background:var(--fail)':''}"></div></div>`;
  const parts=[];
  if(g.zero.length) parts.push(`${esc(childLabel)}이 전혀 없는 항목 <b>${nf(g.zero.length,0)}개</b>(<span class="dim">${nf(g.zero.reduce((s,x)=>s+x.installed,0))}MW</span>)`);
  if(g.under.length) parts.push(`부분 계약분 미달 <b>${nf(g.under.length,0)}개</b>(<span class="dim">${nf(g.under.reduce((s,x)=>s+x.gap,0))}MW</span>)`);
  if(g.over.length) parts.push(`설비용량 초과 계약 <b>${nf(g.over.length,0)}개</b>(<span class="dim">${nf(g.over.reduce((s,x)=>s-x.gap,0))}MW 초과</span>) — 데이터 확인 필요`);
  const reasonText=parts.length?`<p style="margin:2px 0 10px;color:var(--sub);font-size:12.5px">${parts.join(' + ')}로 설명됩니다.</p>`:'';
  const list=g.items.filter(x=>Math.abs(x.gap)>0.0001).sort((a,b)=>Math.abs(b.gap)-Math.abs(a.gap)).slice(0,8)
    .map(x=>{
      const kind=x.gap>0?(x.contracted<=0?'미계약':'부분계약'):'초과계약';
      const badge=x.gap>0?(x.contracted<=0?'no':'warn'):'info';
      return `<div class="unsecrow" onclick="openDetail('${jsq(parentTk)}',${x.i})">
        <span>${esc(String(x.pk))}${x.name?' · '+esc(x.name):''}</span>
        <span class="chgval mono">${nf(x.contracted)} / ${nf(x.installed)} MW</span>
        <span class="badge ${badge}">${kind}</span></div>`;
    }).join('');
  return `<div style="margin-bottom:14px"><div style="font-weight:700;font-size:12.5px;margin-bottom:4px;color:var(--sub)">${esc(title)}</div>
    <div style="font-size:13.5px;font-weight:600">${headLine}</div>
    ${barHtml}${reasonText}${list||'<div class="nocand">모든 항목이 정확히 일치합니다.</div>'}</div>`;
}
function capacityGapPanel(){
  const b1=capacityGapBlock('발전소 → 구매계약 (공급 측)','T_발전소','설비용량(MW)','발전소명','T_구매계약','발전소ID','구매계약용량(MW)','구매계약');
  const b2=capacityGapBlock('판매계약 → 전기사용지 (수요 측)','T_판매계약','판매계약용량(MW)','판매 담당자','T_전기사용지','판매계약ID','전기사용지계약용량(MW)','전기사용지');
  if(!b1&&!b2) return '';
  return panel('용량 정합성 분석','구매계약/판매계약이 상위 표의 용량을 얼마나 반영하는지, 차이가 왜 나는지 발전소·판매계약 단위로 분해해 보여줍니다',b1+b2);
}

/* ── 우선순위 조치 필요 항목 ─────────────────────────────────────────────
   PPA 운영 실무 관점 3개 카테고리로 나눠 보여줍니다:
     A. 계약/공급 임박 알림 — 공급기한 D-30/D-60, 이미 만료
     B. 수급 불균형·미확보 모니터링 — 잔여(미매칭) 용량, 초과계약, 미확보
     C. 데이터 정합성 이상치 — 검증 오류, 날짜 역전
   각 행은 건수와 함께(가능하면) 용량(MW)을 보여주고, 심각도 뱃지
   (위험/주의/정상)와 클릭 시 해당 조건으로 이동하는 바로가기를 갖습니다. */
function actionRow(label,n,capMW,badgeCls,badgeText,action){
  const has=n>0;
  const finalCls=has?badgeCls:'ok';
  const finalText=has?badgeText:'정상';
  const capStr=(has&&capMW!=null)?` · ${nf(capMW)} MW`:'';
  return `<div class="unsecrow${has?'':' isok'}"${action?` onclick="${action}"`:''}>
    <span>${esc(label)}</span>
    <span class="chgval mono">${nf(n,0)}건${capStr}</span>
    <span class="badge ${finalCls}">${finalText}</span></div>`;
}
function actionCategory(title,rows){
  const activeN=rows.filter(r=>r.n>0).length;
  const body=rows.map(r=>actionRow(r.label,r.n,r.cap,r.badgeCls,r.badgeText,r.action)).join('');
  return `<div class="actioncat">
      <div class="actioncathead">${esc(title)}${activeN?`<span class="actioncatcount">${activeN}</span>`:''}</div>
      ${body}
    </div>`;
}
function actionItemsPanel(){
  const P=byKey['T_발전소'],B=byKey['T_구매계약'],S=byKey['T_판매계약'];

  /* A. 계약/공급 임박 알림 — 이미 만료된 건은 지금 손을 쓸 수 없는(과거가
     된) 사안이라 "조치 필요" 목록에서 완전히 제외합니다. 아직 조치할 시간이
     남아 있는(공급기한이 다가오는) 항목만 보여줍니다. */
  const b30=B?dateBucket('T_구매계약','공급기한_구매','구매계약용량(MW)',0,30):{n:0,cap:0};
  const s30=S?dateBucket('T_판매계약','공급기한_판매','판매계약용량(MW)',0,30):{n:0,cap:0};
  const b60=B?dateBucket('T_구매계약','공급기한_구매','구매계약용량(MW)',31,60):{n:0,cap:0};
  const s60=S?dateBucket('T_판매계약','공급기한_판매','판매계약용량(MW)',31,60):{n:0,cap:0};
  const catA=[
    {label:'구매계약 — 공급기한 D-30 이내',n:b30.n,cap:b30.cap,badgeCls:'no',badgeText:'위험',action:"jumpToDateWindow('T_구매계약','공급기한_구매',30)"},
    {label:'판매계약 — 공급기한 D-30 이내',n:s30.n,cap:s30.cap,badgeCls:'no',badgeText:'위험',action:"jumpToDateWindow('T_판매계약','공급기한_판매',30)"},
    {label:'구매계약 — 공급기한 D-31~60',n:b60.n,cap:b60.cap,badgeCls:'warn',badgeText:'주의',action:"jumpToDateWindowRange('T_구매계약','공급기한_구매',31,60)"},
    {label:'판매계약 — 공급기한 D-31~60',n:s60.n,cap:s60.cap,badgeCls:'warn',badgeText:'주의',action:"jumpToDateWindowRange('T_판매계약','공급기한_판매',31,60)"},
  ];

  /* B. 수급 불균형·미확보 모니터링 — "발전소 탭으로 이동" 정도가 아니라
     실제 문제가 있는 발전소 PK만 골라 그 표를 필터링해서 보여줍니다
     (jumpToPkSet). g.zero/g.under/g.over는 전부 발전소 행 기준입니다 —
     "초과 계약"도 어떤 구매계약이 초과인지가 아니라 어떤 발전소가 설비
     용량보다 많이 계약됐는지를 가리키므로 발전소 표로 드릴다운합니다. */
  const g=(P&&B&&P.rows.length)?capacityGap('T_발전소','설비용량(MW)','발전소명','T_구매계약','발전소ID','구매계약용량(MW)'):null;
  const residualPks=g?[...g.zero,...g.under].map(x=>String(x.pk)):[];
  const residualMW=g?(g.zero.reduce((s,x)=>s+x.installed,0)+g.under.reduce((s,x)=>s+x.gap,0)):0;
  const overPks=g?g.over.map(x=>String(x.pk)):[];
  const overMW=g?g.over.reduce((s,x)=>s-x.gap,0):0;
  const bUn=B?countWhere('T_구매계약','수요기업 미확보','TRUE'):0;
  const bUnMW=B?B.rows.filter(r=>r.cells['수요기업 미확보']==='TRUE').reduce((s,r)=>{const n=Number(r.cells['구매계약용량(MW)']);return s+(isNaN(n)?0:n);},0):0;
  const sUn=S?countWhere('T_판매계약','공급자원 미확보','TRUE'):0;
  const sUnMW=S?S.rows.filter(r=>r.cells['공급자원 미확보']==='TRUE').reduce((s,r)=>{const n=Number(r.cells['판매계약용량(MW)']);return s+(isNaN(n)?0:n);},0):0;
  const catB=[
    {label:'발전 설비용량 대비 수요계약 미매칭 잔여용량',n:residualPks.length,cap:residualMW,badgeCls:'warn',badgeText:'주의',action:`jumpToPkSet('T_발전소',${pkArrLiteral(residualPks)})`},
    {label:'발전소 — 설비용량 초과 계약(구매계약 합계 > 설비용량)',n:overPks.length,cap:overMW,badgeCls:'no',badgeText:'위험',action:`jumpToPkSet('T_발전소',${pkArrLiteral(overPks)})`},
    {label:'구매계약 — 수요기업 미확보',n:bUn,cap:bUnMW,badgeCls:'warn',badgeText:'주의',action:"jumpToFilter('T_구매계약','수요기업 미확보','TRUE')"},
    {label:'판매계약 — 공급자원 미확보',n:sUn,cap:sUnMW,badgeCls:'warn',badgeText:'주의',action:"jumpToFilter('T_판매계약','공급자원 미확보','TRUE')"},
  ];

  /* C. 데이터 정합성 이상치 */
  const errN=DATA.validation.total_errors;
  const dateInvPks=S?dateInversionPks('T_판매계약','계약일','공급기한_판매'):[];
  const catC=[
    {label:'검증 오류 (PK 누락·중복, FK 누락·참조, 조합중복)',n:errN,cap:null,badgeCls:'no',badgeText:'위험',action:"state.tab='검증';render()"},
    {label:'날짜 역전 오류 (판매계약 계약일 > 공급기한)',n:dateInvPks.length,cap:null,badgeCls:'no',badgeText:'위험',action:dateInvPks.length?`jumpToPkSet('T_판매계약',${pkArrLiteral(dateInvPks)})`:''},
  ];

  const allRows=[...catA,...catB,...catC];
  const dangerN=allRows.filter(r=>r.n>0&&r.badgeCls==='no').length;
  const warnN=allRows.filter(r=>r.n>0&&r.badgeCls==='warn').length;
  const totalIssues=dangerN+warnN;
  const summaryHtml=`<div class="actionsummary">
      ${dangerN?`<span class="badge no">위험 ${nf(dangerN,0)}</span>`:''}
      ${warnN?`<span class="badge warn">주의 ${nf(warnN,0)}</span>`:''}
      ${!totalIssues?`<span class="badge ok">위험·주의 항목 없음</span>`:''}
    </div>`;
  return panel('우선순위 조치 필요 항목',
    totalIssues?`${totalIssues}개 항목에서 확인이 필요합니다 · 클릭하면 해당 조건으로 이동`:'지금 확인이 필요한 항목이 없습니다',
    summaryHtml
    +actionCategory('계약/공급 임박 알림',catA)
    +actionCategory('수급 불균형 · 미확보 모니터링',catB)
    +actionCategory('데이터 정합성 이상치',catC));
}

const STATUS_COLOR={ok:'var(--pass)',warn:'var(--amber)',mute:'var(--sub)',no:'var(--fail)',info:'var(--info)'};
function statusColor(v){return STATUS_COLOR[statusClass(v)]||'var(--info)';}

/* 년월(YYYY-MM) 단위 집계 — 날짜 컬럼 하나를 기준으로 건수·용량 */
function monthKey(v){const s=String(v||'');return /^\d{4}-\d{2}/.test(s)?s.slice(0,7):null;}
function monthLabel(ym){const p=ym.split('-');return "'"+p[0].slice(2)+'.'+p[1];}
function groupByMonth(tk,dateCol,capCol){
  const t=byKey[tk];if(!t) return {};
  const m={};
  t.rows.forEach(r=>{
    const ym=monthKey(r.cells[dateCol]);if(!ym) return;
    if(!m[ym]) m[ym]={cnt:0,cap:0};
    m[ym].cnt++;
    const n=Number(r.cells[capCol]);m[ym].cap+=isNaN(n)?0:n;
  });
  return m;
}
function dayKey(v){const s=String(v||'');return /^\d{4}-\d{2}-\d{2}$/.test(s)?s:null;}
function dayLabel(d){return String(Number(d.slice(8,10)))+'일';}
function groupByDay(tk,dateCol,capCol){
  const t=byKey[tk];if(!t) return {};
  const m={};
  t.rows.forEach(r=>{
    const d=dayKey(r.cells[dateCol]);if(!d) return;
    if(!m[d]) m[d]={cnt:0,cap:0};
    m[d].cnt++;
    const n=Number(r.cells[capCol]);m[d].cap+=isNaN(n)?0:n;
  });
  return m;
}
function daysOfMonth(ym){
  const p=ym.split('-').map(Number),last=new Date(p[0],p[1],0).getDate(),out=[];
  for(let d=1;d<=last;d++) out.push(ym+'-'+String(d).padStart(2,'0'));
  return out;
}
function monthRange(n){ // 이번 달을 포함해 최근 n개월의 YYYY-MM 목록 (과거→현재)
  const out=[];const base=TODAY?new Date(TODAY):new Date();
  for(let i=n-1;i>=0;i--){
    const d=new Date(base.getFullYear(),base.getMonth()-i,1);
    out.push(d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0'));
  }
  return out;
}
function jumpToMonth(tk,col,ym){
  state.tab=tk;state.dateFilters[tk]=state.dateFilters[tk]||{};
  const p=ym.split('-').map(Number);
  const lastDay=new Date(p[0],p[1],0).getDate();
  state.dateFilters[tk][col]={from:ym+'-01',to:ym+'-'+String(lastDay).padStart(2,'0')};
  state.page[tk]=1;render();
}
function jumpToDay(tk,col,d){
  state.tab=tk;state.dateFilters[tk]=state.dateFilters[tk]||{};
  state.dateFilters[tk][col]={from:d,to:d};
  state.page[tk]=1;render();
}
function drillMonth(ym){state.homeTrend.drillYm=ym;render();}
function undrillMonth(){state.homeTrend.drillYm=null;render();}
function schemaBox(tk){const t=byKey[tk];
  return `<button class="schemabox" onclick="state.tab='${jsq(tk)}';render()">${esc(t.label)}<span class="schemacount mono">${nf(t.rows.length,0)}</span></button>`;}
function schemaDiagram(){
  const need=['T_발전소','T_구매계약','T_수급매칭','T_수요기업','T_판매계약','T_전기사용지'];
  if(!need.every(k=>byKey[k])) return '';
  return `<div class="schemarow"><span class="schematag">공급측</span>${schemaBox('T_발전소')}<span class="schemaarrow">→</span>${schemaBox('T_구매계약')}<span class="schemaarrow">→</span>${schemaBox('T_수급매칭')}</div>
    <div class="schemarow"><span class="schematag">수요측</span>${schemaBox('T_수요기업')}<span class="schemaarrow">→</span>${schemaBox('T_판매계약')}<span class="schemaarrow">→</span>${schemaBox('T_전기사용지')}<span class="schemaarrow">→</span>${schemaBox('T_수급매칭')}</div>`;}

/* ── 홈: 현황별 비율·용량 ───────────────────────────────────────────────── */
function statusPanel(){
  const M=byKey['T_수급매칭'];if(!M||!M.rows.length) return '';
  const groups=groupSumJoinFK('T_수급매칭','현황','구매계약ID','T_구매계약','구매계약용량(MW)');
  if(!groups.length) return '';
  /* "N. 라벨" 값에서 코드를 뽑아 코드 오름차순(1→99)으로 정렬 — 화면엔
     라벨만 보여주고 원래 값(g.g)은 필터링용으로만 씁니다. */
  const withMeta=groups.map(g=>({...g,parsed:parseStatus(g.g)}));
  withMeta.sort((a,b)=>{
    if(a.parsed.code!=null&&b.parsed.code!=null) return a.parsed.code-b.parsed.code;
    if(a.parsed.code!=null) return -1;
    if(b.parsed.code!=null) return 1;
    return a.parsed.label.localeCompare(b.parsed.label,'ko');
  });
  const total=withMeta.reduce((s,g)=>s+g.cnt,0);
  const totalCap=withMeta.reduce((s,g)=>s+g.cap,0);
  const segs=withMeta.map(g=>{
    const pct=total?g.cnt/total*100:0;
    return `<div class="statusseg" style="width:${pct}%;background:${statusColor(g.g)}" title="${esc(g.parsed.label)} · ${nf(g.cnt,0)}건 (${nf(pct,0)}%)"></div>`;
  }).join('');
  const legend=withMeta.map(g=>{
    const pct=total?g.cnt/total*100:0;
    return `<div class="statusrow" onclick="jumpToFilter('T_수급매칭','현황','${jsq(g.g)}')">
      <span class="statusdot" style="background:${statusColor(g.g)}"></span>
      <span class="statuslabel">${esc(g.parsed.label)}</span>
      <span class="statuspct mono">${nf(pct,0)}%</span>
      <span class="statuscnt mono">${nf(g.cnt,0)}건</span>
      <span class="statuscap mono">${nf(g.cap)} MW</span></div>`;
  }).join('');
  return panel('수급매칭 — 현황별 비율 · 용량','칸/행을 누르면 그 현황만 보기로 이동합니다 (용량은 연결된 구매계약 기준)',
    `<div class="statusbar">${segs}</div>
     <div class="statuslegendhead"><span></span><span>현황</span><span>비중</span><span>건수</span><span>용량</span></div>
     ${legend}
     <div class="statustotal">전체 ${nf(total,0)}건 · 용량 합계 ${nf(totalCap)} MW</div>`);
}

/* ── 홈: 년월별 추이 ────────────────────────────────────────────────────── */
function ymIdx(ym){const p=ym.split('-').map(Number);return p[0]*12+(p[1]-1);}
function idxYm(i){const y=Math.floor(i/12),mo=i%12;return y+'-'+String(mo+1).padStart(2,'0');}
/* 두 YYYY-MM 사이(포함) 월 목록 — 순서가 뒤바뀌어 입력돼도 알아서
   맞바꾸고, 화면이 너무 넓어지지 않도록 한 번에 최대 36개월까지만. */
function monthRangeBetween(fromYm,toYm){
  let a=ymIdx(fromYm),b=ymIdx(toYm);
  if(a>b){const t=a;a=b;b=t;}
  if(b-a>35) b=a+35;
  const out=[];for(let i=a;i<=b;i++) out.push(idxYm(i));
  return out;
}
/* 조회 기간 필터가 설정돼 있으면 그 구간을, 아니면 기존 기본값(최근
   12개월)을 돌려줍니다. */
function trendRangeKeys(){
  const {rangeFrom,rangeTo}=state.homeTrend;
  if(rangeFrom&&rangeTo) return monthRangeBetween(rangeFrom,rangeTo);
  return monthRange(12);
}
function resetHomeTrendRange(){state.homeTrend.rangeFrom=null;state.homeTrend.rangeTo=null;render();}
const TREND_METRICS={
  new:{label:'신규 판매계약',tk:'T_판매계약',col:'계약일',cap:'판매계약용량(MW)'},
  buyexp:{label:'구매계약 공급기한',tk:'T_구매계약',col:'공급기한_구매',cap:'구매계약용량(MW)'},
  saleexp:{label:'판매계약 공급기한',tk:'T_판매계약',col:'공급기한_판매',cap:'판매계약용량(MW)'},
};
function setHomeTrend(k,v){state.homeTrend[k]=v;render();}
function fmtTick(v,unit){return unit==='cap'?nf(v):nf(v,0);}
/* "보기 좋은" 눈금 간격 — 4등분 근처에서 1/2/5×10ⁿ 중 하나를 고릅니다 */
function niceStep(range,targetTicks){
  const raw=range/Math.max(1,targetTicks);
  if(raw<=0) return 1;
  const mag=Math.pow(10,Math.floor(Math.log10(raw)));
  const norm=raw/mag;
  let step;
  if(norm<1.5) step=1;else if(norm<3) step=2;else if(norm<7) step=5;else step=10;
  return step*mag;
}
function niceTicks(max,integer){
  if(max<=0) return [0,1];
  let step=niceStep(max,4);
  if(integer) step=Math.max(1,Math.round(step)); /* 건수처럼 정수만 뜻이 있는 값은 0.2 같은 눈금이 안 나오게 */
  const top=Math.ceil(max/step)*step,out=[];
  for(let v=0;v<=top+1e-9;v+=step) out.push(Math.round(v*1000)/1000);
  return out;
}
/* 월별/일별 추이 공용 렌더러 — keys 순서대로 막대를 그리고, "보기 좋은"
   눈금의 격자선·왼쪽 축 라벨을 같이 그려서 값을 가늠하기 쉽게 합니다.
   onclickFn(key)가 null을 돌려주면 그 막대는 클릭 비활성. */
function buildTrendChart(keys,m,unit,labelFn,onclickFn,titleFn){
  const vals=keys.map(k=>{const b=m[k];return b?(unit==='cap'?b.cap:b.cnt):0;});
  const rawMax=Math.max(0,...vals);
  const ticks=niceTicks(rawMax,unit==='cnt');
  const scaleTop=ticks[ticks.length-1]||1;
  const peakV=Math.max(0,...vals);
  const axis=ticks.map(t=>`<span class="trendaxistick" style="bottom:${t/scaleTop*100}%">${fmtTick(t,unit)}</span>`).join('');
  const grid=ticks.map(t=>`<div class="trendgridline" style="bottom:${t/scaleTop*100}%"></div>`).join('');
  const bars=keys.map((k,i)=>{
    const v=vals[i];
    const h=v>0?Math.max(3,v/scaleTop*100):0;
    const isPeak=v>0&&v===peakV;
    const act=onclickFn(k);
    const label=isPeak?`<span class="trendlabel">${fmtTick(v,unit)}</span>`:'';
    return `<div class="trendcol${isPeak?' peak':''}"${act?` onclick="${act}"`:''} title="${esc(titleFn(k))}">${label}<div class="trendbar" style="height:${h}%"></div></div>`;
  }).join('');
  const tickStep=Math.max(1,Math.ceil(keys.length/9));
  const xticks=keys.map((k,i)=>`<span class="trendtick">${i%tickStep===0?esc(labelFn(k)):''}</span>`).join('');
  const totalCnt=keys.reduce((s,k)=>s+(m[k]?m[k].cnt:0),0);
  const totalCap=keys.reduce((s,k)=>s+(m[k]?m[k].cap:0),0);
  const chart=`<div class="trendplot">
      <div class="trendaxis">${axis}</div>
      <div class="trendarea"><div class="trendgrid">${grid}</div><div class="trendchart">${bars}</div></div>
    </div>
    <div class="trendxrow"><div class="trendaxisspacer"></div><div class="trendticks">${xticks}</div></div>`;
  return {chart,totalCnt,totalCap};
}
/* 월별 추이 — 선/영역 차트. 최근 12개월(실선, 채워진 영역) 위에 정확히
   1년 전 같은 12개월(점선, 무채색)을 겹쳐서 전년 동기와 비교합니다.
   작년 구간에 데이터가 하나도 없으면(워크북 역사가 짧은 경우) 비교선은
   조용히 숨기고 이번 구간만 보여줍니다 — 텅 빈 선을 그리지 않습니다. */
function buildTrendLineChart(cfg,m,unit,curKeys){
  const prevKeys=curKeys.map(ym=>idxYm(ymIdx(ym)-12));
  const valOf=k=>{const b=m[k];return b?(unit==='cap'?b.cap:b.cnt):0;};
  const curVals=curKeys.map(valOf);
  const prevVals=prevKeys.map(valOf);
  const hasPrev=prevVals.some(v=>v>0);
  const rawMax=Math.max(0,...curVals,...(hasPrev?prevVals:[]));
  const ticks=niceTicks(rawMax,unit==='cnt');
  const scaleTop=ticks[ticks.length-1]||1;
  const axis=ticks.map(t=>`<span class="trendaxistick" style="bottom:${t/scaleTop*100}%">${fmtTick(t,unit)}</span>`).join('');
  const grid=ticks.map(t=>`<div class="trendgridline" style="bottom:${t/scaleTop*100}%"></div>`).join('');

  const W=1200,H=180,padX=16,n=curKeys.length;
  const xAt=i=>n>1?padX+i*((W-2*padX)/(n-1)):W/2;
  const yAt=v=>H-(scaleTop>0?Math.min(1,v/scaleTop)*H:0);
  const linePath=vals=>vals.map((v,i)=>(i===0?'M':'L')+xAt(i).toFixed(1)+','+yAt(v).toFixed(1)).join(' ');
  const areaPath=vals=>linePath(vals)+` L${xAt(n-1).toFixed(1)},${H} L${xAt(0).toFixed(1)},${H} Z`;
  const dots=(vals,keysArr,color,clickable)=>vals.map((v,i)=>{
    const cx=xAt(i).toFixed(1),cy=yAt(v).toFixed(1),k=keysArr[i];
    const b=m[k]||{cnt:0,cap:0};
    const title=monthLabel(k)+' · '+nf(b.cnt,0)+'건 · '+nf(b.cap)+'MW'+(clickable?' · 눌러서 일별 보기':'');
    const act=(clickable&&b.cnt>0)?` style="cursor:pointer" onclick="drillMonth('${jsq(k)}')"`:'';
    return `<circle cx="${cx}" cy="${cy}" r="10" fill="transparent"${act}><title>${esc(title)}</title></circle>
      <circle cx="${cx}" cy="${cy}" r="4" fill="${color}" stroke="var(--panel)" stroke-width="2" style="pointer-events:none"/>`;
  }).join('');
  const curEnd=curVals[n-1],curEndY=yAt(curEnd);
  const endLabel=`<text class="trendendlabel" x="${xAt(n-1).toFixed(1)}" y="${Math.max(11,curEndY-10).toFixed(1)}" text-anchor="end">${esc(fmtTick(curEnd,unit))}</text>`;

  let svg=`<svg class="trendsvg" viewBox="0 0 ${W} ${H}" preserveAspectRatio="none">`;
  if(hasPrev){
    svg+=`<path d="${areaPath(prevVals)}" fill="var(--mute)" opacity=".10" stroke="none"/>`;
    svg+=`<path d="${linePath(prevVals)}" fill="none" stroke="var(--mute)" stroke-width="2" stroke-linejoin="round" stroke-linecap="round" stroke-dasharray="5 4" vector-effect="non-scaling-stroke"/>`;
  }
  svg+=`<path d="${areaPath(curVals)}" fill="var(--teal)" opacity=".14" stroke="none"/>`;
  svg+=`<path d="${linePath(curVals)}" fill="none" stroke="var(--teal)" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round" vector-effect="non-scaling-stroke"/>`;
  if(hasPrev) svg+=dots(prevVals,prevKeys,'var(--mute)',false);
  svg+=dots(curVals,curKeys,'var(--teal)',true);
  svg+=endLabel;
  svg+='</svg>';

  /* 범위가 12개월을 넘거나 연도 경계를 넘나들면 "3월"만으론 헷갈리므로
     "'26.03" 형식(monthLabel)으로, 점 개수가 많을 때는 겹치지 않게
     솎아서(tickStep) 보여줍니다. */
  const tickStep=Math.max(1,Math.ceil(n/12));
  const xticks=curKeys.map((k,i)=>{
    const show=i%tickStep===0||i===n-1;
    const lbl=show?(n>12?monthLabel(k):Number(k.slice(5,7))+'월'):'';
    return `<span class="trendtick">${esc(lbl)}</span>`;
  }).join('');
  const rangeLabel=n===12&&!(state.homeTrend.rangeFrom&&state.homeTrend.rangeTo)?'최근 12개월':`${monthLabel(curKeys[0])} ~ ${monthLabel(curKeys[n-1])}`;
  const legend=hasPrev?`<div class="trendlegend">
      <span><span class="trendlegendline"></span>${esc(rangeLabel)}</span>
      <span><span class="trendlegendline prev"></span>전년 동기</span>
    </div>`:'';
  const chart=`<div class="trendplot">
      <div class="trendaxis">${axis}</div>
      <div class="trendarea"><div class="trendgrid">${grid}</div>${svg}</div>
    </div>
    <div class="trendxrow"><div class="trendaxisspacer"></div><div class="trendticks">${xticks}</div></div>
    ${legend}`;
  const totalCnt=curKeys.reduce((s,k)=>s+(m[k]?m[k].cnt:0),0);
  const totalCap=curKeys.reduce((s,k)=>s+(m[k]?m[k].cap:0),0);
  const prevTotalCnt=prevKeys.reduce((s,k)=>s+(m[k]?m[k].cnt:0),0);
  return {chart,totalCnt,totalCap,hasPrev,prevTotalCnt};
}
function trendToolsHtml(avail,unit){
  return `<div class="trendtools">
      <div class="segtoggle">${avail.map(([k,c])=>
        `<button class="${state.homeTrend.metric===k?'on':''}" onclick="setHomeTrend('metric','${k}')">${esc(c.label)}</button>`).join('')}</div>
      <div class="segtoggle">
        <button class="${unit==='cnt'?'on':''}" onclick="setHomeTrend('unit','cnt')">건수</button>
        <button class="${unit==='cap'?'on':''}" onclick="setHomeTrend('unit','cap')">용량(MW)</button>
      </div>
    </div>`;
}
/* 조회 시작/종료 년월 필터 — <input type=month>는 브라우저 기본 달력
   선택기와 직접 타이핑을 동시에 지원합니다(공급기한 입력과 동일한 이유로
   선택). 값이 바뀌면 즉시 다시 그립니다. */
function trendRangeToolsHtml(){
  const {rangeFrom,rangeTo}=state.homeTrend;
  const custom=!!(rangeFrom&&rangeTo);
  const clamped=custom&&(Math.abs(ymIdx(rangeTo)-ymIdx(rangeFrom))>35);
  return `<div class="trendrange">
      <label>조회 시작<input type="month" class="trendrangeinput" value="${esc(rangeFrom||'')}" onchange="setHomeTrend('rangeFrom',this.value)"></label>
      <span class="trendrangesep">~</span>
      <label>조회 종료<input type="month" class="trendrangeinput" value="${esc(rangeTo||'')}" onchange="setHomeTrend('rangeTo',this.value)"></label>
      ${custom?`<button class="btn" onclick="resetHomeTrendRange()">최근 12개월로</button>`:''}
      ${clamped?`<span class="trendrangenote">한 번에 최대 36개월까지 표시됩니다</span>`:''}
    </div>`;
}
function trendPanel(){
  const avail=Object.entries(TREND_METRICS).filter(([,c])=>byKey[c.tk]);
  if(!avail.length) return '';
  if(!byKey[TREND_METRICS[state.homeTrend.metric]?.tk]) state.homeTrend.metric=avail[0][0];
  const cfg=TREND_METRICS[state.homeTrend.metric];
  const unit=state.homeTrend.unit;
  const tools=trendToolsHtml(avail,unit);

  if(state.homeTrend.drillYm){
    const ym=state.homeTrend.drillYm;
    const m=groupByDay(cfg.tk,cfg.col,cfg.cap);
    const days=daysOfMonth(ym);
    const {chart,totalCnt,totalCap}=buildTrendChart(days,m,unit,
      d=>dayLabel(d),
      d=>(m[d]&&m[d].cnt>0)?`jumpToDay('${jsq(cfg.tk)}','${jsq(cfg.col)}','${jsq(d)}')`:null,
      d=>{const b=m[d]||{cnt:0,cap:0};return d+' · '+nf(b.cnt,0)+'건 · '+nf(b.cap)+'MW';});
    const ymLabel=ym.slice(0,4)+'년 '+Number(ym.slice(5,7))+'월';
    return panel('년월별 추이','일별 막대를 누르면 그 날짜만 보기로 이동합니다',`
      ${tools}
      <div class="trendbreadcrumb">
        <button class="trendback" onclick="undrillMonth()">← 월별로</button>
        <span class="trendcrumbnow">${esc(ymLabel)} · 일별 (${esc(cfg.label)})</span>
        <button class="btn" style="margin-left:auto" onclick="jumpToMonth('${jsq(cfg.tk)}','${jsq(cfg.col)}','${jsq(ym)}')">이 달 전체를 표에서 보기</button>
      </div>
      ${chart}
      <div class="trendfoot"><span>${esc(ymLabel)} · ${days.length}일</span><span>합계 ${nf(totalCnt,0)}건 · ${nf(totalCap)} MW</span></div>`);
  }

  const m=groupByMonth(cfg.tk,cfg.col,cfg.cap);
  const keys=Object.keys(m);
  const rangeTools=trendRangeToolsHtml();
  if(!keys.length) return panel('년월별 추이','',`<div class="nocand">${esc(cfg.label)} 항목에 날짜가 입력된 데이터가 없습니다.</div>
    <div style="margin-top:10px">${tools}</div>`);
  const curKeys=trendRangeKeys();
  const {chart,totalCnt,totalCap,hasPrev,prevTotalCnt}=buildTrendLineChart(cfg,m,unit,curKeys);
  const isCustom=!!(state.homeTrend.rangeFrom&&state.homeTrend.rangeTo);
  const rangeDesc=isCustom?`${monthLabel(curKeys[0])} ~ ${monthLabel(curKeys[curKeys.length-1])} (${curKeys.length}개월)`:'최근 12개월';
  const yoy=hasPrev
    ?(prevTotalCnt>0?` · 전년 동기 대비 ${(totalCnt-prevTotalCnt)>=0?'+':''}${nf(totalCnt-prevTotalCnt,0)}건`:'')
    :' · 전년 동기 데이터 없음';
  return panel('년월별 추이','선 위의 점을 누르면 그 달의 일별 추이를 볼 수 있습니다',`
    ${tools}
    ${rangeTools}
    ${chart}
    <div class="trendfoot"><span>${esc(rangeDesc)}${esc(yoy)}</span><span>합계 ${nf(totalCnt,0)}건 · ${nf(totalCap)} MW</span></div>`);
}

/* ── 홈: 한눈에 보기 인사이트 문장 ──────────────────────────────────────── */
/* 공급기한 컬럼에서 "아직 지나지 않은 것 중 가장 임박한" 한 건을 찾음 —
   "지금 당장 뭘 봐야 하는지"를 숫자 하나로 압축해 보여주는 용도. */
function nearestDeadline(tk,col){
  const t=byKey[tk];if(!t||!TODAY) return null;
  let best=null;
  t.rows.forEach(r=>{
    const v=String(r.cells[col]||'');if(!/^\d{4}-\d{2}-\d{2}$/.test(v)) return;
    const d=daysBetween(TODAY,v);
    if(d>=0&&(best===null||d<best.days)) best={days:d,pk:String(r.cells[t.pk]??'')};
  });
  return best;
}
/* 한 문단짜리 문장을 쭉 읽어야 했던 예전 방식 대신, 숫자·아이콘 중심으로
   3초 안에 훑을 수 있는 3단 구성으로 바꿨습니다:
     1) 핵심 지표 칩 — 발전소/구매/판매/매칭률/최근 추이를 아이콘+숫자로
     2) 현황 아이콘 스트립 — 수급매칭 8개 현황 각각의 현재 건수(0건은 흐리게)
     3) 지금 확인할 것 — 이슈 발생·미확보 건수와 가장 임박한 공급기한을
        한 줄로 모아, 문제가 있으면 눈에 띄게(흰 배경+빨간 글씨) 강조 */
function homeInsight(){
  const P=byKey['T_발전소'],B=byKey['T_구매계약'],S=byKey['T_판매계약'],M=byKey['T_수급매칭'];
  const supplyMW=P?sumCol('T_발전소','설비용량(MW)'):0;
  const bs=B?sumCapSplit('T_구매계약','구매계약용량(MW)',purchaseTerminatedIds()):null;
  const ss=S?sumCapSplit('T_판매계약','판매계약용량(MW)',saleTerminatedIds()):null;
  if(!P&&!bs&&!ss) return '';

  const chips=[];
  if(P) chips.push({icon:'🏭',label:'발전소',value:nf(P.rows.length,0)+'개',sub:nf(supplyMW)+'MW'});
  if(bs) chips.push({icon:'🧾',label:'구매계약',value:nf(bs.activeN,0)+'건',sub:nf(bs.activeMW)+'MW 유효'});
  if(ss) chips.push({icon:'🧾',label:'판매계약',value:nf(ss.activeN,0)+'건',sub:nf(ss.activeMW)+'MW 유효'});
  if(bs&&supplyMW>0){
    const pct=nf(Math.min(999,bs.activeMW/supplyMW*100),0);
    chips.push({icon:'📊',label:'설비 대비 매칭률',value:pct+'%',sub:'',
      title:'발전소 설비용량 합계 대비 구매계약 유효 용량의 비율 - 100%에 가까울수록 설비 여유가 없다는 뜻입니다.'});
  }
  if(S){
    const m=groupByMonth('T_판매계약','계약일','판매계약용량(MW)');
    let cnt=0,cap=0;
    monthRange(3).forEach(ym=>{if(m[ym]){cnt+=m[ym].cnt;cap+=m[ym].cap;}});
    if(cnt>0) chips.push({icon:'📈',label:'최근 3개월 신규 판매',value:nf(cnt,0)+'건',sub:nf(cap)+'MW'});
  }
  const chipHtml=chips.map(c=>`<span class="ichip"${c.title?` title="${esc(c.title)}"`:''}>
      <span class="ichipicon">${c.icon}</span>
      <span class="ichiptext"><b>${esc(c.value)}</b>${c.sub?`<span class="dim">${esc(c.sub)}</span>`:''}<span class="ichiplabel">${esc(c.label)}</span></span>
    </span>`).join('');

  const counts=M?statusCountsAll():{};
  const statusStripHtml=(M&&M.rows.length)?`<div class="istatusrow">${
    Object.entries(STATUS_META).map(([label,meta])=>{
      const n=counts[label]||0;
      return `<span class="ischip${n?'':' zero'}" title="${esc(label)}: ${nf(n,0)}건 / 전체 ${nf(M.rows.length,0)}건">${meta.icon} <b>${nf(n,0)}</b></span>`;
    }).join('')
  }</div>`:'';

  /* 잔여 공급기한 — 두 표를 통틀어 가장 임박한(D값이 가장 작은) 한 건만
     짚어서 "지금 당장 뭘 챙겨야 하는지"를 바로 알 수 있게 합니다. */
  const nb=B?nearestDeadline('T_구매계약','공급기한_구매'):null;
  const ns=S?nearestDeadline('T_판매계약','공급기한_판매'):null;
  const nearest=[nb&&{...nb,label:'구매계약'},ns&&{...ns,label:'판매계약'}].filter(Boolean).sort((a,b)=>a.days-b.days)[0];
  const issueN=counts['이슈 발생']||0;
  const unsecuredN=counts['미확보']||0;
  const alertParts=[];
  if(issueN>0) alertParts.push(`⚠️ 이슈 발생 <b>${nf(issueN,0)}건</b>`);
  if(unsecuredN>0) alertParts.push(`❓ 미확보 <b>${nf(unsecuredN,0)}건</b>`);
  if(nearest) alertParts.push(`⏰ 가장 임박한 공급기한 ${esc(nearest.label)} <b>${esc(nearest.pk)}</b> D-<b>${nf(nearest.days,0)}</b>`);
  const hasUrgent=issueN>0||unsecuredN>0;
  const alertHtml=alertParts.length
    ?`<div class="ialert${hasUrgent?' urgent':''}">${alertParts.join(' · ')}</div>`
    :`<div class="ialert ok">✅ 지금 확인이 필요한 이슈가 없습니다.</div>`;

  return `<div class="insight"><p class="ieyebrow">한눈에 보기</p>
    <div class="ichiprow">${chipHtml}</div>
    ${statusStripHtml}
    ${alertHtml}</div>`;
}

function tHome(){
  const P=byKey['T_발전소'],B=byKey['T_구매계약'],S=byKey['T_판매계약'];
  const supplyMW=P?sumCol('T_발전소','설비용량(MW)'):0;
  /* 메인 용량 KPI는 "계약 종료" 건(수급매칭 현황이 전부 종료류인 계약)을
     빼고 집계하고, 뺀 만큼은 아래 "종료/만료 용량" 카드로 따로 보여줍니다
     — sumCapSplit()이 활성/종료 두 값을 한 번에 계산해서 서로 앞뒤가
     맞게 유지합니다. */
  const purchSplit=B?sumCapSplit('T_구매계약','구매계약용량(MW)',purchaseTerminatedIds()):{activeN:0,activeMW:0,termN:0,termMW:0};
  const saleSplit=S?sumCapSplit('T_판매계약','판매계약용량(MW)',saleTerminatedIds()):{activeN:0,activeMW:0,termN:0,termMW:0};
  const purchMW=purchSplit.activeMW,saleMW=saleSplit.activeMW;
  const termN=purchSplit.termN+saleSplit.termN,termMW=purchSplit.termMW+saleSplit.termMW;
  const bUn=B?countWhere('T_구매계약','수요기업 미확보','TRUE'):0;
  const sUn=S?countWhere('T_판매계약','공급자원 미확보','TRUE'):0;
  const bExp=B?countExpiring('T_구매계약','공급기한_구매',SOON_DAYS):0;
  const sExp=S?countExpiring('T_판매계약','공급기한_판매',SOON_DAYS):0;
  const mix=P?groupSum('T_발전소','발전원','설비용량(MW)'):[];
  const mixMax=mix.length?mix[0][1]:0;
  const mixBars=mix.map(([g,v],i)=>{
    const col=CAT_COLORS[i%CAT_COLORS.length];
    return `<div class="mixrow"><span class="mixdot" style="background:${col}"></span><span class="mixlabel">${esc(g)}</span>
      <div class="mixbar"><div class="mixfill" style="width:${mixMax>0?(v/mixMax*100):0}%;background:${col}"></div></div>
      <span class="mixval mono">${nf(v)} MW (${mixMax>0?nf(v/mix.reduce((s,x)=>s+x[1],0)*100,0):0}%)</span></div>`;
  }).join('')||'<div class="nocand">데이터 없음</div>';
  const ok=DATA.validation.total_errors===0;
  const balance=purchMW-saleMW;
  const chgTotal=CHANGES.has_prev?(CHANGES.total_added+CHANGES.total_changed+CHANGES.total_removed):0;

  return `<section>${printHead('PPA 계약관리 현황 요약','')}
    <div class="homehead"><span class="sub mono">기준 시각: ${esc((DATA.generated_at||'').replace('T',' '))}</span>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <button class="btn primary" onclick="copySummary()">요약 복사 (보고용)</button>
        <details class="drop"><summary>전체 내려받기</summary><div class="dropbody right">
          <button class="dlopt" onclick="downloadAllXlsx()">Excel (.xlsx)<small>6개 표 전체 · 시트별 분리</small></button>
          <button class="dlopt" onclick="downloadSummaryMd()">Markdown (.md)<small>보고서용 요약 + 표별 건수</small></button>
        </div></details>
        <button class="btn" onclick="window.print()">인쇄 / PDF</button>
      </div></div>
    ${homeInsight()}
    <div class="kpis">
    ${kpi('발전소 설비용량 합계',nf(supplyMW)+' MW',(P?nf(P.rows.length,0)+'개 발전소':'')+deltaBadge(capacityDelta('T_발전소','설비용량(MW)')),'accent',"state.tab='T_발전소';render()")}
    ${kpi('구매계약 총 용량 (유효)'+infoTip('공급기한과 무관하게, 수급매칭 현황이 전부 "종료"류인 계약만 뺀 값입니다.'),nf(purchMW)+' MW',(B?(purchSplit.termN?`${nf(purchSplit.activeN,0)}건 · 종료 ${nf(purchSplit.termN,0)}건 제외`:nf(purchSplit.activeN,0)+'건'):'')+deltaBadge(capacityDelta('T_구매계약','구매계약용량(MW)')),'',"state.tab='T_구매계약';render()")}
    ${kpi('판매계약 총 용량 (유효)'+infoTip('수급매칭 현황이 전부 "종료"류인 계약은 제외한 값입니다.'),nf(saleMW)+' MW',(S?(saleSplit.termN?`${nf(saleSplit.activeN,0)}건 · 종료 ${nf(saleSplit.termN,0)}건 제외`:nf(saleSplit.activeN,0)+'건'):'')+deltaBadge(capacityDelta('T_판매계약','판매계약용량(MW)')),'',"state.tab='T_판매계약';render()")}
    ${kpi('구매 − 판매 밸런스'+infoTip('구매계약 유효 용량에서 판매계약 유효 용량을 뺀 값 — 양수면 구매(공급)가 판매(수요)보다 여유 있다는 뜻입니다.'),(balance>=0?'+':'')+nf(balance)+' MW',balance>=0?'구매 우위(여유)':'판매 우위(부족)',balance<0?'warn':'')}
    ${kpi('종료/만료 용량'+infoTip('메인 용량 지표에서 제외된 "계약 종료" 건들의 용량을 모은 값입니다 — 없어진 게 아니라 여기로 옮겨 보이는 것입니다.'),nf(termMW)+' MW',termN?`구매 ${purchSplit.termN}건 · 판매 ${saleSplit.termN}건 (메인 지표 제외됨)`:'해당 없음','')}
    </div>
    <div class="kpis">
    ${kpi('검증 오류',nf(DATA.validation.total_errors,0),ok?'전 표 정상':'클릭해서 확인',ok?'':'warn',"state.tab='검증';render()")}
    ${kpi('지난 대비 변경',CHANGES.has_prev?nf(chgTotal,0):'—',
      CHANGES.has_prev?`추가 ${CHANGES.total_added} · 수정 ${CHANGES.total_changed} · 삭제 ${CHANGES.total_removed}`:'이전 스냅샷 없음',
      '',CHANGES.has_prev?"state.tab='변경';render()":'')}
    ${kpi('미확보 계약',nf(bUn+sUn,0),`구매 ${bUn} · 판매 ${sUn}`,(bUn+sUn)>0?'warn':'')}
    ${kpi(SOON_DAYS+'일 내 공급기한',nf(bExp+sExp,0),`구매 ${bExp} · 판매 ${sExp}`,'')}
    </div>
    ${statusPanel()}
    ${trendPanel()}
    <div class="grid2">
      ${panel('발전원별 설비용량 비중','',mixBars)}
      ${actionItemsPanel()}
    </div>
    ${capacityGapPanel()}
    ${schemaDiagram()?panel('표 관계 구조','박스를 누르면 해당 표로 이동합니다',schemaDiagram()):''}</section>`;
}

/* ── 변경 탭 ────────────────────────────────────────────────────────────── */
function tChanges(){
  if(!CHANGES.has_prev)
    return `<section><div class="panel"><div class="ph"><h3>이전 스냅샷이 없습니다</h3></div>
      <div class="nocand">대시보드를 만들 때마다 그 시점의 데이터가 <b>스냅샷 파일</b>로 옆에 저장됩니다.
      <b>다음번에 다시 생성하면</b> 이 탭에 "지난번 대비 추가 / 수정 / 삭제된 항목"이 자동으로 표시됩니다.</div></div></section>`;
  const rowsBy=(k)=>(CHANGES.summary||{})[k]||{added:0,removed:0,changed:0};
  const tableRows=DATA.tables.map(t=>{
    const s=rowsBy(t.key);
    const tot=s.added+s.changed+s.removed;
    return `<tr class="${tot?'clickrow':''}" ${tot?`onclick="state.tab='${jsq(t.key)}';state.onlyChg['${jsq(t.key)}']=true;state.page['${jsq(t.key)}']=1;render()"`:''}>
      <td>${esc(t.label)}</td><td class="num">${s.added}</td><td class="num">${s.changed}</td><td class="num">${s.removed}</td></tr>`;}).join('');
  const details=(CHANGES.details||[]).map(d=>{
    const t=byKey[d.table];
    return `<div class="chk chg" onclick="jumpTo('${jsq(d.table)}','${jsq(d.pk)}')">
      <span>${esc(t?t.label:d.table)} · ${esc(d.pk)} <span class="dkey">${esc(d.col)}</span></span>
      <span class="chgval"><span class="chgold">${esc(d.old||'(공란)')}</span>→ <span class="chgnew">${esc(d.new||'(공란)')}</span></span>
      <span></span><span class="badge info">수정</span></div>`;}).join('')
    ||'<div class="nocand">값이 수정된 항목은 없습니다.</div>';
  const trunc=CHANGES.truncated?`<div class="nocand">표시 한도를 넘어 일부만 보여줍니다 (집계 건수는 정확합니다).</div>`:'';
  const removed=DATA.tables.map(t=>{
    const list=(CHANGES.removed_rows||{})[t.key]||[];
    if(!list.length) return '';
    return list.map((cells,idx)=>`<div class="chk del" onclick="openRemovedDetail('${jsq(t.key)}',${idx})">
      <span>${esc(t.label)} · ${esc(cells[t.pk]||'(PK 공란)')}</span>
      <span class="chgval">${esc(t.columns.slice(1,4).map(c=>cells[c]).filter(Boolean).join(' · '))}</span>
      <span></span><span class="badge mute">삭제 · 클릭해서 전체 보기</span></div>`).join('');}).join('')
    ||'<div class="nocand">삭제된 항목은 없습니다.</div>';
  const added=DATA.tables.map(t=>
    t.rows.map((r,i)=>({r,i})).filter(({r})=>r.change==='added').map(({r,i})=>
      `<div class="chk add" onclick="openDetail('${jsq(t.key)}',${i})">
        <span>${esc(t.label)} · ${esc(r.cells[t.pk]||'')}</span>
        <span class="chgval">${esc(t.columns.slice(1,3).map(c=>r.cells[c]).filter(Boolean).join(' · '))}</span>
        <span></span><span class="badge ok">추가</span></div>`).join('')).join('')
    ||'<div class="nocand">추가된 항목은 없습니다.</div>';

  return `<section>${printHead('변경 내역','기준: '+esc((CHANGES.prev_generated_at||'').replace('T',' '))+' → '+esc((DATA.generated_at||'').replace('T',' ')))}
    <div style="margin:-4px 0 12px">
      <button class="btn" onclick="openInlineResetBaseline()" title="지금 시점을 새 비교 기준으로 리셋합니다(전체 변경 이력은 유지됨)">변경 비교 기준 리셋</button>
    </div>
    <div class="kpis">
      ${kpi('추가',nf(CHANGES.total_added,0),'새로 생긴 행','')}
      ${kpi('수정',nf(CHANGES.total_changed,0),'값이 바뀐 행 ('+nf((CHANGES.details||[]).length,0)+'개 항목)','')}
      ${kpi('삭제',nf(CHANGES.total_removed,0),'사라진 행','')}
      ${kpi('비교 기준',esc((CHANGES.prev_generated_at||'—').slice(0,16).replace('T',' ')),'직전 생성 시점','')}
    </div>
    ${panel('표별 변경 건수(행 기준)','행을 클릭하면 그 표의 변경분만 보기로 이동합니다',
      `<table><thead><tr><th class="nosort">표</th><th class="nosort num">추가</th><th class="nosort num">수정</th><th class="nosort num">삭제</th></tr></thead><tbody>${tableRows}</tbody></table>`)}
    ${panel(`수정된 항목 (이전값 → 새값) · ${nf((CHANGES.details||[]).length,0)}개`,'클릭하면 관계조회로 이동',details+trunc)}
    <div class="grid2">${panel('추가된 항목','',added)}${panel('삭제된 항목','',removed)}</div>
    ${panel(`전체 변경 이력 · 최근 ${nf(CHANGELOG.length,0)}건 (최대 1,000건 보존)`,
      '여러 번의 생성에 걸쳐 계속 쌓입니다 — 이번 생성분만이 아니라 지금까지의 추가/수정/삭제를 검색할 수 있습니다',
      changelogView())}
    </section>`;
}

/* ── 전체 변경 이력(여러 생성에 걸쳐 누적, 최대 1,000건) ──────────────────── */
function changelogFiltered(){
  const q=(state.clog.q||'').trim();
  const kind=state.clog.kind||'';
  const table=state.clog.table||'';
  return CHANGELOG.filter(e=>{
    if(kind&&e.kind!==kind) return false;
    if(table&&e.table!==table) return false;
    if(q){
      const hay=[e.pk,byKey[e.table]?byKey[e.table].label:e.table,e.actor||'']
        .concat(e.changed_cols||[]).concat(Object.values(e.cells||{}))
        .join(' ');
      if(!matchesSearch(hay,q)) return false;
    }
    return true;
  }).slice().reverse();
}
function changelogRow(e){
  const t=byKey[e.table];
  const label=t?t.label:e.table;
  const exists=rowIndex[e.table]&&rowIndex[e.table][String(e.pk)]!==undefined;
  const kindBadge=e.kind==='added'?'<span class="badge ok">추가</span>'
    :e.kind==='removed'?'<span class="badge mute">삭제</span>'
    :'<span class="badge info">수정</span>';
  const cls=e.kind==='added'?'add':(e.kind==='removed'?'del':'chg');
  let detail;
  if(e.kind==='changed'){
    const cols=e.changed_cols||[];
    detail=cols.slice(0,3).map(c=>
      `<span class="dkey">${esc(c)}</span> <span class="chgold">${esc((e.prev||{})[c]||'(공란)')}</span>→<span class="chgnew">${esc((e.cells||{})[c]||'(공란)')}</span>`
    ).join(' · ')+(cols.length>3?` 외 ${cols.length-3}건`:'');
  } else {
    const cols=(t?t.columns:[]).filter(c=>c!==(t?t.pk:'')).slice(0,3);
    detail=cols.map(c=>(e.cells||{})[c]).filter(Boolean).join(' · ');
  }
  const clickable=exists?` onclick="jumpTo('${jsq(e.table)}','${jsq(e.pk)}')"`:' style="cursor:default"';
  const actorTag=e.actor?`<span class="dkey" style="margin-left:8px">👤 ${esc(e.actor)}</span>`:'';
  return `<div class="chk ${cls}"${clickable}>
    <span>${esc(label)} · ${esc(e.pk||'(PK 공란)')}<span class="dkey" style="margin-left:8px">${esc((e.generated_at||'').replace('T',' ').slice(0,16))}</span>${actorTag}</span>
    <span class="chgval">${detail}</span>
    <span></span>${kindBadge}</div>`;
}
function setClogQ(v){state.clog.q=v;state.page.clog=1;render();}
function setClogKind(v){state.clog.kind=v;state.page.clog=1;render();}
function setClogTable(v){state.clog.table=v;state.page.clog=1;render();}
function changelogView(){
  if(!CHANGELOG.length) return '<div class="nocand">아직 쌓인 이력이 없습니다 — 다음 생성부터 여기 표시됩니다.</div>';
  const kOpts=[['','전체 종류'],['added','추가'],['changed','수정'],['removed','삭제']]
    .map(([v,l])=>`<option value="${v}" ${state.clog.kind===v?'selected':''}>${l}</option>`).join('');
  const tOpts=['<option value="">전체 표</option>'].concat(
    DATA.tables.map(t=>`<option value="${jsq(t.key)}" ${state.clog.table===t.key?'selected':''}>${esc(t.label)}</option>`)).join('');
  const rows=changelogFiltered();
  const total=rows.length;
  const pages=Math.max(1,Math.ceil(total/state.pageSize));
  const page=Math.min(Math.max(1,state.page.clog||1),pages);
  state.page.clog=page;
  const pageRows=rows.slice((page-1)*state.pageSize,page*state.pageSize);
  const list=pageRows.map(changelogRow).join('')||'<div class="nocand">조건에 맞는 이력이 없습니다.</div>';
  return `<div class="toolbar">
      <input id="clog-q" class="search" placeholder="PK·값으로 이력 검색…" value="${esc(state.clog.q)}" oninput="onSearchType(event,setClogQ)">
      <select class="filtersel" onchange="setClogKind(this.value)">${kOpts}</select>
      <select class="filtersel" onchange="setClogTable(this.value)">${tOpts}</select>
      <span class="count">${nf(total,0)} / ${nf(CHANGELOG.length,0)}건</span>
    </div>
    <div>${list}</div>
    ${pager({key:'clog'},total,page,pages)}`;
}

/* ── 검증 탭 ────────────────────────────────────────────────────────────── */
function tVerify(){
  const v=DATA.validation;
  const byTable=Object.entries(v.by_table).map(([k,c])=>
    `<tr class="clickrow" onclick="state.tab='${jsq(k)}';state.onlyErr['${jsq(k)}']=true;state.page['${jsq(k)}']=1;render()">
      <td>${esc(byKey[k]?byKey[k].label:k)}</td><td class="num">${c}</td></tr>`).join('')
    ||'<tr><td colspan="2" class="emptyrow">오류 없음</td></tr>';
  const byItem=Object.entries(v.by_item).map(([k,c])=>
    `<tr><td>${esc(k)}</td><td class="num">${c}</td></tr>`).join('')
    ||'<tr><td colspan="2" class="emptyrow">오류 없음</td></tr>';
  const detail=v.errors.map(e=>{
    const pkv=e.pk_value;
    const act=pkv?`jumpTo('${jsq(e.table)}','${jsq(pkv)}')`
      :`state.tab='${jsq(e.table)}';state.onlyErr['${jsq(e.table)}']=true;render()`;
    return `<div class="chk no" onclick="${act}">
      <span>${esc(byKey[e.table]?byKey[e.table].label:e.table)} · PK=${esc(pkv||'(공란)')}</span>
      <span class="mono">행 ${e.row_index+1}</span><span>${esc(e.error_item)}</span>
      <span class="badge no">클릭해서 확인</span></div>`;}).join('')
    ||'<div class="nocand" style="text-align:center;padding:20px">오류가 없습니다.</div>';
  return `<section>${printHead('검증 결과','총 '+v.total_errors+'건')}
    <div class="kpis">${kpi('총 오류 건수',nf(v.total_errors,0),v.total_errors>0?'표별 세부는 아래':'전 표 정상',v.total_errors>0?'warn':'accent')}</div>
    <div class="grid2">
      ${panel('표별 오류 건수','행 클릭 시 오류만 보기',`<table><thead><tr><th class="nosort">표</th><th class="nosort num">건수</th></tr></thead><tbody>${byTable}</tbody></table>`)}
      ${panel('오류항목별 건수','',`<table><thead><tr><th class="nosort">오류항목</th><th class="nosort num">건수</th></tr></thead><tbody>${byItem}</tbody></table>`)}
    </div>
    ${panel('상세 오류 목록','행을 클릭하면 문제의 레코드로 이동합니다',detail)}</section>`;
}

/* ── 전역 검색 ──────────────────────────────────────────────────────────── */
function runGlobalSearch(v){
  const box=document.getElementById('globalResults');
  const q=(v||'').trim();
  if(!q){box.classList.remove('show');box.innerHTML='';return;}
  const res=[];
  DATA.tables.forEach(t=>t.rows.forEach((r,i)=>{
    if(matchesSearch(t.columns.map(c=>r.cells[c]??'').join(' '),q)){
      res.push({table:t.key,tlabel:t.label,pk:r.cells[t.pk],idx:i,
        text:t.columns.slice(0,3).map(c=>r.cells[c]).filter(x=>x!==undefined&&x!=='').join(' · ')});}}));
  const shown=res.slice(0,20);
  box.innerHTML=shown.length
    ? shown.map(r=>`<button class="gresrow" onclick="closeGlobalSearch();openDetail('${jsq(r.table)}',${r.idx})">
        <span class="tag">${esc(r.tlabel)}</span>${esc(r.text)}</button>`).join('')
      +(res.length>20?`<div class="nocand" style="padding:8px 16px">${nf(res.length-20,0)}건 더 있음 — 검색어를 좁혀보세요.</div>`:'')
    : '<div class="nocand" style="padding:10px 16px">일치하는 항목이 없습니다.</div>';
  box.classList.add('show');
}
/* 즉시 호출용(포커스 시 이전 검색 결과 다시 보여주기) — 타이핑 중엔
   onSearchType(공용 디바운스+IME 안전 핸들러, 위 함수 정의 참고)을 씁니다. */
function onGlobalSearch(v){runGlobalSearch(v);}
function closeGlobalSearch(){
  const i=document.getElementById('globalSearch');if(i) i.value='';
  const b=document.getElementById('globalResults');if(b){b.classList.remove('show');b.innerHTML='';}
}

/* ── 내려받기 (현재 조건이 적용된 데이터 기준) ──────────────────────────── */
function exportRows(t){
  const cols=visibleColumns(t);
  const rows=filteredRows(t).map(({r})=>cols.map(c=>String(r.cells[c]??'')));
  return {cols,rows};
}
function stamp(){
  const d=new Date();
  const p=n=>String(n).padStart(2,'0');
  return `${d.getFullYear()}${p(d.getMonth()+1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`;
}
function saveBlob(blob,name){
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url;a.download=name;document.body.appendChild(a);a.click();
  setTimeout(()=>{document.body.removeChild(a);URL.revokeObjectURL(url);},0);
  toast(name+' 내려받기를 시작했습니다.');
}
function downloadMenu(k){
  const t=byKey[k];
  const n=filteredRows(t).length;
  return `<div class="nocand" style="padding:2px 2px 8px">현재 조건 <b>${nf(n,0)}건</b> · 표시 중인 컬럼 기준</div>
    <button class="dlopt" onclick="downloadCsv('${jsq(k)}')">CSV (.csv)<small>엑셀에서 바로 열림 (UTF-8 BOM)</small></button>
    <button class="dlopt" onclick="downloadXlsx('${jsq(k)}')">Excel (.xlsx)<small>서식 있는 실제 엑셀 파일</small></button>
    <button class="dlopt" onclick="downloadMd('${jsq(k)}')">Markdown (.md)<small>메일·위키·보고서 붙여넣기용</small></button>
    <button class="dlopt" onclick="window.print()">PDF / 인쇄<small>브라우저 인쇄창에서 "PDF로 저장"</small></button>`;
}
/* 아래 3개는 "컬럼 이름 배열 + 문자열 행 배열"만 받으므로 표 탭과 탐색 탭이
   똑같이 씁니다 (탐색 탭은 여러 표를 조인한 결과를 그대로 넘깁니다). */
function downloadCsvRows(name,cols,rows){
  const q=s=>'"'+String(s).replace(/"/g,'""')+'"';
  const csv=[cols.map(q).join(',')].concat(rows.map(r=>r.map(q).join(','))).join('\r\n');
  saveBlob(new Blob(['﻿'+csv],{type:'text/csv;charset=utf-8'}),`${name}_${stamp()}.csv`);
}
function downloadMdRows(name,cols,rows,meta){
  const cell=s=>String(s).replace(/\|/g,'\\|');
  const md=[`# ${name}`,'',`- 기준 시각: ${(DATA.generated_at||'').replace('T',' ')}`]
    .concat(meta||[])
    .concat([`- 건수: ${rows.length}건`,'',
      '| '+cols.map(cell).join(' | ')+' |',
      '|'+cols.map(()=>'---').join('|')+'|'])
    .concat(rows.map(r=>'| '+r.map(cell).join(' | ')+' |')).join('\n');
  saveBlob(new Blob([md],{type:'text/markdown;charset=utf-8'}),`${name}_${stamp()}.md`);
}
function downloadXlsxRows(name,cols,rows){
  saveBlob(buildXlsx([{name:name,cols,rows}]),`${name}_${stamp()}.xlsx`);
}
function downloadCsv(k){
  const t=byKey[k],{cols,rows}=exportRows(t);
  downloadCsvRows(`PPA_${t.label}`,cols,rows);
}
function downloadMd(k){
  const t=byKey[k],{cols,rows}=exportRows(t);
  downloadMdRows(`PPA ${t.label}`,cols,rows,[`- 조건: ${filterDescription(k)}`]);
}

/* 최소 구현 XLSX 작성기 — 외부 라이브러리 없이 압축 없는(store) ZIP을 직접
   만들어 실제 .xlsx 파일을 생성합니다. 사내망에 CDN을 못 쓰기 때문에 필요. */
const CRC_TABLE=(()=>{const t=new Uint32Array(256);
  for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=(c&1)?(0xEDB88320^(c>>>1)):(c>>>1);t[n]=c>>>0;}return t;})();
function crc32(buf){let c=0xFFFFFFFF;
  for(let i=0;i<buf.length;i++)c=CRC_TABLE[(c^buf[i])&0xFF]^(c>>>8);
  return (c^0xFFFFFFFF)>>>0;}
function u8(s){return new TextEncoder().encode(s);}
function zipStore(files){
  const now=new Date();
  const dosTime=((now.getHours()&31)<<11)|((now.getMinutes()&63)<<5)|((now.getSeconds()/2)&31);
  const dosDate=(((now.getFullYear()-1980)&127)<<9)|(((now.getMonth()+1)&15)<<5)|(now.getDate()&31);
  const locals=[],centrals=[];let offset=0,total=0;
  files.forEach(f=>{
    const name=u8(f.name),data=f.data,crc=crc32(data);
    const lh=new Uint8Array(30+name.length),lv=new DataView(lh.buffer);
    lv.setUint32(0,0x04034b50,true);lv.setUint16(4,20,true);lv.setUint16(6,0x0800,true);
    lv.setUint16(8,0,true);lv.setUint16(10,dosTime,true);lv.setUint16(12,dosDate,true);
    lv.setUint32(14,crc,true);lv.setUint32(18,data.length,true);lv.setUint32(22,data.length,true);
    lv.setUint16(26,name.length,true);lv.setUint16(28,0,true);lh.set(name,30);
    const cd=new Uint8Array(46+name.length),cv=new DataView(cd.buffer);
    cv.setUint32(0,0x02014b50,true);cv.setUint16(4,20,true);cv.setUint16(6,20,true);
    cv.setUint16(8,0x0800,true);cv.setUint16(10,0,true);cv.setUint16(12,dosTime,true);
    cv.setUint16(14,dosDate,true);cv.setUint32(16,crc,true);cv.setUint32(20,data.length,true);
    cv.setUint32(24,data.length,true);cv.setUint16(28,name.length,true);cv.setUint32(42,offset,true);
    cd.set(name,46);
    locals.push(lh,data);centrals.push(cd);
    offset+=lh.length+data.length;total+=lh.length+data.length;
  });
  const cdSize=centrals.reduce((s,c)=>s+c.length,0);
  const eo=new Uint8Array(22),ev=new DataView(eo.buffer);
  ev.setUint32(0,0x06054b50,true);ev.setUint16(8,files.length,true);ev.setUint16(10,files.length,true);
  ev.setUint32(12,cdSize,true);ev.setUint32(16,offset,true);
  const out=new Uint8Array(total+cdSize+22);let p=0;
  locals.concat(centrals).concat([eo]).forEach(c=>{out.set(c,p);p+=c.length;});
  return out;
}
function xesc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
  .replace(/"/g,'&quot;').replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g,'');}
function colLetter(n){let s='';n++;while(n>0){const m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s;}
function sheetXml(cols,rows){
  const isNum=cols.map(c=>isNumCol(c));
  const head='<row r="1">'+cols.map((c,j)=>
    `<c r="${colLetter(j)}1" s="1" t="inlineStr"><is><t xml:space="preserve">${xesc(c)}</t></is></c>`).join('')+'</row>';
  const body=rows.map((r,i)=>'<row r="'+(i+2)+'">'+r.map((v,j)=>{
    const ref=colLetter(j)+(i+2);
    if(v==='') return '';
    if(isNum[j]&&/^-?\d+(\.\d+)?$/.test(v)) return `<c r="${ref}"><v>${v}</v></c>`;
    return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xesc(v)}</t></is></c>`;
  }).join('')+'</row>').join('');
  const widths=cols.map((c,j)=>`<col min="${j+1}" max="${j+1}" width="${Math.min(42,Math.max(11,c.length*1.9+4))}" customWidth="1"/>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
<cols>${widths}</cols><sheetData>${head}${body}</sheetData></worksheet>`;
}
function safeSheetName(name,used){
  let n=String(name).replace(/[\\\/\?\*\[\]:]/g,'_').slice(0,31)||'Sheet';
  let base=n,i=2;while(used.has(n)){n=(base.slice(0,28)+'_'+i);i++;}
  used.add(n);return n;
}
function buildXlsx(sheets){
  const used=new Set();
  const names=sheets.map(s=>safeSheetName(s.name,used));
  const files=[
    {name:'[Content_Types].xml',data:u8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
${sheets.map((s,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}
</Types>`)},
    {name:'_rels/.rels',data:u8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`)},
    {name:'xl/workbook.xml',data:u8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>${names.map((n,i)=>`<sheet name="${xesc(n)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('')}</sheets></workbook>`)},
    {name:'xl/_rels/workbook.xml.rels',data:u8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${sheets.map((s,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('')}
<Relationship Id="rId${sheets.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`)},
    {name:'xl/styles.xml',data:u8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><sz val="11"/><name val="맑은 고딕"/></font><font><b/><sz val="11"/><name val="맑은 고딕"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFEFEEE7"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
<dxfs count="0"/>
</styleSheet>`)},
  ].concat(sheets.map((s,i)=>({name:`xl/worksheets/sheet${i+1}.xml`,data:u8(sheetXml(s.cols,s.rows))})));
  return new Blob([zipStore(files)],
    {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
function downloadXlsx(k){
  const t=byKey[k],{cols,rows}=exportRows(t);
  saveBlob(buildXlsx([{name:t.label,cols,rows}]),`PPA_${t.label}_${stamp()}.xlsx`);
}
function downloadAllXlsx(){
  const sheets=DATA.tables.map(t=>({name:t.label,cols:t.columns,rows:t.rows.map(r=>t.columns.map(c=>String(r.cells[c]??'')))}));
  saveBlob(buildXlsx(sheets),`PPA_전체_${stamp()}.xlsx`);
}
function downloadSummaryMd(){
  const lines=[buildSummaryText(),'','## 표별 건수','','| 표 | 건수 | 검증오류 | 변경 |','|---|---|---|---|'];
  DATA.tables.forEach(t=>{
    const e=t.rows.filter(r=>(r.error_cols||[]).length>0).length;
    const c=t.rows.filter(r=>!!r.change).length;
    lines.push(`| ${t.label} | ${t.rows.length} | ${e} | ${CHANGES.has_prev?c:'-'} |`);});
  saveBlob(new Blob([lines.join('\n')],{type:'text/markdown;charset=utf-8'}),`PPA_요약_${stamp()}.md`);
}

/* ── 보고용 요약 텍스트 / 클립보드 ──────────────────────────────────────── */
function toast(msg){
  const el=document.getElementById('toast');if(!el) return;
  el.textContent=msg;el.classList.add('show');
  clearTimeout(window._toastTimer);
  window._toastTimer=setTimeout(()=>el.classList.remove('show'),2200);
}

/* ── 표 안에서 바로 수정/추가/삭제 ─────────────────────────────────────────
   실제 입력/저장 로직은 여기가 아니라 dashboard_form.js(실시간 입력 서버가
   붙어 있을 때만 로드됨)가 갖고 있고, window.PPA_FORM 으로 그 진입점만
   노출됩니다. 이 대시보드 화면은 그 창을 원하는 표/레코드로 열어달라고
   요청만 합니다 - 검증·FK 확인·삭제 시 참조 검사 등은 전부 기존 폼이
   그대로 재사용하므로 여기서 따로 구현하지 않습니다. */
function needsLiveServer(){
  toast('실시간 입력 서버 실행 중에만 사용할 수 있습니다 (간편 입력/저장 버튼 참고).');
}
function openInlineAdd(tk){
  if(!window.PPA_FORM){needsLiveServer();return;}
  window.PPA_FORM.add(tk);
}
function openInlineEdit(tk,pkv){
  if(!window.PPA_FORM){needsLiveServer();return;}
  closeDetail();
  window.PPA_FORM.edit(tk,pkv);
}
function openInlineDelete(tk,pkv){
  if(!window.PPA_FORM){needsLiveServer();return;}
  closeDetail();
  window.PPA_FORM.del(tk,pkv);
}
function openInlineResetBaseline(){
  if(!window.PPA_FORM){needsLiveServer();return;}
  window.PPA_FORM.resetBaseline();
}
function fallbackCopy(text){
  const ta=document.createElement('textarea');
  ta.value=text;ta.style.position='fixed';ta.style.left='-9999px';
  document.body.appendChild(ta);ta.focus();ta.select();
  try{document.execCommand('copy');toast('클립보드에 복사했습니다.');}
  catch(e){toast('복사에 실패했습니다 — 직접 선택해서 복사해주세요.');}
  document.body.removeChild(ta);
}
function copyText(text){
  if(navigator.clipboard&&navigator.clipboard.writeText)
    navigator.clipboard.writeText(text).then(()=>toast('클립보드에 복사했습니다.')).catch(()=>fallbackCopy(text));
  else fallbackCopy(text);
}
function buildSummaryText(){
  const P=byKey['T_발전소'],B=byKey['T_구매계약'],S=byKey['T_판매계약'];
  const supplyMW=P?sumCol('T_발전소','설비용량(MW)'):0;
  const purchMW=B?sumCol('T_구매계약','구매계약용량(MW)'):0;
  const saleMW=S?sumCol('T_판매계약','판매계약용량(MW)'):0;
  const bUn=B?countWhere('T_구매계약','수요기업 미확보','TRUE'):0;
  const sUn=S?countWhere('T_판매계약','공급자원 미확보','TRUE'):0;
  const bExp=B?countExpiring('T_구매계약','공급기한_구매',SOON_DAYS):0;
  const sExp=S?countExpiring('T_판매계약','공급기한_판매',SOON_DAYS):0;
  const mix=P?groupSum('T_발전소','발전원','설비용량(MW)'):[];
  const L=[];
  L.push('[PPA 계약관리 현황 요약]');
  L.push('기준: '+(DATA.generated_at||'').replace('T',' '));
  L.push('');
  L.push('- 발전소 '+(P?P.rows.length:0)+'개 · 설비용량 합계 '+nf(supplyMW)+' MW');
  L.push('- 구매계약 '+(B?B.rows.length:0)+'건 · 총 '+nf(purchMW)+' MW');
  L.push('- 판매계약 '+(S?S.rows.length:0)+'건 · 총 '+nf(saleMW)+' MW');
  L.push('- 구매-판매 밸런스: '+((purchMW-saleMW)>=0?'+':'')+nf(purchMW-saleMW)+' MW');
  if(mix.length) L.push('- 발전원 비중: '+mix.map(([g,v])=>g+' '+nf(v)+'MW').join(', '));
  L.push('- 미확보: 구매계약 '+bUn+'건 · 판매계약 '+sUn+'건');
  L.push('- '+SOON_DAYS+'일 내 공급기한 도래: 구매 '+bExp+'건 · 판매 '+sExp+'건');
  L.push('- 검증 오류: '+DATA.validation.total_errors+'건'+(DATA.validation.total_errors?' (검증 탭 확인 필요)':' (전 표 정상)'));
  if(CHANGES.has_prev)
    L.push('- 직전 대비 변경: 추가 '+CHANGES.total_added+' · 수정 '+CHANGES.total_changed+' · 삭제 '+CHANGES.total_removed);
  return L.join('\n');
}
function copySummary(){copyText(buildSummaryText());}

/* ── 인쇄 머리글 (화면에는 숨김, 인쇄할 때만 보임) ──────────────────────── */
function printHead(title,cond){
  return `<div class="printhead"><h2>${esc(title)}</h2>
    <div class="pmeta">PPA 계약관리 · 기준 ${esc((DATA.generated_at||'').replace('T',' '))}${cond?' · 조건: '+esc(cond):''}</div></div>`;
}

/* ── 테마 / URL / 렌더 ──────────────────────────────────────────────────── */
function applyTheme(){
  document.documentElement.setAttribute('data-theme',state.theme);
  const b=document.getElementById('themeIcon');if(b) b.textContent=state.theme==='dark'?'☀️':'🌙';
}
function toggleTheme(){state.theme=state.theme==='dark'?'light':'dark';writeLS('ppa_theme',state.theme);applyTheme();}

/* ── 뒤로가기 ────────────────────────────────────────────────────────────
   "화면이 바뀌었다고 체감하는" 상태만 history entry로 쌓는다. 검색어/
   정렬/페이지 이동/컬럼 숨김 같은 잦은 조작은 일부러 스냅샷에서 뺐다 —
   그 값만 바뀌어서는 navSnapshot() 문자열이 그대로라 push가 안 일어남.
   render() 한 곳(syncHash 호출부)에서만 다루므로 개별 onclick 50여 곳은
   손대지 않는다. */
function navSnapshot(){
  return {
    tab:state.tab,lookupTable:state.lookupTable,lookup:state.lookup,
    filters:state.filters,dateFilters:state.dateFilters,
    onlyErr:state.onlyErr,onlyChg:state.onlyChg,
    modal:state.modal,homeTrend:state.homeTrend,
    explore:state.explore?{base:state.explore.base,tables:state.explore.tables,
      cols:state.explore.cols,missing:state.explore.missing}:null,
  };
}
function applyNavSnapshot(s){
  if(!s) return;
  state.tab=s.tab;state.lookupTable=s.lookupTable;state.lookup=s.lookup;
  state.filters=s.filters||{};state.dateFilters=s.dateFilters||{};
  state.onlyErr=s.onlyErr||{};state.onlyChg=s.onlyChg||{};
  state.modal=s.modal;state.homeTrend=s.homeTrend||state.homeTrend;
  if(s.explore){
    if(!state.explore) initExplore(s.explore.base);
    Object.assign(state.explore,s.explore);
  }
}
let g_lastNavKey=null;
function syncHash(){
  /* 공유 가능한 딥링크는 지금처럼 관계조회+lookup일 때만 해시에 싣는다
     (그 외 굵직한 상태는 history.state에만 실림 — URL 길이/가독성 때문에
     해시로는 안 뺌). 새로고침해도 이 해시만으로 lookup이 복원된다. */
  const h=(state.tab==='관계조회'&&state.lookup)
    ?'#lookup='+encodeURIComponent(state.lookup.table)+':'+encodeURIComponent(state.lookup.pk):'';
  if((location.hash||'')!==h){
    if(h) history.replaceState(history.state,'',h);
    else history.replaceState(history.state,'',location.pathname+location.search);
  }
  const snap=navSnapshot(),key=JSON.stringify(snap);
  if(key!==g_lastNavKey){
    history.pushState({nav:snap},'',location.href);
    g_lastNavKey=key;
  }
}
function parseHash(){
  const m=location.hash.match(/^#lookup=([^:]+):(.+)$/);
  if(m){
    const table=decodeURIComponent(m[1]),pk=decodeURIComponent(m[2]);
    if(byKey[table]){state.tab='관계조회';state.lookupTable=table;setLookup(table,pk);}
  }
}
/* 간편 입력/저장으로 값을 바꾼 뒤 location.reload() 하면 원래 "지금 보던
   화면 그대로" 돌아와야 하는데, 아무 손도 안 대면 state.tab 기본값인 '홈'
   으로 매번 튕깁니다. dashboard_form.js가 reload 직전에 sessionStorage에
   남겨둔 탭을 여기서 1회성으로 복원합니다(#lookup= 해시가 있으면 그 딥링크가
   더 명시적인 의도이므로, 이 함수 뒤에 이어서 실행되는 parseHash()가
   우선권을 가져가도록 순서를 둡니다 — 아래 초기화 IIFE 참고). */
function restoreTabFromSession(){
  var saved;
  try{ saved=sessionStorage.getItem('ppa_return_tab'); sessionStorage.removeItem('ppa_return_tab'); }
  catch(e){ return; }
  if(!saved) return;
  if(saved==='홈'||saved==='관계조회'||saved==='탐색'||saved==='비교'||saved==='변경'||saved==='검증'||byKey[saved]){
    state.tab=saved;
  }
}
window.addEventListener('popstate',e=>{
  if(!e.state||!e.state.nav) return; /* 우리가 만들지 않은 진입점 — 무시하고 추가 push도 안 만듦 */
  g_lastNavKey=JSON.stringify(e.state.nav);
  applyNavSnapshot(e.state.nav);
  render();
});
/* 사이드바 내비 아이콘 — 외부 아이콘 폰트 없이 유니코드 이모지만 사용(오프라인 제약) */
const NAV_ICONS={'홈':'🏠','관계조회':'🔗','탐색':'🧭','비교':'⚖️',
  'T_발전소':'🏭','T_구매계약':'🧾','T_수요기업':'🏢','T_판매계약':'📑','T_전기사용지':'📍','T_수급매칭':'🔀',
  '변경':'🕓','검증':'✅'};
/* 상단바 제목이 참조하는 단일 기준 — 사이드바에 지금 그 항목이 실제로
   그려졌는지와 무관하게(예: 비교 탭에서 마지막 고정을 해제해도) 항상
   정확한 라벨을 돌려준다 */
function tabLabel(k){
  if(k==='비교') return '비교';
  if(byKey[k]) return byKey[k].label;
  return k;
}
function navItem(k,label,dot){
  return `<button class="navitem${k===state.tab?' on':''}" data-k="${esc(k)}"
    onclick="state.tab='${jsq(k)}';closeSidebar();render()">
    <span class="navicon">${NAV_ICONS[k]||'📄'}</span><span class="navlabel">${esc(label)}</span>${dot}</button>`;
}
function navGroup(label,items){
  if(!items.length) return '';
  return `<div class="navgroup"><div class="navgrouplabel">${esc(label)}</div>${items.map(([k,l,d])=>navItem(k,l,d)).join('')}</div>`;
}
function renderTabs(){
  const overview=[['홈','홈',''],['관계조회','관계조회',''],['탐색','탐색','']];
  if(state.pinned.length) overview.push(['비교','비교 ('+state.pinned.length+')','']);
  const tables=DATA.tables.map(t=>{
    const e=t.rows.filter(r=>(r.error_cols||[]).length>0).length;
    return [t.key,t.label,e?`<span class="tabdot">${e}</span>`:''];
  });
  const chgTot=CHANGES.has_prev?(CHANGES.total_added+CHANGES.total_changed+CHANGES.total_removed):0;
  const manage=[
    ['변경','변경',chgTot?`<span class="tabdot info">${chgTot}</span>`:''],
    ['검증','검증',DATA.validation.total_errors?`<span class="tabdot">${DATA.validation.total_errors}</span>`:''],
  ];
  document.getElementById('tabbar').innerHTML=
    navGroup('개요',overview)+navGroup('데이터',tables)+navGroup('관리',manage);
}
function openSidebar(){document.getElementById('sidebar').classList.add('open');document.getElementById('sidebarBackdrop').classList.add('open');}
function closeSidebar(){document.getElementById('sidebar').classList.remove('open');document.getElementById('sidebarBackdrop').classList.remove('open');}
function toggleSidebar(){document.getElementById('sidebar').classList.contains('open')?closeSidebar():openSidebar();}
/* 성능 메모(자율 점검): 이 대시보드는 매 렌더마다 문자열로 조립한 HTML을
   #view.innerHTML에 통째로 한 번만 대입합니다 — 브라우저가 innerHTML을
   파싱할 때 자체적으로 배치(batch) 최적화를 하므로, createElement로 노드를
   하나씩 만들어 DocumentFragment에 담아 붙이는 방식보다 실제로 더 느리지
   않고 코드도 훨씬 단순합니다(리플로우도 어차피 둘 다 "교체 1회"로 동일).
   그래서 이 파일 전체에서 노드별 DOM API 대신 문자열 템플릿 방식을 그대로
   유지했습니다. 대신 실측으로 느렸던 지점(타이핑 시 매 키 입력마다 전체
   재렌더)은 onSearchType()의 디바운스(140ms)로 렌더 횟수 자체를 줄여
   해결했습니다 — 이게 대량 데이터에서 체감되는 "렉"의 실제 원인이었습니다. */
function render(){
  renderTabs();
  const view=document.getElementById('view');
  /* innerHTML을 새로 그리면 입력창이 매번 재생성되어 포커스가 풀립니다 —
     그대로 두면 한 글자 칠 때마다 커서가 빠져 "검색이 안 되는" 것처럼 보입니다.
     그래서 포커스/커서 위치를 기억했다가 같은 id의 요소에 복원합니다. */
  const act=document.activeElement;
  let fid=null,ss=null,se=null;
  if(act&&act.id&&act!==document.body){
    fid=act.id;
    if(typeof act.selectionStart==='number'){ss=act.selectionStart;se=act.selectionEnd;}
  }
  let html;
  if(state.tab==='홈') html=tHome();
  else if(state.tab==='관계조회') html=tLookup();
  else if(state.tab==='탐색') html=tExplore();
  else if(state.tab==='비교') html=tCompare();
  else if(state.tab==='변경') html=tChanges();
  else if(state.tab==='검증') html=tVerify();
  else html=tData(byKey[state.tab]);
  view.innerHTML=html;
  document.getElementById('pageTitle').textContent=tabLabel(state.tab);
  document.getElementById('modalHost').innerHTML=modalHtml();
  if(fid){
    const el=document.getElementById(fid);
    if(el){el.focus();if(ss!==null&&el.setSelectionRange){try{el.setSelectionRange(ss,se);}catch(e){}}}
  }
  const st=document.getElementById('status');
  const ok=DATA.validation.total_errors===0;
  st.className='pill '+(ok?'ok':'no');
  st.textContent=ok?'검증 통과':`검증 오류 ${DATA.validation.total_errors}건`;
  const cp=document.getElementById('chgpill');
  const chgTot=CHANGES.has_prev?(CHANGES.total_added+CHANGES.total_changed+CHANGES.total_removed):0;
  cp.style.display=chgTot?'':'none';
  cp.textContent=`변경 ${chgTot}건`;
  syncHash();
}
document.addEventListener('click',e=>{
  const w=document.querySelector('.gsearchwrap');
  if(w&&!w.contains(e.target)) closeGlobalSearch();
});
document.addEventListener('keydown',e=>{
  if(e.key==='Escape'){
    if(state.modal){closeDetail();return;}
    if(document.getElementById('sidebar').classList.contains('open')){closeSidebar();return;}
    closeGlobalSearch();
  }
});
(function(){
  loadHidden();applyTheme();restoreTabFromSession();parseHash();
  document.getElementById('foot-src').textContent=(DATA.is_demo?'데모 데이터':'실 데이터')+
    ' · '+DATA.tables.map(t=>t.label+' '+t.rows.length).join(' · ');
  /* 시작 entry를 자기 자신으로 확정 — 첫 render()가 이 값과 똑같은
     스냅샷을 만들므로 불필요한 pushState 없이 replaceState 한 번뿐 */
  g_lastNavKey=JSON.stringify(navSnapshot());
  history.replaceState({nav:JSON.parse(g_lastNavKey)},'',location.href);
  render();
})();
"""

# ─────────────────────────────────────────────────────────────────────────────
# 뼈대
# ─────────────────────────────────────────────────────────────────────────────
_HTML = r"""<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>PPA 계약관리 현황</title>
<style>/*__CSS__*/</style></head><body>
<div class="shell">
  <div class="sidebarbackdrop" id="sidebarBackdrop" onclick="closeSidebar()"></div>
  <aside class="sidebar" id="sidebar">
    <div class="brand"><span class="brandmark">P</span>
      <span class="brandtext"><b>PPA 계약관리</b><span>정적 스냅샷 · 조회 전용</span></span></div>
    <nav class="navwrap" id="tabbar"></nav>
    <div class="sidefoot">
      <button id="themeBtn" class="sidethemebtn" onclick="toggleTheme()" title="화면 테마 전환"><span id="themeIcon">🌙</span><span>화면 테마 전환</span></button>
    </div>
  </aside>
  <div class="main">
    <header class="topbar">
      <div class="topbar-left">
        <button class="menubtn" onclick="toggleSidebar()" title="메뉴" aria-label="메뉴 열기">☰</button>
        <h1 id="pageTitle">홈</h1>
      </div>
      <div class="topbar-right">
        <div class="gsearchwrap">
          <input id="globalSearch" class="search" style="width:100%" placeholder="전체 표에서 검색 (ID, 발전소명, 기업명, 담당자 등)…"
            oninput="onSearchType(event,runGlobalSearch)" onfocus="onGlobalSearch(this.value)">
          <div id="globalResults" class="globalresults"></div>
        </div>
        <button id="chgpill" class="pill chg" onclick="state.tab='변경';render()" style="display:none">변경</button>
        <button id="status" class="pill" onclick="state.tab='검증';render()">—</button>
      </div>
    </header>{{DEMO}}
    <div class="wrap"><div id="view"></div>
      <footer><span>생성 {{NOW}}</span><span id="foot-src"></span></footer></div>
  </div>
</div>
<div id="modalHost"></div><div id="toast"></div>
<script>/*__JS__*/</script></body></html>"""

TEMPLATE = _HTML.replace("/*__CSS__*/", _CSS).replace("/*__JS__*/", _JS)


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
