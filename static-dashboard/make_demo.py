#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""make_demo.py — 화면 확인용 데모 데이터로 PPA현황_데모.html 생성.

일부러 오류 케이스(PK 중복, FK 참조 실패, 조합중복)를 섞어 넣어서
검증 탭과 셀 하이라이트가 실제로 어떻게 보이는지 확인할 수 있게 했습니다.
[변경] 탭도 보여주려고 "지난번 생성분"에 해당하는 가상의 이전 스냅샷을
만들어서 추가/수정/삭제가 각각 잡히도록 해뒀습니다.
"""
import copy
import datetime
import os
import sys

# Windows embeddable Python(python.org "embeddable package")은 pythonXXX._pth
# 파일이 있으면 실행하는 스크립트의 폴더를 자동으로 sys.path에 넣어주지 않습니다
# (일반 Python 설치본과 다른 부분). 그래서 같은 폴더의 build_dashboard.py 등을
# import하려면 이 폴더를 직접 sys.path에 추가해줘야 합니다.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from build_dashboard import build_payload
from ppa_changes import compute_changes
from ppa_dashboard_render import render_dashboard

tables_data = {
    "T_발전소": [
        {"발전소ID": "P001", "발전소명": "신안태양광1호", "발전법인명": "그린에너지(주)",
         "설비용량(MW)": "3.0", "발전원": "태양광", "Readiness": "New", "MGA_Supply": "0.133333333"},
        {"발전소ID": "P002", "발전소명": "영광풍력1호", "발전법인명": "클린파워(주)",
         "설비용량(MW)": "5.0", "발전원": "풍력", "Readiness": "New", "MGA_Supply": "0.133333333"},
        {"발전소ID": "P003", "발전소명": "완도해상풍력", "발전법인명": "오션윈드(주)",
         "설비용량(MW)": "8.0", "발전원": "풍력", "Readiness": "Ready", "MGA_Supply": "0.12"},
    ],
    "T_구매계약": [
        {"구매계약ID": "구매-P001-2024-01", "발전소ID": "P001", "구매계약용량(MW)": "3.0",
         "구매단가(원/kWh)": "142", "공급기한_구매": "2030-12-31", "계약기간(년)": "10",
         "수요기업 미확보": "FALSE", "구매 담당자": "최기철"},
        {"구매계약ID": "구매-BAD-2024-01", "발전소ID": "P999", "구매계약용량(MW)": "1.0",
         "구매단가(원/kWh)": "150", "공급기한_구매": "2026-09-30", "계약기간(년)": "5",
         "수요기업 미확보": "TRUE", "구매 담당자": "박창옥"},  # 발전소ID 참조 오류 (P999 없음)
        {"구매계약ID": "구매-P003-2026-01", "발전소ID": "P003", "구매계약용량(MW)": "8.0",
         "구매단가(원/kWh)": "138", "공급기한_구매": "2032-06-30", "계약기간(년)": "15",
         "수요기업 미확보": "TRUE", "구매 담당자": "최기철"},
    ],
    "T_수요기업": [
        {"수요기업ID": "D001", "기업명": "㈜그린전자"},
        {"수요기업ID": "D002", "기업명": "한빛반도체㈜"},
    ],
    "T_판매계약": [
        {"판매계약ID": "판매-D001-2024-01", "수요기업ID": "D001", "판매계약용량(MW)": "1.0",
         "계약일": "2024-01-01", "공급기한_판매": "2027-01-01", "계약유형": "D",
         "판매단가(원/kWh)": "160", "공급자원 미확보": "FALSE", "판매 담당자": "홍원정",
         "계약기간(년)": "10", "Requirement": "New", "MGA_Demand": "0.1"},
        {"판매계약ID": "판매-D001-2024-01", "수요기업ID": "D001", "판매계약용량(MW)": "1.2",  # PK 중복
         "계약일": "2024-06-01", "공급기한_판매": "2027-06-01", "계약유형": "V",
         "판매단가(원/kWh)": "158", "공급자원 미확보": "FALSE", "판매 담당자": "홍원정",
         "계약기간(년)": "10", "Requirement": "New", "MGA_Demand": "0.1"},
        {"판매계약ID": "판매-D002-2026-01", "수요기업ID": "D002", "판매계약용량(MW)": "4.0",
         "계약일": "2026-02-01", "공급기한_판매": "2026-11-30", "계약유형": "D",
         "판매단가(원/kWh)": "165", "공급자원 미확보": "TRUE", "판매 담당자": "김서연",
         "계약기간(년)": "8", "Requirement": "New", "MGA_Demand": "0.11"},
    ],
    "T_전기사용지": [
        {"전기사용지ID": "전기사용지-D001-2024-01", "판매계약ID": "판매-D001-2024-01",
         "전기사용지명": "그린전자 본사", "전기사용지계약용량(MW)": "1.0"},
        {"전기사용지ID": "전기사용지-D002-2026-01", "판매계약ID": "판매-D002-2026-01",
         "전기사용지명": "한빛반도체 청주캠퍼스", "전기사용지계약용량(MW)": "4.0"},
    ],
    "T_수급매칭": [
        {"수급매칭ID": "매칭-001", "전기사용지ID": "전기사용지-D001-2024-01",
         "구매계약ID": "구매-P001-2024-01", "현황": "1. 공급 중"},
        {"수급매칭ID": "매칭-002", "전기사용지ID": "전기사용지-D001-2024-01",
         "구매계약ID": "구매-P001-2024-01", "현황": "1. 공급 중"},  # 조합중복 (매칭-001과 동일 조합)
        {"수급매칭ID": "매칭-003", "전기사용지ID": "전기사용지-D002-2026-01",
         "구매계약ID": "구매-P003-2026-01", "현황": "2. 공급 예정"},
    ],
}


def fake_previous(cur: dict) -> dict:
    """[변경] 탭 시연용 — '지난번 생성분'을 흉내낸 이전 스냅샷.

    현재 데이터에서 한 건을 빼고(→ 이번에 추가된 것으로 잡힘), 한 건의 값을
    되돌리고(→ 수정으로 잡힘), 지금은 없는 한 건을 넣어(→ 삭제로 잡힘) 만듭니다.
    """
    prev = copy.deepcopy(cur)
    # 1) 완도해상풍력(P003)과 그 구매계약은 지난번엔 없었다 → 이번에 "추가"
    prev["T_발전소"] = [r for r in prev["T_발전소"] if r["발전소ID"] != "P003"]
    prev["T_구매계약"] = [r for r in prev["T_구매계약"] if r["구매계약ID"] != "구매-P003-2026-01"]
    # 2) 구매단가/담당자가 지난번과 다르다 → "수정"
    for r in prev["T_구매계약"]:
        if r["구매계약ID"] == "구매-P001-2024-01":
            r["구매단가(원/kWh)"] = "138"
            r["구매 담당자"] = "박창옥"
    # 3) 지난번엔 있었는데 지금은 없는 수요기업 → "삭제"
    prev["T_수요기업"] = prev["T_수요기업"] + [{"수요기업ID": "D009", "기업명": "종료된거래처㈜"}]
    return {
        "generated_at": (datetime.datetime.now() - datetime.timedelta(days=7)).isoformat(timespec="seconds"),
        "tables": prev,
    }


if __name__ == "__main__":
    changes, marks = compute_changes(tables_data, fake_previous(tables_data))
    payload = build_payload(tables_data, is_demo=True, changes=changes, marks=marks)
    with open("PPA현황_데모.html", "w", encoding="utf-8") as f:
        f.write(render_dashboard(payload))
    print("생성: PPA현황_데모.html")
    print(f"검증 오류: {payload['validation']['total_errors']}건 "
          f"(의도적으로 넣은 PK중복/발전소ID참조/조합중복 케이스)")
    print(f"변경 시연: 추가 {changes['total_added']} · 수정 {changes['total_changed']} · "
          f"삭제 {changes['total_removed']}건")
