#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ppa_ids.py — 구매계약/판매계약/수급매칭/전기사용지 4개 표의 PK를
관련 필드로부터 자동으로 계산하는 순수 함수 모음(COM/HTTP에 의존하지
않아 그대로 단위 테스트할 수 있습니다).

규칙(사용자 확정):
  구매계약ID   = 구매-{발전소ID}-{공급기한_구매 연도}-{순번}
  판매계약ID   = 판매-{수요기업ID}-{공급기한_판매 연도}-{순번}
  수급매칭ID   = 매칭-{수요기업ID}-{공급기한_판매 연도}-{순번}
                (수요기업ID·연도는 선택된 전기사용지ID → 그 행의 판매계약ID
                 → 그 판매계약의 수요기업ID/공급기한_판매로 찾음)
  전기사용지ID = 전기사용지-{수요기업ID}-{공급기한_판매 연도}-{순번}-{전기사용지명}
                (수요기업ID·연도는 선택된 판매계약ID로 바로 찾음)

연도 자리에 쓸 날짜가 비어 있으면(계약을 아직 안 맺은 "임시" 건) "T"를
씁니다. 순번은 같은 접두어(발전소/수요기업 + 연도)를 가진 기존 ID들 중
가장 큰 번호에 1을 더해 2자리 0채움으로 만듭니다.

편집 시 근거 필드가 안 바뀌었으면(=새로 계산한 접두어가 기존 PK와
같으면) 번호를 그대로 유지합니다 - 저장할 때마다 번호가 계속 올라가는
것을 막기 위함입니다. 전기사용지명처럼 접두어에 포함되지 않는 값만
바뀐 경우에도 번호는 유지한 채 그 부분만 새 값으로 바꿉니다.
"""
from __future__ import annotations

import re
from typing import Optional


def _s(v) -> str:
    return str(v or "").strip()


def _year_or_t(date_str) -> str:
    """"YYYY-MM-DD" 형식에서 앞 4자리만 뽑습니다. 비어 있거나 형식이 아니면
    "T"(임시)로 취급합니다."""
    s = _s(date_str)
    if len(s) >= 4 and s[:4].isdigit():
        return s[:4]
    return "T"


def find_row(tables_data: dict, table_key: str, pk_col: str, pk_value: str) -> Optional[dict]:
    target = _s(pk_value)
    if not target:
        return None
    for r in tables_data.get(table_key, []):
        if _s(r.get(pk_col)) == target:
            return r
    return None


def _resolve_via_sale_contract(sale_pk: str, tables_data: dict) -> Optional[dict]:
    """판매계약ID 하나로 수요기업ID+연도를 찾음 - 전기사용지·수급매칭이 공용으로 씀."""
    sale_row = find_row(tables_data, "T_판매계약", "판매계약ID", sale_pk)
    if not sale_row:
        return None
    demand_id = _s(sale_row.get("수요기업ID"))
    if not demand_id:
        return None
    return {"id_part": demand_id, "year": _year_or_t(sale_row.get("공급기한_판매"))}


def _resolve_purchase(record: dict, tables_data: dict) -> Optional[dict]:
    plant_id = _s(record.get("발전소ID"))
    if not plant_id:
        return None
    return {"id_part": plant_id, "year": _year_or_t(record.get("공급기한_구매"))}


def _resolve_sale(record: dict, tables_data: dict) -> Optional[dict]:
    demand_id = _s(record.get("수요기업ID"))
    if not demand_id:
        return None
    return {"id_part": demand_id, "year": _year_or_t(record.get("공급기한_판매"))}


def _resolve_electric_site(record: dict, tables_data: dict) -> Optional[dict]:
    sale_pk = _s(record.get("판매계약ID"))
    if not sale_pk:
        return None
    base = _resolve_via_sale_contract(sale_pk, tables_data)
    if not base:
        return None
    base["suffix"] = _s(record.get("전기사용지명"))
    return base


def _resolve_matching(record: dict, tables_data: dict) -> Optional[dict]:
    site_pk = _s(record.get("전기사용지ID"))
    if not site_pk:
        return None
    site_row = find_row(tables_data, "T_전기사용지", "전기사용지ID", site_pk)
    if not site_row:
        return None
    sale_pk = _s(site_row.get("판매계약ID"))
    if not sale_pk:
        return None
    return _resolve_via_sale_contract(sale_pk, tables_data)


# table_key -> (표시 접두어, PK 컬럼명, 근거 필드 계산 함수, 이 표의 ID를
# 참조하는 자식 표 중 "부모 값이 바뀌면 자기 ID도 다시 계산해야 하는" 목록)
RULES: dict[str, dict] = {
    "T_구매계약": {
        "prefix": "구매", "pk": "구매계약ID", "resolver": _resolve_purchase,
        "id_dependents": [],  # 구매계약 변경은 수급매칭의 FK만 갱신, ID 재계산은 없음
    },
    "T_판매계약": {
        "prefix": "판매", "pk": "판매계약ID", "resolver": _resolve_sale,
        "id_dependents": ["T_전기사용지"],  # 전기사용지ID가 판매계약 데이터에서 파생됨
    },
    "T_전기사용지": {
        "prefix": "전기사용지", "pk": "전기사용지ID", "resolver": _resolve_electric_site,
        "id_dependents": ["T_수급매칭"],  # 수급매칭ID가 (전기사용지 경유) 판매계약 데이터에서 파생됨
    },
    "T_수급매칭": {
        "prefix": "매칭", "pk": "수급매칭ID", "resolver": _resolve_matching,
        "id_dependents": [],  # 수급매칭ID는 아무도 참조하지 않음(연쇄 종점)
    },
}

ID_TABLE_KEYS = tuple(RULES.keys())


def _next_seq(tables_data: dict, table_key: str, pk_col: str, prefix_str: str, exclude_pk: Optional[str]) -> int:
    pattern = re.compile("^" + re.escape(prefix_str) + r"(\d+)")
    max_n = 0
    exclude = _s(exclude_pk)
    for r in tables_data.get(table_key, []):
        pk = _s(r.get(pk_col))
        if exclude and pk == exclude:
            continue
        m = pattern.match(pk)
        if m:
            max_n = max(max_n, int(m.group(1)))
    return max_n + 1


def compute_id(table_key: str, record: dict, tables_data: dict, current_pk: Optional[str] = None) -> dict:
    """반환:
      성공 - {"ok": True, "id": "구매-P001-2025-01", "changed": bool}
      실패(연결된 값이 아직 없음) - {"ok": False, "reason": "..."}
    "changed"는 current_pk(편집 중인 기존 PK, 신규면 None)와 비교한 결과이며,
    호출부는 changed=True일 때만 실제로 PK를 바꾸고 연쇄 갱신을 수행하면 됩니다.
    """
    rule = RULES.get(table_key)
    if not rule:
        return {"ok": False, "reason": f"ID 자동생성 대상이 아닌 표입니다: {table_key}"}

    resolved = rule["resolver"](record, tables_data)
    if not resolved:
        return {"ok": False, "reason": "ID를 계산하려면 연결된 항목(발전소/수요기업/판매계약/전기사용지)을 먼저 선택해주세요."}

    prefix_str = f"{rule['prefix']}-{resolved['id_part']}-{resolved['year']}-"
    suffix = resolved.get("suffix") or ""
    old_pk = _s(current_pk)

    if old_pk and old_pk.startswith(prefix_str):
        # 근거 필드(발전소/수요기업/연도)는 안 바뀌었으니 번호는 유지 - 이름처럼
        # 접두어 밖의 값만 바뀌었으면 그 부분만 새 값으로 다시 붙입니다.
        rest = old_pk[len(prefix_str):]
        seq_part = rest.split("-", 1)[0]
        if seq_part.isdigit():
            candidate = prefix_str + seq_part + (f"-{suffix}" if suffix else "")
            return {"ok": True, "id": candidate, "changed": candidate != old_pk}

    seq = _next_seq(tables_data, table_key, rule["pk"], prefix_str, exclude_pk=old_pk or None)
    new_id = f"{prefix_str}{seq:02d}"
    if suffix:
        new_id = f"{new_id}-{suffix}"
    return {"ok": True, "id": new_id, "changed": new_id != old_pk}


def id_dependents(table_key: str) -> list[str]:
    """이 표의 PK가 바뀌면 "FK 갱신만"이 아니라 "자기 ID도 다시 계산"까지
    해야 하는 자식 표 목록(수급매칭·전기사용지 방향으로만 존재)."""
    rule = RULES.get(table_key)
    return list(rule["id_dependents"]) if rule else []
