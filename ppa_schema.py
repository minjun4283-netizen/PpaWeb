#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ppa_schema.py — PPA 6개 표의 스키마 정의 + 검증 엔진 (순수 로직).

실제 xlsm 화면 캡처로 확인한 컬럼을 그대로 미러링합니다
(PpaWeb/backend/src/schema/tableDefs.ts 와 동일한 정의). 검증 열
(PK중복/PK공란/*참조/조합중복)은 원본 매크로처럼 시트에 저장하지 않고,
이 모듈이 그때그때 계산합니다 — 웹앱 백엔드의 validation.ts 포팅.
"""
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class TableSchema:
    key: str
    label: str
    pk: str
    columns: list[str]
    fk: dict[str, str] = field(default_factory=dict)  # col -> ref table key
    unique_groups: list[list[str]] = field(default_factory=list)


# 발전소→구매계약→수요기업→판매계약→전기사용지→수급매칭 순서로 정리
# (두 체인을 각각 쌓은 뒤 수급매칭으로 마지막에 결합하는 자연스러운 입력 흐름).
TABLES: list[TableSchema] = [
    TableSchema(
        key="T_발전소", label="발전소", pk="발전소ID",
        columns=["발전소ID", "발전소명", "발전법인명", "설비용량(MW)", "발전원", "Readiness", "MGA_Supply"],
    ),
    TableSchema(
        key="T_구매계약", label="구매계약", pk="구매계약ID",
        columns=["구매계약ID", "발전소ID", "구매계약용량(MW)", "구매단가(원/kWh)", "공급기한_구매",
                 "계약기간(년)", "수요기업 미확보", "구매 담당자"],
        fk={"발전소ID": "T_발전소"},
    ),
    TableSchema(
        key="T_수요기업", label="수요기업", pk="수요기업ID",
        columns=["수요기업ID", "기업명"],
    ),
    TableSchema(
        key="T_판매계약", label="판매계약", pk="판매계약ID",
        columns=["판매계약ID", "수요기업ID", "판매계약용량(MW)", "계약일", "공급기한_판매", "계약유형",
                 "판매단가(원/kWh)", "공급자원 미확보", "판매 담당자", "계약기간(년)", "Requirement", "MGA_Demand"],
        fk={"수요기업ID": "T_수요기업"},
    ),
    TableSchema(
        key="T_전기사용지", label="전기사용지", pk="전기사용지ID",
        columns=["전기사용지ID", "판매계약ID", "전기사용지명", "전기사용지계약용량(MW)"],
        fk={"판매계약ID": "T_판매계약"},
    ),
    TableSchema(
        key="T_수급매칭", label="수급매칭", pk="수급매칭ID",
        columns=["수급매칭ID", "전기사용지ID", "구매계약ID", "현황"],
        fk={"전기사용지ID": "T_전기사용지", "구매계약ID": "T_구매계약"},
        unique_groups=[["전기사용지ID", "구매계약ID"]],
    ),
]

TABLE_BY_KEY: dict[str, TableSchema] = {t.key: t for t in TABLES}


def _s(v: Optional[object]) -> str:
    if v is None:
        return ""
    return str(v).strip()


def validate(tables_data: dict[str, list[dict]]) -> dict:
    """검증실행 포팅: PK공란/PK중복, FK공란/참조, 조합중복.

    반환값: {
      "errors": [{"table":, "row_index":, "pk_value":, "error_item":}],
      "error_cols": {(table_key, row_index): {col, ...}},   # 셀 단위 하이라이트용
      "summary_by_table": {table_key: count},
      "summary_by_item": {error_item: count},
    }
    """
    errors: list[dict] = []
    error_cols: dict[tuple[str, int], set[str]] = {}

    def add_error(table_key: str, row_index: int, pk_value: str, error_item: str, cols: list[str]):
        errors.append(
            {"table": table_key, "row_index": row_index, "pk_value": pk_value, "error_item": error_item}
        )
        error_cols.setdefault((table_key, row_index), set()).update(cols)

    for t in TABLES:
        rows = tables_data.get(t.key, [])

        # PK 공란 / PK 중복
        counts: dict[str, int] = {}
        for r in rows:
            v = _s(r.get(t.pk))
            if v:
                counts[v] = counts.get(v, 0) + 1
        for i, r in enumerate(rows):
            v = _s(r.get(t.pk))
            if v == "":
                add_error(t.key, i, v, "PK 공란", [t.pk])
            elif counts.get(v, 0) > 1:
                add_error(t.key, i, v, "PK 중복", [t.pk])

        # FK 공란 / FK 참조
        for fk_col, ref_key in t.fk.items():
            ref_rows = tables_data.get(ref_key, [])
            ref_pk = TABLE_BY_KEY[ref_key].pk
            ref_values = {_s(rr.get(ref_pk)) for rr in ref_rows if _s(rr.get(ref_pk))}
            for i, r in enumerate(rows):
                v = _s(r.get(fk_col))
                pkv = _s(r.get(t.pk))
                if v == "":
                    add_error(t.key, i, pkv, f"{fk_col} 공란", [fk_col])
                elif v not in ref_values:
                    add_error(t.key, i, pkv, f"{fk_col} 참조", [fk_col])

        # 조합중복
        for group in t.unique_groups:
            combo_counts: dict[str, int] = {}
            for r in rows:
                parts = [_s(r.get(c)) for c in group]
                if any(parts):
                    key_str = "|".join(parts)
                    combo_counts[key_str] = combo_counts.get(key_str, 0) + 1
            for i, r in enumerate(rows):
                parts = [_s(r.get(c)) for c in group]
                if any(parts):
                    key_str = "|".join(parts)
                    if combo_counts.get(key_str, 0) > 1:
                        add_error(t.key, i, _s(r.get(t.pk)), "조합중복", group)

    by_table: dict[str, int] = {}
    by_item: dict[str, int] = {}
    for e in errors:
        by_table[e["table"]] = by_table.get(e["table"], 0) + 1
        by_item[e["error_item"]] = by_item.get(e["error_item"], 0) + 1

    return {
        "errors": errors,
        "error_cols": error_cols,
        "summary_by_table": by_table,
        "summary_by_item": by_item,
    }
