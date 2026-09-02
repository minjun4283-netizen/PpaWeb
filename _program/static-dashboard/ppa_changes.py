#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ppa_changes.py — 직전 생성분과 비교해 "무엇이 바뀌었는지"를 계산합니다.

정적 대시보드는 서버도 DB도 없어서 원본 웹앱처럼 실시간 변경이력을 남길 수
없습니다. 대신 대시보드를 만들 때마다 그때의 데이터를 스냅샷(JSON)으로 옆에
저장해두고, 다음번에 만들 때 그 스냅샷과 비교하는 방식으로 같은 목적을
달성합니다 — "지난번 보고 이후 무엇이 달라졌는지"가 화면에 바로 뜹니다.

PK 기준으로 추가/삭제를 판정하고, 같은 PK끼리는 컬럼별로 값을 비교해
바뀐 항목(이전값 → 새값)을 기록합니다.
"""
import json
import os
from typing import Optional

from ppa_schema import TABLE_BY_KEY, TABLES

# 변경 상세는 화면 표시용이라, 대량 편집이 있었을 때 HTML이 과도하게
# 커지지 않도록 상한을 둡니다(집계 건수는 상한과 무관하게 정확합니다).
MAX_DETAILS = 1000

# 이번 한 번의 생성에서 바뀐 것만 보여주는 위 MAX_DETAILS와 달리, 이건 여러 번의
# 생성에 걸쳐 계속 쌓이는 "전체 변경 이력"의 상한입니다 - 최근 1,000건만 남기고
# 오래된 것부터 버립니다(집계가 아니라 실제 항목 개수 기준).
CHANGELOG_MAX = 1000


def _s(v: object) -> str:
    if v is None:
        return ""
    return str(v).strip()


def load_snapshot(path: str) -> Optional[dict]:
    if not path or not os.path.exists(path):
        return None
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError) as exc:
        print(f"[알림] 이전 스냅샷을 읽지 못해 변경 비교를 건너뜁니다: {exc}")
        return None
    if not isinstance(data, dict) or "tables" not in data:
        print("[알림] 이전 스냅샷 형식이 올바르지 않아 변경 비교를 건너뜁니다.")
        return None
    return data


def save_snapshot(path: str, tables_data: dict[str, list[dict]], generated_at: str) -> None:
    payload = {"generated_at": generated_at, "tables": tables_data}
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False)
    except OSError as exc:
        print(f"[알림] 스냅샷을 저장하지 못했습니다(다음 실행 시 변경 비교 불가): {exc}")


def _index_by_pk(rows: list[dict], pk: str) -> dict[str, dict]:
    """PK → 행. PK가 공란이거나 중복이면 첫 번째 것만 씁니다.

    (중복/공란 자체는 검증 탭에서 별도로 오류로 잡히므로, 여기서는 비교가
    깨지지 않게만 처리합니다.)
    """
    out: dict[str, dict] = {}
    for r in rows:
        key = _s(r.get(pk))
        if key and key not in out:
            out[key] = r
    return out


def compute_changes(tables_data: dict[str, list[dict]], prev: Optional[dict]) -> tuple[dict, dict]:
    """반환: (changes 요약, marks)

    marks 는 {(table_key, row_index): {"change": "added"|"changed",
    "changed_cols": [...], "prev": {col: 이전값}}} 형태로, build_payload 가
    각 행에 붙여 화면에서 색으로 표시하는 데 씁니다.
    """
    changes: dict = {
        "has_prev": prev is not None,
        "prev_generated_at": (prev or {}).get("generated_at"),
        "summary": {},
        "total_added": 0,
        "total_removed": 0,
        "total_changed": 0,
        "removed_rows": {},
        "details": [],
        "truncated": False,
    }
    marks: dict = {}

    if prev is None:
        return changes, marks

    prev_tables = prev.get("tables") or {}

    for t in TABLES:
        cur_rows = tables_data.get(t.key, [])
        old_rows = prev_tables.get(t.key, [])
        cur_by_pk = _index_by_pk(cur_rows, t.pk)
        old_by_pk = _index_by_pk(old_rows, t.pk)

        added = [pk for pk in cur_by_pk if pk not in old_by_pk]
        removed = [pk for pk in old_by_pk if pk not in cur_by_pk]
        changed_pks: list[str] = []

        for i, r in enumerate(cur_rows):
            pk_val = _s(r.get(t.pk))
            if not pk_val:
                continue
            if pk_val in old_by_pk:
                old = old_by_pk[pk_val]
                diff_cols = [c for c in t.columns if _s(r.get(c)) != _s(old.get(c))]
                if diff_cols:
                    changed_pks.append(pk_val)
                    marks[(t.key, i)] = {
                        "change": "changed",
                        "changed_cols": diff_cols,
                        "prev": {c: _s(old.get(c)) for c in diff_cols},
                    }
                    for c in diff_cols:
                        if len(changes["details"]) < MAX_DETAILS:
                            changes["details"].append(
                                {
                                    "table": t.key,
                                    "pk": pk_val,
                                    "col": c,
                                    "old": _s(old.get(c)),
                                    "new": _s(r.get(c)),
                                }
                            )
                        else:
                            changes["truncated"] = True
            elif cur_by_pk.get(pk_val) is r:
                # 새로 생긴 PK (같은 PK가 중복이면 첫 행만 '추가'로 표시)
                marks[(t.key, i)] = {"change": "added", "changed_cols": [], "prev": {}}

        changes["removed_rows"][t.key] = [
            {c: _s(old_by_pk[pk].get(c)) for c in t.columns} for pk in removed
        ]
        changes["summary"][t.key] = {
            "added": len(added),
            "removed": len(removed),
            "changed": len(changed_pks),
        }
        changes["total_added"] += len(added)
        changes["total_removed"] += len(removed)
        changes["total_changed"] += len(changed_pks)

    return changes, marks


def default_changelog_path(out_path: str) -> str:
    stem, _ = os.path.splitext(out_path)
    return stem + "_changelog.json"


def default_lastbuild_path(out_path: str) -> str:
    """[변경] 탭에 보이는 "지난 기준 대비" 비교는 리셋 전까지 고정된 기준
    스냅샷(default_snapshot_path)과 비교합니다 - 리셋을 누르기 전까지는 계속
    쌓여서 보입니다. 하지만 "전체 변경 이력"(누적 changelog)에 매번 그 커진
    전체 diff를 통째로 다시 적으면 같은 항목이 실행마다 중복으로 쌓이므로,
    changelog에는 이번 한 번의 실행에서 실제로 바뀐 것만 넣어야 합니다 - 그
    계산에 쓰는, 매 실행마다 갱신되는 별도의 "직전 실행" 스냅샷입니다."""
    stem, _ = os.path.splitext(out_path)
    return stem + "_lastbuild.json"


def load_changelog(path: str) -> list[dict]:
    if not path or not os.path.exists(path):
        return []
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError) as exc:
        print(f"[알림] 이전 변경 이력을 읽지 못해 새로 시작합니다: {exc}")
        return []
    return data if isinstance(data, list) else []


def build_changelog_entries(
    tables_data: dict[str, list[dict]], changes: dict, marks: dict, generated_at: str,
    actor: str | None = None,
) -> list[dict]:
    """이번 생성에서 감지된 추가/수정/삭제를, 누적 이력에 그대로 추가할 수 있는
    평평한 항목 리스트로 만듭니다(marks는 "added"/"changed"만, 삭제는
    changes["removed_rows"]에 전체 스냅샷 값으로 들어있음).

    actor: 이 변경을 누가 했는지(보통 실시간 입력 서버를 실행 중인 사용자의
    Windows 로그인 계정, 또는 사용자가 직접 정한 표시 이름) - 없으면 "알 수
    없음"으로 남깁니다(정적 생성 스크립트를 실행한 사람을 특정할 수 없는
    경우 등)."""
    actor_val = actor or "알 수 없음"
    entries: list[dict] = []

    for (table_key, row_idx), mark in marks.items():
        schema = TABLE_BY_KEY.get(table_key)
        if not schema:
            continue
        row = tables_data.get(table_key, [])[row_idx]
        pk_val = _s(row.get(schema.pk))

        if mark["change"] == "added":
            entries.append(
                {
                    "generated_at": generated_at,
                    "kind": "added",
                    "table": table_key,
                    "pk": pk_val,
                    "actor": actor_val,
                    "cells": {c: _s(row.get(c)) for c in schema.columns},
                }
            )
        elif mark["change"] == "changed":
            entries.append(
                {
                    "generated_at": generated_at,
                    "kind": "changed",
                    "table": table_key,
                    "pk": pk_val,
                    "actor": actor_val,
                    "changed_cols": mark["changed_cols"],
                    "prev": mark["prev"],
                    "cells": {c: _s(row.get(c)) for c in mark["changed_cols"]},
                }
            )

    for table_key, rows in (changes.get("removed_rows") or {}).items():
        schema = TABLE_BY_KEY.get(table_key)
        if not schema:
            continue
        for cells in rows:
            entries.append(
                {
                    "generated_at": generated_at,
                    "kind": "removed",
                    "table": table_key,
                    "pk": _s(cells.get(schema.pk)),
                    "actor": actor_val,
                    "cells": dict(cells),
                }
            )

    return entries


def append_changelog(path: str, new_entries: list[dict], cap: int = CHANGELOG_MAX) -> list[dict]:
    """기존 이력 뒤에 새 항목을 붙이고, cap을 넘으면 오래된 것부터 버립니다.

    저장까지 하고, 최종 저장된(=화면에 그대로 쓸 수 있는) 리스트를 돌려줍니다.
    """
    combined = load_changelog(path) + new_entries
    if len(combined) > cap:
        combined = combined[-cap:]

    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(combined, f, ensure_ascii=False)
    except OSError as exc:
        print(f"[알림] 변경 이력을 저장하지 못했습니다(다음 실행까지는 화면에서만 보임): {exc}")

    return combined
