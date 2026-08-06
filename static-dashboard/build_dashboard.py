#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""build_dashboard.py — xlsm 또는 CSV 폴더 → PPA현황.html 정적 대시보드 생성.

사용법:
    python3 build_dashboard.py --xlsm=/path/to/PPA파일.xlsm
    python3 build_dashboard.py --csv-dir=/path/to/csv폴더
    python3 build_dashboard.py --xlsm=... --out=다른이름.html

서버 없이 이 스크립트를 실행할 때마다 최신 상태의 HTML 파일 하나가 생성됩니다.
그 파일을 팀원들과 공유(사내망 공유폴더, 메일, Teams 등)하면 됩니다. 편집은
여전히 엑셀에서 하고, 바뀐 내용을 반영하려면 이 스크립트를 다시 실행하세요.

실행할 때마다 데이터 스냅샷(<출력파일>_snapshot.json)을 같은 폴더에 남깁니다.
다음번에 실행하면 그 스냅샷과 비교해서 "지난번 대비 추가/삭제/수정된 항목"이
대시보드의 [변경] 탭과 표 화면에 자동으로 표시됩니다. 별도 설정은 필요 없고,
끄고 싶으면 --no-snapshot 을 붙이면 됩니다.
"""
import argparse
import datetime
import os
import sys

# Windows embeddable Python(python.org "embeddable package")은 pythonXXX._pth
# 파일이 있으면 실행하는 스크립트의 폴더를 자동으로 sys.path에 넣어주지 않습니다
# (일반 Python 설치본과 다른 부분). 그래서 같은 폴더의 ppa_loader.py 등을
# import하려면 이 폴더를 직접 sys.path에 추가해줘야 합니다.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ppa_changes import compute_changes, load_snapshot, save_snapshot
from ppa_dashboard_render import render_dashboard
from ppa_loader import load_from_csv_dir, load_from_xlsm
from ppa_schema import TABLES, validate


def default_snapshot_path(out_path: str) -> str:
    stem, _ = os.path.splitext(out_path)
    return stem + "_snapshot.json"


def build_payload(
    tables_data: dict[str, list[dict]],
    is_demo: bool,
    changes: dict | None = None,
    marks: dict | None = None,
) -> dict:
    validation = validate(tables_data)
    marks = marks or {}

    tables_payload = []
    for t in TABLES:
        rows = tables_data.get(t.key, [])
        rows_payload = []
        for i, r in enumerate(rows):
            err_cols = sorted(validation["error_cols"].get((t.key, i), []))
            mark = marks.get((t.key, i))
            row_obj = {"cells": {c: r.get(c, "") for c in t.columns}, "error_cols": err_cols}
            if mark:
                row_obj["change"] = mark["change"]
                if mark["changed_cols"]:
                    row_obj["changed_cols"] = mark["changed_cols"]
                    row_obj["prev"] = mark["prev"]
            rows_payload.append(row_obj)
        tables_payload.append(
            {
                "key": t.key,
                "label": t.label,
                "pk": t.pk,
                "columns": t.columns,
                "fk_columns": list(t.fk.keys()),
                "fk": t.fk,
                "rows": rows_payload,
            }
        )

    return {
        "generated_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "is_demo": is_demo,
        "tables": tables_payload,
        "validation": {
            "total_errors": len(validation["errors"]),
            "by_table": validation["summary_by_table"],
            "by_item": validation["summary_by_item"],
            "errors": validation["errors"],
        },
        "changes": changes or {"has_prev": False},
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--xlsm", help="원본 xlsm 파일 경로 (openpyxl 필요)")
    parser.add_argument("--csv-dir", help="표별 CSV 파일들이 있는 폴더 (T_발전소.csv 등)")
    parser.add_argument("--out", default="PPA현황.html", help="출력 HTML 파일명 (기본: PPA현황.html)")
    parser.add_argument(
        "--prev",
        help="비교할 이전 스냅샷 파일 경로 (기본: <출력파일>_snapshot.json 을 자동 사용)",
    )
    parser.add_argument(
        "--no-snapshot",
        action="store_true",
        help="이번 실행 결과를 스냅샷으로 저장하지 않음 (변경 비교 기능을 쓰지 않을 때)",
    )
    args = parser.parse_args()

    if not args.xlsm and not args.csv_dir:
        parser.error("--xlsm 또는 --csv-dir 중 하나는 반드시 지정해야 합니다.")

    if args.xlsm:
        try:
            tables_data = load_from_xlsm(args.xlsm)
        except ModuleNotFoundError:
            sys.exit(
                "openpyxl을 불러올 수 없습니다. vendor/ 폴더가 build_dashboard.py와 "
                "같은 위치에 있는지 확인하거나, 'pip install openpyxl'을 시도해보세요. "
                "그래도 안 된다면 --csv-dir 방식(표준 라이브러리만 사용)을 이용하세요."
            )
    else:
        tables_data = load_from_csv_dir(args.csv_dir)

    snapshot_path = args.prev or default_snapshot_path(args.out)
    prev = load_snapshot(snapshot_path)
    changes, marks = compute_changes(tables_data, prev)

    payload = build_payload(tables_data, is_demo=False, changes=changes, marks=marks)

    with open(args.out, "w", encoding="utf-8") as f:
        f.write(render_dashboard(payload))

    if not args.no_snapshot:
        save_snapshot(default_snapshot_path(args.out), tables_data, payload["generated_at"])

    print(f"\n생성 완료: {args.out}")
    print(f"검증 오류: {payload['validation']['total_errors']}건")
    if changes.get("has_prev"):
        print(
            f"지난 생성분 대비 변경: 추가 {changes['total_added']}건 · "
            f"수정 {changes['total_changed']}건 · 삭제 {changes['total_removed']}건"
        )
    else:
        print("이전 스냅샷이 없어 이번에는 변경 비교를 건너뛰었습니다 "
              "(다음 실행부터 [변경] 탭에 표시됩니다).")


if __name__ == "__main__":
    main()
