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
"""
import argparse
import datetime
import sys

from ppa_dashboard_render import render_dashboard
from ppa_loader import load_from_csv_dir, load_from_xlsm
from ppa_schema import TABLES, validate


def build_payload(tables_data: dict[str, list[dict]], is_demo: bool) -> dict:
    validation = validate(tables_data)

    tables_payload = []
    for t in TABLES:
        rows = tables_data.get(t.key, [])
        rows_payload = []
        for i, r in enumerate(rows):
            err_cols = sorted(validation["error_cols"].get((t.key, i), []))
            rows_payload.append({"cells": {c: r.get(c, "") for c in t.columns}, "error_cols": err_cols})
        tables_payload.append(
            {
                "key": t.key,
                "label": t.label,
                "pk": t.pk,
                "columns": t.columns,
                "fk_columns": list(t.fk.keys()),
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
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--xlsm", help="원본 xlsm 파일 경로 (openpyxl 필요)")
    parser.add_argument("--csv-dir", help="표별 CSV 파일들이 있는 폴더 (T_발전소.csv 등)")
    parser.add_argument("--out", default="PPA현황.html", help="출력 HTML 파일명 (기본: PPA현황.html)")
    args = parser.parse_args()

    if not args.xlsm and not args.csv_dir:
        parser.error("--xlsm 또는 --csv-dir 중 하나는 반드시 지정해야 합니다.")

    if args.xlsm:
        try:
            tables_data = load_from_xlsm(args.xlsm)
        except ModuleNotFoundError:
            sys.exit(
                "openpyxl이 설치되어 있지 않습니다. 'pip install openpyxl'이 안 된다면 "
                "--csv-dir 방식(표준 라이브러리만 사용)을 이용하세요."
            )
    else:
        tables_data = load_from_csv_dir(args.csv_dir)

    payload = build_payload(tables_data, is_demo=False)

    with open(args.out, "w", encoding="utf-8") as f:
        f.write(render_dashboard(payload))

    print(f"\n생성 완료: {args.out}")
    print(f"검증 오류: {payload['validation']['total_errors']}건")


if __name__ == "__main__":
    main()
