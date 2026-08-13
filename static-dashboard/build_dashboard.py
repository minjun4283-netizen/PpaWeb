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
import zipfile

# Windows embeddable Python(python.org "embeddable package")은 pythonXXX._pth
# 파일이 있으면 실행하는 스크립트의 폴더를 자동으로 sys.path에 넣어주지 않습니다
# (일반 Python 설치본과 다른 부분). 그래서 같은 폴더의 ppa_loader.py 등을
# import하려면 이 폴더를 직접 sys.path에 추가해줘야 합니다.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ppa_changes import (
    append_changelog,
    build_changelog_entries,
    compute_changes,
    default_changelog_path,
    default_lastbuild_path,
    load_changelog,
    load_snapshot,
    save_snapshot,
)
from ppa_dashboard_render import render_dashboard
from ppa_loader import load_from_csv_dir, load_from_xlsm
from ppa_schema import TABLES, validate


def diagnose_xlsm_open_failure(path: str) -> str:
    """xlsm을 열지 못했을 때, 흔한 원인을 스스로 점검해서 원인에 맞는 안내를 만듭니다."""
    lines = [f"엑셀 파일을 열지 못했습니다: {path}", ""]

    if not os.path.exists(path):
        lines.append("→ 이 경로에 파일이 없습니다. 경로(특히 폴더 이름의 띄어쓰기/오타)를 다시 확인해주세요.")
        return "\n".join(lines)

    folder = os.path.dirname(os.path.abspath(path))
    lock_path = os.path.join(folder, "~$" + os.path.basename(path))
    size = os.path.getsize(path)

    checked = False
    if size == 0:
        lines.append("→ 원인으로 보이는 것: 파일 크기가 0바이트입니다.")
        lines.append("   OneDrive 동기화가 아직 끝나지 않았거나 파일 복사가 중간에 끊겼을 수 있습니다.")
        lines.append("   해당 파일을 탐색기에서 더블클릭해 엑셀로 완전히 열어본 뒤 닫고 다시 시도해주세요.")
        checked = True
    elif size < 10_000:
        lines.append(f"→ 원인으로 보이는 것: 파일 크기가 {size:,}바이트로 비정상적으로 작습니다.")
        lines.append("   정상 xlsm이 아니라 OneDrive 자리표시자(placeholder)이거나 손상된 파일일 수 있습니다.")
        checked = True

    if os.path.exists(lock_path):
        lines.append("→ 원인으로 보이는 것: 같은 폴더에 임시 잠금 파일(" + os.path.basename(lock_path) + ")이 있습니다.")
        lines.append("   → 지금 엑셀에서 이 파일이 열려 있다는 뜻입니다. 엑셀을 완전히 닫은 뒤(저장은 먼저 해두고) 다시 실행해주세요.")
        checked = True

    if not checked:
        lines.append("→ 흔한 원인 체크리스트:")
        lines.append("   1) OneDrive에서 이 파일이 구름 아이콘(클라우드 전용)으로 표시되나요?")
        lines.append("      탐색기에서 더블클릭해 완전히 내려받아지도록(초록 체크로 바뀔 때까지) 기다린 뒤 다시 시도하세요.")
        lines.append("   2) 지금 엑셀에서 이 파일이 열려 있다면, 저장 후 닫고 다시 시도하세요.")
        lines.append("   3) 확장자만 .xlsm이고 실제로는 예전 .xls(97-2003) 형식으로 저장된 파일은 아닌지 확인하세요.")
        lines.append("      (엑셀에서 열어 '다른 이름으로 저장' → '엑셀 매크로 사용 통합 문서(*.xlsm)'으로 다시 저장해보세요.)")
        lines.append("   4) 위 방법으로도 안 되면, 엑셀에서 이 파일을 열어 '다른 이름으로 저장'으로 새 사본을 만들고")
        lines.append("      그 사본 경로로 --xlsm 을 다시 지정해보세요 (원본이 손상됐을 가능성을 우회합니다).")

    return "\n".join(lines)


def default_snapshot_path(out_path: str) -> str:
    stem, _ = os.path.splitext(out_path)
    return stem + "_snapshot.json"


def build_payload(
    tables_data: dict[str, list[dict]],
    is_demo: bool,
    changes: dict | None = None,
    marks: dict | None = None,
    changelog: list | None = None,
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
        "changelog": changelog or [],
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
    parser.add_argument(
        "--reset-snapshot",
        action="store_true",
        help="[변경] 탭의 비교 기준점을 이번 실행 결과로 리셋합니다(그 전까지 누적된 "
             "변경 표시가 초기화됩니다). 지정하지 않으면 기준점은 그대로 유지되고 "
             "변경이 계속 누적되어 보입니다. 전체 변경 이력(_changelog.json)은 이 "
             "옵션과 무관하게 계속 보존됩니다.",
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
        except zipfile.BadZipFile:
            sys.exit(diagnose_xlsm_open_failure(args.xlsm))
        except FileNotFoundError:
            sys.exit(f"파일을 찾을 수 없습니다: {args.xlsm}\n경로를 다시 확인해주세요.")
        except PermissionError:
            sys.exit(
                f"파일을 읽을 권한이 없습니다: {args.xlsm}\n"
                "엑셀에서 이 파일을 열어둔 상태라면 닫고 다시 시도해주세요."
            )
    else:
        tables_data = load_from_csv_dir(args.csv_dir)

    # baseline: [변경] 탭에 보이는 "지난 기준 대비" 비교 대상 - 리셋을 누르기
    # 전까지는 절대 안 바뀝니다(그래서 그 사이 변경이 계속 누적돼 보임).
    # lastbuild: 이번 실행 한 번의 변경분만 골라 전체 변경 이력(changelog)에
    # 정확히 한 번씩만 추가하기 위한 내부용 스냅샷 - 항상 매 실행마다 갱신.
    baseline_path = args.prev or default_snapshot_path(args.out)
    lastbuild_path = default_lastbuild_path(args.out)
    baseline = load_snapshot(baseline_path)
    changes, marks = compute_changes(tables_data, baseline)

    lastbuild = load_snapshot(lastbuild_path)
    build_changes, build_marks = compute_changes(tables_data, lastbuild)

    changelog_path = default_changelog_path(args.out)
    changelog = load_changelog(changelog_path)

    generated_at = datetime.datetime.now().isoformat(timespec="seconds")
    if build_changes.get("has_prev") and not args.no_snapshot:
        new_entries = build_changelog_entries(tables_data, build_changes, build_marks, generated_at)
        if new_entries:
            changelog = append_changelog(changelog_path, new_entries)

    payload = build_payload(tables_data, is_demo=False, changes=changes, marks=marks, changelog=changelog)
    payload["generated_at"] = generated_at

    with open(args.out, "w", encoding="utf-8") as f:
        f.write(render_dashboard(payload))

    if not args.no_snapshot:
        save_snapshot(lastbuild_path, tables_data, payload["generated_at"])
        if args.reset_snapshot or baseline is None:
            save_snapshot(baseline_path, tables_data, payload["generated_at"])

    print(f"\n생성 완료: {args.out}")
    print(f"검증 오류: {payload['validation']['total_errors']}건")
    if changes.get("has_prev"):
        print(
            f"기준점 대비 누적 변경: 추가 {changes['total_added']}건 · "
            f"수정 {changes['total_changed']}건 · 삭제 {changes['total_removed']}건"
        )
        if args.reset_snapshot:
            print("--reset-snapshot 지정됨: 이번 실행 결과를 새 기준점으로 리셋했습니다.")
        else:
            print("기준점은 --reset-snapshot 을 주기 전까지 그대로 유지됩니다.")
    else:
        print("기준 스냅샷이 없어 이번에는 변경 비교를 건너뛰었습니다 "
              "(다음 실행부터 [변경] 탭에 표시되고, 이번 실행 결과가 새 기준점이 됩니다).")


if __name__ == "__main__":
    main()
