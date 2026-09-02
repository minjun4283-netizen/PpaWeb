#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ppa_archive.py — 엑셀 저장/대시보드 종료 시점마다 xlsm+html을 타임스탬프
사본으로 archive/ 폴더에 남기는 간단한 백업.

excel_com.py(웹 간편입력/탐색/일괄편집으로 저장할 때마다)와
ppa_liveserver.py(서버 종료 시점)가 이 모듈의 backup_now() 하나만 불러
씁니다 - 로직을 한 곳에만 두어 두 호출부가 어긋나지 않게 합니다.

백업은 있으면 좋은 안전장치일 뿐이라, 여기서 나는 오류(디스크 공간 부족
등)가 원래 하려던 저장/종료 자체를 막으면 안 됩니다 - 그래서 backup_now는
예외를 절대 밖으로 던지지 않고, 실패하면 조용히 넘어갑니다(호출부에서
로그만 남기고 싶으면 반환값의 error를 보면 됩니다).
"""
from __future__ import annotations

import datetime
import glob
import os
import shutil
from typing import Optional

KEEP_PER_EXT = 100  # 확장자별로 보관할 백업 최대 개수


def backup_now(xlsm_path: Optional[str], html_path: Optional[str], archive_dir: str) -> dict:
    """xlsm_path/html_path 각각을 archive_dir에 타임스탬프 사본으로 복사하고,
    오래된 백업을 정리합니다. 두 경로 모두 선택 사항(존재하는 파일만 복사) -
    아직 html이 한 번도 안 만들어졌으면 xlsm만 백업합니다.
    """
    result = {"ok": True, "copied": [], "error": None}
    try:
        os.makedirs(archive_dir, exist_ok=True)
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

        if xlsm_path and os.path.exists(xlsm_path):
            dest = os.path.join(archive_dir, f"PPA파일_{ts}.xlsm")
            shutil.copy2(xlsm_path, dest)
            result["copied"].append(dest)

        if html_path and os.path.exists(html_path):
            dest = os.path.join(archive_dir, f"PPA현황_{ts}.html")
            shutil.copy2(html_path, dest)
            result["copied"].append(dest)

        _prune(archive_dir, "*.xlsm")
        _prune(archive_dir, "*.html")
    except Exception as e:  # noqa: BLE001 - 백업 실패가 저장/종료를 막으면 안 됨
        result["ok"] = False
        result["error"] = str(e)
    return result


def _prune(archive_dir: str, pattern: str) -> None:
    """파일명에 타임스탬프가 들어 있어 이름순 정렬 = 시간순 정렬입니다.
    KEEP_PER_EXT개를 넘으면 가장 오래된 것부터 지웁니다."""
    files = sorted(glob.glob(os.path.join(archive_dir, pattern)))
    excess = len(files) - KEEP_PER_EXT
    if excess <= 0:
        return
    for path in files[:excess]:
        try:
            os.remove(path)
        except OSError:
            pass
