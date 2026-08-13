#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""excel_com.py — Windows COM(pywin32)으로 "지금 열려 있는" 엑셀에 직접 읽고 씁니다.

openpyxl은 최신 zip 기반 xlsm만 열 수 있는데, 운영 파일이 예전 바이너리 포맷
(레거시 CFB, 확장자만 .xlsm)으로 저장돼 있으면 열지 못합니다(BadZipFile).
이 모듈은 파일을 직접 읽지 않고 "엑셀 프로그램 자체"를 원격조종해서 읽고
쓰므로, 파일 포맷과 무관하게 동작하고, 지금 화면에 열려 있는 통합문서에
바로 반영됩니다.

동작 방식:
  - 이미 엑셀이 그 파일을 열어둔 상태면 → 그 세션에 그대로 씀(사용자가 보는
    화면이 바로 바뀌고, 저장도 그 세션에서 이뤄집니다).
  - 열려 있지 않으면 → 보이지 않는 새 엑셀 인스턴스를 띄워 열고 저장한 뒤,
    그 인스턴스만 종료합니다(사용자가 나중에 직접 여는 세션에는 영향 없음).

모든 COM 호출은 전용 백그라운드 스레드 하나에서만 실행됩니다(COM은 이 호출을
만든 스레드에서만 안전하게 쓸 수 있고, 저장 요청이 몰려도 한 번에 하나씩만
처리되도록 직렬화하기 위함 — 동시에 두 저장이 같은 행을 두고 경쟁하는 상황을
막습니다).
"""
from __future__ import annotations

import datetime
import os
import queue
import re
import threading
from typing import Optional

from ppa_schema import TABLE_BY_KEY, TABLES

try:
    import pythoncom
    import win32com.client
    _IMPORT_ERROR: Optional[Exception] = None
except ImportError as exc:  # pywin32 미설치 환경에서도 모듈 자체는 import 가능해야 함
    pythoncom = None
    win32com = None
    _IMPORT_ERROR = exc


LABEL_CANDIDATES: dict[str, list[str]] = {
    "T_발전소": ["발전소명", "발전법인명"],
    "T_구매계약": ["발전소ID", "구매 담당자"],
    "T_수요기업": ["기업명"],
    "T_판매계약": ["수요기업ID", "계약유형"],
    "T_전기사용지": ["전기사용지명", "판매계약ID"],
    "T_수급매칭": ["현황", "전기사용지ID", "구매계약ID"],
}


class ExcelComError(RuntimeError):
    pass


def _require_pywin32() -> None:
    if _IMPORT_ERROR is not None:
        raise ExcelComError(
            "pywin32가 설치되어 있지 않습니다 (import 오류: " + str(_IMPORT_ERROR) + ").\n"
            "이 기능(웹 화면에서 입력 → 엑셀에 실시간 반영)은 Windows + pywin32가 있어야 "
            "동작합니다. static-dashboard/README.md 의 'pywin32 설치' 안내를 참고하세요."
        )


def _is_date_column(name: str) -> bool:
    return bool(re.search(r"일자|날짜|시작일|종료일|체결일|계약일", name))


def _is_number_column(name: str) -> bool:
    return bool(re.search(r"용량|금액|단가|비율|수량|면적|사용량|발전량|REC|MW|MWh|kW|kWh|개월|년수", name))


def _coerce_for_excel(col_name: str, raw_value):
    value = "" if raw_value is None else str(raw_value).strip()
    if value == "":
        return ""

    if value.upper() in ("TRUE", "FALSE"):
        return value.upper() == "TRUE"

    if _is_date_column(col_name):
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
            try:
                return datetime.datetime.strptime(value, fmt)
            except ValueError:
                continue
        return value

    if _is_number_column(col_name):
        try:
            f = float(value)
            return int(f) if f.is_integer() else f
        except ValueError:
            return value

    return value


def _cell_to_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, datetime.datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else str(value)
    return str(value).strip()


class _Job:
    __slots__ = ("fn", "args", "result_q")

    def __init__(self, fn, args):
        self.fn = fn
        self.args = args
        self.result_q: "queue.Queue" = queue.Queue(maxsize=1)


class ExcelBridge:
    """전용 스레드 하나에서 모든 COM 호출을 직렬화해서 실행합니다."""

    def __init__(self, xlsm_path: str):
        _require_pywin32()
        self.xlsm_path = os.path.abspath(xlsm_path)
        self._jobs: "queue.Queue" = queue.Queue()
        self._app = None
        self._we_launched_app = False
        self._thread = threading.Thread(target=self._run, name="ExcelComWorker", daemon=True)
        self._thread.start()

    # ---- 공개 API: 각각 워커 스레드로 작업을 넘기고 결과를 기다립니다 ----
    def read_table(self, table_key: str) -> dict:
        return self._call(self._read_table, table_key)

    def read_all_tables(self) -> dict[str, list[dict]]:
        return self._call(self._read_all_tables)

    def get_record(self, table_key: str, pk_value: str) -> dict:
        return self._call(self._get_record, table_key, pk_value)

    def get_options(self, table_key: str) -> list[dict]:
        return self._call(self._get_options, table_key)

    def save_record(self, table_key: str, record: dict) -> dict:
        return self._call(self._save_record, table_key, record)

    def shutdown(self) -> None:
        self._jobs.put(None)

    # ---- 내부: 작업 큐 ----
    def _call(self, fn, *args):
        job = _Job(fn, args)
        self._jobs.put(job)
        ok, payload = job.result_q.get()
        if not ok:
            raise payload
        return payload

    def _run(self) -> None:
        pythoncom.CoInitialize()
        try:
            while True:
                job = self._jobs.get()
                if job is None:
                    break
                try:
                    result = job.fn(*job.args)
                    job.result_q.put((True, result))
                except Exception as exc:  # noqa: BLE001 - 워커는 죽지 않고 호출자에게 되돌려줌
                    job.result_q.put((False, exc))
        finally:
            self._close_if_we_launched()
            pythoncom.CoUninitialize()

    # ---- 워크북 확보: 이미 열려 있으면 그 세션, 아니면 숨김 인스턴스로 새로 염 ----
    #
    # 예전에는 GetObject(Class="Excel.Application")로 "떠 있는 아무 엑셀"을
    # 가져온 뒤 그 Workbooks를 우리가 직접 순회하며 전체 경로 문자열을
    # 비교했는데, 두 가지로 깨졌습니다:
    #   1) 엑셀 프로세스가 여러 개 떠 있으면 Class= 로는 "그" 프로세스가
    #      아니라 그냥 아무 인스턴스가 잡혀서, 정작 파일이 열려 있는
    #      프로세스를 놓칠 수 있음
    #   2) 경로 문자열 비교(normcase/abspath)가 한글 자모 정규화(NFC/NFD)나
    #      OneDrive 경로 표기 차이로 어긋나면, 실제로는 같은 파일인데도
    #      "안 열려 있다"고 오판 → 이미 열려 있는 파일을 또 열려다 엑셀이
    #      "같은 이름의 통합 문서가 이미 열려 있다"며 거부
    # 지금은 파일 경로 자체를 모니커로 GetObject에 넘깁니다 - 엑셀은 열려
    # 있는 통합문서마다 이 방식으로 바로 찾아지도록 스스로 등록해두므로,
    # 어느 프로세스에 열려 있든, 문자열 비교 없이 정확히 그 워크북을
    # 돌려받습니다(안 열려 있으면 예외만 남기고 정상적으로 다음 단계로).
    def _ensure_workbook(self):
        try:
            return win32com.client.GetObject(self.xlsm_path)
        except Exception:
            pass

        # 우리가 이전에 띄워둔 숨김 인스턴스가 있으면 재사용(매번 새로 띄우면
        # 느립니다). 같은 이름의 파일을 동시에 두 번 열 수 없다는 엑셀 자체의
        # 제약 덕분에, 파일명만 같아도 그게 우리 파일이라고 확신할 수 있습니다.
        if self._app is not None and self._we_launched_app:
            try:
                target_name = os.path.basename(self.xlsm_path).lower()
                for wb in self._app.Workbooks:
                    if os.path.basename(wb.FullName).lower() == target_name:
                        return wb
            except Exception:
                self._app = None
                self._we_launched_app = False

        if not os.path.exists(self.xlsm_path):
            raise ExcelComError(f"엑셀 파일을 찾을 수 없습니다: {self.xlsm_path}")

        try:
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            wb = app.Workbooks.Open(self.xlsm_path)
        except Exception as exc:
            raise ExcelComError(
                f"엑셀 파일을 열지 못했습니다: {self.xlsm_path}\n"
                "다른 프로그램이 사용 중이거나 경로가 올바르지 않을 수 있습니다.\n"
                f"원본 오류: {exc}"
            ) from exc

        self._app = app
        self._we_launched_app = True
        return wb

    def _close_if_we_launched(self) -> None:
        if self._app is not None and self._we_launched_app:
            try:
                for wb in list(self._app.Workbooks):
                    wb.Close(SaveChanges=False)
                self._app.Quit()
            except Exception:
                pass
        self._app = None
        self._we_launched_app = False

    def _worksheet(self, wb, table_key: str):
        try:
            return wb.Worksheets(table_key)
        except Exception as exc:
            raise ExcelComError(f"시트를 찾을 수 없습니다: {table_key}") from exc

    def _read_grid(self, ws):
        """UsedRange 전체를 한 번에 읽어옵니다(셀 단위 COM 호출을 피해 속도 확보)."""
        used = ws.UsedRange
        values = used.Value
        if values is None:
            return [], 1, 1
        if not isinstance(values, tuple):
            values = ((values,),)
        elif values and not isinstance(values[0], tuple):
            values = (values,)
        return values, used.Row, used.Column

    def _header_map(self, values, start_col: int) -> dict[str, int]:
        """열 이름 → 절대 열 번호(Cells()에 그대로 쓸 수 있는 1-based 인덱스)."""
        if not values:
            return {}
        header_row = values[0]
        out: dict[str, int] = {}
        for i, h in enumerate(header_row):
            name = _cell_to_text(h)
            if name:
                out[name] = start_col + i
        return out

    # ---- 읽기 ----
    def _read_table(self, table_key: str) -> dict:
        schema = TABLE_BY_KEY[table_key]
        wb = self._ensure_workbook()
        ws = self._worksheet(wb, table_key)
        values, start_row, start_col = self._read_grid(ws)

        if not values:
            return {"headers": schema.columns, "rows": []}

        col_at = self._header_map(values, start_col)
        col_idx = {name: (abs_col - start_col) for name, abs_col in col_at.items()}

        rows = []
        for raw in values[1:]:
            record = {}
            has_any = False
            for col in schema.columns:
                idx = col_idx.get(col)
                val = _cell_to_text(raw[idx]) if idx is not None and idx < len(raw) else ""
                record[col] = val
                if val != "":
                    has_any = True
            if has_any:
                rows.append(record)

        return {"headers": schema.columns, "rows": rows}

    def _read_all_tables(self) -> dict[str, list[dict]]:
        return {t.key: self._read_table(t.key)["rows"] for t in TABLES}

    def _get_record(self, table_key: str, pk_value: str) -> dict:
        schema = TABLE_BY_KEY[table_key]
        target = str(pk_value or "").strip()
        for row in self._read_table(table_key)["rows"]:
            if str(row.get(schema.pk) or "").strip() == target:
                return row
        return {}

    def _get_options(self, table_key: str) -> list[dict]:
        schema = TABLE_BY_KEY[table_key]
        label_fields = LABEL_CANDIDATES.get(table_key, [])
        seen = set()
        options = []
        for row in self._read_table(table_key)["rows"]:
            value = str(row.get(schema.pk) or "").strip()
            if not value or value in seen:
                continue
            seen.add(value)
            extra = ""
            for field in label_fields:
                v = str(row.get(field) or "").strip()
                if v:
                    extra = v
                    break
            options.append({"value": value, "label": f"{value} | {extra}" if extra else value})
        return options

    # ---- 검증 (엑셀에 쓰기 전에 확인 — VBA 입력폼과 동일한 최소 기준) ----
    def _validate(self, table_key: str, record: dict) -> list[str]:
        schema = TABLE_BY_KEY[table_key]
        errors = []

        pk_val = str(record.get(schema.pk) or "").strip()
        if not pk_val:
            errors.append(f"{schema.pk}는 필수입니다.")

        for fk_col, ref_key in schema.fk.items():
            val = str(record.get(fk_col) or "").strip()
            if not val:
                errors.append(f"{fk_col}는 필수입니다.")
                continue
            ref_schema = TABLE_BY_KEY[ref_key]
            ref_values = {
                str(r.get(ref_schema.pk) or "").strip() for r in self._read_table(ref_key)["rows"]
            }
            if val not in ref_values:
                errors.append(f"{fk_col} 값 '{val}'을(를) {ref_key}에서 찾을 수 없습니다.")

        return errors

    # ---- 쓰기 ----
    def _save_record(self, table_key: str, record: dict) -> dict:
        schema = TABLE_BY_KEY[table_key]

        errors = self._validate(table_key, record)
        if errors:
            raise ExcelComError(" / ".join(errors))

        pk_value = str(record.get(schema.pk) or "").strip()
        wb = self._ensure_workbook()
        ws = self._worksheet(wb, table_key)

        values, start_row, start_col = self._read_grid(ws)
        if not values:
            raise ExcelComError(f"{table_key} 시트에서 머리글 행을 찾지 못했습니다.")

        col_at = self._header_map(values, start_col)
        missing = [c for c in schema.columns if c not in col_at]
        if missing:
            raise ExcelComError(f"{table_key} 시트에서 다음 열을 찾지 못했습니다: {', '.join(missing)}")

        pk_col_abs = col_at[schema.pk]
        pk_idx = pk_col_abs - start_col
        target_row = None
        for offset, raw in enumerate(values[1:], start=1):
            val = _cell_to_text(raw[pk_idx]) if pk_idx < len(raw) else ""
            if val == pk_value:
                target_row = start_row + offset
                break

        action = "updated"
        if target_row is None:
            action = "inserted"
            target_row = self._append_row_target(ws, start_row, len(values))

        for col, abs_col in col_at.items():
            if col not in record:
                continue
            ws.Cells(target_row, abs_col).Value = _coerce_for_excel(col, record.get(col, ""))

        wb.Save()

        return {"action": action, "table": table_key, "pk_value": pk_value, "row": target_row}

    def _append_row_target(self, ws, start_row: int, used_row_count: int) -> int:
        """새 행을 추가할 위치. 시트에 표(ListObject)가 있으면 표를 확장해서
        서식/필터가 그대로 이어지게 하고, 없으면 마지막 사용 행 바로 다음에 씁니다.
        """
        try:
            if ws.ListObjects.Count >= 1:
                lo = ws.ListObjects(1)
                new_row = lo.ListRows.Add()
                return new_row.Range.Row
        except Exception:
            pass
        return start_row + used_row_count
