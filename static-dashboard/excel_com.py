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

    def get_references(self, table_key: str, pk_value: str) -> list[dict]:
        return self._call(self._referencing_records, table_key, pk_value)

    def delete_record(self, table_key: str, pk_value: str, force: bool = False) -> dict:
        return self._call(self._delete_record, table_key, pk_value, force)

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
                    # 우리가 이번 작업 때문에 숨김 엑셀을 새로 띄웠다면, 작업이 끝나는
                    # 즉시 닫습니다. 서버가 떠 있는 동안 계속 붙잡고 있으면(이전
                    # 방식), 그 파일을 아무도 안 쓰고 있어도 "다른 프로그램이 사용
                    # 중"이라며 읽기 전용으로 열리는 원인이 됩니다 — 특히 서버가
                    # 정상 종료되지 않고 콘솔 창만 닫히면 그 숨김 인스턴스가 백그
                    # 라운드에 영영 남아있을 수 있었습니다. 파일이 이미 사용자의
                    # 엑셀에 열려 있는 평소 상황(가장 흔한 경우)에는 우리가 새로
                    # 띄운 게 없으므로 이 정리는 아무 일도 하지 않습니다.
                    self._close_if_we_launched()
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

        # 이 경로가 애초에 로컬 경로가 아니면(OneDrive/SharePoint 자동 저장 중
        # ThisWorkbook.Path가 "https://..."를 돌려주는 경우), 아래의 파일 존재
        # 확인이나 DispatchEx로 넘어가봐야 이유를 알 수 없는 오류만 납니다.
        # VBA 쪽에 이미 같은 검사(로컬경로_확인)가 있어 정상적인 경로로는 이
        # 값이 여기까지 넘어오지 않지만, .bat을 직접 실행하는 등 그 검사를
        # 거치지 않은 경우를 대비해 여기서도 명확한 원인을 알려줍니다.
        if self.xlsm_path.lower().startswith("http"):
            raise ExcelComError(
                "이 파일 경로가 OneDrive/SharePoint 클라우드 주소(https://...)로 "
                "전달됐습니다 - 로컬 경로(C:\\...)가 필요합니다.\n"
                "엑셀에서 파일 → 정보 → 자동 저장(AutoSave)을 끄고 다시 시도하거나, "
                "탐색기의 OneDrive 동기화 폴더에서 이 파일을 직접 열고 다시 시도해주세요."
            )

        target_name = os.path.basename(self.xlsm_path).lower()

        # 위의 파일 경로 모니커로 못 찾았어도, 실제로는 열려 있을 수 있습니다 -
        # OneDrive/SharePoint 자동 저장(AutoSave)으로 열린 통합문서는 엑셀이
        # 자기 자신을 로컬 경로가 아니라 클라우드 URL 신원으로 등록해두는
        # 경우가 있어서, 로컬 경로 모니커로는 찾아지지 않습니다. 이럴 때를
        # 대비해 떠 있는 엑셀 프로세스의 워크북들을 파일명만으로 다시
        # 확인합니다 - 같은 이름의 파일을 엑셀에서 동시에 두 개 열 수 없다는
        # 제약 덕분에, 파일명만 같아도 그게 우리 파일이라고 확신할 수 있습니다.
        try:
            running = win32com.client.GetObject(Class="Excel.Application")
            for wb in running.Workbooks:
                try:
                    if os.path.basename(wb.Name).lower() == target_name:
                        return wb
                except Exception:
                    continue
        except Exception:
            pass

        # 우리가 이전에 띄워둔 숨김 인스턴스가 있으면 재사용(매번 새로 띄우면
        # 느립니다).
        if self._app is not None and self._we_launched_app:
            try:
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
            app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable - 매크로 자동 실행 금지

            # 새로 띄운 인스턴스는 사용자가 평소 쓰는 애드인(OpenSolver.xlam 등)을
            # 자동으로 같이 불러옵니다. 우리는 데이터만 읽고 쓰면 되고, 이 애드인들을
            # 열어두면 그 파일들까지 우리가 잠그게 되어(사용자 본인 이름으로 "편집
            # 중"이라고 뜸) 나중에 사용자가 직접 그 애드인을 열려고 할 때 방해가
            # 됩니다. 우리 파일을 열기 전에 자동으로 딸려온 것들을 먼저 다 닫습니다.
            for auto_wb in list(app.Workbooks):
                try:
                    auto_wb.Close(SaveChanges=False)
                except Exception:
                    pass

            wb = app.Workbooks.Open(self.xlsm_path)
        except Exception as exc:
            raise ExcelComError(f"{self._diagnose_open_failure()}\n\n원본 오류: {exc}") from exc

        self._app = app
        self._we_launched_app = True
        return wb

    def _diagnose_open_failure(self) -> str:
        """열기가 실패했을 때, 흔한 원인을 스스로 점검해서 원인에 맞는 안내를 만듭니다."""
        path = self.xlsm_path
        lines = [f"엑셀 파일을 열지 못했습니다: {path}"]

        if not os.path.exists(path):
            lines.append("→ 이 경로에 파일이 없습니다. 경로(특히 폴더 이름의 띄어쓰기/오타)를 다시 확인해주세요.")
            return "\n".join(lines)

        size = os.path.getsize(path)
        folder = os.path.dirname(path)
        lock_path = os.path.join(folder, "~$" + os.path.basename(path))

        if size == 0:
            lines.append("→ 파일 크기가 0바이트입니다. OneDrive에서 아직 완전히 내려받아지지 않은")
            lines.append("   (클라우드 전용) 파일일 수 있습니다. 탐색기에서 이 파일을 더블클릭해")
            lines.append("   완전히 내려받아지도록(초록 체크로 바뀔 때까지) 한 뒤 다시 시도해주세요.")
        elif os.path.exists(lock_path):
            lines.append(f"→ 잠금 파일이 있습니다({os.path.basename(lock_path)}) - 이미 다른 곳에서")
            lines.append("   열려 있다는 뜻입니다. 그 세션에서 저장 후 닫고 다시 시도해주세요.")
        else:
            lines.append("→ 다른 프로그램이 이 파일을 열어두고 있거나, OneDrive 동기화가 아직")
            lines.append("   끝나지 않았을 수 있습니다.")

        return "\n".join(lines)

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

    # ---- PK로 행 찾기 (저장/삭제가 공용으로 씀) ----
    def _locate_row(self, ws, schema, pk_value: str):
        values, start_row, start_col = self._read_grid(ws)
        if not values:
            raise ExcelComError(f"{schema.key} 시트에서 머리글 행을 찾지 못했습니다.")

        col_at = self._header_map(values, start_col)
        if schema.pk not in col_at:
            raise ExcelComError(f"{schema.key} 시트에서 '{schema.pk}' 열을 찾지 못했습니다.")

        pk_idx = col_at[schema.pk] - start_col
        target_row = None
        for offset, raw in enumerate(values[1:], start=1):
            val = _cell_to_text(raw[pk_idx]) if pk_idx < len(raw) else ""
            if val == pk_value:
                target_row = start_row + offset
                break

        return values, start_row, start_col, col_at, target_row

    # ---- 쓰기 ----
    def _save_record(self, table_key: str, record: dict) -> dict:
        schema = TABLE_BY_KEY[table_key]

        errors = self._validate(table_key, record)
        if errors:
            raise ExcelComError(" / ".join(errors))

        pk_value = str(record.get(schema.pk) or "").strip()
        wb = self._ensure_workbook()
        ws = self._worksheet(wb, table_key)

        values, start_row, start_col, col_at, target_row = self._locate_row(ws, schema, pk_value)

        missing = [c for c in schema.columns if c not in col_at]
        if missing:
            raise ExcelComError(f"{table_key} 시트에서 다음 열을 찾지 못했습니다: {', '.join(missing)}")

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

    # ---- 참조 확인 (삭제 전에 다른 표가 이 PK를 쓰고 있는지) ----
    def _referencing_records(self, table_key: str, pk_value: str) -> list[dict]:
        target = str(pk_value or "").strip()
        refs: list[dict] = []

        for t in TABLES:
            for fk_col, ref_key in t.fk.items():
                if ref_key != table_key:
                    continue
                rows = self._read_table(t.key)["rows"]
                count = sum(1 for r in rows if str(r.get(fk_col) or "").strip() == target)
                if count > 0:
                    refs.append({"table": t.key, "label": t.label, "fk_col": fk_col, "count": count})

        return refs

    # ---- 삭제 ----
    def _delete_record(self, table_key: str, pk_value: str, force: bool = False) -> dict:
        schema = TABLE_BY_KEY[table_key]
        pk_value = str(pk_value or "").strip()
        if not pk_value:
            raise ExcelComError(f"{schema.pk}는 필수입니다.")

        refs = self._referencing_records(table_key, pk_value)
        if refs and not force:
            detail = ", ".join(f"{r['label']} {r['count']}건({r['fk_col']})" for r in refs)
            raise ExcelComError(
                f"다른 표에서 이 {schema.pk}({pk_value})를 참조하고 있어 삭제할 수 없습니다: {detail}. "
                "참조하는 데이터를 먼저 정리한 뒤 다시 시도해주세요."
            )

        wb = self._ensure_workbook()
        ws = self._worksheet(wb, table_key)
        values, start_row, start_col, col_at, target_row = self._locate_row(ws, schema, pk_value)

        if target_row is None:
            raise ExcelComError(
                f"{schema.pk} '{pk_value}'을(를) {table_key}에서 찾지 못했습니다(이미 삭제됐을 수 있습니다)."
            )

        # 표(ListObject) 안에 있는 행이든 아니든, 행 전체 삭제는 엑셀이 알아서
        # 표 범위/서식을 줄이고 나머지 행을 한 칸씩 끌어올립니다 - 별도의
        # ListObject 전용 삭제 API가 필요 없습니다. 이 시트들은 검증/파생값을
        # 전부 이 시스템이 계산하고 엑셀 수식을 쓰지 않으므로(스키마 설계상),
        # 셀 참조가 깨질 걱정도 없습니다.
        ws.Rows(target_row).Delete()
        wb.Save()

        return {"action": "deleted", "table": table_key, "pk_value": pk_value}

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
