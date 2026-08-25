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
    # "공급기한_구매"/"공급기한_판매"처럼 "기한"으로 끝나는 컬럼도 날짜로
    # 다뤄야 엑셀에 진짜 날짜 값(datetime)으로 써져서 정렬/기간필터/D-day
    # 계산이 제대로 됩니다 - 문자열로 남으면 그런 계산이 깨집니다.
    return bool(re.search(r"일자|날짜|기한|시작일|종료일|체결일|계약일", name))


def _is_number_column(name: str) -> bool:
    return bool(re.search(r"용량|금액|단가|비율|수량|면적|사용량|발전량|REC|MW|MWh|kW|kWh|개월|년수", name))


# 엑셀 날짜 일련번호의 기준일(1900년 날짜 체계, 1900을 윤년으로 잘못 셌던
# 옛 Lotus 1-2-3 버그를 엑셀이 그대로 물려받은 결과 - 1899-12-30을 0일로
# 두면 실제 사용 범위(1901년 이후)의 날짜는 전부 정확히 맞아떨어집니다.
_EXCEL_DATE_EPOCH = datetime.date(1899, 12, 30)


def _date_to_excel_serial(d: datetime.date) -> int:
    return (d - _EXCEL_DATE_EPOCH).days


def _coerce_for_excel(col_name: str, raw_value):
    value = "" if raw_value is None else str(raw_value).strip()
    if value == "":
        return ""

    if value.upper() in ("TRUE", "FALSE"):
        return value.upper() == "TRUE"

    if _is_date_column(col_name):
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
            try:
                # 주의: 여기서 datetime.datetime을 그대로 돌려주면 안 됩니다.
                # pywin32가 naive datetime을 COM VARIANT(DATE)로 바꿀 때 UTC
                # 기준으로 변환하는 경로를 타서, 한국(UTC+9)에서는 자정 값이
                # 전날 오후로 밀려 엑셀에 하루 전 날짜로 저장되는 문제가
                # 있었습니다(예: 2026-09-01 입력 → 2026-08-31 저장). 시간대
                # 변환이 아예 끼어들 수 없도록, 엑셀 날짜 일련번호(정수)를
                # 직접 계산해 순수 숫자로 씁니다 - _write_row가 이 숫자를 쓴
                # 뒤 셀 서식을 날짜로 지정합니다.
                parsed = datetime.datetime.strptime(value, fmt).date()
                return _date_to_excel_serial(parsed)
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

    def batch_apply(self, operations: list[dict]) -> dict:
        return self._call(self._batch_apply, operations)

    def workbook_reachable(self) -> bool:
        """사용자가 자기 엑셀에서 이 통합문서를 아직 열어두고 있는지 - 서버
        자동 종료 워치독(ppa_liveserver.py)이 주기적으로 불러 "엑셀을
        닫으면 서버도 같이 꺼지는" 동작을 구현하는 데 씁니다."""
        return self._call(self._workbook_reachable)

    def refresh_from_disk(self) -> dict:
        """지금 붙잡고 있는 워크북을 닫았다가 다시 엽니다 - "대시보드 새로고침"
        버튼에서 씁니다. read_all_tables()는 매번 이미 열려 있는 같은 워크북
        객체를 그대로 돌려받으므로(세션 내내 Excel이 메모리에 들고 있는 상태),
        다른 컴퓨터에서 저장한 변경사항이 여기 자동으로 반영되지 않습니다
        (Excel 자체의 공동편집 기능이 그 순간 정상 작동하지 않는 한 - 매크로
        포함 파일은 이마저도 불안정합니다). 닫았다 다시 열면 그 시점 디스크
        (OneDrive 동기화분)의 최신 내용을 확실히 새로 읽어옵니다."""
        return self._call(self._refresh_from_disk)

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
                # 예전에는 여기서 매 작업이 끝날 때마다 우리가 새로 띄운 숨김
                # 엑셀을 즉시 닫았습니다("아무도 안 쓰는데 읽기 전용" 방지 목적).
                # 문제는 엑셀이 사용자 화면에 안 열려 있는 흔한 상황(서버를 먼저
                # 켜고 나중에 엑셀을 여는 경우 등)에서는 조회 한 번, 저장 한 번마다
                # 매번 엑셀 프로세스를 새로 띄우고 통합문서를 다시 여는 셈이라
                # — COM으로 엑셀 프로세스를 새로 띄우는 것 자체가 수 초씩 걸릴 수
                # 있어 — "저장할 때마다 몇 초씩 멈칫"하는 게 체감 속도의 가장 큰
                # 원인이었습니다. 이제는 세션(서버가 떠 있는 동안) 내내 붙잡아두고
                # 서버가 완전히 종료될 때(아래 바깥쪽 finally, 콘솔 닫힘/Ctrl+C/
                # 원격 종료 요청 전부 이 경로를 탐)만 한 번 정리합니다 — 두 번째
                # 요청부터는 이미 열려 있는 인스턴스를 그대로 재사용해 훨씬
                # 빨라집니다.
        finally:
            self._close_if_we_launched()
            pythoncom.CoUninitialize()

    def _workbook_reachable(self) -> bool:
        # 우리가 스스로 띄운 숨김 인스턴스는 서버가 세션 내내 직접 붙잡고
        # 있으므로(위 _run 참고) 그 생사는 이 워치독이 신경 쓸 대상이 아닙니다
        # - "사용자가 자기 엑셀에서 이 통합문서를 닫았는지"만 감시합니다.
        if self._we_launched_app:
            return True
        # 아래 _ensure_workbook()과 정확히 같은 2단계 확인을 거칩니다 - 처음엔
        # 경로 모니커 확인 하나만 썼더니, OneDrive/SharePoint 자동 저장으로
        # 열린 통합문서(엑셀이 로컬 경로가 아니라 클라우드 URL 신원으로
        # 등록)에서 파일이 멀쩡히 열려 있는데도 "닫혔다"고 오판해 사용하는
        # 도중에 서버가 꺼져버리는 문제가 있었습니다(수정 버튼을 눌렀을 때
        # "Failed to fetch"로 나타남) - _ensure_workbook이 이미 이 문제를
        # 겪고 고쳐둔 방식을 그대로 재사용합니다.
        try:
            win32com.client.GetObject(self.xlsm_path)
            return True
        except Exception:
            pass
        try:
            running = win32com.client.GetObject(Class="Excel.Application")
            target_name = os.path.basename(self.xlsm_path).lower()
            for wb in running.Workbooks:
                try:
                    if os.path.basename(wb.Name).lower() == target_name:
                        return True
                except Exception:
                    continue
        except Exception:
            pass
        return False

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

    def _refresh_from_disk(self) -> dict:
        """[중요한 안전 제약] 사용자 자신의 화면에 열려 있는 Excel 세션은 여기서
        **절대 강제로 닫지 않습니다.** 원래는 어느 쪽이든(사용자 세션이든 우리가
        띄운 숨김 인스턴스든) 닫았다 다시 여는 방식으로 만들었었는데, 실제
        환경(부서 공유 OneDrive 파일을 여러 명이 각자 서버로 붙잡고 있는 상황)
        에서 서버가 멈추는 문제가 났습니다: Workbooks.Open()/Close() 같은 COM
        호출은 Excel이 어떤 대화상자(다른 버전 병합 안내, 읽기 전용 권장 등)를
        띄우면 사람이 그 창을 닫을 때까지 그대로 멈춰버리는데, 이 서버는 모든
        COM 호출을 스레드 하나에서 순서대로 처리하므로(ExcelBridge._run 참고)
        그 호출 하나가 멈추면 서버 전체가 응답을 멈춥니다("서버 연결 오류").
        사용자가 직접 보고 있는 세션은 하필 그런 대화상자가 뜨기 가장 쉬운
        바로 그 상황(다른 사람과 동시 편집)에서 새로고침되므로, 이 위험을
        아예 감수하지 않고 그 자리에서 이미 읽을 수 있는 값만 돌려줍니다.

        서버가 스스로 띄운 숨김 인스턴스(_we_launched_app)는 화면에 아무도
        안 보고 있고 우리만 붙잡고 있는 세션이라 위 위험이 훨씬 낮으므로,
        그 경우에만 실제로 닫았다 다시 엽니다."""
        try:
            wb = self._ensure_workbook()
        except ExcelComError:
            # 애초에 열려 있지 않으면 다음 읽기가 어차피 새로 여니, 새로고침할
            # 대상이 없는 것과 같습니다 - 조용히 넘어갑니다.
            return {"reopened": False}

        if not self._we_launched_app:
            return {
                "reopened": False,
                "note": (
                    "지금 사용자 화면에 열려 있는 엑셀 세션에서 읽었습니다. 다른 "
                    "컴퓨터에서 저장한 내용까지 확실히 보려면 엑셀 파일을 직접 "
                    "닫았다가 다시 열어주세요."
                ),
            }

        try:
            saved = bool(wb.Saved)
        except Exception:
            saved = True

        if not saved:
            raise ExcelComError(
                "지금 저장되지 않은 변경사항이 있어 새로고침(다시 열기)을 "
                "진행할 수 없습니다. 잠시 후 다시 시도해주세요."
            )

        app = wb.Application
        try:
            app.DisplayAlerts = False
        except Exception:
            pass

        try:
            wb.Close(SaveChanges=False)
        except Exception as exc:
            raise ExcelComError(
                f"기존 워크북을 닫는 중 오류가 났습니다: {exc}"
            ) from exc

        try:
            app.Workbooks.Open(self.xlsm_path)
        except Exception as exc:
            raise ExcelComError(f"{self._diagnose_open_failure()}\n\n원본 오류: {exc}") from exc

        self._app = app
        return {"reopened": True}

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
    def _validate(self, table_key: str, record: dict, extra_valid_fk: dict | None = None) -> list[str]:
        """extra_valid_fk: {표이름: {아직 저장 전이지만 이번 배치에서 같이
        생기는 PK, ...}} - 그룹 일괄 입력에서 "새 부모 + 그 부모를 참조하는
        새 자식"을 한 배치로 같이 만들 때, 부모가 아직 엑셀에 없어도 FK
        검증을 통과시키기 위해 씁니다."""
        extra_valid_fk = extra_valid_fk or {}
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
            ref_values |= extra_valid_fk.get(ref_key, set())
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

        wb = self._ensure_workbook()
        result = self._write_row(wb, table_key, record)
        wb.Save()
        return result

    def _write_row(self, wb, table_key: str, record: dict) -> dict:
        """검증은 이미 끝났다고 보고 실제로 셀에 씁니다(저장은 호출자 책임 -
        배치 작업은 마지막에 한 번만 저장해야 해서 여기서 매번 저장하지 않음)."""
        schema = TABLE_BY_KEY[table_key]
        pk_value = str(record.get(schema.pk) or "").strip()
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
            coerced = _coerce_for_excel(col, record.get(col, ""))
            cell = ws.Cells(target_row, abs_col)
            cell.Value = coerced
            # 일련번호(정수)로 썼으니, 셀에 이미 날짜 서식이 없었던 경우에도
            # 숫자가 아니라 날짜로 보이도록 명시적으로 지정합니다(기존에
            # 데이터가 있던 열이면 보통 이미 날짜 서식이라 그대로여도 무방하나,
            # 새로 추가되는 행/시트라면 이 지정이 없으면 46266 같은 숫자로
            # 보일 수 있습니다).
            if _is_date_column(col) and isinstance(coerced, int):
                cell.NumberFormat = "yyyy-mm-dd"

        return {"action": action, "table": table_key, "pk_value": pk_value, "row": target_row}

    # ---- 참조 확인 (삭제 전에 다른 표가 이 PK를 쓰고 있는지) ----
    def _referencing_rows(self, table_key: str, pk_value: str) -> list[dict]:
        """참조하는 개별 행까지 돌려줍니다(표별 집계가 아니라) - 배치 삭제에서
        "이번 배치에서 같이 지워지는 행"은 걸림돌에서 빼기 위해 필요합니다."""
        target = str(pk_value or "").strip()
        out: list[dict] = []

        for t in TABLES:
            for fk_col, ref_key in t.fk.items():
                if ref_key != table_key:
                    continue
                for r in self._read_table(t.key)["rows"]:
                    if str(r.get(fk_col) or "").strip() == target:
                        out.append({"table": t.key, "fk_col": fk_col, "pk": str(r.get(t.pk) or "").strip()})

        return out

    def _referencing_records(
        self, table_key: str, pk_value: str, exclude: set | None = None
    ) -> list[dict]:
        exclude = exclude or set()
        rows = [r for r in self._referencing_rows(table_key, pk_value) if (r["table"], r["pk"]) not in exclude]

        agg: dict[tuple[str, str], int] = {}
        for r in rows:
            key = (r["table"], r["fk_col"])
            agg[key] = agg.get(key, 0) + 1

        return [
            {"table": tk, "label": TABLE_BY_KEY[tk].label, "fk_col": fc, "count": c}
            for (tk, fc), c in agg.items()
        ]

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
        result = self._delete_row(wb, table_key, pk_value)
        wb.Save()
        return result

    def _delete_row(self, wb, table_key: str, pk_value: str) -> dict:
        """참조 검사는 이미 끝났다고 보고 실제로 행을 지웁니다(저장은 호출자
        책임). 표(ListObject) 안에 있는 행이든 아니든, 행 전체 삭제는 엑셀이
        알아서 표 범위/서식을 줄이고 나머지 행을 한 칸씩 끌어올립니다 - 별도의
        ListObject 전용 삭제 API가 필요 없습니다. 이 시트들은 검증/파생값을
        전부 이 시스템이 계산하고 엑셀 수식을 쓰지 않으므로(스키마 설계상),
        셀 참조가 깨질 걱정도 없습니다."""
        schema = TABLE_BY_KEY[table_key]
        ws = self._worksheet(wb, table_key)
        values, start_row, start_col, col_at, target_row = self._locate_row(ws, schema, pk_value)

        if target_row is None:
            raise ExcelComError(
                f"{schema.pk} '{pk_value}'을(를) {table_key}에서 찾지 못했습니다(이미 삭제됐을 수 있습니다)."
            )

        ws.Rows(target_row).Delete()
        return {"action": "deleted", "table": table_key, "pk_value": pk_value}

    # ---- 일괄(배치) 작업: 여러 표에 걸친 저장/삭제를 한 번에 검증 → 전부 적용
    # → 마지막에 한 번만 저장. 검증에서 하나라도 실패하면 아무것도 쓰지
    # 않습니다("트랜잭션"에 가장 가까운, COM으로 실제 구현 가능한 형태). ----
    def _batch_apply(self, operations: list[dict]) -> dict:
        if not operations:
            raise ExcelComError("적용할 작업이 없습니다.")

        scheduled_deletes = {
            (op.get("table"), str(op.get("pk") or "").strip())
            for op in operations
            if op.get("action") == "delete"
        }

        # 배치 안에서 "이번에 새로 같이 생기는" PK들 - 아직 저장 전이라도
        # 부모+자식을 한 번에 새로 만들 때 FK 검증을 통과시키기 위함.
        batch_new_pks: dict[str, set] = {}
        for op in operations:
            if op.get("action") != "save":
                continue
            table = op.get("table")
            schema = TABLE_BY_KEY.get(table)
            if not schema:
                continue
            pk_value = str((op.get("record") or {}).get(schema.pk) or "").strip()
            if pk_value:
                batch_new_pks.setdefault(table, set()).add(pk_value)

        errors: list[str] = []
        for idx, op in enumerate(operations, start=1):
            table = op.get("table")
            action = op.get("action")

            if table not in TABLE_BY_KEY:
                errors.append(f"{idx}번째 작업: 지원하지 않는 표입니다 ({table}).")
                continue

            if action == "save":
                errs = self._validate(table, op.get("record") or {}, extra_valid_fk=batch_new_pks)
                if errs:
                    errors.append(f"{idx}번째 작업({TABLE_BY_KEY[table].label}): " + " / ".join(errs))
            elif action == "delete":
                pk_value = str(op.get("pk") or "").strip()
                if not pk_value:
                    errors.append(f"{idx}번째 작업({TABLE_BY_KEY[table].label}): 삭제할 PK가 없습니다.")
                    continue
                refs = self._referencing_records(table, pk_value, exclude=scheduled_deletes)
                if refs:
                    detail = ", ".join(f"{r['label']} {r['count']}건({r['fk_col']})" for r in refs)
                    errors.append(
                        f"{idx}번째 작업({TABLE_BY_KEY[table].label} {pk_value}): "
                        f"이 배치 밖에서 참조하고 있어 삭제할 수 없습니다 - {detail}."
                    )
            else:
                errors.append(f"{idx}번째 작업: 알 수 없는 action '{action}' 입니다.")

        if errors:
            raise ExcelComError(
                "일괄 작업을 적용하지 못했습니다(하나도 반영되지 않았습니다):\n" + "\n".join(errors)
            )

        wb = self._ensure_workbook()
        results = []
        for op in operations:
            if op.get("action") == "save":
                results.append(self._write_row(wb, op["table"], op.get("record") or {}))
            else:
                results.append(self._delete_row(wb, op["table"], str(op.get("pk") or "").strip()))

        wb.Save()
        return {"results": results}

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
