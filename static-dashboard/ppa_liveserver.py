#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ppa_liveserver.py — 대시보드 화면에서 입력/저장하면 엑셀에도 그대로 반영합니다.

기존 dashboard_recreate.bat(더블클릭 → HTML 한 번 생성 → 끝) 과 달리, 이
스크립트는 계속 떠 있는 작은 웹 서버입니다. 브라우저에서 열어두고 쓰는
동안:

  1) 화면 오른쪽 아래 "간편 입력/저장" 버튼으로 표를 고르고 값을 입력해
     저장하면, excel_com.py 가 Windows COM으로 **엑셀 프로그램 자체**를
     조작해서 그 값을 실제 통합문서에 씁니다(파일을 직접 열어 덮어쓰는 게
     아니라, 지금 열려 있으면 그 세션 그대로, 안 열려 있으면 새로 열어서).
  2) 저장이 끝나면 서버가 방금 바뀐 내용을 다시 읽어 PPA현황.html 을 그
     자리에서 다시 만들고, 브라우저 화면도 새로고침됩니다.

즉 dashboard_recreate.bat(openpyxl로 파일을 직접 읽음, 서버 없음, 조회 전용)
과 이 스크립트(COM으로 살아있는 엑셀을 직접 조작, 서버 필요, 입력 가능)는
서로 다른 상황을 위한 것입니다 — 각각의 장단점은 README.md 참고.

이 스크립트는 Windows + Microsoft Excel + pywin32 가 있어야 동작합니다.
"""
from __future__ import annotations

import argparse
import datetime
import json
import os
import sys
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from excel_com import ExcelBridge, ExcelComError  # noqa: E402
from ppa_schema import TABLE_BY_KEY, TABLES  # noqa: E402


def find_first_xlsm(base_dir: Path) -> Path:
    files = sorted(base_dir.glob("*.xlsm"))
    if not files:
        raise FileNotFoundError(f"이 폴더에서 xlsm 파일을 찾지 못했습니다: {base_dir}")
    if len(files) > 1:
        raise FileNotFoundError(
            f"이 폴더에 xlsm 파일이 여러 개 있어 자동으로 고를 수 없습니다: {base_dir}\n"
            "--xlsm 으로 직접 지정해주세요."
        )
    return files[0]


class App:
    def __init__(self, xlsm_path: str, html_path: str, form_js_path: str):
        self.xlsm_path = os.path.abspath(xlsm_path)
        self.html_path = os.path.abspath(html_path)
        self.form_js_path = os.path.abspath(form_js_path)
        self.bridge = ExcelBridge(self.xlsm_path)
        self._rebuild_lock = threading.Lock()
        self._last_payload: dict | None = None

    def schema_json(self) -> list[dict]:
        return [
            {
                "key": t.key,
                "label": t.label,
                "pk": t.pk,
                "columns": t.columns,
                "fk": t.fk,
            }
            for t in TABLES
        ]

    def rebuild(self) -> dict:
        """엑셀에서 6개 표를 다시 읽어 PPA현황.html 을 갱신합니다.

        [변경] 탭에 보이는 "기준점 대비 누적 변경"은 리셋 버튼(reset_baseline)을
        누르기 전까지 고정된 기준 스냅샷과 비교합니다 - 그 사이에는 저장할
        때마다 계속 쌓여서 보입니다. 전체 변경 이력(changelog)에는 그 커진
        누적 diff를 그대로 매번 다시 넣지 않도록, 이번 한 번의 저장에서 실제로
        바뀐 것만 별도(lastbuild) 스냅샷으로 계산해 넣습니다.
        """
        from build_dashboard import build_payload, default_snapshot_path
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

        with self._rebuild_lock:
            tables_data = self.bridge.read_all_tables()

            baseline_path = default_snapshot_path(self.html_path)
            baseline = load_snapshot(baseline_path)
            changes, marks = compute_changes(tables_data, baseline)

            lastbuild_path = default_lastbuild_path(self.html_path)
            lastbuild = load_snapshot(lastbuild_path)
            build_changes, build_marks = compute_changes(tables_data, lastbuild)

            changelog_path = default_changelog_path(self.html_path)
            changelog = load_changelog(changelog_path)
            generated_at = datetime.datetime.now().isoformat(timespec="seconds")
            if build_changes.get("has_prev"):
                new_entries = build_changelog_entries(tables_data, build_changes, build_marks, generated_at)
                if new_entries:
                    changelog = append_changelog(changelog_path, new_entries)

            payload = build_payload(tables_data, is_demo=False, changes=changes, marks=marks, changelog=changelog)
            payload["generated_at"] = generated_at

            with open(self.html_path, "w", encoding="utf-8") as f:
                f.write(render_dashboard(payload))
            save_snapshot(lastbuild_path, tables_data, payload["generated_at"])
            if baseline is None:
                save_snapshot(baseline_path, tables_data, payload["generated_at"])

            self._last_payload = payload
            return payload

    def reset_baseline(self) -> dict:
        """[변경] 탭의 비교 기준점을 지금 시점으로 되돌립니다 - 그동안 쌓여
        보이던 "기준점 대비 누적 변경"이 초기화됩니다. 여러 생성에 걸쳐 계속
        남는 "전체 변경 이력"(changelog)은 건드리지 않습니다."""
        from build_dashboard import default_snapshot_path
        from ppa_changes import save_snapshot

        with self._rebuild_lock:
            tables_data = self.bridge.read_all_tables()
            generated_at = datetime.datetime.now().isoformat(timespec="seconds")
            save_snapshot(default_snapshot_path(self.html_path), tables_data, generated_at)
        return self.rebuild()

    def dashboard_html(self) -> str:
        if not os.path.exists(self.html_path):
            self.rebuild()
        html = Path(self.html_path).read_text(encoding="utf-8", errors="replace")
        tag = f'<script src="/dashboard_form.js?v={int(time.time())}"></script>'
        idx = html.lower().rfind("</body>")
        if idx >= 0:
            return html[:idx] + tag + "\n" + html[idx:]
        return html + "\n" + tag

    def save_and_rebuild(self, table_key: str, record: dict) -> dict:
        result = self.bridge.save_record(table_key, record)
        payload = self.rebuild()
        return {
            "save": result,
            "validation_errors": payload["validation"]["total_errors"],
        }

    def delete_and_rebuild(self, table_key: str, pk_value: str) -> dict:
        result = self.bridge.delete_record(table_key, pk_value)
        payload = self.rebuild()
        return {
            "delete": result,
            "validation_errors": payload["validation"]["total_errors"],
        }

    def batch_and_rebuild(self, operations: list[dict]) -> dict:
        result = self.bridge.batch_apply(operations)
        payload = self.rebuild()
        return {
            "batch": result,
            "validation_errors": payload["validation"]["total_errors"],
        }


def make_handler(app: App):
    class Handler(BaseHTTPRequestHandler):
        def _send(self, code=200, content_type="text/plain; charset=utf-8", body=b""):
            try:
                self.send_response(code)
                self.send_header("Content-Type", content_type)
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except (BrokenPipeError, ConnectionAbortedError, ConnectionResetError):
                pass

        def _send_json(self, obj: dict, code: int = 200):
            self._send(code, "application/json; charset=utf-8", json.dumps(obj, ensure_ascii=False).encode("utf-8"))

        def do_GET(self):
            parsed = urlparse(self.path)
            path, qs = parsed.path, parse_qs(parsed.query)

            if path in ("/", "/index.html"):
                try:
                    self._send(200, "text/html; charset=utf-8", app.dashboard_html().encode("utf-8"))
                except Exception as e:
                    self._send(500, "text/plain; charset=utf-8", str(e).encode("utf-8"))
                return

            if path == "/dashboard_form.js":
                try:
                    self._send(200, "application/javascript; charset=utf-8", Path(app.form_js_path).read_bytes())
                except Exception as e:
                    self._send_json({"ok": False, "error": str(e)})
                return

            if path == "/api/schema":
                self._send_json({"ok": True, "schema": app.schema_json()})
                return

            if path == "/api/options":
                table = (qs.get("table") or [""])[0]
                if table not in TABLE_BY_KEY:
                    self._send_json({"ok": False, "error": "지원하지 않는 표입니다."})
                    return
                try:
                    self._send_json({"ok": True, "options": app.bridge.get_options(table)})
                except ExcelComError as e:
                    self._send_json({"ok": False, "error": str(e)})
                except Exception as e:
                    self._send_json({"ok": False, "error": f"예상하지 못한 오류: {e}"})
                return

            if path == "/api/records":
                table = (qs.get("table") or [""])[0]
                if table not in TABLE_BY_KEY:
                    self._send_json({"ok": False, "error": "지원하지 않는 표입니다."})
                    return
                try:
                    self._send_json({"ok": True, "rows": app.bridge.read_table(table)["rows"]})
                except ExcelComError as e:
                    self._send_json({"ok": False, "error": str(e)})
                except Exception as e:
                    self._send_json({"ok": False, "error": f"예상하지 못한 오류: {e}"})
                return

            if path == "/api/record":
                table = (qs.get("table") or [""])[0]
                pk = (qs.get("pk") or [""])[0]
                if table not in TABLE_BY_KEY:
                    self._send_json({"ok": False, "error": "지원하지 않는 표입니다."})
                    return
                try:
                    self._send_json({"ok": True, "record": app.bridge.get_record(table, pk)})
                except ExcelComError as e:
                    self._send_json({"ok": False, "error": str(e)})
                except Exception as e:
                    self._send_json({"ok": False, "error": f"예상하지 못한 오류: {e}"})
                return

            if path == "/api/references":
                table = (qs.get("table") or [""])[0]
                pk = (qs.get("pk") or [""])[0]
                if table not in TABLE_BY_KEY:
                    self._send_json({"ok": False, "error": "지원하지 않는 표입니다."})
                    return
                if not pk:
                    self._send_json({"ok": True, "references": []})
                    return
                try:
                    self._send_json({"ok": True, "references": app.bridge.get_references(table, pk)})
                except ExcelComError as e:
                    self._send_json({"ok": False, "error": str(e)})
                except Exception as e:
                    self._send_json({"ok": False, "error": f"예상하지 못한 오류: {e}"})
                return

            self._send(404, "text/plain; charset=utf-8", b"Not Found")

        def do_POST(self):
            path = urlparse(self.path).path
            try:
                length = int(self.headers.get("Content-Length", "0"))
                raw = self.rfile.read(length)
                payload = json.loads(raw.decode("utf-8")) if raw else {}
            except Exception as e:
                self._send_json({"ok": False, "error": f"잘못된 요청 본문: {e}"})
                return

            if path == "/api/save":
                table = str(payload.get("table") or "").strip()
                record = payload.get("record") or {}
                if table not in TABLE_BY_KEY:
                    self._send_json({"ok": False, "error": "지원하지 않는 표입니다."})
                    return
                try:
                    result = app.save_and_rebuild(table, record)
                    self._send_json({"ok": True, **result, "message": "엑셀에 저장하고 대시보드를 갱신했습니다."})
                except ExcelComError as e:
                    self._send_json({"ok": False, "error": str(e)})
                except Exception as e:
                    self._send_json({"ok": False, "error": f"예상하지 못한 오류: {e}"})
                return

            if path == "/api/delete":
                table = str(payload.get("table") or "").strip()
                pk = str(payload.get("pk") or "").strip()
                if table not in TABLE_BY_KEY:
                    self._send_json({"ok": False, "error": "지원하지 않는 표입니다."})
                    return
                if not pk:
                    self._send_json({"ok": False, "error": "삭제할 PK를 지정해주세요."})
                    return
                try:
                    result = app.delete_and_rebuild(table, pk)
                    self._send_json({"ok": True, **result, "message": "엑셀에서 삭제하고 대시보드를 갱신했습니다."})
                except ExcelComError as e:
                    self._send_json({"ok": False, "error": str(e)})
                except Exception as e:
                    self._send_json({"ok": False, "error": f"예상하지 못한 오류: {e}"})
                return

            if path == "/api/batch":
                operations = payload.get("operations") or []
                if not isinstance(operations, list) or not operations:
                    self._send_json({"ok": False, "error": "적용할 작업이 없습니다."})
                    return
                bad = [op for op in operations if str(op.get("table") or "") not in TABLE_BY_KEY]
                if bad:
                    self._send_json({"ok": False, "error": "지원하지 않는 표가 포함돼 있습니다."})
                    return
                try:
                    result = app.batch_and_rebuild(operations)
                    self._send_json({"ok": True, **result, "message": "일괄 작업을 엑셀에 반영하고 대시보드를 갱신했습니다."})
                except ExcelComError as e:
                    self._send_json({"ok": False, "error": str(e)})
                except Exception as e:
                    self._send_json({"ok": False, "error": f"예상하지 못한 오류: {e}"})
                return

            if path == "/api/rebuild":
                try:
                    payload = app.rebuild()
                    self._send_json({"ok": True, "message": "대시보드를 갱신했습니다.", "generated_at": payload["generated_at"]})
                except Exception as e:
                    self._send_json({"ok": False, "error": str(e)})
                return

            if path == "/api/reset_snapshot":
                try:
                    payload = app.reset_baseline()
                    self._send_json({
                        "ok": True,
                        "message": "변경 비교 기준점을 지금 시점으로 리셋했습니다.",
                        "generated_at": payload["generated_at"],
                    })
                except Exception as e:
                    self._send_json({"ok": False, "error": str(e)})
                return

            self._send_json({"ok": False, "error": "Not Found"}, 404)

        def log_message(self, fmt, *args):
            print("[HTTP]", fmt % args)

    return Handler


def _install_console_close_handler(app: App) -> None:
    """콘솔 창을 X로 닫거나 로그오프/종료할 때도 정리가 되도록 합니다.

    Ctrl+C(KeyboardInterrupt)는 이미 main()에서 처리되지만, 콘솔 창의 X 버튼은
    파이썬 예외가 아니라 Windows 콘솔 제어 이벤트라서 그걸로는 안 잡힙니다.
    이걸 놓치면, 우리가 새로 띄운 숨김 엑셀 인스턴스(파일이 아무 데도 안
    열려 있어서 어쩔 수 없이 띄운 경우)가 프로세스와 함께 뒤에 그대로 남아
    파일을 계속 붙잡고 있을 수 있습니다 - "아무도 안 쓰는데 읽기 전용"의
    흔한 원인.
    """
    try:
        import win32api
        import win32con

        def handler(ctrl_type):
            if ctrl_type in (
                win32con.CTRL_CLOSE_EVENT,
                win32con.CTRL_LOGOFF_EVENT,
                win32con.CTRL_SHUTDOWN_EVENT,
            ):
                print("[종료] 콘솔이 닫혀 엑셀 연결을 정리합니다...")
                app.bridge.shutdown()
                app.bridge._thread.join(timeout=3)
                return True
            return False

        win32api.SetConsoleCtrlHandler(handler, True)
    except Exception:
        pass


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--xlsm", default="", help="원본 xlsm 파일 경로 (생략하면 상위 폴더에서 자동 탐색)")
    parser.add_argument("--out", default="PPA현황.html", help="대시보드 HTML 출력 경로 (기본: PPA현황.html)")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8842)
    parser.add_argument("--no-browser", action="store_true", help="시작할 때 브라우저를 자동으로 열지 않음")
    args = parser.parse_args()

    script_dir = Path(__file__).resolve().parent

    if args.xlsm:
        xlsm_path = str(Path(args.xlsm).resolve())
    else:
        try:
            xlsm_path = str(find_first_xlsm(script_dir.parent))
        except FileNotFoundError as e:
            sys.exit(str(e))

    if not os.path.exists(xlsm_path):
        sys.exit(f"엑셀 파일을 찾을 수 없습니다: {xlsm_path}")

    html_path = str((script_dir / args.out).resolve())
    form_js_path = str(script_dir / "dashboard_form.js")

    print(f"[정보] XLSM: {xlsm_path}")
    print(f"[정보] HTML: {html_path}")

    try:
        app = App(xlsm_path=xlsm_path, html_path=html_path, form_js_path=form_js_path)
    except ExcelComError as e:
        sys.exit(str(e))

    try:
        app.rebuild()
        print("[정보] 초기 대시보드 생성 완료")
    except Exception as e:
        print(f"[경고] 초기 대시보드 생성 실패 (서버는 계속 시작합니다): {e}")

    _install_console_close_handler(app)

    server = ThreadingHTTPServer((args.host, args.port), make_handler(app))
    url = f"http://{args.host}:{args.port}"
    print(f"[OK] 서버 시작: {url}")
    print("[안내] 이 창을 닫으면 서버가 종료됩니다. 끝내려면 Ctrl+C 를 누르세요.")

    if not args.no_browser:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n[종료] 서버를 멈춥니다.")
    finally:
        app.bridge.shutdown()
        server.server_close()


if __name__ == "__main__":
    main()
