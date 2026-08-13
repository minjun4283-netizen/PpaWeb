#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL)

NS = {"a": NS_MAIN, "r": NS_REL, "pr": NS_PKG_REL}


def q(tag: str) -> str:
    return f"{{{NS_MAIN}}}{tag}"


def normalize_name(text: str) -> str:
    return str(text or "").replace(" ", "").replace("_", "").replace("-", "").strip().lower()


def cell_to_text(v) -> str:
    if v is None:
        return ""
    return str(v).strip()


def col_letters_to_index(s: str) -> int:
    n = 0
    for ch in s.upper():
        if "A" <= ch <= "Z":
            n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def col_index_to_letters(n: int) -> str:
    out = []
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out.append(chr(ord("A") + rem))
    return "".join(reversed(out)) or "A"


def split_ref(ref: str):
    m = re.match(r"^([A-Z]+)(\d+)$", str(ref or "").upper())
    if not m:
        return None, None
    return col_letters_to_index(m.group(1)), int(m.group(2))


def resolve_sheet_target(target: str) -> str:
    t = str(target or "").replace("\\", "/")
    if t.startswith("/"):
        t = t.lstrip("/")
    while t.startswith("../"):
        t = t[3:]
    if not t.startswith("xl/"):
        t = "xl/" + t
    return t


def extract_xlsm_to_temp(xlsm_path: str):
    temp_dir = tempfile.mkdtemp(prefix="ppa_xlsm_xml_")
    with zipfile.ZipFile(xlsm_path, "r") as zf:
        zf.extractall(temp_dir)
    return temp_dir


def pack_temp_to_xlsm(temp_dir: str, out_path: str):
    with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(temp_dir):
            for name in files:
                full = os.path.join(root, name)
                rel = os.path.relpath(full, temp_dir).replace("\\", "/")
                zf.write(full, rel)


def load_shared_strings(temp_dir: str):
    path = os.path.join(temp_dir, "xl", "sharedStrings.xml")
    if not os.path.exists(path):
        return []

    tree = ET.parse(path)
    root = tree.getroot()
    out = []

    for si in root.findall("a:si", NS):
        out.append("".join(si.itertext()))

    return out


def get_sheet_part(temp_dir: str, table_name: str) -> str:
    workbook_xml = os.path.join(temp_dir, "xl", "workbook.xml")
    rels_xml = os.path.join(temp_dir, "xl", "_rels", "workbook.xml.rels")

    wb_tree = ET.parse(workbook_xml)
    wb_root = wb_tree.getroot()

    rel_tree = ET.parse(rels_xml)
    rel_root = rel_tree.getroot()

    rel_map = {}
    for rel in rel_root.findall(f"{{{NS_PKG_REL}}}Relationship"):
        rel_map[rel.get("Id")] = rel.get("Target")

    sheets = []
    for s in wb_root.findall("a:sheets/a:sheet", NS):
        name = s.get("name") or ""
        rid = s.get(f"{{{NS_REL}}}id")
        target = rel_map.get(rid, "")
        sheets.append((name, resolve_sheet_target(target)))

    for name, target in sheets:
        if name == table_name:
            return target

    target_norm = normalize_name(table_name)
    target_no_t = normalize_name(table_name.replace("T_", ""))

    for name, target in sheets:
        n = normalize_name(name)
        if n == target_norm or n == target_no_t:
            return target

    for name, target in sheets:
        n = normalize_name(name)
        if target_no_t and (target_no_t in n or n in target_no_t):
            return target

    raise RuntimeError(f"Worksheet not found: {table_name}")


def cell_elem_to_text(cell, shared_strings) -> str:
    t = cell.get("t")
    if t == "inlineStr":
        is_elem = cell.find("a:is", NS)
        return "".join(is_elem.itertext()).strip() if is_elem is not None else ""

    v = cell.find("a:v", NS)
    if v is None or v.text is None:
        return ""

    if t == "s":
        try:
            idx = int(v.text)
            return shared_strings[idx] if 0 <= idx < len(shared_strings) else ""
        except Exception:
            return ""

    return str(v.text).strip()


def parse_sheet(temp_dir: str, sheet_part: str):
    sheet_xml = os.path.join(temp_dir, *sheet_part.split("/"))
    tree = ET.parse(sheet_xml)
    root = tree.getroot()
    sheet_data = root.find("a:sheetData", NS)
    if sheet_data is None:
        raise RuntimeError("sheetData not found")

    shared_strings = load_shared_strings(temp_dir)

    rows = {}
    max_row = 1
    max_col = 1

    for row_elem in sheet_data.findall("a:row", NS):
        row_idx = int(row_elem.get("r") or "0")
        if row_idx <= 0:
            continue

        row_map = {}
        for cell in row_elem.findall("a:c", NS):
            ref = cell.get("r") or ""
            col_idx, row_from_ref = split_ref(ref)
            if not col_idx:
                continue
            text = cell_elem_to_text(cell, shared_strings)
            row_map[col_idx] = text
            max_col = max(max_col, col_idx)
            if row_from_ref:
                max_row = max(max_row, row_from_ref)

        rows[row_idx] = row_map
        max_row = max(max_row, row_idx)

    return tree, root, sheet_data, rows, max_row, max_col, sheet_xml


def detect_header_row(rows: dict, max_col: int):
    best_row = 1
    best_headers = []

    for row_idx in range(1, 21):
        seen = set()
        cur = []
        row_map = rows.get(row_idx, {})
        for col_idx in range(1, max_col + 1):
            name = str(row_map.get(col_idx, "")).strip()
            if not name or name in seen:
                continue
            seen.add(name)
            cur.append((col_idx, name))

        if len(cur) > len(best_headers):
            best_row = row_idx
            best_headers = cur

    return best_row, best_headers


def extract_table_rows(rows: dict, max_row: int, header_row: int, header_pairs):
    headers = [name for _, name in header_pairs]
    out = []

    for row_idx in range(header_row + 1, max_row + 1):
        row_map = rows.get(row_idx, {})
        obj = {}
        non_empty = 0

        for col_idx, name in header_pairs:
            text = str(row_map.get(col_idx, "")).strip()
            if text:
                non_empty += 1
            obj[name] = text

        if non_empty > 0:
            out.append(obj)

    return headers, out


def read_table(xlsm_path: str, table_name: str):
    temp_dir = None
    try:
        temp_dir = extract_xlsm_to_temp(xlsm_path)
        sheet_part = get_sheet_part(temp_dir, table_name)
        _, _, _, rows, max_row, max_col, _ = parse_sheet(temp_dir, sheet_part)
        header_row, header_pairs = detect_header_row(rows, max_col)
        headers, out_rows = extract_table_rows(rows, max_row, header_row, header_pairs)
        return {"headers": headers, "rows": out_rows}
    finally:
        if temp_dir and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)


def find_row_elem(sheet_data, target_row: int):
    for row_elem in sheet_data.findall("a:row", NS):
        if int(row_elem.get("r") or "0") == target_row:
            return row_elem
    return None


def ensure_row_elem(sheet_data, target_row: int):
    row_elem = find_row_elem(sheet_data, target_row)
    if row_elem is not None:
        return row_elem

    new_row = ET.Element(q("row"), {"r": str(target_row)})

    inserted = False
    row_elems = sheet_data.findall("a:row", NS)
    for idx, existing in enumerate(row_elems):
        r = int(existing.get("r") or "0")
        if r > target_row:
            sheet_data.insert(idx, new_row)
            inserted = True
            break

    if not inserted:
        sheet_data.append(new_row)

    return new_row


def cell_ref(col_idx: int, row_idx: int) -> str:
    return f"{col_index_to_letters(col_idx)}{row_idx}"


def ensure_cell_elem(row_elem, col_idx: int, row_idx: int):
    ref = cell_ref(col_idx, row_idx)

    for cell in row_elem.findall("a:c", NS):
        if (cell.get("r") or "").upper() == ref.upper():
            return cell

    new_cell = ET.Element(q("c"), {"r": ref})

    inserted = False
    cells = row_elem.findall("a:c", NS)
    for idx, existing in enumerate(cells):
        existing_ref = existing.get("r") or ""
        existing_col, _ = split_ref(existing_ref)
        if existing_col and existing_col > col_idx:
            row_elem.insert(idx, new_cell)
            inserted = True
            break

    if not inserted:
        row_elem.append(new_cell)

    return new_cell


def set_cell_inline_text(cell, value: str):
    for child in list(cell):
        cell.remove(child)

    cell.set("t", "inlineStr")
    is_elem = ET.SubElement(cell, q("is"))
    t_elem = ET.SubElement(is_elem, q("t"))
    t_elem.text = "" if value is None else str(value)


def update_dimension(root, max_row: int, max_col: int):
    dim = root.find("a:dimension", NS)
    if dim is None:
        return
    dim.set("ref", f"A1:{col_index_to_letters(max_col)}{max_row}")


def write_table(xlsm_path: str, table_name: str, pk_name: str, record: dict):
    pk_value = str(record.get(pk_name) or "").strip()
    if not pk_value:
        raise RuntimeError(f"{pk_name}는 필수입니다.")

    temp_dir = None
    temp_xlsm = None

    try:
        temp_dir = extract_xlsm_to_temp(xlsm_path)
        sheet_part = get_sheet_part(temp_dir, table_name)
        tree, root, sheet_data, rows, max_row, max_col, sheet_xml = parse_sheet(temp_dir, sheet_part)

        header_row, header_pairs = detect_header_row(rows, max_col)
        header_map = {name: col_idx for col_idx, name in header_pairs}

        if pk_name not in header_map:
            raise RuntimeError(f"PK column not found: {pk_name}")

        pk_col = header_map[pk_name]
        target_row = None

        for row_idx in range(header_row + 1, max_row + 1):
            row_map = rows.get(row_idx, {})
            if str(row_map.get(pk_col, "")).strip() == pk_value:
                target_row = row_idx
                break

        action = "updated"
        if target_row is None:
            target_row = max(max_row, header_row) + 1
            action = "inserted"

        row_elem = ensure_row_elem(sheet_data, target_row)

        for key, value in record.items():
            if key in header_map:
                col_idx = header_map[key]
                cell = ensure_cell_elem(row_elem, col_idx, target_row)
                set_cell_inline_text(cell, "" if value is None else str(value))
                max_col = max(max_col, col_idx)

        max_row = max(max_row, target_row)
        update_dimension(root, max_row, max_col)

        tree.write(sheet_xml, encoding="utf-8", xml_declaration=True)

        fd, temp_xlsm = tempfile.mkstemp(prefix="ppa_saved_", suffix=".xlsm")
        os.close(fd)
        pack_temp_to_xlsm(temp_dir, temp_xlsm)

        shutil.copy2(temp_xlsm, xlsm_path)

        return {
            "action": action,
            "row": target_row,
            "pk_name": pk_name,
            "pk_value": pk_value,
            "table": table_name,
        }

    except PermissionError:
        raise RuntimeError("원본 XLSM 파일이 열려 있습니다. Excel에서 닫고 다시 시도하세요.")
    finally:
        if temp_xlsm and os.path.exists(temp_xlsm):
            try:
                os.remove(temp_xlsm)
            except Exception:
                pass
        if temp_dir and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)
2) static-dashboard\ppa_server.py 전체 교체
python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from xlsm_store import read_table, write_table


STATIC_SCHEMA = {
    "T_발전소": ["발전소ID", "발전소명", "발전법인명", "설비용량(MW)", "발전원", "Readiness", "MGA_Supply"],
    "T_구매계약": ["구매계약ID", "발전소ID", "구매계약용량(MW)", "구매단가(원/kWh)", "공급기한_구매", "계약기간(년)", "수요기업 미확보", "구매 담당자"],
    "T_수요기업": ["수요기업ID", "기업명"],
    "T_판매계약": ["판매계약ID", "수요기업ID", "판매계약용량(MW)", "계약일", "공급기한_판매", "계약유형", "판매단가(원/kWh)", "공급자원 미확보", "판매 담당자", "계약기간(년)", "Requirement", "MGA_Demand"],
    "T_전기사용지": ["전기사용지ID", "판매계약ID", "전기사용지명", "전기사용지계약용량(MW)"],
    "T_수급매칭": ["수급매칭ID", "전기사용지ID", "구매계약ID", "현황"],
}

PK_BY_TABLE = {
    "T_발전소": "발전소ID",
    "T_구매계약": "구매계약ID",
    "T_수요기업": "수요기업ID",
    "T_판매계약": "판매계약ID",
    "T_전기사용지": "전기사용지ID",
    "T_수급매칭": "수급매칭ID",
}

LABEL_CANDIDATES = {
    "T_발전소": ["발전소명", "발전법인명"],
    "T_구매계약": ["발전소ID", "구매 담당자"],
    "T_수요기업": ["기업명"],
    "T_판매계약": ["수요기업ID", "계약유형"],
    "T_전기사용지": ["전기사용지명", "판매계약ID"],
    "T_수급매칭": ["현황", "전기사용지ID", "구매계약ID"],
}


def resolve_path(path_str: str, base_dir: Path) -> str:
    p = Path(path_str)
    if p.is_absolute():
        return str(p.resolve())
    return str((base_dir / p).resolve())


def find_first_xlsm(base_dir: Path) -> str:
    files = sorted(base_dir.glob("*.xlsm"))
    if not files:
        raise FileNotFoundError(f"XLSM not found in: {base_dir}")
    return str(files[0].resolve())


class App:
    def __init__(self, xlsm_path: str, html_path: str, build_python: str, build_script: str, form_js: str):
        self.xlsm_path = os.path.abspath(xlsm_path)
        self.html_path = os.path.abspath(html_path)
        self.build_python = os.path.abspath(build_python)
        self.build_script = os.path.abspath(build_script)
        self.form_js = os.path.abspath(form_js)

        self._rebuild_lock = threading.Lock()
        self._rebuild_running = False
        self._last_rebuild_error = None
        self._last_rebuild_ok_at = None

    def get_schema(self):
        return STATIC_SCHEMA

    def read_table(self, table_name: str):
        return read_table(self.xlsm_path, table_name)

    def save_table(self, table_name: str, record: dict):
        return write_table(self.xlsm_path, table_name, PK_BY_TABLE[table_name], record)

    def make_options_for_table(self, table_name: str):
        data = self.read_table(table_name)
        rows = data.get("rows") or []

        pk_name = PK_BY_TABLE[table_name]
        label_fields = LABEL_CANDIDATES.get(table_name, [])

        options = []
        seen = set()

        for row in rows:
            value = str(row.get(pk_name) or "").strip()
            if not value or value in seen:
                continue
            seen.add(value)

            extra = ""
            for field in label_fields:
                v = str(row.get(field) or "").strip()
                if v:
                    extra = v
                    break

            label = f"{value} | {extra}" if extra else value
            options.append({"value": value, "label": label})

        return options

    def get_record(self, table_name: str, pk_value: str):
        data = self.read_table(table_name)
        rows = data.get("rows") or []
        pk_name = PK_BY_TABLE[table_name]
        target = str(pk_value or "").strip()

        for row in rows:
            if str(row.get(pk_name) or "").strip() == target:
                return row
        return {}

    def rebuild_dashboard(self):
        script_dir = str(Path(self.build_script).resolve().parent)
        last_err = None

        for _ in range(5):
            temp_dir = None
            temp_xlsm = None
            try:
                temp_dir = tempfile.mkdtemp(prefix="ppa_build_")
                temp_xlsm = os.path.join(temp_dir, "source.xlsm")
                shutil.copy2(self.xlsm_path, temp_xlsm)

                cmd = [
                    self.build_python,
                    self.build_script,
                    "--xlsm", temp_xlsm,
                    "--out", self.html_path,
                ]

                proc = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    cwd=script_dir,
                )

                if proc.returncode == 0:
                    self._last_rebuild_error = None
                    self._last_rebuild_ok_at = time.strftime("%Y-%m-%d %H:%M:%S")
                    return

                last_err = (proc.stdout or "") + "\n" + (proc.stderr or "")
            except Exception as e:
                last_err = str(e)
            finally:
                if temp_dir:
                    shutil.rmtree(temp_dir, ignore_errors=True)

            time.sleep(1.0)

        self._last_rebuild_error = last_err
        raise RuntimeError(last_err or "대시보드 재생성 실패")

    def schedule_rebuild(self):
        with self._rebuild_lock:
            if self._rebuild_running:
                return False
            self._rebuild_running = True

        def worker():
            try:
                self.rebuild_dashboard()
            except Exception as e:
                self._last_rebuild_error = str(e)
            finally:
                with self._rebuild_lock:
                    self._rebuild_running = False

        threading.Thread(target=worker, daemon=True).start()
        return True

    def dashboard_html(self) -> str:
        html = Path(self.html_path).read_text(encoding="utf-8", errors="replace")
        tag = f'<script src="/dashboard_form.js?v={int(time.time())}"></script>'
        idx = html.lower().rfind("</body>")
        if idx >= 0:
            return html[:idx] + tag + "\n" + html[idx:]
        return html + "\n" + tag


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
            body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
            self._send(code, "application/json; charset=utf-8", body)

        def do_GET(self):
            parsed = urlparse(self.path)
            path = parsed.path
            qs = parse_qs(parsed.query)

            if path == "/dashboard_form.js":
                try:
                    self._send(200, "application/javascript; charset=utf-8", Path(app.form_js).read_bytes())
                except Exception as e:
                    self._send_json({"ok": False, "error": str(e)})
                return

            if path == "/api/schema":
                self._send_json({"ok": True, "schema": app.get_schema()})
                return

            if path == "/api/list-records":
                try:
                    table_name = (qs.get("table") or [""])[0]
                    if table_name not in PK_BY_TABLE:
                        self._send_json({"ok": False, "error": "지원하지 않는 테이블입니다."})
                        return
                    self._send_json({"ok": True, "options": app.make_options_for_table(table_name)})
                except Exception as e:
                    self._send_json({"ok": False, "error": str(e)})
                return

            if path == "/api/get-record":
                try:
                    table_name = (qs.get("table") or [""])[0]
                    pk_value = (qs.get("pk_value") or [""])[0]
                    if table_name not in PK_BY_TABLE:
                        self._send_json({"ok": False, "error": "지원하지 않는 테이블입니다."})
                        return
                    self._send_json({"ok": True, "record": app.get_record(table_name, pk_value)})
                except Exception as e:
                    self._send_json({"ok": False, "error": str(e)})
                return

            if path == "/api/options":
                self._send_json({"ok": True})
                return

            if path in ("/", "/index.html"):
                try:
                    self._send(200, "text/html; charset=utf-8", app.dashboard_html().encode("utf-8"))
                except Exception as e:
                    self._send(200, "text/plain; charset=utf-8", str(e).encode("utf-8"))
                return

            if path == "/favicon.ico":
                self._send(204, "image/x-icon", b"")
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

            if path == "/api/rebuild":
                try:
                    started = app.schedule_rebuild()
                    self._send_json({"ok": True, "started": started, "message": "대시보드 재생성을 시작했습니다."})
                except Exception as e:
                    self._send_json({"ok": False, "error": str(e)})
                return

            if path == "/api/save-table":
                try:
                    table_name = str(payload.get("table") or "").strip()
                    record = payload.get("record") or {}

                    if table_name not in PK_BY_TABLE:
                        self._send_json({"ok": False, "error": "지원하지 않는 테이블입니다."})
                        return

                    pk_name = PK_BY_TABLE[table_name]
                    if not str(record.get(pk_name) or "").strip():
                        self._send_json({"ok": False, "error": f"{pk_name}는 필수입니다."})
                        return

                    result = app.save_table(table_name, record)
                    self._send_json({"ok": True, "result": result, "message": "저장 완료"})
                except Exception as e:
                    self._send_json({"ok": False, "error": str(e)})
                return

            self._send_json({"ok": False, "error": "Not Found"})

        def log_message(self, fmt, *args):
            print("[HTTP]", fmt % args)

    return Handler


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--xlsm", default="")
    parser.add_argument("--html", default="dashboard.html")
    parser.add_argument("--build-script", default="build_dashboard.py")
    parser.add_argument("--python", default="")
    parser.add_argument("--form-js", default="dashboard_form.js")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    args = parser.parse_args()

    script_dir = Path(__file__).resolve().parent
    base_dir = script_dir.parent

    xlsm_path = resolve_path(args.xlsm, base_dir) if args.xlsm else find_first_xlsm(base_dir)
    html_path = resolve_path(args.html, script_dir)
    build_script = resolve_path(args.build_script, script_dir)
    form_js = resolve_path(args.form_js, script_dir)

    if args.python:
        build_python = resolve_path(args.python, base_dir) if not Path(args.python).is_absolute() else str(Path(args.python).resolve())
    else:
        build_python = sys.executable

    app = App(
        xlsm_path=xlsm_path,
        html_path=html_path,
        build_python=build_python,
        build_script=build_script,
        form_js=form_js,
    )

    if not os.path.exists(app.html_path):
        app.rebuild_dashboard()

    server = ThreadingHTTPServer((args.host, args.port), make_handler(app))
    print(f"[OK] 서버 시작: http://{args.host}:{args.port}")
    print(f"[INFO] XLSM: {app.xlsm_path}")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n[STOP] 종료")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()