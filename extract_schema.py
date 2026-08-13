#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import json
import re
import zipfile
from xml.etree import ElementTree as ET

TARGET_TABLES = [
    "T_발전소",
    "T_구매계약",
    "T_수요기업",
    "T_판매계약",
    "T_전기사용지",
    "T_수급매칭",
]

NS_MAIN = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
NS_REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL_PKG = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}


def is_blank(value) -> bool:
    return value is None or str(value).strip() == ""


def normalize_name(text: str) -> str:
    return re.sub(r"[\s_\-]+", "", str(text or "")).strip().lower()


def is_helper_column(name: str) -> bool:
    if not name:
        return True

    bad_exact = {
        "PK중복",
        "PK공란",
        "조합중복",
        "열1",
    }

    if name in bad_exact:
        return True
    if name.endswith("참조"):
        return True
    if name.endswith("공란"):
        return True
    if name.endswith("중복"):
        return True

    return False


def col_ref_to_index(cell_ref: str) -> int:
    m = re.match(r"([A-Z]+)", (cell_ref or "").upper())
    if not m:
        return 0

    letters = m.group(1)
    result = 0
    for ch in letters:
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


def read_xml_from_zip(zf: zipfile.ZipFile, path: str):
    data = zf.read(path)
    return ET.fromstring(data)


def load_shared_strings(zf: zipfile.ZipFile):
    path = "xl/sharedStrings.xml"
    if path not in zf.namelist():
        return []

    root = read_xml_from_zip(zf, path)
    items = []

    for si in root.findall(".//x:si", NS_MAIN):
        texts = []
        for t in si.findall(".//x:t", NS_MAIN):
            texts.append(t.text or "")
        items.append("".join(texts))

    return items


def load_sheet_name_to_path(zf: zipfile.ZipFile):
    wb_root = read_xml_from_zip(zf, "xl/workbook.xml")
    rel_root = read_xml_from_zip(zf, "xl/_rels/workbook.xml.rels")

    rel_map = {}
    for rel in rel_root.findall("r:Relationship", NS_REL_PKG):
        rel_id = rel.attrib.get("Id", "")
        target = rel.attrib.get("Target", "")
        if target.startswith("/"):
            internal = target.lstrip("/")
        elif target.startswith("xl/"):
            internal = target
        else:
            internal = "xl/" + target.lstrip("/")
        internal = internal.replace("\\", "/")
        rel_map[rel_id] = internal

    result = {}
    for sheet in wb_root.findall(".//x:sheets/x:sheet", NS_MAIN):
        name = sheet.attrib.get("name", "")
        rel_id = sheet.attrib.get(f"{{{NS_REL_DOC}}}id", "")
        sheet_path = rel_map.get(rel_id, "")
        if name and sheet_path:
            result[name] = sheet_path

    return result


def extract_cell_value(cell, shared_strings):
    cell_type = cell.attrib.get("t", "")

    if cell_type == "inlineStr":
        texts = [t.text or "" for t in cell.findall(".//x:t", NS_MAIN)]
        return "".join(texts).strip()

    v = cell.find("x:v", NS_MAIN)
    raw = "" if v is None or v.text is None else v.text

    if cell_type == "s":
        try:
            idx = int(raw)
            if 0 <= idx < len(shared_strings):
                return str(shared_strings[idx]).strip()
        except Exception:
            pass
        return ""

    if cell_type == "b":
        return "TRUE" if raw == "1" else "FALSE"

    return str(raw).strip()


def extract_headers_from_sheet_xml(zf: zipfile.ZipFile, sheet_path: str, shared_strings):
    root = read_xml_from_zip(zf, sheet_path)

    best_headers = []
    rows = root.findall(".//x:sheetData/x:row", NS_MAIN)

    for row in rows[:20]:
        row_map = {}

        for cell in row.findall("x:c", NS_MAIN):
            ref = cell.attrib.get("r", "")
            col_idx = col_ref_to_index(ref)
            if col_idx <= 0:
                continue

            value = extract_cell_value(cell, shared_strings)
            if is_blank(value):
                continue

            row_map[col_idx] = value

        ordered = [row_map[k] for k in sorted(row_map.keys())]
        headers = [str(v).strip() for v in ordered if not is_blank(v)]

        if len(headers) > len(best_headers):
            best_headers = headers

    seen = set()
    result = []
    for h in best_headers:
        if not h:
            continue
        if h in seen:
            continue
        if is_helper_column(h):
            continue
        seen.add(h)
        result.append(h)

    return result


def get_schema(xlsm_path: str):
    with zipfile.ZipFile(xlsm_path, "r") as zf:
        shared_strings = load_shared_strings(zf)
        sheet_map = load_sheet_name_to_path(zf)

        normalized_sheet_map = {normalize_name(k): v for k, v in sheet_map.items()}

        schema = {}
        for table_name in TARGET_TABLES:
            sheet_path = sheet_map.get(table_name)

            if not sheet_path:
                sheet_path = normalized_sheet_map.get(normalize_name(table_name), "")

            if not sheet_path:
                sheet_path = normalized_sheet_map.get(normalize_name(table_name.replace("T_", "")), "")

            if sheet_path:
                try:
                    schema[table_name] = extract_headers_from_sheet_xml(zf, sheet_path, shared_strings)
                except Exception:
                    schema[table_name] = []
            else:
                schema[table_name] = []

        return schema


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--xlsm", required=True)
    args = parser.parse_args()

    schema = get_schema(args.xlsm)
    print(json.dumps({"ok": True, "schema": schema}, ensure_ascii=False))


if __name__ == "__main__":
    main()