#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""excel_com.ExcelBridge의 _save_record/_write_row/_cascade_rename(4개 표
ID 자동생성 + 연쇄 갱신)을 COM 없이 검증하는 통합 테스트 - 실제 클래스
메서드를 그대로 쓰되 워크시트 I/O만 인메모리로 흉내낸다.

실행: python3 test_id_cascade.py (이 폴더 안에서, 또는 아무 위치에서나)
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from ppa_schema import TABLE_BY_KEY
from excel_com import _cell_to_text


class FakeListObjects:
    Count = 0

    def __call__(self, i):
        raise IndexError("no ListObjects in fake sheet")


class FakeCell:
    def __init__(self, sheet, row, col):
        self.sheet, self.row, self.col = sheet, row, col

    @property
    def Value(self):
        return self.sheet.get(self.row, self.col)

    @Value.setter
    def Value(self, v):
        self.sheet.set(self.row, self.col, v)

    @property
    def NumberFormat(self):
        return "General"

    @NumberFormat.setter
    def NumberFormat(self, v):
        pass


class FakeWorksheet:
    def __init__(self, headers):
        self.headers = list(headers)
        self.data_rows = []  # list[list]
        self.ListObjects = FakeListObjects()

    def get(self, row, col):
        if row == 1:
            return self.headers[col - 1] if col - 1 < len(self.headers) else None
        idx = row - 2
        if idx < 0 or idx >= len(self.data_rows):
            return None
        r = self.data_rows[idx]
        return r[col - 1] if col - 1 < len(r) else None

    def set(self, row, col, v):
        if row == 1:
            while len(self.headers) < col:
                self.headers.append(None)
            self.headers[col - 1] = v
            return
        idx = row - 2
        while len(self.data_rows) <= idx:
            self.data_rows.append([None] * len(self.headers))
        r = self.data_rows[idx]
        while len(r) < col:
            r.append(None)
        r[col - 1] = v

    def Cells(self, row, col):
        return FakeCell(self, row, col)

    @property
    def UsedRange(self):
        headers = tuple(self.headers)
        rows = [headers]
        for r in self.data_rows:
            padded = list(r) + [None] * (len(self.headers) - len(r))
            rows.append(tuple(padded))
        outer = self

        class _Used:
            Row = 1
            Column = 1
            Value = tuple(rows)
        return _Used()


class FakeWorkbook:
    def __init__(self, sheets):
        self.sheets = sheets
        self.save_count = 0

    def Worksheets(self, name):
        return self.sheets[name]

    def Save(self):
        self.save_count += 1


def make_bridge_and_wb(seed_rows: dict[str, list[dict]]):
    from excel_com import ExcelBridge
    bridge = ExcelBridge.__new__(ExcelBridge)
    bridge._last_unmatched = {}

    sheets = {}
    for key, schema in TABLE_BY_KEY.items():
        ws = FakeWorksheet(schema.columns)
        for row in seed_rows.get(key, []):
            ws.data_rows.append([row.get(c, "") for c in schema.columns])
        sheets[key] = ws
    wb = FakeWorkbook(sheets)
    bridge._ensure_workbook = lambda: wb  # noqa: E731
    return bridge, wb


def dump(wb, table_key):
    schema = TABLE_BY_KEY[table_key]
    ws = wb.sheets[table_key]
    out = []
    for r in ws.data_rows:
        out.append({c: _cell_to_text(r[i]) if i < len(r) else "" for i, c in enumerate(schema.columns)})
    return out


def find(rows, pk_col, pk_val):
    for r in rows:
        if str(r.get(pk_col) or "") == pk_val:
            return r
    return None


# ---------------------------------------------------------------
# 시나리오 데이터: 구매-P001-2024-01 / 판매-D001-2024-01 /
# 전기사용지-D001-2024-01-공장A / 매칭-D001-2024-01
seed = {
    "T_발전소": [
        {"발전소ID": "P001", "발전소명": "테스트발전소1", "설비용량(MW)": "10", "발전원": "태양광"},
        {"발전소ID": "P002", "발전소명": "테스트발전소2", "설비용량(MW)": "8", "발전원": "풍력"},
    ],
    "T_구매계약": [{
        "구매계약ID": "구매-P001-2024-01", "발전소ID": "P001", "구매계약용량(MW)": "5",
        "공급기한_구매": "2024-01-01", "계약기간(년)": "10",
    }],
    "T_수요기업": [
        {"수요기업ID": "D001", "기업명": "테스트수요기업1"},
        {"수요기업ID": "D002", "기업명": "테스트수요기업2"},
    ],
    "T_판매계약": [{
        "판매계약ID": "판매-D001-2024-01", "수요기업ID": "D001", "판매계약용량(MW)": "5",
        "계약일": "2024-01-01", "공급기한_판매": "2024-01-01",
    }],
    "T_전기사용지": [{
        "전기사용지ID": "전기사용지-D001-2024-01-공장A", "판매계약ID": "판매-D001-2024-01",
        "전기사용지명": "공장A", "전기사용지계약용량(MW)": "5",
    }],
    "T_수급매칭": [{
        "수급매칭ID": "매칭-D001-2024-01", "전기사용지ID": "전기사용지-D001-2024-01-공장A",
        "구매계약ID": "구매-P001-2024-01", "현황": "1. 공급 중",
    }],
}


def test_1_new_purchase_id():
    bridge, wb = make_bridge_and_wb(seed)
    record = {"발전소ID": "P001", "구매계약용량(MW)": "3", "공급기한_구매": "2024-06-01", "계약기간(년)": "5"}
    result = bridge._save_record("T_구매계약", record, original_pk=None)
    assert result["pk_value"] == "구매-P001-2024-02", result
    rows = dump(wb, "T_구매계약")
    assert find(rows, "구매계약ID", "구매-P001-2024-02") is not None
    assert find(rows, "구매계약ID", "구매-P001-2024-01") is not None  # 기존 행 유지
    print("test_1_new_purchase_id OK ->", result["pk_value"])


def test_2_edit_purchase_no_id_change():
    bridge, wb = make_bridge_and_wb(seed)
    record = {
        "구매계약ID": "구매-P001-2024-01", "발전소ID": "P001", "구매계약용량(MW)": "9",
        "공급기한_구매": "2024-01-01", "계약기간(년)": "10",
    }
    result = bridge._save_record("T_구매계약", record, original_pk="구매-P001-2024-01")
    assert result["pk_value"] == "구매-P001-2024-01"
    assert result["action"] == "updated"
    rows = dump(wb, "T_구매계약")
    row = find(rows, "구매계약ID", "구매-P001-2024-01")
    assert row["구매계약용량(MW)"] == "9"
    assert len(rows) == 1  # 새 행이 안 생겨야 함
    print("test_2_edit_purchase_no_id_change OK")


def test_3_edit_purchase_plant_changes_cascades_fk_only():
    bridge, wb = make_bridge_and_wb(seed)
    record = {
        "구매계약ID": "구매-P001-2024-01", "발전소ID": "P002", "구매계약용량(MW)": "5",
        "공급기한_구매": "2024-01-01", "계약기간(년)": "10",
    }
    result = bridge._save_record("T_구매계약", record, original_pk="구매-P001-2024-01")
    new_pk = "구매-P002-2024-01"
    assert result["pk_value"] == new_pk, result
    # 구매계약 표: 이름 바뀐 행 하나, 옛 이름 행 없음
    purch_rows = dump(wb, "T_구매계약")
    assert find(purch_rows, "구매계약ID", new_pk) is not None
    assert find(purch_rows, "구매계약ID", "구매-P001-2024-01") is None
    # 수급매칭의 구매계약ID FK가 갱신됐는지 - 수급매칭ID 자체는 안 바뀜(구매 쪽은 FK만)
    match_rows = dump(wb, "T_수급매칭")
    m = find(match_rows, "수급매칭ID", "매칭-D001-2024-01")
    assert m is not None, "수급매칭ID가 그대로 유지돼야 함"
    assert m["구매계약ID"] == new_pk, m
    print("test_3_edit_purchase_plant_changes_cascades_fk_only OK ->", result.get("cascaded"))


def test_4_edit_sale_demand_changes_cascades_two_levels():
    bridge, wb = make_bridge_and_wb(seed)
    record = {
        "판매계약ID": "판매-D001-2024-01", "수요기업ID": "D002", "판매계약용량(MW)": "5",
        "계약일": "2024-01-01", "공급기한_판매": "2024-01-01",
    }
    result = bridge._save_record("T_판매계약", record, original_pk="판매-D001-2024-01")
    new_sale_pk = "판매-D002-2024-01"
    assert result["pk_value"] == new_sale_pk, result

    site_rows = dump(wb, "T_전기사용지")
    new_site_pk = "전기사용지-D002-2024-01-공장A"
    site = find(site_rows, "전기사용지ID", new_site_pk)
    assert site is not None, site_rows
    assert site["판매계약ID"] == new_sale_pk
    assert find(site_rows, "전기사용지ID", "전기사용지-D001-2024-01-공장A") is None

    match_rows = dump(wb, "T_수급매칭")
    new_match_pk = "매칭-D002-2024-01"
    m = find(match_rows, "수급매칭ID", new_match_pk)
    assert m is not None, match_rows
    assert m["전기사용지ID"] == new_site_pk
    assert find(match_rows, "수급매칭ID", "매칭-D001-2024-01") is None

    cascaded = result.get("cascaded", [])
    tables_touched = {c["table"] for c in cascaded}
    assert tables_touched == {"T_전기사용지", "T_수급매칭"}, cascaded
    print("test_4_edit_sale_demand_changes_cascades_two_levels OK ->", cascaded)


def test_5_rename_electric_site_name_only_keeps_seq():
    bridge, wb = make_bridge_and_wb(seed)
    record = {
        "전기사용지ID": "전기사용지-D001-2024-01-공장A", "판매계약ID": "판매-D001-2024-01",
        "전기사용지명": "공장B", "전기사용지계약용량(MW)": "5",
    }
    result = bridge._save_record("T_전기사용지", record, original_pk="전기사용지-D001-2024-01-공장A")
    new_pk = "전기사용지-D001-2024-01-공장B"
    assert result["pk_value"] == new_pk, result
    match_rows = dump(wb, "T_수급매칭")
    m = find(match_rows, "수급매칭ID", "매칭-D001-2024-01")  # 수급매칭ID 자체는 전기사용지명에 안 의존 -> 안 바뀜
    assert m is not None
    assert m["전기사용지ID"] == new_pk
    print("test_5_rename_electric_site_name_only_keeps_seq OK ->", result.get("cascaded"))


def test_6_temporary_contract_gets_T_year():
    bridge, wb = make_bridge_and_wb(seed)
    record = {"발전소ID": "P001", "구매계약용량(MW)": "1", "공급기한_구매": "", "계약기간(년)": ""}
    result = bridge._save_record("T_구매계약", record, original_pk=None)
    assert result["pk_value"] == "구매-P001-T-01", result
    print("test_6_temporary_contract_gets_T_year OK ->", result["pk_value"])


def test_7_temporary_to_real_date_cascades():
    bridge, wb = make_bridge_and_wb(seed)
    # 먼저 임시 판매계약 하나 신규 생성
    r1 = bridge._save_record(
        "T_판매계약",
        {"수요기업ID": "D001", "판매계약용량(MW)": "1", "계약일": "", "공급기한_판매": ""},
        original_pk=None,
    )
    temp_pk = r1["pk_value"]
    assert temp_pk == "판매-D001-T-01", r1
    # 전기사용지를 그 임시 계약에 연결해서 신규 생성
    r2 = bridge._save_record(
        "T_전기사용지",
        {"판매계약ID": temp_pk, "전기사용지명": "신규공장", "전기사용지계약용량(MW)": "1"},
        original_pk=None,
    )
    assert r2["pk_value"] == "전기사용지-D001-T-01-신규공장", r2

    # 이제 실제 날짜가 정해져서 판매계약을 편집 -> T가 실제 연도로 바뀌고 연쇄돼야 함
    r3 = bridge._save_record(
        "T_판매계약",
        {"판매계약ID": temp_pk, "수요기업ID": "D001", "판매계약용량(MW)": "1",
         "계약일": "2026-03-01", "공급기한_판매": "2026-03-01"},
        original_pk=temp_pk,
    )
    new_sale_pk = "판매-D001-2026-01"
    assert r3["pk_value"] == new_sale_pk, r3
    site_rows = dump(wb, "T_전기사용지")
    new_site_pk = "전기사용지-D001-2026-01-신규공장"
    assert find(site_rows, "전기사용지ID", new_site_pk) is not None, site_rows
    assert find(site_rows, "전기사용지ID", "전기사용지-D001-T-01-신규공장") is None
    print("test_7_temporary_to_real_date_cascades OK ->", r3.get("cascaded"))


def test_8_sequence_numbering_avoids_collision():
    bridge, wb = make_bridge_and_wb(seed)
    ids = []
    for i in range(3):
        r = bridge._save_record(
            "T_구매계약",
            {"발전소ID": "P001", "구매계약용량(MW)": "1", "공급기한_구매": "2030-01-01", "계약기간(년)": "1"},
            original_pk=None,
        )
        ids.append(r["pk_value"])
    assert ids == ["구매-P001-2030-01", "구매-P001-2030-02", "구매-P001-2030-03"], ids
    print("test_8_sequence_numbering_avoids_collision OK ->", ids)


def test_9_matching_id_ignores_purchase_side_changes():
    bridge, wb = make_bridge_and_wb(seed)
    # 구매계약 발전소를 바꿔도(캐스케이드 1단계) 수급매칭ID 문자열 자체는 절대 안 바뀜(이미 test_3에서
    # 확인) - 여기서는 반대로 수급매칭 자체를 신규 생성할 때 규칙이 맞는지 확인.
    r = bridge._save_record(
        "T_수급매칭",
        {"전기사용지ID": "전기사용지-D001-2024-01-공장A", "구매계약ID": "구매-P001-2024-01", "현황": "1. 공급 중"},
        original_pk=None,
    )
    assert r["pk_value"] == "매칭-D001-2024-02", r  # 기존 매칭-D001-2024-01 다음 번호
    print("test_9_matching_id_ignores_purchase_side_changes OK ->", r["pk_value"])


if __name__ == "__main__":
    tests = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    for t in tests:
        t()
    print(f"\n=== ALL {len(tests)} TESTS PASSED ===")
