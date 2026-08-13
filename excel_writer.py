#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import contextlib
import time


class ExcelWriteError(Exception):
    pass


def _norm(v) -> str:
    return "" if v is None else str(v).strip()


def _find_listobject(workbook, table_name: str):
    for ws in workbook.Worksheets:
        for lo in ws.ListObjects:
            if str(lo.Name).strip() == table_name:
                return ws, lo
    raise ExcelWriteError(f"엑셀 테이블(ListObject)을 찾지 못했습니다: {table_name}")


def _headers_of(lo) -> list[str]:
    headers = []
    for i in range(1, lo.ListColumns.Count + 1):
        headers.append(str(lo.HeaderRowRange.Cells(1, i).Value).strip())
    return headers


def upsert_row_to_table(xlsm_path: str, table_name: str, pk_col: str, row: dict) -> dict:
    try:
        import pythoncom
        import win32com.client
    except Exception as e:
        raise ExcelWriteError(
            "pywin32가 필요합니다. `pip install pywin32` 후 다시 시도하세요."
        ) from e

    pythoncom.CoInitialize()
    excel = None
    wb = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False

        try:
            wb = excel.Workbooks.Open(
                xlsm_path,
                UpdateLinks=0,
                ReadOnly=False,
                IgnoreReadOnlyRecommended=True,
            )
        except Exception as e:
            raise ExcelWriteError(
                f"엑셀 파일을 열지 못했습니다: {xlsm_path}\n"
                f"엑셀에서 파일을 열어둔 상태라면 닫고 다시 시도해주세요.\n"
                f"원본 오류: {e}"
            ) from e

        _, lo = _find_listobject(wb, table_name)
        headers = _headers_of(lo)

        if pk_col not in headers:
            raise ExcelWriteError(f"PK 컬럼이 엑셀 테이블 헤더에 없습니다: {pk_col}")

        pk_val = _norm(row.get(pk_col))
        if not pk_val:
            raise ExcelWriteError(f"PK 값이 비어 있습니다: {pk_col}")

        pk_idx = headers.index(pk_col) + 1
        row_values = {h: row.get(h, "") for h in headers}

        target_row_index = None
        data_body = lo.DataBodyRange

        if data_body is not None:
            for r in range(1, data_body.Rows.Count + 1):
                cur = _norm(data_body.Cells(r, pk_idx).Value)
                if cur == pk_val:
                    target_row_index = r
                    break

        created = False
        if target_row_index is None:
            list_row = lo.ListRows.Add()
            target_range = list_row.Range
            created = True
        else:
            target_range = lo.DataBodyRange.Rows(target_row_index)

        for c_idx, header in enumerate(headers, start=1):
            val = row_values.get(header, "")
            if val == "":
                target_range.Cells(1, c_idx).Value = ""
            else:
                target_range.Cells(1, c_idx).Value = val

        wb.Save()

        # 저장 직후 파일 잠금/OneDrive 동기화 완화
        time.sleep(2.5)

        return {
            "ok": True,
            "action": "inserted" if created else "updated",
            "table": table_name,
            "pk_col": pk_col,
            "pk_value": pk_val,
        }

    except ExcelWriteError:
        raise
    except Exception as e:
        raise ExcelWriteError(f"엑셀 저장 중 오류: {e}") from e

    finally:
        with contextlib.suppress(Exception):
            if wb is not None:
                wb.Close(SaveChanges=False)

        time.sleep(1.0)

        with contextlib.suppress(Exception):
            if excel is not None:
                excel.Quit()

        time.sleep(1.5)

        with contextlib.suppress(Exception):
            pythoncom.CoUninitialize()