import { db } from "../db.js";

export type ChangeType = "추가" | "수정" | "삭제";

export interface RowSnapshot {
  [column: string]: unknown;
}

function diffDescription(oldRow: RowSnapshot | null, newRow: RowSnapshot | null): string {
  if (!oldRow) return "신규 레코드 추가";
  if (!newRow) return "기존 레코드 삭제";

  const keys = new Set([...Object.keys(oldRow), ...Object.keys(newRow)]);
  const changed: string[] = [];
  for (const key of keys) {
    if (key.startsWith("_")) continue;
    const before = oldRow[key] ?? "";
    const after = newRow[key] ?? "";
    if (String(before) !== String(after)) changed.push(key);
  }
  return changed.length === 0 ? "변경항목: 없음" : `변경항목(${changed.length}): ${changed.join(", ")}`;
}

export function recordChange(params: {
  username: string;
  tableKey: string;
  rowId: number | null;
  pkValue: string;
  changeType: ChangeType;
  oldRow: RowSnapshot | null;
  newRow: RowSnapshot | null;
}) {
  const { username, tableKey, rowId, pkValue, changeType, oldRow, newRow } = params;
  db.prepare(
    `INSERT INTO change_log (username, table_key, row_id, pk_value, change_type, old_data, new_data, description)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
  ).run(
    username,
    tableKey,
    rowId,
    pkValue,
    changeType,
    oldRow ? JSON.stringify(oldRow) : null,
    newRow ? JSON.stringify(newRow) : null,
    diffDescription(oldRow, newRow)
  );
}

export interface ChangeLogRow {
  id: number;
  changed_at: string;
  username: string;
  table_key: string;
  row_id: number | null;
  pk_value: string;
  change_type: ChangeType;
  old_data: string | null;
  new_data: string | null;
  description: string;
}

export function listChangeLog(filters: { tableKey?: string; limit?: number }): ChangeLogRow[] {
  const limit = Math.min(filters.limit ?? 200, 1000);
  if (filters.tableKey) {
    return db
      .prepare(
        `SELECT * FROM change_log WHERE table_key = ? ORDER BY id DESC LIMIT ?`
      )
      .all(filters.tableKey, limit) as ChangeLogRow[];
  }
  return db
    .prepare(`SELECT * FROM change_log ORDER BY id DESC LIMIT ?`)
    .all(limit) as ChangeLogRow[];
}
