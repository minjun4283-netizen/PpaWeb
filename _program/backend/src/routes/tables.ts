import { Router } from "express";
import { db, dataTableName, getTableDef, getColumnDefs, quoteIdent } from "../db.js";
import { requireAuth } from "../middleware/auth.js";
import { recordChange, type RowSnapshot } from "../services/changeLog.js";

export const tablesRouter = Router();
tablesRouter.use(requireAuth);

function loadTableOr404(tableKey: string, res: import("express").Response) {
  const tableDef = getTableDef(tableKey);
  if (!tableDef) {
    res.status(404).json({ error: `알 수 없는 표: ${tableKey}` });
    return null;
  }
  return tableDef;
}

function pickKnownColumns(
  tableKey: string,
  body: Record<string, unknown>
): { columns: string[]; values: unknown[] } {
  const knownKeys = new Set(getColumnDefs(tableKey).map((c) => c.col_key));
  const columns: string[] = [];
  const values: unknown[] = [];
  for (const [key, value] of Object.entries(body)) {
    if (knownKeys.has(key)) {
      columns.push(key);
      values.push(value);
    }
  }
  return { columns, values };
}

tablesRouter.get("/:tableKey/rows", (req, res) => {
  const { tableKey } = req.params;
  const tableDef = loadTableOr404(tableKey, res);
  if (!tableDef) return;

  const physical = dataTableName(tableKey);
  const rows = db.prepare(`SELECT * FROM ${quoteIdent(physical)} ORDER BY _id ASC`).all();
  res.json({ rows });
});

tablesRouter.post("/:tableKey/rows", (req, res) => {
  const { tableKey } = req.params;
  const tableDef = loadTableOr404(tableKey, res);
  if (!tableDef) return;

  const { columns, values } = pickKnownColumns(tableKey, req.body ?? {});
  const physical = dataTableName(tableKey);
  const username = req.user!.username;

  const colSql = ["_created_by", "_updated_by", ...columns]
    .map((c) => quoteIdent(c))
    .join(", ");
  const placeholders = ["?", "?", ...columns.map(() => "?")].join(", ");

  const info = db
    .prepare(`INSERT INTO ${quoteIdent(physical)} (${colSql}) VALUES (${placeholders})`)
    .run(username, username, ...values);

  const newRow = db
    .prepare(`SELECT * FROM ${quoteIdent(physical)} WHERE _id = ?`)
    .get(info.lastInsertRowid) as RowSnapshot;

  recordChange({
    username,
    tableKey,
    rowId: Number(info.lastInsertRowid),
    pkValue: String(newRow[tableDef.pk_column] ?? ""),
    changeType: "추가",
    oldRow: null,
    newRow,
  });

  res.status(201).json({ row: newRow });
});

tablesRouter.put("/:tableKey/rows/:id", (req, res) => {
  const { tableKey, id } = req.params;
  const tableDef = loadTableOr404(tableKey, res);
  if (!tableDef) return;

  const physical = dataTableName(tableKey);
  const oldRow = db
    .prepare(`SELECT * FROM ${quoteIdent(physical)} WHERE _id = ?`)
    .get(id) as RowSnapshot | undefined;

  if (!oldRow) {
    res.status(404).json({ error: "행을 찾을 수 없습니다." });
    return;
  }

  const { columns, values } = pickKnownColumns(tableKey, req.body ?? {});
  const username = req.user!.username;

  if (columns.length > 0) {
    const setSql = columns.map((c) => `${quoteIdent(c)} = ?`).join(", ");
    db.prepare(
      `UPDATE ${quoteIdent(physical)} SET ${setSql}, _updated_by = ?, _updated_at = datetime('now') WHERE _id = ?`
    ).run(...values, username, id);
  }

  const newRow = db
    .prepare(`SELECT * FROM ${quoteIdent(physical)} WHERE _id = ?`)
    .get(id) as RowSnapshot;

  recordChange({
    username,
    tableKey,
    rowId: Number(id),
    pkValue: String(newRow[tableDef.pk_column] ?? ""),
    changeType: "수정",
    oldRow,
    newRow,
  });

  res.json({ row: newRow });
});

tablesRouter.delete("/:tableKey/rows/:id", (req, res) => {
  const { tableKey, id } = req.params;
  const tableDef = loadTableOr404(tableKey, res);
  if (!tableDef) return;

  const physical = dataTableName(tableKey);
  const oldRow = db
    .prepare(`SELECT * FROM ${quoteIdent(physical)} WHERE _id = ?`)
    .get(id) as RowSnapshot | undefined;

  if (!oldRow) {
    res.status(404).json({ error: "행을 찾을 수 없습니다." });
    return;
  }

  db.prepare(`DELETE FROM ${quoteIdent(physical)} WHERE _id = ?`).run(id);

  recordChange({
    username: req.user!.username,
    tableKey,
    rowId: Number(id),
    pkValue: String(oldRow[tableDef.pk_column] ?? ""),
    changeType: "삭제",
    oldRow,
    newRow: null,
  });

  res.json({ ok: true });
});
