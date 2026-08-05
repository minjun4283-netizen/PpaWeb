import fs from "node:fs";
import path from "node:path";
import Database from "better-sqlite3";
import { config } from "./config.js";
import { TABLE_SEEDS } from "./schema/tableDefs.js";

fs.mkdirSync(path.dirname(config.dbPath), { recursive: true });

export const db = new Database(config.dbPath);
db.pragma("journal_mode = WAL");
db.pragma("foreign_keys = ON");

// Only letters (incl. Korean), digits, and underscore are allowed in dynamic
// identifiers we splice into SQL. Anything else is rejected before it ever
// reaches a query string.
const SAFE_IDENTIFIER = /^[A-Za-z0-9_가-힣]+$/;

export function assertSafeIdentifier(name: string): string {
  if (!SAFE_IDENTIFIER.test(name)) {
    throw new Error(`Unsafe identifier: ${name}`);
  }
  return name;
}

export function quoteIdent(name: string): string {
  assertSafeIdentifier(name);
  return `"${name}"`;
}

export function dataTableName(tableKey: string): string {
  return `data_${tableKey}`;
}

function ensureMetaTables() {
  db.exec(`
    CREATE TABLE IF NOT EXISTS users (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT UNIQUE NOT NULL,
      password_hash TEXT NOT NULL,
      display_name TEXT NOT NULL,
      role TEXT NOT NULL CHECK (role IN ('admin','user')),
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    );

    CREATE TABLE IF NOT EXISTS table_defs (
      table_key TEXT PRIMARY KEY,
      label TEXT NOT NULL,
      pk_column TEXT NOT NULL,
      sort_order INTEGER NOT NULL
    );

    CREATE TABLE IF NOT EXISTS column_defs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      table_key TEXT NOT NULL,
      col_key TEXT NOT NULL,
      label TEXT NOT NULL,
      type TEXT NOT NULL CHECK (type IN ('text','number','date')),
      sort_order INTEGER NOT NULL,
      is_fk INTEGER NOT NULL DEFAULT 0,
      ref_table TEXT,
      ref_column TEXT,
      UNIQUE(table_key, col_key)
    );

    CREATE TABLE IF NOT EXISTS unique_groups (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      table_key TEXT NOT NULL,
      group_index INTEGER NOT NULL,
      col_key TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS change_log (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      changed_at TEXT NOT NULL DEFAULT (datetime('now')),
      username TEXT NOT NULL,
      table_key TEXT NOT NULL,
      row_id INTEGER,
      pk_value TEXT NOT NULL,
      change_type TEXT NOT NULL CHECK (change_type IN ('추가','수정','삭제')),
      old_data TEXT,
      new_data TEXT,
      description TEXT
    );

    CREATE TABLE IF NOT EXISTS access_sync_log (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      synced_at TEXT NOT NULL DEFAULT (datetime('now')),
      username TEXT NOT NULL,
      table_key TEXT NOT NULL,
      status TEXT NOT NULL,
      message TEXT
    );
  `);
}

// Keeps table_defs.sort_order in sync with TABLE_SEEDS' declaration order
// even on a database that was already seeded under an older ordering —
// INSERT OR IGNORE alone would never touch sort_order on existing rows.
function syncTableSortOrder() {
  const updateOrder = db.prepare(`UPDATE table_defs SET sort_order = ? WHERE table_key = ?`);
  TABLE_SEEDS.forEach((table, index) => updateOrder.run(index, table.key));
}

function seedTableDefsIfMissing() {
  const insertTable = db.prepare(
    `INSERT OR IGNORE INTO table_defs (table_key, label, pk_column, sort_order) VALUES (?, ?, ?, ?)`
  );
  const insertColumn = db.prepare(
    `INSERT OR IGNORE INTO column_defs (table_key, col_key, label, type, sort_order, is_fk, ref_table, ref_column)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
  );
  const insertUniqueGroupCol = db.prepare(
    `INSERT INTO unique_groups (table_key, group_index, col_key) VALUES (?, ?, ?)`
  );
  const hasUniqueGroups = db.prepare(
    `SELECT COUNT(*) as c FROM unique_groups WHERE table_key = ?`
  );

  TABLE_SEEDS.forEach((table, index) => {
    insertTable.run(table.key, table.label, table.pk, index);

    table.columns.forEach((col, colIndex) => {
      const fk = table.foreignKeys.find((f) => f.column === col.key);
      insertColumn.run(
        table.key,
        col.key,
        col.label,
        col.type,
        colIndex,
        fk ? 1 : 0,
        fk?.refTable ?? null,
        fk?.refColumn ?? null
      );
    });

    const existing = hasUniqueGroups.get(table.key) as { c: number };
    if (existing.c === 0) {
      table.uniqueGroups.forEach((group, groupIndex) => {
        group.forEach((colKey) => insertUniqueGroupCol.run(table.key, groupIndex, colKey));
      });
    }
  });
}

export interface ColumnDefRow {
  id: number;
  table_key: string;
  col_key: string;
  label: string;
  type: "text" | "number" | "date";
  sort_order: number;
  is_fk: number;
  ref_table: string | null;
  ref_column: string | null;
}

export function getColumnDefs(tableKey: string): ColumnDefRow[] {
  return db
    .prepare(`SELECT * FROM column_defs WHERE table_key = ? ORDER BY sort_order ASC`)
    .all(tableKey) as ColumnDefRow[];
}

function syncPhysicalTable(tableKey: string) {
  const physical = dataTableName(tableKey);
  quoteIdent(physical);

  db.exec(`
    CREATE TABLE IF NOT EXISTS ${quoteIdent(physical)} (
      _id INTEGER PRIMARY KEY AUTOINCREMENT,
      _created_at TEXT NOT NULL DEFAULT (datetime('now')),
      _updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      _created_by TEXT,
      _updated_by TEXT
    );
  `);

  const existingCols = new Set(
    (db.prepare(`PRAGMA table_info(${quoteIdent(physical)})`).all() as { name: string }[]).map(
      (c) => c.name
    )
  );

  for (const col of getColumnDefs(tableKey)) {
    if (!existingCols.has(col.col_key)) {
      const sqlType = col.type === "number" ? "REAL" : "TEXT";
      db.exec(
        `ALTER TABLE ${quoteIdent(physical)} ADD COLUMN ${quoteIdent(col.col_key)} ${sqlType};`
      );
    }
  }
}

export function initDatabase() {
  ensureMetaTables();
  seedTableDefsIfMissing();
  syncTableSortOrder();
  for (const table of TABLE_SEEDS) {
    syncPhysicalTable(table.key);
  }
}

export interface TableDefRow {
  table_key: string;
  label: string;
  pk_column: string;
  sort_order: number;
}

export function listTableDefs(): TableDefRow[] {
  return db.prepare(`SELECT * FROM table_defs ORDER BY sort_order ASC`).all() as TableDefRow[];
}

export function getTableDef(tableKey: string): TableDefRow | undefined {
  return db.prepare(`SELECT * FROM table_defs WHERE table_key = ?`).get(tableKey) as
    | TableDefRow
    | undefined;
}

export function addColumnDef(
  tableKey: string,
  colKey: string,
  label: string,
  type: "text" | "number" | "date"
): void {
  assertSafeIdentifier(colKey);
  const maxOrder = db
    .prepare(`SELECT COALESCE(MAX(sort_order), -1) as m FROM column_defs WHERE table_key = ?`)
    .get(tableKey) as { m: number };

  db.prepare(
    `INSERT INTO column_defs (table_key, col_key, label, type, sort_order, is_fk, ref_table, ref_column)
     VALUES (?, ?, ?, ?, ?, 0, NULL, NULL)`
  ).run(tableKey, colKey, label, type, maxOrder.m + 1);

  const physical = dataTableName(tableKey);
  const sqlType = type === "number" ? "REAL" : "TEXT";
  db.exec(`ALTER TABLE ${quoteIdent(physical)} ADD COLUMN ${quoteIdent(colKey)} ${sqlType};`);
}

export function getUniqueGroups(tableKey: string): string[][] {
  const rows = db
    .prepare(
      `SELECT group_index, col_key FROM unique_groups WHERE table_key = ? ORDER BY group_index ASC`
    )
    .all(tableKey) as { group_index: number; col_key: string }[];

  const groups = new Map<number, string[]>();
  for (const row of rows) {
    if (!groups.has(row.group_index)) groups.set(row.group_index, []);
    groups.get(row.group_index)!.push(row.col_key);
  }
  return [...groups.values()];
}
