// Ports 엑셀시트별_Access테이블전송 / AE_ExportOneSheetToAccess from the original VBA macro.
//
// This only works when the Node process itself runs on Windows with the
// Microsoft Access Database Engine (ACE OLEDB) redistributable installed —
// identical to the requirement the VBA macro already had. On any other
// platform (e.g. a Linux intranet server), calling this will fail fast with
// a clear error instead of crashing the whole app, so the rest of the
// website keeps working; run this piece on a Windows box that can see the
// .accdb file (same machine or a mapped network share) if you need it.
import { config } from "../config.js";
import { db, dataTableName, getColumnDefs, getTableDef, quoteIdent } from "../db.js";

export function accessSyncAvailable(): boolean {
  return process.platform === "win32" && config.accessDbPath !== "";
}

function toSqlValue(value: unknown): string {
  if (value === null || value === undefined) return "Null";
  const str = String(value).trim();
  if (str === "") return "Null";

  if (typeof value === "number" || /^-?\d+(\.\d+)?$/.test(str)) {
    return str;
  }

  const dateMatch = /^\d{4}-\d{2}-\d{2}([ T]\d{2}:\d{2}(:\d{2})?)?$/.test(str);
  if (dateMatch) {
    const d = new Date(str);
    if (!Number.isNaN(d.getTime())) {
      const pad = (n: number) => String(n).padStart(2, "0");
      const formatted = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(
        d.getHours()
      )}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
      return `#${formatted}#`;
    }
  }

  return `'${str.replace(/'/g, "''")}'`;
}

async function loadAdodb() {
  try {
    // Optional dependency: only present/usable on Windows. Imported lazily so
    // the rest of the backend runs fine on Linux/macOS dev machines without it.
    const mod = await import("node-adodb" as string);
    return (mod as { default?: typeof mod }).default ?? mod;
  } catch {
    throw new Error(
      "node-adodb 모듈을 불러올 수 없습니다. Windows + ACE OLEDB 드라이버 환경에서만 Access 동기화가 가능합니다."
    );
  }
}

export interface AccessSyncResult {
  tableKey: string;
  ok: boolean;
  message: string;
}

export async function syncTableToAccess(tableKey: string): Promise<AccessSyncResult> {
  if (!accessSyncAvailable()) {
    return {
      tableKey,
      ok: false,
      message: "이 서버 환경에서는 Access 동기화를 사용할 수 없습니다 (Windows 전용 기능).",
    };
  }

  const tableDef = getTableDef(tableKey);
  if (!tableDef) {
    return { tableKey, ok: false, message: `알 수 없는 표: ${tableKey}` };
  }

  const columnDefs = getColumnDefs(tableKey);
  const physical = dataTableName(tableKey);
  const rows = db
    .prepare(`SELECT * FROM ${quoteIdent(physical)} ORDER BY _id ASC`)
    .all() as Record<string, unknown>[];

  try {
    const ADODB = await loadAdodb();
    const connection = ADODB.open(
      `Provider=Microsoft.ACE.OLEDB.12.0;Data Source=${config.accessDbPath};Persist Security Info=False;`
    );

    await connection.execute(`DELETE FROM [${tableKey}]`);

    for (const row of rows) {
      const fieldList = columnDefs.map((c) => `[${c.col_key}]`).join(",");
      const valueList = columnDefs.map((c) => toSqlValue(row[c.col_key])).join(",");
      await connection.execute(`INSERT INTO [${tableKey}] (${fieldList}) VALUES (${valueList})`);
    }

    return { tableKey, ok: true, message: `${rows.length}건 전송 완료` };
  } catch (err) {
    return { tableKey, ok: false, message: (err as Error).message };
  }
}
