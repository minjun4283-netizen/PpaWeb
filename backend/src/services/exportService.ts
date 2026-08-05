import ExcelJS from "exceljs";
import { db, dataTableName, getColumnDefs, listTableDefs, quoteIdent } from "../db.js";

export type ExportType = "added" | "modified";

function collectPks(tableKey: string, since: string): { added: Set<string>; modified: Set<string> } {
  const entries = db
    .prepare(
      `SELECT change_type, pk_value, changed_at FROM change_log
       WHERE table_key = ? AND changed_at > ? ORDER BY id ASC`
    )
    .all(tableKey, since) as { change_type: string; pk_value: string; changed_at: string }[];

  const added = new Set<string>();
  const modified = new Set<string>();
  const deleted = new Set<string>();

  for (const entry of entries) {
    if (entry.change_type === "추가") added.add(entry.pk_value);
    else if (entry.change_type === "수정") {
      if (!added.has(entry.pk_value)) modified.add(entry.pk_value);
    } else if (entry.change_type === "삭제") {
      added.delete(entry.pk_value);
      modified.delete(entry.pk_value);
      deleted.add(entry.pk_value);
    }
  }

  return { added, modified };
}

export async function buildExportWorkbook(type: ExportType, since: string): Promise<ExcelJS.Workbook> {
  const workbook = new ExcelJS.Workbook();

  for (const tableDef of listTableDefs()) {
    const columnDefs = getColumnDefs(tableDef.table_key);
    const sheet = workbook.addWorksheet(tableDef.table_key);

    sheet.columns = columnDefs.map((c) => ({ header: c.label, key: c.col_key, width: 18 }));
    sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
    sheet.getRow(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF5B9BD5" },
    };

    const { added, modified } = collectPks(tableDef.table_key, since);
    const targetPks = type === "added" ? added : modified;
    if (targetPks.size === 0) continue;

    const physical = dataTableName(tableDef.table_key);
    const rows = db
      .prepare(`SELECT * FROM ${quoteIdent(physical)} ORDER BY _id ASC`)
      .all() as Record<string, unknown>[];

    for (const row of rows) {
      const pkValue = String(row[tableDef.pk_column] ?? "");
      if (!targetPks.has(pkValue)) continue;
      const record: Record<string, unknown> = {};
      for (const col of columnDefs) record[col.col_key] = row[col.col_key];
      sheet.addRow(record);
    }
  }

  return workbook;
}
