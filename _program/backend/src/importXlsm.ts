// One-off data migration: load the real PPA xlsm workbook into this app's
// SQLite database. Meant to be run once (or re-run with --replace) directly
// on the intranet host/container that can see the file — this backend has
// no network path to the corporate VDI where the source workbook lives, so
// there is no "fetch it from a URL" option here.
//
// Usage:
//   npm run import -- --file=/path/to/PPA파일.xlsm [--replace] [--dry-run]
//
// --replace   clears each table's existing rows before inserting the sheet's
//             rows (otherwise rows are appended, which is safe to run
//             against an already-seeded/empty table but would create
//             duplicate PKs if run twice against the same loaded data).
// --dry-run   only reports what would be imported, does not write anything.
import path from "node:path";
import ExcelJS from "exceljs";
import { db, dataTableName, getColumnDefs, initDatabase, listTableDefs, quoteIdent } from "./db.js";

interface ImportOptions {
  file?: string;
  replace: boolean;
  dryRun: boolean;
}

function parseArgs(argv: string[]): ImportOptions {
  const opts: ImportOptions = { replace: false, dryRun: false };
  for (const arg of argv) {
    if (arg.startsWith("--file=")) opts.file = arg.slice("--file=".length);
    else if (arg === "--replace") opts.replace = true;
    else if (arg === "--dry-run") opts.dryRun = true;
  }
  return opts;
}

// Cell values from exceljs come back in several shapes depending on
// formatting/content: primitives, Date objects (for date-formatted cells),
// rich-text runs, hyperlink objects, or formula results. Flatten all of
// those down to a plain string/number/null the DB can store.
function extractCellValue(raw: ExcelJS.CellValue): string | number | null {
  if (raw === null || raw === undefined) return null;

  if (raw instanceof Date) {
    // exceljs builds date cells straight from the serial number using UTC
    // fields, with no timezone attached — read them back the same way so a
    // local-timezone conversion doesn't shift the calendar day by one.
    const y = raw.getUTCFullYear();
    const m = String(raw.getUTCMonth() + 1).padStart(2, "0");
    const d = String(raw.getUTCDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }

  if (typeof raw === "boolean") return raw ? "TRUE" : "FALSE";
  if (typeof raw === "number") return raw;
  if (typeof raw === "string") return raw.trim();

  if (typeof raw === "object") {
    const obj = raw as unknown as Record<string, unknown>;
    if ("result" in obj) return extractCellValue(obj.result as ExcelJS.CellValue);
    if (typeof obj.text === "string") return obj.text.trim();
    if (Array.isArray(obj.richText)) {
      return (obj.richText as { text?: string }[]).map((rt) => rt.text ?? "").join("").trim();
    }
  }

  return String(raw).trim();
}

async function main() {
  const opts = parseArgs(process.argv.slice(2));
  if (!opts.file) {
    console.error(
      "사용법: npm run import -- --file=/path/to/엑셀파일.xlsm [--replace] [--dry-run]"
    );
    process.exit(1);
  }

  initDatabase();

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(path.resolve(opts.file));

  for (const tableDef of listTableDefs()) {
    const sheet = workbook.getWorksheet(tableDef.table_key);
    if (!sheet) {
      console.warn(`[건너뜀] 시트를 찾을 수 없음: ${tableDef.table_key}`);
      continue;
    }

    const columnDefs = getColumnDefs(tableDef.table_key);
    const labelToKey = new Map(columnDefs.map((c) => [c.label, c.col_key]));

    const colIndexToKey = new Map<number, string>();
    const unmatchedHeaders: string[] = [];

    sheet.getRow(1).eachCell((cell, colNumber) => {
      const label = extractCellValue(cell.value);
      const labelStr = label == null ? "" : String(label).trim();
      if (!labelStr) return;
      const key = labelToKey.get(labelStr);
      if (key) colIndexToKey.set(colNumber, key);
      else unmatchedHeaders.push(labelStr);
    });

    if (colIndexToKey.size === 0) {
      console.warn(
        `[건너뜀] ${tableDef.table_key}: 인식 가능한 컬럼이 없습니다. (헤더: ${unmatchedHeaders.join(", ")})`
      );
      continue;
    }

    const rows: Record<string, unknown>[] = [];
    let skippedBlank = 0;

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const record: Record<string, unknown> = {};
      let hasAny = false;
      for (const [colNumber, key] of colIndexToKey) {
        const value = extractCellValue(row.getCell(colNumber).value);
        record[key] = value;
        if (value !== null && value !== "") hasAny = true;
      }
      if (!hasAny) {
        skippedBlank++;
        return;
      }
      rows.push(record);
    });

    console.log(
      `${tableDef.table_key}: ${rows.length}행 인식 (공란행 ${skippedBlank}건 제외).` +
        (unmatchedHeaders.length > 0
          ? ` 인식 안 된 헤더(검증열 등, 정상): ${unmatchedHeaders.join(", ")}`
          : "")
    );

    if (opts.dryRun) continue;

    const physical = dataTableName(tableDef.table_key);
    const insertColumns = [...colIndexToKey.values()];
    const insertStmt = db.prepare(
      `INSERT INTO ${quoteIdent(physical)} (_created_by, _updated_by, ${insertColumns
        .map((c) => quoteIdent(c))
        .join(", ")})
       VALUES (?, ?, ${insertColumns.map(() => "?").join(", ")})`
    );

    let imported = 0;
    const runImport = db.transaction(() => {
      if (opts.replace) {
        db.exec(`DELETE FROM ${quoteIdent(physical)}`);
      }
      for (const record of rows) {
        const values = insertColumns.map((c) => record[c] ?? null);
        insertStmt.run("excel-import", "excel-import", ...values);
        imported++;
      }
    });
    runImport();

    db.prepare(
      `INSERT INTO change_log (username, table_key, row_id, pk_value, change_type, old_data, new_data, description)
       VALUES ('excel-import', ?, NULL, '(일괄 가져오기)', '추가', NULL, NULL, ?)`
    ).run(
      tableDef.table_key,
      `엑셀에서 ${imported}건 가져옴${opts.replace ? " (기존 데이터 교체)" : ""}`
    );

    console.log(`  → ${imported}건 저장 완료${opts.replace ? " (기존 데이터 교체)" : ""}`);
  }

  console.log(opts.dryRun ? "확인 완료 (--dry-run, 실제로 저장되지 않았습니다)." : "가져오기 완료.");
}

main().catch((err) => {
  console.error("가져오기 실패:", err);
  process.exit(1);
});
