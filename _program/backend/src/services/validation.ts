import { db, dataTableName, getColumnDefs, getTableDef, getUniqueGroups, listTableDefs, quoteIdent } from "../db.js";

export interface ValidationError {
  tableKey: string;
  rowId: number;
  pkValue: string;
  errorItem: string;
  result: "오류";
}

export interface ValidationReport {
  errors: ValidationError[];
  totalErrors: number;
  runAt: string;
  summaryByErrorItem: { key: string; count: number }[];
  summaryByTable: { key: string; count: number }[];
}

function isBlank(value: unknown): boolean {
  return value === null || value === undefined || String(value).trim() === "";
}

function validateTable(tableKey: string): ValidationError[] {
  const tableDef = getTableDef(tableKey)!;
  const columnDefs = getColumnDefs(tableKey);
  const fkColumns = columnDefs.filter((c) => c.is_fk === 1);
  const uniqueGroups = getUniqueGroups(tableKey);
  const physical = dataTableName(tableKey);

  const rows = db
    .prepare(`SELECT * FROM ${quoteIdent(physical)} ORDER BY _id ASC`)
    .all() as Record<string, unknown>[];

  const errors: ValidationError[] = [];
  const pushError = (row: Record<string, unknown>, errorItem: string) => {
    errors.push({
      tableKey,
      rowId: Number(row._id),
      pkValue: String(row[tableDef.pk_column] ?? ""),
      errorItem,
      result: "오류",
    });
  };

  // PK 공란 / PK 중복 (ValidateBlank + ValidateDuplicate on the PK column)
  const pkCounts = new Map<string, number>();
  for (const row of rows) {
    const v = String(row[tableDef.pk_column] ?? "").trim();
    if (v !== "") pkCounts.set(v, (pkCounts.get(v) ?? 0) + 1);
  }
  for (const row of rows) {
    const v = String(row[tableDef.pk_column] ?? "").trim();
    if (v === "") {
      pushError(row, "PK 공란");
    } else if ((pkCounts.get(v) ?? 0) > 1) {
      pushError(row, "PK 중복");
    }
  }

  // FK 공란 / FK 참조 (ValidateBlank + ValidateReference per FK column)
  for (const fk of fkColumns) {
    const refPhysical = dataTableName(fk.ref_table!);
    const refValues = new Set(
      (
        db
          .prepare(`SELECT ${quoteIdent(fk.ref_column!)} as v FROM ${quoteIdent(refPhysical)}`)
          .all() as { v: unknown }[]
      )
        .map((r) => String(r.v ?? "").trim())
        .filter((v) => v !== "")
    );

    for (const row of rows) {
      const v = String(row[fk.col_key] ?? "").trim();
      if (v === "") {
        pushError(row, `${fk.col_key} 공란`);
      } else if (!refValues.has(v)) {
        pushError(row, `${fk.col_key} 참조`);
      }
    }
  }

  // 조합중복 (ValidateCombinationDuplicate)
  for (const group of uniqueGroups) {
    const comboCounts = new Map<string, number>();
    for (const row of rows) {
      const parts = group.map((col) => String(row[col] ?? "").trim());
      const key = parts.join("|");
      if (parts.some((p) => p !== "")) {
        comboCounts.set(key, (comboCounts.get(key) ?? 0) + 1);
      }
    }
    for (const row of rows) {
      const parts = group.map((col) => String(row[col] ?? "").trim());
      const key = parts.join("|");
      if (parts.some((p) => p !== "") && (comboCounts.get(key) ?? 0) > 1) {
        pushError(row, "조합중복");
      }
    }
  }

  return errors;
}

export function runValidation(): ValidationReport {
  const allErrors: ValidationError[] = [];
  for (const table of listTableDefs()) {
    allErrors.push(...validateTable(table.table_key));
  }

  const byErrorItem = new Map<string, number>();
  const byTable = new Map<string, number>();
  for (const err of allErrors) {
    byErrorItem.set(err.errorItem, (byErrorItem.get(err.errorItem) ?? 0) + 1);
    byTable.set(err.tableKey, (byTable.get(err.tableKey) ?? 0) + 1);
  }

  return {
    errors: allErrors,
    totalErrors: allErrors.length,
    runAt: new Date().toISOString(),
    summaryByErrorItem: [...byErrorItem.entries()].map(([key, count]) => ({ key, count })),
    summaryByTable: [...byTable.entries()].map(([key, count]) => ({ key, count })),
  };
}
