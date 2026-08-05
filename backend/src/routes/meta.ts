import { Router } from "express";
import { z } from "zod";
import {
  addColumnDef,
  assertSafeIdentifier,
  getColumnDefs,
  getTableDef,
  getUniqueGroups,
  listTableDefs,
} from "../db.js";
import { requireAdmin, requireAuth } from "../middleware/auth.js";

export const metaRouter = Router();
metaRouter.use(requireAuth);

metaRouter.get("/tables", (_req, res) => {
  const tables = listTableDefs().map((t) => ({
    ...t,
    columns: getColumnDefs(t.table_key),
    uniqueGroups: getUniqueGroups(t.table_key),
  }));
  res.json({ tables });
});

const addColumnSchema = z.object({
  colKey: z.string().min(1).max(64),
  label: z.string().min(1).max(128),
  type: z.enum(["text", "number", "date"]),
});

metaRouter.post("/tables/:tableKey/columns", requireAdmin, (req, res) => {
  const { tableKey } = req.params;
  const tableDef = getTableDef(tableKey);
  if (!tableDef) {
    res.status(404).json({ error: `알 수 없는 표: ${tableKey}` });
    return;
  }

  const parsed = addColumnSchema.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: "입력값을 확인해주세요.", details: parsed.error.flatten() });
    return;
  }

  const existing = getColumnDefs(tableKey).some((c) => c.col_key === parsed.data.colKey);
  if (existing) {
    res.status(409).json({ error: "이미 존재하는 컬럼입니다." });
    return;
  }

  try {
    assertSafeIdentifier(parsed.data.colKey);
    addColumnDef(tableKey, parsed.data.colKey, parsed.data.label, parsed.data.type);
  } catch (err) {
    res.status(400).json({ error: (err as Error).message });
    return;
  }

  res.status(201).json({ columns: getColumnDefs(tableKey) });
});
