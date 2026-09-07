import { Router } from "express";
import { db, listTableDefs } from "../db.js";
import { requireAdmin, requireAuth } from "../middleware/auth.js";
import { accessSyncAvailable, syncTableToAccess } from "../services/accessSync.js";

export const accessSyncRouter = Router();
accessSyncRouter.use(requireAuth);

function logResult(username: string, tableKey: string, ok: boolean, message: string) {
  db.prepare(
    `INSERT INTO access_sync_log (username, table_key, status, message) VALUES (?, ?, ?, ?)`
  ).run(username, tableKey, ok ? "성공" : "실패", message);
}

accessSyncRouter.get("/status", (_req, res) => {
  res.json({ available: accessSyncAvailable() });
});

accessSyncRouter.post("/:tableKey", requireAdmin, async (req, res) => {
  const result = await syncTableToAccess(req.params.tableKey);
  logResult(req.user!.username, result.tableKey, result.ok, result.message);
  res.json(result);
});

accessSyncRouter.post("/", requireAdmin, async (req, res) => {
  const results = [];
  for (const table of listTableDefs()) {
    const result = await syncTableToAccess(table.table_key);
    logResult(req.user!.username, result.tableKey, result.ok, result.message);
    results.push(result);
  }
  res.json({ results });
});

accessSyncRouter.get("/log", (_req, res) => {
  const entries = db
    .prepare(`SELECT * FROM access_sync_log ORDER BY id DESC LIMIT 200`)
    .all();
  res.json({ entries });
});
