import { Router } from "express";
import { requireAuth } from "../middleware/auth.js";
import { listChangeLog } from "../services/changeLog.js";
import { isoToSqliteUtc } from "../util/time.js";

export const changeLogRouter = Router();
changeLogRouter.use(requireAuth);

changeLogRouter.get("/", (req, res) => {
  const tableKey = typeof req.query.tableKey === "string" ? req.query.tableKey : undefined;
  const limit = req.query.limit ? Number(req.query.limit) : undefined;
  const since = typeof req.query.since === "string" ? isoToSqliteUtc(req.query.since) : undefined;
  const until = typeof req.query.until === "string" ? isoToSqliteUtc(req.query.until) : undefined;
  res.json({ entries: listChangeLog({ tableKey, since, until, limit }) });
});
