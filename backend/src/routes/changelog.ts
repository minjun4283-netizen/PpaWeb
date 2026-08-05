import { Router } from "express";
import { requireAuth } from "../middleware/auth.js";
import { listChangeLog } from "../services/changeLog.js";

export const changeLogRouter = Router();
changeLogRouter.use(requireAuth);

changeLogRouter.get("/", (req, res) => {
  const tableKey = typeof req.query.tableKey === "string" ? req.query.tableKey : undefined;
  const limit = req.query.limit ? Number(req.query.limit) : undefined;
  res.json({ entries: listChangeLog({ tableKey, limit }) });
});
