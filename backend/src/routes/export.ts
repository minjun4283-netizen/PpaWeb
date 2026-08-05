import { Router } from "express";
import { requireAuth } from "../middleware/auth.js";
import { buildExportWorkbook, type ExportType } from "../services/exportService.js";

export const exportRouter = Router();
exportRouter.use(requireAuth);

exportRouter.get("/:type", async (req, res) => {
  const type = req.params.type as ExportType;
  if (type !== "added" && type !== "modified") {
    res.status(400).json({ error: "type은 added 또는 modified 여야 합니다." });
    return;
  }

  const since = typeof req.query.since === "string" ? req.query.since : "1970-01-01T00:00:00.000Z";
  const workbook = await buildExportWorkbook(type, since);

  const label = type === "added" ? "추가" : "수정";
  const ts = new Date().toISOString().replace(/[:.]/g, "-");
  const filename = `${label}_PPA_${ts}.xlsx`;
  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  // Korean filenames aren't valid raw Content-Disposition bytes; use the
  // RFC 5987 filename* form with an ASCII fallback.
  res.setHeader(
    "Content-Disposition",
    `attachment; filename="export.xlsx"; filename*=UTF-8''${encodeURIComponent(filename)}`
  );

  await workbook.xlsx.write(res);
  res.end();
});
