import { Router } from "express";
import { requireAuth } from "../middleware/auth.js";
import { buildBuyContractTooltip, buildElectricUseSiteTooltip } from "../services/tooltip.js";

export const tooltipRouter = Router();
tooltipRouter.use(requireAuth);

tooltipRouter.get("/supply-match", (req, res) => {
  const { column, value } = req.query;
  if (typeof column !== "string" || typeof value !== "string") {
    res.status(400).json({ error: "column, value 쿼리 파라미터가 필요합니다." });
    return;
  }

  if (column === "전기사용지ID") {
    res.json({ fields: buildElectricUseSiteTooltip(value) });
  } else if (column === "구매계약ID") {
    res.json({ fields: buildBuyContractTooltip(value) });
  } else {
    res.json({ fields: [] });
  }
});
