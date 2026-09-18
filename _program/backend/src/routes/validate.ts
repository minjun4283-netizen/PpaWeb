import { Router } from "express";
import { requireAuth } from "../middleware/auth.js";
import { runValidation } from "../services/validation.js";

export const validateRouter = Router();
validateRouter.use(requireAuth);

validateRouter.post("/", (_req, res) => {
  res.json(runValidation());
});

validateRouter.get("/", (_req, res) => {
  res.json(runValidation());
});
