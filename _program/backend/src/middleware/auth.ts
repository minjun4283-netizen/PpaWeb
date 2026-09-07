import type { NextFunction, Request, Response } from "express";
import { config } from "../config.js";
import { verifyToken, type AuthUser } from "../auth/authService.js";

declare global {
  // eslint-disable-next-line @typescript-eslint/no-namespace
  namespace Express {
    interface Request {
      user?: AuthUser;
    }
  }
}

export function requireAuth(req: Request, res: Response, next: NextFunction) {
  const token = req.cookies?.[config.cookieName];
  const user = token ? verifyToken(token) : null;
  if (!user) {
    res.status(401).json({ error: "로그인이 필요합니다." });
    return;
  }
  req.user = user;
  next();
}

export function requireAdmin(req: Request, res: Response, next: NextFunction) {
  if (req.user?.role !== "admin") {
    res.status(403).json({ error: "관리자 권한이 필요합니다." });
    return;
  }
  next();
}
