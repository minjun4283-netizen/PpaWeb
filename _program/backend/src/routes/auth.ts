import { Router } from "express";
import { config } from "../config.js";
import { findUserByUsername, signToken, toAuthUser, verifyPassword } from "../auth/authService.js";
import { requireAuth } from "../middleware/auth.js";

export const authRouter = Router();

const cookieOptions = {
  httpOnly: true,
  sameSite: "lax" as const,
  secure: config.isProduction,
  maxAge: 12 * 60 * 60 * 1000,
};

authRouter.post("/login", (req, res) => {
  const { username, password } = req.body ?? {};
  if (typeof username !== "string" || typeof password !== "string") {
    res.status(400).json({ error: "아이디와 비밀번호를 입력해주세요." });
    return;
  }

  const user = findUserByUsername(username);
  if (!user || !verifyPassword(user, password)) {
    res.status(401).json({ error: "아이디 또는 비밀번호가 올바르지 않습니다." });
    return;
  }

  const authUser = toAuthUser(user);
  const token = signToken(authUser);
  res.cookie(config.cookieName, token, cookieOptions);
  res.json({ user: authUser });
});

authRouter.post("/logout", (_req, res) => {
  res.clearCookie(config.cookieName);
  res.json({ ok: true });
});

authRouter.get("/me", requireAuth, (req, res) => {
  res.json({ user: req.user });
});
