import { Router } from "express";
import { z } from "zod";
import { db } from "../db.js";
import { createUser, findUserByUsername } from "../auth/authService.js";
import { requireAdmin, requireAuth } from "../middleware/auth.js";

export const usersRouter = Router();
usersRouter.use(requireAuth, requireAdmin);

usersRouter.get("/", (_req, res) => {
  const users = db
    .prepare(`SELECT id, username, display_name, role, created_at FROM users ORDER BY id ASC`)
    .all();
  res.json({ users });
});

const createUserSchema = z.object({
  username: z.string().min(1).max(64),
  password: z.string().min(4).max(200),
  displayName: z.string().min(1).max(100),
  role: z.enum(["admin", "user"]),
});

usersRouter.post("/", (req, res) => {
  const parsed = createUserSchema.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: "입력값을 확인해주세요.", details: parsed.error.flatten() });
    return;
  }

  if (findUserByUsername(parsed.data.username)) {
    res.status(409).json({ error: "이미 존재하는 아이디입니다." });
    return;
  }

  const user = createUser(
    parsed.data.username,
    parsed.data.password,
    parsed.data.displayName,
    parsed.data.role
  );
  res.status(201).json({ user });
});

usersRouter.delete("/:id", (req, res) => {
  const id = Number(req.params.id);
  if (id === req.user!.id) {
    res.status(400).json({ error: "본인 계정은 삭제할 수 없습니다." });
    return;
  }
  db.prepare(`DELETE FROM users WHERE id = ?`).run(id);
  res.json({ ok: true });
});
