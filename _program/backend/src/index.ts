import "dotenv/config";
import path from "node:path";
import { fileURLToPath } from "node:url";
import fs from "node:fs";
import cookieParser from "cookie-parser";
import cors from "cors";
import express from "express";
import { config } from "./config.js";
import { initDatabase } from "./db.js";
import { countUsers, createUser } from "./auth/authService.js";
import { authRouter } from "./routes/auth.js";
import { metaRouter } from "./routes/meta.js";
import { tablesRouter } from "./routes/tables.js";
import { validateRouter } from "./routes/validate.js";
import { changeLogRouter } from "./routes/changelog.js";
import { exportRouter } from "./routes/export.js";
import { accessSyncRouter } from "./routes/accessSync.js";
import { tooltipRouter } from "./routes/tooltip.js";
import { usersRouter } from "./routes/users.js";

initDatabase();

if (countUsers() === 0) {
  const password = process.env.ADMIN_PASSWORD ?? "changeme123";
  createUser("admin", password, "관리자", "admin");
  // eslint-disable-next-line no-console
  console.warn(
    `[초기 설정] 관리자 계정을 생성했습니다. 아이디: admin / 비밀번호: ${password}\n` +
      `로그인 후 반드시 비밀번호를 변경하세요 (또는 ADMIN_PASSWORD 환경변수로 최초 비밀번호를 지정할 수 있습니다).`
  );
}

const app = express();
app.use(express.json());
app.use(cookieParser());
app.use(
  cors({
    origin: process.env.CORS_ORIGIN?.split(",") ?? true,
    credentials: true,
  })
);

app.use("/api/auth", authRouter);
app.use("/api/meta", metaRouter);
app.use("/api/tables", tablesRouter);
app.use("/api/validate", validateRouter);
app.use("/api/change-log", changeLogRouter);
app.use("/api/export", exportRouter);
app.use("/api/access-sync", accessSyncRouter);
app.use("/api/tooltip", tooltipRouter);
app.use("/api/users", usersRouter);

app.get("/api/health", (_req, res) => res.json({ ok: true }));

// In production the frontend is pre-built and copied next to this file
// (see the Dockerfile) so a single container/process can serve both the API
// and the static site — no separate web server needed on the intranet host.
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const publicDir = path.join(__dirname, "public");
if (fs.existsSync(publicDir)) {
  app.use(express.static(publicDir));
  app.get(/^(?!\/api).*/, (_req, res) => {
    res.sendFile(path.join(publicDir, "index.html"));
  });
}

app.listen(config.port, () => {
  // eslint-disable-next-line no-console
  console.log(`PpaWeb backend listening on port ${config.port}`);
});
