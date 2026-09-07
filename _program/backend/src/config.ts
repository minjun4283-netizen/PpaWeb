import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

export const config = {
  port: Number(process.env.PORT ?? 4000),
  jwtSecret: process.env.JWT_SECRET ?? "change-me-in-production",
  dbPath: process.env.DB_PATH ?? path.join(__dirname, "..", "data", "ppaweb.db"),
  cookieName: "ppaweb_token",
  isProduction: process.env.NODE_ENV === "production",
  // Access DB sync target (only reachable on Windows hosts with the ACE OLEDB driver installed,
  // same requirement as the original VBA macro).
  accessDbPath: process.env.ACCESS_DB_PATH ?? "",
};
