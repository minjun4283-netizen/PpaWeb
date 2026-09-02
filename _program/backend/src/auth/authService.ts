import bcrypt from "bcryptjs";
import jwt from "jsonwebtoken";
import { db } from "../db.js";
import { config } from "../config.js";

export type Role = "admin" | "user";

export interface UserRecord {
  id: number;
  username: string;
  password_hash: string;
  display_name: string;
  role: Role;
}

export interface AuthUser {
  id: number;
  username: string;
  displayName: string;
  role: Role;
}

export function findUserByUsername(username: string): UserRecord | undefined {
  return db.prepare(`SELECT * FROM users WHERE username = ?`).get(username) as
    | UserRecord
    | undefined;
}

export function createUser(
  username: string,
  password: string,
  displayName: string,
  role: Role
): AuthUser {
  const passwordHash = bcrypt.hashSync(password, 10);
  const info = db
    .prepare(
      `INSERT INTO users (username, password_hash, display_name, role) VALUES (?, ?, ?, ?)`
    )
    .run(username, passwordHash, displayName, role);
  return { id: Number(info.lastInsertRowid), username, displayName, role };
}

export function verifyPassword(user: UserRecord, password: string): boolean {
  return bcrypt.compareSync(password, user.password_hash);
}

export function countUsers(): number {
  return (db.prepare(`SELECT COUNT(*) as c FROM users`).get() as { c: number }).c;
}

export function toAuthUser(user: UserRecord): AuthUser {
  return { id: user.id, username: user.username, displayName: user.display_name, role: user.role };
}

export function signToken(user: AuthUser): string {
  return jwt.sign(user, config.jwtSecret, { expiresIn: "12h" });
}

export function verifyToken(token: string): AuthUser | null {
  try {
    return jwt.verify(token, config.jwtSecret) as AuthUser;
  } catch {
    return null;
  }
}
