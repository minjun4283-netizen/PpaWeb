// change_log.changed_at is written via SQLite's datetime('now'), which
// produces "YYYY-MM-DD HH:MM:SS" (space-delimited, UTC, no milliseconds).
// Timestamps coming from the frontend are ISO 8601 ("...T...Z"). Comparing
// those two formats directly with plain string >/< is wrong: at the same
// instant, ' ' (0x20) sorts before 'T' (0x54), so same-calendar-day
// comparisons silently misbehave. Always convert incoming ISO timestamps
// through this before comparing against changed_at.
export function isoToSqliteUtc(iso: string): string {
  const d = new Date(iso);
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())} ${pad(
    d.getUTCHours()
  )}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}`;
}
