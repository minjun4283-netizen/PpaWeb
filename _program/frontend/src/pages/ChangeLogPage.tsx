import { useEffect, useState } from "react";
import { api, type ChangeLogEntry } from "../api";
import { useTables } from "../TablesContext";

function formatData(raw: string | null): string {
  if (!raw) return "";
  try {
    const obj = JSON.parse(raw) as Record<string, unknown>;
    return Object.entries(obj)
      .filter(([k]) => !k.startsWith("_"))
      .map(([k, v]) => `${k}=${v ?? ""}`)
      .join("; ");
  } catch {
    return raw;
  }
}

function toIso(localDatetime: string): string | undefined {
  if (!localDatetime) return undefined;
  const d = new Date(localDatetime);
  return Number.isNaN(d.getTime()) ? undefined : d.toISOString();
}

export function ChangeLogPage() {
  const { tables } = useTables();
  const [tableKey, setTableKey] = useState("");
  const [since, setSince] = useState("");
  const [until, setUntil] = useState("");
  const [entries, setEntries] = useState<ChangeLogEntry[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    setLoading(true);
    api
      .listChangeLog({ tableKey: tableKey || undefined, since: toIso(since), until: toIso(until) })
      .then((res) => setEntries(res.entries))
      .finally(() => setLoading(false));
  }, [tableKey, since, until]);

  function labelOf(key: string): string {
    return tables.find((t) => t.table_key === key)?.label ?? key;
  }

  return (
    <div className="change-log-page">
      <div className="table-page-header">
        <h2>변경이력</h2>
        <div className="table-page-actions changelog-filters">
          <select value={tableKey} onChange={(e) => setTableKey(e.target.value)}>
            <option value="">전체 표</option>
            {tables.map((t) => (
              <option key={t.table_key} value={t.table_key}>
                {t.label}
              </option>
            ))}
          </select>
          <label className="inline-label">
            시작
            <input type="datetime-local" value={since} onChange={(e) => setSince(e.target.value)} />
          </label>
          <label className="inline-label">
            종료
            <input type="datetime-local" value={until} onChange={(e) => setUntil(e.target.value)} />
          </label>
          {(since || until) && (
            <button
              className="ghost-btn"
              onClick={() => {
                setSince("");
                setUntil("");
              }}
            >
              기간 초기화
            </button>
          )}
        </div>
      </div>

      {loading ? (
        <p>불러오는 중...</p>
      ) : (
        <div className="grid-scroll">
          <table className="data-grid">
            <thead>
              <tr>
                <th>변경시각</th>
                <th>사용자</th>
                <th>표</th>
                <th>PK값</th>
                <th>변경유형</th>
                <th>설명</th>
                <th>이전값</th>
                <th>현재값</th>
              </tr>
            </thead>
            <tbody>
              {entries.length === 0 && (
                <tr>
                  <td colSpan={8} className="empty-row">
                    변경 이력이 없습니다.
                  </td>
                </tr>
              )}
              {entries.map((e) => (
                <tr key={e.id}>
                  <td>{new Date(e.changed_at.replace(" ", "T") + "Z").toLocaleString()}</td>
                  <td>{e.username}</td>
                  <td>{labelOf(e.table_key)}</td>
                  <td>{e.pk_value}</td>
                  <td>
                    <span className={`change-badge change-${e.change_type}`}>{e.change_type}</span>
                  </td>
                  <td>{e.description}</td>
                  <td className="wrap-cell">{formatData(e.old_data)}</td>
                  <td className="wrap-cell">{formatData(e.new_data)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
