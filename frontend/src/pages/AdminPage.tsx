import { useEffect, useState } from "react";
import { api, ApiError, type ColumnType } from "../api";
import { useTables } from "../TablesContext";
import { useAuth } from "../AuthContext";

interface UserRow {
  id: number;
  username: string;
  display_name: string;
  role: "admin" | "user";
  created_at: string;
}

function UserManagementSection() {
  const { user: currentUser } = useAuth();
  const [users, setUsers] = useState<UserRow[]>([]);
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [displayName, setDisplayName] = useState("");
  const [role, setRole] = useState<"admin" | "user">("user");
  const [message, setMessage] = useState<string | null>(null);

  function load() {
    api.listUsers().then((res) => setUsers(res.users));
  }

  useEffect(() => {
    load();
  }, []);

  async function handleCreate(e: React.FormEvent) {
    e.preventDefault();
    setMessage(null);
    try {
      await api.createUser(username, password, displayName, role);
      setMessage(`"${displayName}" 계정을 만들었습니다.`);
      setUsername("");
      setPassword("");
      setDisplayName("");
      setRole("user");
      load();
    } catch (err) {
      setMessage(err instanceof ApiError ? err.message : "계정 생성 중 오류가 발생했습니다.");
    }
  }

  async function handleDelete(id: number) {
    if (!window.confirm("이 계정을 삭제할까요?")) return;
    await api.deleteUser(id);
    load();
  }

  return (
    <section className="admin-section">
      <h3>팀원 계정 관리</h3>
      <p className="hint">여기서 만든 아이디/비밀번호로 팀원들이 로그인합니다.</p>

      <table className="mini-table user-table">
        <thead>
          <tr>
            <th>아이디</th>
            <th>이름</th>
            <th>권한</th>
            <th></th>
          </tr>
        </thead>
        <tbody>
          {users.map((u) => (
            <tr key={u.id}>
              <td>{u.username}</td>
              <td>{u.display_name}</td>
              <td>{u.role === "admin" ? "관리자" : "일반"}</td>
              <td>
                {u.id !== currentUser?.id && (
                  <button className="danger-btn" onClick={() => handleDelete(u.id)}>
                    삭제
                  </button>
                )}
              </td>
            </tr>
          ))}
        </tbody>
      </table>

      <form className="add-column-form" onSubmit={handleCreate}>
        <label>
          아이디
          <input value={username} onChange={(e) => setUsername(e.target.value)} required />
        </label>
        <label>
          초기 비밀번호
          <input
            type="text"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            required
            minLength={4}
          />
        </label>
        <label>
          이름
          <input value={displayName} onChange={(e) => setDisplayName(e.target.value)} required />
        </label>
        <label>
          권한
          <select value={role} onChange={(e) => setRole(e.target.value as "admin" | "user")}>
            <option value="user">일반</option>
            <option value="admin">관리자</option>
          </select>
        </label>
        <button type="submit">계정 추가</button>
      </form>
      {message && <div className="info-banner">{message}</div>}
    </section>
  );
}

export function AdminPage() {
  const { tables, refresh } = useTables();
  const [tableKey, setTableKey] = useState("");
  const [colKey, setColKey] = useState("");
  const [label, setLabel] = useState("");
  const [type, setType] = useState<ColumnType>("text");
  const [message, setMessage] = useState<string | null>(null);

  const [accessAvailable, setAccessAvailable] = useState(false);
  const [syncResults, setSyncResults] = useState<{ tableKey: string; ok: boolean; message: string }[]>(
    []
  );
  const [syncing, setSyncing] = useState<string | null>(null);
  const [syncLog, setSyncLog] = useState<
    { id: number; synced_at: string; username: string; table_key: string; status: string; message: string }[]
  >([]);

  useEffect(() => {
    if (tables.length > 0 && !tableKey) setTableKey(tables[0].table_key);
  }, [tables, tableKey]);

  useEffect(() => {
    api.accessSyncStatus().then((r) => setAccessAvailable(r.available));
    loadSyncLog();
  }, []);

  function loadSyncLog() {
    api.accessSyncLog().then((r) => setSyncLog(r.entries));
  }

  async function handleAddColumn(e: React.FormEvent) {
    e.preventDefault();
    setMessage(null);
    try {
      await api.addColumn(tableKey, colKey, label, type);
      setMessage(`"${label}" 컬럼을 추가했습니다.`);
      setColKey("");
      setLabel("");
      await refresh();
    } catch (err) {
      setMessage(err instanceof ApiError ? err.message : "컬럼 추가 중 오류가 발생했습니다.");
    }
  }

  async function handleSyncAll() {
    setSyncing("__all__");
    try {
      const res = await api.accessSyncAll();
      setSyncResults(res.results);
      loadSyncLog();
    } finally {
      setSyncing(null);
    }
  }

  async function handleSyncOne(tk: string) {
    setSyncing(tk);
    try {
      const result = await api.accessSyncTable(tk);
      setSyncResults((prev) => [result, ...prev.filter((r) => r.tableKey !== tk)]);
      loadSyncLog();
    } finally {
      setSyncing(null);
    }
  }

  return (
    <div className="admin-page">
      <h2>관리자</h2>

      <UserManagementSection />

      <section className="admin-section">
        <h3>컬럼관리</h3>
        <p className="hint">
          초기 스키마는 추측으로 채워졌습니다. 실제 엑셀 표에 있는 컬럼을 여기서 추가해 맞춰주세요.
        </p>
        <form className="add-column-form" onSubmit={handleAddColumn}>
          <label>
            표
            <select value={tableKey} onChange={(e) => setTableKey(e.target.value)}>
              {tables.map((t) => (
                <option key={t.table_key} value={t.table_key}>
                  {t.label}
                </option>
              ))}
            </select>
          </label>
          <label>
            컬럼 키 (영문/숫자/한글, 공백 불가)
            <input value={colKey} onChange={(e) => setColKey(e.target.value)} required />
          </label>
          <label>
            표시 이름
            <input value={label} onChange={(e) => setLabel(e.target.value)} required />
          </label>
          <label>
            유형
            <select value={type} onChange={(e) => setType(e.target.value as ColumnType)}>
              <option value="text">텍스트</option>
              <option value="number">숫자</option>
              <option value="date">날짜</option>
            </select>
          </label>
          <button type="submit">컬럼 추가</button>
        </form>
        {message && <div className="info-banner">{message}</div>}
      </section>

      <section className="admin-section">
        <h3>Access DB 동기화</h3>
        {!accessAvailable && (
          <div className="warn-banner">
            이 서버는 Windows + ACE OLEDB 드라이버 환경이 아니어서 Access 동기화를 사용할 수 없습니다.
            운영 시에는 이 기능을 Access 파일을 볼 수 있는 Windows 서버(또는 예약 작업)에서 실행하세요.
          </div>
        )}
        <div className="sync-buttons">
          <button onClick={handleSyncAll} disabled={!accessAvailable || syncing !== null}>
            {syncing === "__all__" ? "전송 중..." : "전체 표 Access로 전송"}
          </button>
          {tables.map((t) => (
            <button
              key={t.table_key}
              className="ghost-btn"
              onClick={() => handleSyncOne(t.table_key)}
              disabled={!accessAvailable || syncing !== null}
            >
              {syncing === t.table_key ? "전송 중..." : `${t.label} 전송`}
            </button>
          ))}
        </div>
        {syncResults.length > 0 && (
          <ul className="sync-results">
            {syncResults.map((r) => (
              <li key={r.tableKey} className={r.ok ? "sync-ok" : "sync-fail"}>
                {r.tableKey}: {r.message}
              </li>
            ))}
          </ul>
        )}

        <h4>최근 동기화 이력</h4>
        {syncLog.length === 0 ? (
          <p className="hint">아직 동기화 이력이 없습니다.</p>
        ) : (
          <table className="mini-table">
            <thead>
              <tr>
                <th>시각</th>
                <th>사용자</th>
                <th>표</th>
                <th>결과</th>
                <th>메시지</th>
              </tr>
            </thead>
            <tbody>
              {syncLog.map((entry) => (
                <tr key={entry.id}>
                  <td>{new Date(entry.synced_at.replace(" ", "T") + "Z").toLocaleString()}</td>
                  <td>{entry.username}</td>
                  <td>{entry.table_key}</td>
                  <td className={entry.status === "성공" ? "sync-ok" : "sync-fail"}>{entry.status}</td>
                  <td>{entry.message}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </section>
    </div>
  );
}
