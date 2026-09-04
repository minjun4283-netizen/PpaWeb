import { useState } from "react";
import { useAuth } from "../AuthContext";
import { ApiError } from "../api";

export function LoginPage() {
  const { login } = useAuth();
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    setError(null);
    setBusy(true);
    try {
      await login(username, password);
    } catch (err) {
      setError(err instanceof ApiError ? err.message : "로그인 중 오류가 발생했습니다.");
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="login-page">
      <form className="login-card" onSubmit={handleSubmit}>
        <h1>PPA 계약관리</h1>
        <p className="login-sub">사내망 전용 · 팀 협업 데이터 관리</p>

        <label>
          아이디
          <input
            value={username}
            onChange={(e) => setUsername(e.target.value)}
            autoCapitalize="off"
            autoCorrect="off"
            autoComplete="username"
            required
          />
        </label>

        <label>
          비밀번호
          <input
            type="password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            autoComplete="current-password"
            required
          />
        </label>

        {error && <div className="error-banner">{error}</div>}

        <button type="submit" disabled={busy}>
          {busy ? "로그인 중..." : "로그인"}
        </button>
      </form>
    </div>
  );
}
