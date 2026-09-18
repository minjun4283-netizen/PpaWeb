import { useState } from "react";
import { NavLink, Outlet } from "react-router-dom";
import { useAuth } from "../AuthContext";
import { useTables } from "../TablesContext";

export function Layout() {
  const { user, logout } = useAuth();
  const { tables } = useTables();
  const [menuOpen, setMenuOpen] = useState(false);

  return (
    <div className="app-shell">
      <header className="topbar">
        <button className="menu-btn" onClick={() => setMenuOpen((v) => !v)} aria-label="메뉴 열기">
          ☰
        </button>
        <h1>PPA 계약관리</h1>
        <div className="topbar-user">
          <span>{user?.displayName}</span>
          <button className="link-btn" onClick={() => logout()}>
            로그아웃
          </button>
        </div>
      </header>

      <div className="app-body">
        <nav className={`sidebar ${menuOpen ? "open" : ""}`} onClick={() => setMenuOpen(false)}>
          <div className="nav-group-label">데이터 표</div>
          {tables.map((t) => (
            <NavLink key={t.table_key} to={`/tables/${t.table_key}`} className="nav-link">
              {t.label}
            </NavLink>
          ))}

          <div className="nav-group-label">보고서</div>
          <NavLink to="/validation" className="nav-link">
            검증 리포트
          </NavLink>
          <NavLink to="/change-log" className="nav-link">
            변경이력
          </NavLink>

          {user?.role === "admin" && (
            <>
              <div className="nav-group-label">관리자</div>
              <NavLink to="/admin" className="nav-link">
                컬럼관리 / Access 동기화
              </NavLink>
            </>
          )}
        </nav>

        <main className="content">
          <Outlet />
        </main>
      </div>
    </div>
  );
}
