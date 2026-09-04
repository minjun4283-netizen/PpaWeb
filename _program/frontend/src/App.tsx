import { Navigate, Route, Routes } from "react-router-dom";
import { AuthProvider, useAuth } from "./AuthContext";
import { TablesProvider, useTables } from "./TablesContext";
import { ValidationProvider } from "./ValidationContext";
import { Layout } from "./components/Layout";
import { LoginPage } from "./pages/LoginPage";
import { TablePage } from "./pages/TablePage";
import { ValidationPage } from "./pages/ValidationPage";
import { ChangeLogPage } from "./pages/ChangeLogPage";
import { AdminPage } from "./pages/AdminPage";

function AuthedApp() {
  const { tables, loading } = useTables();

  if (loading) return <div className="full-page-loading">불러오는 중...</div>;

  return (
    <ValidationProvider>
      <Routes>
        <Route element={<Layout />}>
          <Route index element={<Navigate to={`/tables/${tables[0]?.table_key ?? ""}`} replace />} />
          <Route path="/tables/:tableKey" element={<TablePage />} />
          <Route path="/validation" element={<ValidationPage />} />
          <Route path="/change-log" element={<ChangeLogPage />} />
          <Route path="/admin" element={<AdminPage />} />
          <Route path="*" element={<Navigate to="/" replace />} />
        </Route>
      </Routes>
    </ValidationProvider>
  );
}

function Gate() {
  const { user, loading } = useAuth();

  if (loading) return <div className="full-page-loading">불러오는 중...</div>;
  if (!user) return <LoginPage />;

  return (
    <TablesProvider>
      <AuthedApp />
    </TablesProvider>
  );
}

export default function App() {
  return (
    <AuthProvider>
      <Gate />
    </AuthProvider>
  );
}
