import { createContext, useCallback, useContext, useEffect, useState, type ReactNode } from "react";
import { api, type TableDef } from "./api";

interface TablesContextValue {
  tables: TableDef[];
  loading: boolean;
  refresh: () => Promise<void>;
}

const TablesContext = createContext<TablesContextValue | null>(null);

export function TablesProvider({ children }: { children: ReactNode }) {
  const [tables, setTables] = useState<TableDef[]>([]);
  const [loading, setLoading] = useState(true);

  const refresh = useCallback(async () => {
    const res = await api.listTables();
    setTables(res.tables);
  }, []);

  useEffect(() => {
    refresh().finally(() => setLoading(false));
  }, [refresh]);

  return (
    <TablesContext.Provider value={{ tables, loading, refresh }}>{children}</TablesContext.Provider>
  );
}

export function useTables(): TablesContextValue {
  const ctx = useContext(TablesContext);
  if (!ctx) throw new Error("useTables must be used inside TablesProvider");
  return ctx;
}
