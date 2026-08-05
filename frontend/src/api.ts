export type ColumnType = "text" | "number" | "date";

export interface ColumnDef {
  id: number;
  table_key: string;
  col_key: string;
  label: string;
  type: ColumnType;
  sort_order: number;
  is_fk: number;
  ref_table: string | null;
  ref_column: string | null;
}

export interface TableDef {
  table_key: string;
  label: string;
  pk_column: string;
  sort_order: number;
  columns: ColumnDef[];
  uniqueGroups: string[][];
}

export interface RowRecord {
  _id: number;
  _created_at: string;
  _updated_at: string;
  _created_by: string | null;
  _updated_by: string | null;
  [key: string]: unknown;
}

export interface AuthUser {
  id: number;
  username: string;
  displayName: string;
  role: "admin" | "user";
}

export interface ValidationError {
  tableKey: string;
  rowId: number;
  pkValue: string;
  errorItem: string;
  result: "오류";
}

export interface ValidationReport {
  errors: ValidationError[];
  totalErrors: number;
  runAt: string;
  summaryByErrorItem: { key: string; count: number }[];
  summaryByTable: { key: string; count: number }[];
}

export interface ChangeLogEntry {
  id: number;
  changed_at: string;
  username: string;
  table_key: string;
  row_id: number | null;
  pk_value: string;
  change_type: "추가" | "수정" | "삭제";
  old_data: string | null;
  new_data: string | null;
  description: string;
}

export interface TooltipField {
  label: string;
  value: string;
}

class ApiError extends Error {
  status: number;
  constructor(message: string, status: number) {
    super(message);
    this.status = status;
  }
}

async function request<T>(path: string, options: RequestInit = {}): Promise<T> {
  const res = await fetch(`/api${path}`, {
    credentials: "include",
    headers: { "Content-Type": "application/json", ...options.headers },
    ...options,
  });

  if (!res.ok) {
    let message = `요청 실패 (${res.status})`;
    try {
      const body = await res.json();
      if (body?.error) message = body.error;
    } catch {
      // ignore JSON parse failure
    }
    throw new ApiError(message, res.status);
  }

  if (res.status === 204) return undefined as T;
  return (await res.json()) as T;
}

export const api = {
  login: (username: string, password: string) =>
    request<{ user: AuthUser }>("/auth/login", {
      method: "POST",
      body: JSON.stringify({ username, password }),
    }),
  logout: () => request<{ ok: true }>("/auth/logout", { method: "POST" }),
  me: () => request<{ user: AuthUser }>("/auth/me"),

  listTables: () => request<{ tables: TableDef[] }>("/meta/tables"),
  addColumn: (tableKey: string, colKey: string, label: string, type: ColumnType) =>
    request<{ columns: ColumnDef[] }>(`/meta/tables/${encodeURIComponent(tableKey)}/columns`, {
      method: "POST",
      body: JSON.stringify({ colKey, label, type }),
    }),

  listRows: (tableKey: string) =>
    request<{ rows: RowRecord[] }>(`/tables/${encodeURIComponent(tableKey)}/rows`),
  createRow: (tableKey: string, data: Record<string, unknown>) =>
    request<{ row: RowRecord }>(`/tables/${encodeURIComponent(tableKey)}/rows`, {
      method: "POST",
      body: JSON.stringify(data),
    }),
  updateRow: (tableKey: string, id: number, data: Record<string, unknown>) =>
    request<{ row: RowRecord }>(`/tables/${encodeURIComponent(tableKey)}/rows/${id}`, {
      method: "PUT",
      body: JSON.stringify(data),
    }),
  deleteRow: (tableKey: string, id: number) =>
    request<{ ok: true }>(`/tables/${encodeURIComponent(tableKey)}/rows/${id}`, {
      method: "DELETE",
    }),

  runValidation: () => request<ValidationReport>("/validate", { method: "POST" }),

  listChangeLog: (filters: { tableKey?: string; since?: string; until?: string } = {}) => {
    const params = new URLSearchParams();
    if (filters.tableKey) params.set("tableKey", filters.tableKey);
    if (filters.since) params.set("since", filters.since);
    if (filters.until) params.set("until", filters.until);
    const qs = params.toString();
    return request<{ entries: ChangeLogEntry[] }>(`/change-log${qs ? `?${qs}` : ""}`);
  },

  tooltip: (column: string, value: string) =>
    request<{ fields: TooltipField[] }>(
      `/tooltip/supply-match?column=${encodeURIComponent(column)}&value=${encodeURIComponent(value)}`
    ),

  accessSyncStatus: () => request<{ available: boolean }>("/access-sync/status"),
  accessSyncAll: () => request<{ results: { tableKey: string; ok: boolean; message: string }[] }>(
    "/access-sync",
    { method: "POST" }
  ),
  accessSyncTable: (tableKey: string) =>
    request<{ tableKey: string; ok: boolean; message: string }>(
      `/access-sync/${encodeURIComponent(tableKey)}`,
      { method: "POST" }
    ),
  accessSyncLog: () =>
    request<{ entries: { id: number; synced_at: string; username: string; table_key: string; status: string; message: string }[] }>(
      "/access-sync/log"
    ),

  exportUrl: (type: "added" | "modified", since: string) =>
    `/api/export/${type}?since=${encodeURIComponent(since)}`,

  listUsers: () =>
    request<{ users: { id: number; username: string; display_name: string; role: "admin" | "user"; created_at: string }[] }>(
      "/users"
    ),
  createUser: (username: string, password: string, displayName: string, role: "admin" | "user") =>
    request<{ user: AuthUser }>("/users", {
      method: "POST",
      body: JSON.stringify({ username, password, displayName, role }),
    }),
  deleteUser: (id: number) => request<{ ok: true }>(`/users/${id}`, { method: "DELETE" }),
};

export { ApiError };
