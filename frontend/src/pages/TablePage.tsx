import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { useParams } from "react-router-dom";
import { api, type ColumnDef, type RowRecord } from "../api";
import { useTables } from "../TablesContext";
import { useValidation } from "../ValidationContext";
import { InfoPopover } from "../components/InfoPopover";

const TOOLTIP_COLUMNS = new Set(["전기사용지ID", "구매계약ID"]);

interface Draft {
  clientId: string;
  data: Record<string, string>;
}

function emptyDraftData(columns: ColumnDef[]): Record<string, string> {
  const data: Record<string, string> = {};
  for (const c of columns) data[c.col_key] = "";
  return data;
}

export function TablePage() {
  const { tableKey = "" } = useParams();
  const { tables } = useTables();
  const { report } = useValidation();
  const tableDef = tables.find((t) => t.table_key === tableKey);

  const [rows, setRows] = useState<RowRecord[]>([]);
  const [drafts, setDrafts] = useState<Draft[]>([]);
  const [loading, setLoading] = useState(true);
  const [savedFlash, setSavedFlash] = useState<number | null>(null);
  const snapshots = useRef<Record<number, Record<string, unknown>>>({});

  const load = useCallback(async () => {
    if (!tableKey) return;
    setLoading(true);
    try {
      const res = await api.listRows(tableKey);
      setRows(res.rows);
      snapshots.current = Object.fromEntries(res.rows.map((r) => [r._id, { ...r }]));
    } finally {
      setLoading(false);
    }
  }, [tableKey]);

  useEffect(() => {
    load();
    setDrafts([]);
  }, [load]);

  const errorsByRowId = useMemo(() => {
    const map = new Map<number, string[]>();
    if (report) {
      for (const err of report.errors) {
        if (err.tableKey !== tableKey) continue;
        if (!map.has(err.rowId)) map.set(err.rowId, []);
        map.get(err.rowId)!.push(err.errorItem);
      }
    }
    return map;
  }, [report, tableKey]);

  if (!tableDef) return <p>불러오는 중...</p>;
  const columns = tableDef.columns;

  function updateRowLocal(rowId: number, colKey: string, value: string) {
    setRows((prev) => prev.map((r) => (r._id === rowId ? { ...r, [colKey]: value } : r)));
  }

  async function handleBlurSaved(rowId: number, colKey: string, value: string) {
    const before = snapshots.current[rowId]?.[colKey];
    if (String(before ?? "") === value) return;
    const res = await api.updateRow(tableKey, rowId, { [colKey]: value });
    snapshots.current[rowId] = { ...res.row };
    setRows((prev) => prev.map((r) => (r._id === rowId ? res.row : r)));
    setSavedFlash(rowId);
    setTimeout(() => setSavedFlash((cur) => (cur === rowId ? null : cur)), 900);
  }

  async function handleDelete(rowId: number) {
    if (!window.confirm("이 행을 삭제할까요? 되돌릴 수 없습니다.")) return;
    await api.deleteRow(tableKey, rowId);
    setRows((prev) => prev.filter((r) => r._id !== rowId));
    delete snapshots.current[rowId];
  }

  function addDraft() {
    setDrafts((prev) => [
      ...prev,
      { clientId: `draft-${Date.now()}-${Math.random()}`, data: emptyDraftData(columns) },
    ]);
  }

  function updateDraft(clientId: string, colKey: string, value: string) {
    setDrafts((prev) =>
      prev.map((d) => (d.clientId === clientId ? { ...d, data: { ...d.data, [colKey]: value } } : d))
    );
  }

  async function saveDraft(clientId: string) {
    const draft = drafts.find((d) => d.clientId === clientId);
    if (!draft) return;
    if (!draft.data[tableDef!.pk_column]?.trim()) {
      window.alert(`${tableDef!.pk_column} 값을 입력해주세요.`);
      return;
    }
    const res = await api.createRow(tableKey, draft.data);
    snapshots.current[res.row._id] = { ...res.row };
    setRows((prev) => [...prev, res.row]);
    setDrafts((prev) => prev.filter((d) => d.clientId !== clientId));
  }

  function discardDraft(clientId: string) {
    setDrafts((prev) => prev.filter((d) => d.clientId !== clientId));
  }

  function renderInput(
    col: ColumnDef,
    value: string,
    onChange: (v: string) => void,
    onBlur?: () => void
  ) {
    const inputType = col.type === "number" ? "number" : col.type === "date" ? "date" : "text";
    return (
      <input
        className="cell-input"
        type={inputType}
        step={col.type === "number" ? "any" : undefined}
        value={value ?? ""}
        onChange={(e) => onChange(e.target.value)}
        onBlur={onBlur}
      />
    );
  }

  return (
    <div className="table-page">
      <div className="table-page-header">
        <h2>{tableDef.label}</h2>
        <button onClick={addDraft}>+ 행 추가</button>
      </div>

      {loading ? (
        <p>불러오는 중...</p>
      ) : (
        <div className="grid-scroll">
          <table className="data-grid">
            <thead>
              <tr>
                {columns.map((c) => (
                  <th key={c.col_key}>
                    {c.label}
                    {c.col_key === tableDef.pk_column && <span className="pk-badge">PK</span>}
                    {c.is_fk === 1 && <span className="fk-badge">FK→{c.ref_table}</span>}
                  </th>
                ))}
                <th>작업</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((row) => {
                const errs = errorsByRowId.get(row._id);
                return (
                  <tr key={row._id} className={errs ? "row-error" : ""} title={errs?.join(", ")}>
                    {columns.map((col) => (
                      <td key={col.col_key}>
                        <div className="cell-with-info">
                          {renderInput(
                            col,
                            String(row[col.col_key] ?? ""),
                            (v) => updateRowLocal(row._id, col.col_key, v),
                            () => handleBlurSaved(row._id, col.col_key, String(row[col.col_key] ?? ""))
                          )}
                          {TOOLTIP_COLUMNS.has(col.col_key) && (
                            <InfoPopover column={col.col_key} value={String(row[col.col_key] ?? "")} />
                          )}
                        </div>
                      </td>
                    ))}
                    <td className="row-actions">
                      {savedFlash === row._id && <span className="saved-flash">저장됨</span>}
                      <button className="danger-btn" onClick={() => handleDelete(row._id)}>
                        삭제
                      </button>
                    </td>
                  </tr>
                );
              })}

              {drafts.map((draft) => (
                <tr key={draft.clientId} className="row-draft">
                  {columns.map((col) => (
                    <td key={col.col_key}>
                      {renderInput(col, draft.data[col.col_key], (v) =>
                        updateDraft(draft.clientId, col.col_key, v)
                      )}
                    </td>
                  ))}
                  <td className="row-actions">
                    <button onClick={() => saveDraft(draft.clientId)}>저장</button>
                    <button className="ghost-btn" onClick={() => discardDraft(draft.clientId)}>
                      취소
                    </button>
                  </td>
                </tr>
              ))}

              {rows.length === 0 && drafts.length === 0 && (
                <tr>
                  <td colSpan={columns.length + 1} className="empty-row">
                    데이터가 없습니다. "행 추가"로 새 행을 만들어보세요.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
