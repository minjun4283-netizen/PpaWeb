import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { useParams } from "react-router-dom";
import { api, type ColumnDef, type RowRecord, type TableDef } from "../api";
import { useTables } from "../TablesContext";
import { useValidation } from "../ValidationContext";
import { InfoPopover } from "../components/InfoPopover";

// The original macro only wired up the hyperlink tooltip on the T_수급매칭
// sheet (see Build전기사용지Tooltip_Expanded / Build구매계약Tooltip_Expanded).
// Gating on table + column (not column name alone) matters because these
// same column keys are also the PK of T_전기사용지 / T_구매계약 themselves —
// without the table check the info button would wrongly show up next to a
// table's own primary key.
const TOOLTIP_TABLE = "T_수급매칭";
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

// Original VBA stored one flag column per check (PK공란, 발전소ID 참조, ...) so
// a glance at the sheet showed exactly which field was wrong. We don't persist
// flag columns, so instead map each reported error item back to the specific
// column(s) it concerns, to highlight just that cell rather than the whole row.
function errorItemToColumns(tableDef: TableDef, errorItem: string): string[] {
  if (errorItem === "PK 공란" || errorItem === "PK 중복") return [tableDef.pk_column];
  if (errorItem === "조합중복") return tableDef.uniqueGroups.flat();
  const match = errorItem.match(/^(.*) (공란|참조)$/);
  return match ? [match[1]] : [];
}

type SortDir = "asc" | "desc" | null;

export function TablePage() {
  const { tableKey = "" } = useParams();
  const { tables } = useTables();
  const { report } = useValidation();
  const tableDef = tables.find((t) => t.table_key === tableKey);

  const [rows, setRows] = useState<RowRecord[]>([]);
  const [drafts, setDrafts] = useState<Draft[]>([]);
  const [loading, setLoading] = useState(true);
  const [savedFlash, setSavedFlash] = useState<number | null>(null);
  const [search, setSearch] = useState("");
  const [sortCol, setSortCol] = useState<string | null>(null);
  const [sortDir, setSortDir] = useState<SortDir>(null);
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
    setSearch("");
    setSortCol(null);
    setSortDir(null);
  }, [load]);

  const errorsByRow = useMemo(() => {
    const map = new Map<number, { items: string[]; columns: Set<string> }>();
    if (report && tableDef) {
      for (const err of report.errors) {
        if (err.tableKey !== tableKey) continue;
        if (!map.has(err.rowId)) map.set(err.rowId, { items: [], columns: new Set() });
        const entry = map.get(err.rowId)!;
        entry.items.push(err.errorItem);
        for (const col of errorItemToColumns(tableDef, err.errorItem)) entry.columns.add(col);
      }
    }
    return map;
  }, [report, tableKey, tableDef]);

  const displayRows = useMemo(() => {
    let result = rows;
    const q = search.trim().toLowerCase();
    if (q) {
      result = result.filter((row) =>
        Object.entries(row).some(
          ([key, value]) => !key.startsWith("_") && String(value ?? "").toLowerCase().includes(q)
        )
      );
    }
    if (sortCol && sortDir) {
      result = [...result].sort((a, b) => {
        const av = a[sortCol];
        const bv = b[sortCol];
        const an = Number(av);
        const bn = Number(bv);
        let cmp: number;
        if (av != null && bv != null && av !== "" && bv !== "" && !Number.isNaN(an) && !Number.isNaN(bn)) {
          cmp = an - bn;
        } else {
          cmp = String(av ?? "").localeCompare(String(bv ?? ""), "ko");
        }
        return sortDir === "asc" ? cmp : -cmp;
      });
    }
    return result;
  }, [rows, search, sortCol, sortDir]);

  if (!tableDef) return <p>불러오는 중...</p>;
  const columns = tableDef.columns;

  function toggleSort(colKey: string) {
    if (sortCol !== colKey) {
      setSortCol(colKey);
      setSortDir("asc");
    } else if (sortDir === "asc") {
      setSortDir("desc");
    } else {
      setSortCol(null);
      setSortDir(null);
    }
  }

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
    onBlur?: () => void,
    hasError?: boolean
  ) {
    const inputType = col.type === "number" ? "number" : col.type === "date" ? "date" : "text";
    return (
      <input
        className={`cell-input${hasError ? " cell-input-error" : ""}`}
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
        <div className="table-page-actions">
          <input
            className="search-input"
            type="search"
            placeholder="검색..."
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
          <button onClick={addDraft}>+ 행 추가</button>
        </div>
      </div>

      {loading ? (
        <p>불러오는 중...</p>
      ) : (
        <div className="grid-scroll">
          <table className="data-grid">
            <thead>
              <tr>
                {columns.map((c) => (
                  <th key={c.col_key} className="sortable-th" onClick={() => toggleSort(c.col_key)}>
                    {c.label}
                    {c.col_key === tableDef.pk_column && <span className="pk-badge">PK</span>}
                    {c.is_fk === 1 && <span className="fk-badge">FK→{c.ref_table}</span>}
                    {sortCol === c.col_key && (
                      <span className="sort-indicator">{sortDir === "asc" ? " ▲" : " ▼"}</span>
                    )}
                  </th>
                ))}
                <th>작업</th>
              </tr>
            </thead>
            <tbody>
              {displayRows.map((row) => {
                const errInfo = errorsByRow.get(row._id);
                return (
                  <tr
                    key={row._id}
                    className={errInfo ? "row-error" : ""}
                    title={errInfo?.items.join(", ")}
                  >
                    {columns.map((col) => (
                      <td key={col.col_key}>
                        <div className="cell-with-info">
                          {renderInput(
                            col,
                            String(row[col.col_key] ?? ""),
                            (v) => updateRowLocal(row._id, col.col_key, v),
                            () => handleBlurSaved(row._id, col.col_key, String(row[col.col_key] ?? "")),
                            errInfo?.columns.has(col.col_key)
                          )}
                          {tableKey === TOOLTIP_TABLE && TOOLTIP_COLUMNS.has(col.col_key) && (
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

              {displayRows.length === 0 && drafts.length === 0 && (
                <tr>
                  <td colSpan={columns.length + 1} className="empty-row">
                    {rows.length === 0
                      ? '데이터가 없습니다. "행 추가"로 새 행을 만들어보세요.'
                      : "검색 결과가 없습니다."}
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
