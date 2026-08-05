import { useState } from "react";
import { useValidation } from "../ValidationContext";
import { api } from "../api";
import { useTables } from "../TablesContext";

function defaultSince(): string {
  const d = new Date(Date.now() - 24 * 60 * 60 * 1000);
  d.setSeconds(0, 0);
  return d.toISOString().slice(0, 16);
}

export function ValidationPage() {
  const { report, running, run } = useValidation();
  const { tables } = useTables();
  const [since, setSince] = useState(defaultSince());

  function labelOf(tableKey: string): string {
    return tables.find((t) => t.table_key === tableKey)?.label ?? tableKey;
  }

  function downloadUrl(type: "added" | "modified"): string {
    const iso = since ? new Date(since).toISOString() : "1970-01-01T00:00:00.000Z";
    return api.exportUrl(type, iso);
  }

  return (
    <div className="validation-page">
      <div className="table-page-header">
        <h2>검증 리포트</h2>
        <button onClick={() => run()} disabled={running}>
          {running ? "검증 중..." : "검증 실행"}
        </button>
      </div>

      {!report && <p>아직 실행되지 않았습니다. "검증 실행" 버튼을 눌러주세요.</p>}

      {report && (
        <>
          <div className="summary-cards">
            <div className="summary-card">
              <div className="summary-title">총 오류건수</div>
              <div className="summary-value">{report.totalErrors}</div>
            </div>
            <div className="summary-card">
              <div className="summary-title">실행시각</div>
              <div className="summary-value small">{new Date(report.runAt).toLocaleString()}</div>
            </div>
          </div>

          <div className="summary-grid">
            <div>
              <h3>오류항목별 건수</h3>
              <table className="mini-table">
                <thead>
                  <tr>
                    <th>오류항목</th>
                    <th>건수</th>
                  </tr>
                </thead>
                <tbody>
                  {report.summaryByErrorItem.length === 0 && (
                    <tr>
                      <td colSpan={2}>오류 없음</td>
                    </tr>
                  )}
                  {report.summaryByErrorItem.map((s) => (
                    <tr key={s.key}>
                      <td>{s.key}</td>
                      <td>{s.count}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div>
              <h3>표별 건수</h3>
              <table className="mini-table">
                <thead>
                  <tr>
                    <th>표</th>
                    <th>건수</th>
                  </tr>
                </thead>
                <tbody>
                  {report.summaryByTable.length === 0 && (
                    <tr>
                      <td colSpan={2}>오류 없음</td>
                    </tr>
                  )}
                  {report.summaryByTable.map((s) => (
                    <tr key={s.key}>
                      <td>{labelOf(s.key)}</td>
                      <td>{s.count}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          <h3>상세 오류 목록</h3>
          <div className="grid-scroll">
            <table className="data-grid">
              <thead>
                <tr>
                  <th>원본</th>
                  <th>행ID</th>
                  <th>ID값</th>
                  <th>오류항목</th>
                  <th>결과</th>
                </tr>
              </thead>
              <tbody>
                {report.errors.length === 0 && (
                  <tr>
                    <td colSpan={5} className="empty-row">
                      오류 없음
                    </td>
                  </tr>
                )}
                {report.errors.map((e, i) => (
                  <tr key={i}>
                    <td>{labelOf(e.tableKey)}</td>
                    <td>{e.rowId}</td>
                    <td>{e.pkValue}</td>
                    <td>{e.errorItem}</td>
                    <td className="cell-error">{e.result}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}

      <div className="export-section">
        <h3>추가/수정 파일 내보내기</h3>
        <p className="hint">기준 시각 이후 변경된 행을 표별 시트로 묶은 xlsx 파일을 내려받습니다.</p>
        <label className="inline-label">
          기준 시각
          <input type="datetime-local" value={since} onChange={(e) => setSince(e.target.value)} />
        </label>
        <div className="export-buttons">
          <a className="button-link" href={downloadUrl("added")}>
            추가분 다운로드
          </a>
          <a className="button-link" href={downloadUrl("modified")}>
            수정분 다운로드
          </a>
        </div>
      </div>
    </div>
  );
}
