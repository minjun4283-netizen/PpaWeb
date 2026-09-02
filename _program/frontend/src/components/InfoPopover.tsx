import { useRef, useState } from "react";
import { createPortal } from "react-dom";
import { api, type TooltipField } from "../api";

// Touch-friendly replacement for the original VBA UserForm hover tooltip:
// tap the info button to fetch and show related-record fields, tap again to close.
// Rendered through a portal (fixed position, viewport coordinates) so it can't
// be clipped by the data grid's own scroll container.
export function InfoPopover({ column, value }: { column: string; value: string }) {
  const [open, setOpen] = useState(false);
  const [fields, setFields] = useState<TooltipField[] | null>(null);
  const [loading, setLoading] = useState(false);
  const [coords, setCoords] = useState({ top: 0, left: 0 });
  const btnRef = useRef<HTMLButtonElement>(null);

  async function toggle() {
    if (open) {
      setOpen(false);
      return;
    }
    const rect = btnRef.current?.getBoundingClientRect();
    if (rect) setCoords({ top: rect.bottom + 6, left: rect.left });

    setOpen(true);
    if (!value.trim()) {
      setFields([]);
      return;
    }
    setLoading(true);
    try {
      const res = await api.tooltip(column, value.trim());
      setFields(res.fields);
    } finally {
      setLoading(false);
    }
  }

  return (
    <span className="info-popover-wrap">
      <button
        type="button"
        ref={btnRef}
        className="info-btn"
        onClick={toggle}
        aria-label={`${column} 관련정보 보기`}
      >
        ⓘ
      </button>
      {open &&
        createPortal(
          <>
            <div className="info-popover-backdrop" onClick={() => setOpen(false)} />
            <div className="info-popover" style={{ top: coords.top, left: coords.left }}>
              {loading && <div className="info-loading">불러오는 중...</div>}
              {!loading && fields && fields.length === 0 && (
                <div className="info-empty">값을 입력하세요.</div>
              )}
              {!loading &&
                fields &&
                fields.map((f) => (
                  <div key={f.label} className="info-row">
                    <span className="info-label">{f.label}</span>
                    <span className="info-value">{f.value}</span>
                  </div>
                ))}
            </div>
          </>,
          document.body
        )}
    </span>
  );
}
