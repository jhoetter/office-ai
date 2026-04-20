"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { useFocusTrap } from "@officeai/ui";
import { XlsxAgent, cellKey, colToLetter, formatA1, type Cell, type CellValue } from "@officeai/xlsx";

/**
 * "Edit Data" modal for embedded spreadsheets / charts in Word and
 * PowerPoint. Hosts a transient {@link XlsxAgent} loaded from the
 * embedded `.xlsx` bytes, lets the user edit cells through a
 * deliberately minimal grid, and emits the new bytes (and a plain
 * 2D mirror for chart preview regeneration) back to the host editor
 * on save.
 *
 * The host editor is responsible for dispatching the right bus
 * command with these bytes:
 *
 *   - OLE spreadsheet edits → `*:update-spreadsheet`
 *   - Chart edits → `*:update-spreadsheet` (refresh embedded book)
 *     plus `*:set-chart-data` (refresh the chart XML's typed
 *     categories / series so PowerPoint re-renders the visual)
 *
 * The grid intentionally stays minimal: it is a Word/PowerPoint
 * inline edit affordance, not a full xlsx editor. Power users still
 * have the option to "Open in spreadsheet editor" (a future button)
 * that round-trips the bytes through the standalone XlsxEditor.
 */

interface Props {
  readonly open: boolean;
  /** Initial workbook bytes to seed the transient agent. */
  readonly bytes: Uint8Array | null;
  /** Title shown in the dialog header (e.g. `"Edit chart data"`). */
  readonly title?: string;
  readonly onCancel: () => void;
  readonly onSave: (result: {
    readonly bytes: Uint8Array;
    readonly grid: ReadonlyArray<ReadonlyArray<string | number | null>>;
  }) => void;
}

export function EmbeddedXlsxModal(props: Props): React.ReactElement | null {
  const { open, bytes, title = "Edit data", onCancel, onSave } = props;
  const [agent, setAgent] = useState<XlsxAgent | null>(null);
  const [tick, setTick] = useState(0);
  const [activeSheet, setActiveSheet] = useState<string>("");
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const panelRef = useRef<HTMLDivElement | null>(null);

  useFocusTrap(panelRef, { enabled: open, onEscape: onCancel });

  useEffect(() => {
    if (!open || !bytes) {
      setAgent(null);
      setActiveSheet("");
      setError(null);
      setSaving(false);
      return;
    }
    let cancelled = false;
    setLoading(true);
    setError(null);
    void (async () => {
      try {
        const a = await XlsxAgent.fromBuffer(bytes);
        if (cancelled) return;
        setAgent(a);
        const sheets = a.listSheets().filter((s) => s.kind === "worksheet");
        if (sheets.length > 0) setActiveSheet(sheets[0]!.name);
      } catch (err) {
        if (!cancelled) setError(err instanceof Error ? err.message : String(err));
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [open, bytes]);

  const sheets = useMemo(() => {
    if (!agent) return [];
    return agent.listSheets().filter((s) => s.kind === "worksheet");
  }, [agent, tick]);

  const sheet = useMemo(() => {
    if (!agent || !activeSheet) return null;
    return agent.getSnapshot().root.sheets.find((s) => s.name === activeSheet) ?? null;
  }, [agent, activeSheet, tick]);

  const dims = useMemo(() => computeDims(sheet?.cells), [sheet]);

  const setCell = async (row: number, col: number, raw: string) => {
    if (!agent || !activeSheet) return;
    try {
      const value = parseCellInput(raw);
      await agent.applyCommand({
        type: "xlsx:set-cell-value",
        payload: { sheet: activeSheet, ref: formatA1({ row, col }), value },
        source: "human",
      });
      setTick((t) => t + 1);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    }
  };

  const submit = async () => {
    if (!agent) return;
    setSaving(true);
    setError(null);
    try {
      const ab = await agent.exportFile();
      const out = new Uint8Array(ab);
      const grid = projectGrid(sheet?.cells, dims.rows, dims.cols);
      onSave({ bytes: out, grid });
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setSaving(false);
    }
  };

  if (!open) return null;

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label={title}
      data-testid="embedded-xlsx-modal"
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(0,0,0,0.45)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        zIndex: 1100,
      }}
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onCancel();
      }}
    >
      <div
        ref={panelRef}
        tabIndex={-1}
        style={{
          width: "min(900px, 92vw)",
          maxHeight: "90vh",
          background: "var(--background)",
          color: "var(--foreground)",
          border: "1px solid var(--border)",
          borderRadius: 8,
          boxShadow: "0 16px 48px rgba(0,0,0,0.22)",
          display: "flex",
          flexDirection: "column",
          outline: "none",
        }}
      >
        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            padding: "14px 18px",
            borderBottom: "1px solid var(--border)",
          }}
        >
          <div style={{ fontSize: 14, fontWeight: 600 }}>{title}</div>
          {sheets.length > 1 && (
            <select
              value={activeSheet}
              onChange={(e) => setActiveSheet(e.target.value)}
              style={{
                padding: "4px 6px",
                border: "1px solid var(--border)",
                borderRadius: 4,
                fontSize: 12,
                background: "var(--background)",
                color: "var(--foreground)",
              }}
            >
              {sheets.map((s) => (
                <option key={s.name} value={s.name}>
                  {s.name}
                </option>
              ))}
            </select>
          )}
        </div>

        <div
          style={{
            flex: 1,
            minHeight: 200,
            overflow: "auto",
            padding: 12,
          }}
        >
          {loading && <div style={{ fontSize: 12 }}>Loading workbook…</div>}
          {!loading && !sheet && !error && <div style={{ fontSize: 12 }}>No worksheet to edit.</div>}
          {sheet && <SimpleGrid sheet={sheet} rows={dims.rows} cols={dims.cols} onCommit={setCell} />}
        </div>

        {error && (
          <div
            style={{
              padding: "0 18px 8px",
              fontSize: 12,
              color: "var(--destructive, #b91c1c)",
            }}
          >
            {error}
          </div>
        )}

        <div
          style={{
            display: "flex",
            justifyContent: "flex-end",
            gap: 8,
            padding: "12px 18px",
            borderTop: "1px solid var(--border)",
          }}
        >
          <button
            type="button"
            onClick={onCancel}
            style={{
              padding: "6px 12px",
              border: "1px solid var(--border)",
              borderRadius: 4,
              background: "var(--background)",
              color: "var(--foreground)",
              fontSize: 12,
              cursor: "pointer",
            }}
          >
            Cancel
          </button>
          <button
            type="button"
            data-testid="embedded-xlsx-modal-save"
            onClick={() => void submit()}
            disabled={!agent || saving}
            style={{
              padding: "6px 12px",
              border: "1.5px solid var(--ai-violet)",
              borderRadius: 4,
              background: agent && !saving ? "var(--ai-violet)" : "var(--muted)",
              color: agent && !saving ? "white" : "var(--muted-foreground)",
              fontSize: 12,
              cursor: agent && !saving ? "pointer" : "not-allowed",
            }}
          >
            {saving ? "Saving…" : "Save"}
          </button>
        </div>
      </div>
    </div>
  );
}

interface SimpleGridProps {
  readonly sheet: { readonly cells: ReadonlyMap<string, Cell> };
  readonly rows: number;
  readonly cols: number;
  readonly onCommit: (row: number, col: number, raw: string) => void;
}

/**
 * Bare-minimum table editor: header row of column letters + 1-based
 * row numbers + per-cell `<input>`. Sufficient for the embed-edit
 * workflow where users tweak a handful of values; not meant to
 * replicate the full XlsxEditor's keyboard model. Headroom rows /
 * columns (`+2 rows, +1 col`) are added so the user can extend the
 * range without an explicit "add row" UI — typing into a virgin
 * cell just commits the value through the bus and the next render
 * widens the grid further.
 */
function SimpleGrid(props: SimpleGridProps): React.ReactElement {
  const { sheet, rows, cols, onCommit } = props;
  const renderRows = Math.max(rows + 2, 6);
  const renderCols = Math.max(cols + 1, 4);

  return (
    <table
      style={{
        borderCollapse: "collapse",
        fontSize: 12,
        fontFamily: "ui-monospace, SFMono-Regular, Menlo, monospace",
      }}
    >
      <thead>
        <tr>
          <th style={cellHeaderStyle} />
          {Array.from({ length: renderCols }, (_, c) => (
            <th key={c} style={cellHeaderStyle}>
              {colToLetter(c)}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {Array.from({ length: renderRows }, (_, r) => (
          <tr key={r}>
            <th style={cellHeaderStyle}>{r + 1}</th>
            {Array.from({ length: renderCols }, (_, c) => {
              const cell = sheet.cells.get(cellKey(r, c));
              return (
                <td key={c} style={cellTdStyle}>
                  <CellInput initial={cellDisplay(cell)} onCommit={(raw) => onCommit(r, c, raw)} />
                </td>
              );
            })}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

interface CellInputProps {
  readonly initial: string;
  readonly onCommit: (raw: string) => void;
}

/**
 * Local-mirror input that only emits a commit on blur or Enter so
 * the bus dispatch / re-render churn stays bounded — we don't want
 * to fire a `set-cell-value` per keystroke.
 */
function CellInput(props: CellInputProps): React.ReactElement {
  const [value, setValue] = useState(props.initial);
  useEffect(() => {
    setValue(props.initial);
  }, [props.initial]);
  return (
    <input
      type="text"
      value={value}
      onChange={(e) => setValue(e.target.value)}
      onBlur={() => {
        if (value !== props.initial) props.onCommit(value);
      }}
      onKeyDown={(e) => {
        if (e.key === "Enter") {
          e.preventDefault();
          (e.currentTarget as HTMLInputElement).blur();
        } else if (e.key === "Escape") {
          e.preventDefault();
          setValue(props.initial);
          (e.currentTarget as HTMLInputElement).blur();
        }
      }}
      style={{
        width: 80,
        padding: "2px 4px",
        border: "none",
        background: "transparent",
        color: "var(--foreground)",
        fontFamily: "inherit",
        fontSize: "inherit",
        outline: "none",
      }}
    />
  );
}

const cellHeaderStyle: React.CSSProperties = {
  border: "1px solid var(--border)",
  background: "var(--muted, #f4f4f5)",
  color: "var(--muted-foreground)",
  padding: "2px 6px",
  minWidth: 32,
  textAlign: "center",
  fontWeight: 500,
};

const cellTdStyle: React.CSSProperties = {
  border: "1px solid var(--border)",
  padding: 0,
  background: "var(--background)",
};

function computeDims(cells: ReadonlyMap<string, Cell> | null | undefined): {
  readonly rows: number;
  readonly cols: number;
} {
  if (!cells || cells.size === 0) return { rows: 0, cols: 0 };
  let maxRow = 0;
  let maxCol = 0;
  for (const key of cells.keys()) {
    const idx = key.indexOf(":");
    if (idx === -1) continue;
    const r = Number(key.slice(0, idx));
    const c = Number(key.slice(idx + 1));
    if (r > maxRow) maxRow = r;
    if (c > maxCol) maxCol = c;
  }
  return { rows: maxRow + 1, cols: maxCol + 1 };
}

function projectGrid(
  cells: ReadonlyMap<string, Cell> | null | undefined,
  rows: number,
  cols: number
): ReadonlyArray<ReadonlyArray<string | number | null>> {
  const out: Array<Array<string | number | null>> = [];
  for (let r = 0; r < rows; r++) {
    const row: Array<string | number | null> = [];
    for (let c = 0; c < cols; c++) {
      const cell = cells?.get(cellKey(r, c));
      row.push(cellToScalar(cell?.value));
    }
    out.push(row);
  }
  return out;
}

function cellToScalar(value: CellValue | undefined): string | number | null {
  if (value === undefined || value === null) return null;
  if (typeof value === "number") return value;
  if (typeof value === "string") return value;
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return value.code;
}

function cellDisplay(cell: Cell | undefined): string {
  if (!cell) return "";
  const v = cell.value;
  if (v === null || v === undefined) return "";
  if (typeof v === "number") return String(v);
  if (typeof v === "string") return v;
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  return v.code;
}

/**
 * Promote a raw input string into a typed `CellValue`. Mirrors
 * `xlsx:set-cell-value` semantics on a deliberately small subset:
 * empty → null, parsable number → number, `TRUE`/`FALSE` → boolean,
 * everything else → string. Formula authoring (`=SUM(...)`) is
 * intentionally not exposed here — the modal is for value tweaks.
 */
function parseCellInput(raw: string): CellValue {
  const trimmed = raw.trim();
  if (trimmed === "") return null;
  if (/^-?\d+(\.\d+)?$/.test(trimmed)) return Number(trimmed);
  if (trimmed.toUpperCase() === "TRUE") return true;
  if (trimmed.toUpperCase() === "FALSE") return false;
  return raw;
}
