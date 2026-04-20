"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { useFocusTrap } from "@officeai/ui";
import {
  XlsxAgent,
  colToLetter,
  parseCellKey,
  type XlsxClipboardSnapshot,
} from "@officeai/xlsx";
import type { XlsxEmbedMode } from "./xlsxEmbedShared";

/**
 * Pick an xlsx file → sheet → range, choose a paste mode, and emit a
 * fully-resolved {@link XlsxClipboardSnapshot} ready for the embed
 * dispatcher (`applyXlsxEmbed`). Used by the DOCX/PPTX editors'
 * "Insert from xlsx" actions on the action catalogue.
 *
 * Intentionally minimal:
 *   - File chooser → loads into a transient `XlsxAgent`.
 *   - Sheet dropdown defaults to the first worksheet.
 *   - Range defaults to the populated bounding box of the selected
 *     sheet (`A1:` + last column letter + last row number) so the
 *     common "import the whole sheet" workflow is one click.
 *   - Three mode buttons mirror the paste-mode taxonomy.
 *
 * No live preview today — that lands in a follow-up once the editor
 * canvas can host a thumbnail. The point of this dialog is parity
 * with the action catalogue, not a full picker UX.
 */

export interface XlsxRangePickerResult {
  readonly snapshot: XlsxClipboardSnapshot;
  readonly mode: XlsxEmbedMode;
}

interface Props {
  readonly open: boolean;
  /** Pre-selected mode; the dialog still lets the user change it. */
  readonly defaultMode?: XlsxEmbedMode;
  readonly onCancel: () => void;
  readonly onSubmit: (result: XlsxRangePickerResult) => void;
}

interface Loaded {
  readonly fileName: string;
  readonly agent: XlsxAgent;
  readonly sheets: ReadonlyArray<{ name: string; index: number }>;
}

const MODE_OPTIONS: ReadonlyArray<{ readonly id: XlsxEmbedMode; readonly label: string; readonly hint: string }> = [
  { id: "materialized", label: "Table", hint: "Native editable cells" },
  { id: "live", label: "Live spreadsheet", hint: "Double-click → Excel" },
  { id: "chart", label: "Chart", hint: "First row = series, first col = categories" },
];

export function XlsxRangePickerDialog(props: Props): React.ReactElement | null {
  const { open, defaultMode = "materialized", onCancel, onSubmit } = props;
  const [loaded, setLoaded] = useState<Loaded | null>(null);
  const [sheetName, setSheetName] = useState<string>("");
  const [range, setRange] = useState<string>("");
  const [mode, setMode] = useState<XlsxEmbedMode>(defaultMode);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState<boolean>(false);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const panelRef = useRef<HTMLDivElement | null>(null);

  useFocusTrap(panelRef, { enabled: open, onEscape: onCancel });

  useEffect(() => {
    if (!open) {
      setLoaded(null);
      setSheetName("");
      setRange("");
      setMode(defaultMode);
      setError(null);
      setBusy(false);
    }
  }, [open, defaultMode]);

  const usedRange = useMemo(() => {
    if (!loaded) return "";
    const sheet = loaded.agent
      .getSnapshot()
      .root.sheets.find((s) => s.name === sheetName);
    if (!sheet) return "A1:A1";
    return computeUsedRange(sheet.cells);
  }, [loaded, sheetName]);

  useEffect(() => {
    if (usedRange) setRange(usedRange);
  }, [usedRange]);

  const onPickFile = async (file: File) => {
    setError(null);
    setBusy(true);
    try {
      const buf = new Uint8Array(await file.arrayBuffer());
      const agent = await XlsxAgent.fromBuffer(buf);
      const sheets = agent.listSheets().filter((s) => s.kind === "worksheet");
      if (sheets.length === 0) {
        throw new Error("Workbook has no worksheets.");
      }
      setLoaded({ fileName: file.name, agent, sheets });
      setSheetName(sheets[0]!.name);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  };

  const submit = () => {
    if (!loaded || !sheetName) return;
    setError(null);
    try {
      const snapshot = loaded.agent.getClipboardSnapshot({
        sheet: sheetName,
        range: range.trim() || "A1:A1",
      });
      if (snapshot.width <= 0 || snapshot.height <= 0) {
        throw new Error("Selected range is empty.");
      }
      if (mode === "chart" && (snapshot.width < 2 || snapshot.height < 2)) {
        throw new Error("Chart needs at least 2 columns × 2 rows.");
      }
      onSubmit({ snapshot, mode });
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    }
  };

  if (!open) return null;

  const canSubmit = !!loaded && !!sheetName && !!range.trim() && !busy;

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label="Insert from xlsx"
      data-testid="xlsx-range-picker-dialog"
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(0,0,0,0.35)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        zIndex: 1000,
      }}
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onCancel();
      }}
    >
      <div
        ref={panelRef}
        tabIndex={-1}
        style={{
          width: 460,
          maxHeight: "90vh",
          overflowY: "auto",
          background: "var(--background)",
          color: "var(--foreground)",
          border: "1px solid var(--border)",
          borderRadius: 8,
          boxShadow: "0 12px 32px rgba(0,0,0,0.18)",
          padding: 20,
          display: "flex",
          flexDirection: "column",
          gap: 14,
          outline: "none",
        }}
      >
        <div style={{ fontSize: 14, fontWeight: 600 }}>Insert from xlsx</div>

        {!loaded && (
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              disabled={busy}
              style={{
                padding: "10px 16px",
                border: "1px dashed var(--border)",
                borderRadius: 6,
                background: "var(--background)",
                color: "var(--foreground)",
                fontSize: 13,
                cursor: busy ? "wait" : "pointer",
              }}
              data-testid="xlsx-picker-choose-file"
            >
              {busy ? "Loading…" : "Choose .xlsx file"}
            </button>
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              style={{ display: "none" }}
              onChange={(e) => {
                const file = e.target.files?.[0];
                e.target.value = "";
                if (file) void onPickFile(file);
              }}
            />
          </div>
        )}

        {loaded && (
          <>
            <div style={{ fontSize: 12, color: "var(--muted-foreground)" }}>
              {loaded.fileName}
            </div>

            <label style={{ display: "flex", flexDirection: "column", gap: 4 }}>
              <span style={{ fontSize: 12, fontWeight: 500 }}>Sheet</span>
              <select
                value={sheetName}
                onChange={(e) => setSheetName(e.target.value)}
                style={{
                  padding: "6px 8px",
                  border: "1px solid var(--border)",
                  borderRadius: 4,
                  fontSize: 12,
                  background: "var(--background)",
                  color: "var(--foreground)",
                }}
                data-testid="xlsx-picker-sheet"
              >
                {loaded.sheets.map((s) => (
                  <option key={s.name} value={s.name}>
                    {s.name}
                  </option>
                ))}
              </select>
            </label>

            <label style={{ display: "flex", flexDirection: "column", gap: 4 }}>
              <span style={{ fontSize: 12, fontWeight: 500 }}>Range</span>
              <input
                type="text"
                value={range}
                onChange={(e) => setRange(e.target.value)}
                placeholder="A1:E10"
                style={{
                  padding: "6px 8px",
                  border: "1px solid var(--border)",
                  borderRadius: 4,
                  fontSize: 12,
                  fontFamily: "monospace",
                  background: "var(--background)",
                  color: "var(--foreground)",
                }}
                data-testid="xlsx-picker-range"
              />
            </label>

            <div>
              <div style={{ fontSize: 12, fontWeight: 500, marginBottom: 6 }}>
                Insert as
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 6 }}>
                {MODE_OPTIONS.map((opt) => {
                  const active = mode === opt.id;
                  return (
                    <button
                      key={opt.id}
                      type="button"
                      data-testid={`xlsx-picker-mode-${opt.id}`}
                      aria-pressed={active}
                      onClick={() => setMode(opt.id)}
                      style={{
                        padding: "8px 6px",
                        border: active ? "1.5px solid var(--ai-violet)" : "1px solid var(--border)",
                        borderRadius: 6,
                        background: active ? "var(--ai-violet-light)" : "var(--background)",
                        color: "var(--foreground)",
                        fontSize: 12,
                        cursor: "pointer",
                        display: "flex",
                        flexDirection: "column",
                        gap: 2,
                        textAlign: "center",
                      }}
                    >
                      <span style={{ fontWeight: 500 }}>{opt.label}</span>
                      <span style={{ fontSize: 10, color: "var(--muted-foreground)" }}>
                        {opt.hint}
                      </span>
                    </button>
                  );
                })}
              </div>
            </div>
          </>
        )}

        {error && (
          <div style={{ fontSize: 12, color: "var(--destructive, #b91c1c)" }}>{error}</div>
        )}

        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 4 }}>
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
            onClick={submit}
            disabled={!canSubmit}
            data-testid="xlsx-picker-insert"
            style={{
              padding: "6px 12px",
              border: "1.5px solid var(--ai-violet)",
              borderRadius: 4,
              background: canSubmit ? "var(--ai-violet)" : "var(--muted)",
              color: canSubmit ? "white" : "var(--muted-foreground)",
              fontSize: 12,
              cursor: canSubmit ? "pointer" : "not-allowed",
            }}
          >
            Insert
          </button>
        </div>
      </div>
    </div>
  );
}

/**
 * Compute the populated bounding box of a sheet's typed `cells` map
 * and emit it as an A1 range string. Empty sheets fall back to
 * `"A1:A1"` so the dialog doesn't ship an invalid-looking default.
 */
function computeUsedRange(cells: ReadonlyMap<string, unknown>): string {
  if (cells.size === 0) return "A1:A1";
  let maxRow = 0;
  let maxCol = 0;
  for (const key of cells.keys()) {
    const { row, col } = parseCellKey(key);
    if (row > maxRow) maxRow = row;
    if (col > maxCol) maxCol = col;
  }
  return `A1:${colToLetter(maxCol)}${maxRow + 1}`;
}
