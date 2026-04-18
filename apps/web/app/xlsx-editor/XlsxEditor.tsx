"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { DragEvent, ReactNode } from "react";
import { FolderOpen, Loader2, Sparkles, Download } from "lucide-react";
import { Button, cn } from "@officeai/ui";
import {
  XlsxAgent,
  cellKey,
  colToLetter,
  formatA1,
  type CellValue,
  type Sheet,
  type XlsxSnapshot,
} from "@officeai/xlsx";
import { buildSampleXlsx } from "@/lib/sample-xlsx";
import { Grid, type GridSelection } from "./Grid";

interface ToastMessage {
  id: number;
  kind: "info" | "warn" | "error";
  text: string;
}

const SAMPLE_NAME = "sample.xlsx";

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/**
 * Top-level XLSX editor surface for /xlsx-editor.
 *
 * Lifecycle (mirrors `DocxEditor`):
 *   1. On mount, build the synthetic xlsx via `buildSampleXlsx()` and
 *      load it into a fresh `XlsxAgent`.
 *   2. Subscribe to mutations to keep `revision`, `pendingCount`, and
 *      the visible cell snapshot in sync.
 *   3. Render header → formula bar → grid → sheet tabs → agent prompt.
 *
 * All cell mutations dispatch through `agent.applyCommand` so the
 * single command-bus invariant holds for the agent path too.
 */
export function XlsxEditor(): ReactNode {
  const agentRef = useRef<XlsxAgent | null>(null);
  const [agent, setAgent] = useState<XlsxAgent | null>(null);
  const [snapshot, setSnapshot] = useState<XlsxSnapshot | null>(null);
  const [activeSheetName, setActiveSheetName] = useState<string | null>(null);
  const [selection, setSelection] = useState<GridSelection | null>({ row: 0, col: 0 });
  const [pendingCount, setPendingCount] = useState(0);
  const [toasts, setToasts] = useState<ToastMessage[]>([]);
  const [agentBusy, setAgentBusy] = useState(false);
  const [agentPrompt, setAgentPrompt] = useState("");
  const [formulaDraft, setFormulaDraft] = useState("");
  const [formulaFocused, setFormulaFocused] = useState(false);
  const formulaInputRef = useRef<HTMLInputElement | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [filename, setFilename] = useState<string>(SAMPLE_NAME);
  const [dragOver, setDragOver] = useState(false);

  const toastIdRef = useRef(0);
  const pushToast = useCallback((kind: ToastMessage["kind"], text: string) => {
    const id = ++toastIdRef.current;
    setToasts((prev) => [...prev, { id, kind, text }]);
    setTimeout(() => setToasts((prev) => prev.filter((t) => t.id !== id)), 4000);
  }, []);

  // Holds the unsubscribe handle for the active agent's `subscribe()`
  // callback so we can swap agents (Open file) without leaking listeners.
  const offRef = useRef<(() => void) | null>(null);

  const mountAgent = useCallback(
    (a: XlsxAgent, name: string) => {
      offRef.current?.();
      agentRef.current = a;
      setAgent(a);
      setFilename(name);
      const snap = a.getSnapshot();
      setSnapshot(snap);
      setActiveSheetName(snap.root.sheets[0]?.name ?? null);
      setSelection({ row: 0, col: 0 });
      setPendingCount(a.getPendingMutations().length);
      offRef.current = a.subscribe((s) => {
        setSnapshot(s);
        setPendingCount(a.getPendingMutations().length);
      });
    },
    [setAgent, setSnapshot, setActiveSheetName, setSelection, setPendingCount]
  );

  useEffect(() => {
    let cancelled = false;
    void (async () => {
      try {
        const buf = await buildSampleXlsx();
        if (cancelled) return;
        const a = await XlsxAgent.fromBuffer(buf);
        if (cancelled) return;
        mountAgent(a, SAMPLE_NAME);
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    })();
    return () => {
      cancelled = true;
      offRef.current?.();
      offRef.current = null;
      agentRef.current = null;
      setAgent(null);
    };
  }, [pushToast, mountAgent]);

  const openFile = useCallback(
    async (file: File) => {
      const lower = file.name.toLowerCase();
      if (!lower.endsWith(".xlsx")) {
        pushToast("error", `Unsupported file: ${file.name} (only .xlsx)`);
        return;
      }
      try {
        const buf = await file.arrayBuffer();
        const a = await XlsxAgent.fromBuffer(buf);
        mountAgent(a, file.name);
        pushToast("info", `Opened ${file.name}`);
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [mountAgent, pushToast]
  );

  const onPickFile = useCallback(() => {
    fileInputRef.current?.click();
  }, []);

  const onFileInputChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (file) void openFile(file);
      e.target.value = "";
    },
    [openFile]
  );

  const onDragOver = useCallback((e: DragEvent<HTMLDivElement>) => {
    if (Array.from(e.dataTransfer.items).some((it) => it.kind === "file")) {
      e.preventDefault();
      setDragOver(true);
    }
  }, []);
  const onDragLeave = useCallback((e: DragEvent<HTMLDivElement>) => {
    if (e.currentTarget === e.target) setDragOver(false);
  }, []);
  const onDrop = useCallback(
    (e: DragEvent<HTMLDivElement>) => {
      e.preventDefault();
      setDragOver(false);
      const file = e.dataTransfer.files?.[0];
      if (file) void openFile(file);
    },
    [openFile]
  );

  const activeSheet: Sheet | null = useMemo(() => {
    if (!snapshot || !activeSheetName) return null;
    return snapshot.root.sheets.find((s) => s.name === activeSheetName) ?? null;
  }, [snapshot, activeSheetName]);

  const selectedCell = useMemo(() => {
    if (!activeSheet || !selection) return null;
    return activeSheet.cells.get(cellKey(selection.row, selection.col)) ?? null;
  }, [activeSheet, selection]);

  const selectedRef = selection ? formatA1({ row: selection.row, col: selection.col }) : "";

  // Derived display for the formula bar when the user is NOT actively
  // editing it. While the input has focus we surface `formulaDraft`
  // (uncommitted user keystrokes) instead so the snapshot subscription
  // can't clobber typing.
  const derivedFormulaDisplay = (() => {
    if (!selectedCell) return "";
    if (selectedCell.formula) return `=${selectedCell.formula.text}`;
    return formatCellValue(selectedCell.value);
  })();
  const formulaValue = formulaFocused ? formulaDraft : derivedFormulaDisplay;

  const dispatchCellEdit = useCallback(
    async (sheetName: string, ref: string, raw: string) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        if (raw.startsWith("=")) {
          await a.applyCommand({
            type: "xlsx:set-cell-formula",
            payload: { sheet: sheetName, ref, formula: raw.slice(1) },
            source: "human",
          });
        } else {
          const value: CellValue = parseLiteral(raw);
          await a.applyCommand({
            type: "xlsx:set-cell-value",
            payload: { sheet: sheetName, ref, value },
            source: "human",
          });
        }
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const onFormulaSubmit = useCallback(() => {
    if (!activeSheet || !selection) return;
    const ref = formatA1({ row: selection.row, col: selection.col });
    void dispatchCellEdit(activeSheet.name, ref, formulaDraft);
    setFormulaFocused(false);
    setFormulaDraft("");
    formulaInputRef.current?.blur();
  }, [activeSheet, selection, formulaDraft, dispatchCellEdit]);

  const onCommitGridEdit = useCallback(
    (sel: GridSelection, value: string) => {
      if (!activeSheet) return;
      const ref = formatA1({ row: sel.row, col: sel.col });
      void dispatchCellEdit(activeSheet.name, ref, value);
    },
    [activeSheet, dispatchCellEdit]
  );

  const onSave = useCallback(async () => {
    const a = agentRef.current;
    if (!a) return;
    try {
      const buf = await a.exportFile();
      const blob = new Blob([buf], { type: XLSX_MIME });
      const url = URL.createObjectURL(blob);
      const a2 = document.createElement("a");
      a2.href = url;
      a2.download = filename;
      a2.click();
      URL.revokeObjectURL(url);
      pushToast("info", `Exported ${filename}`);
    } catch (err) {
      pushToast("error", err instanceof Error ? err.message : String(err));
    }
  }, [filename, pushToast]);

  const onAgentRun = useCallback(async () => {
    const a = agentRef.current;
    if (!a || !activeSheet) return;
    const prompt = agentPrompt.trim();
    if (!prompt) return;
    setAgentBusy(true);
    try {
      // Demo recipe — the parallel of DocxEditor's "[AI] " prefix.
      // The agent stages a `xlsx:set-cell-value` against a free cell
      // (col 3 / "D") in the active sheet's bounding row range. It is
      // marked `source: "agent"` so it lands in the pending queue for
      // human review, surfacing the badge in the header.
      const target = pickAgentTarget(activeSheet);
      await a.applyCommand({
        type: "xlsx:set-cell-value",
        payload: {
          sheet: activeSheet.name,
          ref: target,
          value: `[AI] ${prompt}`,
        },
        source: "agent",
        agentId: "demo-agent",
      });
      setAgentPrompt("");
      pushToast("info", `Agent staged a change at ${target}.`);
    } catch (err) {
      pushToast("error", err instanceof Error ? err.message : String(err));
    } finally {
      setAgentBusy(false);
    }
  }, [activeSheet, agentPrompt, pushToast]);

  const sheets = snapshot?.root.sheets ?? [];
  const revision = snapshot?.revision ?? 0;

  return (
    <div
      className={cn(
        "xlsx-editor relative flex h-full min-h-0 flex-col gap-3",
        dragOver && "ring-2 ring-[var(--ai-violet)] ring-offset-2"
      )}
      onDragOver={onDragOver}
      onDragLeave={onDragLeave}
      onDrop={onDrop}
    >
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx"
        data-testid="open-xlsx-input"
        className="sr-only"
        onChange={onFileInputChange}
      />
      <header className="flex flex-wrap items-center justify-between gap-2 rounded-md border border-divider bg-surface px-3 py-2">
        <div className="flex items-center gap-3">
          <span data-testid="filename" className="text-sm font-medium text-foreground">
            {filename}
          </span>
          <span
            data-testid="revision-badge"
            className="rounded-full border border-divider bg-background px-2 py-0.5 text-[10px] font-medium text-secondary"
          >
            rev {revision}
          </span>
          <span
            data-testid="pending-badge"
            className={cn(
              "rounded-full px-2 py-0.5 text-[10px] font-medium",
              pendingCount > 0
                ? "bg-[var(--ai-violet-light)] text-[var(--ai-violet)]"
                : "bg-background text-secondary"
            )}
          >
            {pendingCount} pending
          </span>
        </div>
        <div className="flex items-center gap-2">
          <Button
            size="sm"
            variant="ghost"
            onClick={onPickFile}
            data-testid="open-xlsx"
            title="Open .xlsx from disk"
          >
            <FolderOpen size={14} />
            Open
          </Button>
          <Button
            size="sm"
            variant="primary"
            onClick={() => void onSave()}
            disabled={!agent}
            data-testid="save-xlsx"
          >
            <Download size={14} />
            Save
          </Button>
        </div>
      </header>
      {dragOver ? (
        <div
          aria-hidden
          className="pointer-events-none fixed inset-0 z-40 flex items-center justify-center bg-[var(--ai-violet)]/10"
        >
          <div className="rounded-lg border border-[var(--ai-violet)] bg-surface px-4 py-2 text-sm text-foreground shadow-md">
            Drop a .xlsx file to open
          </div>
        </div>
      ) : null}

      <div className="formula-bar flex items-center gap-2 rounded-md border border-divider bg-surface px-2 py-1.5">
        <span
          data-testid="cell-ref"
          className="inline-flex h-7 min-w-[60px] items-center justify-center rounded border border-divider bg-background px-2 text-xs font-mono text-foreground"
        >
          {selectedRef || "—"}
        </span>
        <span className="text-secondary text-xs font-mono">fx</span>
        <input
          ref={formulaInputRef}
          data-testid="formula-input"
          aria-label="Formula bar"
          value={formulaValue}
          onChange={(e) => setFormulaDraft(e.target.value)}
          onFocus={() => {
            setFormulaDraft(derivedFormulaDisplay);
            setFormulaFocused(true);
          }}
          onBlur={() => {
            setFormulaFocused(false);
            setFormulaDraft("");
          }}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              onFormulaSubmit();
            } else if (e.key === "Escape") {
              e.preventDefault();
              setFormulaFocused(false);
              setFormulaDraft("");
              formulaInputRef.current?.blur();
            }
          }}
          placeholder={selection ? "Type a value or =formula" : "Select a cell to edit"}
          disabled={!selection || !agent}
          className="flex-1 bg-transparent px-1 py-1 text-xs text-foreground placeholder:text-tertiary focus:outline-none"
        />
      </div>

      <div className="relative flex-1 min-h-0">
        {activeSheet ? (
          <Grid
            sheet={activeSheet}
            selection={selection}
            onSelect={setSelection}
            onCommitEdit={onCommitGridEdit}
          />
        ) : (
          <div className="flex h-full items-center justify-center rounded-md border border-divider bg-background text-sm text-secondary">
            <Loader2 className="mr-2 animate-spin" size={14} />
            Loading workbook…
          </div>
        )}
      </div>

      <div
        data-testid="sheet-tabs"
        className="sheet-tabs flex items-center gap-1 overflow-x-auto rounded-md border border-divider bg-surface px-2 py-1"
      >
        {sheets.length === 0 ? (
          <span className="text-xs text-secondary">No sheets</span>
        ) : (
          sheets.map((s) => {
            const active = s.name === activeSheetName;
            return (
              <button
                key={s.id}
                type="button"
                data-testid={`sheet-tab-${s.name}`}
                onClick={() => setActiveSheetName(s.name)}
                className={cn(
                  "shrink-0 rounded px-3 py-1 text-xs font-medium transition-colors",
                  active
                    ? "bg-background text-foreground shadow-sm border border-divider"
                    : "text-secondary hover:text-foreground hover:bg-hover"
                )}
              >
                {s.name}
              </button>
            );
          })
        )}
      </div>

      <div className="agent-bar flex items-center gap-2 rounded-md border border-[var(--ai-violet-muted)] bg-[var(--ai-violet-light)]/40 px-2 py-1.5">
        <Sparkles size={14} className="text-[var(--ai-violet)] shrink-0" />
        <input
          data-testid="agent-prompt"
          aria-label="Agent prompt"
          value={agentPrompt}
          onChange={(e) => setAgentPrompt(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              void onAgentRun();
            }
          }}
          placeholder="Ask the agent to stage a change…"
          disabled={!agent || agentBusy}
          className="flex-1 bg-transparent px-1 py-1 text-xs text-foreground placeholder:text-tertiary focus:outline-none"
        />
        <Button
          size="sm"
          variant="accent"
          onClick={() => void onAgentRun()}
          disabled={!agent || agentBusy || agentPrompt.trim().length === 0}
          data-testid="agent-run"
          className="bg-[var(--ai-violet)] hover:bg-[var(--ai-violet)]/90"
        >
          {agentBusy ? <Loader2 className="animate-spin" size={14} /> : <Sparkles size={14} />}
          Propose
        </Button>
      </div>

      <div className="pointer-events-none fixed bottom-6 left-1/2 z-50 flex -translate-x-1/2 flex-col items-center gap-2">
        {toasts.map((t) => (
          <div
            key={t.id}
            role="status"
            className={cn(
              "pointer-events-auto rounded-md border px-3 py-1.5 text-xs shadow-sm",
              t.kind === "info" && "border-divider bg-surface text-foreground",
              t.kind === "warn" && "border-[var(--warning)] bg-[var(--warning)]/10 text-[var(--warning)]",
              t.kind === "error" &&
                "border-[var(--error-border)] bg-[var(--error-bg)] text-[var(--error-text)]"
            )}
          >
            {t.text}
          </div>
        ))}
      </div>
    </div>
  );
}

function formatCellValue(value: CellValue): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number") return String(value);
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  switch (value.kind) {
    case "error":
      return value.code;
    default: {
      const _exhaustive: never = value.kind;
      void _exhaustive;
      return "";
    }
  }
}

/**
 * Best-effort literal parsing for the formula bar / in-cell editor:
 *   - `""`        → null (clear cell)
 *   - `"123.45"`  → number
 *   - `"true"`    → boolean
 *   - everything else stays a string
 */
function parseLiteral(raw: string): CellValue {
  const t = raw.trim();
  if (t === "") return null;
  if (/^-?\d+(?:\.\d+)?$/.test(t)) {
    const n = Number(t);
    if (Number.isFinite(n)) return n;
  }
  const lower = t.toLowerCase();
  if (lower === "true") return true;
  if (lower === "false") return false;
  return raw;
}

/**
 * Pick a free-ish target cell for the demo agent recipe. Falls back to
 * the column after the populated bounding box, on the first empty row.
 */
function pickAgentTarget(sheet: Sheet): string {
  let maxRow = -1;
  let maxCol = -1;
  for (const c of sheet.cells.values()) {
    if (c.row > maxRow) maxRow = c.row;
    if (c.col > maxCol) maxCol = c.col;
  }
  const targetRow = Math.max(maxRow + 1, 0);
  const targetCol = Math.min(Math.max(maxCol + 1, 3), 25);
  return `${colToLetter(targetCol)}${targetRow + 1}`;
}
