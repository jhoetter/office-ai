"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { DragEvent, ReactNode } from "react";
import { FolderOpen, Loader2, Sparkles, Download } from "lucide-react";
import { Button, cn } from "@officeai/ui";
import {
  XlsxAgent,
  cellKey,
  colToLetter,
  flattenCellXf,
  formatA1,
  formatRange,
  type CellFormatPatch,
  type CellValue,
  type Sheet,
  type XlsxSnapshot,
} from "@officeai/xlsx";
import { buildSampleXlsx } from "@/lib/sample-xlsx";
import { Grid } from "./Grid";
import {
  formatSelection,
  isSingle,
  selectionToRange,
  singleSelection,
  type CellPos,
  type Selection,
} from "./selection";
import {
  FormulaSuggest,
  applySuggestion,
  getSuggestions,
} from "./FormulaSuggest";
import { Toolbar } from "./Toolbar";
import { formatCellValue as renderCellValue } from "./styles";

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
  const [selection, setSelection] = useState<Selection | null>(singleSelection({ row: 0, col: 0 }));
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
      setSelection(singleSelection({ row: 0, col: 0 }));
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
    // Formula bar / derived display always reflect the *anchor* cell
    // (Excel matches this — the active cell stays white in a range).
    return activeSheet.cells.get(cellKey(selection.anchor.row, selection.anchor.col)) ?? null;
  }, [activeSheet, selection]);

  const selectedRef = selection ? formatSelection(selection) : "";

  // Derived display for the formula bar when the user is NOT actively
  // editing it. While the input has focus we surface `formulaDraft`
  // (uncommitted user keystrokes) instead so the snapshot subscription
  // can't clobber typing.
  const derivedFormulaDisplay = (() => {
    if (!selectedCell) return "";
    if (selectedCell.formula) return `=${selectedCell.formula.text}`;
    // Formula bar shows the raw value (e.g. `0.25`, not `25.00%`) so
    // edits don't accidentally rewrite the underlying numeric value.
    return renderCellValue(selectedCell.value, 0);
  })();
  const formulaValue = formulaFocused ? formulaDraft : derivedFormulaDisplay;

  // Caret offset inside the formula-bar input. We snapshot it on every
  // selectionchange / keystroke so click-to-insert-ref knows where to
  // splice the picked cell reference. The ref is read by ref-insertion
  // logic (no re-render needed); the state shadow drives autocomplete
  // reactivity (matches list refreshes when the caret moves).
  const formulaCaretRef = useRef<number>(0);
  const [formulaCaret, setFormulaCaret] = useState(0);
  const captureCaret = useCallback(() => {
    const el = formulaInputRef.current;
    if (!el) return;
    const next = el.selectionStart ?? el.value.length;
    formulaCaretRef.current = next;
    setFormulaCaret(next);
  }, []);
  const [suggestHighlight, setSuggestHighlight] = useState(0);

  // We're in "formula edit mode" (Excel's "point mode") whenever the
  // formula bar is focused AND the draft starts with `=`. Cell clicks
  // in this mode insert the ref at the caret instead of moving the
  // selection.
  const formulaEditing = formulaFocused && formulaDraft.startsWith("=");

  // Pending range we're rendering inside the formula draft as the user
  // hovers over a cell after Shift+click. While `pendingRefSpan` is
  // non-null, the slice [from..to] of `formulaDraft` is the live cell
  // reference text being extended.
  const pendingRefSpanRef = useRef<{ from: number; to: number } | null>(null);
  // Anchor cell of the in-progress "point mode" range — captured on
  // the first cell mousedown after entering formula edit mode so a
  // subsequent Shift-click / drag can extend the ref into A1:C3.
  const pendingRefAnchorRef = useRef<CellPos | null>(null);

  const insertRefAtCaret = useCallback(
    (ref: string) => {
      const span = pendingRefSpanRef.current;
      const draft = formulaDraft;
      let next: string;
      let nextCaret: number;
      if (span) {
        next = draft.slice(0, span.from) + ref + draft.slice(span.to);
        nextCaret = span.from + ref.length;
      } else {
        const caret = formulaCaretRef.current;
        next = draft.slice(0, caret) + ref + draft.slice(caret);
        nextCaret = caret + ref.length;
      }
      pendingRefSpanRef.current = { from: nextCaret - ref.length, to: nextCaret };
      formulaCaretRef.current = nextCaret;
      setFormulaDraft(next);
      // Re-focus + place caret at the end of the inserted ref so
      // subsequent typing continues the formula.
      requestAnimationFrame(() => {
        const el = formulaInputRef.current;
        if (!el) return;
        el.focus();
        el.setSelectionRange(nextCaret, nextCaret);
      });
    },
    [formulaDraft]
  );

  const handleGridSelect = useCallback(
    (pos: CellPos, opts?: { extend?: boolean }) => {
      // Click-to-insert-ref: while the formula bar is in point mode,
      // a plain click inserts the cell ref at the caret; a drag /
      // Shift-click extends the previously inserted ref into a range.
      // The active selection (and therefore the formula's *target*
      // cell) is intentionally NOT moved here — Excel keeps the
      // original cell as the destination while the user picks refs.
      if (formulaEditing) {
        if (opts?.extend && pendingRefSpanRef.current) {
          // Build "anchor:focus" from the *first* clicked ref (stored
          // implicitly in pendingRefAnchorRef) and the new pos.
          const anchor = pendingRefAnchorRef.current ?? pos;
          const sel: Selection = {
            anchor,
            focus: pos,
          };
          const ref =
            isSingle(sel) ? formatA1(sel.anchor) : formatRange(selectionToRange(sel));
          insertRefAtCaret(ref);
        } else {
          pendingRefSpanRef.current = null;
          pendingRefAnchorRef.current = pos;
          insertRefAtCaret(formatA1(pos));
        }
        return;
      }

      // Normal (non-formula) selection behaviour.
      if (opts?.extend) {
        setSelection((prev) =>
          prev ? { anchor: prev.anchor, focus: pos } : singleSelection(pos)
        );
      } else {
        setSelection(singleSelection(pos));
      }
      // Pull keyboard focus back to the surface so the next printable
      // key starts type-to-edit on the new anchor. Focus synchronously
      // so the active element is already the surface by the time the
      // user's mouseup completes.
      surfaceRef.current?.focus({ preventScroll: true });
    },
    [formulaEditing, insertRefAtCaret]
  );

  // Grid-level keyboard handler: when nothing else has focus, a
  // printable key starts in-formula-bar editing for the active
  // single-cell anchor. F2 enters with the existing value; Backspace /
  // Delete clears the cell.
  const surfaceRef = useRef<HTMLDivElement | null>(null);
  const onSurfaceKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLDivElement>) => {
      // Don't steal keys destined for inputs / buttons / the formula
      // bar — they have their own onKeyDown.
      const target = e.target as HTMLElement;
      if (target.tagName === "INPUT" || target.tagName === "BUTTON" || target.isContentEditable) {
        return;
      }
      if (!selection || !isSingle(selection)) {
        // Multi-cell selections defer to dedicated commands later
        // (Delete is the one exception we want to keep).
        if (e.key !== "Backspace" && e.key !== "Delete") return;
      }

      if (e.key === "F2") {
        e.preventDefault();
        const fi = formulaInputRef.current;
        if (!fi) return;
        setFormulaDraft(derivedFormulaDisplay);
        setFormulaFocused(true);
        requestAnimationFrame(() => {
          fi.focus();
          const len = fi.value.length;
          fi.setSelectionRange(len, len);
        });
        return;
      }

      if (e.key === "Backspace" || e.key === "Delete") {
        if (!activeSheet || !selection) return;
        e.preventDefault();
        const a = agentRef.current;
        if (!a) return;
        // Range-aware clear: dispatch one set-cell-value over each
        // cell in the normalized selection. Simple loop is fine for
        // the in-app sizes we expect; a true range-clear command is a
        // future optimisation.
        const range = selectionToRange(selection);
        for (let r = range.start.row; r <= range.end.row; r++) {
          for (let c = range.start.col; c <= range.end.col; c++) {
            void a
              .applyCommand({
                type: "xlsx:set-cell-value",
                payload: {
                  sheet: activeSheet.name,
                  ref: formatA1({ row: r, col: c }),
                  value: null,
                },
                source: "human",
              })
              .catch((err: unknown) => {
                pushToast("error", err instanceof Error ? err.message : String(err));
              });
          }
        }
        return;
      }

      // Type-to-edit: a single printable key starts edit mode and
      // pre-fills the formula bar with that key.
      const isPrintable =
        e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey && e.key !== " " ||
        // Treat Space as an explicit edit-start (replaces existing
        // contents), since that matches Excel's behaviour.
        (e.key === " " && !e.ctrlKey && !e.metaKey && !e.altKey);
      if (!isPrintable) return;
      if (!selection || !isSingle(selection)) return;
      e.preventDefault();
      setFormulaDraft(e.key === " " ? "" : e.key);
      setFormulaFocused(true);
      requestAnimationFrame(() => {
        const fi = formulaInputRef.current;
        if (!fi) return;
        fi.focus();
        const len = fi.value.length;
        fi.setSelectionRange(len, len);
      });
    },
    [activeSheet, derivedFormulaDisplay, pushToast, selection]
  );

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
    // Formula-bar Enter applies to the *anchor* cell, mirroring Excel
    // (the moving end of the range doesn't receive the value).
    const ref = formatA1({ row: selection.anchor.row, col: selection.anchor.col });
    void dispatchCellEdit(activeSheet.name, ref, formulaDraft);
    setFormulaFocused(false);
    setFormulaDraft("");
    formulaInputRef.current?.blur();
  }, [activeSheet, selection, formulaDraft, dispatchCellEdit]);

  const onCommitGridEdit = useCallback(
    (pos: CellPos, value: string) => {
      if (!activeSheet) return;
      const ref = formatA1({ row: pos.row, col: pos.col });
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

  // Recompute autocomplete matches whenever the user types or moves
  // the caret while the formula bar is focused. `getSuggestions`
  // returns the empty list when the caret isn't on a partial function
  // token, which makes the popover render `null`.
  const { matches: suggestionMatches, active: suggestionSpan } = formulaFocused
    ? getSuggestions(formulaDraft, formulaCaret)
    : { matches: [], active: null };

  // Reset the highlight cursor whenever the prefix changes so the
  // first match is always selected by default. Done in render via a
  // ref-tracked previous-prefix shadow to avoid the lint-flagged
  // setState-in-effect pattern; the conditional setState is safe
  // because it short-circuits when the prefix is unchanged.
  const prevPrefixRef = useRef<string | null>(null);
  if (prevPrefixRef.current !== (suggestionSpan?.prefix ?? null)) {
    prevPrefixRef.current = suggestionSpan?.prefix ?? null;
    if (suggestHighlight !== 0) setSuggestHighlight(0);
  }

  const onApplyFormat = useCallback(
    (patch: CellFormatPatch) => {
      const a = agentRef.current;
      if (!a || !activeSheet || !selection) return;
      const range = formatSelection(selection);
      void a
        .applyCommand({
          type: "xlsx:set-cell-format",
          payload: {
            sheet: activeSheet.name,
            range,
            format: patch,
          },
          source: "human",
        })
        .catch((err: unknown) => {
          pushToast("error", err instanceof Error ? err.message : String(err));
        });
    },
    [activeSheet, selection, pushToast]
  );

  const dispatchOrToast = useCallback(
    (
      type:
        | "xlsx:merge-cells"
        | "xlsx:unmerge-cells"
        | "xlsx:insert-row"
        | "xlsx:insert-column"
        | "xlsx:delete-row"
        | "xlsx:delete-column",
      payload: Record<string, unknown>
    ) => {
      const a = agentRef.current;
      if (!a) return;
      void a
        .applyCommand({ type, payload, source: "human" } as Parameters<
          typeof a.applyCommand
        >[0])
        .catch((err: unknown) => {
          pushToast("error", err instanceof Error ? err.message : String(err));
        });
    },
    [pushToast]
  );

  // Range eligible for merge: at least 2 cells.
  const canMerge = !!(selection && !isSingle(selection));
  // Resolve the merge currently under the selection (if any). We
  // accept either an exact-match range OR a single cell that lives
  // inside a merge — Excel-style "click the merged surface, then
  // unmerge". The matched merge becomes the unmerge target so we
  // don't have to grow the selection ourselves.
  const matchedMerge = useMemo(() => {
    if (!activeSheet || !selection) return null;
    const n = selectionToRange(selection);
    return (
      activeSheet.merges.find(
        (m) =>
          m.r1 === n.start.row &&
          m.c1 === n.start.col &&
          m.r2 === n.end.row &&
          m.c2 === n.end.col
      ) ??
      activeSheet.merges.find(
        (m) =>
          isSingle(selection) &&
          n.start.row >= m.r1 &&
          n.start.row <= m.r2 &&
          n.start.col >= m.c1 &&
          n.start.col <= m.c2
      ) ??
      null
    );
  }, [activeSheet, selection]);
  const canUnmerge = !!matchedMerge;

  const onMerge = useCallback(() => {
    if (!activeSheet || !selection) return;
    dispatchOrToast("xlsx:merge-cells", {
      sheet: activeSheet.name,
      range: formatSelection(selection),
    });
  }, [activeSheet, selection, dispatchOrToast]);

  const onUnmerge = useCallback(() => {
    if (!activeSheet || !matchedMerge) return;
    const range = formatRange({
      start: { row: matchedMerge.r1, col: matchedMerge.c1 },
      end: { row: matchedMerge.r2, col: matchedMerge.c2 },
    });
    dispatchOrToast("xlsx:unmerge-cells", { sheet: activeSheet.name, range });
  }, [activeSheet, matchedMerge, dispatchOrToast]);

  const onInsertRowAbove = useCallback(() => {
    if (!activeSheet || !selection) return;
    const r = selectionToRange(selection);
    dispatchOrToast("xlsx:insert-row", {
      sheet: activeSheet.name,
      at: r.start.row + 1,
      count: 1,
    });
  }, [activeSheet, selection, dispatchOrToast]);

  const onInsertRowBelow = useCallback(() => {
    if (!activeSheet || !selection) return;
    const r = selectionToRange(selection);
    dispatchOrToast("xlsx:insert-row", {
      sheet: activeSheet.name,
      at: r.end.row + 2,
      count: 1,
    });
  }, [activeSheet, selection, dispatchOrToast]);

  const onInsertColumnLeft = useCallback(() => {
    if (!activeSheet || !selection) return;
    const r = selectionToRange(selection);
    dispatchOrToast("xlsx:insert-column", {
      sheet: activeSheet.name,
      at: r.start.col + 1,
      count: 1,
    });
  }, [activeSheet, selection, dispatchOrToast]);

  const onInsertColumnRight = useCallback(() => {
    if (!activeSheet || !selection) return;
    const r = selectionToRange(selection);
    dispatchOrToast("xlsx:insert-column", {
      sheet: activeSheet.name,
      at: r.end.col + 2,
      count: 1,
    });
  }, [activeSheet, selection, dispatchOrToast]);

  const onDeleteRow = useCallback(() => {
    if (!activeSheet || !selection) return;
    const r = selectionToRange(selection);
    dispatchOrToast("xlsx:delete-row", {
      sheet: activeSheet.name,
      at: r.start.row + 1,
      count: r.end.row - r.start.row + 1,
    });
  }, [activeSheet, selection, dispatchOrToast]);

  const onDeleteColumn = useCallback(() => {
    if (!activeSheet || !selection) return;
    const r = selectionToRange(selection);
    dispatchOrToast("xlsx:delete-column", {
      sheet: activeSheet.name,
      at: r.start.col + 1,
      count: r.end.col - r.start.col + 1,
    });
  }, [activeSheet, selection, dispatchOrToast]);

  const acceptSuggestion = useCallback(
    (info: Parameters<typeof applySuggestion>[1]) => {
      if (!suggestionSpan) return;
      const { next, caret } = applySuggestion(formulaDraft, info, suggestionSpan);
      setFormulaDraft(next);
      formulaCaretRef.current = caret;
      setFormulaCaret(caret);
      // Re-focus + park caret inside the parens so the user can keep
      // typing arguments without grabbing the mouse.
      requestAnimationFrame(() => {
        const el = formulaInputRef.current;
        if (!el) return;
        el.focus();
        el.setSelectionRange(caret, caret);
      });
    },
    [formulaDraft, suggestionSpan]
  );

  return (
    <div
      ref={surfaceRef}
      tabIndex={0}
      onKeyDown={onSurfaceKeyDown}
      className={cn(
        "xlsx-editor relative flex h-full min-h-0 flex-col gap-3 outline-none",
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

      {snapshot ? (
        <Toolbar
          disabled={!agent || !selection}
          anchorStyleId={selectedCell?.styleId}
          styles={snapshot.root.styles}
          selection={selection}
          onApply={onApplyFormat}
          canMerge={canMerge}
          canUnmerge={canUnmerge}
          onMerge={onMerge}
          onUnmerge={onUnmerge}
          onInsertRowAbove={onInsertRowAbove}
          onInsertRowBelow={onInsertRowBelow}
          onInsertColumnLeft={onInsertColumnLeft}
          onInsertColumnRight={onInsertColumnRight}
          onDeleteRow={onDeleteRow}
          onDeleteColumn={onDeleteColumn}
        />
      ) : null}

      <div className="formula-bar relative flex items-center gap-2 rounded-md border border-divider bg-surface px-2 py-1.5">
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
          onChange={(e) => {
            // A user keystroke invalidates the click-to-insert pending
            // span — anything they type from here adds to / replaces
            // the formula instead of extending the picked ref.
            pendingRefSpanRef.current = null;
            pendingRefAnchorRef.current = null;
            setFormulaDraft(e.target.value);
            captureCaret();
          }}
          onSelect={captureCaret}
          onClick={captureCaret}
          onFocus={() => {
            // Only seed the draft from the resolved cell value when the
            // user is focusing the bar fresh (mouse click, Tab). When
            // type-to-edit has already pre-filled `formulaDraft`, leave
            // it alone — otherwise the just-typed character would be
            // clobbered by the cell's prior value.
            if (formulaDraft === "") setFormulaDraft(derivedFormulaDisplay);
            setFormulaFocused(true);
            requestAnimationFrame(captureCaret);
          }}
          onBlur={() => {
            setFormulaFocused(false);
            setFormulaDraft("");
            pendingRefSpanRef.current = null;
            pendingRefAnchorRef.current = null;
          }}
          onKeyDown={(e) => {
            const hasSuggestions = suggestionMatches.length > 0;
            if (hasSuggestions && (e.key === "ArrowDown" || e.key === "ArrowUp")) {
              e.preventDefault();
              setSuggestHighlight((prev) => {
                const dir = e.key === "ArrowDown" ? 1 : -1;
                const n = suggestionMatches.length;
                return (prev + dir + n) % n;
              });
              return;
            }
            if (hasSuggestions && (e.key === "Tab" || e.key === "Enter")) {
              // Enter accepts a suggestion only while the popover is
              // open; otherwise it submits the formula.
              const pick = suggestionMatches[Math.min(suggestHighlight, suggestionMatches.length - 1)];
              if (pick) {
                e.preventDefault();
                acceptSuggestion(pick);
                return;
              }
            }
            if (e.key === "Enter") {
              e.preventDefault();
              onFormulaSubmit();
            } else if (e.key === "Escape") {
              e.preventDefault();
              setFormulaFocused(false);
              setFormulaDraft("");
              pendingRefSpanRef.current = null;
              pendingRefAnchorRef.current = null;
              formulaInputRef.current?.blur();
              surfaceRef.current?.focus();
            } else {
              captureCaret();
            }
          }}
          placeholder={selection ? "Type a value or =formula" : "Select a cell to edit"}
          disabled={!selection || !agent}
          className="flex-1 bg-transparent px-1 py-1 text-xs text-foreground placeholder:text-tertiary focus:outline-none"
        />
        <div className="absolute left-[68px] right-2 top-full z-40">
          <FormulaSuggest
            matches={suggestionMatches}
            highlight={Math.min(suggestHighlight, Math.max(suggestionMatches.length - 1, 0))}
            onPick={acceptSuggestion}
            onHighlight={setSuggestHighlight}
          />
        </div>
      </div>

      <div className="relative flex-1 min-h-0">
        {activeSheet && snapshot ? (
          <Grid
            sheet={activeSheet}
            styles={snapshot.root.styles}
            selection={selection}
            onSelect={handleGridSelect}
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
