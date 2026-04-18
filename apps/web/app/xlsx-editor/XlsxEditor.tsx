"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { DragEvent, ReactNode } from "react";
import { FolderOpen, Loader2, Download } from "lucide-react";
import { Button, CommentComposer, CommentsSidebar, TextFormatBar, cn } from "@officeai/ui";
import { createXlsxCommentsProvider } from "./xlsxCommentsProvider";
import {
  XlsxAgent,
  assignRefColors,
  cellKey,
  flattenCellXf,
  formatA1,
  formatRange,
  tokenizeForDisplay,
  type CellFormatPatch,
  type CellValue,
  type DisplayToken,
  type Sheet,
  type StyleTable,
  type XlsxSnapshot,
} from "@officeai/xlsx";
import type { ActiveTextFormat, TextFormatProvider } from "@officeai/text-formatting";
import { computeXlsxActive, createXlsxFormatProvider } from "./xlsxFormatProvider";
import { buildSampleXlsx } from "@/lib/sample-xlsx";
import { Grid, type RefRect, type GridContextTarget, type MarchingAntsRect } from "./Grid";
import { ContextMenu, type ContextMenuItem } from "./ContextMenu";
import { FormulaHighlight } from "./FormulaHighlight";
import {
  formatSelection,
  isSingle,
  selectionToRange,
  singleSelection,
  type CellPos,
  type Selection,
} from "./selection";
import { FormulaSuggest, applySuggestion, getSuggestions } from "./FormulaSuggest";
import { Toolbar } from "./Toolbar";
import { useShortcutsDialog } from "@/lib/shortcuts/useShortcutsDialog";
import { KeyboardShortcutsDialog } from "@/lib/shortcuts/KeyboardShortcutsDialog";
import { TextToColumnsPopover } from "./TextToColumnsPopover";
import { sniffDelimiter } from "@officeai/xlsx";
import { formatCellValue as renderCellValue } from "./styles";
import {
  marshalClipboard,
  parseClipboardPayload,
  writeToSystemClipboard,
  readFromSystemClipboard,
} from "./clipboard";

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
 *   3. Render header → formula bar → grid → sheet tabs.
 *
 * All cell mutations dispatch through `agent.applyCommand` so the
 * single command-bus invariant holds for both human edits and any
 * external agent driving the same `XlsxAgent` over the headless
 * `office-agent` CLI. The editor surface itself is human-only — agent
 * affordances live in the CLI, not the UI.
 */
export function XlsxEditor(): ReactNode {
  const agentRef = useRef<XlsxAgent | null>(null);
  const [agent, setAgent] = useState<XlsxAgent | null>(null);
  const [snapshot, setSnapshot] = useState<XlsxSnapshot | null>(null);
  const [activeSheetName, setActiveSheetName] = useState<string | null>(null);
  const [selection, setSelection] = useState<Selection | null>(singleSelection({ row: 0, col: 0 }));
  const [pendingCount, setPendingCount] = useState(0);
  const [toasts, setToasts] = useState<ToastMessage[]>([]);
  const [formulaDraft, setFormulaDraft] = useState("");
  const [formulaFocused, setFormulaFocused] = useState(false);
  const formulaInputRef = useRef<HTMLInputElement | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [filename, setFilename] = useState<string>(SAMPLE_NAME);
  const [dragOver, setDragOver] = useState(false);
  const [ctxMenu, setCtxMenu] = useState<{
    target: GridContextTarget;
    x: number;
    y: number;
  } | null>(null);
  // Phase 13d — clipboard source overlay ("marching ants"). Tracks the
  // last range Cmd+C / Cmd+X copied from THIS app so we can render
  // the dashed border AND, on a follow-up Cmd+V, clear the source if
  // the original op was a Cut. Cleared on Escape or any model edit.
  const [marchingAnts, setMarchingAnts] = useState<(MarchingAntsRect & { readonly sheet: string }) | null>(
    null
  );
  const shortcutsDialog = useShortcutsDialog();

  const toastIdRef = useRef(0);
  const pushToast = useCallback((kind: ToastMessage["kind"], text: string) => {
    const id = ++toastIdRef.current;
    setToasts((prev) => [...prev, { id, kind, text }]);
    setTimeout(() => setToasts((prev) => prev.filter((t) => t.id !== id)), 4000);
  }, []);

  // Refs that mirror the latest selection / sheet / styles so the
  // shared TextFormatBar provider can read them at event-handler
  // time without us having to rebuild the provider on every render.
  // Updated below after the relevant `useMemo`s have run.
  const selectionRef = useRef<Selection | null>(selection);
  const activeSheetRef = useRef<Sheet | null>(null);
  const stylesRef = useRef<StyleTable | null>(null);

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

  // Mirror React state into refs so the shared TextFormatBar provider
  // (built once via the lazy useState below) can read the latest
  // selection / sheet / style table when the user clicks a control.
  useEffect(() => {
    selectionRef.current = selection;
  }, [selection]);
  useEffect(() => {
    activeSheetRef.current = activeSheet;
  }, [activeSheet]);
  useEffect(() => {
    stylesRef.current = snapshot?.root.styles ?? null;
  }, [snapshot]);

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

  // Track the formula input's horizontal scroll so the colour
  // overlay underneath stays aligned for long formulas. Updated on
  // every input scroll event below.
  const [formulaScrollLeft, setFormulaScrollLeft] = useState(0);

  // Phase 12a/c — tokenise the formula bar contents (whether it's a
  // live draft or the resolved derived display) so we can paint
  // colored ref tokens in the bar AND draw matching coloured borders
  // on the referenced cells in the grid. We only highlight when the
  // value is actually a formula (`=…`) — plain literals get the
  // default text colour everywhere.
  const formulaTokens: ReadonlyArray<DisplayToken> = useMemo(() => {
    if (!formulaValue.startsWith("=")) return [];
    return tokenizeForDisplay(formulaValue);
  }, [formulaValue]);
  const refColors = useMemo(() => assignRefColors(formulaTokens), [formulaTokens]);

  const refRects: ReadonlyArray<RefRect> = useMemo(() => {
    if (formulaTokens.length === 0) return [];
    if (!activeSheet) return [];
    const out: RefRect[] = [];
    const seen = new Set<string>();
    for (const t of formulaTokens) {
      if (!t.target || !t.refKey) continue;
      // Skip cross-sheet refs: only colour rects on the active sheet.
      // (Highlighting other sheets would require navigating tabs.)
      if (t.target.sheet && t.target.sheet !== activeSheet.name) continue;
      if (seen.has(t.refKey)) continue;
      seen.add(t.refKey);
      const color = refColors.get(t.refKey);
      if (!color) continue;
      if (t.target.kind === "ref") {
        out.push({ r1: t.target.row, c1: t.target.col, r2: t.target.row, c2: t.target.col, color });
      } else {
        out.push({ r1: t.target.r1, c1: t.target.c1, r2: t.target.r2, c2: t.target.c2, color });
      }
    }
    return out;
  }, [formulaTokens, refColors, activeSheet]);

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

  // Total grid bounds — must stay in sync with `Grid.tsx`. These
  // bound arrow / Home / Ctrl+End navigation; the underlying model
  // accepts arbitrary indices, but rendering only goes this far.
  const GRID_ROWS = 1000;
  const GRID_COLS = 26;

  const handleAxisSelect = useCallback((axis: "row" | "col", index: number, opts?: { extend?: boolean }) => {
    // Row click → select the entire row (col 0 .. GRID_COLS-1).
    // Column click → select the entire column (row 0 .. GRID_ROWS-1).
    // Shift-click extends from the existing anchor along the same
    // axis so users can rubber-band multi-row / multi-col ranges.
    setSelection((prev) => {
      const focus: CellPos =
        axis === "row" ? { row: index, col: GRID_COLS - 1 } : { row: GRID_ROWS - 1, col: index };
      const anchor: CellPos = axis === "row" ? { row: index, col: 0 } : { row: 0, col: index };
      if (opts?.extend && prev) {
        // Keep the prior anchor; replace the focus on the matching
        // axis only (so a row-select extends rows, col-select cols).
        if (axis === "row") {
          return { anchor: { row: prev.anchor.row, col: 0 }, focus };
        }
        return { anchor: { row: 0, col: prev.anchor.col }, focus };
      }
      return { anchor, focus };
    });
    surfaceRef.current?.focus({ preventScroll: true });
  }, []);

  // Detect whether the current selection covers entire rows / cols
  // — used to decide whether Cmd/Ctrl+− deletes a row or a column.
  const wholeRowSelection = useMemo(() => {
    if (!selection) return false;
    const r = selectionToRange(selection);
    return r.start.col === 0 && r.end.col >= GRID_COLS - 1;
  }, [selection]);
  const wholeColSelection = useMemo(() => {
    if (!selection) return false;
    const r = selectionToRange(selection);
    return r.start.row === 0 && r.end.row >= GRID_ROWS - 1;
  }, [selection]);

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
          const ref = isSingle(sel) ? formatA1(sel.anchor) : formatRange(selectionToRange(sel));
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
        setSelection((prev) => (prev ? { anchor: prev.anchor, focus: pos } : singleSelection(pos)));
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
  // Move the active selection one cell in the given direction, with
  // optional Shift-extend. Pure helper — no side effects beyond
  // calling setSelection.
  const moveSelection = useCallback((dRow: number, dCol: number, opts: { extend: boolean }) => {
    setSelection((prev) => {
      const base: CellPos = prev?.focus ?? { row: 0, col: 0 };
      const next: CellPos = {
        row: Math.max(0, Math.min(GRID_ROWS - 1, base.row + dRow)),
        col: Math.max(0, Math.min(GRID_COLS - 1, base.col + dCol)),
      };
      if (opts.extend && prev) return { anchor: prev.anchor, focus: next };
      return singleSelection(next);
    });
  }, []);

  // Cmd/Ctrl+arrow Excel-style "jump to data edge". When stationed
  // on a non-empty cell, jump to the last non-empty cell in the run;
  // when stationed on an empty cell, jump to the next non-empty one.
  // Falls back to the grid edge when no transition is found.
  const jumpToDataEdge = useCallback(
    (dRow: number, dCol: number, opts: { extend: boolean }) => {
      if (!activeSheet) return;
      setSelection((prev) => {
        const base: CellPos = prev?.focus ?? { row: 0, col: 0 };
        const isFilled = (r: number, c: number): boolean => {
          if (r < 0 || c < 0 || r >= GRID_ROWS || c >= GRID_COLS) return false;
          const cell = activeSheet.cells.get(cellKey(r, c));
          return !!cell && cell.value !== null && cell.value !== undefined;
        };
        const startFilled = isFilled(base.row, base.col);
        let r = base.row;
        let c = base.col;
        const step = (): boolean => {
          const nr = r + dRow;
          const nc = c + dCol;
          if (nr < 0 || nc < 0 || nr >= GRID_ROWS || nc >= GRID_COLS) return false;
          r = nr;
          c = nc;
          return true;
        };
        if (startFilled) {
          // Walk while the *next* cell is also filled; stop just
          // before a transition into emptiness.
          while (isFilled(r + dRow, c + dCol)) {
            if (!step()) break;
          }
        } else {
          // Walk until we hit the next filled cell or the edge.
          while (step()) {
            if (isFilled(r, c)) break;
          }
        }
        const next: CellPos = { row: r, col: c };
        if (opts.extend && prev) return { anchor: prev.anchor, focus: next };
        return singleSelection(next);
      });
    },
    [activeSheet]
  );

  /**
   * Capture the current selection into a {@link XlsxClipboardSnapshot}
   * and write the TSV + HTML pair to the system clipboard. Also
   * primes `marchingAnts` so the Grid draws the source overlay.
   */
  const copySelection = useCallback(
    async (mode: "copy" | "cut"): Promise<boolean> => {
      const a = agentRef.current;
      if (!a || !activeSheet || !selection) return false;
      const range = selectionToRange(selection);
      let snap;
      try {
        snap = a.getClipboardSnapshot({
          sheet: activeSheet.name,
          range: formatRange(range),
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
        return false;
      }
      const payload = marshalClipboard(snap);
      try {
        await writeToSystemClipboard(payload);
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
        return false;
      }
      setMarchingAnts({
        sheet: activeSheet.name,
        r1: range.start.row,
        c1: range.start.col,
        r2: range.end.row,
        c2: range.end.col,
        mode,
      });
      return true;
    },
    [activeSheet, selection, pushToast]
  );

  /**
   * Try the synchronous `event.clipboardData` channel first (works
   * inside a real `paste` event handler). Falls back to the async
   * `navigator.clipboard.read()` permission dance for keyboard
   * shortcut handlers that don't sit inside a paste event.
   */
  const pasteAtSelection = useCallback(
    async (direct?: { html?: string | null; text?: string | null }): Promise<boolean> => {
      const a = agentRef.current;
      if (!a || !activeSheet || !selection) return false;
      const target = formatA1(selection.anchor);
      const snap = direct ? parseClipboardPayload(direct) : await readFromSystemClipboard();
      if (!snap) {
        pushToast("warn", "Clipboard is empty.");
        return false;
      }
      try {
        await a.applyCommand({
          type: "xlsx:paste-range",
          payload: { sheet: activeSheet.name, target, source: snap },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
        return false;
      }

      // If the source was a Cut from THIS app, clear the source range
      // now (Excel parity — Cut doesn't actually mutate until Paste).
      if (marchingAnts?.mode === "cut" && marchingAnts.sheet === activeSheet.name) {
        const r0 = Math.min(marchingAnts.r1, marchingAnts.r2);
        const r1 = Math.max(marchingAnts.r1, marchingAnts.r2);
        const c0 = Math.min(marchingAnts.c1, marchingAnts.c2);
        const c1 = Math.max(marchingAnts.c1, marchingAnts.c2);
        for (let r = r0; r <= r1; r++) {
          for (let c = c0; c <= c1; c++) {
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
      }
      setMarchingAnts(null);

      // Move the selection to cover the pasted block so subsequent
      // Cmd+V / arrow keys feel "Excel-y".
      const end: CellPos = {
        row: selection.anchor.row + Math.max(0, snap.height - 1),
        col: selection.anchor.col + Math.max(0, snap.width - 1),
      };
      setSelection({ anchor: selection.anchor, focus: end });
      return true;
    },
    [activeSheet, selection, marchingAnts, pushToast]
  );

  // Native `copy` / `cut` / `paste` events fire on the focused element
  // and bubble. The surface div is `tabIndex=0`, so when the user
  // hits Cmd+C/X/V outside an input we receive them here. We
  // `preventDefault` and use the synchronous `event.clipboardData`
  // channel which avoids the async permission dance entirely.
  const onSurfaceCopy = useCallback(
    (e: React.ClipboardEvent<HTMLDivElement>) => {
      const tgt = e.target as HTMLElement;
      if (tgt.tagName === "INPUT" || tgt.tagName === "BUTTON" || tgt.isContentEditable) return;
      const a = agentRef.current;
      if (!a || !activeSheet || !selection) return;
      e.preventDefault();
      try {
        const range = selectionToRange(selection);
        const snap = a.getClipboardSnapshot({
          sheet: activeSheet.name,
          range: formatRange(range),
        });
        const payload = marshalClipboard(snap);
        e.clipboardData.setData("text/plain", payload.tsv);
        e.clipboardData.setData("text/html", payload.html);
        setMarchingAnts({
          sheet: activeSheet.name,
          r1: range.start.row,
          c1: range.start.col,
          r2: range.end.row,
          c2: range.end.col,
          mode: "copy",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [activeSheet, selection, pushToast]
  );

  const onSurfaceCut = useCallback(
    (e: React.ClipboardEvent<HTMLDivElement>) => {
      const tgt = e.target as HTMLElement;
      if (tgt.tagName === "INPUT" || tgt.tagName === "BUTTON" || tgt.isContentEditable) return;
      const a = agentRef.current;
      if (!a || !activeSheet || !selection) return;
      e.preventDefault();
      try {
        const range = selectionToRange(selection);
        const snap = a.getClipboardSnapshot({
          sheet: activeSheet.name,
          range: formatRange(range),
        });
        const payload = marshalClipboard(snap);
        e.clipboardData.setData("text/plain", payload.tsv);
        e.clipboardData.setData("text/html", payload.html);
        setMarchingAnts({
          sheet: activeSheet.name,
          r1: range.start.row,
          c1: range.start.col,
          r2: range.end.row,
          c2: range.end.col,
          mode: "cut",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [activeSheet, selection, pushToast]
  );

  const onSurfacePaste = useCallback(
    (e: React.ClipboardEvent<HTMLDivElement>) => {
      const tgt = e.target as HTMLElement;
      if (tgt.tagName === "INPUT" || tgt.tagName === "BUTTON" || tgt.isContentEditable) return;
      if (!agentRef.current || !activeSheet || !selection) return;
      e.preventDefault();
      const html = e.clipboardData.getData("text/html");
      const text = e.clipboardData.getData("text/plain");
      void pasteAtSelection({ html, text });
    },
    [activeSheet, selection, pasteAtSelection]
  );

  const onSurfaceKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLDivElement>) => {
      // Don't steal keys destined for inputs / buttons / the formula
      // bar — they have their own onKeyDown.
      const target = e.target as HTMLElement;
      if (target.tagName === "INPUT" || target.tagName === "BUTTON" || target.isContentEditable) {
        return;
      }

      // ── Navigation keys (work whether or not we have a single-cell
      // selection — extending a range is the whole point).
      const arrowDelta: Record<string, [number, number]> = {
        ArrowUp: [-1, 0],
        ArrowDown: [1, 0],
        ArrowLeft: [0, -1],
        ArrowRight: [0, 1],
      };
      const arrow = arrowDelta[e.key];
      if (arrow) {
        e.preventDefault();
        if (e.metaKey || e.ctrlKey) {
          jumpToDataEdge(arrow[0], arrow[1], { extend: e.shiftKey });
        } else {
          moveSelection(arrow[0], arrow[1], { extend: e.shiftKey });
        }
        return;
      }

      if (e.key === "Home") {
        e.preventDefault();
        setSelection((prev) => {
          if (!prev) return singleSelection({ row: 0, col: 0 });
          if (e.metaKey || e.ctrlKey) {
            return e.shiftKey
              ? { anchor: prev.anchor, focus: { row: 0, col: 0 } }
              : singleSelection({ row: 0, col: 0 });
          }
          const focus: CellPos = { row: prev.focus.row, col: 0 };
          return e.shiftKey ? { anchor: prev.anchor, focus } : singleSelection(focus);
        });
        return;
      }

      if (e.key === "End" && (e.metaKey || e.ctrlKey)) {
        // Ctrl+End → bottom-right of the *used* range (proxy: max
        // row/col across non-empty cells; falls back to A1).
        e.preventDefault();
        if (!activeSheet) return;
        let maxRow = 0;
        let maxCol = 0;
        for (const cell of activeSheet.cells.values()) {
          if (cell.value === null || cell.value === undefined) continue;
          if (cell.row > maxRow) maxRow = cell.row;
          if (cell.col > maxCol) maxCol = cell.col;
        }
        const focus: CellPos = { row: maxRow, col: maxCol };
        setSelection((prev) =>
          e.shiftKey && prev ? { anchor: prev.anchor, focus } : singleSelection(focus)
        );
        return;
      }

      if (e.key === "Tab") {
        e.preventDefault();
        moveSelection(0, e.shiftKey ? -1 : 1, { extend: false });
        return;
      }

      if (e.key === "Enter") {
        e.preventDefault();
        moveSelection(e.shiftKey ? -1 : 1, 0, { extend: false });
        return;
      }

      if (e.key === "F2") {
        if (!selection || !isSingle(selection)) return;
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

      if (e.key === "Escape") {
        // Plain Escape on the surface clears the selection back to a
        // single anchor — handy after Shift-extending a range. It
        // also dismisses the clipboard "marching ants" overlay, which
        // mirrors Excel's behaviour exactly.
        if (marchingAnts) {
          e.preventDefault();
          setMarchingAnts(null);
          return;
        }
        if (!selection) return;
        e.preventDefault();
        setSelection(singleSelection(selection.anchor));
        return;
      }

      // Undo / Redo — match Excel exactly:
      //   Cmd/Ctrl+Z       → undo
      //   Cmd/Ctrl+Shift+Z → redo (Cmd/Ctrl+Y also accepted)
      if ((e.metaKey || e.ctrlKey) && (e.key === "z" || e.key === "Z")) {
        e.preventDefault();
        const a = agentRef.current;
        if (!a) return;
        if (e.shiftKey) {
          if (a.canRedo()) a.redo();
        } else {
          if (a.canUndo()) a.undo();
        }
        return;
      }
      if ((e.metaKey || e.ctrlKey) && (e.key === "y" || e.key === "Y")) {
        e.preventDefault();
        const a = agentRef.current;
        if (!a) return;
        if (a.canRedo()) a.redo();
        return;
      }

      // Inline marks: Cmd/Ctrl + B / I / U toggle the mark over the
      // current selection. The active anchor's effective style drives
      // the toggle direction so a second press flips back, matching
      // Excel exactly. Skipped when no real selection / no agent.
      if ((e.metaKey || e.ctrlKey) && !e.shiftKey && !e.altKey) {
        const markKey =
          e.key === "b" || e.key === "B"
            ? "bold"
            : e.key === "i" || e.key === "I"
              ? "italic"
              : e.key === "u" || e.key === "U"
                ? "underline"
                : null;
        if (markKey && activeSheet && selection) {
          e.preventDefault();
          const a = agentRef.current;
          if (!a) return;
          const styleTable = stylesRef.current;
          // Probe the anchor cell's effective style to flip the
          // toggle direction. When the styles table isn't loaded
          // (early frames) treat the mark as "off" so the first
          // press always turns it on.
          const eff = styleTable ? flattenCellXf(styleTable, selectedCell?.styleId) : null;
          const currentlyOn = Boolean(
            (eff?.font as { bold?: boolean; italic?: boolean; underline?: unknown } | undefined)?.[markKey]
          );
          const range = formatSelection(selection);
          void a
            .applyCommand({
              type: "xlsx:set-cell-format",
              payload: {
                sheet: activeSheet.name,
                range,
                format: { font: { [markKey]: !currentlyOn } } as never,
              },
              source: "human",
            })
            .catch((err: unknown) =>
              pushToast("error", err instanceof Error ? err.message : String(err))
            );
          return;
        }
      }

      // Number-format shortcuts. Use `event.code` so Shift+Digit5
      // (which yields "%") still maps to the physical "5" key. We
      // dispatch the *built-in* numFmtId by id (as a numeric string)
      // for Number/Percent so the grid renderer's `formatNumber()`
      // recognises it via its 0..49 fast-path. Currency is a custom
      // format string because no built-in id renders the prefix the
      // way our toolbar advertises.
      if ((e.metaKey || e.ctrlKey) && e.shiftKey && !e.altKey && activeSheet && selection) {
        const numberFormat =
          e.code === "Digit1"
            ? "4" // #,##0.00 (built-in)
            : e.code === "Digit4"
              ? "$#,##0.00"
              : e.code === "Digit5"
                ? "9" // 0% (built-in)
                : null;
        if (numberFormat) {
          e.preventDefault();
          const a = agentRef.current;
          if (!a) return;
          const range = formatSelection(selection);
          void a
            .applyCommand({
              type: "xlsx:set-cell-format",
              payload: { sheet: activeSheet.name, range, format: { numberFormat } },
              source: "human",
            })
            .catch((err: unknown) =>
              pushToast("error", err instanceof Error ? err.message : String(err))
            );
          return;
        }
      }

      if (e.key === "Backspace" || e.key === "Delete") {
        if (!activeSheet || !selection) return;
        e.preventDefault();
        const a = agentRef.current;
        if (!a) return;
        const range = selectionToRange(selection);

        // Whole-row / whole-col selection → Delete actually drops
        // the rows / cols from the sheet (matches the user's
        // Excel-adjacent muscle memory: "select row → Delete →
        // row gone"). For partial selections we fall back to the
        // range-clear behaviour below.
        if (wholeRowSelection) {
          const count = range.end.row - range.start.row + 1;
          void a
            .applyCommand({
              type: "xlsx:delete-row",
              payload: { sheet: activeSheet.name, at: range.start.row + 1, count },
              source: "human",
            })
            .catch((err: unknown) => pushToast("error", err instanceof Error ? err.message : String(err)));
          return;
        }
        if (wholeColSelection) {
          const count = range.end.col - range.start.col + 1;
          void a
            .applyCommand({
              type: "xlsx:delete-column",
              payload: { sheet: activeSheet.name, at: range.start.col + 1, count },
              source: "human",
            })
            .catch((err: unknown) => pushToast("error", err instanceof Error ? err.message : String(err)));
          return;
        }

        // Range-aware clear: dispatch one set-cell-value over each
        // cell in the normalized selection. Simple loop is fine for
        // the in-app sizes we expect; a true range-clear command is a
        // future optimisation.
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
      // pre-fills the formula bar with that key. Only on a single-
      // cell anchor so we don't accidentally clobber a multi-cell
      // selection.
      if (!selection || !isSingle(selection)) return;
      const isPrintable =
        (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey && e.key !== " ") ||
        // Treat Space as an explicit edit-start (replaces existing
        // contents), since that matches Excel's behaviour.
        (e.key === " " && !e.ctrlKey && !e.metaKey && !e.altKey);
      if (!isPrintable) return;
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
    [
      activeSheet,
      derivedFormulaDisplay,
      jumpToDataEdge,
      marchingAnts,
      moveSelection,
      pushToast,
      selectedCell,
      selection,
      wholeColSelection,
      wholeRowSelection,
    ]
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

  const onFormulaSubmit = useCallback(
    (move: { row: number; col: number } = { row: 1, col: 0 }) => {
      if (!activeSheet || !selection) return;
      // Formula-bar Enter applies to the *anchor* cell, mirroring Excel
      // (the moving end of the range doesn't receive the value).
      const anchor = selection.anchor;
      const ref = formatA1({ row: anchor.row, col: anchor.col });
      void dispatchCellEdit(activeSheet.name, ref, formulaDraft);
      setFormulaFocused(false);
      setFormulaDraft("");
      formulaInputRef.current?.blur();
      // Move the selection in Excel-style: Enter→down, Shift+Enter→up,
      // Tab→right, Shift+Tab→left. Caller passes the delta.
      const next: CellPos = {
        row: Math.max(0, Math.min(GRID_ROWS - 1, anchor.row + move.row)),
        col: Math.max(0, Math.min(GRID_COLS - 1, anchor.col + move.col)),
      };
      setSelection(singleSelection(next));
      surfaceRef.current?.focus({ preventScroll: true });
    },
    [activeSheet, selection, formulaDraft, dispatchCellEdit]
  );

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
        | "xlsx:delete-column"
        | "xlsx:set-column-width"
        | "xlsx:set-row-height"
        | "xlsx:text-to-columns"
        | "xlsx:fill-range",
      payload: Record<string, unknown>
    ) => {
      const a = agentRef.current;
      if (!a) return;
      void a
        .applyCommand({ type, payload, source: "human" } as Parameters<typeof a.applyCommand>[0])
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
        (m) => m.r1 === n.start.row && m.c1 === n.start.col && m.r2 === n.end.row && m.c2 === n.end.col
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

  // P13g — Smart fill handle. Grid calls back once on mouse-up with
  // source/target/direction; we forward to xlsx:fill-range.
  const onFill = useCallback(
    (args: {
      source: { r1: number; c1: number; r2: number; c2: number };
      target: { r1: number; c1: number; r2: number; c2: number };
      direction: "down" | "right" | "up" | "left";
    }) => {
      if (!activeSheet) return;
      const sourceRange = formatRange({
        start: { row: args.source.r1, col: args.source.c1 },
        end: { row: args.source.r2, col: args.source.c2 },
      });
      const targetRange = formatRange({
        start: { row: args.target.r1, col: args.target.c1 },
        end: { row: args.target.r2, col: args.target.c2 },
      });
      if (sourceRange === targetRange) return;
      dispatchOrToast("xlsx:fill-range", {
        sheet: activeSheet.name,
        source: sourceRange,
        target: targetRange,
        direction: args.direction,
      });
    },
    [activeSheet, dispatchOrToast]
  );

  // P13f — Text to Columns popover state.
  const [ttocOpen, setTtocOpen] = useState(false);
  const ttocDefaultDelim = useMemo(() => {
    if (!activeSheet || !selection) return ",";
    const r = selectionToRange(selection);
    const sample = activeSheet.cells.get(cellKey(r.start.row, r.start.col))?.value;
    if (typeof sample === "string" && sample.length > 0) return sniffDelimiter(sample);
    return ",";
  }, [activeSheet, selection]);
  const canTextToColumns = !!(
    activeSheet &&
    selection &&
    selection.anchor.col === selection.focus.col
  );
  const onTextToColumns = useCallback(() => {
    if (!canTextToColumns) return;
    setTtocOpen(true);
  }, [canTextToColumns]);
  const onTextToColumnsConfirm = useCallback(
    (opts: { delimiter: string; treatConsecutiveAsOne: boolean }) => {
      setTtocOpen(false);
      if (!activeSheet || !selection) return;
      dispatchOrToast("xlsx:text-to-columns", {
        sheet: activeSheet.name,
        range: formatSelection(selection),
        delimiter: opts.delimiter,
        treatConsecutiveAsOne: opts.treatConsecutiveAsOne,
      });
    },
    [activeSheet, selection, dispatchOrToast]
  );

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

  const onResizeColumn = useCallback(
    (col: number, widthPx: number) => {
      if (!activeSheet) return;
      dispatchOrToast("xlsx:set-column-width", {
        sheet: activeSheet.name,
        column: col + 1,
        width: widthPx,
      });
    },
    [activeSheet, dispatchOrToast]
  );

  const onResizeRow = useCallback(
    (row: number, heightPx: number) => {
      if (!activeSheet) return;
      dispatchOrToast("xlsx:set-row-height", {
        sheet: activeSheet.name,
        row: row + 1,
        height: heightPx,
      });
    },
    [activeSheet, dispatchOrToast]
  );

  const onDeleteColumn = useCallback(() => {
    if (!activeSheet || !selection) return;
    const r = selectionToRange(selection);
    dispatchOrToast("xlsx:delete-column", {
      sheet: activeSheet.name,
      at: r.start.col + 1,
      count: r.end.col - r.start.col + 1,
    });
  }, [activeSheet, selection, dispatchOrToast]);

  const onClearContents = useCallback(() => {
    if (!activeSheet || !selection) return;
    const a = agentRef.current;
    if (!a) return;
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
  }, [activeSheet, selection, pushToast]);

  const onClearFormats = useCallback(() => {
    if (!activeSheet || !selection) return;
    onApplyFormat({
      font: { color: undefined, bold: undefined, italic: undefined, underline: undefined, strike: undefined },
      fill: { color: undefined, pattern: undefined },
      alignment: { horizontal: undefined, vertical: undefined },
      numberFormat: undefined,
    });
  }, [activeSheet, selection, onApplyFormat]);

  const onContextMenuOpen = useCallback((target: GridContextTarget, coords: { x: number; y: number }) => {
    setCtxMenu({ target, x: coords.x, y: coords.y });
  }, []);

  const closeCtxMenu = useCallback(() => setCtxMenu(null), []);

  const onCutMenu = useCallback(() => {
    void copySelection("cut");
  }, [copySelection]);
  const onCopyMenu = useCallback(() => {
    void copySelection("copy");
  }, [copySelection]);
  const onPasteMenu = useCallback(() => {
    void pasteAtSelection();
  }, [pasteAtSelection]);

  const ctxMenuItems = useMemo<ReadonlyArray<ContextMenuItem>>(() => {
    if (!ctxMenu) return [];
    const target = ctxMenu.target;
    const canCopyHere = !!(activeSheet && selection);
    const cellEntries: ContextMenuItem[] = [
      {
        kind: "action",
        id: "cut",
        label: "Cut",
        shortcut: "⌘X",
        disabled: !canCopyHere,
        onSelect: onCutMenu,
      },
      {
        kind: "action",
        id: "copy",
        label: "Copy",
        shortcut: "⌘C",
        disabled: !canCopyHere,
        onSelect: onCopyMenu,
      },
      {
        kind: "action",
        id: "paste",
        label: "Paste",
        shortcut: "⌘V",
        disabled: !canCopyHere,
        onSelect: onPasteMenu,
      },
      { kind: "divider", id: "div-clipboard" },
      {
        kind: "action",
        id: "insert-row-above",
        label: "Insert row above",
        onSelect: onInsertRowAbove,
      },
      {
        kind: "action",
        id: "insert-row-below",
        label: "Insert row below",
        onSelect: onInsertRowBelow,
      },
      {
        kind: "action",
        id: "insert-col-left",
        label: "Insert column left",
        onSelect: onInsertColumnLeft,
      },
      {
        kind: "action",
        id: "insert-col-right",
        label: "Insert column right",
        onSelect: onInsertColumnRight,
      },
      { kind: "divider", id: "div-insert" },
      { kind: "action", id: "delete-row", label: "Delete row", onSelect: onDeleteRow },
      { kind: "action", id: "delete-col", label: "Delete column", onSelect: onDeleteColumn },
      { kind: "divider", id: "div-delete" },
      { kind: "action", id: "clear-contents", label: "Clear contents", onSelect: onClearContents },
      { kind: "action", id: "clear-formats", label: "Clear formats", onSelect: onClearFormats },
      { kind: "divider", id: "div-data" },
      {
        kind: "action",
        id: "text-to-columns",
        label: "Text to Columns…",
        disabled: !canTextToColumns,
        onSelect: onTextToColumns,
      },
    ];
    if (target.kind === "row-header") {
      return [
        {
          kind: "action",
          id: "cut",
          label: "Cut",
          shortcut: "⌘X",
          disabled: !canCopyHere,
          onSelect: onCutMenu,
        },
        {
          kind: "action",
          id: "copy",
          label: "Copy",
          shortcut: "⌘C",
          disabled: !canCopyHere,
          onSelect: onCopyMenu,
        },
        {
          kind: "action",
          id: "paste",
          label: "Paste",
          shortcut: "⌘V",
          disabled: !canCopyHere,
          onSelect: onPasteMenu,
        },
        { kind: "divider", id: "div-clip-row" },
        { kind: "action", id: "insert-row-above", label: "Insert row above", onSelect: onInsertRowAbove },
        { kind: "action", id: "insert-row-below", label: "Insert row below", onSelect: onInsertRowBelow },
        { kind: "action", id: "delete-row", label: "Delete row", onSelect: onDeleteRow },
        { kind: "divider", id: "div-row-clear" },
        { kind: "action", id: "clear-contents", label: "Clear contents", onSelect: onClearContents },
      ];
    }
    if (target.kind === "col-header") {
      return [
        {
          kind: "action",
          id: "cut",
          label: "Cut",
          shortcut: "⌘X",
          disabled: !canCopyHere,
          onSelect: onCutMenu,
        },
        {
          kind: "action",
          id: "copy",
          label: "Copy",
          shortcut: "⌘C",
          disabled: !canCopyHere,
          onSelect: onCopyMenu,
        },
        {
          kind: "action",
          id: "paste",
          label: "Paste",
          shortcut: "⌘V",
          disabled: !canCopyHere,
          onSelect: onPasteMenu,
        },
        { kind: "divider", id: "div-clip-col" },
        { kind: "action", id: "insert-col-left", label: "Insert column left", onSelect: onInsertColumnLeft },
        {
          kind: "action",
          id: "insert-col-right",
          label: "Insert column right",
          onSelect: onInsertColumnRight,
        },
        { kind: "action", id: "delete-col", label: "Delete column", onSelect: onDeleteColumn },
        { kind: "divider", id: "div-col-clear" },
        { kind: "action", id: "clear-contents", label: "Clear contents", onSelect: onClearContents },
        { kind: "divider", id: "div-col-data" },
        {
          kind: "action",
          id: "text-to-columns",
          label: "Text to Columns…",
          disabled: !canTextToColumns,
          onSelect: onTextToColumns,
        },
      ];
    }
    return cellEntries;
  }, [
    ctxMenu,
    activeSheet,
    selection,
    onCutMenu,
    onCopyMenu,
    onPasteMenu,
    onInsertRowAbove,
    onInsertRowBelow,
    onInsertColumnLeft,
    onInsertColumnRight,
    onDeleteRow,
    onDeleteColumn,
    onClearContents,
    onClearFormats,
    canTextToColumns,
    onTextToColumns,
  ]);

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

  // Shared text-formatting bar wiring. The provider is built once via
  // useState's lazy initialiser; it closes over the refs above so it
  // always sees the latest selection/sheet/styles when the user
  // clicks a control. (Same pattern as DocxEditor — the React
  // Compiler can't see through the closure boundary, so we silence
  // the rule for this construction site.)
  /* eslint-disable react-hooks/refs */
  const [textFormatProvider] = useState<TextFormatProvider>(() =>
    createXlsxFormatProvider({
      agentRef,
      selectionRef,
      sheetRef: activeSheetRef,
      stylesRef,
      pushToast,
    })
  );
  /* eslint-enable react-hooks/refs */
  const textFormatActive: ActiveTextFormat = computeXlsxActive(
    activeSheet,
    snapshot?.root.styles ?? null,
    selection
  );

  return (
    <div
      ref={surfaceRef}
      tabIndex={0}
      onKeyDown={onSurfaceKeyDown}
      onCopy={onSurfaceCopy}
      onCut={onSurfaceCut}
      onPaste={onSurfacePaste}
      data-testid="xlsx-surface"
      data-whole-row={wholeRowSelection ? "1" : "0"}
      data-whole-col={wholeColSelection ? "1" : "0"}
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
          textFormatProvider={textFormatProvider}
          textFormatActive={textFormatActive}
          canMerge={canMerge}
          canUnmerge={canUnmerge}
          onMerge={onMerge}
          onUnmerge={onUnmerge}
          canUndo={agent?.canUndo() ?? false}
          canRedo={agent?.canRedo() ?? false}
          onUndo={() => {
            const a = agentRef.current;
            if (a && a.canUndo()) a.undo();
          }}
          onRedo={() => {
            const a = agentRef.current;
            if (a && a.canRedo()) a.redo();
          }}
          canTextToColumns={canTextToColumns}
          onTextToColumns={onTextToColumns}
          onOpenShortcuts={() => shortcutsDialog.setOpen(true)}
        />
      ) : null}

      <TextToColumnsPopover
        open={ttocOpen}
        defaultDelimiter={ttocDefaultDelim}
        onCancel={() => setTtocOpen(false)}
        onConfirm={onTextToColumnsConfirm}
      />

      <div className="formula-bar relative flex items-center gap-2 rounded-md border border-divider bg-surface px-2 py-1.5">
        <span
          data-testid="cell-ref"
          className="inline-flex h-7 min-w-[60px] items-center justify-center rounded border border-divider bg-background px-2 text-xs font-mono text-foreground"
        >
          {selectedRef || "—"}
        </span>
        <span className="text-secondary text-xs font-mono">fx</span>
        <div className="relative flex-1 font-mono text-xs">
          <FormulaHighlight
            value={formulaValue}
            tokens={formulaTokens}
            refColors={refColors}
            scrollLeft={formulaScrollLeft}
          />
          <input
            ref={formulaInputRef}
            data-testid="formula-input"
            aria-label="Formula bar"
            value={formulaValue}
            onScroll={(e) => setFormulaScrollLeft(e.currentTarget.scrollLeft)}
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
                onFormulaSubmit({ row: e.shiftKey ? -1 : 1, col: 0 });
              } else if (e.key === "Tab") {
                e.preventDefault();
                onFormulaSubmit({ row: 0, col: e.shiftKey ? -1 : 1 });
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
            // When the formula starts with `=`, the FormulaHighlight
            // overlay is responsible for the visible glyphs — make the
            // input's own text transparent (but keep the caret visible
            // via `caretColor`). Plain literals stay rendered by the
            // input itself so we don't have to model number / string
            // colours in the overlay too.
            style={{
              position: "relative",
              zIndex: 1,
              background: "transparent",
              color: formulaValue.startsWith("=") ? "transparent" : undefined,
              caretColor: "var(--foreground)",
            }}
            className="block w-full bg-transparent p-1 text-xs text-foreground placeholder:text-tertiary focus:outline-none"
          />
        </div>
        <div className="absolute left-[68px] right-2 top-full z-40">
          <FormulaSuggest
            matches={suggestionMatches}
            highlight={Math.min(suggestHighlight, Math.max(suggestionMatches.length - 1, 0))}
            onPick={acceptSuggestion}
            onHighlight={setSuggestHighlight}
          />
        </div>
      </div>

      <div className="relative flex flex-1 min-h-0 gap-2">
        <div className="relative flex-1 min-h-0">
        {activeSheet && snapshot ? (
          <Grid
            sheet={activeSheet}
            styles={snapshot.root.styles}
            selection={selection}
            onSelect={handleGridSelect}
            onCommitEdit={onCommitGridEdit}
            onResizeColumn={onResizeColumn}
            onResizeRow={onResizeRow}
            refRects={refRects}
            onSelectAxis={handleAxisSelect}
            onContextMenu={onContextMenuOpen}
            marchingAnts={
              marchingAnts && marchingAnts.sheet === activeSheet.name
                ? {
                    r1: marchingAnts.r1,
                    c1: marchingAnts.c1,
                    r2: marchingAnts.r2,
                    c2: marchingAnts.c2,
                    mode: marchingAnts.mode,
                  }
                : null
            }
            onFill={onFill}
            liveEditDraft={
              formulaFocused && selection && isSingle(selection)
                ? {
                    row: selection.anchor.row,
                    col: selection.anchor.col,
                    draft: formulaDraft,
                  }
                : null
            }
          />
        ) : (
          <div className="flex h-full items-center justify-center rounded-md border border-divider bg-background text-sm text-secondary">
            <Loader2 className="mr-2 animate-spin" size={14} />
            Loading workbook…
          </div>
        )}
        </div>
        {agent && activeSheet ? (
          <aside
            data-testid="xlsx-comments-sidebar"
            className="hidden w-[260px] shrink-0 flex-col gap-2 overflow-y-auto rounded-md border border-divider bg-surface p-2 lg:flex"
          >
            <CommentsSidebar
              key={`xlsx-comments-${activeSheet.name}-${revision}`}
              provider={createXlsxCommentsProvider({ agent, sheetName: activeSheet.name })}
              author="You"
              emptyHint="No comments on this sheet yet."
            />
            {selection ? (
              <CommentComposer
                provider={createXlsxCommentsProvider({ agent, sheetName: activeSheet.name })}
                anchor={{
                  kind: "xlsx-cell",
                  sheet: activeSheet.name,
                  ref: formatA1(selection.anchor),
                }}
                placeholder={`Comment on ${formatA1(selection.anchor)}…`}
              />
            ) : null}
          </aside>
        ) : null}
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

      <ContextMenu
        open={ctxMenu !== null}
        x={ctxMenu?.x ?? 0}
        y={ctxMenu?.y ?? 0}
        items={ctxMenuItems}
        onClose={closeCtxMenu}
      />

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
      <KeyboardShortcutsDialog
        product="xlsx"
        open={shortcutsDialog.open}
        onClose={() => shortcutsDialog.setOpen(false)}
      />
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
