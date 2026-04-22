"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { ReactNode } from "react";
import { CommentComposer, CommentsSidebar, cn } from "@officeai/ui";
import { I18nProvider, useTranslator, type Locale } from "@/lib/i18n";
import { createXlsxCommentsProvider } from "./xlsxCommentsProvider";
import {
  EditorShell,
  EmptyState,
  buildPaletteFromCatalogue,
  createToastId,
  type ExportFormat,
  type ExportOptionValues,
  type FindAdapter,
  type FindMatch,
  type FindOptions,
  type PaletteCommand,
  type PaletteRunners,
  type ProductAdapter,
  type SaveState,
  type ToastItem,
} from "@/lib/shell";
import { xlsxActions } from "@officeai/xlsx";
import {
  PRODUCT_FILE_TYPES,
  downloadBlob,
  openFile as openFileViaService,
  saveFile as saveFileViaService,
} from "@/lib/files/file-service";
import { convertViaServer } from "@/lib/files/convert-client";
import { sheetToCsv, sheetToTsv, workbookToCsvZip, workbookToJson } from "./lib/export-data";
import {
  XlsxAgent,
  assignRefColors,
  cellKey,
  colToLetter,
  evaluateConditionalFormats,
  flattenCellXf,
  formatA1,
  formatRange,
  parseA1,
  parseRange,
  tokenizeForDisplay,
  type Cell,
  type CellFormatPatch,
  type CellValue,
  type DisplayToken,
  type SetCellFormatPayload,
  type Sheet,
  type StyleTable,
  type XlsxClipboardSnapshot,
  type XlsxSnapshot,
} from "@officeai/xlsx";
import type { ActiveTextFormat, TextFormatProvider } from "@officeai/text-formatting";
import { computeXlsxActive, createXlsxFormatProvider } from "./xlsxFormatProvider";
import { buildBlankXlsx, buildSampleXlsx } from "@/lib/sample-xlsx";
import { handleUndoRedo } from "@/lib/undo-redo";
import {
  Grid,
  type CommentMarker,
  type GridContextTarget,
  type MarchingAntsRect,
  type RefRect,
} from "./Grid";
import { ContextMenu, type ContextMenuItem } from "./ContextMenu";
import { computeUsedRange } from "./gridDimensions";
import { FormulaHighlight } from "./FormulaHighlight";
import {
  allAreas,
  areasContainCell,
  forEachUnionSparseCell,
  formatAreas,
  formatSelection,
  isSingle,
  normalizeSelection,
  selectionToRange,
  singleSelection,
  unionSpanUpperBound,
  type CellPos,
  type Selection,
} from "./selection";
import { FormulaSuggest, applySuggestion, getSuggestions } from "./FormulaSuggest";
import { Toolbar, type BorderPreset } from "./Toolbar";
import { SheetTabBar } from "./SheetTabBar";
import { FormatCellsDialog, type TabId as FormatTabId } from "./FormatCellsDialog";
import { PasteSpecialDialog, type PasteSpecialOptions } from "./PasteSpecialDialog";
import { ConditionalFormatDialog } from "./ConditionalFormatDialog";
import { DataValidationDialog } from "./DataValidationDialog";
import { NameBox } from "./NameBox";
import { NameManagerDialog } from "./NameManagerDialog";
import { ChartDialog, InsertChartDialog } from "./InsertChartDialog";
import {
  PageSetupDialog,
  type PageSetupSubmit,
  type PageSetupTab,
} from "./PageSetupDialog";
import { ZoomDialog } from "./ZoomDialog";
import { ProtectSheetDialog, type ProtectSheetValues } from "./ProtectSheetDialog";
import { ProtectWorkbookDialog, type ProtectWorkbookValues } from "./ProtectWorkbookDialog";
import { InsertPivotTableDialog, type PivotDialogSubmit } from "./InsertPivotTableDialog";
import type { ChartKind, ConditionalFormat, DataValidation, DefinedName } from "@officeai/xlsx";
import { useShortcutsDialog } from "@/lib/shortcuts/useShortcutsDialog";
import { KeyboardShortcutsDialog } from "@/lib/shortcuts/KeyboardShortcutsDialog";
import {
  PresenceSlot,
  readExplicitRoomFromUrl,
  RemotePresenceList,
  roomIdForSource,
  useCommandBroadcast,
  usePublishPresence,
  useRealtimeRoom,
  useStableTabId,
} from "@/lib/realtime";
import { TextToColumnsPopover } from "./TextToColumnsPopover";
import { FilterDropdown } from "./FilterDropdown";
import { sniffDelimiter } from "@officeai/xlsx";
import { formatCellValue as renderCellValue } from "./styles";
import {
  marshalClipboard,
  parseClipboardPayload,
  writeToSystemClipboard,
  readFromSystemClipboard,
} from "./clipboard";
import { EMBED_MIME } from "@/lib/embed/envelope";

const SAMPLE_NAME = "sample.xlsx";
const BLANK_NAME = "Untitled.xlsx";

const XLSX_EXPORT_FORMATS: ReadonlyArray<ExportFormat> = [
  {
    id: "xlsx",
    label: "Excel workbook (.xlsx)",
    description: "Round-trip native OOXML with all sheets, formatting and formulas.",
    extension: "xlsx",
    mime: PRODUCT_FILE_TYPES.xlsx.primaryMime,
    kind: "instant",
    group: "native",
    icon: "sheet",
  },
  {
    id: "pdf",
    label: "PDF document (.pdf)",
    description: "Server-side conversion via LibreOffice. Print-ready output of the full workbook.",
    extension: "pdf",
    mime: "application/pdf",
    kind: "dialog",
    group: "pdf-web",
    icon: "pdf",
    optionFields: [
      {
        id: "orientation",
        label: "Orientation",
        control: {
          type: "select",
          defaultId: "source",
          options: [
            { id: "source", label: "Use sheet setting" },
            { id: "portrait", label: "Portrait" },
            { id: "landscape", label: "Landscape" },
          ],
        },
        hint: "LibreOffice respects each sheet's own page setup unless overridden.",
      },
      {
        id: "fitToWidth",
        label: "Fit to page width",
        control: { type: "toggle", defaultValue: false },
        hint: "Scales wide tables down so columns don't overflow.",
      },
    ],
  },
  {
    id: "html",
    label: "Web page (.html)",
    description: "Server-side HTML export. Renders with sheet tabs and styles.",
    extension: "html",
    mime: "text/html",
    kind: "instant",
    group: "pdf-web",
    icon: "code",
  },
  // Data exports are sorted active-sheet first (the common ask),
  // then the all-sheets variants. The "<Scope> — <FORMAT> (.ext)"
  // shape mirrors PPTX's "Current slide — PDF (.pdf)" /
  // "All slides — PNG (.zip)" so the dropdown reads consistently
  // across products. We use "Active sheet" rather than "Current
  // sheet" because that's Excel's own vernacular in Save As
  // (likewise "Current slide" matches PowerPoint).
  {
    id: "csv",
    label: "Active sheet — CSV (.csv)",
    description: "Comma-separated values for the active sheet only.",
    extension: "csv",
    mime: "text/csv;charset=utf-8",
    kind: "instant",
    group: "data",
    icon: "text",
  },
  {
    id: "tsv",
    label: "Active sheet — TSV (.tsv)",
    description: "Tab-separated values for the active sheet.",
    extension: "tsv",
    mime: "text/tab-separated-values;charset=utf-8",
    kind: "instant",
    group: "data",
    icon: "text",
  },
  {
    id: "csv-all",
    label: "All sheets — CSV (.zip)",
    description: "One CSV per worksheet, bundled as a zip.",
    extension: "zip",
    mime: "application/zip",
    kind: "instant",
    group: "data",
    icon: "text",
  },
  {
    id: "json",
    label: "All sheets — JSON (.json)",
    description: "Structured cell values per sheet, keyed by column letter.",
    extension: "json",
    mime: "application/json;charset=utf-8",
    kind: "instant",
    group: "data",
    icon: "code",
  },
];

function stripXlsxExtension(name: string): string {
  return name.replace(/\.xlsx$/i, "");
}

/**
 * True when an event target is a form control / editable surface that
 * owns its own key handling (inputs, textareas, selects, buttons,
 * contenteditable). Used by the surface-level keydown / clipboard
 * guards so typing in the comments composer doesn't get hijacked by
 * the grid's "type-to-edit" handler.
 */
/**
 * Decode an image blob just enough to learn its intrinsic pixel
 * size. We need this when inserting because xlsx:add-image insists
 * on positive width/height in CSS pixels — the XLSX serializer then
 * converts to EMU. We use an `Image` element rather than parsing
 * PNG/JPEG headers so we transparently support whatever the browser
 * supports (incl. JPEG variants we don't want to hand-roll).
 */
const SUPPORTED_IMAGE_MIME = new Set(["image/png", "image/jpeg", "image/gif"]);

function isSupportedImageFile(file: { type?: string; name?: string }): boolean {
  if (file.type && SUPPORTED_IMAGE_MIME.has(file.type)) return true;
  // Some browsers leave `type` empty for OS-drag payloads; fall back to
  // the extension so a dragged-in `.png` still goes through the
  // image insertion path instead of being treated as an xlsx file.
  const name = (file.name ?? "").toLowerCase();
  return name.endsWith(".png") || name.endsWith(".jpg") || name.endsWith(".jpeg") || name.endsWith(".gif");
}

function defaultInsertAnchor(selection: Selection | null): { fromRow: number; fromCol: number } {
  if (!selection) return { fromRow: 0, fromCol: 0 };
  return { fromRow: selection.anchor.row, fromCol: selection.anchor.col };
}

async function measureImage(
  buf: ArrayBuffer,
  contentType: string
): Promise<{ width: number; height: number }> {
  const blob = new Blob([buf], { type: contentType });
  const url = URL.createObjectURL(blob);
  try {
    const dims = await new Promise<{ width: number; height: number }>((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve({ width: img.naturalWidth || 96, height: img.naturalHeight || 96 });
      img.onerror = () => reject(new Error("Failed to decode image"));
      img.src = url;
    });
    return dims;
  } finally {
    URL.revokeObjectURL(url);
  }
}

function isFormControlTarget(t: EventTarget | null): boolean {
  if (!t || !(t instanceof HTMLElement)) return false;
  const tag = t.tagName;
  if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT" || tag === "BUTTON") return true;
  return t.isContentEditable;
}

/**
 * Lift the truthy state out of a worksheet's opaque `<printOptions>`
 * element so the Page Layout toggle buttons can mirror reality.
 */
interface PageState {
  readonly printGridLines: boolean;
  readonly printHeadings: boolean;
}

function readPageState(sheet: Sheet | null): PageState {
  if (!sheet?.printOptionsXml) return { printGridLines: false, printHeadings: false };
  const truthy = (xml: string, attr: string): boolean => {
    const m = xml.match(new RegExp(`\\b${attr}\\s*=\\s*"(1|true)"`));
    return !!m;
  };
  return {
    printGridLines: truthy(sheet.printOptionsXml, "gridLines"),
    printHeadings: truthy(sheet.printOptionsXml, "headings"),
  };
}

interface CalcState {
  readonly mode: "auto" | "autoNoTable" | "manual";
  readonly calcOnSave: boolean;
}

function readCalcState(snapshot: XlsxSnapshot | null): CalcState {
  const xml = snapshot?.root.calcPrXml;
  if (!xml) return { mode: "auto", calcOnSave: true };
  const m = xml.match(/\bcalcMode\s*=\s*"([^"]*)"/);
  const modeStr = m?.[1];
  const mode: CalcState["mode"] =
    modeStr === "manual" || modeStr === "autoNoTable" ? modeStr : "auto";
  const calcOnSaveAttr = xml.match(/\bcalcOnSave\s*=\s*"([^"]*)"/);
  const calcOnSave = calcOnSaveAttr ? calcOnSaveAttr[1] !== "0" && calcOnSaveAttr[1] !== "false" : true;
  return { mode, calcOnSave };
}

interface SheetViewState {
  readonly view: "normal" | "pageBreakPreview" | "pageLayout";
  readonly showGridLines: boolean;
  readonly showRowColHeaders: boolean;
  readonly showRuler: boolean;
  readonly rightToLeft: boolean;
  readonly zoomScale: number;
  readonly showFormulas: boolean;
}

function readSheetViewState(sheet: Sheet | null): SheetViewState {
  const xml = sheet?.sheetViewsXml ?? "";
  const m = xml.match(/<sheetView\b[^>]*>/i);
  const tag = m?.[0] ?? "";
  const truthy = (attr: string, dflt: boolean): boolean => {
    const a = tag.match(new RegExp(`\\b${attr}\\s*=\\s*"([^"]*)"`));
    if (!a) return dflt;
    const v = a[1];
    return v !== "0" && v !== "false";
  };
  const num = (attr: string, dflt: number): number => {
    const a = tag.match(new RegExp(`\\b${attr}\\s*=\\s*"([^"]*)"`));
    const n = a ? Number(a[1]) : NaN;
    return Number.isFinite(n) && n > 0 ? n : dflt;
  };
  const viewAttr = tag.match(/\bview\s*=\s*"([^"]*)"/);
  const view: SheetViewState["view"] =
    viewAttr?.[1] === "pageBreakPreview" || viewAttr?.[1] === "pageLayout"
      ? viewAttr[1]
      : "normal";
  return {
    view,
    showGridLines: truthy("showGridLines", true),
    showRowColHeaders: truthy("showRowColHeaders", true),
    showRuler: truthy("showRuler", true),
    rightToLeft: truthy("rightToLeft", false),
    zoomScale: num("zoomScale", 100),
    showFormulas: truthy("showFormulas", false),
  };
}

/**
 * Strip a sheet name + absolute markers from the OOXML `refersTo`
 * text of a defined name, leaving the bare A1 range list. Used by
 * the Print Area "Add to print area" path so we can union with the
 * existing area text.
 */
function stripSheetPrefixUtil(refersTo: string, sheetName: string): string {
  const escaped = sheetName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const re = new RegExp(`(?:^|,)\\s*(?:'?${escaped}'?!)([^,]+)`, "g");
  const parts: string[] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(refersTo)) !== null) {
    if (m[1]) parts.push(m[1].replace(/\$/g, ""));
  }
  return parts.length ? parts.join(",") : refersTo;
}

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
export interface XlsxEditorProps {
  /** Fired whenever the editor's bootstrap-ready state changes. The
   * page-level splash listens to this to know when to fade out and
   * unveil the workbook. Stays `false` until the agent is mounted
   * AND the first snapshot + active sheet are resolved (or until the
   * initial sample load has failed, in which case we treat the
   * EmptyState recovery affordance as "ready" so the splash unveils
   * it). */
  readonly onBootstrapReady?: (ready: boolean) => void;
  /** Optional pre-loaded workbook. When provided, the editor fetches
   * the bytes at `url` instead of building the synthetic sample, and
   * uses `name` as the workbook title (so subsequent Save / Export
   * keep the original filename). Used by the home page's "sample
   * files" listing. */
  readonly initialSource?: { readonly url: string; readonly name: string };
  /** When true, the editor bootstraps with a truly empty workbook
   * (one blank sheet, no rows) instead of the synthetic sample. Used
   * by the home page's "New spreadsheet" action. Ignored when
   * `initialSource` is set. */
  readonly initialBlank?: boolean;
  /** Optional pre-loaded workbook bytes. When set, takes priority
   * over `initialSource` and `initialBlank` so embedding hosts can
   * stream a `Uint8Array` straight into the editor without first
   * stashing it under a URL. */
  readonly initialBytes?: Uint8Array;
  /** Filename to display + use on Save when `initialBytes` is set. */
  readonly initialFilename?: string;
  /** Host save handler. When provided, Save invokes this with the
   * exported bytes, OOXML MIME, and working filename — instead of
   * falling back to File-System-Access. */
  readonly onSave?: (bytes: Uint8Array, mime: string, filename: string) => Promise<void>;
  /** Host close handler — surfaces a "Back" affordance in editor
   * chrome. Embedding route owns the actual navigation. */
  readonly onClose?: () => void;
  /** Override the i18n locale; mounts a self-contained
   * `<I18nProvider initialLocale={locale}>` so the editor renders in
   * the requested language regardless of host provider state. */
  readonly locale?: Locale;
  /** Theme override placeholder; wired in Phase 1. */
  readonly theme?: "light" | "dark";
  /** Realtime presence identity (host-supplied). When set, replaces
   * the default anonymous identity on the awareness payload so cell
   * cursors / avatars show the authenticated user's real name. */
  readonly presenceUser?: { readonly id: string; readonly name: string; readonly color?: string };
  /** Explicit realtime room id (host-supplied). Pin two browsers
   * viewing the same workbook into the same room without
   * coordinating URLs. Pass `null` to disable realtime. */
  readonly room?: string | null;
  /** Hide the 📁 Open toolbar affordance. Set by embedded hosts that
   * own their document corpus — see
   * `EmbeddedEditorProps.hideLocalFileOpen` in
   * `@officeai/react-editors/contract`. */
  readonly hideLocalFileOpen?: boolean;
}

export function XlsxEditor(props: XlsxEditorProps = {}): ReactNode {
  const { locale } = props;
  if (locale !== undefined) {
    return (
      <I18nProvider initialLocale={locale}>
        <XlsxEditorInner {...props} />
      </I18nProvider>
    );
  }
  return <XlsxEditorInner {...props} />;
}

function XlsxEditorInner({
  onBootstrapReady,
  initialSource,
  initialBlank,
  initialBytes,
  initialFilename,
  onSave: onSaveProp,
  onClose: onCloseProp,
  presenceUser,
  room: roomOverride,
  hideLocalFileOpen,
}: XlsxEditorProps = {}): ReactNode {
  const { t } = useTranslator();
  const agentRef = useRef<XlsxAgent | null>(null);
  const [agent, setAgent] = useState<XlsxAgent | null>(null);
  const [snapshot, setSnapshot] = useState<XlsxSnapshot | null>(null);
  const [activeSheetName, setActiveSheetName] = useState<string | null>(null);
  const [selection, setSelection] = useState<Selection | null>(singleSelection({ row: 0, col: 0 }));
  // C13 — Disjoint extra areas accumulated by Ctrl/Cmd-click. The
  // active area still lives in `selection`; everything here is just
  // additional marquee + cells that participate in clear-contents,
  // formatting commands, and the status-bar aggregates.
  const [extraAreas, setExtraAreas] = useState<ReadonlyArray<Selection>>([]);
  const [selectedImageId, setSelectedImageId] = useState<string | null>(null);
  const [selectedChartId, setSelectedChartId] = useState<string | null>(null);
  const [insertChartOpen, setInsertChartOpen] = useState(false);
  const [insertPivotOpen, setInsertPivotOpen] = useState(false);
  // Page Layout / Review / View dialogs.
  const [pageSetupOpen, setPageSetupOpen] = useState<{ tab?: PageSetupTab } | null>(null);
  const [protectSheetOpen, setProtectSheetOpen] = useState(false);
  const [protectWorkbookOpen, setProtectWorkbookOpen] = useState(false);
  const [zoomDialogOpen, setZoomDialogOpen] = useState(false);
  /**
   * `editChartId !== null` opens the Edit-chart dialog pre-filled
   * from `activeSheet.charts.find(...)`. We keep an id (not the
   * chart object) here so the dialog always reflects the latest
   * snapshot if the chart was mutated by another command in the
   * background while the dialog is open.
   */
  const [editChartId, setEditChartId] = useState<string | null>(null);
  const [pendingCount, setPendingCount] = useState(0);
  const [toasts, setToasts] = useState<ReadonlyArray<ToastItem>>([]);
  const [formulaDraft, setFormulaDraft] = useState("");
  const [formulaFocused, setFormulaFocused] = useState(false);
  const formulaInputRef = useRef<HTMLInputElement | null>(null);
  // Sheet that was active when the user entered the formula bar. While
  // editing, clicking another sheet tab navigates the grid view but
  // keeps the formula's destination cell pinned to this sheet, and any
  // cell picks on a different sheet are inserted as `Sheet!A1` refs
  // (Excel "point mode" parity). Cleared on commit / cancel.
  const [formulaOriginSheet, setFormulaOriginSheet] = useState<string | null>(null);
  const formulaOriginSheetRef = useRef<string | null>(null);
  // Dedicated hidden <input type=file> for the toolbar's "Insert image"
  // affordance — kept separate from the workbook open input so the
  // accept= filter doesn't bleed across the two flows.
  const imageInputRef = useRef<HTMLInputElement | null>(null);
  const [filename, setFilename] = useState<string>(
    initialFilename ?? initialSource?.name ?? (initialBlank || initialBytes ? BLANK_NAME : SAMPLE_NAME)
  );
  const [saveState, setSaveState] = useState<SaveState>("saved");
  const fileHandleRef = useRef<FileSystemFileHandle | undefined>(undefined);
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
  // Tracks the initial sample-load lifecycle. The page-level splash
  // (see `apps/web/app/xlsx-editor/page.tsx`) stays up until the
  // agent + first snapshot are ready, so the editor surface looks
  // alive on the very first paint instead of asking the user to
  // "Open a workbook" for the ~150-300 ms it takes JSZip + the
  // parser to materialise the synthetic sample. Flips to true only
  // on failure, at which point we report ready up so the splash
  // unveils EmptyState as the recovery affordance.
  const [initialLoadFailed, setInitialLoadFailed] = useState(false);
  const shortcutsDialog = useShortcutsDialog();

  const pushToast = useCallback((kind: ToastItem["kind"], text: string) => {
    const id = createToastId("xlsx");
    setToasts((prev) => [...prev, { id, kind, text }]);
  }, []);
  const dismissToast = useCallback((id: string) => {
    setToasts((prev) => prev.filter((t) => t.id !== id));
  }, []);

  // Refs that mirror the latest selection / sheet / styles so the
  // shared TextFormatBar provider can read them at event-handler
  // time without us having to rebuild the provider on every render.
  // Updated below after the relevant `useMemo`s have run.
  const selectionRef = useRef<Selection | null>(selection);
  const activeSheetRef = useRef<Sheet | null>(null);
  const stylesRef = useRef<StyleTable | null>(null);
  const snapshotRef = useRef<XlsxSnapshot | null>(null);

  // Holds the unsubscribe handle for the active agent's `subscribe()`
  // callback so we can swap agents (Open file) without leaking listeners.
  const offRef = useRef<(() => void) | null>(null);

  const mountAgent = useCallback(
    (a: XlsxAgent, name: string, handle?: FileSystemFileHandle) => {
      offRef.current?.();
      agentRef.current = a;
      setAgent(a);
      setFilename(name);
      fileHandleRef.current = handle;
      setSaveState("saved");
      const snap = a.getSnapshot();
      setSnapshot(snap);
      setActiveSheetName(snap.root.sheets[0]?.name ?? null);
      setSelection(singleSelection({ row: 0, col: 0 }));
      setExtraAreas([]);
      setPendingCount(a.getPendingMutations().length);
      let first = true;
      offRef.current = a.subscribe((s, mutation) => {
        setSnapshot(s);
        setPendingCount(a.getPendingMutations().length);
        // Skip the synchronous initial snapshot most agents emit.
        if (first) {
          first = false;
          return;
        }
        // Surface bus rebase rejections (see
        // packages/core/src/commands/bus.ts.recomputeWorking).
        // Without this toast a pending agent mutation that
        // becomes inconsistent after an undo just vanishes —
        // the user has no way to know what happened.
        if (mutation.status === "rejected" && mutation.rejection?.code === "rebase-failed") {
          pushToast(
            "warn",
            `An agent suggestion couldn't be re-applied after the last edit (${mutation.rejection.message})`
          );
          return;
        }
        setSaveState("modified");
      });
    },
    [setAgent, setSnapshot, setActiveSheetName, setSelection, setPendingCount]
  );

  useEffect(() => {
    let cancelled = false;
    void (async () => {
      try {
        // Four bootstrap paths, picked in priority order:
        //   1. `initialBytes` — host streams the workbook straight
        //      in (used when the editor is embedded by hof-os and
        //      the bytes already came back from S3).
        //   2. `initialSource` — fetch a pre-existing .xlsx.
        //   3. `initialBlank` — build an empty workbook.
        //   4. Default — build the synthetic sample.
        let buf: ArrayBuffer;
        if (initialBytes) {
          const copy = new Uint8Array(initialBytes.byteLength);
          copy.set(initialBytes);
          buf = copy.buffer;
        } else if (initialSource) {
          const res = await fetch(initialSource.url);
          if (!res.ok) {
            throw new Error(`Failed to load ${initialSource.name} (${res.status})`);
          }
          buf = await res.arrayBuffer();
        } else if (initialBlank) {
          buf = await buildBlankXlsx();
        } else {
          buf = await buildSampleXlsx();
        }
        if (cancelled) return;
        const a = await XlsxAgent.fromBuffer(buf);
        if (cancelled) return;
        const name =
          initialFilename ?? initialSource?.name ?? (initialBlank || initialBytes ? BLANK_NAME : SAMPLE_NAME);
        mountAgent(a, name);
      } catch (err) {
        if (cancelled) return;
        // Sample build failed — fall back to the EmptyState so the
        // user has a visible "Open file" affordance instead of being
        // stuck on the skeleton forever.
        setInitialLoadFailed(true);
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
  }, [pushToast, mountAgent, initialSource, initialBlank, initialBytes, initialFilename]);

  const openFile = useCallback(
    async (file: File, handle?: FileSystemFileHandle) => {
      const lower = file.name.toLowerCase();
      if (!lower.endsWith(".xlsx")) {
        pushToast("error", `Unsupported file: ${file.name} (only .xlsx)`);
        return;
      }
      try {
        const buf = await file.arrayBuffer();
        const a = await XlsxAgent.fromBuffer(buf);
        mountAgent(a, file.name, handle);
        pushToast("success", `Opened ${file.name}`);
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [mountAgent, pushToast]
  );

  const onPickFile = useCallback(async () => {
    try {
      const opened = await openFileViaService({
        description: PRODUCT_FILE_TYPES.xlsx.description,
        mimeToExt: PRODUCT_FILE_TYPES.xlsx.mimeToExt,
        accept: PRODUCT_FILE_TYPES.xlsx.accept,
      });
      if (!opened) return;
      const file = new File([opened.bytes as BlobPart], opened.name, {
        type: PRODUCT_FILE_TYPES.xlsx.primaryMime,
      });
      await openFile(file, opened.handle);
    } catch (err) {
      pushToast("error", err instanceof Error ? err.message : String(err));
    }
  }, [openFile, pushToast]);

  const dispatchAddImage = useCallback(
    async (
      file: Blob,
      anchor: { fromRow: number; fromCol: number; fromOffsetXPx?: number; fromOffsetYPx?: number }
    ) => {
      const sheet = activeSheetRef.current;
      if (!sheet) return;
      const a = agentRef.current;
      if (!a) return;
      const contentType = file.type;
      if (contentType !== "image/png" && contentType !== "image/jpeg" && contentType !== "image/gif") {
        pushToast("error", `Unsupported image type: ${contentType || "unknown"}`);
        return;
      }
      const buf = await file.arrayBuffer();
      const bytes = new Uint8Array(buf);
      const { width, height } = await measureImage(buf, contentType);
      try {
        await a.applyCommand({
          type: "xlsx:add-image",
          payload: {
            sheet: sheet.name,
            bytes,
            contentType,
            fromRow: anchor.fromRow,
            fromCol: anchor.fromCol,
            fromOffsetXPx: anchor.fromOffsetXPx ?? 0,
            fromOffsetYPx: anchor.fromOffsetYPx ?? 0,
            widthPx: width,
            heightPx: height,
          },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const onShellFileDrop = useCallback(
    (file: File) => {
      if (isSupportedImageFile(file)) {
        void dispatchAddImage(file, defaultInsertAnchor(selectionRef.current));
        return;
      }
      void openFile(file);
    },
    [dispatchAddImage, openFile]
  );

  const onInsertImageClick = useCallback(() => {
    imageInputRef.current?.click();
  }, []);

  const onImageInputChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (file) void dispatchAddImage(file, defaultInsertAnchor(selectionRef.current));
      e.target.value = "";
    },
    [dispatchAddImage]
  );

  const activeSheet: Sheet | null = useMemo(() => {
    if (!snapshot || !activeSheetName) return null;
    return snapshot.root.sheets.find((s) => s.name === activeSheetName) ?? null;
  }, [snapshot, activeSheetName]);

  // Pivot tables anchored on the active sheet. Phase 1 of the pivot
  // work treats them as read-only — `Grid` paints a tinted rectangle
  // + dashed outline + name badge over each pivot's `location.ref`
  // so the user can see pivot output as something other than a
  // static range, even though the underlying cells already carry the
  // cached values Excel wrote. We filter by `sheetId` (set at parse
  // time from the rels graph) rather than name so renaming a sheet
  // wouldn't break the linkage.
  const pivotsForActiveSheet = useMemo(() => {
    if (!snapshot || !activeSheet) return undefined;
    const all = snapshot.root.pivotTables;
    if (!all || all.length === 0) return undefined;
    const filtered = all.filter((p) => p.sheetId === activeSheet.sheetId && p.location !== undefined);
    return filtered.length > 0 ? filtered : undefined;
  }, [snapshot, activeSheet]);

  // Mirror bootstrap-ready up to the page-level splash so it can
  // fade out and unveil either the grid or the EmptyState. Owning
  // the splash at page scope (not here) keeps the badge `<span>`
  // mounted across the dynamic-import handoff — see
  // `apps/web/app/xlsx-editor/page.tsx`. We treat
  // `initialLoadFailed` as "ready" because once the failure has
  // surfaced, the EmptyState is the recovery UI the user should see
  // — the splash's job is done.
  useEffect(() => {
    const ready = initialLoadFailed ? true : Boolean(agent && activeSheet && snapshot);
    onBootstrapReady?.(ready);
  }, [agent, activeSheet, snapshot, initialLoadFailed, onBootstrapReady]);

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
    snapshotRef.current = snapshot ?? null;
  }, [snapshot]);
  useEffect(() => {
    formulaOriginSheetRef.current = formulaOriginSheet;
  }, [formulaOriginSheet]);

  const selectedCell = useMemo(() => {
    if (!activeSheet || !selection) return null;
    // Formula bar / derived display always reflect the *anchor* cell
    // (Excel matches this — the active cell stays white in a range).
    return activeSheet.cells.get(cellKey(selection.anchor.row, selection.anchor.col)) ?? null;
  }, [activeSheet, selection]);

  const selectedRef = selection
    ? extraAreas.length > 0
      ? formatAreas(selection, extraAreas)
      : formatSelection(selection)
    : "";

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

  // Cells that carry an unresolved (top-level) comment. Walks the
  // active sheet's comments — replies and resolved threads are
  // skipped so the yellow marker reads "open conversation here".
  const commentMarkers: ReadonlyArray<CommentMarker> = useMemo(() => {
    if (!activeSheet) return [];
    const out: CommentMarker[] = [];
    const seen = new Set<string>();
    for (const c of activeSheet.comments) {
      if (c.parentId) continue;
      if (c.resolved) continue;
      if (seen.has(c.ref)) continue;
      seen.add(c.ref);
      try {
        const addr = parseA1(c.ref);
        out.push({ row: addr.row, col: addr.col });
      } catch {
        // Malformed ref — skip silently rather than break the grid.
      }
    }
    return out;
  }, [activeSheet]);

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

  // Wrap a sheet name in single quotes and double inner apostrophes when
  // it isn't a bare A1-safe identifier (Excel quoting parity, e.g.
  // `'My Sheet'!A1`, `'Bob''s'!B2`).
  const formatSheetPrefixedRef = useCallback((sheetName: string, body: string): string => {
    const needsQuoting = /[^A-Za-z0-9_.]/.test(sheetName) || /^\d/.test(sheetName);
    const quoted = needsQuoting ? `'${sheetName.replace(/'/g, "''")}'` : sheetName;
    return `${quoted}!${body}`;
  }, []);

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

  // C1 — Real Excel bounds, kept in lockstep with `Grid.tsx`. The
  // virtualised renderer never instantiates a per-cell array of this
  // size; selection rectangles are stored as `(row, col)` pairs, not
  // expanded into per-cell sets, so a "select entire column"
  // operation costs O(1) regardless of `GRID_ROWS`.
  const GRID_ROWS = 1_048_576;
  const GRID_COLS = 16_384;

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
    (pos: CellPos, opts?: { extend?: boolean; additive?: boolean }) => {
      // Click-to-insert-ref: while the formula bar is in point mode,
      // a plain click inserts the cell ref at the caret; a drag /
      // Shift-click extends the previously inserted ref into a range.
      // The active selection (and therefore the formula's *target*
      // cell) is intentionally NOT moved here — Excel keeps the
      // original cell as the destination while the user picks refs.
      if (formulaEditing) {
        // When the active sheet differs from where the formula edit
        // started, qualify the picked ref with `Sheet!` so the
        // formula points at the cell on the *other* sheet (Excel's
        // sheet-tab-during-edit behaviour).
        const origin = formulaOriginSheetRef.current;
        const currentSheet = activeSheetRef.current?.name ?? null;
        const crossSheet = origin !== null && currentSheet !== null && currentSheet !== origin;
        const qualify = (body: string): string =>
          crossSheet && currentSheet ? formatSheetPrefixedRef(currentSheet, body) : body;
        if (opts?.extend && pendingRefSpanRef.current) {
          // Build "anchor:focus" from the *first* clicked ref (stored
          // implicitly in pendingRefAnchorRef) and the new pos.
          const anchor = pendingRefAnchorRef.current ?? pos;
          const sel: Selection = {
            anchor,
            focus: pos,
          };
          const body = isSingle(sel) ? formatA1(sel.anchor) : formatRange(selectionToRange(sel));
          insertRefAtCaret(qualify(body));
        } else {
          pendingRefSpanRef.current = null;
          pendingRefAnchorRef.current = pos;
          insertRefAtCaret(qualify(formatA1(pos)));
        }
        return;
      }

      // C13 — Ctrl/Cmd-click without Shift starts a new disjoint
      // area. We push the previous active selection into `extraAreas`
      // (Excel does not dedup or test for overlap; matching the
      // behaviour keeps the model simple and round-trippable to a
      // sequential clear / format dispatch).
      if (opts?.additive) {
        setSelection((prev) => {
          if (prev) {
            setExtraAreas((es) => [...es, prev]);
          }
          return singleSelection(pos);
        });
        setSelectedImageId(null);
        setSelectedChartId(null);
        surfaceRef.current?.focus({ preventScroll: true });
        return;
      }

      // Normal (non-formula) selection behaviour.
      if (opts?.extend) {
        setSelection((prev) => (prev ? { anchor: prev.anchor, focus: pos } : singleSelection(pos)));
      } else {
        setSelection(singleSelection(pos));
        // Plain (non-additive) click clears any extra areas — Excel
        // parity. Shift-extend keeps them around.
        setExtraAreas([]);
        // C8 — Arm Format Painter drop on this fresh mousedown. The
        // eventual global mouseup picks up `selectionRef` (which by
        // then includes any drag-extend) and applies the captured
        // formats over it.
        if (formatPainterRef.current) {
          formatPainterPendingRef.current = true;
        }
      }
      setSelectedImageId(null);
      setSelectedChartId(null);
      // Pull keyboard focus back to the surface so the next printable
      // key starts type-to-edit on the new anchor. Focus synchronously
      // so the active element is already the surface by the time the
      // user's mouseup completes.
      surfaceRef.current?.focus({ preventScroll: true });
    },
    [formulaEditing, formatSheetPrefixedRef, insertRefAtCaret]
  );

  // Grid-level keyboard handler: when nothing else has focus, a
  // printable key starts in-formula-bar editing for the active
  // single-cell anchor. F2 enters with the existing value; Backspace /
  // Delete clears the cell.
  const surfaceRef = useRef<HTMLDivElement | null>(null);
  // Auto-focus the surface as soon as a workbook is mounted so the
  // anchor cell (A1 by default) is immediately "live" — typing a
  // printable key kicks off type-to-edit on the active cell without
  // the user having to click the grid first. Re-runs on every
  // workbook swap (open file, new sheet) so the same affordance
  // applies to subsequent loads. Skipped when the user is already
  // editing the formula bar so we don't yank focus mid-keystroke.
  useEffect(() => {
    if (!agent) return;
    if (formulaFocused) return;
    surfaceRef.current?.focus({ preventScroll: true });
  }, [agent, formulaFocused]);
  // Bumped each time the comments sidebar requests "scroll to this
  // cell". The Grid effect keys off `nonce` so clicking the same
  // comment twice still re-scrolls and re-flashes.
  const [commentScrollTarget, setCommentScrollTarget] = useState<{
    row: number;
    col: number;
    nonce: number;
  } | null>(null);
  const focusCommentComposer = useCallback(() => {
    requestAnimationFrame(() => {
      const tab = document.querySelector<HTMLButtonElement>('[data-testid="rail-tab-comments"]');
      tab?.click();
      requestAnimationFrame(() => {
        const textarea = document.querySelector<HTMLTextAreaElement>(
          'aside [data-testid="comment-composer"] textarea'
        );
        textarea?.focus();
      });
    });
  }, []);
  // Locate the cell carrying `commentId` on the active sheet, move
  // the selection there (so the marquee + name-box update too) and
  // ask the Grid to scroll + flash the cell. Wired into the comments
  // provider's `onScrollTo` hook.
  const scrollToComment = useCallback(
    (commentId: string) => {
      const a = agentRef.current;
      if (!a) return;
      const snap = a.getSnapshot();
      const sheet = snap.root.sheets.find((s) => s.name === activeSheetName);
      if (!sheet) return;
      const target = sheet.comments.find((c) => c.id === commentId);
      if (!target) return;
      let addr: { row: number; col: number };
      try {
        addr = parseA1(target.ref);
      } catch {
        return;
      }
      setSelection(singleSelection(addr));
      setCommentScrollTarget((prev) => ({
        row: addr.row,
        col: addr.col,
        nonce: (prev?.nonce ?? 0) + 1,
      }));
    },
    [activeSheetName]
  );
  // Move the active selection one cell in the given direction, with
  // optional Shift-extend. Pure helper — no side effects beyond
  // calling setSelection.
  const moveSelection = useCallback((dRow: number, dCol: number, opts: { extend: boolean }) => {
    // C13 — Keyboard navigation collapses any Ctrl-click extras back
    // to a single active area, mirroring Excel's behaviour where
    // arrow / Tab / Enter abandons the disjoint selection.
    setExtraAreas([]);
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
  // Falls back to the *used range* edge (not the worksheet edge) so
  // a stray Cmd+Down on an empty column doesn't try to walk
  // 1,048,576 rows linearly. C1 introduced the real Excel bounds and
  // this guard keeps the operation O(used-range height/width).
  const jumpToDataEdge = useCallback(
    (dRow: number, dCol: number, opts: { extend: boolean }) => {
      if (!activeSheet) return;
      setSelection((prev) => {
        const base: CellPos = prev?.focus ?? { row: 0, col: 0 };
        const used = computeUsedRange(activeSheet.cells);
        const isFilled = (r: number, c: number): boolean => {
          if (r < 0 || c < 0 || r >= GRID_ROWS || c >= GRID_COLS) return false;
          const cell = activeSheet.cells.get(cellKey(r, c));
          return !!cell && cell.value !== null && cell.value !== undefined;
        };
        // Far edge for the search: the used-range bound in the
        // direction we're moving, falling back to the caret position
        // (we won't move further than there is data).
        const maxR = used ? used.r2 : base.row;
        const maxC = used ? used.c2 : base.col;
        const startFilled = isFilled(base.row, base.col);
        let r = base.row;
        let c = base.col;
        const inSearchBounds = (nr: number, nc: number): boolean => {
          if (nr < 0 || nc < 0 || nr >= GRID_ROWS || nc >= GRID_COLS) return false;
          if (dRow > 0 && nr > maxR) return false;
          if (dRow < 0 && nr < 0) return false;
          if (dCol > 0 && nc > maxC) return false;
          if (dCol < 0 && nc < 0) return false;
          return true;
        };
        const step = (): boolean => {
          const nr = r + dRow;
          const nc = c + dCol;
          if (!inSearchBounds(nr, nc)) return false;
          r = nr;
          c = nc;
          return true;
        };
        if (startFilled) {
          while (isFilled(r + dRow, c + dCol)) {
            if (!step()) break;
          }
        } else {
          while (step()) {
            if (isFilled(r, c)) break;
          }
          // No data ahead → snap to the worksheet edge (Excel
          // parity). Cheap because it's a single index assignment.
          if (!isFilled(r, c)) {
            if (dRow > 0) r = GRID_ROWS - 1;
            else if (dRow < 0) r = 0;
            if (dCol > 0) c = GRID_COLS - 1;
            else if (dCol < 0) c = 0;
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
    async (
      direct?: { html?: string | null; text?: string | null },
      opts?: { mode?: "all" | "values" | "formulas" | "formats"; transpose?: boolean }
    ): Promise<boolean> => {
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
          payload: {
            sheet: activeSheet.name,
            target,
            source: snap,
            mode: opts?.mode ?? "all",
            transpose: opts?.transpose ?? false,
          },
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
      // Cmd+V / arrow keys feel "Excel-y". When transposing, the
      // pasted block is rotated — height becomes width and vice versa.
      const transposed = !!opts?.transpose;
      const pastedH = transposed ? snap.width : snap.height;
      const pastedW = transposed ? snap.height : snap.width;
      const end: CellPos = {
        row: selection.anchor.row + Math.max(0, pastedH - 1),
        col: selection.anchor.col + Math.max(0, pastedW - 1),
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
      if (isFormControlTarget(e.target)) return;
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
        if (payload.embed) e.clipboardData.setData(EMBED_MIME, payload.embed);
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
      if (isFormControlTarget(e.target)) return;
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
        if (payload.embed) e.clipboardData.setData(EMBED_MIME, payload.embed);
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
      if (isFormControlTarget(e.target)) return;
      if (!agentRef.current || !activeSheet || !selection) return;
      // Image clipboard branch — when the user pastes a screenshot or
      // copied image we route to xlsx:add-image instead of the cell
      // paste flow. We pick the FIRST image item (the rest are usually
      // alternative encodings of the same image: png + jpg fallbacks).
      const items = Array.from(e.clipboardData.items ?? []);
      const imgItem = items.find((it) => it.kind === "file" && SUPPORTED_IMAGE_MIME.has(it.type));
      if (imgItem) {
        const file = imgItem.getAsFile();
        if (file) {
          e.preventDefault();
          void dispatchAddImage(file, defaultInsertAnchor(selection));
          return;
        }
      }
      e.preventDefault();
      const html = e.clipboardData.getData("text/html");
      const text = e.clipboardData.getData("text/plain");
      void pasteAtSelection({ html, text });
    },
    [activeSheet, selection, pasteAtSelection, dispatchAddImage]
  );

  const onSurfaceKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLDivElement>) => {
      // Don't steal keys destined for inputs / textareas / selects /
      // buttons / contenteditable surfaces — they have their own
      // onKeyDown. Without this, typing in the comments composer (a
      // <textarea> inside the floating rail) would be hijacked by the
      // grid's "type-to-edit" handler and start editing the active
      // cell's formula instead.
      if (isFormControlTarget(e.target)) return;

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
        // Ctrl+End → bottom-right corner of the used range. C1
        // unified the bounds calc in `computeUsedRange` so the
        // navigation, viewport-fit and (later) print-area logic all
        // agree on what "used" means.
        e.preventDefault();
        if (!activeSheet) return;
        const used = computeUsedRange(activeSheet.cells);
        const focus: CellPos = used ? { row: used.r2, col: used.c2 } : { row: 0, col: 0 };
        setSelection((prev) =>
          e.shiftKey && prev ? { anchor: prev.anchor, focus } : singleSelection(focus)
        );
        return;
      }

      // Mod+A — Excel-style "Select All". First press expands the
      // selection to cover the used range; a follow-up press (or
      // pressing it on an empty sheet) goes all the way to the
      // worksheet's full extent. Matches Excel's two-stage chord.
      if ((e.metaKey || e.ctrlKey) && !e.shiftKey && !e.altKey && (e.key === "a" || e.key === "A")) {
        e.preventDefault();
        if (!activeSheet) return;
        const used = computeUsedRange(activeSheet.cells);
        const range = selection ? selectionToRange(selection) : null;
        const matchesUsed =
          !!used &&
          !!range &&
          range.start.row === used.r1 &&
          range.start.col === used.c1 &&
          range.end.row === used.r2 &&
          range.end.col === used.c2;
        setExtraAreas([]);
        if (!used || matchesUsed) {
          setSelection({
            anchor: { row: 0, col: 0 },
            focus: { row: GRID_ROWS - 1, col: GRID_COLS - 1 },
          });
        } else {
          setSelection({
            anchor: { row: used.r1, col: used.c1 },
            focus: { row: used.r2, col: used.c2 },
          });
        }
        return;
      }

      // Mod+Shift+Space — promote the selection to the entire sheet.
      if ((e.metaKey || e.ctrlKey) && e.shiftKey && !e.altKey && e.code === "Space") {
        e.preventDefault();
        setExtraAreas([]);
        setSelection({
          anchor: { row: 0, col: 0 },
          focus: { row: GRID_ROWS - 1, col: GRID_COLS - 1 },
        });
        return;
      }

      // Mod+Space — promote the selection to the full column(s) it
      // touches. macOS swallows Cmd+Space (Spotlight); Excel for Mac
      // documents Ctrl+Space for column-select, which still reaches
      // us through `e.ctrlKey`.
      if ((e.metaKey || e.ctrlKey) && !e.shiftKey && !e.altKey && e.code === "Space") {
        if (!selection) return;
        e.preventDefault();
        const r = selectionToRange(selection);
        setExtraAreas([]);
        setSelection({
          anchor: { row: 0, col: r.start.col },
          focus: { row: GRID_ROWS - 1, col: r.end.col },
        });
        return;
      }

      // Shift+Space — promote the selection to the full row(s) it touches.
      if (!e.metaKey && !e.ctrlKey && !e.altKey && e.shiftKey && e.code === "Space") {
        if (!selection) return;
        e.preventDefault();
        const r = selectionToRange(selection);
        setExtraAreas([]);
        setSelection({
          anchor: { row: r.start.row, col: 0 },
          focus: { row: r.end.row, col: GRID_COLS - 1 },
        });
        return;
      }

      // Mod+PageUp / Mod+PageDown — switch to the previous / next
      // visible sheet (Excel parity). Wraps around either end.
      if (
        (e.metaKey || e.ctrlKey) &&
        !e.shiftKey &&
        !e.altKey &&
        (e.key === "PageUp" || e.key === "PageDown")
      ) {
        const allSheets = snapshotRef.current?.root.sheets ?? [];
        // Skip hidden / very-hidden sheets so the chord matches what
        // the tab strip actually shows the user.
        const visible = allSheets.filter((s) => s.state !== "hidden" && s.state !== "veryHidden");
        if (visible.length === 0) return;
        e.preventDefault();
        const idx = Math.max(
          0,
          visible.findIndex((s) => s.name === activeSheetName)
        );
        const dir = e.key === "PageDown" ? 1 : -1;
        const next = visible[(idx + dir + visible.length) % visible.length];
        if (next) setActiveSheetName(next.name);
        return;
      }

      // PageUp / PageDown — Excel's "page" of vertical movement. We
      // don't track the visible row count at this layer, so 24 rows
      // (≈ a default Excel viewport) is a reasonable constant. Shift
      // extends the active range, matching arrow-key behaviour.
      if (!e.metaKey && !e.ctrlKey && !e.altKey && (e.key === "PageUp" || e.key === "PageDown")) {
        e.preventDefault();
        const dir = e.key === "PageDown" ? 1 : -1;
        moveSelection(dir * 24, 0, { extend: e.shiftKey });
        return;
      }

      // Mod+Enter — commit the active anchor cell's value/formula
      // across every other cell in the selection. This is the
      // "fill all selected" chord Excel uses when you've typed into
      // a multi-cell selection. Skipped on a single-cell anchor
      // (nothing else to fan out to).
      if ((e.metaKey || e.ctrlKey) && !e.shiftKey && !e.altKey && e.key === "Enter") {
        if (!activeSheet || !selection || isSingle(selection)) return;
        e.preventDefault();
        const a = agentRef.current;
        if (!a) return;
        const r = selectionToRange(selection);
        const anchorCell = activeSheet.cells.get(cellKey(selection.anchor.row, selection.anchor.col));
        const anchorFormula = anchorCell?.formula?.text ?? null;
        const anchorValue: CellValue = anchorCell?.value ?? null;
        for (let row = r.start.row; row <= r.end.row; row++) {
          for (let col = r.start.col; col <= r.end.col; col++) {
            if (row === selection.anchor.row && col === selection.anchor.col) continue;
            const ref = formatA1({ row, col });
            const cmd = anchorFormula
              ? {
                  type: "xlsx:set-cell-formula" as const,
                  payload: { sheet: activeSheet.name, ref, formula: anchorFormula },
                  source: "human" as const,
                }
              : {
                  type: "xlsx:set-cell-value" as const,
                  payload: { sheet: activeSheet.name, ref, value: anchorValue },
                  source: "human" as const,
                };
            void a
              .applyCommand(cmd)
              .catch((err: unknown) => pushToast("error", err instanceof Error ? err.message : String(err)));
          }
        }
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

      if (e.key === "F3") {
        e.preventDefault();
        setNameManagerOpen(true);
        return;
      }

      // C14 — Excel-parity "Format as Table" shortcut (Mod+T).
      // Promotes the current selection (or single anchor's
      // contiguous block, in a future pass) to an Excel Table.
      if ((e.metaKey || e.ctrlKey) && !e.shiftKey && !e.altKey && (e.key === "t" || e.key === "T")) {
        if (!activeSheet || !selection) return;
        e.preventDefault();
        const range = formatSelection(selection);
        dispatchOrToast("xlsx:add-table", {
          sheet: activeSheet.name,
          range,
        });
        return;
      }

      // Mod+D — Excel "Fill Down". Replicate the top row of the
      // selection over every row beneath it via the shared
      // `xlsx:fill-range` handler (so series detection / formula
      // re-anchoring stay consistent with the drag-handle path).
      if ((e.metaKey || e.ctrlKey) && !e.shiftKey && !e.altKey && (e.key === "d" || e.key === "D")) {
        if (!activeSheet || !selection) return;
        const r = selectionToRange(selection);
        if (r.start.row === r.end.row) return;
        e.preventDefault();
        const sourceRange = formatRange({
          start: { row: r.start.row, col: r.start.col },
          end: { row: r.start.row, col: r.end.col },
        });
        const targetRange = formatRange({ start: r.start, end: r.end });
        dispatchOrToast("xlsx:fill-range", {
          sheet: activeSheet.name,
          source: sourceRange,
          target: targetRange,
          direction: "down",
        });
        return;
      }

      // Mod+R — Excel "Fill Right". Same deal, but the leftmost
      // column propagates across the rest of the selection.
      if ((e.metaKey || e.ctrlKey) && !e.shiftKey && !e.altKey && (e.key === "r" || e.key === "R")) {
        if (!activeSheet || !selection) return;
        const r = selectionToRange(selection);
        if (r.start.col === r.end.col) return;
        e.preventDefault();
        const sourceRange = formatRange({
          start: { row: r.start.row, col: r.start.col },
          end: { row: r.end.row, col: r.start.col },
        });
        const targetRange = formatRange({ start: r.start, end: r.end });
        dispatchOrToast("xlsx:fill-range", {
          sheet: activeSheet.name,
          source: sourceRange,
          target: targetRange,
          direction: "right",
        });
        return;
      }

      // Mod+Shift+L — toggle the AutoFilter band on / off (Excel
      // parity). Routes through the same callback the toolbar's
      // Filter button uses, so the active-sheet detection + range
      // inference stays in one place.
      if ((e.metaKey || e.ctrlKey) && e.shiftKey && !e.altKey && (e.key === "l" || e.key === "L")) {
        if (!activeSheet) return;
        e.preventDefault();
        onToggleFilter();
        return;
      }

      if (e.key === "F2") {
        if (!selection || !isSingle(selection)) return;
        e.preventDefault();
        const fi = formulaInputRef.current;
        if (!fi) return;
        setFormulaDraft(derivedFormulaDisplay);
        setFormulaFocused(true);
        if (activeSheet) setFormulaOriginSheet(activeSheet.name);
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
        // Floating-object selection (image / chart) takes priority
        // over collapsing the cell range — Excel parity, and
        // matches the user's mental model that Escape "drops" the
        // currently-armed thing.
        if (selectedChartId !== null || selectedImageId !== null) {
          e.preventDefault();
          setSelectedChartId(null);
          setSelectedImageId(null);
          return;
        }
        if (!selection) return;
        e.preventDefault();
        setSelection(singleSelection(selection.anchor));
        return;
      }

      // Undo / Redo — Excel-parity (Cmd/Ctrl+Z, Cmd/Ctrl+Shift+Z,
      // Cmd/Ctrl+Y). All three editors share `handleUndoRedo` so
      // the chord detection, the form-field guard, and the
      // dispatch-onto-the-bus contract are identical. The bus is
      // the single source of truth for history (see
      // spec/shared/agent-api.md).
      if (handleUndoRedo(e, agentRef.current)) return;

      // C7 — Excel-parity Paste Special shortcut (Cmd+Shift+V).
      // Native paste (Cmd+V) is captured via `onSurfacePaste`;
      // Cmd+Shift+V deliberately bypasses the native event so we can
      // open a dialog and let the user choose what to paste before
      // we read the clipboard.
      if ((e.metaKey || e.ctrlKey) && e.shiftKey && !e.altKey && (e.key === "v" || e.key === "V")) {
        e.preventDefault();
        setPasteSpecialOpen(true);
        return;
      }

      // C8 — Format Painter shortcut (Mod+Shift+C). Captures the
      // current selection's formats and arms the painter; the next
      // mousedown-drag-up on the grid drops the formats over the
      // destination range. Esc cancels.
      if ((e.metaKey || e.ctrlKey) && e.shiftKey && !e.altKey && (e.key === "c" || e.key === "C")) {
        e.preventDefault();
        activateFormatPainterRef.current(false);
        return;
      }

      // C5 — Excel-parity Format Cells shortcut (Mod+1). Opens the
      // dialog on the Number tab. We intentionally guard on the
      // *physical* "1" key (`Digit1`) and require no Shift / Alt
      // modifiers so Mod+Shift+1 still maps to the number-format
      // shortcut below without conflict.
      if ((e.metaKey || e.ctrlKey) && !e.shiftKey && !e.altKey && (e.code === "Digit1" || e.key === "1")) {
        e.preventDefault();
        setFormatCellsTab("number");
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
          // C13 — Use the central `onApplyFormat` so the toggle fans
          // out across every area in the multi-area selection.
          onApplyFormat({ font: { [markKey]: !currentlyOn } } as never);
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
          // C13 — Same fan-out treatment as the Bold/Italic shortcut
          // so number-format chords land on every disjoint area.
          onApplyFormat({ numberFormat });
          return;
        }
      }

      // Mod+; / Mod+Shift+; — insert today's date / current time at
      // the active anchor. We write a real Excel date serial AND
      // apply the matching built-in number format so the cell reads
      // "10/19/2026" / "11:30:00 AM" instead of a literal string.
      // Use `e.code === "Semicolon"` so Shift+; (which yields ":"
      // on most layouts) still maps to the physical key.
      if ((e.metaKey || e.ctrlKey) && !e.altKey && e.code === "Semicolon") {
        if (!activeSheet || !selection) return;
        e.preventDefault();
        const a = agentRef.current;
        if (!a) return;
        const insertTime = e.shiftKey;
        const now = new Date();
        const serial = insertTime ? excelTimeSerial(now) : excelDateSerial(now);
        const ref = formatA1({ row: selection.anchor.row, col: selection.anchor.col });
        void a
          .applyCommand({
            type: "xlsx:set-cell-value",
            payload: { sheet: activeSheet.name, ref, value: serial },
            source: "human",
          })
          .catch((err: unknown) => pushToast("error", err instanceof Error ? err.message : String(err)));
        // Built-in numFmtIds: 14 = m/d/yyyy, 19 = h:mm:ss AM/PM.
        // Pushed through `onApplyFormat` so multi-area selections
        // (Ctrl-clicked extras) also pick up the format.
        onApplyFormat({ numberFormat: insertTime ? "19" : "14" });
        return;
      }

      if (e.key === "Backspace" || e.key === "Delete") {
        if (!activeSheet) return;
        if (selectedImageId) {
          e.preventDefault();
          const a0 = agentRef.current;
          if (!a0) return;
          const removedId = selectedImageId;
          setSelectedImageId(null);
          void a0
            .applyCommand({
              type: "xlsx:remove-image",
              payload: { sheet: activeSheet.name, imageId: removedId },
              source: "human",
            })
            .catch((err: unknown) => pushToast("error", err instanceof Error ? err.message : String(err)));
          return;
        }
        if (!selection) return;
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

        // Range-aware clear: only the *populated* cells in the
        // union need a clear command — empty cells are already
        // empty. Walking the sparse `cells` map keeps "Delete on
        // entire column" from issuing a million no-op commands.
        forEachUnionSparseCell(activeSheet.cells.values(), allAreas(selection, extraAreas), (cell) => {
          void a
            .applyCommand({
              type: "xlsx:set-cell-value",
              payload: {
                sheet: activeSheet.name,
                ref: formatA1({ row: cell.row, col: cell.col }),
                value: null,
              },
              source: "human",
            })
            .catch((err: unknown) => {
              pushToast("error", err instanceof Error ? err.message : String(err));
            });
        });
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
      if (activeSheet) setFormulaOriginSheet(activeSheet.name);
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
      activeSheetName,
      derivedFormulaDisplay,
      jumpToDataEdge,
      marchingAnts,
      moveSelection,
      pushToast,
      selectedCell,
      selectedChartId,
      selectedImageId,
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
      // (the moving end of the range doesn't receive the value). When
      // the user navigated to another sheet during the edit (point
      // mode), commit to — and return to — the origin sheet.
      const anchor = selection.anchor;
      const ref = formatA1({ row: anchor.row, col: anchor.col });
      const targetSheet = formulaOriginSheetRef.current ?? activeSheet.name;
      void dispatchCellEdit(targetSheet, ref, formulaDraft);
      setFormulaFocused(false);
      setFormulaDraft("");
      setFormulaOriginSheet(null);
      formulaInputRef.current?.blur();
      if (targetSheet !== activeSheet.name) setActiveSheetName(targetSheet);
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

  const onMoveImage = useCallback(
    (
      imageId: string,
      anchor: { fromRow: number; fromCol: number; fromOffsetXPx: number; fromOffsetYPx: number }
    ) => {
      if (!activeSheet) return;
      const a = agentRef.current;
      if (!a) return;
      void a
        .applyCommand({
          type: "xlsx:move-image",
          payload: {
            sheet: activeSheet.name,
            imageId,
            fromRow: anchor.fromRow,
            fromCol: anchor.fromCol,
            fromOffsetXPx: anchor.fromOffsetXPx,
            fromOffsetYPx: anchor.fromOffsetYPx,
          },
          source: "human",
        })
        .catch((err: unknown) => pushToast("error", err instanceof Error ? err.message : String(err)));
    },
    [activeSheet, pushToast]
  );

  const onResizeImage = useCallback(
    (imageId: string, size: { widthPx: number; heightPx: number }) => {
      if (!activeSheet) return;
      const a = agentRef.current;
      if (!a) return;
      void a
        .applyCommand({
          type: "xlsx:resize-image",
          payload: {
            sheet: activeSheet.name,
            imageId,
            widthPx: size.widthPx,
            heightPx: size.heightPx,
          },
          source: "human",
        })
        .catch((err: unknown) => pushToast("error", err instanceof Error ? err.message : String(err)));
    },
    [activeSheet, pushToast]
  );

  const onRemoveImage = useCallback(
    (imageId: string) => {
      if (!activeSheet) return;
      const a = agentRef.current;
      if (!a) return;
      setSelectedImageId(null);
      void a
        .applyCommand({
          type: "xlsx:remove-image",
          payload: { sheet: activeSheet.name, imageId },
          source: "human",
        })
        .catch((err: unknown) => pushToast("error", err instanceof Error ? err.message : String(err)));
    },
    [activeSheet, pushToast]
  );

  /**
   * Chart move/resize mirror their image counterparts: the overlay
   * commits final geometry on mouse-up, we forward to the typed
   * command bus. Single dispatch per gesture keeps undo single-step.
   */
  const onMoveChart = useCallback(
    (
      chartId: string,
      anchor: { fromRow: number; fromCol: number; fromOffsetXPx: number; fromOffsetYPx: number }
    ) => {
      if (!activeSheet) return;
      const a = agentRef.current;
      if (!a) return;
      void a
        .applyCommand({
          type: "xlsx:move-chart",
          payload: {
            sheet: activeSheet.name,
            chartId,
            fromRow: anchor.fromRow,
            fromCol: anchor.fromCol,
            fromOffsetXPx: anchor.fromOffsetXPx,
            fromOffsetYPx: anchor.fromOffsetYPx,
          },
          source: "human",
        })
        .catch((err: unknown) => pushToast("error", err instanceof Error ? err.message : String(err)));
    },
    [activeSheet, pushToast]
  );

  const onResizeChart = useCallback(
    (chartId: string, size: { widthPx: number; heightPx: number }) => {
      if (!activeSheet) return;
      const a = agentRef.current;
      if (!a) return;
      void a
        .applyCommand({
          type: "xlsx:resize-chart",
          payload: {
            sheet: activeSheet.name,
            chartId,
            widthPx: size.widthPx,
            heightPx: size.heightPx,
          },
          source: "human",
        })
        .catch((err: unknown) => pushToast("error", err instanceof Error ? err.message : String(err)));
    },
    [activeSheet, pushToast]
  );

  /**
   * Resolve the chart currently targeted by the Edit dialog from the
   * live snapshot. Returning `null` (rather than reading the dialog's
   * stale copy) means the dialog auto-closes if the chart got removed
   * out from under it via `xlsx:remove-chart` (e.g. by the agent).
   */
  const editingChart = useMemo(() => {
    if (editChartId === null || !activeSheet) return null;
    return activeSheet.charts.find((c) => c.id === editChartId) ?? null;
  }, [editChartId, activeSheet]);

  const onSave = useCallback(async () => {
    const a = agentRef.current;
    if (!a) return;
    setSaveState("saving");
    try {
      const buf = await a.exportFile();
      const bytes = new Uint8Array(buf);
      const mime = PRODUCT_FILE_TYPES.xlsx.primaryMime;
      if (onSaveProp) {
        await onSaveProp(bytes, mime, filename);
        setSaveState("saved");
        pushToast("success", `Saved ${filename}`);
        return;
      }
      const wroteInPlace = await saveFileViaService(bytes, filename, mime, fileHandleRef.current);
      setSaveState("saved");
      pushToast("success", wroteInPlace ? `Saved ${filename}` : `Downloaded ${filename}`);
    } catch (err) {
      setSaveState("error");
      pushToast("error", err instanceof Error ? err.message : String(err));
    }
  }, [filename, onSaveProp, pushToast]);

  const onExport = useCallback(
    async (format: ExportFormat, options?: ExportOptionValues) => {
      const a = agentRef.current;
      if (!a) return;
      const baseName = stripXlsxExtension(filename);
      const downloadName = `${baseName}.${format.extension}`;
      try {
        switch (format.id) {
          case "xlsx": {
            const buf = await a.exportFile();
            downloadBlob(new Blob([buf as BlobPart], { type: format.mime }), downloadName);
            break;
          }
          case "pdf":
          case "html": {
            const buf = await a.exportFile();
            const out = await convertViaServer({
              bytes: new Uint8Array(buf),
              sourceExt: "xlsx",
              targetExt: format.id,
              filename: baseName,
            });
            downloadBlob(out, downloadName);
            break;
          }
          case "csv": {
            const sheet = activeSheetRef.current ?? snapshot?.root.sheets[0] ?? null;
            if (!sheet) throw new Error("No sheet to export.");
            const csv = sheetToCsv(sheet);
            downloadBlob(new Blob([csv], { type: format.mime }), downloadName);
            break;
          }
          case "csv-all": {
            const snap = a.getSnapshot();
            const blob = await workbookToCsvZip(snap);
            downloadBlob(blob, downloadName);
            break;
          }
          case "tsv": {
            const sheet = activeSheetRef.current ?? snapshot?.root.sheets[0] ?? null;
            if (!sheet) throw new Error("No sheet to export.");
            const tsv = sheetToTsv(sheet);
            downloadBlob(new Blob([tsv], { type: format.mime }), downloadName);
            break;
          }
          case "json": {
            const snap = a.getSnapshot();
            const json = workbookToJson(snap);
            downloadBlob(new Blob([json], { type: format.mime }), downloadName);
            break;
          }
          default:
            throw new Error(`Unsupported export format: ${format.id}`);
        }
        pushToast("success", `Exported ${downloadName}`);
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
        throw err;
      }
      void options;
    },
    [filename, pushToast, snapshot]
  );

  const sheets = snapshot?.root.sheets ?? [];
  const revision = snapshot?.revision ?? 0;

  // Build object URLs for image media. Re-built only when the workbook
  // image map changes (identity), so editing other parts of the sheet
  // doesn't churn the URLs and force <img> reloads. We revoke on the
  // next build / unmount.
  const imageObjectUrls = useMemo<ReadonlyMap<string, string>>(() => {
    const m = new Map<string, string>();
    if (!snapshot) return m;
    for (const [path, blob] of snapshot.root.images.entries()) {
      const b = new Blob([blob.bytes as BlobPart], { type: blob.contentType });
      m.set(path, URL.createObjectURL(b));
    }
    return m;
  }, [snapshot?.root.images]);
  useEffect(() => {
    return () => {
      for (const url of imageObjectUrls.values()) URL.revokeObjectURL(url);
    };
  }, [imageObjectUrls]);

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
      // C13 — Fan out across every area in the multi-area selection
      // so Bold / Fill / Number-format / Borders all "just work" on
      // disjoint Ctrl-click rectangles.
      const areas = allAreas(selection, extraAreas);
      const cmds = areas.map((area) => ({
        type: "xlsx:set-cell-format" as const,
        payload: {
          sheet: activeSheet.name,
          range: formatSelection(area),
          format: patch,
        },
        source: "human" as const,
      }));
      void a.applyCommands(cmds).catch((err: unknown) => {
        pushToast("error", err instanceof Error ? err.message : String(err));
      });
    },
    [activeSheet, selection, extraAreas, pushToast]
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
        | "xlsx:fill-range"
        | "xlsx:set-auto-filter"
        | "xlsx:set-filter-column"
        | "xlsx:clear-filter-column"
        | "xlsx:sort-range"
        | "xlsx:freeze-panes"
        | "xlsx:unfreeze-panes"
        | "xlsx:add-sheet"
        | "xlsx:rename-sheet"
        | "xlsx:delete-sheet"
        | "xlsx:move-sheet"
        | "xlsx:set-sheet-state"
        | "xlsx:add-conditional-format"
        | "xlsx:remove-conditional-format"
        | "xlsx:clear-conditional-formats"
        | "xlsx:add-data-validation"
        | "xlsx:remove-data-validation"
        | "xlsx:clear-data-validations"
        | "xlsx:add-defined-name"
        | "xlsx:update-defined-name"
        | "xlsx:remove-defined-name"
        | "xlsx:add-table"
        | "xlsx:remove-table"
        | "xlsx:add-chart"
        | "xlsx:remove-chart"
        | "xlsx:move-chart"
        | "xlsx:resize-chart"
        | "xlsx:update-chart"
        | "xlsx:set-range-values",
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

  /**
   * Lower-friction dispatcher for the recently-landed Page Layout /
   * Formulas / Review / View commands. The strict union in
   * `dispatchOrToast` predates them; rather than enumerate every
   * variant we accept a raw type string here. Errors still flow into
   * the toast surface.
   */
  const dispatchAny = useCallback(
    (type: string, payload: Record<string, unknown>): Promise<void> => {
      const a = agentRef.current;
      if (!a) return Promise.resolve();
      return a
        .applyCommand({ type, payload, source: "human" } as Parameters<typeof a.applyCommand>[0])
        .then(() => undefined)
        .catch((err: unknown) => {
          pushToast("error", err instanceof Error ? err.message : String(err));
        });
    },
    [pushToast]
  );

  /**
   * Quick chart-type switch from the chart toolbar. Goes through
   * `xlsx:update-chart` so undo/redo + diff tracking pick it up the
   * same way as data-range edits made via the dialog.
   */
  const onChangeChartKind = useCallback(
    (chartId: string, kind: ChartKind) => {
      if (!activeSheet) return;
      dispatchOrToast("xlsx:update-chart", {
        sheet: activeSheet.name,
        chartId,
        kind,
      });
    },
    [activeSheet, dispatchOrToast]
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

  // ── Page Layout / Formulas / Review / View derived state ────────
  // All four toolbar tabs lift opaque XML attributes off the typed
  // workbook/sheet projection so the buttons can mirror truth (e.g.
  // a green "Show formulas" toggle when the active sheet has it on).
  const pageState = useMemo(() => readPageState(activeSheet), [activeSheet]);
  const calcState = useMemo(() => readCalcState(snapshot), [snapshot]);
  const sheetView = useMemo(() => readSheetViewState(activeSheet), [activeSheet]);
  const sheetProtected = !!activeSheet?.sheetProtectionXml;
  const workbookProtected = !!snapshot?.root.workbookProtectionXml;

  const onOpenPageSetup = useCallback((tab?: PageSetupTab) => {
    setPageSetupOpen({ tab });
  }, []);

  const onApplyMarginsPreset = useCallback(
    (preset: "normal" | "wide" | "narrow") => {
      if (!activeSheet) return;
      void dispatchAny("xlsx:set-page-margins", { sheet: activeSheet.name, preset });
    },
    [activeSheet, dispatchAny]
  );

  const onApplyOrientation = useCallback(
    (orientation: "portrait" | "landscape") => {
      if (!activeSheet) return;
      void dispatchAny("xlsx:set-page-setup", { sheet: activeSheet.name, orientation });
    },
    [activeSheet, dispatchAny]
  );

  const onApplyPaperSize = useCallback(
    (paperSize: number) => {
      if (!activeSheet) return;
      void dispatchAny("xlsx:set-page-setup", { sheet: activeSheet.name, paperSize });
    },
    [activeSheet, dispatchAny]
  );

  const onPrintArea = useCallback(
    (mode: "set" | "clear" | "add") => {
      if (!activeSheet) return;
      if (mode === "clear") {
        void dispatchAny("xlsx:set-print-area", { sheet: activeSheet.name, clear: true });
        return;
      }
      if (!selection) {
        pushToast("error", "Select a range first.");
        return;
      }
      const newRange = formatSelection(selection);
      if (mode === "set") {
        void dispatchAny("xlsx:set-print-area", { sheet: activeSheet.name, range: newRange });
        return;
      }
      // mode === "add": union with the existing print area, if any.
      const existing = snapshot?.root.definedNames.find(
        (n) => n.name === "_xlnm.Print_Area" && n.scope === activeSheet.name
      );
      const existingClean = existing
        ? stripSheetPrefixUtil(existing.refersTo, activeSheet.name)
        : "";
      const merged = existingClean ? `${existingClean},${newRange}` : newRange;
      void dispatchAny("xlsx:set-print-area", { sheet: activeSheet.name, range: merged });
    },
    [activeSheet, selection, snapshot, dispatchAny, pushToast]
  );

  const onTogglePrintFlag = useCallback(
    (flag: "gridLines" | "headings", value: boolean) => {
      if (!activeSheet) return;
      void dispatchAny("xlsx:set-print-options", {
        sheet: activeSheet.name,
        [flag]: value,
      });
    },
    [activeSheet, dispatchAny]
  );

  const onSetCalcMode = useCallback(
    (mode: "auto" | "autoNoTable" | "manual") => {
      void dispatchAny("xlsx:set-calc-mode", { calcMode: mode });
    },
    [dispatchAny]
  );

  const onSetCalcOnSave = useCallback(
    (value: boolean) => {
      void dispatchAny("xlsx:set-calc-mode", { calcOnSave: value });
    },
    [dispatchAny]
  );

  const onToggleShowFormulas = useCallback(() => {
    if (!activeSheet) return;
    void dispatchAny("xlsx:set-show-formulas", {
      sheet: activeSheet.name,
      show: !sheetView.showFormulas,
    });
  }, [activeSheet, dispatchAny, sheetView.showFormulas]);

  const onSetSheetView = useCallback(
    (view: "normal" | "pageBreakPreview" | "pageLayout") => {
      if (!activeSheet) return;
      void dispatchAny("xlsx:set-sheet-view", { sheet: activeSheet.name, view });
    },
    [activeSheet, dispatchAny]
  );

  const onToggleViewFlag = useCallback(
    (flag: "showGridLines" | "showRowColHeaders" | "showRuler" | "rightToLeft", value: boolean) => {
      if (!activeSheet) return;
      void dispatchAny("xlsx:set-sheet-view", {
        sheet: activeSheet.name,
        [flag]: value,
      });
    },
    [activeSheet, dispatchAny]
  );

  const onSetZoom = useCallback(
    (zoom: number) => {
      if (!activeSheet) return;
      void dispatchAny("xlsx:set-sheet-view", { sheet: activeSheet.name, zoomScale: zoom });
    },
    [activeSheet, dispatchAny]
  );

  const onSubmitPageSetup = useCallback(
    (submit: PageSetupSubmit) => {
      if (!activeSheet) return;
      const sheet = activeSheet.name;
      if (submit.setup) {
        const s = submit.setup;
        const payload: Record<string, unknown> = { sheet, orientation: s.orientation, paperSize: s.paperSize };
        if (s.fitToWidth !== null || s.fitToHeight !== null) {
          payload.fitToWidth = s.fitToWidth ?? 1;
          payload.fitToHeight = s.fitToHeight ?? 1;
        } else if (s.scale !== null) {
          payload.scale = s.scale;
        }
        payload.blackAndWhite = s.blackAndWhite;
        payload.draft = s.draft;
        void dispatchAny("xlsx:set-page-setup", payload);
      }
      if (submit.margins) {
        const m = submit.margins;
        if (m.preset === "custom") {
          void dispatchAny("xlsx:set-page-margins", {
            sheet,
            leftIn: m.leftIn,
            rightIn: m.rightIn,
            topIn: m.topIn,
            bottomIn: m.bottomIn,
            headerIn: m.headerIn,
            footerIn: m.footerIn,
          });
        } else {
          void dispatchAny("xlsx:set-page-margins", { sheet, preset: m.preset });
        }
      }
      if (submit.printOptions) {
        void dispatchAny("xlsx:set-print-options", { sheet, ...submit.printOptions });
      }
      if (submit.printArea) {
        if (submit.printArea.range === null) {
          void dispatchAny("xlsx:set-print-area", { sheet, clear: true });
        } else {
          void dispatchAny("xlsx:set-print-area", { sheet, range: submit.printArea.range });
        }
      }
      if (submit.printTitles) {
        const { rows, cols } = submit.printTitles;
        if (rows === null && cols === null) {
          void dispatchAny("xlsx:set-print-titles", { sheet, clear: true });
        } else {
          const payload: Record<string, unknown> = { sheet };
          if (rows !== null) payload.rows = rows;
          if (cols !== null) payload.cols = cols;
          void dispatchAny("xlsx:set-print-titles", payload);
        }
      }
    },
    [activeSheet, dispatchAny]
  );

  const onProtectSheetSubmit = useCallback(
    (values: ProtectSheetValues) => {
      if (!activeSheet) return;
      void dispatchAny("xlsx:set-sheet-protection", {
        sheet: activeSheet.name,
        enabled: true,
        ...values,
      });
    },
    [activeSheet, dispatchAny]
  );

  const onProtectSheetClear = useCallback(() => {
    if (!activeSheet) return;
    void dispatchAny("xlsx:set-sheet-protection", { sheet: activeSheet.name, enabled: false });
  }, [activeSheet, dispatchAny]);

  const onProtectWorkbookSubmit = useCallback(
    (values: ProtectWorkbookValues) => {
      void dispatchAny("xlsx:set-workbook-protection", {
        enabled: true,
        ...values,
      });
    },
    [dispatchAny]
  );

  const onProtectWorkbookClear = useCallback(() => {
    void dispatchAny("xlsx:set-workbook-protection", { enabled: false });
  }, [dispatchAny]);

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
  const canTextToColumns = !!(activeSheet && selection && selection.anchor.col === selection.focus.col);
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

  // ── AutoFilter (P17) ───────────────────────────────────────────────
  // Filter dropdown anchor + active column. The dropdown is rendered
  // at the editor surface level (not inside the Grid) so it can
  // escape the scrollable viewport's clip rect.
  const [filterDropdown, setFilterDropdown] = useState<{
    colId: number;
    anchor: DOMRect;
  } | null>(null);

  // C5 — Format Cells dialog. `null` = closed; otherwise the value
  // names the tab the dialog should open onto so different entry
  // points (Mod+1 vs context-menu "Borders…" vs the toolbar borders
  // splitter) can land the user where they expect.
  const [formatCellsTab, setFormatCellsTab] = useState<FormatTabId | null>(null);

  // C7 — Paste Special dialog. Opened on Cmd+Shift+V (Excel parity)
  // and from the context menu / command palette. The dialog itself
  // doesn't read the clipboard — confirm bounces back into
  // `pasteAtSelection` with the chosen mode/transpose.
  const [pasteSpecialOpen, setPasteSpecialOpen] = useState<boolean>(false);

  // C10 — Conditional Formatting dialog. Lists existing typed rules
  // for the active sheet plus a "New rule…" form. Opaque rules
  // imported from the original file aren't surfaced here (they round-
  // trip via `Sheet.opaqueConditionalFormats` and stay byte-equal).
  const [conditionalFormatOpen, setConditionalFormatOpen] = useState<boolean>(false);

  // C11 — Data Validation dialog. Same shape as the CF dialog above.
  const [dataValidationOpen, setDataValidationOpen] = useState<boolean>(false);

  // C8 — Format Painter state. `null` = inactive. When active we
  // hold onto the source clipboard snapshot (formats only) so each
  // subsequent target click/drag can reapply the captured formatting.
  // `sticky=true` mirrors Excel's double-click-to-pin behaviour: the
  // painter stays active across multiple paints until the user hits
  // Esc or clicks the toolbar button again.
  const [formatPainter, setFormatPainter] = useState<{
    sheet: string;
    snap: XlsxClipboardSnapshot;
    sticky: boolean;
  } | null>(null);
  const formatPainterRef = useRef(formatPainter);
  useEffect(() => {
    formatPainterRef.current = formatPainter;
  }, [formatPainter]);

  // Late-bound entry point for the Format Painter activation. We
  // need to reference the activation callback from `onSurfaceKeyDown`
  // (declared earlier in the file) without creating a TDZ cycle, so
  // we expose it via a ref that's updated after the callback is
  // defined further below.
  const activateFormatPainterRef = useRef<(sticky: boolean) => void>(() => {});

  // Auto-detect the used range on a sheet — the smallest A1 rectangle
  // covering every populated cell. Mirrors Excel's behaviour when the
  // user toggles AutoFilter from a single-cell selection.
  const detectUsedRange = useCallback((sheet: Sheet): string | null => {
    let minR = Infinity;
    let minC = Infinity;
    let maxR = -1;
    let maxC = -1;
    for (const cell of sheet.cells.values()) {
      if (cell.value === null || cell.value === "") continue;
      if (cell.row < minR) minR = cell.row;
      if (cell.col < minC) minC = cell.col;
      if (cell.row > maxR) maxR = cell.row;
      if (cell.col > maxC) maxC = cell.col;
    }
    if (maxR < 0) return null;
    return formatRange({
      start: { row: minR, col: minC },
      end: { row: maxR, col: maxC },
    });
  }, []);

  const onFreeze = useCallback(
    (rows: number, cols: number) => {
      if (!activeSheet) return;
      if (rows === 0 && cols === 0) {
        dispatchOrToast("xlsx:unfreeze-panes", { sheet: activeSheet.name });
        return;
      }
      dispatchOrToast("xlsx:freeze-panes", { sheet: activeSheet.name, rows, cols });
    },
    [activeSheet, dispatchOrToast]
  );

  /**
   * C9 — Sheet management. Each callback dispatches a single
   * command and lets the bus handle undo/redo + dirty tracking.
   * Errors surface through the existing toast queue so the user
   * always sees Excel's own validation messages (duplicate sheet
   * name, last visible sheet, etc.).
   */
  const onAddSheet = useCallback(() => {
    const a = agentRef.current;
    if (!a) return;
    const existing = new Set(snapshot?.root.sheets.map((s) => s.name.toLowerCase()) ?? []);
    let i = snapshot?.root.sheets.length ?? 0;
    let name = `Sheet${i + 1}`;
    while (existing.has(name.toLowerCase())) {
      i += 1;
      name = `Sheet${i + 1}`;
    }
    void a
      .applyCommand({ type: "xlsx:add-sheet", payload: { name }, source: "human" })
      .then(() => setActiveSheetName(name))
      .catch((err: unknown) => pushToast("error", err instanceof Error ? err.message : String(err)));
  }, [snapshot, pushToast]);

  const onRenameSheet = useCallback(
    (currentName: string, nextName: string) => {
      dispatchOrToast("xlsx:rename-sheet", { name: currentName, newName: nextName });
      setActiveSheetName((prev) => (prev === currentName ? nextName : prev));
    },
    [dispatchOrToast]
  );

  const onDeleteSheet = useCallback(
    (name: string) => {
      const sheets = snapshot?.root.sheets ?? [];
      const idx = sheets.findIndex((s) => s.name === name);
      const fallback =
        sheets.find((s) => s.name !== name && s.state === "visible")?.name ??
        sheets.find((s) => s.name !== name)?.name ??
        null;
      dispatchOrToast("xlsx:delete-sheet", { name });
      setActiveSheetName((prev) => (prev === name ? fallback : prev));
      void idx;
    },
    [snapshot, dispatchOrToast]
  );

  const onMoveSheet = useCallback(
    (name: string, to: number) => {
      dispatchOrToast("xlsx:move-sheet", { name, to });
    },
    [dispatchOrToast]
  );

  const onSetSheetState = useCallback(
    (name: string, state: "visible" | "hidden" | "veryHidden") => {
      dispatchOrToast("xlsx:set-sheet-state", { name, state });
      // If we just hid the active sheet, swing focus to the next
      // visible sheet so the editor never bottoms out on a hidden
      // surface.
      if (state !== "visible") {
        const sheets = snapshot?.root.sheets ?? [];
        setActiveSheetName((prev) => {
          if (prev !== name) return prev;
          const next = sheets.find((s) => s.name !== name && s.state === "visible");
          return next?.name ?? prev;
        });
      }
    },
    [dispatchOrToast, snapshot]
  );

  // C10 — Conditional Formatting commands. Each one routes through
  // dispatchOrToast so undo/redo + dirty tracking + toast on error
  // come for free.
  const onAddConditionalFormat = useCallback(
    (rule: ConditionalFormat) => {
      if (!activeSheet) return;
      dispatchOrToast("xlsx:add-conditional-format", { sheet: activeSheet.name, rule });
    },
    [activeSheet, dispatchOrToast]
  );

  const onRemoveConditionalFormat = useCallback(
    (id: string) => {
      if (!activeSheet) return;
      dispatchOrToast("xlsx:remove-conditional-format", { sheet: activeSheet.name, id });
    },
    [activeSheet, dispatchOrToast]
  );

  const onClearConditionalFormats = useCallback(() => {
    if (!activeSheet) return;
    dispatchOrToast("xlsx:clear-conditional-formats", { sheet: activeSheet.name });
  }, [activeSheet, dispatchOrToast]);

  // C11 — Data Validation commands.
  const onAddDataValidation = useCallback(
    (rule: DataValidation) => {
      if (!activeSheet) return;
      dispatchOrToast("xlsx:add-data-validation", { sheet: activeSheet.name, rule });
    },
    [activeSheet, dispatchOrToast]
  );

  const onRemoveDataValidation = useCallback(
    (id: string) => {
      if (!activeSheet) return;
      dispatchOrToast("xlsx:remove-data-validation", { sheet: activeSheet.name, id });
    },
    [activeSheet, dispatchOrToast]
  );

  const onClearDataValidations = useCallback(() => {
    if (!activeSheet) return;
    dispatchOrToast("xlsx:clear-data-validations", { sheet: activeSheet.name });
  }, [activeSheet, dispatchOrToast]);

  // C12 — Defined names (named ranges).
  const definedNames = useMemo<ReadonlyArray<DefinedName>>(
    () => snapshot?.root.definedNames ?? [],
    [snapshot]
  );
  const [nameManagerOpen, setNameManagerOpen] = useState<boolean>(false);

  /**
   * Format the current selection as an absolute, sheet-qualified
   * `Sheet1!$A$1:$C$5` reference. Used as the default `refersTo`
   * value when minting a new defined name from the Name Box or the
   * Name Manager dialog.
   */
  const selectionRefersTo = useMemo(() => {
    if (!activeSheet || !selection) return "";
    const r = selectionToRange(selection);
    const a1 = (row: number, col: number) => `$${colToLetter(col)}$${row + 1}`;
    const sheet = activeSheet.name.match(/^[A-Za-z_][\w. ]*$/)
      ? activeSheet.name
      : `'${activeSheet.name.replace(/'/g, "''")}'`;
    if (r.start.row === r.end.row && r.start.col === r.end.col) {
      return `${sheet}!${a1(r.start.row, r.start.col)}`;
    }
    return `${sheet}!${a1(r.start.row, r.start.col)}:${a1(r.end.row, r.end.col)}`;
  }, [activeSheet, selection]);

  const onAddDefinedName = useCallback(
    (entry: { name: string; refersTo: string; scope?: string; comment?: string }) => {
      dispatchOrToast("xlsx:add-defined-name", entry);
    },
    [dispatchOrToast]
  );

  const onUpdateDefinedName = useCallback(
    (entry: { name: string; scope?: string; nextName?: string; refersTo?: string; comment?: string }) => {
      dispatchOrToast("xlsx:update-defined-name", entry);
    },
    [dispatchOrToast]
  );

  const onRemoveDefinedName = useCallback(
    (entry: { name: string; scope?: string }) => {
      dispatchOrToast("xlsx:remove-defined-name", entry);
    },
    [dispatchOrToast]
  );

  /**
   * Resolve a Name Box input to a navigable target: a defined name
   * resolves to the range it points at; an A1 cell or range
   * resolves to itself. Returns true when the input could be
   * navigated, false otherwise (lets the caller treat the input as
   * a "create-new" intent).
   */
  const onJumpFromNameBox = useCallback(
    (input: string): boolean => {
      const trimmed = input.trim();
      if (!trimmed) return false;
      // Sheet-qualified ref (`Sheet1!A1` or `'My Sheet'!A1:C5`).
      const sheetQualified = /^(?:'((?:[^']|'')+)'|([A-Za-z_][\w. ]*))!(.+)$/.exec(trimmed);
      let targetSheetName: string | undefined;
      let body = trimmed;
      if (sheetQualified) {
        targetSheetName = (sheetQualified[1] ?? sheetQualified[2] ?? "").replace(/''/g, "'");
        body = sheetQualified[3]!;
      }
      const cleaned = body.replace(/\$/g, "").toUpperCase();
      const tryNavigateRef = (sheetName: string | undefined, ref: string): boolean => {
        try {
          if (ref.includes(":")) {
            const range = parseRange(ref);
            const sheetTarget =
              sheetName && snapshot?.root.sheets.some((s) => s.name === sheetName)
                ? sheetName
                : activeSheetName;
            if (sheetTarget && sheetTarget !== activeSheetName) setActiveSheetName(sheetTarget);
            setSelection({
              anchor: { row: range.start.row, col: range.start.col },
              focus: { row: range.end.row, col: range.end.col },
            });
            return true;
          }
          const cell = parseA1(ref);
          if (cell) {
            const sheetTarget =
              sheetName && snapshot?.root.sheets.some((s) => s.name === sheetName)
                ? sheetName
                : activeSheetName;
            if (sheetTarget && sheetTarget !== activeSheetName) setActiveSheetName(sheetTarget);
            setSelection(singleSelection({ row: cell.row, col: cell.col }));
            return true;
          }
        } catch {
          /* fall through */
        }
        return false;
      };
      // 1) Try a direct cell/range parse first.
      if (tryNavigateRef(targetSheetName, cleaned)) return true;
      // 2) Try a defined-name lookup (workbook + active sheet scope).
      const dn = definedNames.find(
        (d) => d.name === trimmed && (d.scope === undefined || d.scope === activeSheetName)
      );
      if (dn) {
        const m = /^(?:'((?:[^']|'')+)'|([A-Za-z_][\w. ]*))!(.+)$/.exec(dn.refersTo.trim().replace(/^=/, ""));
        const sheetName = m
          ? (m[1] ?? m[2] ?? activeSheetName ?? "").replace(/''/g, "'")
          : (activeSheetName ?? "");
        const refBody = (m ? m[3]! : dn.refersTo.trim().replace(/^=/, "")).replace(/\$/g, "");
        return tryNavigateRef(sheetName || undefined, refBody.toUpperCase());
      }
      return false;
    },
    [definedNames, activeSheetName, snapshot]
  );

  const onCreateNameFromBox = useCallback(
    (name: string) => {
      if (!selectionRefersTo) {
        pushToast("warn", "Select a range first");
        return;
      }
      onAddDefinedName({ name, refersTo: selectionRefersTo });
    },
    [selectionRefersTo, onAddDefinedName, pushToast]
  );

  const onToggleFilter = useCallback(() => {
    if (!activeSheet) return;
    if (activeSheet.autoFilter) {
      dispatchOrToast("xlsx:set-auto-filter", { sheet: activeSheet.name, range: null });
      setFilterDropdown(null);
      return;
    }
    let range: string | null = null;
    if (selection && !isSingle(selection)) {
      range = formatSelection(selection);
    } else {
      range = detectUsedRange(activeSheet);
    }
    if (!range) {
      pushToast("warn", "No data to filter on this sheet");
      return;
    }
    dispatchOrToast("xlsx:set-auto-filter", { sheet: activeSheet.name, range });
  }, [activeSheet, selection, detectUsedRange, dispatchOrToast, pushToast]);

  const onOpenFilter = useCallback((colId: number, anchor: DOMRect) => {
    setFilterDropdown({ colId, anchor });
  }, []);

  const onCloseFilter = useCallback(() => setFilterDropdown(null), []);

  const onApplyFilterColumn = useCallback(
    (criterion: import("@officeai/xlsx").FilterColumn) => {
      if (!activeSheet || !filterDropdown) return;
      dispatchOrToast("xlsx:set-filter-column", {
        sheet: activeSheet.name,
        colId: filterDropdown.colId,
        criterion,
      });
      setFilterDropdown(null);
    },
    [activeSheet, filterDropdown, dispatchOrToast]
  );

  const onClearFilterColumn = useCallback(() => {
    if (!activeSheet || !filterDropdown) return;
    dispatchOrToast("xlsx:clear-filter-column", {
      sheet: activeSheet.name,
      colId: filterDropdown.colId,
    });
    setFilterDropdown(null);
  }, [activeSheet, filterDropdown, dispatchOrToast]);

  const onSortFromFilter = useCallback(
    (order: "asc" | "desc") => {
      if (!activeSheet || !activeSheet.autoFilter || !filterDropdown) return;
      const af = activeSheet.autoFilter;
      const range = formatRange({
        start: { row: af.range.r1, col: af.range.c1 },
        end: { row: af.range.r2, col: af.range.c2 },
      });
      dispatchOrToast("xlsx:sort-range", {
        sheet: activeSheet.name,
        range,
        sortBy: { colId: filterDropdown.colId, order },
      });
      setFilterDropdown(null);
    },
    [activeSheet, filterDropdown, dispatchOrToast]
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

  /**
   * C6 — Borders splitter dispatcher.
   *
   * For uniform-side presets (all, top, bottom, left, right,
   * top-bottom, top-thick-bottom, none) we dispatch a single
   * `xlsx:set-cell-format` over the whole selection. For outside /
   * thick-outside on a multi-cell selection we dispatch up to 4
   * sub-range patches so the perimeter cells get the right sides.
   *
   * Notes:
   *   - "none" clears all 4 sides via `style: "none"` (the patch
   *     handler interprets that as a deletion).
   *   - We always include `border` even for single-side presets so
   *     the existing other sides stay intact.
   */
  const onApplyBorderPreset = useCallback(
    (preset: BorderPreset) => {
      const a = agentRef.current;
      if (!a || !activeSheet || !selection) return;
      const sheet = activeSheet.name;

      const thin = (color?: string): { style: "thin"; color?: string } =>
        color ? { style: "thin", color } : { style: "thin" };
      const thick = (color?: string): { style: "medium"; color?: string } =>
        color ? { style: "medium", color } : { style: "medium" };
      const noneSide = { style: "none" as const };

      // C13 — Build the full command list across every area and
      // dispatch as one batch so undo / redo treats the multi-area
      // border change as a single user gesture.
      const allCmds: Array<{
        type: "xlsx:set-cell-format";
        payload: SetCellFormatPayload;
        source: "human";
      }> = [];

      const buildForArea = (area: Selection): void => {
        const n = normalizeSelection(area);
        const wholeRange = formatSelection(area);
        const push = (range: string, format: CellFormatPatch) =>
          allCmds.push({
            type: "xlsx:set-cell-format",
            payload: { sheet, range, format },
            source: "human",
          });

        switch (preset) {
          case "all":
            push(wholeRange, {
              border: { top: thin(), right: thin(), bottom: thin(), left: thin() },
            });
            return;
          case "none":
            push(wholeRange, {
              border: { top: noneSide, right: noneSide, bottom: noneSide, left: noneSide },
            });
            return;
          case "top":
            push(wholeRange, { border: { top: thin() } });
            return;
          case "bottom":
            push(wholeRange, { border: { bottom: thin() } });
            return;
          case "left":
            push(wholeRange, { border: { left: thin() } });
            return;
          case "right":
            push(wholeRange, { border: { right: thin() } });
            return;
          case "top-bottom":
            push(wholeRange, { border: { top: thin(), bottom: thin() } });
            return;
          case "top-thick-bottom":
            push(wholeRange, { border: { top: thin(), bottom: thick() } });
            return;
          case "outside":
          case "thick-outside": {
            const side = preset === "thick-outside" ? thick : thin;
            if (n.r0 === n.r1 && n.c0 === n.c1) {
              push(wholeRange, {
                border: { top: side(), right: side(), bottom: side(), left: side() },
              });
              return;
            }
            push(formatRange({ start: { row: n.r0, col: n.c0 }, end: { row: n.r0, col: n.c1 } }), {
              border: { top: side() },
            });
            push(formatRange({ start: { row: n.r1, col: n.c0 }, end: { row: n.r1, col: n.c1 } }), {
              border: { bottom: side() },
            });
            push(formatRange({ start: { row: n.r0, col: n.c0 }, end: { row: n.r1, col: n.c0 } }), {
              border: { left: side() },
            });
            push(formatRange({ start: { row: n.r0, col: n.c1 }, end: { row: n.r1, col: n.c1 } }), {
              border: { right: side() },
            });
            return;
          }
          default: {
            const _exhaustive: never = preset;
            void _exhaustive;
          }
        }
      };

      for (const area of allAreas(selection, extraAreas)) buildForArea(area);

      void a.applyCommands(allCmds).catch((err: unknown) => pushToast("error", String(err)));
    },
    [activeSheet, selection, extraAreas, pushToast]
  );

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

  // C7 — Confirm callback from the Paste Special dialog. The dialog
  // owns no clipboard state; on confirm we close it and bounce back
  // into `pasteAtSelection` with the chosen mode/transpose so the
  // existing async clipboard read + permission dance stays in one
  // place.
  const onPasteSpecialConfirm = useCallback(
    (opts: PasteSpecialOptions) => {
      setPasteSpecialOpen(false);
      void pasteAtSelection(undefined, opts);
    },
    [pasteAtSelection]
  );

  // C8 — Format Painter activate / cancel. Single click pins for one
  // paint, double-click pins until cancelled. When activated without
  // a selection we no-op + toast (same as Excel: nothing to copy
  // from). Re-activating while already on toggles off so the
  // toolbar button feels like a regular toggle.
  const activateFormatPainter = useCallback(
    (sticky: boolean) => {
      const a = agentRef.current;
      if (!a || !activeSheet || !selection) {
        pushToast("warn", "Select a cell first to copy its format.");
        return;
      }
      // Toggle off if already active (matches the toolbar button's
      // press/unpress affordance).
      if (formatPainter) {
        setFormatPainter(null);
        return;
      }
      try {
        const range = formatSelection(selection);
        const snap = a.getClipboardSnapshot({ sheet: activeSheet.name, range });
        setFormatPainter({ sheet: activeSheet.name, snap, sticky });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [activeSheet, selection, formatPainter, pushToast]
  );

  // Mirror the activate callback into the late-bound ref so the
  // earlier `onSurfaceKeyDown` handler can dispatch it without
  // creating a TDZ-illegal forward reference in its dependency
  // array.
  useEffect(() => {
    activateFormatPainterRef.current = activateFormatPainter;
  }, [activateFormatPainter]);

  // C8 — Track whether the user is *currently* dragging out the
  // destination range while Format Painter is armed. We set this
  // when `handleGridSelect` is called with `extend=false` (a fresh
  // mousedown on the grid) and clear it on the global `mouseup`,
  // at which point we apply the captured format to whatever the
  // selection ended up covering.
  const formatPainterPendingRef = useRef<boolean>(false);

  /**
   * C8 — Apply the captured Format Painter source to the current
   * selection. If the source is 1×1 we expand it across the
   * destination grid so the single style fills the whole target
   * range (Excel parity). For multi-cell sources we paste once at
   * the destination's anchor with the source's natural HxW.
   */
  const applyFormatPainterAt = useCallback(
    (target: Selection) => {
      const a = agentRef.current;
      const fp = formatPainterRef.current;
      if (!a || !fp) return;
      const sh = snapshotRef.current?.root.sheets.find((s) => s.name === fp.sheet);
      if (!sh) return;
      const n = normalizeSelection(target);
      const targetH = n.r1 - n.r0 + 1;
      const targetW = n.c1 - n.c0 + 1;

      let source: XlsxClipboardSnapshot = fp.snap;
      if (fp.snap.height === 1 && fp.snap.width === 1) {
        const cell = fp.snap.cells[0]?.[0] ?? null;
        const expanded: XlsxClipboardSnapshot = {
          origin: fp.snap.origin,
          width: targetW,
          height: targetH,
          cells: Array.from({ length: targetH }, () =>
            Array.from({ length: targetW }, () => (cell ? { ...cell } : null))
          ),
          merges: [],
        };
        source = expanded;
      }

      void a
        .applyCommand({
          type: "xlsx:paste-range",
          payload: {
            sheet: fp.sheet,
            target: formatA1({ row: n.r0, col: n.c0 }),
            source,
            mode: "formats",
          },
          source: "human",
        })
        .catch((err: unknown) => pushToast("error", err instanceof Error ? err.message : String(err)));

      if (!fp.sticky) setFormatPainter(null);
    },
    [pushToast]
  );

  // C8 — Global mouseup listener that finishes a Format Painter
  // drop. When the user mousedowns on the grid while painter is on
  // we set `formatPainterPendingRef`; the eventual mouseup reads
  // the latest selection from `selectionRef` and dispatches the
  // formats paste. Esc cancels the painter without dropping.
  useEffect(() => {
    const onUp = () => {
      if (!formatPainterPendingRef.current) return;
      formatPainterPendingRef.current = false;
      const sel = selectionRef.current;
      if (sel) applyFormatPainterAt(sel);
    };
    window.addEventListener("mouseup", onUp);
    return () => window.removeEventListener("mouseup", onUp);
  }, [applyFormatPainterAt]);

  // Esc cancels Format Painter even outside the surface focus.
  useEffect(() => {
    if (!formatPainter) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        e.preventDefault();
        setFormatPainter(null);
        formatPainterPendingRef.current = false;
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [formatPainter]);

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
      {
        kind: "action",
        id: "paste-special",
        label: "Paste Special…",
        shortcut: "⇧⌘V",
        disabled: !canCopyHere,
        onSelect: () => setPasteSpecialOpen(true),
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
      { kind: "divider", id: "div-format" },
      {
        kind: "action",
        id: "format-cells",
        label: "Format cells…",
        shortcut: "⌘1",
        onSelect: () => setFormatCellsTab("number"),
      },
      {
        kind: "action",
        id: "data-validation",
        label: "Data validation…",
        onSelect: () => setDataValidationOpen(true),
      },
      {
        kind: "action",
        id: "conditional-format",
        label: "Conditional formatting…",
        onSelect: () => setConditionalFormatOpen(true),
      },
    ];
    if (target.kind === "image") {
      const imageId = target.imageId;
      return [
        {
          kind: "action",
          id: "delete-image",
          label: "Delete image",
          shortcut: "⌫",
          onSelect: () => onRemoveImage(imageId),
        },
      ];
    }
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
        { kind: "divider", id: "div-col-filter" },
        {
          kind: "action",
          id: "sort-asc",
          label: "Sort A → Z",
          disabled: !activeSheet?.autoFilter,
          onSelect: () => {
            if (!activeSheet?.autoFilter) return;
            const af = activeSheet.autoFilter;
            const range = formatRange({
              start: { row: af.range.r1, col: af.range.c1 },
              end: { row: af.range.r2, col: af.range.c2 },
            });
            const colId = target.col - af.range.c1;
            if (colId < 0 || colId > af.range.c2 - af.range.c1) return;
            dispatchOrToast("xlsx:sort-range", {
              sheet: activeSheet.name,
              range,
              sortBy: { colId, order: "asc" },
            });
          },
        },
        {
          kind: "action",
          id: "sort-desc",
          label: "Sort Z → A",
          disabled: !activeSheet?.autoFilter,
          onSelect: () => {
            if (!activeSheet?.autoFilter) return;
            const af = activeSheet.autoFilter;
            const range = formatRange({
              start: { row: af.range.r1, col: af.range.c1 },
              end: { row: af.range.r2, col: af.range.c2 },
            });
            const colId = target.col - af.range.c1;
            if (colId < 0 || colId > af.range.c2 - af.range.c1) return;
            dispatchOrToast("xlsx:sort-range", {
              sheet: activeSheet.name,
              range,
              sortBy: { colId, order: "desc" },
            });
          },
        },
        {
          kind: "action",
          id: "clear-filter-from-col",
          label: "Clear filter from column",
          disabled:
            !activeSheet?.autoFilter ||
            !activeSheet.autoFilter.columns.has(target.col - activeSheet.autoFilter.range.c1),
          onSelect: () => {
            if (!activeSheet?.autoFilter) return;
            const colId = target.col - activeSheet.autoFilter.range.c1;
            dispatchOrToast("xlsx:clear-filter-column", { sheet: activeSheet.name, colId });
          },
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
    dispatchOrToast,
    onRemoveImage,
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

  // C2 — Status-bar selection summary, Excel-parity: chips for Sum,
  // Avg, Count (non-empty), Numerical Count, Min, Max. We *iterate
  // the sparse `cells` map and filter to the selection rectangle*
  // rather than walking every (r, c) in range — Ctrl+A on a 1M ×
  // 16K sheet would otherwise OOM the main thread. The map is
  // already keyed by the cells the user has touched.
  // C10 — Pre-compute conditional-format overlays for the active
  // sheet. Memoised on `activeSheet` so we recompute on cell edits
  // (which produce a fresh sheet object) and on rule changes
  // (which also re-emit the sheet via the bus).
  const cfOverlays = useMemo(() => {
    if (!activeSheet) return undefined;
    if (activeSheet.conditionalFormats.length === 0) return undefined;
    return evaluateConditionalFormats(activeSheet);
  }, [activeSheet]);

  // C11 — Per-cell data-validation index. Maps `r:c` → resolved
  // dropdown options for cells covered by a typed `list` rule. We
  // only resolve literal lists here; formula refs are surfaced as
  // an empty option list with a placeholder hint so the user still
  // sees the dropdown arrow.
  const dvIndex = useMemo(() => {
    if (!activeSheet || activeSheet.dataValidations.length === 0) return undefined;
    const out = new Map<string, ReadonlyArray<string>>();
    for (const dv of activeSheet.dataValidations) {
      if (dv.kind !== "list") continue;
      let range;
      try {
        range = parseRange(dv.range);
      } catch {
        continue;
      }
      const options: string[] = dv.formula
        ? resolveFormulaListOptions(dv.source, activeSheet)
        : dv.source
            .split(",")
            .map((s) => s.trim())
            .filter((s) => s.length > 0);
      for (let r = range.start.row; r <= range.end.row; r++) {
        for (let c = range.start.col; c <= range.end.col; c++) {
          out.set(cellKey(r, c), options);
        }
      }
    }
    return out;
  }, [activeSheet]);

  const selectionAggregates = useMemo(() => {
    if (!activeSheet || !selection) return undefined;
    const areas = allAreas(selection, extraAreas);
    // Span = number of cells covered by the union (rectangle sum,
    // not deduped — overlapping ranges are rare and a small over-
    // estimate is harmless here). We compute this from rectangle
    // math so a "select entire column" gesture stays O(1) instead of
    // walking 1M coordinates.
    const span = unionSpanUpperBound(areas);
    if (span < 2) return undefined;
    let sum = 0;
    let count = 0;
    let countNum = 0;
    let min = Number.POSITIVE_INFINITY;
    let max = Number.NEGATIVE_INFINITY;
    // Iterate the SPARSE cells map and filter to the union; this is
    // O(numPopulatedCells) regardless of how large the selection
    // rectangle is. The previous code paid O(spanOfRectangle) which
    // froze the page on whole-column selections.
    for (const cell of activeSheet.cells.values()) {
      if (!areasContainCell(areas, cell.row, cell.col)) continue;
      const v: CellValue = cell.value;
      if (v === null || v === undefined) continue;
      count += 1;
      if (typeof v === "number" && Number.isFinite(v)) {
        countNum += 1;
        sum += v;
        if (v < min) min = v;
        if (v > max) max = v;
      }
    }
    if (count === 0) return undefined;
    const fmt = (n: number) => (Math.abs(n) >= 1e6 ? n.toExponential(2) : Number(n.toFixed(4)).toString());
    const out: { readonly label: string; readonly value: string }[] = [];
    if (countNum > 0) {
      out.push({ label: "Sum", value: fmt(sum) });
      out.push({ label: "Avg", value: fmt(sum / countNum) });
      out.push({ label: "Min", value: fmt(min) });
      out.push({ label: "Max", value: fmt(max) });
    }
    // "Count" mirrors Excel's "Count" (non-empty cells); when there
    // are non-numeric values mixed in we expose the numeric count
    // separately so the user can disambiguate.
    out.push({ label: "Count", value: String(count) });
    if (countNum > 0 && countNum !== count) {
      out.push({ label: "Numerical Count", value: String(countNum) });
    }
    return out;
  }, [activeSheet, selection, extraAreas]);

  const selectionText = useMemo(() => {
    if (!activeSheet) return "";
    if (selectionAggregates && selectionAggregates.length > 0) return undefined;
    const ref = selectedRef || "—";
    return `${activeSheet.name}!${ref}`;
  }, [activeSheet, selectionAggregates, selectedRef]);

  // Open comment count for the right-rail badge.
  const openCommentCount = useMemo(() => {
    if (!activeSheet) return 0;
    let n = 0;
    for (const c of activeSheet.comments) {
      if (c.parentId) continue;
      if (c.resolved) continue;
      n += 1;
    }
    return n;
  }, [activeSheet]);

  // C4 — Find/Replace adapter for XLSX. Excel-parity behaviour:
  //   * One {@link FindMatch} per *occurrence* (not per cell), so
  //     "Find Next" walks through every hit, including multiple
  //     matches inside a single cell.
  //   * Search runs against displayed text and (when the cell carries
  //     a formula) the formula text prefixed with `=`. Replacements
  //     write back through `xlsx:set-cell-value`, so a formula edit
  //     keeps the formula and a literal edit re-parses through
  //     `parseLiteral` (so "12,300" doesn't decay into a string).
  //   * Walks cells in row-major order. With the C1 sparse model this
  //     stays linear in the number of *populated* cells regardless of
  //     the 16K × 1M virtual bounds.
  //   * `gotoMatch` snaps the selection to the cell and fires the
  //     existing scroll-into-view-with-flash plumbing the comments
  //     rail uses, so the user always sees what was found.
  const findAdapter = useMemo<FindAdapter | undefined>(() => {
    if (!activeSheet) return undefined;
    const sheet = activeSheet;
    const buildRegex = (q: string, opts: FindOptions): RegExp | null => {
      if (q.length === 0) return null;
      try {
        const flags = opts.caseSensitive ? "g" : "gi";
        const body = opts.regex ? q : q.replace(/[-/\\^$*+?.()|[\]{}]/g, "\\$&");
        const wrapped = opts.wholeWord ? `\\b${body}\\b` : body;
        return new RegExp(wrapped, flags);
      } catch {
        return null;
      }
    };
    const cellText = (cell: { formula?: { text: string }; value: unknown }): string => {
      if (cell.formula) return `=${cell.formula.text}`;
      if (cell.value === null || cell.value === undefined) return "";
      return String(cell.value);
    };
    const cellsInOrder = (): Array<{ row: number; col: number; text: string }> => {
      const arr: Array<{ row: number; col: number; text: string }> = [];
      for (const cell of sheet.cells.values()) {
        const t = cellText(cell);
        if (t.length === 0) continue;
        arr.push({ row: cell.row, col: cell.col, text: t });
      }
      arr.sort((a, b) => a.row - b.row || a.col - b.col);
      return arr;
    };
    return {
      findAll(query, opts) {
        const re = buildRegex(query, opts);
        if (!re) return [];
        const results: FindMatch[] = [];
        for (const c of cellsInOrder()) {
          re.lastIndex = 0;
          let m: RegExpExecArray | null;
          while ((m = re.exec(c.text)) !== null) {
            const ref = formatA1({ row: c.row, col: c.col });
            results.push({
              id: `${c.row}:${c.col}:${m.index}:${m[0].length}`,
              preview: `${sheet.name}!${ref}  ${c.text}`,
            });
            if (m[0].length === 0) re.lastIndex += 1;
          }
        }
        return results;
      },
      gotoMatch(match) {
        const [r, c] = match.id.split(":").map(Number);
        if (Number.isFinite(r) && Number.isFinite(c)) {
          setSelection(singleSelection({ row: r, col: c }));
          setCommentScrollTarget((prev) => ({
            row: r,
            col: c,
            nonce: (prev?.nonce ?? 0) + 1,
          }));
        }
      },
      async replaceMatch(match, replacement) {
        const a = agentRef.current;
        if (!a) return;
        const parts = match.id.split(":");
        const r = Number(parts[0]);
        const c = Number(parts[1]);
        const idx = Number(parts[2] ?? "");
        const len = Number(parts[3] ?? "");
        if (!Number.isFinite(r) || !Number.isFinite(c) || !Number.isFinite(idx) || !Number.isFinite(len))
          return;
        const cell = sheet.cells.get(cellKey(r, c));
        if (!cell) return;
        const text = cellText(cell);
        if (idx < 0 || idx + len > text.length) return;
        const replaced = text.slice(0, idx) + replacement + text.slice(idx + len);
        if (replaced === text) return;
        await a.applyCommand({
          type: "xlsx:set-cell-value",
          payload: {
            sheet: sheet.name,
            ref: formatA1({ row: r, col: c }),
            value: replaced.startsWith("=") ? replaced : parseLiteral(replaced),
          },
          source: "human",
        });
      },
      async replaceAll(query, replacement, opts) {
        const a = agentRef.current;
        if (!a) return 0;
        const re = buildRegex(query, opts);
        if (!re) return 0;
        let count = 0;
        for (const cell of sheet.cells.values()) {
          const text = cellText(cell);
          if (text.length === 0) continue;
          re.lastIndex = 0;
          if (!re.test(text)) continue;
          re.lastIndex = 0;
          const replaced = text.replace(re, replacement);
          if (replaced === text) continue;
          await a.applyCommand({
            type: "xlsx:set-cell-value",
            payload: {
              sheet: sheet.name,
              ref: formatA1({ row: cell.row, col: cell.col }),
              value: replaced.startsWith("=") ? replaced : parseLiteral(replaced),
            },
            source: "human",
          });
          // Each cell counts once even if it carried multiple matches —
          // matches Excel's "N replacements made" status which counts
          // *cells changed*, not individual occurrences.
          count += 1;
        }
        return count;
      },
    };
  }, [activeSheet]);

  const renderCommentsPanel = useCallback((): ReactNode => {
    if (!agent || !activeSheet) {
      return <div className="p-4 text-sm text-secondary">No sheet selected.</div>;
    }
    return (
      <div className="flex min-h-0 flex-1 flex-col">
        <div className="min-h-0 flex-1 overflow-y-auto">
          <CommentsSidebar
            key={`xlsx-comments-${activeSheet.name}-${revision}`}
            provider={createXlsxCommentsProvider({
              agent,
              sheetName: activeSheet.name,
              onScrollTo: scrollToComment,
            })}
            author="You"
            {...(realtimeRoom.room?.identity
              ? {
                  authorIdentity: {
                    name: realtimeRoom.room.identity.name,
                    id: realtimeRoom.room.identity.id,
                    color: realtimeRoom.room.identity.color,
                  },
                }
              : {})}
            emptyHint="No comments on this sheet yet. Select a cell and press Add comment in the toolbar."
            onScrollTo={scrollToComment}
          />
        </div>
        {selection ? (
          <div className="border-t border-divider p-2">
            <CommentComposer
              provider={createXlsxCommentsProvider({ agent, sheetName: activeSheet.name })}
              anchor={{
                kind: "xlsx-cell",
                sheet: activeSheet.name,
                ref: formatA1(selection.anchor),
              }}
              placeholder={`Comment on ${formatA1(selection.anchor)}…`}
            />
          </div>
        ) : null}
      </div>
    );
  }, [activeSheet, agent, revision, scrollToComment, selection]);

  // Palette is generated from the central xlsx action catalogue (see
  // packages/xlsx/src/actions/catalogue.ts). Labels/sections/shortcuts
  // flow from the catalogue; this map only carries the closure-bound
  // side-effects + per-id `enabled` gating.
  const paletteCommands = useMemo<ReadonlyArray<PaletteCommand>>(() => {
    const runners: PaletteRunners = {
      "xlsx.toggle-filter": { run: () => onToggleFilter(), enabled: Boolean(agent) },
      "xlsx.text-to-columns": { run: () => onTextToColumns(), enabled: canTextToColumns },
      "xlsx.format-cells": { run: () => setFormatCellsTab("number"), enabled: Boolean(agent && selection) },
      "xlsx.format-cells-alignment": {
        run: () => setFormatCellsTab("alignment"),
        enabled: Boolean(agent && selection),
      },
      "xlsx.format-cells-font": {
        run: () => setFormatCellsTab("font"),
        enabled: Boolean(agent && selection),
      },
      "xlsx.format-cells-border": {
        run: () => setFormatCellsTab("border"),
        enabled: Boolean(agent && selection),
      },
      "xlsx.format-cells-fill": {
        run: () => setFormatCellsTab("fill"),
        enabled: Boolean(agent && selection),
      },
      "xlsx.format-cells-protection": {
        run: () => setFormatCellsTab("protection"),
        enabled: Boolean(agent && selection),
      },
      "xlsx.borders-all": { run: () => onApplyBorderPreset("all"), enabled: Boolean(agent && selection) },
      "xlsx.borders-outside": {
        run: () => onApplyBorderPreset("outside"),
        enabled: Boolean(agent && selection),
      },
      "xlsx.borders-thick-outside": {
        run: () => onApplyBorderPreset("thick-outside"),
        enabled: Boolean(agent && selection),
      },
      "xlsx.borders-none": { run: () => onApplyBorderPreset("none"), enabled: Boolean(agent && selection) },
      "xlsx.paste-special": { run: () => setPasteSpecialOpen(true), enabled: Boolean(agent && selection) },
      "xlsx.format-painter": {
        run: () => activateFormatPainter(false),
        enabled: Boolean(agent && selection),
      },
      "xlsx.add-sheet": { run: () => onAddSheet(), enabled: Boolean(agent) },
      "xlsx.rename-sheet": {
        run: () => {
          if (!activeSheet) return;
          const next = window.prompt("New sheet name", activeSheet.name);
          if (next && next.trim() && next.trim() !== activeSheet.name) {
            onRenameSheet(activeSheet.name, next.trim());
          }
        },
        enabled: Boolean(agent && activeSheet),
      },
      "xlsx.delete-sheet": {
        run: () => {
          if (activeSheet) onDeleteSheet(activeSheet.name);
        },
        enabled: Boolean(agent && activeSheet && (snapshot?.root.sheets.length ?? 0) > 1),
      },
      "xlsx.hide-sheet": {
        run: () => {
          if (activeSheet) onSetSheetState(activeSheet.name, "hidden");
        },
        enabled: Boolean(
          agent && activeSheet && (snapshot?.root.sheets.filter((s) => s.state === "visible").length ?? 0) > 1
        ),
      },
      "xlsx.conditional-format": {
        run: () => setConditionalFormatOpen(true),
        enabled: Boolean(agent && activeSheet),
      },
      "xlsx.conditional-format-clear": {
        run: () => onClearConditionalFormats(),
        enabled: Boolean(agent && activeSheet && (activeSheet.conditionalFormats.length ?? 0) > 0),
      },
      "xlsx.data-validation": {
        run: () => setDataValidationOpen(true),
        enabled: Boolean(agent && activeSheet),
      },
      "xlsx.data-validation-clear": {
        run: () => onClearDataValidations(),
        enabled: Boolean(
          agent &&
          activeSheet &&
          ((activeSheet.dataValidations.length ?? 0) > 0 || activeSheet.opaqueDataValidations)
        ),
      },
      "xlsx.name-manager": { run: () => setNameManagerOpen(true), enabled: Boolean(agent) },
      "xlsx.define-name": { run: () => setNameManagerOpen(true), enabled: Boolean(agent && selection) },
      "xlsx.merge": { run: () => onMerge(), enabled: canMerge },
      "xlsx.unmerge": { run: () => onUnmerge(), enabled: canUnmerge },
      "xlsx.insert-image": { run: () => onInsertImageClick() },
      "xlsx.format-as-table": {
        run: () => {
          if (!activeSheet || !selection) return;
          dispatchOrToast("xlsx:add-table", {
            sheet: activeSheet.name,
            range: formatSelection(selection),
          });
        },
        enabled: Boolean(activeSheet && selection),
      },
      "xlsx.insert-chart": {
        run: () => setInsertChartOpen(true),
        enabled: Boolean(activeSheet && selection),
      },
      "xlsx.insert-pivot-table": {
        run: () => setInsertPivotOpen(true),
        enabled: Boolean(activeSheet && selection),
      },
      "xlsx.edit-chart": {
        run: () => {
          if (selectedChartId !== null) setEditChartId(selectedChartId);
        },
        enabled: Boolean(activeSheet && selectedChartId),
      },
      "xlsx.add-comment": { run: () => focusCommentComposer(), enabled: Boolean(selection) },

      // ── Page Layout / Formulas / Review / View (UI parity) ──────────
      "xlsx.set-page-setup": { run: () => onOpenPageSetup("page"), enabled: Boolean(activeSheet) },
      "xlsx.set-page-margins": { run: () => onOpenPageSetup("margins"), enabled: Boolean(activeSheet) },
      "xlsx.set-print-options": { run: () => onOpenPageSetup("sheet"), enabled: Boolean(activeSheet) },
      "xlsx.set-print-area": { run: () => onPrintArea("set"), enabled: Boolean(activeSheet && selection) },
      "xlsx.set-print-titles": { run: () => onOpenPageSetup("sheet"), enabled: Boolean(activeSheet) },
      "xlsx.set-calc-mode": {
        run: () => onSetCalcMode(calcState.mode === "auto" ? "manual" : "auto"),
        enabled: Boolean(snapshot),
      },
      "xlsx.set-show-formulas": { run: () => onToggleShowFormulas(), enabled: Boolean(activeSheet) },
      "xlsx.set-sheet-protection": {
        run: () => setProtectSheetOpen(true),
        enabled: Boolean(activeSheet),
      },
      "xlsx.set-workbook-protection": {
        run: () => setProtectWorkbookOpen(true),
        enabled: Boolean(snapshot),
      },
      "xlsx.set-sheet-view": { run: () => setZoomDialogOpen(true), enabled: Boolean(activeSheet) },
    };
    return buildPaletteFromCatalogue(xlsxActions, runners, t);
  }, [
    t,
    agent,
    canMerge,
    canTextToColumns,
    canUnmerge,
    focusCommentComposer,
    onApplyBorderPreset,
    onInsertImageClick,
    onMerge,
    onTextToColumns,
    onToggleFilter,
    onUnmerge,
    selection,
    activateFormatPainter,
    onAddSheet,
    onRenameSheet,
    onDeleteSheet,
    onSetSheetState,
    snapshot,
    activeSheet,
    onClearConditionalFormats,
    onClearDataValidations,
    dispatchOrToast,
    onOpenPageSetup,
    onPrintArea,
    onSetCalcMode,
    calcState.mode,
    onToggleShowFormulas,
    selectedChartId,
  ]);

  const tabFallback = useStableTabId("xlsx");
  const realtimeRoomId = useMemo<string | null>(() => {
    if (!agent) return null;
    if (roomOverride === null) return null;
    if (typeof roomOverride === "string" && roomOverride.length > 0) {
      return `oai/xlsx/host/${roomOverride}`;
    }
    if (!tabFallback && !initialSource) return null;
    return roomIdForSource({
      product: "xlsx",
      src: initialSource?.url,
      tabFallback,
      explicitRoom: readExplicitRoomFromUrl(),
    });
  }, [agent, initialSource, tabFallback, roomOverride]);
  const realtimeRoom = useRealtimeRoom({
    roomId: realtimeRoomId,
    product: "xlsx",
    ...(presenceUser
      ? {
          identity: {
            id: presenceUser.id,
            name: presenceUser.name,
            ...(presenceUser.color ? { color: presenceUser.color } : {}),
          },
        }
      : {}),
  });
  useCommandBroadcast({
    agent,
    room: realtimeRoom.room,
  });

  // Publish XLSX selection (sheet + A1 range) on every change so
  // peers see "Quick Quokka is on Sheet1!B4:D7" in real time.
  const presenceCursor = useMemo(() => {
    if (!activeSheet || !selection) return null;
    const r = selectionToRange(selection);
    const a = `${colToLetter(r.start.col)}${r.start.row + 1}`;
    const b = `${colToLetter(r.end.col)}${r.end.row + 1}`;
    return {
      product: "xlsx" as const,
      sheetName: activeSheet.name,
      anchor: a,
      range: a === b ? a : `${a}:${b}`,
    };
  }, [activeSheet, selection]);
  usePublishPresence({ room: realtimeRoom.room, cursor: presenceCursor });

  const adapter = useMemo<ProductAdapter>(
    () => ({
      product: "xlsx",
      filename,
      saveState,
      comments: { openCount: openCommentCount, resolvedCount: 0 },
      selectionSummary:
        selectionAggregates && selectionAggregates.length > 0
          ? { aggregates: selectionAggregates }
          : { text: selectionText ?? "" },
      canOpen: true,
      hideOpen: hideLocalFileOpen,
      canSave: Boolean(agent),
      canExport: Boolean(agent),
      exportFormats: XLSX_EXPORT_FORMATS,
      onOpenFile: () => void onPickFile(),
      onSave: () => onSave(),
      onExport: (format, options) => onExport(format, options),
      canUndo: agent?.canUndo() ?? false,
      canRedo: agent?.canRedo() ?? false,
      onUndo: () => {
        const a = agentRef.current;
        if (a && a.canUndo()) a.undo();
      },
      onRedo: () => {
        const a = agentRef.current;
        if (a && a.canRedo()) a.redo();
      },
      onOpenShortcuts: () => shortcutsDialog.setOpen(true),
      paletteCommands,
      findAdapter,
      renderCommentsPanel,
      onAddComment: focusCommentComposer,
    }),
    [
      agent,
      filename,
      findAdapter,
      focusCommentComposer,
      hideLocalFileOpen,
      onExport,
      onPickFile,
      onSave,
      openCommentCount,
      paletteCommands,
      renderCommentsPanel,
      saveState,
      selectionAggregates,
      selectionText,
      shortcutsDialog,
    ]
  );

  return (
    <>
      <RemotePresenceList peers={realtimeRoom.remotePeers} />
      <EditorShell
        adapter={adapter}
        onBack={onCloseProp}
        topBarExtras={<PresenceSlot state={realtimeRoom} />}
        toolbar={
          snapshot ? (
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
              onAddComment={focusCommentComposer}
              onToggleFilter={onToggleFilter}
              filterActive={!!activeSheet?.autoFilter}
              onInsertImage={onInsertImageClick}
              onFreeze={onFreeze}
              freeze={activeSheet?.freeze}
              freezeAnchor={selection ? { row: selection.anchor.row, col: selection.anchor.col } : null}
              onApplyBorderPreset={onApplyBorderPreset}
              onOpenMoreBorders={() => setFormatCellsTab("border")}
              onActivateFormatPainter={activateFormatPainter}
              formatPainterActive={formatPainter !== null}
              onOpenInsertChart={() => setInsertChartOpen(true)}
              onOpenInsertPivot={() => setInsertPivotOpen(true)}
              selectedChartId={selectedChartId}
              onEditSelectedChart={() => {
                if (selectedChartId !== null) setEditChartId(selectedChartId);
              }}
              onOpenPageSetup={onOpenPageSetup}
              onApplyMarginsPreset={onApplyMarginsPreset}
              onApplyOrientation={onApplyOrientation}
              onApplyPaperSize={onApplyPaperSize}
              onPrintArea={onPrintArea}
              onTogglePrintFlag={onTogglePrintFlag}
              printGridLines={pageState.printGridLines}
              printHeadings={pageState.printHeadings}
              onOpenNameManager={() => setNameManagerOpen(true)}
              onSetCalcMode={onSetCalcMode}
              onSetCalcOnSave={onSetCalcOnSave}
              onToggleShowFormulas={onToggleShowFormulas}
              calcMode={calcState.mode}
              calcOnSave={calcState.calcOnSave}
              showFormulas={sheetView.showFormulas}
              onOpenProtectSheet={() => setProtectSheetOpen(true)}
              onOpenProtectWorkbook={() => setProtectWorkbookOpen(true)}
              sheetProtected={sheetProtected}
              workbookProtected={workbookProtected}
              onSetSheetView={onSetSheetView}
              onToggleViewFlag={onToggleViewFlag}
              onOpenZoom={() => setZoomDialogOpen(true)}
              viewMode={sheetView.view}
              showGridLinesView={sheetView.showGridLines}
              showRowColHeadersView={sheetView.showRowColHeaders}
              showRulerView={sheetView.showRuler}
              rightToLeft={sheetView.rightToLeft}
              zoomScale={sheetView.zoomScale}
            />
          ) : (
            // Placeholder before the workbook loads so opening the
            // first file doesn't push the grid down by the toolbar's
            // height. Visually inert (matches the surface background)
            // and aria-hidden because there's nothing to interact with.
            <div aria-hidden data-testid="xlsx-toolbar-placeholder" className="h-[34px] w-full" />
          )
        }
        statusBarLeft={
          <span className="flex items-center gap-3 text-[11px] tabular-nums text-tertiary">
            <span>
              rev {revision} · {pendingCount} pending
            </span>
            {formulaEditing ? (
              <span
                className="inline-flex items-center gap-1 rounded bg-[var(--accent-light)] px-1.5 py-0.5 text-[10px] font-medium text-[var(--accent)] shadow-sm"
                data-testid="xlsx-formula-ref-hint"
                aria-live="polite"
              >
                {/*
                  Surfaces what cell-grid clicks do while the formula
                  bar is open. Without this hint users routinely think
                  clicks deselect their formula draft; in fact they
                  insert the picked cell as a reference.
                */}
                {t("xlsx.selection.formulaPickHint")}
              </span>
            ) : null}
          </span>
        }
        body={
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
            className="relative flex min-h-0 min-w-0 flex-1 flex-col gap-2 p-3 outline-none"
          >
            <input
              ref={imageInputRef}
              type="file"
              accept="image/png,image/jpeg,image/gif"
              data-testid="insert-image-input"
              className="sr-only"
              onChange={onImageInputChange}
            />
            {!agent ? (
              initialLoadFailed ? (
                <EmptyState product="xlsx" onOpen={() => void onPickFile()} />
              ) : null
            ) : (
              <>
                <TextToColumnsPopover
                  open={ttocOpen}
                  defaultDelimiter={ttocDefaultDelim}
                  onCancel={() => setTtocOpen(false)}
                  onConfirm={onTextToColumnsConfirm}
                />

                {activeSheet?.autoFilter && filterDropdown ? (
                  <FilterDropdown
                    open
                    sheet={activeSheet}
                    styles={snapshot!.root.styles}
                    autoFilter={activeSheet.autoFilter}
                    colId={filterDropdown.colId}
                    anchor={filterDropdown.anchor}
                    onClose={onCloseFilter}
                    onSort={onSortFromFilter}
                    onClear={onClearFilterColumn}
                    onApply={onApplyFilterColumn}
                  />
                ) : null}

                {snapshot && formatCellsTab !== null ? (
                  <FormatCellsDialog
                    open
                    styles={snapshot.root.styles}
                    anchorStyleId={selectedCell?.styleId}
                    initialTab={formatCellsTab}
                    onClose={() => setFormatCellsTab(null)}
                    onApply={(patch) => onApplyFormat(patch)}
                  />
                ) : null}

                <PasteSpecialDialog
                  open={pasteSpecialOpen}
                  onClose={() => setPasteSpecialOpen(false)}
                  onConfirm={onPasteSpecialConfirm}
                />

                {activeSheet ? (
                  <ConditionalFormatDialog
                    open={conditionalFormatOpen}
                    onClose={() => setConditionalFormatOpen(false)}
                    defaultRange={selection ? formatSelection(selection) : "A1"}
                    rules={activeSheet.conditionalFormats}
                    onAddRule={onAddConditionalFormat}
                    onRemoveRule={onRemoveConditionalFormat}
                    onClearRules={onClearConditionalFormats}
                  />
                ) : null}

                {activeSheet ? (
                  <DataValidationDialog
                    open={dataValidationOpen}
                    onClose={() => setDataValidationOpen(false)}
                    defaultRange={selection ? formatSelection(selection) : "A1"}
                    rules={activeSheet.dataValidations}
                    hasOpaqueRules={Boolean(activeSheet.opaqueDataValidations)}
                    onAddRule={onAddDataValidation}
                    onRemoveRule={onRemoveDataValidation}
                    onClearRules={onClearDataValidations}
                  />
                ) : null}

                <NameManagerDialog
                  open={nameManagerOpen}
                  onClose={() => setNameManagerOpen(false)}
                  definedNames={definedNames}
                  sheetNames={(snapshot?.root.sheets ?? []).map((s) => s.name)}
                  defaultRefersTo={selectionRefersTo}
                  onAdd={onAddDefinedName}
                  onUpdate={onUpdateDefinedName}
                  onRemove={onRemoveDefinedName}
                />

                <PageSetupDialog
                  open={pageSetupOpen !== null}
                  initialTab={pageSetupOpen?.tab}
                  snapshot={snapshot}
                  sheetName={activeSheet?.name ?? null}
                  onClose={() => setPageSetupOpen(null)}
                  onSubmit={onSubmitPageSetup}
                />

                <ProtectSheetDialog
                  open={protectSheetOpen}
                  sheetName={activeSheet?.name ?? null}
                  currentlyProtected={sheetProtected}
                  onClose={() => setProtectSheetOpen(false)}
                  onProtect={onProtectSheetSubmit}
                  onUnprotect={onProtectSheetClear}
                />

                <ProtectWorkbookDialog
                  open={protectWorkbookOpen}
                  currentlyProtected={workbookProtected}
                  onClose={() => setProtectWorkbookOpen(false)}
                  onProtect={onProtectWorkbookSubmit}
                  onUnprotect={onProtectWorkbookClear}
                />

                <ZoomDialog
                  open={zoomDialogOpen}
                  currentZoom={sheetView.zoomScale}
                  onClose={() => setZoomDialogOpen(false)}
                  onSubmit={onSetZoom}
                />

                <div className="formula-bar relative flex items-center gap-2 rounded-md border border-divider bg-surface px-2 py-1.5">
                  <NameBox
                    selectionRef={selectedRef}
                    definedNames={definedNames}
                    activeSheet={activeSheet?.name}
                    onJump={onJumpFromNameBox}
                    onCreateName={onCreateNameFromBox}
                    onOpenManager={() => setNameManagerOpen(true)}
                    disabled={!agent}
                  />
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
                        // Pin the origin sheet on the *first* focus of an
                        // edit session — re-focus from a sheet-tab roundtrip
                        // must not re-pin to the new sheet.
                        if (formulaOriginSheetRef.current === null && activeSheet) {
                          setFormulaOriginSheet(activeSheet.name);
                        }
                        requestAnimationFrame(captureCaret);
                      }}
                      onBlur={() => {
                        // Clicking outside the formula bar (e.g. a toolbar
                        // button) cancels the in-flight edit; restore the
                        // origin sheet view so the user lands back where
                        // they started.
                        const origin = formulaOriginSheetRef.current;
                        setFormulaFocused(false);
                        setFormulaDraft("");
                        setFormulaOriginSheet(null);
                        pendingRefSpanRef.current = null;
                        pendingRefAnchorRef.current = null;
                        if (origin && activeSheetRef.current?.name !== origin) {
                          setActiveSheetName(origin);
                        }
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
                          const pick =
                            suggestionMatches[Math.min(suggestHighlight, suggestionMatches.length - 1)];
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
                          const origin = formulaOriginSheetRef.current;
                          setFormulaFocused(false);
                          setFormulaDraft("");
                          setFormulaOriginSheet(null);
                          pendingRefSpanRef.current = null;
                          pendingRefAnchorRef.current = null;
                          if (origin && activeSheetRef.current?.name !== origin) {
                            setActiveSheetName(origin);
                          }
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

                <div className="relative flex flex-1 min-h-0 min-w-0">
                  <div
                    className="relative flex-1 min-h-0 min-w-0"
                    data-format-painter={formatPainter ? "1" : undefined}
                    style={
                      formatPainter
                        ? {
                            cursor:
                              "url('data:image/svg+xml;utf8,<svg xmlns=%22http://www.w3.org/2000/svg%22 width=%2220%22 height=%2220%22 viewBox=%220 0 24 24%22 fill=%22none%22 stroke=%22currentColor%22 stroke-width=%222%22 stroke-linecap=%22round%22 stroke-linejoin=%22round%22><path d=%22M14.622 17.897l-10.68-2.913%22/><path d=%22M18.376 2.622a1 1 0 1 1 3.002 3.002L17.36 9.643a.5.5 0 0 0 0 .707l.944.944a2.41 2.41 0 0 1 0 3.408l-.944.944a.5.5 0 0 1-.707 0L8.354 7.348a.5.5 0 0 1 0-.707l.944-.944a2.41 2.41 0 0 1 3.408 0l.944.944a.5.5 0 0 0 .707 0z%22/><path d=%22M9 8c-1.804 2.71-3.97 3.46-6.583 3.948a.507.507 0 0 0-.302.819l7.32 8.883a1 1 0 0 0 1.185.204C12.735 20.405 16 16.792 16 15%22/></svg>') 4 18, crosshair",
                          }
                        : undefined
                    }
                  >
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
                        commentMarkers={commentMarkers}
                        scrollTarget={commentScrollTarget}
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
                        onOpenFilter={onOpenFilter}
                        imageObjectUrls={imageObjectUrls}
                        selectedImageId={selectedImageId}
                        onSelectImage={(id) => {
                          setSelectedImageId(id);
                          if (id !== null) surfaceRef.current?.focus({ preventScroll: true });
                        }}
                        onMoveImage={onMoveImage}
                        onResizeImage={onResizeImage}
                        onImageContextMenu={(imageId, coords) => {
                          setSelectedImageId(imageId);
                          setCtxMenu({ target: { kind: "image", imageId }, x: coords.x, y: coords.y });
                        }}
                        liveEditDraft={
                          formulaFocused && selection && isSingle(selection)
                            ? {
                                row: selection.anchor.row,
                                col: selection.anchor.col,
                                draft: formulaDraft,
                              }
                            : null
                        }
                        cfOverlays={cfOverlays}
                        dvIndex={dvIndex}
                        extraAreas={extraAreas}
                        selectedChartId={selectedChartId}
                        onSelectChart={(id) => {
                          setSelectedChartId(id);
                          if (id !== null) surfaceRef.current?.focus({ preventScroll: true });
                        }}
                        onRemoveChart={(id) => {
                          if (!activeSheet) return;
                          dispatchOrToast("xlsx:remove-chart", {
                            sheet: activeSheet.name,
                            chartId: id,
                          });
                          setSelectedChartId(null);
                        }}
                        onMoveChart={onMoveChart}
                        onResizeChart={onResizeChart}
                        onChangeChartKind={onChangeChartKind}
                        onRequestEditChart={(id) => setEditChartId(id)}
                        remotePeers={realtimeRoom.remotePeers}
                        pivotsForSheet={pivotsForActiveSheet}
                        pivotBadgeTooltip={t("xlsx.pivot.readOnlyTooltip")}
                      />
                    ) : null}
                  </div>
                </div>

                <SheetTabBar
                  sheets={sheets.map((s) => ({ id: String(s.id), name: s.name, state: s.state }))}
                  activeName={activeSheetName}
                  peers={realtimeRoom.remotePeers
                    .map((p) => {
                      const c = p.state.cursor;
                      if (!c || c.product !== "xlsx") return null;
                      return {
                        clientId: p.clientId,
                        sheetName: c.sheetName,
                        name: p.state.user.name,
                        color: p.state.user.color,
                      };
                    })
                    .filter((x): x is NonNullable<typeof x> => x !== null)}
                  onActivate={(name) => {
                    // If the user clicks a hidden sheet's "unhide" chip the
                    // sheet may not yet be visible — flip it to visible first
                    // so the activation lands on a renderable surface.
                    const s = sheets.find((x) => x.name === name);
                    if (s && s.state !== "visible") {
                      void onSetSheetState(name, "visible");
                    }
                    setActiveSheetName(name);
                    setExtraAreas([]);
                    // While the formula bar is in point mode, switching
                    // tabs should not steal focus or commit — keep the
                    // bar focused so the next cell click inserts a
                    // sheet-qualified ref. The pending ref span is reset
                    // so a click on the new sheet starts a fresh ref
                    // (clicks across sheets cannot extend a range).
                    if (formulaEditing) {
                      pendingRefSpanRef.current = null;
                      pendingRefAnchorRef.current = null;
                      requestAnimationFrame(() => {
                        formulaInputRef.current?.focus();
                      });
                    }
                  }}
                  onRename={(currentName, nextName) => onRenameSheet(currentName, nextName)}
                  onDelete={(name) => onDeleteSheet(name)}
                  onMove={(name, to) => onMoveSheet(name, to)}
                  onAdd={() => onAddSheet()}
                  onSetState={(name, state) => onSetSheetState(name, state)}
                />

                <ContextMenu
                  open={ctxMenu !== null}
                  x={ctxMenu?.x ?? 0}
                  y={ctxMenu?.y ?? 0}
                  items={ctxMenuItems}
                  onClose={closeCtxMenu}
                />
              </>
            )}
          </div>
        }
        toasts={toasts}
        onDismissToast={dismissToast}
        onFileDrop={onShellFileDrop}
        dropExtension=".xlsx"
        onRenameFilename={(next) => setFilename(next)}
      />
      <KeyboardShortcutsDialog
        product="xlsx"
        open={shortcutsDialog.open}
        onClose={() => shortcutsDialog.setOpen(false)}
      />
      <InsertPivotTableDialog
        open={insertPivotOpen}
        defaultSourceRange={selection ? formatSelection(selection) : "A1:C10"}
        defaultDestination={pivotDestinationDefault(selection)}
        headerLabels={pivotHeaderLabelsFromSelection(activeSheet, selection)}
        onCancel={() => setInsertPivotOpen(false)}
        onSubmit={(args) => {
          if (!activeSheet) return;
          const result = computePivotSummary(activeSheet, args);
          if (!result) {
            pushToast("error", "Couldn't read the source range — check the cells exist.");
            setInsertPivotOpen(false);
            return;
          }
          const dest = parseA1(args.destinationTopLeftA1);
          const range = formatRange({
            start: { row: dest.row, col: dest.col },
            end: { row: dest.row + result.rows.length - 1, col: dest.col + 1 },
          });
          dispatchOrToast("xlsx:set-range-values", {
            sheet: activeSheet.name,
            range,
            values: result.rows,
          });
          setInsertPivotOpen(false);
          pushToast(
            "info",
            `Inserted pivot summary at ${args.destinationTopLeftA1} (${result.rows.length - 1} rows)`
          );
        }}
      />
      <InsertChartDialog
        open={insertChartOpen}
        defaultRange={selection ? formatSelection(selection) : "A1:B5"}
        onCancel={() => setInsertChartOpen(false)}
        onSubmit={(args) => {
          if (!activeSheet) return;
          dispatchOrToast("xlsx:add-chart", {
            sheet: activeSheet.name,
            kind: args.kind,
            dataRange: args.dataRange,
            hasHeaderRow: args.hasHeaderRow,
            hasCategoryColumn: args.hasCategoryColumn,
            ...(args.title ? { title: args.title } : {}),
            palette: args.palette,
            showLegend: args.showLegend,
            showDataLabels: args.showDataLabels,
            showGridlines: args.showGridlines,
            ...(args.xAxisTitle ? { xAxisTitle: args.xAxisTitle } : {}),
            ...(args.yAxisTitle ? { yAxisTitle: args.yAxisTitle } : {}),
          });
          setInsertChartOpen(false);
        }}
      />
      <ChartDialog
        open={editChartId !== null && !!editingChart}
        mode="edit"
        initial={editingChart ?? undefined}
        onCancel={() => setEditChartId(null)}
        onSubmit={(args) => {
          if (!activeSheet || !editingChart) return;
          dispatchOrToast("xlsx:update-chart", {
            sheet: activeSheet.name,
            chartId: editingChart.id,
            kind: args.kind,
            dataRange: args.dataRange,
            hasHeaderRow: args.hasHeaderRow,
            hasCategoryColumn: args.hasCategoryColumn,
            title: args.title ?? null,
            palette: args.palette,
            showLegend: args.showLegend,
            showDataLabels: args.showDataLabels,
            showGridlines: args.showGridlines,
            xAxisTitle: args.xAxisTitle ?? null,
            yAxisTitle: args.yAxisTitle ?? null,
          });
          setEditChartId(null);
        }}
      />
    </>
  );
}

/**
 * Best-effort literal parsing for the formula bar / in-cell editor:
 *   - `""`        → null (clear cell)
 *   - `"123.45"`  → number
 *   - `"true"`    → boolean
 *   - everything else stays a string
 */
/**
 * C11 — Resolve a `formula` data-validation source like
 * `=Sheet1!$A$1:$A$5` (or `=$A$1:$A$5` for same-sheet) into its
 * stringified cell values. Cross-sheet refs are best-effort: if the
 * referenced sheet isn't in scope here we fall back to an empty
 * options list (the dropdown still appears so the user can edit).
 */
function resolveFormulaListOptions(source: string, sheet: Sheet): string[] {
  const trimmed = source.replace(/^=/, "").trim();
  // Strip optional "SheetName!" prefix; we only resolve same-sheet refs.
  const bang = trimmed.indexOf("!");
  const ref = bang >= 0 ? trimmed.slice(bang + 1) : trimmed;
  // Drop $ anchors before parsing.
  const cleaned = ref.replace(/\$/g, "");
  let range;
  try {
    range = parseRange(cleaned);
  } catch {
    return [];
  }
  const out: string[] = [];
  for (let r = range.start.row; r <= range.end.row; r++) {
    for (let c = range.start.col; c <= range.end.col; c++) {
      const cell = sheet.cells.get(cellKey(r, c));
      if (!cell || cell.value === null || cell.value === undefined) continue;
      out.push(stringifyCellValue(cell.value));
    }
  }
  return out;
}

function stringifyCellValue(value: CellValue): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number") return String(value);
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "object" && "kind" in value && value.kind === "error") return value.code;
  return String(value);
}

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
 * Excel date serial for the *local* calendar date in `d`. Excel's
 * epoch is 1899-12-30 because the file format treats 1900 as a leap
 * year (Lotus 1-2-3 bug); using Dec 30 1899 as the origin makes the
 * arithmetic work correctly for every modern date without us having
 * to special-case the phantom 2/29/1900. We round to the day in the
 * user's local timezone so "today" matches what their calendar
 * shows, regardless of UTC offset.
 */
function excelDateSerial(d: Date): number {
  const epoch = Date.UTC(1899, 11, 30);
  const dayMs = 24 * 60 * 60 * 1000;
  const local = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
  return Math.floor((local - epoch) / dayMs);
}

/**
 * Excel time-of-day serial — the fractional-day component (0 ≤ t < 1)
 * that pairs with built-in numFmtIds 18 / 19 / 20 to render as
 * "11:30 AM" etc. Computed against `d`'s local clock so the chord
 * inserts the wall-clock time the user sees on screen.
 */
function excelTimeSerial(d: Date): number {
  const seconds = d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds();
  return seconds / 86400;
}

// ─── Quick-pivot helpers ──────────────────────────────────────────────
//
// Phase-1 "values-only pivot" (see `InsertPivotTableDialog.tsx` for the
// scope rationale). We don't write OOXML pivot parts — instead the
// dialog hands us a (sourceRange, group-col, value-col, agg) tuple and
// we compute the result client-side, returning a 2-column matrix that
// the caller writes back via `xlsx:set-range-values`.

function pivotDestinationDefault(selection: Selection | null): string {
  if (!selection) return "E1";
  const n = normalizeSelection(selection);
  return formatA1({ row: n.r0, col: n.c1 + 2 });
}

function pivotHeaderLabelsFromSelection(
  sheet: Sheet | null,
  selection: Selection | null
): ReadonlyArray<string> {
  if (!sheet || !selection) return [];
  const n = normalizeSelection(selection);
  const labels: string[] = [];
  for (let c = n.c0; c <= n.c1; c++) {
    const cell = sheet.cells.get(cellKey(n.r0, c));
    labels.push(cellTextValue(cell));
  }
  return labels;
}

interface PivotComputation {
  readonly rows: ReadonlyArray<ReadonlyArray<string | number>>;
}

function computePivotSummary(sheet: Sheet, args: PivotDialogSubmit): PivotComputation | null {
  let range;
  try {
    range = parseRange(args.sourceRange);
  } catch {
    return null;
  }
  const r0 = range.start.row;
  const r1 = range.end.row;
  const c0 = range.start.col;
  const c1 = range.end.col;
  const colCount = c1 - c0 + 1;
  if (args.groupColumnIndex < 0 || args.groupColumnIndex >= colCount) return null;
  if (args.valueColumnIndex < 0 || args.valueColumnIndex >= colCount) return null;

  const dataStart = args.hasHeaderRow ? r0 + 1 : r0;
  const groupCol = c0 + args.groupColumnIndex;
  const valueCol = c0 + args.valueColumnIndex;

  const buckets = new Map<string, number[]>();
  // Preserve first-seen insertion order so the summary lines up with
  // the data the user is looking at; Map already does that for us.
  for (let r = dataStart; r <= r1; r++) {
    const groupCell = sheet.cells.get(cellKey(r, groupCol));
    const valueCell = sheet.cells.get(cellKey(r, valueCol));
    const groupKey = cellTextValue(groupCell);
    if (groupKey.length === 0) continue;
    const num = cellNumericValue(valueCell);
    const bucket = buckets.get(groupKey);
    if (bucket) {
      bucket.push(num);
    } else {
      buckets.set(groupKey, [num]);
    }
  }

  const headerLabels = pivotHeaderLabelsFromSelection(sheet, {
    anchor: { row: r0, col: c0 },
    focus: { row: r1, col: c1 },
  });
  const groupHeader =
    args.hasHeaderRow && headerLabels[args.groupColumnIndex]?.length
      ? headerLabels[args.groupColumnIndex]!
      : `Column ${args.groupColumnIndex + 1}`;
  const valueHeader =
    args.hasHeaderRow && headerLabels[args.valueColumnIndex]?.length
      ? headerLabels[args.valueColumnIndex]!
      : `Column ${args.valueColumnIndex + 1}`;
  const aggLabel =
    args.aggregation === "count"
      ? `Count of ${valueHeader}`
      : `${args.aggregation[0]!.toUpperCase()}${args.aggregation.slice(1)} of ${valueHeader}`;

  const rows: Array<Array<string | number>> = [[groupHeader, aggLabel]];
  for (const [key, values] of buckets) {
    rows.push([key, aggregate(values, args.aggregation)]);
  }
  return { rows };
}

function aggregate(values: ReadonlyArray<number>, kind: PivotDialogSubmit["aggregation"]): number {
  if (values.length === 0) return 0;
  switch (kind) {
    case "sum":
      return values.reduce((a, b) => a + b, 0);
    case "count":
      return values.length;
    case "average":
      return values.reduce((a, b) => a + b, 0) / values.length;
    case "min":
      return values.reduce((a, b) => Math.min(a, b), Number.POSITIVE_INFINITY);
    case "max":
      return values.reduce((a, b) => Math.max(a, b), Number.NEGATIVE_INFINITY);
  }
}

function cellTextValue(cell: Cell | undefined): string {
  if (!cell) return "";
  const v = cell.value;
  if (v === undefined || v === null) return "";
  return typeof v === "string" ? v : String(v);
}

function cellNumericValue(cell: Cell | undefined): number {
  if (!cell) return 0;
  const v = cell.value;
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    const n = Number(v.replace(/,/g, ""));
    return Number.isFinite(n) ? n : 0;
  }
  return 0;
}
