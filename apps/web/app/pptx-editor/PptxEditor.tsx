"use client";

import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { CommentComposer, CommentsSidebar } from "@officeai/ui";
import { I18nProvider, useTranslator, type Locale } from "@/lib/i18n";
import { rulerUnitForLocale } from "@/lib/ruler/units";
import { createPptxCommentsProvider } from "./pptxCommentsProvider";
import { PptxAgent } from "@officeai/pptx/agent";
import {
  SlideCanvas,
  SlidesSidebar,
  type PptxTextSelection,
  type RemoteSelectionPeer,
  type SlideContextAction,
  type SlideRailPeerDot,
} from "@officeai/pptx/renderer/react";
import { MAX_ZOOM, MIN_ZOOM, clampZoom } from "@officeai/pptx/renderer";
import {
  createPlayback,
  readFillSpec,
  resolveConnectorEndpoint,
  type ConnectorShape,
  type FillSpec,
  type LayoutKindPayload,
  type OpaqueXml,
  type Picture,
  type PlaybackController,
  type PptxSnapshot,
  type Shape,
  type ShapePreset,
  type Slide,
  type TextAnchor,
  type TextShape,
} from "@officeai/pptx";
import { buildBlankPptx, buildSamplePptx } from "@/lib/sample-pptx";
import { PptxToolbar } from "./PptxToolbar";
import { SlideSizeDialog, type SlideSizePreset } from "./SlideSizeDialog";
import { SetUpShowDialog, type SetUpShowValues } from "./SetUpShowDialog";
import { ConnectorContextBar, type ConnectorAction, type ConnectorStylePatch } from "./ConnectorContextBar";
import { ShapeGeometryContextBar, shapeHasAdjustableGeometry } from "./ShapeGeometryContextBar";
import { PresentMode } from "./PresentMode";
import { AnimationsPanel } from "./AnimationsPanel";
import { MasterPanel } from "./MasterPanel";
import { computePptxActive, createPptxFormatProvider } from "./pptxFormatProvider";
import { useShortcutsDialog } from "@/lib/shortcuts/useShortcutsDialog";
import { handleUndoRedo, isFormField } from "@/lib/undo-redo";
import { KeyboardShortcutsDialog } from "@/lib/shortcuts/KeyboardShortcutsDialog";
import { usePptxShortcuts } from "@/lib/pptx-shortcuts";
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
import {
  EditorShell,
  EmptyState,
  ZoomControl,
  buildPaletteFromCatalogue,
  createToastId,
  type ExportFormat,
  type ExportOptionValues,
  type PaletteCommand,
  type PaletteRunners,
  type ProductAdapter,
  type RightRailTab,
  type SaveState,
  type ToastItem,
} from "@/lib/shell";
import { pptxActions } from "@officeai/pptx";
import { PRODUCT_FILE_TYPES, downloadBlob, openFile, saveFile } from "@/lib/files/file-service";
import { convertViaServer } from "@/lib/files/convert-client";
import {
  parseSlideRange,
  snapshotToPngZip,
  snapshotToSlideJpeg,
  snapshotToSlidePng,
  snapshotToSlideSvg,
  snapshotToSvgZip,
} from "./lib/export-images";
import {
  EMBED_MIME,
  isEmbedEnabled,
  makeEnvelope,
  parseEnvelope,
  serializeEnvelope,
} from "@/lib/embed/envelope";
import { applyXlsxRangeToPptx } from "@/lib/embed/applyXlsxRangeToPptx";
import { applyXlsxEmbed } from "@/lib/embed/xlsxEmbedShared";
import type { XlsxEmbedMode } from "@/lib/embed/xlsxEmbedShared";
import { XlsxRangePickerDialog, type XlsxRangePickerResult } from "@/lib/embed/XlsxRangePickerDialog";
import { installAltKeyTracker, isAltKeyPressed } from "@/lib/embed/altKeyTracker";
import { EmbeddedXlsxModal } from "@/lib/embed/EmbeddedXlsxModal";
import { resolveEmbeddedXlsxRef, readEmbeddedXlsxBytes } from "@/lib/embed/getEmbeddedXlsxBytes";

const SCALE_OPTIONS = {
  type: "select" as const,
  defaultId: "2",
  options: [
    { id: "1", label: "1× (standard)" },
    { id: "2", label: "2× (retina)" },
    { id: "3", label: "3× (high-DPI)" },
  ],
};

const PPTX_EXPORT_FORMATS: ReadonlyArray<ExportFormat> = [
  // ── Whole deck ─────────────────────────────────────────────────
  // Deck-level exports lead the dropdown because "download the
  // .pptx / PDF the deck" is overwhelmingly the common ask. We
  // collapse what would otherwise be Native / PDF & web / Images
  // into a single bucket: in PPTX each holds only 1–2 entries, and
  // rendering them as separate sections reads as a chopped-up
  // dropdown. Order: native (round-trip-true) first, then PDF and
  // HTML, then the bundles (PNG before SVG to mirror the per-slide
  // ordering further down).
  {
    id: "pptx",
    label: "PowerPoint presentation (.pptx)",
    description: "Round-trip native OOXML — opens in PowerPoint, Keynote and LibreOffice Impress.",
    extension: "pptx",
    mime: PRODUCT_FILE_TYPES.pptx.primaryMime,
    kind: "instant",
    group: "deck",
    icon: "slides",
  },
  {
    id: "pdf",
    label: "PDF document (.pdf)",
    description: "Server-side conversion via LibreOffice. One slide per PDF page.",
    extension: "pdf",
    mime: "application/pdf",
    kind: "dialog",
    group: "deck",
    icon: "pdf",
    optionFields: [
      {
        id: "slideRange",
        label: "Slide range",
        control: { type: "text", placeholder: "All slides — try 1,3-5" },
        hint: "Leave blank for every slide. Examples: 1,3 — 2-5 — 1,4-7,10.",
      },
      {
        id: "pageSize",
        label: "Page size",
        control: {
          type: "select",
          defaultId: "source",
          options: [
            { id: "source", label: "Use slide size" },
            { id: "a4", label: "A4 landscape" },
            { id: "letter", label: "Letter landscape" },
          ],
        },
      },
    ],
  },
  {
    id: "html",
    label: "Web page (.html)",
    description: "Server-side HTML export. Ships an interactive viewer with each slide.",
    extension: "html",
    mime: "text/html",
    kind: "instant",
    group: "deck",
    icon: "code",
  },
  {
    id: "png-zip",
    label: "All slides — PNG (.zip)",
    description: "One PNG per slide, bundled as a zip. Pick a scale for retina exports.",
    extension: "zip",
    mime: "application/zip",
    kind: "dialog",
    group: "deck",
    icon: "image",
    optionFields: [
      { id: "scale", label: "Resolution", control: SCALE_OPTIONS },
      {
        id: "slideRange",
        label: "Slide range",
        control: { type: "text", placeholder: "All slides — try 1,3-5" },
        hint: "Leave blank for every slide.",
      },
    ],
  },
  {
    id: "svg-zip",
    label: "All slides — SVG (.zip)",
    description: "Resolution-independent SVG per slide.",
    extension: "zip",
    mime: "application/zip",
    kind: "instant",
    group: "deck",
    icon: "image",
  },
  // ── This slide ─────────────────────────────────────────────────
  // Special-case bucket at the bottom: "I just want this one slide
  // out". Order within the group is most-shareable → most-
  // specialised: PDF (universal single-page handout), PNG (default
  // raster for chat / docs), JPEG (smaller raster), SVG (vector
  // for designers).
  {
    id: "slide-pdf",
    label: "Current slide — PDF (.pdf)",
    description: "Single-page PDF of the current slide. Server-side via LibreOffice.",
    extension: "pdf",
    mime: "application/pdf",
    kind: "instant",
    group: "current",
    icon: "pdf",
  },
  {
    id: "slide-png",
    label: "Current slide — PNG (.png)",
    description: "Just the slide that's currently in view. Lossless, with a white matte.",
    extension: "png",
    mime: "image/png",
    kind: "dialog",
    group: "current",
    icon: "image",
    optionFields: [{ id: "scale", label: "Resolution", control: SCALE_OPTIONS }],
  },
  {
    id: "slide-jpeg",
    label: "Current slide — JPEG (.jpg)",
    description: "Smaller, share-friendly version of the current slide.",
    extension: "jpg",
    mime: "image/jpeg",
    kind: "dialog",
    group: "current",
    icon: "image",
    optionFields: [
      { id: "scale", label: "Resolution", control: SCALE_OPTIONS },
      {
        id: "quality",
        label: "Quality",
        control: {
          type: "select",
          defaultId: "high",
          options: [
            { id: "low", label: "Low (~60%)" },
            { id: "medium", label: "Medium (~75%)" },
            { id: "high", label: "High (~92%)" },
            { id: "max", label: "Maximum (~98%)" },
          ],
        },
      },
    ],
  },
  {
    id: "slide-svg",
    label: "Current slide — SVG (.svg)",
    description: "Resolution-independent vector of the current slide.",
    extension: "svg",
    mime: "image/svg+xml",
    kind: "instant",
    group: "current",
    icon: "image",
  },
];

function stripPptxExtension(name: string): string {
  return name.replace(/\.pptx$/i, "");
}

/** Coerce a dialog "scale" option (`"1" | "2" | "3"`) into the
 * narrowed numeric type the renderer accepts. Defaults to 2× to
 * match the export-zip default. */
function parseScale(raw: ExportOptionValues[string] | undefined): 1 | 2 | 3 {
  const s = typeof raw === "string" ? raw : "2";
  return s === "1" ? 1 : s === "3" ? 3 : 2;
}

/** Map the dialog's coarse quality picker to the 0-1 number canvas
 * wants. Mirrors the lossless-vs-share trade-off most users actually
 * care about; the in-between values come from PhotoShop's "Save for
 * web" presets. */
function parseJpegQuality(raw: ExportOptionValues[string] | undefined): number {
  switch (typeof raw === "string" ? raw : "high") {
    case "low":
      return 0.6;
    case "medium":
      return 0.75;
    case "max":
      return 0.98;
    case "high":
    default:
      return 0.92;
  }
}

const SUPPORTED_IMAGE_MIME: ReadonlySet<string> = new Set([
  "image/png",
  "image/jpeg",
  "image/jpg",
  "image/gif",
  "image/bmp",
  "image/webp",
  "image/svg+xml",
]);

// MIME allowlists kept in sync with `pptx:insert-media`'s VIDEO_MIME /
// AUDIO_MIME tables (packages/pptx/src/commands/insert-media.ts). The
// command would reject unsupported types anyway; gating client-side
// gives us a friendlier toast than the raw "invalid-payload" error.
const SUPPORTED_VIDEO_MIME: ReadonlySet<string> = new Set(["video/mp4", "video/webm", "video/quicktime"]);
const SUPPORTED_AUDIO_MIME: ReadonlySet<string> = new Set([
  "audio/mpeg",
  "audio/mp4",
  "audio/wav",
  "audio/ogg",
]);

export interface PptxEditorProps {
  /** Fired whenever the editor's bootstrap-ready state changes. The
   * page-level splash listens to this to know when to fade out and
   * unveil the editor. Stays `false` until the agent is mounted and
   * the first deck has been parsed, then `true`. */
  readonly onBootstrapReady?: (ready: boolean) => void;
  /** Optional pre-loaded deck. When provided, the editor fetches the
   * bytes at `url` instead of building the synthetic welcome sample,
   * and uses `name` as the deck title (so subsequent Save / Export
   * keep the original filename). Used by the home page's "sample
   * files" listing. */
  readonly initialSource?: { readonly url: string; readonly name: string };
  /** When true, the editor bootstraps with an empty deck (one blank
   * slide, no title/subtitle) instead of the synthetic welcome
   * sample. Used by the home page's "New presentation" action.
   * Ignored when `initialSource` is set. */
  readonly initialBlank?: boolean;
  /** Optional pre-loaded deck bytes. When set, takes priority over
   * `initialSource` and `initialBlank` so embedding hosts can
   * stream a `Uint8Array` straight into the editor without first
   * stashing it under a URL. */
  readonly initialBytes?: Uint8Array;
  /** Filename to display + use on Save when `initialBytes` is set. */
  readonly initialFilename?: string;
  /** Host save handler. */
  readonly onSave?: (bytes: Uint8Array, mime: string, filename: string) => Promise<void>;
  /** Host close handler — embedding route owns navigation. */
  readonly onClose?: () => void;
  /** Override the i18n locale; mounts a self-contained
   * `<I18nProvider initialLocale={locale}>`. */
  readonly locale?: Locale;
  /** Theme override placeholder; wired in Phase 1. */
  readonly theme?: "light" | "dark";
  /** Realtime presence identity (host-supplied). When set, replaces
   * the default anonymous identity on the awareness payload so slide
   * cursors / avatars show the authenticated user's real name. */
  readonly presenceUser?: { readonly id: string; readonly name: string; readonly color?: string };
  /** Explicit realtime room id (host-supplied). Pin two browsers
   * viewing the same deck into the same room without coordinating
   * URLs. Pass `null` to disable realtime. */
  readonly room?: string | null;
  /** Hide the 📁 Open toolbar affordance. Set by embedded hosts that
   * own their document corpus — see
   * `EmbeddedEditorProps.hideLocalFileOpen` in
   * `@officeai/react-editors/contract`. */
  readonly hideLocalFileOpen?: boolean;
}

export function PptxEditor(props: PptxEditorProps = {}): React.ReactNode {
  const { locale } = props;
  if (locale !== undefined) {
    return (
      <I18nProvider initialLocale={locale}>
        <PptxEditorInner {...props} />
      </I18nProvider>
    );
  }
  return <PptxEditorInner {...props} />;
}

function PptxEditorInner({
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
}: PptxEditorProps = {}): React.ReactNode {
  const { t } = useTranslator();
  const [agent, setAgent] = useState<PptxAgent | null>(null);
  const agentRef = useRef<PptxAgent | null>(null);
  const [ready, setReady] = useState(false);

  // Mirror `ready` up to the page-level splash so it can fade out
  // and unveil the deck. Owning the splash at page scope (not here)
  // keeps the badge `<span>` mounted across the dynamic-import
  // handoff — see `apps/web/app/pptx-editor/page.tsx`.
  useEffect(() => {
    onBootstrapReady?.(ready);
  }, [ready, onBootstrapReady]);
  const [docName, setDocName] = useState(
    initialFilename ??
      initialSource?.name ??
      (initialBlank || initialBytes ? "Untitled.pptx" : "welcome.pptx")
  );
  const [activeIndex, setActiveIndex] = useState(0);
  const [tick, setTick] = useState(0);
  const [toasts, setToasts] = useState<ToastItem[]>([]);
  const [zoom, setZoom] = useState(1);
  const [presenting, setPresenting] = useState(false);
  // Speaker-notes panel is hidden by default — matches PowerPoint's
  // "View ▸ Notes" opt-in. Users toggle it via the trailing toolbar
  // button; we deliberately don't persist this across sessions because
  // the panel takes meaningful vertical space and the default state is
  // what users expect when reopening a deck.
  const [notesOpen, setNotesOpen] = useState(false);
  // PowerPoint-style View-tab preferences. Rulers are on by default
  // (matches PowerPoint), gridlines are off. Both are persisted to
  // localStorage so reopening the editor restores the user's choice
  // — these are pure UI prefs (not part of the deck), so we don't
  // route them through the command bus.
  const [rulersVisible, setRulersVisible] = useState<boolean>(() => {
    if (typeof window === "undefined") return true;
    const raw = window.localStorage.getItem("pptx:view:rulers");
    return raw === null ? true : raw === "1";
  });
  const [gridVisible, setGridVisible] = useState<boolean>(() => {
    if (typeof window === "undefined") return false;
    return window.localStorage.getItem("pptx:view:grid") === "1";
  });
  useEffect(() => {
    if (typeof window === "undefined") return;
    window.localStorage.setItem("pptx:view:rulers", rulersVisible ? "1" : "0");
  }, [rulersVisible]);
  useEffect(() => {
    if (typeof window === "undefined") return;
    window.localStorage.setItem("pptx:view:grid", gridVisible ? "1" : "0");
  }, [gridVisible]);
  // Locale → unit (cm in metric locales, in elsewhere). Memoised so
  // the SlideCanvas prop reference stays stable across re-renders.
  const rulerUnit = useMemo<"in" | "cm">(
    () => rulerUnitForLocale(typeof navigator !== "undefined" ? navigator.language : "en-US").unit,
    []
  );
  const [xlsxPickerOpen, setXlsxPickerOpen] = useState<XlsxEmbedMode | null>(null);
  // Dialog visibility for the Design / Slideshow ribbon entries.
  const [slideSizeDialogOpen, setSlideSizeDialogOpen] = useState(false);
  const [setUpShowDialogOpen, setSetUpShowDialogOpen] = useState(false);
  // Imperative request to open the shared right rail to a specific
  // tab (consumed by `EditorShell` via `requestRailTab`). Bumping the
  // nonce re-fires the effect even when the same tab is requested
  // twice in a row, e.g. running "Add shape animation" from Cmd+K
  // after the user closed the panel.
  const [railRequest, setRailRequest] = useState<{ tab: RightRailTab; nonce: number } | null>(null);
  const requestRail = useCallback((tab: RightRailTab) => {
    setRailRequest({ tab, nonce: Date.now() });
  }, []);
  /**
   * Active "Edit Data" modal state for double-click on a chart or
   * OLE spreadsheet shape. `null` when the modal is closed; otherwise
   * carries the loaded workbook bytes plus enough context to dispatch
   * the right `*:update-spreadsheet` (and optional `*:set-chart-data`)
   * command on save.
   */
  const [editingEmbed, setEditingEmbed] = useState<{
    readonly bytes: Uint8Array;
    readonly embeddingPartPath: string;
    readonly chartPartPath: string | null;
    readonly slideIndex: number;
    readonly shapeId: string;
    readonly title: string;
  } | null>(null);
  const [selectedShapeIds, setSelectedShapeIds] = useState<ReadonlyArray<string>>([]);
  const selectedShapeId = selectedShapeIds[0] ?? null;
  const [textSelection, setTextSelection] = useState<PptxTextSelection | null>(null);
  // Connector tool mode: when set, the canvas surfaces ports on every
  // hovered shape and a press-drag gesture commits a brand-new
  // connector of the chosen type. PowerPoint / Google Slides model the
  // connector picker as a "tool" rather than an instant insert because
  // the user almost always wants to place the line themselves; the
  // toolbar item toggles this state, and `null` is the default
  // selection-mode behaviour. Cleared on commit, on Esc, and when the
  // user re-clicks the same toolbar item.
  const [connectorTool, setConnectorTool] = useState<{
    type: "straight" | "elbow" | "curved";
  } | null>(null);
  const [saveState, setSaveState] = useState<SaveState>("saved");
  const [fileHandle, setFileHandle] = useState<FileSystemFileHandle | undefined>(undefined);
  // Mutable mirrors used by the format provider so it can read live
  // state without triggering rerenders inside its own callbacks.
  const slideIndexRef = useRef(0);
  const textSelectionRef = useRef<PptxTextSelection | null>(null);
  const selectedShapeIdRef = useRef<string | null>(null);
  const pushToastRef = useRef<((kind: ToastItem["kind"], text: string) => void) | null>(null);
  const slideSurfaceRef = useRef<HTMLElement | null>(null);
  // Bumped each time the comments sidebar requests "scroll / focus on
  // this comment". The canvas keys its flash overlay off `nonce` so
  // re-clicking the same comment re-plays the animation.
  const [commentFlashTarget, setCommentFlashTarget] = useState<
    | { kind: "shape"; shapeId: string; nonce: number }
    | { kind: "pin"; xEmu: number; yEmu: number; nonce: number }
    | null
  >(null);
  const shortcutsDialog = useShortcutsDialog();

  const onZoomChange = useCallback((next: number) => {
    setZoom(clampZoom(next));
  }, []);

  const pushToast = useCallback((kind: ToastItem["kind"], text: string) => {
    const id = createToastId("pptx");
    setToasts((prev) => [...prev, { id, kind, text }]);
  }, []);
  const dismissToast = useCallback((id: string) => {
    setToasts((prev) => prev.filter((t) => t.id !== id));
  }, []);
  useEffect(() => {
    pushToastRef.current = pushToast;
  }, [pushToast]);
  // These refs back the format provider's call-time reads. We sync them
  // *during render* (rather than in `useEffect`) so the provider sees
  // the freshest state on the very first render after a state change —
  // otherwise the toolbar's `provider.hasSelection()` evaluates against
  // a stale ref and the controls render disabled for one frame and
  // never re-render to recover (refs don't trigger renders themselves).
  slideIndexRef.current = activeIndex;
  textSelectionRef.current = textSelection;
  selectedShapeIdRef.current = selectedShapeId;

  /* eslint-disable react-hooks/refs -- the provider stores refs for
     call-time reads; it never dereferences them during construction. */
  const [textFormatProvider] = useState(() =>
    createPptxFormatProvider({
      agentRef,
      slideIndexRef,
      selectionRef: textSelectionRef,
      selectedShapeIdRef,
      pushToast: (kind, text) => pushToastRef.current?.(kind, text),
    })
  );
  /* eslint-enable react-hooks/refs */

  const mountAgent = useCallback(async (buf: ArrayBuffer) => {
    const next = await PptxAgent.fromBuffer(buf);
    agentRef.current = next;
    setAgent(next);
    setActiveIndex(0);
    setSelectedShapeIds([]);
    setReady(true);
    setTick((t) => t + 1);
    setSaveState("saved");
    next.subscribe((_snap, mutation) => {
      setTick((t) => t + 1);
      // Surface bus rebase rejections (see
      // packages/core/src/commands/bus.ts.recomputeWorking).
      // A pending agent mutation that no longer applies after
      // an undo gets flipped to `rejected` with the
      // `rebase-failed` code; previously it just disappeared.
      if (mutation.status === "rejected" && mutation.rejection?.code === "rebase-failed") {
        pushToastRef.current?.(
          "warn",
          `An agent suggestion couldn't be re-applied after the last edit (${mutation.rejection.message})`
        );
        return;
      }
      setSaveState("modified");
    });
  }, []);

  useEffect(() => {
    let cancelled = false;
    void (async () => {
      try {
        // Four bootstrap paths, picked in priority order:
        //   1. `initialBytes` — host streams the deck straight in.
        //   2. `initialSource` — fetch a pre-existing .pptx.
        //   3. `initialBlank` — build an empty deck.
        //   4. Default — build the synthetic welcome deck.
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
          buf = await buildBlankPptx();
        } else {
          buf = await buildSamplePptx();
        }
        if (!cancelled) await mountAgent(buf);
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [mountAgent, pushToast, initialSource, initialBlank, initialBytes]);

  const onError = useCallback(
    (err: unknown) => pushToast("error", err instanceof Error ? err.message : String(err)),
    [pushToast]
  );

  // Stash the latest duplicate handler in a ref so the keyboard
  // listener doesn't have to re-bind every time `selectedShapeIds`
  // changes (which would also create a temporal-dead-zone here, since
  // `duplicateSelectedShapes` is declared further down).
  const duplicateRef = useRef<() => void>(() => {});
  const groupRef = useRef<() => void>(() => {});
  const ungroupRef = useRef<() => void>(() => {});

  // D1 — Undo / Redo + duplicate / group / ungroup. We bind on
  // document so they fire regardless of which surface element is
  // focused. Undo / redo go through the shared `handleUndoRedo`
  // helper (chord detection + form-field guard + bus dispatch
  // are identical across DOCX / XLSX / PPTX). The
  // duplicate/group/ungroup chords stay PPTX-local because they
  // aren't part of the cross-editor shortcut catalog.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (handleUndoRedo(e, agentRef.current)) return;
      const mod = e.metaKey || e.ctrlKey;
      if (!mod) return;
      const key = e.key.toLowerCase();
      const isDuplicate = key === "d" && !e.shiftKey && !e.altKey;
      const isUngroup = key === "g" && e.shiftKey && e.altKey;
      const isGroup = key === "g" && e.shiftKey && !e.altKey;
      if (!isDuplicate && !isGroup && !isUngroup) return;
      if (isFormField(e.target)) return;
      const a = agentRef.current;
      if (!a) return;
      e.preventDefault();
      if (isDuplicate) {
        duplicateRef.current();
      } else if (isUngroup) {
        ungroupRef.current();
      } else if (isGroup) {
        groupRef.current();
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  // D3 — Tab cycles selection across the slide's top-level shapes. With
  // Shift it cycles backward. Skipped when typing inside a real form
  // field so the browser's native tab-to-focus still works.
  useEffect(() => {
    const isFormField = (target: EventTarget | null): boolean => {
      if (!(target instanceof HTMLElement)) return false;
      const tag = target.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT") return true;
      if (target.isContentEditable) return true;
      return false;
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== "Tab") return;
      if (e.metaKey || e.ctrlKey || e.altKey) return;
      if (isFormField(e.target)) return;
      const a = agentRef.current;
      if (!a) return;
      const slide = a.getSnapshot().root.slides[activeIndex];
      if (!slide || slide.shapes.length === 0) return;
      e.preventDefault();
      const shapes = slide.shapes;
      const currentId = selectedShapeIds[selectedShapeIds.length - 1] ?? null;
      const currentIdx = currentId ? shapes.findIndex((s) => s.id === currentId) : -1;
      const dir = e.shiftKey ? -1 : 1;
      const len = shapes.length;
      const startIdx = currentIdx === -1 ? (e.shiftKey ? len - 1 : 0) : (currentIdx + dir + len) % len;
      setSelectedShapeIds([shapes[startIdx]!.id]);
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [activeIndex, selectedShapeIds]);

  // Clipboard handlers for the PPTX canvas — both intra-deck shape
  // copy/cut/paste and the existing cross-format XLSX → PPTX paste.
  // Bound to the `window` rather than a canvas-scoped element because
  // the slide surface doesn't own a tab-focusable host (selection is
  // mouse-driven) and the shortcut should always land on whichever
  // slide is currently active.
  //
  // Copy / cut: when one or more shapes are selected and the user is
  // *not* editing text inside a contenteditable / form field, we
  // serialise the selected shapes onto the clipboard via our embed
  // envelope (`pptx-shapes` payload). The browser's default text-
  // copy behaviour is preserved when text is selected inside an
  // editable region — we explicitly bail out in that case.
  //
  // Paste: dispatched against the embed envelope first; falls back to
  // ignoring the event so the browser's native handlers (image
  // paste, plain text into a focused text box) keep working.
  useEffect(() => {
    if (!isEmbedEnabled()) return;
    const isFormField = (target: EventTarget | null): boolean => {
      if (!(target instanceof HTMLElement)) return false;
      const tag = target.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT") return true;
      if (target.isContentEditable) return true;
      return false;
    };

    // Shape kinds we know how to round-trip through the clipboard
    // today. See `paste-shapes.ts` for the matching whitelist on the
    // command side. Pictures / charts / OLE / media require part
    // copying which we deliberately don't do here — same-deck
    // duplication of those is available via Cmd+D.
    const isCopyable = (s: Shape): boolean => {
      switch (s.kind) {
        case "text":
        case "table":
        case "connector":
          return true;
        case "group":
          return s.children.every(isCopyable);
        case "pic":
        case "chart":
        case "ole-spreadsheet":
        case "media":
        case "opaque":
          return false;
      }
    };

    const collectSelectedShapes = (): {
      shapes: Shape[];
      skipped: number;
      slideIndex: number;
    } | null => {
      const a = agentRef.current;
      if (!a) return null;
      const slideIndex = slideIndexRef.current;
      const slide = a.getSnapshot().root.slides[slideIndex];
      if (!slide) return null;
      const ids = selectedShapeIds;
      if (ids.length === 0) return null;
      const shapes: Shape[] = [];
      let skipped = 0;
      for (const id of ids) {
        const s = findShape(slide.shapes, id);
        if (!s) continue;
        if (isCopyable(s)) shapes.push(s);
        else skipped++;
      }
      if (shapes.length === 0) return null;
      return { shapes, skipped, slideIndex };
    };

    const writeShapesToClipboard = (
      e: ClipboardEvent,
      collected: { shapes: Shape[]; skipped: number; slideIndex: number }
    ): boolean => {
      if (!e.clipboardData) return false;
      const env = makeEnvelope("pptx", {
        kind: "pptx-shapes",
        shapes: collected.shapes,
        originLabel: `Slide ${collected.slideIndex + 1}`,
      });
      try {
        e.clipboardData.setData(EMBED_MIME, serializeEnvelope(env));
        // Also paint a plain-text fallback so external apps (email,
        // notes) get something readable. We collapse text shapes to
        // their concatenated paragraphs and skip non-text shapes.
        const fallback = collected.shapes
          .map((s) => (s.kind === "text" ? textShapePlain(s) : ""))
          .filter((t) => t.length > 0)
          .join("\n");
        if (fallback) e.clipboardData.setData("text/plain", fallback);
        return true;
      } catch {
        return false;
      }
    };

    // Sidebar-targeted copy/paste promotes the operation from
    // shape-level to slide-level. We only need to know that the event
    // landed inside the sidebar because that's where Word / PowerPoint
    // route their "duplicate slide" Cmd+C/Cmd+V chord. The marker is a
    // `data-testid="pptx-sidebar"` on the aside; checking
    // `closest()` is enough because every sidebar element bubbles
    // through it. Returns the source slide INDEX in the sidebar's
    // current list (which, by construction, equals the model index).
    const sidebarSlideIndexFromEvent = (target: EventTarget | null): number | null => {
      if (!(target instanceof HTMLElement)) return null;
      const aside = target.closest('[data-testid="pptx-sidebar"]');
      if (!aside) return null;
      // Prefer the explicit slide marker if the user clicked a
      // thumbnail; otherwise fall back to the active slide so a
      // sidebar-scoped Cmd+C with no thumbnail focus still copies the
      // currently-displayed slide.
      const thumb = target.closest<HTMLElement>("[data-slide-index]");
      if (thumb) {
        const v = Number.parseInt(thumb.dataset.slideIndex ?? "", 10);
        if (Number.isFinite(v) && v >= 0) return v;
      }
      return slideIndexRef.current;
    };

    const writeSlideToClipboard = (e: ClipboardEvent, slideIndex: number): boolean => {
      if (!e.clipboardData) return false;
      const a = agentRef.current;
      if (!a) return false;
      const slide = a.getSnapshot().root.slides[slideIndex];
      if (!slide) return false;
      const env = makeEnvelope("pptx", {
        kind: "pptx-slide-ref",
        slideIndex,
        sessionId: a.sessionId ?? "unknown",
        originLabel: `Slide ${slideIndex + 1}`,
      });
      try {
        e.clipboardData.setData(EMBED_MIME, serializeEnvelope(env));
        e.clipboardData.setData("text/plain", `Slide ${slideIndex + 1}`);
        return true;
      } catch {
        return false;
      }
    };

    const onCopy = (e: ClipboardEvent) => {
      if (isFormField(e.target)) return;
      const sidebarSlide = sidebarSlideIndexFromEvent(e.target);
      if (sidebarSlide !== null) {
        if (!writeSlideToClipboard(e, sidebarSlide)) return;
        e.preventDefault();
        pushToast("info", `Copied slide ${sidebarSlide + 1}`);
        return;
      }
      const collected = collectSelectedShapes();
      if (!collected) return;
      if (!writeShapesToClipboard(e, collected)) return;
      e.preventDefault();
      const n = collected.shapes.length;
      const skippedNote =
        collected.skipped > 0
          ? ` (${collected.skipped} unsupported shape${collected.skipped === 1 ? "" : "s"} skipped)`
          : "";
      pushToast("info", `Copied ${n} shape${n === 1 ? "" : "s"}${skippedNote}`);
    };

    const onCut = (e: ClipboardEvent) => {
      if (isFormField(e.target)) return;
      const sidebarSlide = sidebarSlideIndexFromEvent(e.target);
      if (sidebarSlide !== null) {
        // Cut on a slide thumbnail = copy + delete. Refuse to delete
        // the only slide so the deck stays valid.
        const a = agentRef.current;
        if (!a) return;
        const slides = a.getSnapshot().root.slides;
        if (slides.length <= 1) {
          pushToast("error", "Can't cut the only slide");
          return;
        }
        if (!writeSlideToClipboard(e, sidebarSlide)) return;
        e.preventDefault();
        void (async () => {
          try {
            await a.applyCommand({
              type: "pptx:delete-slide",
              payload: { slideIndex: sidebarSlide },
              source: "human",
            });
            pushToast("info", `Cut slide ${sidebarSlide + 1}`);
          } catch (err) {
            pushToast("error", err instanceof Error ? err.message : String(err));
          }
        })();
        return;
      }
      const collected = collectSelectedShapes();
      if (!collected) return;
      if (!writeShapesToClipboard(e, collected)) return;
      e.preventDefault();
      const a = agentRef.current;
      if (!a) return;
      const slideIndex = collected.slideIndex;
      const ids = collected.shapes.map((s) => s.id);
      void (async () => {
        try {
          for (const id of ids) {
            await a.applyCommand({
              type: "pptx:delete-shape",
              payload: { slideIndex, shapeId: id },
              source: "human",
            });
          }
          setSelectedShapeIds([]);
          const n = ids.length;
          pushToast("info", `Cut ${n} shape${n === 1 ? "" : "s"}`);
        } catch (err) {
          pushToast("error", err instanceof Error ? err.message : String(err));
        }
      })();
    };

    const onPaste = (e: ClipboardEvent) => {
      if (isFormField(e.target)) return;
      const raw = e.clipboardData?.getData(EMBED_MIME);
      const env = parseEnvelope(raw);
      if (!env) return;
      const agent = agentRef.current;
      if (!agent) return;
      const slideIndex = slideIndexRef.current;
      const payload = env.payload;
      switch (payload.kind) {
        case "xlsx-range": {
          e.preventDefault();
          e.stopPropagation();
          // Alt held → embed as a live OLE Excel object instead of a
          // materialised table. Mirrors PowerPoint's "Paste Special →
          // Microsoft Excel Worksheet Object" shortcut so power users
          // can opt-in without round-tripping through a dialog.
          const mode = isAltKeyPressed() ? "live" : "materialized";
          void (async () => {
            try {
              await applyXlsxRangeToPptx({
                agent,
                snapshot: payload.snapshot,
                slideIndex,
                mode,
              });
            } catch (err) {
              pushToast("error", err instanceof Error ? err.message : String(err));
            }
          })();
          return;
        }
        case "pptx-shapes": {
          e.preventDefault();
          e.stopPropagation();
          void (async () => {
            try {
              const before = agent.getSnapshot().root.slides[slideIndex]?.shapes.length ?? 0;
              await agent.applyCommand({
                type: "pptx:paste-shapes",
                payload: { slideIndex, shapes: payload.shapes },
                source: "human",
              });
              const slide = agent.getSnapshot().root.slides[slideIndex];
              if (slide) {
                const newIds = slide.shapes.slice(before).map((s) => s.id);
                if (newIds.length > 0) setSelectedShapeIds(newIds);
              }
              const n = payload.shapes.length;
              pushToast("info", `Pasted ${n} shape${n === 1 ? "" : "s"}`);
            } catch (err) {
              pushToast("error", err instanceof Error ? err.message : String(err));
            }
          })();
          return;
        }
        case "xlsx-chart-image": {
          // Charts paste as PNG; the image-insert path already handles
          // generic image clipboard data, so we let the existing
          // window paste handler pick it up rather than duplicating
          // the insert-image flow here.
          return;
        }
        case "pptx-slide-ref": {
          // Same-session slide paste: route through the typed
          // `pptx:duplicate-slide` command so all the slide-level
          // bookkeeping (rels, content types, idGen) stays under one
          // codepath. Cross-session paste (different agent / different
          // browser tab) is intentionally a no-op for now — moving a
          // full slide across documents requires copying media + chart
          // parts which is its own workstream.
          if (payload.sessionId !== (agent.sessionId ?? "unknown")) {
            pushToast(
              "info",
              "Cross-document slide paste isn't supported yet — open both decks in the same session."
            );
            return;
          }
          e.preventDefault();
          e.stopPropagation();
          void (async () => {
            try {
              await agent.applyCommand({
                type: "pptx:duplicate-slide",
                payload: { slideIndex: payload.slideIndex },
                source: "human",
              });
              // duplicate-slide inserts the clone right after the
              // source. Move it to sit immediately after the *active*
              // slide so paste lands where the user is looking — that
              // matches PowerPoint's behaviour and the deck stays in a
              // predictable order even if the source was elsewhere.
              const after = payload.slideIndex + 1;
              const target = slideIndexRef.current + 1;
              if (after !== target) {
                await agent.applyCommand({
                  type: "pptx:move-slide",
                  payload: { from: after, to: target },
                  source: "human",
                });
              }
              setActiveIndex(Math.min(target, agent.getSnapshot().root.slides.length - 1));
              pushToast("info", `Pasted slide ${payload.slideIndex + 1}`);
            } catch (err) {
              pushToast("error", err instanceof Error ? err.message : String(err));
            }
          })();
          return;
        }
      }
    };

    window.addEventListener("copy", onCopy);
    window.addEventListener("cut", onCut);
    window.addEventListener("paste", onPaste);
    const uninstallAlt = installAltKeyTracker();
    return () => {
      window.removeEventListener("copy", onCopy);
      window.removeEventListener("cut", onCut);
      window.removeEventListener("paste", onPaste);
      uninstallAlt();
    };
  }, [pushToast, selectedShapeIds]);

  const handleFile = useCallback(
    async (file: File, handle?: FileSystemFileHandle) => {
      const buf = await file.arrayBuffer();
      setDocName(file.name);
      setFileHandle(handle);
      try {
        await mountAgent(buf);
        pushToast("info", `Opened ${file.name}`);
      } catch (err) {
        onError(err);
      }
    },
    [mountAgent, onError, pushToast]
  );

  const handleOpenFile = useCallback(async () => {
    try {
      const opened = await openFile({
        description: PRODUCT_FILE_TYPES.pptx.description,
        mimeToExt: PRODUCT_FILE_TYPES.pptx.mimeToExt,
        accept: PRODUCT_FILE_TYPES.pptx.accept,
      });
      if (!opened) return;
      const file = new File([opened.bytes as BlobPart], opened.name, {
        type: PRODUCT_FILE_TYPES.pptx.primaryMime,
      });
      await handleFile(file, opened.handle);
    } catch (err) {
      onError(err);
    }
  }, [handleFile, onError]);

  const handleSave = useCallback(async () => {
    const a = agentRef.current;
    if (!a) return;
    try {
      setSaveState("saving");
      const buf = await a.exportFile();
      const bytes = new Uint8Array(buf);
      const mime = PRODUCT_FILE_TYPES.pptx.primaryMime;
      if (onSaveProp) {
        await onSaveProp(bytes, mime, docName);
        setSaveState("saved");
        pushToast("success", `Saved ${docName}`);
        return;
      }
      const inPlace = await saveFile(bytes, docName, mime, fileHandle);
      setSaveState("saved");
      pushToast("success", inPlace ? `Saved ${docName}` : `Downloaded ${docName}`);
    } catch (err) {
      setSaveState("error");
      onError(err);
    }
  }, [docName, fileHandle, onError, onSaveProp, pushToast]);

  const handleExport = useCallback(
    async (format: ExportFormat, options?: ExportOptionValues) => {
      const a = agentRef.current;
      if (!a) return;
      const baseName = stripPptxExtension(docName);
      // Single-slide exports get a `-slideN` suffix so a "current
      // slide" PNG can't silently overwrite a previous deck-level
      // export sitting in the user's Downloads folder.
      const slideIdx = slideIndexRef.current;
      const isCurrentSlideFormat = format.group === "current";
      const slideSuffix = isCurrentSlideFormat ? `-slide${slideIdx + 1}` : "";
      const downloadName = `${baseName}${slideSuffix}.${format.extension}`;
      try {
        switch (format.id) {
          case "pptx": {
            const buf = await a.exportFile();
            await saveFile(new Uint8Array(buf), downloadName, format.mime, undefined);
            break;
          }
          case "pdf":
          case "html": {
            const buf = await a.exportFile();
            const out = await convertViaServer({
              bytes: new Uint8Array(buf),
              sourceExt: "pptx",
              targetExt: format.id,
              filename: baseName,
            });
            downloadBlob(out, downloadName);
            break;
          }
          case "slide-pdf": {
            // Server-converted single-slide PDF. We send the whole
            // pptx and let LibreOffice's `PageRange` filter cut it
            // down — that keeps the renderer a single code path
            // (vs. building a one-slide pptx in the browser) and
            // matches what PowerPoint's "Print → Selection" does
            // under the hood.
            const buf = await a.exportFile();
            const out = await convertViaServer({
              bytes: new Uint8Array(buf),
              sourceExt: "pptx",
              targetExt: "pdf",
              filename: `${baseName}-slide${slideIdx + 1}`,
              pageRange: String(slideIdx + 1),
            });
            downloadBlob(out, downloadName);
            break;
          }
          case "slide-png": {
            const snap = a.getSnapshot();
            const scale = parseScale(options?.scale);
            const blob = await snapshotToSlidePng(snap, slideIdx, { scale });
            downloadBlob(blob, downloadName);
            break;
          }
          case "slide-jpeg": {
            const snap = a.getSnapshot();
            const scale = parseScale(options?.scale);
            const quality = parseJpegQuality(options?.quality);
            const blob = await snapshotToSlideJpeg(snap, slideIdx, { scale, quality });
            downloadBlob(blob, downloadName);
            break;
          }
          case "slide-svg": {
            const snap = a.getSnapshot();
            const svg = snapshotToSlideSvg(snap, slideIdx);
            downloadBlob(new Blob([svg], { type: format.mime }), downloadName);
            break;
          }
          case "png-zip": {
            const snap = a.getSnapshot();
            const total = snap.root.slides.length;
            const rangeRaw = typeof options?.slideRange === "string" ? options.slideRange : "";
            const indices = parseSlideRange(rangeRaw, total);
            if (indices.length === 0) {
              throw new Error("Slide range matched no slides.");
            }
            const scale = parseScale(options?.scale);
            const blob = await snapshotToPngZip(snap, { scale, indices });
            downloadBlob(blob, downloadName);
            break;
          }
          case "svg-zip": {
            const snap = a.getSnapshot();
            const blob = await snapshotToSvgZip(snap);
            downloadBlob(blob, downloadName);
            break;
          }
          default:
            throw new Error(`Unsupported export format: ${format.id}`);
        }
        pushToast("success", `Exported ${downloadName}`);
      } catch (err) {
        onError(err);
        throw err;
      }
    },
    [docName, onError, pushToast]
  );

  const snap = agent?.getSnapshot() ?? null;
  void tick;
  const slides = snap?.root.slides ?? [];
  const slideSize = snap?.root.slideSize ?? { cxEmu: 12192000, cyEmu: 6858000 };
  const themeDefault = snap?.root.themeDefault;
  const slide = slides[activeIndex];

  const selectedShape = useMemo<Shape | null>(() => {
    if (!slide || !selectedShapeId) return null;
    return findShape(slide.shapes, selectedShapeId);
  }, [slide, selectedShapeId]);

  // Shape ids that have an unresolved comment thread on the active
  // slide. Reply comments and resolved threads are filtered out so the
  // yellow indicator reads "open conversation".
  const commentedShapeIds = useMemo<ReadonlyArray<string>>(() => {
    if (!snap || !slide || !slide.commentsPartPath) return [];
    const part = snap.root.commentsByPart.get(slide.commentsPartPath);
    if (!part) return [];
    const ids = new Set<string>();
    for (const c of part.comments) {
      if (c.parentId) continue;
      if (c.resolved) continue;
      if (!c.shapeId) continue;
      ids.add(c.shapeId);
    }
    return [...ids];
  }, [snap, slide]);

  const currentShapeFill = useMemo<FillSpec | null>(() => {
    if (!selectedShape) return null;
    if (selectedShape.kind !== "text" && selectedShape.kind !== "pic") return null;
    return readShapeFill(selectedShape);
  }, [selectedShape]);

  /**
   * `<p:sld show="0">` lives in `slideRootAttrs` (preserved verbatim
   * by the parser; no typed model field). The Slideshow ribbon
   * Hide-Slide toggle reads that flag here so its pressed state
   * tracks the document, even when the slide was hidden via the
   * sidebar context-menu.
   */
  const activeSlideHidden = useMemo<boolean>(() => {
    if (!slide) return false;
    return slide.slideRootAttrs.show === "0";
  }, [slide]);

  /**
   * Parse the deck's `<p:showPr>` element (if any) out of
   * `presentationOpaqueTail` so the {@link SetUpShowDialog} opens
   * with the values from the document. PowerPoint defaults are
   * applied for missing flags, matching the dialog's own defaults.
   */
  const currentShowOptions = useMemo<SetUpShowValues>(() => {
    return readShowOptions(snap);
  }, [snap]);

  const currentSlideBackground = useMemo<FillSpec | null>(() => {
    const s = slides[activeIndex];
    if (!s) return null;
    return readSlideBackgroundFillSpec(s);
  }, [slides, activeIndex]);

  // ActiveTextFormat for the shared TextFormatBar. Recomputed on every
  // render that touches the snapshot tick or text selection so the
  // toolbar reflects the current run-level formatting at the caret.
  const textFormatActive = useMemo(() => {
    void tick;
    return computePptxActive(agent, activeIndex, textSelection, selectedShapeId);
  }, [agent, activeIndex, textSelection, selectedShapeId, tick]);

  // Resolve the text shape the alignment / anchor controls should
  // operate on: prefer a live text-edit caret's shape, fall back to
  // the canvas's primary shape selection. Mirrors what the shared
  // `pptxFormatProvider` does for run-level formatting.
  const activeTextShape = useMemo<TextShape | null>(() => {
    void tick;
    if (!slide) return null;
    if (textSelection) {
      const s = findShape(slide.shapes, textSelection.shapeId);
      if (s && s.kind === "text") return s;
    }
    if (selectedShape && selectedShape.kind === "text") return selectedShape;
    return null;
  }, [slide, textSelection, selectedShape, tick]);

  const activeTextAlignment = useMemo<"left" | "center" | "right" | "justify" | null>(() => {
    if (!activeTextShape) return null;
    const paraIdx = textSelection?.paragraph ?? 0;
    const para = activeTextShape.txBody.paragraphs[paraIdx];
    return para?.properties.alignment ?? null;
  }, [activeTextShape, textSelection]);

  const activeTextAnchor = useMemo<TextAnchor | null>(() => {
    if (!activeTextShape) return null;
    const v =
      activeTextShape.txBody.bodyPrRaw?.attrs.anchor ??
      activeTextShape.txBody.bodyPrRaw?.rawAttrs["@_anchor"];
    if (v === "ctr") return "middle";
    if (v === "b") return "bottom";
    if (v === "t") return "top";
    return null;
  }, [activeTextShape]);

  // Default layout for new slides — driven by the Master panel.
  // `null` keeps PowerPoint-equivalent "auto" behaviour (the
  // `pptx:add-slide` command picks based on what the deck has). When
  // the user explicitly opts into a layout from the Master panel,
  // every subsequent New-slide click respects that choice until they
  // reset back to auto.
  const [defaultLayoutKind, setDefaultLayoutKind] = useState<LayoutKindPayload | null>(null);

  const addSlide = useCallback(async () => {
    const a = agentRef.current;
    if (!a) return;
    try {
      const payload = defaultLayoutKind ? { layoutKind: defaultLayoutKind } : {};
      await a.applyCommand({ type: "pptx:add-slide", payload, source: "human" });
      setActiveIndex(a.getSnapshot().root.slides.length - 1);
    } catch (err) {
      onError(err);
    }
  }, [defaultLayoutKind, onError]);

  const addSlideWithLayout = useCallback(
    async (kind: LayoutKindPayload) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        await a.applyCommand({
          type: "pptx:add-slide",
          payload: { layoutKind: kind },
          source: "human",
        });
        setActiveIndex(a.getSnapshot().root.slides.length - 1);
      } catch (err) {
        onError(err);
      }
    },
    [onError]
  );

  const setSlideLayout = useCallback(
    async (kind: LayoutKindPayload) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        await a.applyCommand({
          type: "pptx:set-slide-layout",
          payload: { slideIndex: activeIndex, layoutKind: kind },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError]
  );

  const deleteSlideAt = useCallback(
    async (slideIndex: number) => {
      const a = agentRef.current;
      if (!a) return;
      const total = a.getSnapshot().root.slides.length;
      if (total <= 1) {
        pushToast("warn", "Cannot delete the last slide.");
        return;
      }
      if (slideIndex < 0 || slideIndex >= total) return;
      try {
        await a.applyCommand({
          type: "pptx:delete-slide",
          payload: { slideIndex },
          source: "human",
        });
        // Keep the selection on the same slot when possible so the
        // user lands on the slide that visually replaced the deleted
        // one. When the last slide was removed, fall back to the new
        // last index.
        const nextTotal = total - 1;
        setActiveIndex(Math.min(slideIndex, nextTotal - 1));
      } catch (err) {
        onError(err);
      }
    },
    [onError, pushToast]
  );

  const deleteSlide = useCallback(async () => {
    await deleteSlideAt(activeIndex);
  }, [activeIndex, deleteSlideAt]);

  // Keyboard Delete / Backspace on the slides sidebar removes the
  // focused thumbnail's slide. Mirrors PowerPoint / Keynote: clicking
  // a thumbnail focuses its button, then pressing Delete drops that
  // slide. We resolve the target slide from the focused element's
  // `data-slide-index` (present on every sidebar row) and fall back
  // to the active slide if the user pressed Delete while the aside
  // itself — but no specific row — held focus.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== "Delete" && e.key !== "Backspace") return;
      if (e.metaKey || e.ctrlKey || e.altKey || e.shiftKey) return;
      const target = e.target;
      if (!(target instanceof HTMLElement)) return;
      // Never hijack typing inside a real input or contenteditable.
      const tag = target.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT") return;
      if (target.isContentEditable) return;
      const aside = target.closest('[data-testid="pptx-sidebar"]');
      if (!aside) return;
      const row = target.closest<HTMLElement>("[data-slide-index]");
      const idx = row ? Number.parseInt(row.dataset.slideIndex ?? "", 10) : activeIndex;
      if (!Number.isFinite(idx) || idx < 0) return;
      e.preventDefault();
      void deleteSlideAt(idx);
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [activeIndex, deleteSlideAt]);

  const duplicateSlide = useCallback(async () => {
    const a = agentRef.current;
    if (!a) return;
    try {
      await a.applyCommand({
        type: "pptx:duplicate-slide",
        payload: { slideIndex: activeIndex },
        source: "human",
      });
      setActiveIndex(activeIndex + 1);
    } catch (err) {
      onError(err);
    }
  }, [activeIndex, onError]);

  // D7 — slide thumbnail drag-reorder. The sidebar reports
  // `(from, to)` indices already adjusted for splice semantics so we
  // can dispatch `pptx:move-slide` directly.
  const moveSlide = useCallback(
    async (from: number, to: number) => {
      const a = agentRef.current;
      if (!a) return;
      if (from === to) return;
      try {
        await a.applyCommand({
          type: "pptx:move-slide",
          payload: { from, to },
          source: "human",
        });
        // Track the moved slide so it stays selected after the
        // reorder. `setActiveIndex` is otherwise sticky to the old
        // index, which would visually jump to the wrong slide.
        setActiveIndex(to);
      } catch (err) {
        onError(err);
      }
    },
    [onError]
  );

  // D7 — slide thumbnail context menu. Handles every entry the sidebar
  // can surface (`SlideContextAction`), reusing the existing
  // add/duplicate/delete/move commands. Insert-before / insert-after
  // dispatch `pptx:add-slide` with the explicit `at` slot rather than
  // the toolbar's append-to-end semantics.
  const handleSlideContextAction = useCallback(
    async (slideIndex: number, action: SlideContextAction) => {
      const a = agentRef.current;
      if (!a) return;
      const total = a.getSnapshot().root.slides.length;
      try {
        switch (action) {
          case "insert-before": {
            await a.applyCommand({
              type: "pptx:add-slide",
              payload: { at: slideIndex },
              source: "human",
            });
            setActiveIndex(slideIndex);
            break;
          }
          case "insert-after": {
            await a.applyCommand({
              type: "pptx:add-slide",
              payload: { at: slideIndex + 1 },
              source: "human",
            });
            setActiveIndex(slideIndex + 1);
            break;
          }
          case "duplicate": {
            await a.applyCommand({
              type: "pptx:duplicate-slide",
              payload: { slideIndex },
              source: "human",
            });
            setActiveIndex(slideIndex + 1);
            break;
          }
          case "delete": {
            if (total <= 1) {
              pushToast("warn", "Cannot delete the last slide.");
              return;
            }
            await a.applyCommand({
              type: "pptx:delete-slide",
              payload: { slideIndex },
              source: "human",
            });
            setActiveIndex(Math.max(0, Math.min(slideIndex, total - 2)));
            break;
          }
          case "move-up": {
            if (slideIndex === 0) return;
            await a.applyCommand({
              type: "pptx:move-slide",
              payload: { from: slideIndex, to: slideIndex - 1 },
              source: "human",
            });
            setActiveIndex(slideIndex - 1);
            break;
          }
          case "move-down": {
            if (slideIndex >= total - 1) return;
            await a.applyCommand({
              type: "pptx:move-slide",
              payload: { from: slideIndex, to: slideIndex + 1 },
              source: "human",
            });
            setActiveIndex(slideIndex + 1);
            break;
          }
          case "move-to-start": {
            if (slideIndex === 0) return;
            await a.applyCommand({
              type: "pptx:move-slide",
              payload: { from: slideIndex, to: 0 },
              source: "human",
            });
            setActiveIndex(0);
            break;
          }
          case "move-to-end": {
            if (slideIndex >= total - 1) return;
            await a.applyCommand({
              type: "pptx:move-slide",
              payload: { from: slideIndex, to: total - 1 },
              source: "human",
            });
            setActiveIndex(total - 1);
            break;
          }
          default: {
            const exhaust: never = action;
            throw new Error(`unhandled slide action: ${String(exhaust)}`);
          }
        }
      } catch (err) {
        onError(err);
      }
    },
    [onError, pushToast]
  );

  // Each "Insert" command drops the new shape near the slide centre at a
  // sensible default size. We avoid stacking duplicates by nudging by an
  // extra (count * step) so successive inserts don't pile up at the exact
  // same coordinates.
  const insertOffset = useCallback((): { x: number; y: number } => {
    const a = agentRef.current;
    const s = a?.getSnapshot().root.slides[activeIndex];
    const count = s?.shapes.length ?? 0;
    const step = 200_000;
    return {
      x: 1_000_000 + (count % 8) * step,
      y: 1_000_000 + (count % 8) * step,
    };
  }, [activeIndex]);

  const handleXlsxPickerSubmit = useCallback(
    async (result: XlsxRangePickerResult) => {
      const a = agentRef.current;
      setXlsxPickerOpen(null);
      if (!a) return;
      try {
        await applyXlsxEmbed({
          target: { kind: "pptx", agent: a, slideIndex: activeIndex },
          snapshot: result.snapshot,
          mode: result.mode,
        });
        const label =
          result.mode === "live" ? "embedded spreadsheet" : result.mode === "chart" ? "chart" : "table";
        pushToast("info", `Inserted ${label} from xlsx.`);
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError, pushToast]
  );

  /**
   * Double-click → "Edit Data" entry point. Resolves the embedded
   * `.xlsx` bytes for the activated chart / OLE spreadsheet shape and
   * pops the {@link EmbeddedXlsxModal}. Emits a friendly toast (and
   * does NOT throw) when the shape has no recoverable bytes — for
   * example, a freshly-inserted chart whose `pendingGrid` was never
   * set (defensive; should not happen with our current handlers).
   */
  const handleShapeActivate = useCallback(
    async (info: { shapeId: string; shape: Shape }) => {
      const a = agentRef.current;
      if (!a) return;
      const sh = info.shape;
      if (sh.kind !== "chart" && sh.kind !== "ole-spreadsheet") return;
      try {
        const ref = resolveEmbeddedXlsxRef({
          source: { kind: "pptx", agent: a },
          ...(sh.kind === "chart"
            ? { chartPartPath: sh.chartPartPath }
            : { embeddingPartPath: sh.embeddingPartPath }),
        });
        if (!ref) {
          pushToast("info", "This object has no embedded workbook to edit.");
          return;
        }
        const bytes = await readEmbeddedXlsxBytes({
          source: { kind: "pptx", agent: a },
          embeddingPartPath: ref.embeddingPartPath,
        });
        if (!bytes) {
          pushToast("info", "Embedded workbook bytes could not be loaded.");
          return;
        }
        setEditingEmbed({
          bytes,
          embeddingPartPath: ref.embeddingPartPath,
          chartPartPath: ref.chartPartPath,
          slideIndex: activeIndex,
          shapeId: info.shapeId,
          title: sh.kind === "chart" ? "Edit chart data" : "Edit spreadsheet",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError, pushToast]
  );

  /**
   * Persist edits made in the {@link EmbeddedXlsxModal} back into the
   * host PPTX. For OLE spreadsheets we only need to refresh the
   * embedded `.xlsx` bytes — Office re-renders the preview lazily.
   * For charts we additionally dispatch `pptx:set-chart-data` from
   * the modal's projected grid: row 1 is treated as the categories
   * column (skipping A1) and row 2+ as series rows whose first cell
   * is the series name and whose remaining cells are numeric values.
   * That mirrors the "first row is header / first column is label"
   * convention that `buildChartGrid` and `set-chart-data` already use
   * elsewhere in the codebase.
   */
  const handleEmbeddedXlsxSave = useCallback(
    async (result: {
      readonly bytes: Uint8Array;
      readonly grid: ReadonlyArray<ReadonlyArray<string | number | null>>;
    }) => {
      const a = agentRef.current;
      const ctx = editingEmbed;
      setEditingEmbed(null);
      if (!a || !ctx) return;
      try {
        await a.applyCommand({
          type: "pptx:update-spreadsheet",
          payload: {
            embeddingPartPath: ctx.embeddingPartPath,
            bytes: result.bytes,
            previewGrid: result.grid,
          },
          source: "human",
        });
        if (ctx.chartPartPath) {
          const chartUpdate = projectGridToChartData(result.grid);
          if (chartUpdate) {
            await a.applyCommand({
              type: "pptx:set-chart-data",
              payload: {
                slideIndex: ctx.slideIndex,
                shapeId: ctx.shapeId,
                categories: chartUpdate.categories,
                series: chartUpdate.series,
              },
              source: "human",
            });
          }
        }
        pushToast("info", "Saved embedded data.");
      } catch (err) {
        onError(err);
      }
    },
    [editingEmbed, onError, pushToast]
  );

  const addTextBox = useCallback(async () => {
    const a = agentRef.current;
    if (!a) return;
    try {
      const off = insertOffset();
      await a.applyCommand({
        type: "pptx:add-text-box",
        payload: {
          slideIndex: activeIndex,
          text: "New text box",
          x: off.x,
          y: off.y,
          width: 4_000_000,
          height: 800_000,
        },
        source: "human",
      });
      const s = a.getSnapshot().root.slides[activeIndex];
      setSelectedShapeIds([s.shapes[s.shapes.length - 1]!.id]);
    } catch (err) {
      onError(err);
    }
  }, [activeIndex, insertOffset, onError]);

  const addShape = useCallback(
    async (preset: ShapePreset) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        const off = insertOffset();
        const isLine = preset === "line";
        await a.applyCommand({
          type: "pptx:add-shape",
          payload: {
            slideIndex: activeIndex,
            preset,
            x: off.x,
            y: off.y,
            width: isLine ? 3_000_000 : 2_500_000,
            height: isLine ? 0 : preset === "ellipse" ? 1_500_000 : 1_500_000,
          },
          source: "human",
        });
        const s = a.getSnapshot().root.slides[activeIndex];
        setSelectedShapeIds([s.shapes[s.shapes.length - 1]!.id]);
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, insertOffset, onError]
  );

  // Toggle connector tool mode for the requested type. Re-clicking
  // the same type while it's already armed exits the mode (matches
  // Figma / Slides). The actual `pptx:add-connector` command runs
  // from inside the canvas once the user finishes their drag — that's
  // what makes the experience feel like a real drawing tool instead
  // of a "click button → guess where the line landed" form.
  const startConnectorTool = useCallback((connectorType: "straight" | "elbow" | "curved") => {
    setConnectorTool((prev) => (prev?.type === connectorType ? null : { type: connectorType }));
  }, []);
  const exitConnectorTool = useCallback(() => setConnectorTool(null), []);

  // Apply a partial style patch from the floating connector mini-bar.
  // The bar passes only the fields the user actually touched; we relay
  // them straight through to `pptx:set-connector-style` so undo/redo
  // sees a single command per click instead of separate commands per
  // field. The empty-patch guard mirrors the command's own validation
  // so a no-op interaction (e.g. picker closed without a change) costs
  // zero round-trips.
  const applyConnectorStylePatch = useCallback(
    async (shapeId: string, patch: ConnectorStylePatch) => {
      const a = agentRef.current;
      if (!a) return;
      const hasField =
        patch.connectorType !== undefined ||
        patch.strokeColor !== undefined ||
        patch.strokeWidthEmu !== undefined ||
        patch.strokeDash !== undefined ||
        patch.headEnd !== undefined ||
        patch.tailEnd !== undefined;
      if (!hasField) return;
      try {
        await a.applyCommand({
          type: "pptx:set-connector-style",
          source: "human",
          payload: {
            slideIndex: activeIndex,
            shapeId,
            ...patch,
          },
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError]
  );

  // Translate a connector mini-bar action tag into the appropriate
  // command(s). `reroute` and `swap` are 1:1 with their commands;
  // `detach` walks the slide once to resolve each anchored endpoint to
  // its current slide-coordinate point and rewrites it as a free
  // endpoint there, so the visual position survives the disconnect.
  const applyConnectorAction = useCallback(
    async (shapeId: string, action: ConnectorAction) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        if (action === "reroute") {
          await a.applyCommand({
            type: "pptx:reroute-connector",
            source: "human",
            payload: { slideIndex: activeIndex, shapeId },
          });
          return;
        }
        if (action === "swap") {
          await a.applyCommand({
            type: "pptx:swap-connector-direction",
            source: "human",
            payload: { slideIndex: activeIndex, shapeId },
          });
          return;
        }
        if (action === "detach") {
          const slide = a.getSnapshot().root.slides[activeIndex];
          if (!slide) return;
          const target = findShape(slide.shapes, shapeId);
          if (!target || target.kind !== "connector") return;
          const map = new Map<number, Shape>();
          for (const s of slide.shapes) collectShapesByCNvPrIdLocal(s, map);
          const connector = target as ConnectorShape;
          for (const which of ["start", "end"] as const) {
            const ep = which === "start" ? connector.start : connector.end;
            if (ep.kind !== "anchored") continue;
            const pt = resolveConnectorEndpoint(ep, map);
            if (!pt) continue;
            await a.applyCommand({
              type: "pptx:set-connector-endpoint",
              source: "human",
              payload: {
                slideIndex: activeIndex,
                shapeId,
                which,
                endpoint: { kind: "free", xEmu: pt.x, yEmu: pt.y },
              },
            });
          }
        }
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError]
  );

  const insertImage = useCallback(
    async (file: File) => {
      const a = agentRef.current;
      if (!a) return;
      const mime = (file.type || "").toLowerCase();
      if (!SUPPORTED_IMAGE_MIME.has(mime)) {
        pushToast(
          "error",
          `Unsupported image type "${mime || "unknown"}". Use PNG, JPEG, GIF, BMP, or WEBP.`
        );
        return;
      }
      try {
        const buf = await file.arrayBuffer();
        const bytes = new Uint8Array(buf);
        const intrinsic = await readIntrinsicSize(bytes, mime);
        // Cap the displayed width at ~half a 16:9 slide; preserves aspect.
        const maxWidthEmu = Math.floor(slideSize.cxEmu / 2);
        let widthEmu = intrinsic.width * 9525; // 1 px ≈ 9525 EMU @ 96 DPI
        let heightEmu = intrinsic.height * 9525;
        if (widthEmu > maxWidthEmu) {
          const r = maxWidthEmu / widthEmu;
          widthEmu = Math.round(widthEmu * r);
          heightEmu = Math.round(heightEmu * r);
        }
        if (widthEmu <= 0 || heightEmu <= 0) {
          widthEmu = 2_000_000;
          heightEmu = 2_000_000;
        }
        const off = insertOffset();
        await a.applyCommand({
          type: "pptx:insert-image",
          payload: {
            slideIndex: activeIndex,
            data: bytes,
            mimeType: mime === "image/jpg" ? "image/jpeg" : mime,
            x: off.x,
            y: off.y,
            width: widthEmu,
            height: heightEmu,
            altText: file.name,
            name: file.name,
          },
          source: "human",
        });
        const s = a.getSnapshot().root.slides[activeIndex];
        setSelectedShapeIds([s.shapes[s.shapes.length - 1]!.id]);
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, insertOffset, onError, pushToast, slideSize.cxEmu]
  );

  /**
   * Insert an embedded video / audio file on the active slide via
   * `pptx:insert-media`. Mirrors `insertImage`: validate the MIME
   * client-side (the command rejects unsupported types anyway, but a
   * toast reads better than an "invalid-payload" trace), centre the
   * default-sized box on the slide, and select the freshly-stamped
   * `MediaShape` so the user can immediately drag/move it.
   */
  const insertMedia = useCallback(
    async (file: File) => {
      const a = agentRef.current;
      if (!a) return;
      const mime = (file.type || "").toLowerCase();
      let mediaType: "video" | "audio";
      if (SUPPORTED_VIDEO_MIME.has(mime)) {
        mediaType = "video";
      } else if (SUPPORTED_AUDIO_MIME.has(mime)) {
        mediaType = "audio";
      } else if (mime.startsWith("video/") || mime.startsWith("audio/")) {
        pushToast(
          "error",
          `Unsupported media type "${mime}". Use MP4, WebM, MOV (video) or MP3, M4A, WAV, OGG (audio).`
        );
        return;
      } else {
        pushToast("error", `"${file.name}" isn't a video or audio file.`);
        return;
      }
      try {
        const buf = await file.arrayBuffer();
        const bytes = new Uint8Array(buf);
        // Default sizes per spec: 4 inch wide for video (16:9), 2 inch
        // wide for audio (rendered as a small bar). 1 in = 914400 EMU.
        const widthEmu = mediaType === "video" ? 4 * 914_400 : 2 * 914_400;
        const heightEmu = mediaType === "video" ? Math.round((widthEmu * 9) / 16) : Math.round(0.6 * 914_400);
        const xEmu = Math.max(0, Math.round((slideSize.cxEmu - widthEmu) / 2));
        const yEmu = Math.max(0, Math.round((slideSize.cyEmu - heightEmu) / 2));
        await a.applyCommand({
          type: "pptx:insert-media",
          payload: {
            slideIndex: activeIndex,
            mediaType,
            contentType: mime,
            bytes,
            position: { xEmu, yEmu },
            size: { cxEmu: widthEmu, cyEmu: heightEmu },
            name: file.name,
          },
          source: "human",
        });
        const s = a.getSnapshot().root.slides[activeIndex];
        const last = s?.shapes[s.shapes.length - 1];
        if (last) setSelectedShapeIds([last.id]);
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError, pushToast, slideSize.cxEmu, slideSize.cyEmu]
  );

  /**
   * D9 — replace the bitmap behind the currently selected `Picture`
   * with a user-picked file. Position, size, alt-text, and any
   * `spPrTail` styling are preserved by `pptx:replace-picture-media`.
   * No-ops when the selection isn't a picture.
   */
  const replaceSelectedPicture = useCallback(
    async (file: File) => {
      const a = agentRef.current;
      if (!a || !selectedShapeId) return;
      const sh = slide ? findShape(slide.shapes, selectedShapeId) : null;
      if (!sh || sh.kind !== "pic") {
        pushToast("error", "Replace image only works on the selected picture.");
        return;
      }
      const mime = (file.type || "").toLowerCase();
      if (!SUPPORTED_IMAGE_MIME.has(mime)) {
        pushToast(
          "error",
          `Unsupported image type "${mime || "unknown"}". Use PNG, JPEG, GIF, BMP, or WEBP.`
        );
        return;
      }
      try {
        const buf = await file.arrayBuffer();
        const bytes = new Uint8Array(buf);
        await a.applyCommand({
          type: "pptx:replace-picture-media",
          payload: {
            slideIndex: activeIndex,
            shapeId: sh.id,
            data: bytes,
            mimeType: mime === "image/jpg" ? "image/jpeg" : mime,
            altText: file.name,
          },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError, pushToast, selectedShapeId, slide]
  );

  /**
   * Fired by the canvas when the user activates an empty placeholder
   * (double-click, or PowerPoint-style "click again to enter") whose
   * type isn't text. For `pic` we open a file picker, drop the chosen
   * image into the slot the placeholder was occupying — same x / y /
   * cx / cy — and then remove the placeholder. The two commands are
   * issued sequentially so each sees the previous mutation; if the
   * insert fails we leave the placeholder untouched, which preserves
   * the slot for a retry.
   *
   * Other placeholder types (chart / tbl / dgm / media) trigger a
   * "not yet implemented" toast for now — the wiring scaffolding is
   * here for when those wizards exist.
   */
  const handlePlaceholderActivate = useCallback(
    async (info: { shapeId: string; placeholder: { type?: string; idx?: number } }) => {
      const a = agentRef.current;
      if (!a) return;
      const type = info.placeholder.type ?? "";
      if (type !== "pic") {
        pushToast(
          "info",
          `Filling "${type}" placeholders from the canvas isn't supported yet — use the toolbar.`
        );
        return;
      }
      const currentSlide = a.getSnapshot().root.slides[activeIndex];
      const placeholderShape = currentSlide ? findShape(currentSlide.shapes, info.shapeId) : null;
      if (!placeholderShape || !placeholderShape.position || !placeholderShape.size) {
        return;
      }
      const slot = {
        x: placeholderShape.position.xEmu,
        y: placeholderShape.position.yEmu,
        cx: placeholderShape.size.cxEmu,
        cy: placeholderShape.size.cyEmu,
      };
      const file = await pickImageFile();
      if (!file) return;
      const mime = (file.type || "").toLowerCase();
      if (!SUPPORTED_IMAGE_MIME.has(mime)) {
        pushToast(
          "error",
          `Unsupported image type "${mime || "unknown"}". Use PNG, JPEG, GIF, BMP, or WEBP.`
        );
        return;
      }
      try {
        const buf = await file.arrayBuffer();
        const bytes = new Uint8Array(buf);
        await a.applyCommand({
          type: "pptx:insert-image",
          payload: {
            slideIndex: activeIndex,
            data: bytes,
            mimeType: mime === "image/jpg" ? "image/jpeg" : mime,
            x: slot.x,
            y: slot.y,
            width: slot.cx,
            height: slot.cy,
            altText: file.name,
            name: file.name,
          },
          source: "human",
        });
        await a.applyCommand({
          type: "pptx:delete-shape",
          payload: { slideIndex: activeIndex, shapeId: info.shapeId },
          source: "human",
        });
        const updated = a.getSnapshot().root.slides[activeIndex];
        const last = updated?.shapes[updated.shapes.length - 1];
        if (last) setSelectedShapeIds([last.id]);
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError, pushToast]
  );

  const deleteSelectedShape = useCallback(async () => {
    const a = agentRef.current;
    if (!a || selectedShapeIds.length === 0) return;
    try {
      // Delete each selected shape independently. We dispatch sequentially
      // so each command sees the previous mutation's snapshot — the
      // command bus rejects "unknown shape" otherwise when later ids
      // happen to share group ancestry.
      for (const id of selectedShapeIds) {
        await a.applyCommand({
          type: "pptx:delete-shape",
          payload: { slideIndex: activeIndex, shapeId: id },
          source: "human",
        });
      }
      setSelectedShapeIds([]);
    } catch (err) {
      onError(err);
    }
  }, [activeIndex, onError, selectedShapeIds]);

  // D3 — Cmd+D duplicate. Clones every selected shape on the active
  // slide using the duplicate-shape command (¼" diagonal nudge), then
  // selects the freshly-minted clones so the user can immediately
  // drag/move them. Skipped when nothing is selected.
  // D3 — Cmd+Shift+G group / Cmd+Shift+Alt+G ungroup. Uses the
  // group-shapes / ungroup-shape commands. We refuse silently if the
  // selection doesn't satisfy the command's preconditions (the command
  // bus would reject anyway, but a toast on every failed Tab cycle is
  // noisy).
  const groupSelectedShapes = useCallback(async () => {
    const a = agentRef.current;
    if (!a || selectedShapeIds.length < 2) return;
    try {
      await a.applyCommand({
        type: "pptx:group-shapes",
        payload: { slideIndex: activeIndex, shapeIds: selectedShapeIds },
        source: "human",
      });
      const slide = a.getSnapshot().root.slides[activeIndex];
      const last = slide.shapes[slide.shapes.length - 1];
      if (last && last.kind === "group") setSelectedShapeIds([last.id]);
    } catch (err) {
      onError(err);
    }
  }, [activeIndex, onError, selectedShapeIds]);

  const ungroupSelectedShape = useCallback(async () => {
    const a = agentRef.current;
    if (!a || selectedShapeIds.length !== 1) return;
    const slide = a.getSnapshot().root.slides[activeIndex];
    const target = slide.shapes.find((s) => s.id === selectedShapeIds[0]);
    if (!target || target.kind !== "group") return;
    const childIds = target.children.map((c) => c.id);
    try {
      await a.applyCommand({
        type: "pptx:ungroup-shape",
        payload: { slideIndex: activeIndex, shapeId: target.id },
        source: "human",
      });
      setSelectedShapeIds(childIds);
    } catch (err) {
      onError(err);
    }
  }, [activeIndex, onError, selectedShapeIds]);

  const duplicateSelectedShapes: () => Promise<void> = useCallback(async () => {
    const a = agentRef.current;
    if (!a || selectedShapeIds.length === 0) return;
    try {
      const newIds: string[] = [];
      for (const id of selectedShapeIds) {
        await a.applyCommand({
          type: "pptx:duplicate-shape",
          payload: { slideIndex: activeIndex, shapeId: id },
          source: "human",
        });
        const slide = a.getSnapshot().root.slides[activeIndex];
        const last = slide.shapes[slide.shapes.length - 1];
        if (last) newIds.push(last.id);
      }
      if (newIds.length > 0) setSelectedShapeIds(newIds);
    } catch (err) {
      onError(err);
    }
  }, [activeIndex, onError, selectedShapeIds]);

  // Keep the refs in sync so the keyboard listener (bound once) always
  // calls the latest versions of the closures.
  useEffect(() => {
    duplicateRef.current = () => void duplicateSelectedShapes();
  }, [duplicateSelectedShapes]);
  useEffect(() => {
    groupRef.current = () => void groupSelectedShapes();
  }, [groupSelectedShapes]);
  useEffect(() => {
    ungroupRef.current = () => void ungroupSelectedShape();
  }, [ungroupSelectedShape]);

  const alignSelected = useCallback(
    async (
      mode: "left" | "center-h" | "right" | "top" | "middle-v" | "bottom",
      relativeTo: "selection" | "slide"
    ) => {
      const a = agentRef.current;
      if (!a) return;
      const minCount = relativeTo === "slide" ? 1 : 2;
      if (selectedShapeIds.length < minCount) return;
      try {
        await a.applyCommand({
          type: "pptx:align-shapes",
          payload: {
            slideIndex: activeIndex,
            shapeIds: selectedShapeIds,
            mode,
            relativeTo,
          },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError, selectedShapeIds]
  );

  const distributeSelected = useCallback(
    async (axis: "horizontal" | "vertical") => {
      const a = agentRef.current;
      if (!a || selectedShapeIds.length < 3) return;
      try {
        await a.applyCommand({
          type: "pptx:distribute-shapes",
          payload: { slideIndex: activeIndex, shapeIds: selectedShapeIds, axis },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError, selectedShapeIds]
  );

  // D5 — z-order. Operates on every selected shape sequentially so
  // multi-select "send to back" / "bring to front" produce the right
  // *relative* stacking even when shapes were originally at different
  // depths.
  const reorderSelected = useCallback(
    async (mode: import("@officeai/pptx").ReorderShapeMode) => {
      const a = agentRef.current;
      if (!a || selectedShapeIds.length === 0) return;
      const ids =
        mode === "to-front" || mode === "forward" ? selectedShapeIds : [...selectedShapeIds].reverse();
      try {
        for (const shapeId of ids) {
          await a.applyCommand({
            type: "pptx:reorder-shape",
            payload: { slideIndex: activeIndex, shapeId, mode },
            source: "human",
          });
        }
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError, selectedShapeIds]
  );

  /**
   * D11 — Animations panel command bridges. Each one is a thin
   * forwarder onto the existing `pptx:set-slide-transition`,
   * `pptx:add-shape-animation`, `pptx:remove-shape-animation`, and
   * `pptx:reorder-shape-animations` handlers. Add-animation requires
   * a selected shape because animations are anchored to a `cNvPrId`.
   */
  const setSlideTransition = useCallback(
    async (
      kind: import("@officeai/pptx").TransitionKind,
      speed: import("@officeai/pptx").TransitionSpeed | null
    ) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        // The "unsupported" landing-pad kind is only ever emitted by
        // the parser; the panel never offers it.
        if (kind === "unsupported") return;
        await a.applyCommand({
          type: "pptx:set-slide-transition",
          payload: {
            slideIndex: activeIndex,
            kind,
            ...(speed ? { speed } : {}),
          },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError]
  );

  /**
   * Apply a slide-size change. The {@link SlideSizeDialog} produces
   * either a preset (widescreen / standard / a4 / letter) or an
   * explicit width × height in EMU; both flows route through this
   * single helper so the splitter's preset shortcuts and the dialog
   * end up at the same handler.
   */
  const applySlideSize = useCallback(
    async (payload: { preset: SlideSizePreset; cxEmu?: number; cyEmu?: number }) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        await a.applyCommand({
          type: "pptx:set-slide-size",
          payload: {
            preset: payload.preset,
            ...(payload.cxEmu !== undefined ? { cxEmu: payload.cxEmu } : {}),
            ...(payload.cyEmu !== undefined ? { cyEmu: payload.cyEmu } : {}),
          },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [onError]
  );

  /**
   * Apply slideshow-wide options (Slideshow → Set Up Slide Show).
   * Dispatches the `pptx:set-show-options` handler, which serialises
   * the values into a `<p:showPr>` element on `presentation.xml`.
   */
  const applyShowOptions = useCallback(
    async (values: SetUpShowValues) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        await a.applyCommand({
          type: "pptx:set-show-options",
          payload: {
            showType: values.showType,
            loop: values.loop,
            showNarration: values.showNarration,
            showAnimation: values.showAnimation,
            useTimings: values.useTimings,
          },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [onError]
  );

  const clearShowOptions = useCallback(async () => {
    const a = agentRef.current;
    if (!a) return;
    try {
      await a.applyCommand({
        type: "pptx:set-show-options",
        payload: { clear: true },
        source: "human",
      });
    } catch (err) {
      onError(err);
    }
  }, [onError]);

  /**
   * Toggle the active slide's `show="0"` flag. Mirrors PowerPoint's
   * Slideshow ▸ Hide Slide button.
   */
  const toggleHideActiveSlide = useCallback(async () => {
    const a = agentRef.current;
    if (!a) return;
    const snap0 = a.getSnapshot();
    const slide0 = snap0.root.slides[activeIndex];
    if (!slide0) return;
    const currentlyHidden = slide0.slideRootAttrs.show === "0";
    try {
      await a.applyCommand({
        type: "pptx:set-slide-hidden",
        payload: { slideIndex: activeIndex, hidden: !currentlyHidden },
        source: "human",
      });
    } catch (err) {
      onError(err);
    }
  }, [activeIndex, onError]);

  const addShapeAnimation = useCallback(
    async (params: import("./AnimationsPanel").AddAnimationParams) => {
      const a = agentRef.current;
      if (!a || !selectedShapeId) return;
      try {
        await a.applyCommand({
          type: "pptx:add-shape-animation",
          payload: {
            slideIndex: activeIndex,
            shapeId: selectedShapeId,
            category: params.category,
            preset: params.preset,
            ...(params.direction ? { direction: params.direction } : {}),
            ...(params.trigger ? { trigger: params.trigger } : {}),
            ...(params.durationMs !== undefined ? { durationMs: params.durationMs } : {}),
            ...(params.delayMs !== undefined ? { delayMs: params.delayMs } : {}),
          },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError, selectedShapeId]
  );

  const setShapeAnimation = useCallback(
    async (params: import("./AnimationsPanel").SetAnimationParams) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        await a.applyCommand({
          type: "pptx:set-shape-animation",
          payload: {
            slideIndex: activeIndex,
            animationId: params.animationId,
            ...(params.category !== undefined ? { category: params.category } : {}),
            ...(params.preset !== undefined ? { preset: params.preset } : {}),
            ...(params.direction !== undefined ? { direction: params.direction } : {}),
            ...(params.trigger !== undefined ? { trigger: params.trigger } : {}),
            ...(params.durationMs !== undefined ? { durationMs: params.durationMs } : {}),
            ...(params.delayMs !== undefined ? { delayMs: params.delayMs } : {}),
          },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError]
  );

  const removeShapeAnimation = useCallback(
    async (animationId: string) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        await a.applyCommand({
          type: "pptx:remove-shape-animation",
          payload: { slideIndex: activeIndex, animationId },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError]
  );

  // Phase 4 — live "Preview" button. The renderer wraps every shape in
  // `<g class="anim-target" data-cnvprid="…">`, so the playback engine
  // can address shapes by walking the SVG inside `slideSurfaceRef`.
  // We hold one controller at a time and reset/destroy it before
  // starting a new preview so transforms never linger between runs.
  const playbackRef = useRef<PlaybackController | null>(null);
  // True when an in-editor "Play from beginning" preview is active.
  // While set, clicks on the slide surface advance the next build step
  // (matches PresentMode's click-to-advance) and Esc exits.
  const [animationStepMode, setAnimationStepMode] = useState(false);
  // Tracks how many groups remain in the active step-by-step preview
  // so we can auto-tear-down once the user has clicked through every
  // build step.
  const animationStepRemainingRef = useRef(0);
  useEffect(() => {
    return () => {
      playbackRef.current?.reset();
      playbackRef.current?.destroy();
      playbackRef.current = null;
    };
  }, []);

  // Reset step-mode whenever the user navigates to a different slide;
  // the controller belongs to the previous slide and would target the
  // wrong SVG.
  useEffect(() => {
    if (!animationStepMode) return;
    setAnimationStepMode(false);
    playbackRef.current?.reset();
    playbackRef.current?.destroy();
    playbackRef.current = null;
    animationStepRemainingRef.current = 0;
  }, [activeIndex]); // eslint-disable-line react-hooks/exhaustive-deps

  const exitAnimationStepMode = useCallback(() => {
    if (playbackRef.current) {
      playbackRef.current.reset();
      playbackRef.current.destroy();
      playbackRef.current = null;
    }
    animationStepRemainingRef.current = 0;
    setAnimationStepMode(false);
  }, []);

  const advanceAnimationStep = useCallback(async () => {
    const controller = playbackRef.current;
    if (!controller) return;
    try {
      await controller.clickAdvance();
    } catch {
      // controller torn down mid-advance — ignore.
    }
    animationStepRemainingRef.current -= 1;
    if (animationStepRemainingRef.current <= 0) {
      exitAnimationStepMode();
    }
  }, [exitAnimationStepMode]);

  // Esc bails out of step mode the same way it does in PresentMode.
  useEffect(() => {
    if (!animationStepMode) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        e.preventDefault();
        exitAnimationStepMode();
      }
    };
    window.addEventListener("keydown", onKey, true);
    return () => window.removeEventListener("keydown", onKey, true);
  }, [animationStepMode, exitAnimationStepMode]);

  const previewAnimation = useCallback(
    (animationId: string | null) => {
      const surface = slideSurfaceRef.current;
      if (!surface || !snap) return;
      const slide = snap.root.slides[activeIndex];
      if (!slide) return;
      const svg = surface.querySelector<SVGSVGElement>("svg");
      if (!svg) return;
      // Tear down any in-flight controller so we don't stack animations
      // on top of each other when the user spams the preview button.
      playbackRef.current?.reset();
      playbackRef.current?.destroy();
      animationStepRemainingRef.current = 0;
      setAnimationStepMode(false);
      const controller = createPlayback(svg, slide, {
        slideSize: snap.root.slideSize,
      });
      playbackRef.current = controller;
      controller.prepare();

      // "Play from beginning" (animationId === null): enter step-by-
      // step mode so the user clicks through builds the same way they
      // would in PresentMode. We pre-compute group count so we know
      // when to auto-exit. This matches the existing PresentMode fix
      // and stops the editor preview from being a one-shot blur.
      if (!animationId) {
        const groups = countAnimationGroups(slide);
        if (groups > 0) {
          animationStepRemainingRef.current = groups;
          setAnimationStepMode(true);
          // Show the baseline (entrance shapes hidden) so the very
          // first click reveals the first build step.
          return;
        }
        // No animations to step through — just tear down silently.
        controller.reset();
        controller.destroy();
        playbackRef.current = null;
        return;
      }

      const run = async () => {
        try {
          if (animationId) {
            // Find the click group that contains the requested step and
            // fast-forward to it. We walk the same grouping the engine
            // uses so the indices stay aligned.
            const ordered = [...slide.animations].sort((a, b) => a.order - b.order);
            const groups: Array<Array<(typeof ordered)[number]>> = [];
            let current: Array<(typeof ordered)[number]> = [];
            for (const a of ordered) {
              const startsGroup = current.length === 0 || a.trigger === "onClick";
              if (startsGroup && current.length > 0) {
                groups.push(current);
                current = [];
              }
              current.push(a);
            }
            if (current.length > 0) groups.push(current);
            const targetGroup = groups.findIndex((g) => g.some((a) => a.id === animationId));
            for (let i = 0; i < groups.length; i++) {
              if (i === targetGroup) {
                await controller.clickAdvance();
                break;
              }
              await controller.clickAdvance();
            }
          } else {
            await controller.playAll();
          }
        } catch {
          // Swallow — the user may have switched slides or the
          // controller was destroyed mid-playback. The reset() in the
          // next preview call (or unmount cleanup) restores state.
        } finally {
          // Restore the slide to its baseline so the editor doesn't get
          // stuck in a half-played state once the preview finishes.
          if (playbackRef.current === controller) {
            controller.reset();
            controller.destroy();
            playbackRef.current = null;
          }
        }
      };
      void run();
    },
    [activeIndex, snap]
  );

  const reorderShapeAnimations = useCallback(
    async (orderIds: ReadonlyArray<string>) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        await a.applyCommand({
          type: "pptx:reorder-shape-animations",
          payload: { slideIndex: activeIndex, order: orderIds },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError]
  );

  const setSlideNotes = useCallback(
    async (text: string) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        await a.applyCommand({
          type: "pptx:set-slide-notes",
          payload: { slideIndex: activeIndex, text },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError]
  );

  const changeFill = useCallback(
    async (next: FillSpec | null) => {
      const a = agentRef.current;
      if (!a || !selectedShapeId) return;
      try {
        // The handler accepts both `string | null` (legacy) and full
        // `FillSpec`; we always pass the typed form so gradient /
        // pattern / picture work end-to-end. `null` clears the explicit
        // fill so the shape inherits its style — equivalent to
        // PowerPoint's "Fill: Automatic".
        await a.applyCommand({
          type: "pptx:set-shape-fill",
          payload: {
            slideIndex: activeIndex,
            shapeId: selectedShapeId,
            fill: next ?? null,
          },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError, selectedShapeId]
  );

  const changeSlideBackground = useCallback(
    async (next: FillSpec | null) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        await a.applyCommand({
          type: "pptx:set-slide-background",
          payload: { slideIndex: activeIndex, fill: next },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError]
  );

  // Dispatcher behind the floating geometry bar's `<input type="range">`.
  // One command per adjustment (mirrors PowerPoint's yellow-handle drag
  // committing a single `<a:gd>` write at mouseup) so the undo stack
  // stays one-entry-per-tweak.
  const setShapeGeometry = useCallback(
    async (shapeId: string, adjName: string, value: number) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        await a.applyCommand({
          type: "pptx:set-shape-geometry",
          payload: { slideIndex: activeIndex, shapeId, adjName, value },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError]
  );

  const setTextAlignment = useCallback(
    async (alignment: "left" | "center" | "right" | "justify" | null) => {
      const a = agentRef.current;
      if (!a) return;
      const target = activeTextShape;
      if (!target) {
        pushToast("info", "Select a text shape first.");
        return;
      }
      // PowerPoint's per-paragraph H-alignment: when a text-edit caret
      // is open we just align the paragraph it's parked in; otherwise
      // we align every paragraph in the shape (matches what users get
      // when they click a shape and hit Align without opening edit
      // mode).
      const paragraphs = textSelection ? [textSelection.paragraph] : undefined;
      try {
        await a.applyCommand({
          type: "pptx:set-paragraph-alignment",
          payload: {
            slideIndex: activeIndex,
            shapeId: target.id,
            alignment,
            ...(paragraphs ? { paragraphs } : {}),
          },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, activeTextShape, onError, pushToast, textSelection]
  );

  const setTextAnchorAction = useCallback(
    async (anchor: TextAnchor | null) => {
      const a = agentRef.current;
      if (!a) return;
      const target = activeTextShape;
      if (!target) {
        pushToast("info", "Select a text shape first.");
        return;
      }
      try {
        await a.applyCommand({
          type: "pptx:set-text-anchor",
          payload: { slideIndex: activeIndex, shapeId: target.id, anchor },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, activeTextShape, onError, pushToast]
  );

  // Toolbar "Add comment" entry point. The shared right-rail owns the
  // comments panel + composer; this just opens it and focuses the
  // textarea via a custom event the rail listens for.
  const focusCommentComposer = useCallback(() => {
    requestAnimationFrame(() => {
      const root = document.querySelector<HTMLElement>('[data-testid="rail-tab-comments"]');
      if (root) (root as HTMLButtonElement).click();
      requestAnimationFrame(() => {
        const textarea = document.querySelector<HTMLTextAreaElement>(
          'aside [data-testid="comment-composer"] textarea'
        );
        textarea?.focus();
      });
    });
  }, []);

  // "Click to locate" handler wired into the comments sidebar. Looks
  // up the comment in the active slide's comments part, switches the
  // canvas selection to the anchored shape (if any) and triggers a
  // yellow flash overlay so the user can find it. Falls back to a
  // pin-position flash for free-floating comments.
  const scrollToComment = useCallback(
    (commentId: string) => {
      const a = agentRef.current;
      if (!a) return;
      const snap = a.getSnapshot();
      const target = snap.root.slides[activeIndex];
      if (!target || !target.commentsPartPath) return;
      const part = snap.root.commentsByPart.get(target.commentsPartPath);
      if (!part) return;
      const c = part.comments.find((cm) => cm.id === commentId);
      if (!c) return;
      const surface = slideSurfaceRef.current;
      if (surface) surface.scrollIntoView({ block: "nearest", behavior: "smooth" });
      if (c.shapeId) {
        setSelectedShapeIds([c.shapeId]);
        setCommentFlashTarget((prev) => ({
          kind: "shape",
          shapeId: c.shapeId!,
          nonce: (prev?.nonce ?? 0) + 1,
        }));
      } else {
        setCommentFlashTarget((prev) => ({
          kind: "pin",
          xEmu: c.xEmu,
          yEmu: c.yEmu,
          nonce: (prev?.nonce ?? 0) + 1,
        }));
      }
    },
    [activeIndex]
  );

  // Auto-clear the flash overlay after the animation completes so we
  // don't keep stale SVG nodes in the tree.
  useEffect(() => {
    if (!commentFlashTarget) return;
    const handle = window.setTimeout(() => setCommentFlashTarget(null), 1700);
    return () => window.clearTimeout(handle);
  }, [commentFlashTarget]);

  usePptxShortcuts({
    surfaceRef: slideSurfaceRef,
    agentRef,
    activeIndex,
    slideCount: slides.length,
    selectedShape,
    selectedShapeIds,
    textFormatProvider,
    onAddSlide: () => void addSlide(),
    onDuplicateSlide: () => void duplicateSlide(),
    onDeleteShape: () => void deleteSelectedShape(),
    onChangeSlide: setActiveIndex,
    onError,
  });

  /**
   * D10 — global F5 / Shift+F5 to enter Present mode. F5 starts from
   * slide 1, Shift+F5 starts from the currently active slide. We
   * consume the event so the browser's native F5 (page reload) is
   * suppressed while a presentation is loaded — matching PowerPoint.
   * The listener is gated on `ready` so it never fires before a
   * snapshot is available.
   */
  const startPresenting = useCallback(
    (fromCurrent: boolean) => {
      if (!ready) return;
      if (!fromCurrent) setActiveIndex(0);
      setPresenting(true);
    },
    [ready]
  );
  useEffect(() => {
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === "F5") {
        const target = e.target as HTMLElement | null;
        if (
          target &&
          (target.tagName === "INPUT" || target.tagName === "TEXTAREA" || target.isContentEditable)
        ) {
          return;
        }
        e.preventDefault();
        startPresenting(e.shiftKey);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [startPresenting]);

  // Build object-URL map for every embedded media part so the renderer
  // can paint <Picture> shapes inserted via the toolbar. We rebuild on
  // every snapshot tick — cheap because URLs are deduplicated by part path.
  const mediaUrls = useMemo(() => {
    const map = new Map<string, string>();
    if (!snap) return map;
    for (const [path, part] of snap.root.media) {
      const blob = new Blob([part.bytes as BlobPart], { type: part.contentType });
      map.set(path, URL.createObjectURL(blob));
    }
    return map;
  }, [snap]);
  useEffect(() => {
    return () => {
      for (const url of mediaUrls.values()) URL.revokeObjectURL(url);
    };
  }, [mediaUrls]);

  // Comment-thread count for the right rail badge.
  const openCommentCount = useMemo(() => {
    if (!snap || !slide || !slide.commentsPartPath) return 0;
    const part = snap.root.commentsByPart.get(slide.commentsPartPath);
    if (!part) return 0;
    return part.comments.filter((c) => !c.parentId && !c.resolved).length;
  }, [snap, slide]);

  const renderCommentsPanel = useCallback((): React.ReactNode => {
    if (!agent || !snap || !slide) {
      return <div className="p-4 text-sm text-secondary">No slide selected.</div>;
    }
    return (
      <div className="flex min-h-0 flex-1 flex-col">
        <div className="min-h-0 flex-1 overflow-y-auto">
          <CommentsSidebar
            key={`comments-${activeIndex}-${tick}`}
            provider={createPptxCommentsProvider({
              agent,
              slideIndex: activeIndex,
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
            emptyHint="No comments on this slide yet. Select a shape and press Add comment, or comment on the slide as a whole."
            onScrollTo={scrollToComment}
          />
        </div>
        <div className="border-t border-divider p-2">
          <CommentComposer
            provider={createPptxCommentsProvider({ agent, slideIndex: activeIndex })}
            anchor={
              selectedShape && selectedShape.position
                ? {
                    kind: "pptx-pin",
                    slideIndex: activeIndex,
                    xEmu: selectedShape.position.xEmu + Math.round((selectedShape.size?.cxEmu ?? 0) / 2),
                    yEmu: selectedShape.position.yEmu + Math.round((selectedShape.size?.cyEmu ?? 0) / 2),
                    shapeId: selectedShape.id,
                  }
                : {
                    kind: "pptx-pin",
                    slideIndex: activeIndex,
                    xEmu: Math.round(slideSize.cxEmu / 2),
                    yEmu: Math.round(slideSize.cyEmu / 2),
                  }
            }
            placeholder={
              selectedShape
                ? `Comment on ${selectedShape.name || "selected shape"}…`
                : "Add a comment to this slide…"
            }
          />
        </div>
      </div>
    );
  }, [
    activeIndex,
    agent,
    scrollToComment,
    selectedShape,
    slide,
    slideSize.cxEmu,
    slideSize.cyEmu,
    snap,
    tick,
  ]);

  // Selection summary centred in the status bar. Slide index lives in
  // the left column (`PptxSelectionHint`); this slot only names the
  // current selection to avoid duplicating "Slide n of m" in two languages.
  const selectionText = useMemo(() => {
    if (!ready) return "";
    if (selectedShapeIds.length === 0) {
      return "";
    }
    if (selectedShapeIds.length === 1) {
      const name = (selectedShape?.name ?? "").trim() || t("status.unnamedShape");
      return t("status.oneShapeSelected", { name });
    }
    return t("status.shapesSelected", { n: selectedShapeIds.length });
  }, [ready, selectedShape?.name, selectedShapeIds.length, t]);

  // Palette is generated from the central pptx action catalogue (see
  // packages/pptx/src/actions/catalogue.ts). Labels, sections, and
  // shortcut hints flow from the catalogue; this map only carries the
  // closure-bound side effects + per-id `enabled` gating.
  const paletteCommands = useMemo<ReadonlyArray<PaletteCommand>>(() => {
    const runners: PaletteRunners = {
      "pptx.add-slide": { run: () => void addSlide() },
      "pptx.duplicate-slide": { run: () => void duplicateSlide() },
      "pptx.delete-slide": { run: () => void deleteSlide(), enabled: slides.length > 1 },
      "pptx.add-text-box": { run: () => void addTextBox() },
      "pptx.insert-image": {
        run: () => {
          void (async () => {
            const file = await pickImageFile();
            if (file) await insertImage(file);
          })();
        },
      },
      "pptx.insert-table-from-xlsx": { run: () => setXlsxPickerOpen("materialized") },
      "pptx.insert-spreadsheet-from-xlsx": { run: () => setXlsxPickerOpen("live") },
      "pptx.insert-chart-from-xlsx": { run: () => setXlsxPickerOpen("chart") },
      "pptx.add-rect": { run: () => void addShape("rect") },
      "pptx.add-ellipse": { run: () => void addShape("ellipse") },
      "pptx.add-arrow": { run: () => void addShape("rightArrow") },
      "pptx.add-connector-elbow": { run: () => startConnectorTool("elbow") },
      "pptx.add-connector-straight": { run: () => startConnectorTool("straight") },
      "pptx.add-connector-curved": { run: () => startConnectorTool("curved") },
      "pptx.add-comment": { run: () => focusCommentComposer() },
      "pptx.delete-shape": { run: () => void deleteSelectedShape(), enabled: selectedShapeIds.length > 0 },
      "pptx.duplicate-shape": {
        run: () => void duplicateSelectedShapes(),
        enabled: selectedShapeIds.length > 0,
      },
      "pptx.group-shapes": { run: () => void groupSelectedShapes(), enabled: selectedShapeIds.length >= 2 },
      "pptx.ungroup-shape": {
        run: () => void ungroupSelectedShape(),
        enabled: selectedShapeIds.length === 1,
      },
      "pptx.zoom-reset": { run: () => setZoom(1) },
      "pptx.present-from-start": { run: () => startPresenting(false), enabled: ready && slides.length > 0 },
      "pptx.present-from-current": { run: () => startPresenting(true), enabled: ready && slides.length > 0 },
      // Animations entries reveal the right-rail Animations tab so the
      // user lands on the panel that has the full preset gallery + the
      // per-step Effect Options. Hard-wiring a Cmd+K command to a
      // single (category, preset) tuple would balloon the palette to
      // 30+ entries; the panel keeps the choice space discoverable
      // without burying the rest of the palette.
      "pptx.set-slide-transition": {
        run: () => requestRail("animations"),
        enabled: ready && slides.length > 0,
      },
      "pptx.set-slide-size": {
        run: () => setSlideSizeDialogOpen(true),
        enabled: ready,
      },
      "pptx.set-show-options": {
        run: () => setSetUpShowDialogOpen(true),
        enabled: ready,
      },
      "pptx.set-slide-hidden": {
        run: () => void toggleHideActiveSlide(),
        enabled: ready && slides.length > 0,
      },
      "pptx.add-shape-animation": {
        run: () => requestRail("animations"),
        enabled: ready && slides.length > 0,
      },
      "pptx.remove-shape-animation": {
        run: () => requestRail("animations"),
        enabled: ready && (slide?.animations.length ?? 0) > 0,
      },
      "pptx.reorder-shape-animations": {
        run: () => requestRail("animations"),
        enabled: ready && (slide?.animations.length ?? 0) >= 2,
      },
    };
    return buildPaletteFromCatalogue(pptxActions, runners, t);
  }, [
    t,
    startConnectorTool,
    addShape,
    addSlide,
    addTextBox,
    deleteSelectedShape,
    deleteSlide,
    duplicateSelectedShapes,
    duplicateSlide,
    focusCommentComposer,
    groupSelectedShapes,
    insertImage,
    ready,
    requestRail,
    selectedShapeIds.length,
    slide?.animations.length,
    slides.length,
    startPresenting,
    toggleHideActiveSlide,
    ungroupSelectedShape,
  ]);

  const tabFallback = useStableTabId("pptx");
  const realtimeRoomId = useMemo<string | null>(() => {
    if (!ready) return null;
    if (roomOverride === null) return null;
    if (typeof roomOverride === "string" && roomOverride.length > 0) {
      return `oai/pptx/host/${roomOverride}`;
    }
    if (!tabFallback && !initialSource) return null;
    return roomIdForSource({
      product: "pptx",
      src: initialSource?.url,
      tabFallback,
      explicitRoom: readExplicitRoomFromUrl(),
    });
  }, [ready, initialSource, tabFallback, roomOverride]);
  const realtimeRoom = useRealtimeRoom({
    roomId: realtimeRoomId,
    product: "pptx",
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

  // Publish PPTX selection (active slide + currently-selected
  // shapes) so peers see "Quick Quokka has 2 shapes selected on
  // slide 3" in real time.
  //
  // We deliberately publish *stable* OOXML identifiers (`partPath`
  // for the slide, `cNvPrId` for shapes) instead of the randomly-
  // minted local `NodeId`s. Two browsers parsing the same .pptx mint
  // independent UUIDs, so the local NodeId never matches across
  // peers — that's why the slide-rail dots and on-canvas selection
  // outlines used to silently miss every match.
  const presenceCursor = useMemo(() => {
    if (!slide) return null;
    const cNvIds: string[] = [];
    for (const localId of selectedShapeIds) {
      const sh = findShape(slide.shapes, localId);
      if (sh && sh.cNvPrId > 0) cNvIds.push(String(sh.cNvPrId));
    }
    return {
      product: "pptx" as const,
      slideId: slide.partPath,
      shapeIds: cNvIds,
    };
  }, [slide, selectedShapeIds]);
  usePublishPresence({ room: realtimeRoom.room, cursor: presenceCursor });

  // PPTX outline = each slide's title placeholder text (when the
  // slide layout exposes a `title` placeholder), with `level: 1`
  // throughout — PowerPoint's outline view is flat, not hierarchical.
  // Slides without a title placeholder fall back to "Slide N" so the
  // rail never strands the user on an unnamed slide.
  const outline = useMemo(() => {
    if (!snap) return undefined;
    return snap.root.slides.map((slide, i) => {
      const title = readSlideTitle(slide);
      return {
        id: slide.partPath,
        level: 1,
        text: title || `Slide ${i + 1}`,
        active: i === activeIndex,
        onActivate: () => setActiveIndex(i),
      };
    });
  }, [snap, activeIndex]);

  const adapter = useMemo<ProductAdapter>(
    () => ({
      product: "pptx",
      filename: docName,
      saveState,
      comments: { openCount: openCommentCount, resolvedCount: 0 },
      ...(outline ? { outline } : {}),
      selectionSummary: { text: selectionText },
      canOpen: true,
      hideOpen: hideLocalFileOpen,
      canSave: ready,
      canExport: ready,
      exportFormats: PPTX_EXPORT_FORMATS,
      onOpenFile: () => void handleOpenFile(),
      onSave: () => handleSave(),
      onExport: (format, options) => handleExport(format, options),
      canUndo: ready ? (agentRef.current?.canUndo() ?? false) : false,
      canRedo: ready ? (agentRef.current?.canRedo() ?? false) : false,
      onUndo: () => {
        const a = agentRef.current;
        if (!a || !a.canUndo()) return;
        a.undo();
      },
      onRedo: () => {
        const a = agentRef.current;
        if (!a || !a.canRedo()) return;
        a.redo();
      },
      onOpenShortcuts: () => shortcutsDialog.setOpen(true),
      paletteCommands,
      renderCommentsPanel,
      renderAnimationsPanel: snap
        ? () => (
            <AnimationsPanel
              snapshot={snap}
              activeIndex={activeIndex}
              selectedShape={selectedShape}
              disabled={!ready}
              onSetTransition={(kind, speed) => void setSlideTransition(kind, speed)}
              onAddAnimation={(params) => void addShapeAnimation(params)}
              onSetAnimation={(params) => void setShapeAnimation(params)}
              onRemoveAnimation={(id) => void removeShapeAnimation(id)}
              onReorderAnimations={(orderIds) => void reorderShapeAnimations(orderIds)}
              onPreviewAnimation={(id) => previewAnimation(id)}
            />
          )
        : undefined,
      renderMasterPanel: snap
        ? () => (
            <MasterPanel
              snapshot={snap}
              defaultLayoutKind={defaultLayoutKind}
              onChangeDefaultLayout={setDefaultLayoutKind}
            />
          )
        : undefined,
      onAddComment: focusCommentComposer,
    }),
    [
      activeIndex,
      addShapeAnimation,
      defaultLayoutKind,
      docName,
      focusCommentComposer,
      handleExport,
      handleOpenFile,
      handleSave,
      hideLocalFileOpen,
      openCommentCount,
      outline,
      paletteCommands,
      previewAnimation,
      ready,
      removeShapeAnimation,
      renderCommentsPanel,
      reorderShapeAnimations,
      saveState,
      selectedShape,
      selectionText,
      setShapeAnimation,
      setSlideTransition,
      shortcutsDialog,
      snap,
      tick,
    ]
  );

  const remotePptxPeers: ReadonlyArray<RemoteSelectionPeer> = useMemo(
    () =>
      realtimeRoom.remotePeers
        .map((p): RemoteSelectionPeer | null => {
          const c = p.state.cursor;
          if (!c || c.product !== "pptx") return null;
          return {
            clientId: p.clientId,
            slideId: c.slideId,
            shapeIds: c.shapeIds,
            name: p.state.user.name,
            color: p.state.user.color,
          };
        })
        .filter((x): x is RemoteSelectionPeer => x !== null),
    [realtimeRoom.remotePeers]
  );
  const slideRailPeers: ReadonlyArray<SlideRailPeerDot> = useMemo(
    () =>
      remotePptxPeers.map((p) => ({
        clientId: p.clientId,
        slideId: p.slideId,
        name: p.name,
        color: p.color,
      })),
    [remotePptxPeers]
  );

  return (
    <>
      <RemotePresenceList peers={realtimeRoom.remotePeers} />
      <EditorShell
        adapter={adapter}
        onBack={onCloseProp}
        topBarExtras={<PresenceSlot state={realtimeRoom} />}
        requestRailTab={railRequest ?? undefined}
        toolbar={
          <PptxToolbar
            disabled={!ready}
            slideCount={slides.length}
            hasSelection={selectedShapeId != null}
            selectionCount={selectedShapeIds.length}
            currentShapeFill={currentShapeFill}
            currentSlideBackground={currentSlideBackground}
            textFormatProvider={textFormatProvider}
            textFormatActive={textFormatActive}
            onAddSlide={() => void addSlide()}
            onAddSlideWithLayout={(k) => void addSlideWithLayout(k)}
            onSetSlideLayout={(k) => void setSlideLayout(k)}
            onDeleteSlide={() => void deleteSlide()}
            onDuplicateSlide={() => void duplicateSlide()}
            onAddTextBox={() => void addTextBox()}
            onAddShape={(p) => void addShape(p)}
            onAddConnector={(t) => startConnectorTool(t)}
            connectorToolType={connectorTool?.type ?? null}
            onInsertImage={(f) => void insertImage(f)}
            onInsertMedia={(f) => void insertMedia(f)}
            onInsertFromXlsx={() => setXlsxPickerOpen("materialized")}
            onReplacePicture={(f) => void replaceSelectedPicture(f)}
            selectedIsPicture={selectedShape?.kind === "pic"}
            onDeleteShape={() => void deleteSelectedShape()}
            onAlign={(mode, relativeTo) => void alignSelected(mode, relativeTo)}
            onDistribute={(axis) => void distributeSelected(axis)}
            onReorder={(mode) => void reorderSelected(mode)}
            onGroup={() => void groupSelectedShapes()}
            onUngroup={() => void ungroupSelectedShape()}
            onDuplicateShape={() => void duplicateSelectedShapes()}
            canGroup={selectedShapeIds.length >= 2}
            canUngroup={selectedShapeIds.length === 1 && selectedShape?.kind === "group"}
            onChangeShapeFill={(spec) => void changeFill(spec)}
            onChangeSlideBackground={(spec) => void changeSlideBackground(spec)}
            onSetTextAlignment={(alignment) => void setTextAlignment(alignment)}
            onSetTextAnchor={(anchor) => void setTextAnchorAction(anchor)}
            activeTextAlignment={activeTextAlignment}
            activeTextAnchor={activeTextAnchor}
            hasTextShapeFocus={activeTextShape != null}
            onAddComment={focusCommentComposer}
            onPresent={() => startPresenting(true)}
            canPresent={ready && slides.length > 0}
            onToggleNotes={() => setNotesOpen((v) => !v)}
            notesOpen={notesOpen}
            onToggleRulers={() => setRulersVisible((v) => !v)}
            rulersVisible={rulersVisible}
            onToggleGrid={() => setGridVisible((v) => !v)}
            gridVisible={gridVisible}
            currentTransitionKind={slides[activeIndex]?.transition?.kind ?? "none"}
            currentTransitionSpeed={slides[activeIndex]?.transition?.speed ?? null}
            onSetSlideTransition={(kind, speed) => void setSlideTransition(kind, speed)}
            onOpenRail={(tab) => requestRail(tab)}
            onOpenSlideSize={() => setSlideSizeDialogOpen(true)}
            onApplySlideSizePreset={(preset) => void applySlideSize({ preset })}
            onOpenSetUpShow={() => setSetUpShowDialogOpen(true)}
            onToggleHideSlide={() => void toggleHideActiveSlide()}
            activeSlideHidden={activeSlideHidden}
          />
        }
        statusBarLeft={
          <PptxSelectionHint
            selectedCount={selectedShapeIds.length}
            slideIndex={activeIndex}
            slideCount={slides.length}
          />
        }
        statusBarRight={
          <ZoomControl
            value={zoom}
            minZoom={MIN_ZOOM}
            maxZoom={MAX_ZOOM}
            onChange={onZoomChange}
            disabled={!ready}
          />
        }
        body={
          <div className="pptx-editor flex min-h-0 flex-1 flex-col gap-2 p-3">
            {!agent ? (
              <EmptyState product="pptx" onOpen={() => void handleOpenFile()} />
            ) : (
              <>
                <div className="relative flex min-h-0 flex-1 gap-3">
                  <aside
                    data-testid="pptx-sidebar"
                    className="hidden w-[200px] shrink-0 overflow-y-auto rounded-md border border-divider bg-surface md:block"
                  >
                    {snap ? (
                      <SlidesSidebar
                        slides={slides}
                        slideSize={slideSize}
                        mediaUrls={mediaUrls}
                        theme={themeDefault}
                        activeIndex={activeIndex}
                        onSelect={setActiveIndex}
                        thumbnailWidth={170}
                        onReorder={(from, to) => void moveSlide(from, to)}
                        onContextAction={(idx, action) => void handleSlideContextAction(idx, action)}
                        peers={slideRailPeers}
                      />
                    ) : null}
                  </aside>
                  <section
                    ref={slideSurfaceRef as React.RefObject<HTMLElement>}
                    tabIndex={-1}
                    data-testid="pptx-slide-surface"
                    className="relative flex min-h-0 flex-1 overflow-hidden rounded-md border border-divider"
                    // The slide surface receives programmatic focus
                    // on pointer-down so keyboard shortcuts work
                    // (see onPointerDown below). It's not a real
                    // interactive control though — the actual focus
                    // affordance is the selection chrome painted by
                    // the canvas — so we suppress the global
                    // `:focus-visible` accent outline here. Done via
                    // inline `outline: none` rather than Tailwind so
                    // it wins against the cascade unconditionally.
                    style={{ backgroundColor: "var(--page-backdrop)", outline: "none" }}
                    onPointerDown={(e) => {
                      // Focus the surface on pointer interaction so
                      // keyboard shortcuts (Delete / arrows / PageUp /
                      // PageDown …) fire after clicking a shape. The
                      // section has tabIndex=-1 so neither the SVG
                      // shape children nor the section itself receive
                      // focus on click — without this, key events go
                      // to <body> and `usePptxShortcuts` skips them.
                      // Skip when the click landed in something with
                      // its own caret (the in-place text editor, a
                      // floating input, …) so we don't yank focus
                      // away from it.
                      const t = e.target as HTMLElement | null;
                      if (
                        t &&
                        (t.isContentEditable ||
                          t.closest('[contenteditable="true"], input, textarea, select') !== null)
                      ) {
                        return;
                      }
                      slideSurfaceRef.current?.focus({ preventScroll: true });
                      // Animation step-by-step preview: every pointer
                      // click advances to the next build group (parity
                      // with PresentMode). We listen on pointerdown so
                      // it fires before SlideCanvas's selection
                      // handlers, and stopPropagation so a click
                      // doesn't simultaneously change selection
                      // mid-preview.
                      if (animationStepMode) {
                        e.stopPropagation();
                        void advanceAnimationStep();
                      }
                    }}
                  >
                    {animationStepMode ? (
                      <div
                        data-testid="pptx-anim-step-hint"
                        className="pointer-events-none absolute left-1/2 top-2 z-30 -translate-x-1/2 rounded-full bg-foreground/85 px-3 py-1 text-[11px] font-medium text-background shadow"
                      >
                        Click to advance · Esc to exit
                      </div>
                    ) : null}
                    <div className="relative min-h-0 w-full flex-1">
                      <SlideCanvas
                        agent={agent}
                        slideIndex={activeIndex}
                        mediaUrls={mediaUrls}
                        onError={onError}
                        zoom={zoom}
                        onSelectionChange={setSelectedShapeIds}
                        onTextSelectionChange={setTextSelection}
                        selectedShapeIds={selectedShapeIds}
                        commentedShapeIds={commentedShapeIds}
                        commentFlashTarget={commentFlashTarget}
                        onPlaceholderActivate={(info) => void handlePlaceholderActivate(info)}
                        onShapeActivate={(info) => void handleShapeActivate(info)}
                        connectorTool={connectorTool}
                        onConnectorToolExit={exitConnectorTool}
                        remotePeers={remotePptxPeers}
                        showRulers={rulersVisible}
                        showGrid={gridVisible}
                        rulerUnit={rulerUnit}
                      />
                      {selectedShape &&
                      selectedShape.kind === "connector" &&
                      selectedShapeIds.length === 1 ? (
                        <div className="pointer-events-auto absolute left-1/2 top-2 z-20 -translate-x-1/2">
                          <ConnectorContextBar
                            connector={selectedShape}
                            onPatch={(patch) => void applyConnectorStylePatch(selectedShape.id, patch)}
                            onAction={(action) => void applyConnectorAction(selectedShape.id, action)}
                          />
                        </div>
                      ) : null}
                      {selectedShape &&
                      selectedShape.kind === "text" &&
                      selectedShapeIds.length === 1 &&
                      shapeHasAdjustableGeometry(selectedShape) ? (
                        <div className="pointer-events-auto absolute left-1/2 top-2 z-20 -translate-x-1/2">
                          <ShapeGeometryContextBar
                            shape={selectedShape}
                            onChange={(adjName, value) =>
                              void setShapeGeometry(selectedShape.id, adjName, value)
                            }
                          />
                        </div>
                      ) : null}
                    </div>
                  </section>
                </div>
                {snap && notesOpen ? (
                  <NotesPanel
                    key={`notes-${activeIndex}-${tick}`}
                    notesText={readNotesText(snap, activeIndex)}
                    disabled={!ready}
                    onChange={(t) => void setSlideNotes(t)}
                  />
                ) : null}
              </>
            )}
          </div>
        }
        toasts={toasts}
        onDismissToast={dismissToast}
        onFileDrop={(file) => void handleFile(file)}
        dropExtension=".pptx"
        onRenameFilename={(next) => setDocName(next)}
      />
      <KeyboardShortcutsDialog
        product="pptx"
        open={shortcutsDialog.open}
        onClose={() => shortcutsDialog.setOpen(false)}
      />
      <XlsxRangePickerDialog
        open={xlsxPickerOpen !== null}
        defaultMode={xlsxPickerOpen ?? "materialized"}
        onCancel={() => setXlsxPickerOpen(null)}
        onSubmit={(result) => void handleXlsxPickerSubmit(result)}
      />
      <SlideSizeDialog
        open={slideSizeDialogOpen}
        currentCxEmu={slideSize.cxEmu}
        currentCyEmu={slideSize.cyEmu}
        onClose={() => setSlideSizeDialogOpen(false)}
        onSubmit={(payload) =>
          void applySlideSize({
            preset: payload.preset,
            cxEmu: payload.cxEmu,
            cyEmu: payload.cyEmu,
          })
        }
      />
      <SetUpShowDialog
        open={setUpShowDialogOpen}
        current={currentShowOptions}
        onClose={() => setSetUpShowDialogOpen(false)}
        onSubmit={(values) => void applyShowOptions(values)}
        onClear={() => void clearShowOptions()}
      />
      <EmbeddedXlsxModal
        open={editingEmbed !== null}
        bytes={editingEmbed?.bytes ?? null}
        title={editingEmbed?.title}
        onCancel={() => setEditingEmbed(null)}
        onSave={(r) => void handleEmbeddedXlsxSave(r)}
      />
      {presenting && snap && agent ? (
        <PresentMode
          snapshot={snap}
          subscribeSnapshot={{
            getSnapshot: () => agent.getSnapshot(),
            subscribe: (listener) => agent.subscribe(() => listener()),
          }}
          initialSlideIndex={activeIndex}
          mediaUrls={mediaUrls}
          charts={snap.root.charts}
          onClose={() => setPresenting(false)}
        />
      ) : null}
    </>
  );
}

/**
 * Walk `presentationOpaqueTail` looking for the verbatim `<p:showPr>`
 * element so the SetUpShow dialog opens with the deck's current
 * settings. PowerPoint defaults are applied when the element is
 * missing, the deck never specified the flag, or `snap` is null
 * (early-mount path before the first agent snapshot exists).
 */
function readShowOptions(snap: PptxSnapshot | null): SetUpShowValues {
  const defaults: SetUpShowValues = {
    showType: "presenter",
    loop: false,
    showNarration: true,
    showAnimation: true,
    useTimings: true,
  };
  if (!snap) return defaults;
  const tail = snap.root.presentationOpaqueTail;
  for (const el of tail) {
    if (el.tag === "p:showPr") {
      const attrs = el.attrs;
      let showType: SetUpShowValues["showType"] = "presenter";
      for (const child of el.subtree) {
        if (typeof child !== "object" || child === null) continue;
        for (const key of Object.keys(child as Record<string, unknown>)) {
          if (key === "p:browse") showType = "browse";
          else if (key === "p:kiosk") showType = "kiosk";
        }
      }
      return {
        showType,
        loop: attrs.loop === "1" || attrs.loop === "true",
        showNarration: attrs.showNarration !== "0" && attrs.showNarration !== "false",
        showAnimation: attrs.showAnimation !== "0" && attrs.showAnimation !== "false",
        useTimings: attrs.useTimings !== "0" && attrs.useTimings !== "false",
      };
    }
  }
  return defaults;
}

/** Count the number of click-groups in a slide's animation list. A
 * "click-group" is the unit advanced per click in PresentMode: every
 * `onClick`-triggered animation starts a new group, the rest chain
 * with/after-prev. Used by the editor preview's step-by-step mode to
 * decide when to auto-tear-down. */
function countAnimationGroups(slide: import("@officeai/pptx").Slide): number {
  if (slide.animations.length === 0) return 0;
  const ordered = [...slide.animations].sort((a, b) => a.order - b.order);
  let groups = 0;
  let inGroup = false;
  for (const a of ordered) {
    if (!inGroup || a.trigger === "onClick") {
      groups += 1;
      inGroup = true;
    }
  }
  return groups;
}

/** Read the (typed) title-placeholder text from a slide for use by
 * the rail's outline. Returns the empty string when the slide has no
 * `<p:ph type="title"/>` shape; callers fall back to "Slide N" so the
 * outline is never blank. */
function readSlideTitle(slide: import("@officeai/pptx").Slide): string {
  for (const shape of slide.shapes) {
    if (shape.kind !== "text") continue;
    if (!shape.placeholder) continue;
    const phType = shape.placeholder.type;
    // PowerPoint emits both `title` and `ctrTitle` for the title
    // placeholder; both feed the outline.
    if (phType !== "title" && phType !== "ctrTitle") continue;
    const txt = shape.txBody.paragraphs
      .map((p) => p.runs.map((r) => r.text).join(""))
      .join(" ")
      .trim();
    if (txt.length > 0) return txt;
  }
  return "";
}

function readNotesText(snap: PptxSnapshot, slideIndex: number): string {
  const slide = snap.root.slides[slideIndex];
  if (!slide || !slide.notesSlidePartPath) return "";
  const notes = snap.root.notesSlides.get(slide.notesSlidePartPath);
  if (!notes) return "";
  return notes.body.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join("\n");
}

interface NotesPanelProps {
  notesText: string;
  disabled: boolean;
  onChange: (text: string) => void;
}

function NotesPanel(props: NotesPanelProps): React.ReactNode {
  // Local mirror of the textarea so per-keystroke typing is fluid; we
  // commit to the snapshot on blur and on a 600 ms idle (debounced)
  // rather than dispatching a command on every character — the typed
  // notes part rebuild is cheap but the React re-render churn isn't.
  const [draft, setDraft] = useState(props.notesText);
  const [dirty, setDirty] = useState(false);
  const lastSyncedRef = useRef(props.notesText);

  useEffect(() => {
    // Pull the upstream value when the slide changes (parent passes a
    // fresh `key` on slide change, so this branch is mainly defensive).
    if (lastSyncedRef.current !== props.notesText && !dirty) {
      setDraft(props.notesText);
      lastSyncedRef.current = props.notesText;
    }
  }, [props.notesText, dirty]);

  useEffect(() => {
    if (!dirty) return;
    const handle = setTimeout(() => {
      props.onChange(draft);
      lastSyncedRef.current = draft;
      setDirty(false);
    }, 600);
    return () => clearTimeout(handle);
  }, [draft, dirty, props]);

  return (
    <section
      data-testid="pptx-notes-panel"
      className="flex shrink-0 flex-col gap-1 rounded-md border border-divider bg-surface p-2"
    >
      <label className="text-xs font-medium text-secondary">Speaker notes</label>
      <textarea
        className="min-h-[64px] resize-y rounded border border-divider bg-background p-2 text-sm focus:outline-none focus:ring-1 focus:ring-[var(--accent)]"
        placeholder="Add notes for this slide…"
        disabled={props.disabled}
        value={draft}
        onChange={(e) => {
          setDraft(e.target.value);
          setDirty(true);
        }}
        onBlur={() => {
          if (!dirty) return;
          props.onChange(draft);
          lastSyncedRef.current = draft;
          setDirty(false);
        }}
      />
    </section>
  );
}

function textShapePlain(shape: Shape): string {
  if (shape.kind !== "text") return "";
  const lines: string[] = [];
  for (const p of shape.txBody.paragraphs) {
    const text = p.runs.map((r) => r.text).join("");
    lines.push(text);
  }
  return lines.join("\n");
}

function findShape(shapes: ReadonlyArray<Shape>, id: string): Shape | null {
  for (const s of shapes) {
    if (s.id === id) return s;
    if (s.kind === "group") {
      const inner = findShape(s.children, id);
      if (inner) return inner;
    }
  }
  return null;
}

function collectShapesByCNvPrIdLocal(shape: Shape, out: Map<number, Shape>): void {
  if (shape.cNvPrId > 0) out.set(shape.cNvPrId, shape);
  if (shape.kind === "group") {
    for (const c of shape.children) collectShapesByCNvPrIdLocal(c, out);
  }
}

/**
 * Resolve the typed `FillSpec` painted on a text or picture shape so the
 * toolbar's `FillPicker` can render the active value (and reopen on the
 * matching tab — solid / gradient / pattern / picture). Returns `null`
 * when no fill node is declared (caller treats as "inherit").
 */
function readShapeFill(shape: TextShape | Picture): FillSpec | null {
  return readFillSpec(shape.spPrTail);
}

/**
 * Resolve the typed `FillSpec` for a slide's `<p:bg>`. The fill nodes
 * live one level deeper than `spPrTail` (inside `<p:bgPr>`), so we walk
 * manually rather than reusing `readShapeFill`.
 */
function readSlideBackgroundFillSpec(slide: { cSldHead: Slide["cSldHead"] }): FillSpec | null {
  for (const node of slide.cSldHead) {
    if (node.tag !== "p:bg") continue;
    for (const inner of node.subtree) {
      if (!inner || typeof inner !== "object" || Array.isArray(inner)) continue;
      const obj = inner as Record<string, unknown>;
      if (!Array.isArray(obj["p:bgPr"])) continue;
      const fillNodes = collectFillNodesFromObjects(obj["p:bgPr"] as unknown[]);
      const spec = readFillSpec(fillNodes);
      if (spec) return spec;
    }
  }
  return null;
}

/**
 * Walk a fast-xml-parser preserveOrder array (`{tag: [...]}` entries)
 * and lift any fill-related children into `OpaqueXml` form so
 * `readFillSpec` can consume them without further translation.
 */
function collectFillNodesFromObjects(arr: ReadonlyArray<unknown>): OpaqueXml[] {
  const out: OpaqueXml[] = [];
  for (const inner of arr) {
    if (!inner || typeof inner !== "object" || Array.isArray(inner)) continue;
    const obj = inner as Record<string, unknown>;
    const keys = Object.keys(obj).filter((k) => k !== ":@");
    if (keys.length !== 1) continue;
    const tag = keys[0];
    if (
      tag !== "a:solidFill" &&
      tag !== "a:noFill" &&
      tag !== "a:gradFill" &&
      tag !== "a:pattFill" &&
      tag !== "a:blipFill"
    ) {
      continue;
    }
    const sub = obj[tag];
    const attrs = (obj[":@"] as Record<string, unknown> | undefined) ?? {};
    const rawAttrs: Record<string, string> = {};
    const flatAttrs: Record<string, string> = {};
    for (const [k, v] of Object.entries(attrs)) {
      const sv = String(v);
      rawAttrs[k] = sv;
      flatAttrs[k.startsWith("@_") ? k.slice(2) : k] = sv;
    }
    out.push({ tag, attrs: flatAttrs, rawAttrs, subtree: Array.isArray(sub) ? (sub as unknown[]) : [] });
  }
  return out;
}

/**
 * Open the OS file picker for raster images and resolve to the chosen
 * `File` (or `null` if the user cancelled). We deliberately go via a
 * one-shot detached `<input>` rather than stashing a hidden ref on
 * the component: placeholder activation is a rare, transient flow and
 * adding a permanent hidden input + ref to the editor surface for
 * each new affordance would noisily accumulate over time.
 *
 * The MIME allowlist matches the one used by the toolbar's "Insert
 * image" input so the canvas-driven path can't sneak in a file type
 * that `pptx:insert-image` would reject downstream.
 */
/**
 * Translate the {@link EmbeddedXlsxModal}'s plain 2D grid into the
 * `categories + series` shape that `pptx:set-chart-data` expects.
 *
 * Convention (matches `buildChartGrid` in `@officeai/xlsx`):
 *   - Row 0 is the header row; cells `[0][1..N]` are series names.
 *   - Column 0 is the category column; cells `[1..M][0]` are
 *     category labels.
 *   - The `M × N` interior is numeric series values.
 *
 * Returns `null` when the grid is too small to interpret as chart
 * data (e.g. user blanked the modal); the caller skips the
 * `set-chart-data` dispatch in that case so the chart's prior
 * categories survive.
 */
function projectGridToChartData(grid: ReadonlyArray<ReadonlyArray<string | number | null>>): {
  readonly categories: ReadonlyArray<string>;
  readonly series: ReadonlyArray<{ readonly name?: string; readonly values: ReadonlyArray<number> }>;
} | null {
  if (grid.length < 2) return null;
  const header = grid[0]!;
  if (header.length < 2) return null;
  const categories: string[] = [];
  for (let r = 1; r < grid.length; r++) {
    const cell = grid[r]?.[0];
    categories.push(cell == null ? "" : String(cell));
  }
  const series: Array<{ readonly name?: string; readonly values: ReadonlyArray<number> }> = [];
  for (let c = 1; c < header.length; c++) {
    const rawName = header[c];
    const values: number[] = [];
    for (let r = 1; r < grid.length; r++) {
      const cell = grid[r]?.[c];
      const n = typeof cell === "number" ? cell : Number(cell);
      values.push(Number.isFinite(n) ? n : 0);
    }
    const name = rawName == null ? undefined : String(rawName);
    series.push(name !== undefined ? { name, values } : { values });
  }
  return { categories, series };
}

function pickImageFile(): Promise<File | null> {
  return new Promise((resolve) => {
    if (typeof document === "undefined") {
      resolve(null);
      return;
    }
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "image/png,image/jpeg,image/gif,image/webp,image/bmp";
    let resolved = false;
    const settle = (file: File | null): void => {
      if (resolved) return;
      resolved = true;
      resolve(file);
    };
    input.onchange = () => settle(input.files?.[0] ?? null);
    // Some browsers fire `cancel`; for the rest, the focus-back-to-window
    // heuristic (next tick after a short delay) is the only way to
    // detect a cancelled picker. We schedule a "no file picked" fallback
    // so callers don't await forever.
    input.oncancel = () => settle(null);
    window.addEventListener(
      "focus",
      () => {
        window.setTimeout(() => settle(null), 500);
      },
      { once: true }
    );
    input.click();
  });
}

async function readIntrinsicSize(
  bytes: Uint8Array,
  mime: string
): Promise<{ width: number; height: number }> {
  if (typeof window === "undefined") return { width: 0, height: 0 };
  const blob = new Blob([bytes as BlobPart], { type: mime });
  if (typeof createImageBitmap === "function" && mime !== "image/svg+xml") {
    try {
      const bitmap = await createImageBitmap(blob);
      const w = bitmap.width;
      const h = bitmap.height;
      bitmap.close?.();
      return { width: w, height: h };
    } catch {
      // fall through to <img> path
    }
  }
  return await new Promise<{ width: number; height: number }>((resolve, reject) => {
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = () => {
      const out = { width: img.naturalWidth, height: img.naturalHeight };
      URL.revokeObjectURL(url);
      resolve(out);
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("Could not decode image to read intrinsic size."));
    };
    img.src = url;
  });
}

/**
 * Status-bar hint that surfaces "what is selected" and "where am I"
 * for the slide canvas. Live region so screen readers announce
 * selection changes; tabular numerals so the slide counter doesn't
 * jitter as the index advances.
 */
function PptxSelectionHint({
  selectedCount,
  slideIndex,
  slideCount,
}: {
  readonly selectedCount: number;
  readonly slideIndex: number;
  readonly slideCount: number;
}): ReactNode {
  const { t } = useTranslator();
  return (
    <span
      className="flex items-center gap-3 text-[11px] tabular-nums text-tertiary"
      data-testid="pptx-selection-hint"
      aria-live="polite"
    >
      {slideCount > 0 ? <span>{t("status.slideOf", { n: slideIndex + 1, total: slideCount })}</span> : null}
      {selectedCount > 0 ? (
        <span className="rounded bg-[var(--accent-light)] px-1.5 py-0.5 text-[10px] font-medium text-[var(--accent)]">
          {t("status.shapesSelected", { n: selectedCount })}
        </span>
      ) : (
        <span className="opacity-60">{t("status.selectionEmpty")}</span>
      )}
    </span>
  );
}
