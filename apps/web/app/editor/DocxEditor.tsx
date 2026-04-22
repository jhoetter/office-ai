"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Button, cn } from "@officeai/ui";
import {
  EditorShell,
  EmptyState,
  ZoomControl,
  buildPaletteFromCatalogue,
  createToastId,
  type ExportFormat,
  type ExportOptionValues,
  type FindAdapter,
  type FindMatch,
  type FindOptions,
  type OutlineEntry,
  type PaletteCommand,
  type PaletteRunners,
  type ProductAdapter,
  type SaveState,
  type ToastItem,
} from "@/lib/shell";
import { docxActions } from "@officeai/docx";
import { useTranslator } from "@/lib/i18n";
import {
  PRODUCT_FILE_TYPES,
  downloadBlob,
  openFile as openFileViaService,
  saveFile as saveFileViaService,
} from "@/lib/files/file-service";
import { convertViaServer } from "@/lib/files/convert-client";
import { docxToMarkdown, docxToText } from "./lib/export-text";
import {
  DocxAgent,
  chunkIntoPages,
  documentPageGeometry,
  documentMaxPageGeometry,
  listBookmarks,
  mountDocxEditor,
  resolveEffectivePpr,
} from "@officeai/docx";
import type { BlockNode, DocxSnapshot, InlineNode, MountResult, UnsupportedTx } from "@officeai/docx";
import {
  getPageChunks,
  gotoPage,
  pageDecorationsPlugin,
  pageNumberForPos,
  PAGE_ZONE_COMMIT_EVENT,
  PAGE_ZONE_FOCUS_EVENT,
  PAGE_ZONE_MINT_EVENT,
  type PageZoneCommitDetail,
  type PageZoneFocusDetail,
  type PageZoneMintDetail,
} from "@/lib/page-decorations";
import { GOTO_PAGE_EVENT, pageKeymapPlugin } from "@/lib/page-keymap";
import {
  SHORTCUT_ADD_COMMENT_EVENT,
  SHORTCUT_INSERT_HYPERLINK_EVENT,
  SHORTCUT_TOGGLE_FORMATTING_MARKS_EVENT,
  wordShortcutsKeymapPlugin,
  type InsertHyperlinkDetail,
} from "@/lib/word-shortcuts-keymap";
import {
  formattingMarksPlugin,
  isFormattingMarksOn,
  toggleFormattingMarks,
} from "@/lib/formatting-marks-plugin";
import { useShortcutsDialog } from "@/lib/shortcuts/useShortcutsDialog";
import { KeyboardShortcutsDialog } from "@/lib/shortcuts/KeyboardShortcutsDialog";
import type { EditorView } from "prosemirror-view";
import { TextSelection } from "prosemirror-state";
import { NotImplementedError } from "@officeai/core";
import { buildBlankDocx, buildSampleDocx } from "@/lib/sample-docx";
import { I18nProvider, type Locale } from "@/lib/i18n";
import {
  activeMarkAttr,
  commentParagraphIndex,
  commentThreads,
  currentParagraphAlignment,
  currentParagraphId,
  currentParagraphIndex,
  paragraphStyle,
  paragraphStyleOptions,
  pmPositionToDocx,
  pmSelectionToRange,
} from "@/lib/format-helpers";
import { Toolbar, type AlignmentValue, type ResolvedSpacingDisplay } from "./Toolbar";
import { BookmarkDialog, type BookmarkRow } from "./BookmarkDialog";
import { FootnotesPanel } from "./FootnotesPanel";
import { computeDocxActive, createDocxFormatProvider } from "./docxFormatProvider";
import { PageRuler } from "./PageRuler";
import { PageSetupDialog, type PageSetupValues } from "./PageSetupDialog";
import {
  ProtectDocumentDialog,
  type ProtectDocumentSubmit,
  type ProtectionEdit,
} from "./ProtectDocumentDialog";
import { TableContextToolbar } from "./TableContextToolbar";
import { ImageContextToolbar, type SelectedImageInfo } from "./ImageContextToolbar";
import { ImageResizeOverlay } from "./ImageResizeOverlay";
import { AltTextDialog } from "./AltTextDialog";
import { GotoDialog } from "./GotoDialog";
import { HyperlinkPopover } from "./HyperlinkPopover";
import { CommentsSidebar } from "./CommentsSidebar";
import { TrackedChangesHover, TrackedChangesMargin } from "./TrackedChangesUI";
import { CommentComposer } from "./CommentComposer";
import { DocxRemoteCursorLayer } from "./DocxRemoteCursorLayer";
import { collectRevisions } from "@/lib/format-helpers";
import { createDocxCommentsProvider } from "./docxCommentsProvider";
import { insertImageIntoDocx, SUPPORTED_IMAGE_MIME } from "@/lib/image-insert";
import { EMBED_MIME, isEmbedEnabled, parseEnvelope } from "@/lib/embed/envelope";
import { applyXlsxRangeToDocx } from "@/lib/embed/applyXlsxRangeToDocx";
import { applyXlsxEmbed } from "@/lib/embed/xlsxEmbedShared";
import type { XlsxEmbedMode } from "@/lib/embed/xlsxEmbedShared";
import { EmbeddedXlsxModal } from "@/lib/embed/EmbeddedXlsxModal";
import { resolveEmbeddedXlsxRef, readEmbeddedXlsxBytes } from "@/lib/embed/getEmbeddedXlsxBytes";
import { XlsxRangePickerDialog, type XlsxRangePickerResult } from "@/lib/embed/XlsxRangePickerDialog";
import { installAltKeyTracker, isAltKeyPressed } from "@/lib/embed/altKeyTracker";
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

const DOCX_EXPORT_FORMATS: ReadonlyArray<ExportFormat> = [
  {
    id: "docx",
    label: "Word document (.docx)",
    description: "Round-trip native OOXML — what Word, Pages and LibreOffice open by default.",
    extension: "docx",
    mime: PRODUCT_FILE_TYPES.docx.primaryMime,
    kind: "instant",
    group: "native",
    icon: "doc",
  },
  {
    id: "pdf",
    label: "PDF document (.pdf)",
    description: "Server-side conversion via LibreOffice. Highest layout fidelity for sharing.",
    extension: "pdf",
    mime: "application/pdf",
    kind: "dialog",
    group: "pdf-web",
    icon: "pdf",
    optionFields: [
      {
        id: "pageRange",
        label: "Page range",
        control: { type: "text", placeholder: "All pages — try 1-3, 5" },
        hint: "Leave blank for the whole document. Examples: 1,3 — 2-5 — 1,4-7,10.",
      },
      {
        id: "pageSize",
        label: "Page size",
        control: {
          type: "select",
          defaultId: "source",
          options: [
            { id: "source", label: "Use document setting" },
            { id: "a4", label: "A4 (210 × 297 mm)" },
            { id: "letter", label: "Letter (8.5 × 11 in)" },
          ],
        },
        hint: "LibreOffice prints with the document's page setup unless overridden.",
      },
      {
        id: "embedFonts",
        label: "Embed fonts",
        control: { type: "toggle", defaultValue: true },
        hint: "Keeps the look intact on systems missing your fonts.",
      },
    ],
  },
  {
    id: "html",
    label: "Web page (.html)",
    description: "Server-side HTML export. Useful for previews, intranets, and email.",
    extension: "html",
    mime: "text/html",
    kind: "instant",
    group: "pdf-web",
    icon: "code",
  },
  {
    id: "txt",
    label: "Plain text (.txt)",
    description: "Body text only — paragraphs, headings and lists. No formatting.",
    extension: "txt",
    mime: "text/plain;charset=utf-8",
    kind: "instant",
    group: "data",
    icon: "text",
  },
  {
    id: "md",
    label: "Markdown (.md)",
    description: "GitHub-flavoured Markdown. Headings, emphasis, lists and tables.",
    extension: "md",
    mime: "text/markdown;charset=utf-8",
    kind: "instant",
    group: "data",
    icon: "text",
  },
];

function stripDocxExtension(name: string): string {
  return name.replace(/\.docx$/i, "");
}

export interface DocxEditorProps {
  /** Fired whenever the editor's bootstrap-ready state changes. The
   * page-level splash listens to this to know when to fade out and
   * unveil the editor. Stays `false` until the agent is mounted and
   * the first snapshot has been rendered, then `true`. */
  readonly onBootstrapReady?: (ready: boolean) => void;
  /** Optional pre-loaded document. When provided, the editor fetches
   * the bytes at `url` instead of building the synthetic welcome
   * sample, and uses `name` as the document title (so subsequent
   * Save / Export keep the original filename). Used by the home
   * page's "sample files" listing. */
  readonly initialSource?: { readonly url: string; readonly name: string };
  /** When true, the editor bootstraps with a truly empty document
   * (single blank paragraph) instead of the synthetic welcome
   * sample. Used by the home page's "New document" action so users
   * land in a fresh file rather than the demo content. Ignored when
   * `initialSource` is set. */
  readonly initialBlank?: boolean;
  /** Optional pre-loaded document bytes. When set, takes priority
   * over `initialSource` and `initialBlank` so embedding hosts can
   * stream a `Uint8Array` straight into the editor without first
   * stashing it under a URL. The companion `initialFilename`
   * controls the working filename (Save / Export). */
  readonly initialBytes?: Uint8Array;
  /** Filename to display + use on Save when `initialBytes` is set.
   * Ignored unless `initialBytes` is provided (since the URL-based
   * `initialSource` already carries `name`). */
  readonly initialFilename?: string;
  /** Host save handler. When provided the editor's Save action
   * invokes this with the freshly-exported bytes, the OOXML MIME,
   * and the working filename — instead of reaching into the browser
   * `saveFile` File-System-Access fallback. Hosts (e.g. hof-os)
   * wire this up to push the bytes back to their own storage. */
  readonly onSave?: (bytes: Uint8Array, mime: string, filename: string) => Promise<void>;
  /** Host close handler. When provided, surfaces a "Back" affordance
   * in the editor chrome so users can return to the host. The
   * embedding route is responsible for the actual navigation. */
  readonly onClose?: () => void;
  /** Override the i18n locale. When set, the editor mounts its own
   * `<I18nProvider initialLocale={locale}>` so a host whose root
   * provider is in a different locale (or absent entirely) still
   * gets the correct UI language for the editor surface. */
  readonly locale?: Locale;
  /** Theme override forwarded to next-themes inside Phase 1's
   * extracted package. Today this is a passthrough placeholder so
   * the embedding contract is stable across phases. */
  readonly theme?: "light" | "dark";
  /** Realtime presence identity (host-supplied). When set, replaces
   * the default anonymous "Adjective Animal" identity on the
   * awareness payload so cursors / avatars / tracked-changes show
   * the authenticated user's real name. See `PresenceUser` in
   * `@officeai/react-editors` for the per-field contract. */
  readonly presenceUser?: { readonly id: string; readonly name: string; readonly color?: string };
  /** Explicit realtime room id (host-supplied). When set, the editor
   * joins this exact room instead of deriving one from `?src=` /
   * `?room=` / a per-tab fallback. Pass `null` to disable realtime
   * for this mount; omit / `undefined` keeps the built-in default. */
  readonly room?: string | null;
  /** Hide the 📁 Open toolbar affordance. Set by embedded hosts that
   * own their document corpus — see
   * `EmbeddedEditorProps.hideLocalFileOpen` in
   * `@officeai/react-editors/contract`. */
  readonly hideLocalFileOpen?: boolean;
}

/**
 * The editor surface composed from the right-hand collaboration panels:
 *
 *   ┌────────────────────────────────────────────┬──────────────┐
 *   │ Toolbar (style/marks/colors/align/lists)   │              │
 *   ├────────────────────────────────────────────┤   Comments   │
 *   │                                            │              │
 *   │           ProseMirror editor surface       │   Tracked    │
 *   │                                            │   changes    │
 *   └────────────────────────────────────────────┴──────────────┘
 *
 * Below 1024px the right column hides behind a "Comments" drawer
 * button anchored bottom-right.
 *
 * AI/agent affordances were intentionally removed from the editor
 * surface: every command-bus mutation that the editor exposes is also
 * reachable via the headless `office-agent` CLI, which is the canonical
 * integration point for third-party agents. The bus stays the same
 * either way; this UI only shows human-driven actions.
 */
export function DocxEditor(props: DocxEditorProps = {}): React.ReactNode {
  const { locale } = props;
  if (locale !== undefined) {
    return (
      <I18nProvider initialLocale={locale}>
        <DocxEditorInner {...props} />
      </I18nProvider>
    );
  }
  return <DocxEditorInner {...props} />;
}

function DocxEditorInner({
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
}: DocxEditorProps = {}): React.ReactNode {
  const { t } = useTranslator();
  // The editor host DOM node is exposed via a callback ref so that
  // descendants (e.g. TrackedChangesUI's hover delegation) can read it
  // from React state during render — accessing `hostRef.current`
  // directly during render trips `react-hooks/refs`.
  const [hostEl, setHostEl] = useState<HTMLDivElement | null>(null);
  // The scroll container around the page card is exposed via a callback
  // ref so TrackedChangesMargin can compute coordinates in its content
  // space (so balloons stay glued to the document on scroll).
  const [scrollEl, setScrollEl] = useState<HTMLDivElement | null>(null);
  const imageInputRef = useRef<HTMLInputElement | null>(null);
  const fileHandleRef = useRef<FileSystemFileHandle | undefined>(undefined);
  const [saveState, setSaveState] = useState<SaveState>("saved");
  // `agentRef` / `mountRef` are kept in addition to the React state
  // mirrors below so that long-lived callbacks (file open, accept
  // change, …) capture a stable reference without re-binding on
  // every state change.
  const agentRef = useRef<DocxAgent | null>(null);
  const mountRef = useRef<MountResult | null>(null);
  const [agent, setAgent] = useState<DocxAgent | null>(null);
  const [view, setView] = useState<EditorView | null>(null);
  const [agentReady, setAgentReady] = useState(false);

  // Mirror `agentReady` up to the page-level splash so it can fade
  // out and unveil the editor. Owning the splash at page scope (not
  // here) keeps the badge `<span>` mounted across the dynamic-import
  // handoff — see `apps/web/app/editor/page.tsx`.
  useEffect(() => {
    onBootstrapReady?.(agentReady);
  }, [agentReady, onBootstrapReady]);
  const [toasts, setToasts] = useState<ReadonlyArray<ToastItem>>([]);
  const [docName, setDocName] = useState(
    initialFilename ??
      initialSource?.name ??
      (initialBlank || initialBytes ? "Untitled.docx" : "welcome.docx")
  );
  const [docInfo, setDocInfo] = useState<{
    paragraphs: number;
    revision: number;
    commentThreads: number;
    pageCount: number;
  } | null>(null);
  const [zoom, setZoom] = useState<number>(1);
  const [pageSetupOpen, setPageSetupOpen] = useState(false);
  const [protectDocumentOpen, setProtectDocumentOpen] = useState(false);
  const [bookmarkDialogOpen, setBookmarkDialogOpen] = useState(false);
  /**
   * Currently-targeted cell within `selectedTableId` for the
   * Tabellentools contextual tab. DOCX tables render today as
   * ProseMirror node atoms — there is no caret-level cell editing
   * yet — so we let the user pick the row/column index explicitly
   * (PowerPoint's table layout dialog uses the same model). The
   * picker resets to {0,0} whenever the user selects a different
   * table, which keeps the dispatch targets coherent with whatever
   * table the surrounding `selectedTableId` resolves to.
   */
  const [activeTableCell, setActiveTableCell] = useState<{ row: number; column: number }>({
    row: 0,
    column: 0,
  });
  const [xlsxPickerOpen, setXlsxPickerOpen] = useState<XlsxEmbedMode | null>(null);
  /**
   * Active "Edit Data" modal for double-click on a chart drawing /
   * embedded spreadsheet. `null` while the modal is closed.
   */
  const [editingEmbed, setEditingEmbed] = useState<{
    readonly bytes: Uint8Array;
    readonly embeddingPartPath: string;
    readonly chartPartPath: string | null;
    readonly title: string;
  } | null>(null);
  // Bumped to force re-derivation of toolbar state (active marks /
  // active style) without keeping a redundant copy of the snapshot.
  const [uiTick, setUiTick] = useState(0);
  const [drawerOpen, setDrawerOpen] = useState(false);
  // Edit mode (Word / Google Docs "Track Changes" surface). The
  // toolbar mode picker writes through to PM via `mount.setEditMode`
  // so the underlying transaction-to-commands pipeline can swap
  // `insert-text` ↔ `insert-text-tracked` without re-mounting.
  // `view` mode flips PM's `editable` prop off so the surface
  // becomes read-only. Author defaults to "You" so suggestions are
  // attributable from day one; a real auth integration would feed
  // the signed-in user's display name here.
  const [editMode, setEditMode] = useState<"edit" | "suggest" | "view">("edit");
  const [trackedAuthor] = useState<string>("You");
  // Pilcrow toggle (Word's "Show formatting marks"). Owned by the PM
  // plugin; the React state below is just a mirror so the toolbar
  // button can show pressed-state. Refresh on every uiTick.
  const [formattingMarksOn, setFormattingMarksOn] = useState(false);
  const shortcutsDialog = useShortcutsDialog();
  // Mirror the mode/author into refs so `mountAgent` (a
  // useCallback) can read the latest value without re-binding on
  // every mode flip. The effect below keeps them in sync.
  const editModeRef = useRef<"edit" | "suggest" | "view">("edit");
  const trackedAuthorRef = useRef<string>("You");
  useEffect(() => {
    editModeRef.current = editMode;
    trackedAuthorRef.current = trackedAuthor;
    mountRef.current?.setEditMode(editMode, trackedAuthor);
  }, [editMode, trackedAuthor]);

  const pushToast = useCallback((kind: ToastItem["kind"], text: string) => {
    const id = createToastId("docx");
    setToasts((prev) => [...prev, { id, kind, text }]);
  }, []);
  const dismissToast = useCallback((id: string) => {
    setToasts((prev) => prev.filter((t) => t.id !== id));
  }, []);

  const onUnsupported = useCallback(
    (events: UnsupportedTx[]) => {
      pushToast("warn", `Action deferred — see build log (${events[0]?.reason ?? ""})`);
    },
    [pushToast]
  );

  const onError = useCallback(
    (err: unknown) => {
      pushToast("error", err instanceof Error ? err.message : String(err));
    },
    [pushToast]
  );

  const mountAgent = useCallback(
    async (buf: ArrayBuffer, host: HTMLDivElement) => {
      mountRef.current?.destroy();
      const agentInstance = await DocxAgent.fromBuffer(buf);
      agentRef.current = agentInstance;
      setAgent(agentInstance);
      const refreshState = () => {
        const snap = agentInstance.getSnapshot();
        const paragraphs = snap.root.body.reduce((n, b) => (b.kind === "paragraph" ? n + 1 : n), 0);
        setDocInfo({
          paragraphs,
          revision: snap.revision,
          commentThreads: commentThreads(snap).length,
          pageCount: chunkIntoPages(snap).length,
        });
        setUiTick((t) => t + 1);
      };
      let firstSnap = true;
      const off = agentInstance.subscribe((_snap, mutation) => {
        refreshState();
        if (firstSnap) {
          firstSnap = false;
          return;
        }
        // Surface rebase rejections — when an undo/redo (or a
        // pending-mutation rebase pass) flips a previously
        // pending agent mutation to "rejected", the bus fires
        // one notify per rejected mutation with the
        // `rebase-failed` code. The user sees a toast instead
        // of the suggestion silently disappearing.
        if (mutation.status === "rejected" && mutation.rejection?.code === "rebase-failed") {
          pushToast(
            "warn",
            `An agent suggestion couldn't be re-applied after the last edit (${mutation.rejection.message})`
          );
          return;
        }
        setSaveState("modified");
      });
      host.innerHTML = "";
      const mount = mountDocxEditor(host, {
        agent: agentInstance,
        source: "human",
        onUnsupported,
        onError,
        extraPlugins: [
          pageDecorationsPlugin(agentInstance),
          pageKeymapPlugin(agentInstance),
          wordShortcutsKeymapPlugin(agentInstance),
          formattingMarksPlugin(),
        ],
        // Re-apply the current mode/author on every (re)mount so
        // file-open keeps the user's mode choice across documents.
        editMode: editModeRef.current,
        trackedAuthor: trackedAuthorRef.current,
      });
      mountRef.current = mount;
      setView(mount.view);
      setAgentReady(true);
      refreshState();
      return () => {
        off();
        mount.destroy();
      };
    },
    [onError, onUnsupported]
  );

  useEffect(() => {
    if (!hostEl) return;
    let cancelled = false;
    let cleanup: (() => void) | undefined;
    void (async () => {
      try {
        // Four bootstrap paths, picked in priority order:
        //   1. `initialBytes` — host streams the document straight
        //      in (used when the editor is embedded by hof-os and
        //      the bytes already came back from S3).
        //   2. `initialSource` — fetch a pre-existing .docx (sample
        //      files listing on the home page).
        //   3. `initialBlank` — build a truly empty document (the
        //      "New document" action on the home page).
        //   4. Default — build the synthetic welcome sample so the
        //      editor route is never empty when navigated to
        //      directly.
        let buf: ArrayBuffer;
        if (initialBytes) {
          // Copy into a fresh ArrayBuffer so downstream consumers
          // can transfer the buffer (PDF.js worker etc.) without
          // detaching the host's pristine view.
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
          buf = await buildBlankDocx();
        } else {
          buf = await buildSampleDocx();
        }
        if (cancelled) return;
        cleanup = await mountAgent(buf, hostEl);
      } catch (err) {
        if (cancelled) return;
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    })();
    return () => {
      cancelled = true;
      cleanup?.();
      mountRef.current?.destroy();
      mountRef.current = null;
      agentRef.current = null;
      setAgent(null);
      setView(null);
    };
  }, [mountAgent, hostEl, initialSource, initialBlank, initialBytes, pushToast]);

  // Re-derive toolbar UI state on every selection change so Bold/Italic
  // pressed-state and the paragraph style picker stay in sync with the
  // caret. PM updates the DOM selection synchronously inside its own
  // dispatchTransaction, so the browser fires `selectionchange` on every
  // PM transaction including pure caret moves.
  useEffect(() => {
    const onSel = () => setUiTick((t) => t + 1);
    document.addEventListener("selectionchange", onSel);
    return () => document.removeEventListener("selectionchange", onSel);
  }, []);

  // Mirror the formatting-marks plugin's state into React so the
  // toolbar pressed-state stays in sync. The plugin is the source
  // of truth (it's what the keyboard shortcut also flips).
  useEffect(() => {
    if (!view) return;
    setFormattingMarksOn(isFormattingMarksOn(view.state));
  }, [view, uiTick]);

  const handleToggleFormattingMarks = useCallback(() => {
    const mount = mountRef.current;
    if (!mount) return;
    const next = toggleFormattingMarks(mount.view);
    setFormattingMarksOn(next);
  }, []);

  const handleFile = useCallback(
    async (file: File, handle?: FileSystemFileHandle) => {
      const buf = await file.arrayBuffer();
      setDocName(file.name);
      fileHandleRef.current = handle;
      const host = hostEl;
      if (!host) {
        pushToast("error", "Editor not yet mounted.");
        return;
      }
      try {
        await mountAgent(buf, host);
        setSaveState("saved");
        pushToast("success", `Opened ${file.name}`);
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [mountAgent, pushToast, hostEl]
  );

  const handleOpenFile = useCallback(async () => {
    try {
      const opened = await openFileViaService({
        description: PRODUCT_FILE_TYPES.docx.description,
        mimeToExt: PRODUCT_FILE_TYPES.docx.mimeToExt,
        accept: PRODUCT_FILE_TYPES.docx.accept,
      });
      if (!opened) return;
      const file = new File([opened.bytes as BlobPart], opened.name, {
        type: PRODUCT_FILE_TYPES.docx.primaryMime,
      });
      await handleFile(file, opened.handle);
    } catch (err) {
      pushToast("error", err instanceof Error ? err.message : String(err));
    }
  }, [handleFile, pushToast]);

  const handleSave = useCallback(async () => {
    const agent = agentRef.current;
    if (!agent) return;
    setSaveState("saving");
    try {
      const buf = await agent.exportFile();
      const bytes = new Uint8Array(buf);
      const mime = PRODUCT_FILE_TYPES.docx.primaryMime;
      // Embedding hosts (hof-os) get first dibs — when `onSave` is
      // provided we hand bytes back to the host instead of going
      // through the browser File-System-Access fallback.
      if (onSaveProp) {
        await onSaveProp(bytes, mime, docName);
        setSaveState("saved");
        pushToast("success", `Saved ${docName}`);
        return;
      }
      const wroteInPlace = await saveFileViaService(bytes, docName, mime, fileHandleRef.current);
      setSaveState("saved");
      pushToast("success", wroteInPlace ? `Saved ${docName}` : `Downloaded ${docName}`);
    } catch (err) {
      setSaveState("error");
      pushToast("error", err instanceof Error ? err.message : String(err));
    }
  }, [docName, onSaveProp, pushToast]);

  const handleExport = useCallback(
    async (format: ExportFormat, options?: ExportOptionValues) => {
      const agent = agentRef.current;
      if (!agent) return;
      const baseName = stripDocxExtension(docName);
      const downloadName = `${baseName}.${format.extension}`;
      try {
        switch (format.id) {
          case "docx": {
            const buf = await agent.exportFile();
            downloadBlob(new Blob([buf as BlobPart], { type: format.mime }), downloadName);
            break;
          }
          case "pdf":
          case "html": {
            const buf = await agent.exportFile();
            // Page-range is PDF-only; HTML export ignores it. We
            // trim whitespace defensively because the dialog's text
            // input can carry leading/trailing spaces from copy-paste.
            const pageRange =
              format.id === "pdf" && typeof options?.pageRange === "string" ? options.pageRange.trim() : "";
            const out = await convertViaServer({
              bytes: new Uint8Array(buf),
              sourceExt: "docx",
              targetExt: format.id,
              filename: baseName,
              pageRange: pageRange.length > 0 ? pageRange : undefined,
            });
            downloadBlob(out, downloadName);
            break;
          }
          case "txt": {
            const text = docxToText(agent.getSnapshot());
            downloadBlob(new Blob([text], { type: format.mime }), downloadName);
            break;
          }
          case "md": {
            const md = docxToMarkdown(agent.getSnapshot());
            downloadBlob(new Blob([md], { type: format.mime }), downloadName);
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
    },
    [docName, pushToast]
  );

  const applyPageSetup = useCallback(
    async (next: PageSetupValues) => {
      const agent = agentRef.current;
      const mount = mountRef.current;
      if (!agent) return;
      const idx = mount ? currentParagraphIndex(mount.view.state) : 0;
      try {
        await agent.applyCommand({
          type: "docx:set-page-setup",
          payload: {
            paragraphIndex: idx,
            pgSz: next.pgSz,
            pgMar: next.pgMar,
          },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  /**
   * Apply a Restrict-Editing change. The dialog has already hashed
   * the optional password (Word-style SHA-512 + 100 000 spins) so we
   * just forward the typed payload onto the `docx:set-protection`
   * handler. Passing `enabled: false` clears the existing
   * `<w:documentProtection>` element entirely.
   */
  const applyProtection = useCallback(
    async (payload: ProtectDocumentSubmit) => {
      const agent = agentRef.current;
      if (!agent) return;
      try {
        await agent.applyCommand({
          type: "docx:set-protection",
          payload: {
            enabled: payload.enabled,
            ...(payload.edit && payload.edit !== "none" ? { edit: payload.edit } : {}),
            ...(payload.enforce !== undefined ? { enforce: payload.enforce } : {}),
            ...(payload.formatting !== undefined ? { formatting: payload.formatting } : {}),
            ...(payload.algorithmName ? { algorithmName: payload.algorithmName } : {}),
            ...(payload.hashValue ? { hashValue: payload.hashValue } : {}),
            ...(payload.saltValue ? { saltValue: payload.saltValue } : {}),
            ...(payload.spinCount !== undefined ? { spinCount: payload.spinCount } : {}),
          },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const handleRulerMargins = useCallback(
    async (next: { left: number; right: number }) => {
      const agent = agentRef.current;
      const mount = mountRef.current;
      if (!agent) return;
      const idx = mount ? currentParagraphIndex(mount.view.state) : 0;
      try {
        await agent.applyCommand({
          type: "docx:set-page-setup",
          payload: {
            paragraphIndex: idx,
            pgMar: { left: next.left, right: next.right },
          },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const setParagraphStyle = useCallback(
    async (style: string) => {
      const agent = agentRef.current;
      const mount = mountRef.current;
      if (!agent || !mount) return;
      const idx = currentParagraphIndex(mount.view.state);
      try {
        await agent.applyCommand({
          type: "docx:set-paragraph-style",
          payload: { at: { paragraph: idx, run: 0, offset: 0 }, style },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  /**
   * Comment composer state (P2.5 / W24).
   *
   * `composer` is null when the popover is hidden. When the user clicks
   * "Add comment" with a non-empty selection we capture the selection
   * range + plain text + an anchor coordinate (via `view.coordsAtPos`)
   * and surface the popover. Submission funnels through the shared
   * `add-comment` command so existing tests, the comments sidebar, and
   * the OOXML round-trip all stay correct.
   */
  const [composer, setComposer] = useState<{
    range: {
      start: { paragraph: number; run: number; offset: number };
      end: { paragraph: number; run: number; offset: number };
    };
    selectionText: string;
    anchor: { left: number; top: number; bottom: number } | null;
  } | null>(null);

  // Realtime author identity, mirrored from `realtimeRoom.room?.identity`
  // by an effect declared after the `useRealtimeRoom` call below.
  // Hoisted up here so memoized renderers (e.g. `renderCommentsPanel`)
  // can read it without forward-referencing `realtimeRoom`.
  const [authorIdentity, setAuthorIdentity] = useState<{
    readonly name: string;
    readonly id: string;
    readonly color: string;
  } | null>(null);

  // B5 — hyperlink popover state. Captures the requesting paragraph
  // + flat-text range (so the eventual command targets the same
  // selection even if the user clicks elsewhere while the popover
  // is open) and an anchor coordinate for visual placement.
  const [linkRequest, setLinkRequest] = useState<{
    paragraphId: string;
    range: { start: number; end: number };
    selectionText: string;
    anchor: { left: number; top: number; bottom: number } | null;
    initialUrl: string;
  } | null>(null);

  // B6 — alt-text dialog state. Driven by the image context toolbar's
  // "Alt text" button; we capture the image id + current value so the
  // dialog can preserve any existing description while the user edits.
  const [altTextRequest, setAltTextRequest] = useState<{
    imageId: string;
    initial: string;
  } | null>(null);

  // B10 — "Go to page" dialog. Opens via Mod+G (dispatched by the
  // page keymap plugin) or by clicking the page status pill in the
  // bottom status bar.
  const [gotoOpen, setGotoOpen] = useState(false);

  // P3.8 / B2 — in-place header/footer authoring. The page-decorations
  // plugin renders each zone's content as a `contenteditable=true`
  // island that fires `pm-page-zone-commit` on blur (or Enter). No
  // separate popover or side panel — the user clicks straight into
  // the header/footer to author it, mirroring Word.

  const openCommentComposer = useCallback(() => {
    const agent = agentRef.current;
    const mount = mountRef.current;
    if (!agent || !mount) return;
    const view = mount.view;
    if (view.state.selection.empty) {
      pushToast("info", "Select some text to comment on.");
      return;
    }
    const range = pmSelectionToRange(view.state);
    const selectionText = view.state.doc.textBetween(
      view.state.selection.from,
      view.state.selection.to,
      " ",
      " "
    );
    let anchor: { left: number; top: number; bottom: number } | null = null;
    try {
      const coords = view.coordsAtPos(view.state.selection.from);
      anchor = { left: coords.left, top: coords.top, bottom: coords.bottom };
    } catch {
      anchor = null;
    }
    setComposer({ range, selectionText, anchor });
  }, [pushToast]);

  // The actual implementation lives later (it depends on the
  // commentsProvider memo, which is declared further below). We
  // expose a stable wrapper here so callers higher in the file can
  // bind it without hoisting the heavy `useMemo`.
  const submitCommentRef = useRef<(text: string) => void>(() => undefined);
  const submitComment = useCallback((text: string) => {
    submitCommentRef.current(text);
  }, []);

  // Word-shortcut bridge: the keymap plugin can't open prompts /
  // composers itself (it lives in the renderer layer), so it
  // surfaces requests as `CustomEvent`s on the editor host. We
  // route those into the existing comment-composer flow and a
  // tiny URL prompt for hyperlink insertion. Both validate again
  // on the React side so a stale event (e.g. selection collapsed
  // before the user typed in the prompt) is a no-op.
  useEffect(() => {
    if (!hostEl) return;
    const onAddComment = () => openCommentComposer();
    const onInsertHyperlink = (event: Event) => {
      const ce = event as CustomEvent<InsertHyperlinkDetail>;
      const detail = ce.detail;
      if (!detail) return;
      const mount = mountRef.current;
      let anchor: { left: number; top: number; bottom: number } | null = null;
      if (mount) {
        try {
          const coords = mount.view.coordsAtPos(mount.view.state.selection.from);
          anchor = { left: coords.left, top: coords.top, bottom: coords.bottom };
        } catch {
          anchor = null;
        }
      }
      setLinkRequest({
        paragraphId: detail.paragraphId,
        range: detail.range,
        selectionText: detail.selectionText,
        anchor,
        initialUrl: "",
      });
    };
    const onToggleFormattingMarks = () => handleToggleFormattingMarks();
    hostEl.addEventListener(SHORTCUT_ADD_COMMENT_EVENT, onAddComment as EventListener);
    hostEl.addEventListener(SHORTCUT_INSERT_HYPERLINK_EVENT, onInsertHyperlink as EventListener);
    hostEl.addEventListener(SHORTCUT_TOGGLE_FORMATTING_MARKS_EVENT, onToggleFormattingMarks as EventListener);
    return () => {
      hostEl.removeEventListener(SHORTCUT_ADD_COMMENT_EVENT, onAddComment as EventListener);
      hostEl.removeEventListener(SHORTCUT_INSERT_HYPERLINK_EVENT, onInsertHyperlink as EventListener);
      hostEl.removeEventListener(
        SHORTCUT_TOGGLE_FORMATTING_MARKS_EVENT,
        onToggleFormattingMarks as EventListener
      );
    };
  }, [hostEl, openCommentComposer, pushToast, handleToggleFormattingMarks]);

  // B10 — Mod+G "Go to page". The page-keymap plugin fires a window
  // CustomEvent so this listener is independent of editor focus and
  // can survive remounts of the PM view; we just open the dialog.
  useEffect(() => {
    const onGoto = () => setGotoOpen(true);
    window.addEventListener(GOTO_PAGE_EVENT, onGoto as EventListener);
    return () => window.removeEventListener(GOTO_PAGE_EVENT, onGoto as EventListener);
  }, []);

  useEffect(() => {
    if (!hostEl) return;
    const onZoneCommit = (event: Event) => {
      const ce = event as CustomEvent<PageZoneCommitDetail>;
      const detail = ce.detail;
      if (!detail || !detail.partPath) return;
      const agent = agentRef.current;
      if (!agent) return;
      // Word-style in-place authoring: when the rich-edit surface
      // produced a multi-paragraph / token-bearing body, dispatch
      // `docx:set-header-footer-blocks` so page-number fields and
      // inline images survive the round-trip. Otherwise fall back
      // to the legacy plain-text command for back-compat with the
      // single-paragraph happy path.
      if (detail.blocks && detail.blocks.length > 0) {
        void agent
          .applyCommand({
            type: "docx:set-header-footer-blocks",
            payload: { partPath: detail.partPath, body: detail.blocks },
            source: "human",
          })
          .catch((err) => pushToast("error", err instanceof Error ? err.message : String(err)));
        return;
      }
      const cmdType = detail.slot === "header" ? "docx:set-header-text" : "docx:set-footer-text";
      void agent
        .applyCommand({
          type: cmdType,
          payload: { partId: detail.partPath, paragraphIndex: 0, text: detail.text },
          source: "human",
        })
        .catch((err) => pushToast("error", err instanceof Error ? err.message : String(err)));
    };
    hostEl.addEventListener(PAGE_ZONE_COMMIT_EVENT, onZoneCommit as EventListener);
    return () => hostEl.removeEventListener(PAGE_ZONE_COMMIT_EVENT, onZoneCommit as EventListener);
  }, [hostEl, pushToast]);

  // Track which header/footer zone is focused so the toolbar can
  // surface the contextual "Header & Footer" button cluster (Word's
  // "Close header and footer" affordance + per-section toggles).
  const [hfZoneFocus, setHfZoneFocus] = useState<PageZoneFocusDetail | null>(null);
  useEffect(() => {
    if (!hostEl) return;
    const onZoneFocus = (event: Event) => {
      const ce = event as CustomEvent<PageZoneFocusDetail>;
      const detail = ce.detail;
      if (!detail || !detail.slot) {
        setHfZoneFocus(null);
        return;
      }
      setHfZoneFocus(detail);
    };
    hostEl.addEventListener(PAGE_ZONE_FOCUS_EVENT, onZoneFocus as EventListener);
    return () => hostEl.removeEventListener(PAGE_ZONE_FOCUS_EVENT, onZoneFocus as EventListener);
  }, [hostEl]);

  // Word-style "double-click an empty header/footer to start typing":
  // the empty zone fires `PAGE_ZONE_MINT_EVENT` on dblclick, we
  // dispatch the mint command, then drop the caret into the freshly
  // rendered editable zone after the next paint. The
  // `requestAnimationFrame` ensures the snapshot has propagated and
  // the page-decorations plugin has rebuilt the widget; otherwise
  // the focus would land on a stale (still-no-part) DOM node.
  useEffect(() => {
    if (!hostEl) return;
    const onZoneMint = (event: Event) => {
      const ce = event as CustomEvent<PageZoneMintDetail>;
      const detail = ce.detail;
      if (!detail) return;
      const agent = agentRef.current;
      if (!agent) return;
      void (async () => {
        try {
          await agent.applyCommand({
            type: "docx:create-header-footer-part",
            payload: { slot: detail.slot },
            source: "human",
          });
        } catch (err) {
          pushToast("error", err instanceof Error ? err.message : String(err));
          return;
        }
        const tryFocus = (attempt: number): void => {
          const target = hostEl.querySelector<HTMLElement>(
            `.pm-page-zone-${detail.slot}[data-page-number="${detail.pageNumber}"] [contenteditable='true']`
          );
          if (target) {
            target.focus();
            const range = document.createRange();
            range.selectNodeContents(target);
            range.collapse(false);
            const sel = window.getSelection();
            if (sel) {
              sel.removeAllRanges();
              sel.addRange(range);
            }
            return;
          }
          if (attempt < 6) {
            requestAnimationFrame(() => tryFocus(attempt + 1));
          }
        };
        requestAnimationFrame(() => tryFocus(0));
      })();
    };
    hostEl.addEventListener(PAGE_ZONE_MINT_EVENT, onZoneMint as EventListener);
    return () => hostEl.removeEventListener(PAGE_ZONE_MINT_EVENT, onZoneMint as EventListener);
  }, [hostEl, pushToast]);

  const closeHeaderFooter = useCallback(() => {
    if (!hostEl) return;
    const focused = hostEl.querySelector<HTMLElement>(".pm-page-zone-focused [contenteditable='true']");
    if (focused) focused.blur();
    if (view) view.focus();
    setHfZoneFocus(null);
  }, [hostEl, view]);

  // Hidden file input reused by the H/F toolbar's "Insert image"
  // affordance. We can't reuse the body-side `imageInputRef` because
  // its onChange routes to `handleImageFile`, which targets the body
  // selection. The H/F image needs to land in the focused part via
  // `docx:insert-header-footer-image`, so it gets its own input.
  const hfImageInputRef = useRef<HTMLInputElement | null>(null);

  // Resolve which paragraph (and char offset) inside the focused H/F
  // zone the caret is in. We rely on `data-paragraph-id` stamped on
  // each `.pm-page-zone-line` by the rich rendering pass — that's the
  // only authoritative link from a DOM line back to a `Paragraph.id`
  // in the snapshot. When no selection exists (toolbar clicked
  // without first putting the caret in the zone), fall back to the
  // last line so newly-inserted fields land at the end of the part.
  const resolveFocusedHFLine = useCallback((): {
    paragraphId: string;
    offset: number;
    paragraphIndex: number;
  } | null => {
    if (!hostEl) return null;
    const zone = hostEl.querySelector<HTMLElement>(".pm-page-zone-focused [contenteditable='true']");
    if (!zone) return null;
    const sel = typeof window !== "undefined" ? window.getSelection() : null;
    let line: HTMLElement | null = null;
    let offset = 0;
    if (sel && sel.focusNode && zone.contains(sel.focusNode)) {
      let node: Node | null = sel.focusNode;
      while (node && node !== zone) {
        if (node instanceof HTMLElement && node.classList.contains("pm-page-zone-line")) {
          line = node;
          break;
        }
        node = node.parentNode;
      }
      if (line) {
        const range = document.createRange();
        range.setStart(line, 0);
        range.setEnd(sel.focusNode, sel.focusOffset);
        offset = range.toString().length;
      }
    }
    if (!line) {
      const lines = zone.querySelectorAll<HTMLElement>(".pm-page-zone-line");
      line = lines[lines.length - 1] ?? null;
      if (line) offset = (line.textContent ?? "").length;
    }
    if (!line) return null;
    const paragraphId = line.dataset.paragraphId;
    if (!paragraphId) return null;
    const lines = Array.from(zone.querySelectorAll<HTMLElement>(".pm-page-zone-line"));
    const paragraphIndex = lines.indexOf(line);
    return { paragraphId, offset, paragraphIndex: paragraphIndex < 0 ? 0 : paragraphIndex };
  }, [hostEl]);

  const insertHeaderFooterField = useCallback(
    async (field: "PAGE" | "NUMPAGES") => {
      const agent = agentRef.current;
      if (!agent) return;
      const focus = resolveFocusedHFLine();
      if (!focus) {
        pushToast("warn", "Click into the header or footer first.");
        return;
      }
      try {
        await agent.applyCommand({
          type: "docx:insert-page-number",
          payload: { paragraphId: focus.paragraphId, offset: focus.offset, field },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast, resolveFocusedHFLine]
  );

  const handleHeaderFooterImageFile = useCallback(
    async (file: File) => {
      const agent = agentRef.current;
      if (!agent || !hfZoneFocus?.partPath) {
        pushToast("warn", "Click into the header or footer first.");
        return;
      }
      const focus = resolveFocusedHFLine();
      try {
        const buffer = await file.arrayBuffer();
        // Default to a modest 96 px display size — Word's "Insert
        // picture" never inserts at the original raster pixel size
        // for H/F images either; users typically resize via the
        // sizing handles afterwards. We pick a square hint and let
        // the natural aspect ratio re-flow once typed sizing handles
        // ship.
        const probe = await probeImageDimensions(file).catch(() => null);
        const cssWidth = probe?.width ?? 96;
        const cssHeight = probe?.height ?? 96;
        await agent.applyCommand({
          type: "docx:insert-header-footer-image",
          payload: {
            partPath: hfZoneFocus.partPath,
            paragraphIndex: focus?.paragraphIndex,
            data: new Uint8Array(buffer),
            mimeType: file.type || "image/png",
            width: cssWidth,
            height: cssHeight,
            name: file.name,
          },
          source: "human",
        });
        pushToast("info", `Inserted ${file.name || "image"} into ${hfZoneFocus.slot}.`);
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [hfZoneFocus, pushToast, resolveFocusedHFLine]
  );

  const toggleSectionDifferentFirst = useCallback(
    async (enabled: boolean) => {
      const agent = agentRef.current;
      if (!agent) return;
      // Target the section under the caret. The handler walks forward
      // from `paragraphIndex` to the next `<w:sectPr>` (or the trailing
      // implicit section), so any paragraph inside the section is fine.
      const v = mountRef.current?.view;
      const paraIdx = v ? currentParagraphIndex(v.state) : 0;
      try {
        await agent.applyCommand({
          type: "docx:set-section-different-first",
          payload: { paragraphIndex: paraIdx, enabled },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const submitHyperlink = useCallback(
    async (next: { url: string; text: string }) => {
      const agent = agentRef.current;
      if (!agent || !linkRequest) return;
      const url = normaliseUrl(next.url);
      try {
        await agent.applyCommand({
          type: "docx:insert-hyperlink",
          payload: {
            paragraphId: linkRequest.paragraphId,
            range: linkRequest.range,
            url,
          },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      } finally {
        setLinkRequest(null);
      }
    },
    [linkRequest, pushToast]
  );

  const openHyperlinkPopover = useCallback(() => {
    const mount = mountRef.current;
    if (!mount) return;
    const view = mount.view;
    if (view.state.selection.empty) {
      pushToast("info", "Select some text to link.");
      return;
    }
    const paragraphId = currentParagraphId(view.state);
    if (!paragraphId) return;
    const range = pmSelectionToRange(view.state);
    if (range.start.paragraph !== range.end.paragraph) {
      pushToast("warn", "Hyperlinks must stay within one paragraph.");
      return;
    }
    const selectionText = view.state.doc.textBetween(
      view.state.selection.from,
      view.state.selection.to,
      " ",
      " "
    );
    let anchor: { left: number; top: number; bottom: number } | null = null;
    try {
      const coords = view.coordsAtPos(view.state.selection.from);
      anchor = { left: coords.left, top: coords.top, bottom: coords.bottom };
    } catch {
      anchor = null;
    }
    setLinkRequest({
      paragraphId,
      range: { start: range.start.offset, end: range.end.offset },
      selectionText,
      anchor,
      initialUrl: "",
    });
  }, [pushToast]);

  const surfaceUnsupported = useCallback(
    (label: string) => {
      pushToast("warn", `${label} is not yet supported in this build.`);
    },
    [pushToast]
  );

  const setAlignment = useCallback(
    async (alignment: AlignmentValue) => {
      const agent = agentRef.current;
      const mount = mountRef.current;
      if (!agent || !mount) return;
      const paragraphId = currentParagraphId(mount.view.state);
      if (!paragraphId) {
        pushToast("info", "Place the caret in a paragraph first.");
        return;
      }
      try {
        await agent.applyCommand({
          type: "docx:set-paragraph-alignment",
          payload: { paragraphId, alignment },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  /**
   * F1-UI — toolbar entry point for `docx:insert-footnote`. Splices a
   * `FootnoteReferenceLeaf` at the caret's flat-text offset inside the
   * paragraph that owns the selection and appends the matching
   * `<w:footnote>` to `footnotesPart`. The caret stays where it is —
   * editing the footnote body in place is deferred until the bottom-of-
   * page lane lands a contenteditable surface.
   */
  const insertFootnote = useCallback(async () => {
    const agent = agentRef.current;
    const mount = mountRef.current;
    if (!agent || !mount) return;
    const paragraphId = currentParagraphId(mount.view.state);
    if (!paragraphId) {
      pushToast("info", "Place the caret in a paragraph first.");
      return;
    }
    const caret = pmPositionToDocx(mount.view.state, mount.view.state.selection.from);
    try {
      await agent.applyCommand({
        type: "docx:insert-footnote",
        payload: { paragraphId, offset: caret.offset },
        source: "human",
      });
    } catch (err) {
      pushToast("error", err instanceof Error ? err.message : String(err));
    }
  }, [pushToast]);

  const adjustIndent = useCallback(
    async (deltaTwips: number) => {
      const agent = agentRef.current;
      const mount = mountRef.current;
      if (!agent || !mount) return;
      const paragraphId = currentParagraphId(mount.view.state);
      if (!paragraphId) {
        pushToast("info", "Place the caret in a paragraph first.");
        return;
      }
      try {
        await agent.applyCommand({
          type: "docx:set-paragraph-indent",
          payload: { paragraphId, deltaTwips },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const setParagraphSpacing = useCallback(
    async (patch: {
      line?: number | null;
      lineRule?: "auto" | "exact" | "atLeast" | null;
      before?: number | null;
      after?: number | null;
    }) => {
      const agent = agentRef.current;
      const mount = mountRef.current;
      if (!agent || !mount) return;
      const paragraphId = currentParagraphId(mount.view.state);
      if (!paragraphId) {
        pushToast("info", "Place the caret in a paragraph first.");
        return;
      }
      try {
        await agent.applyCommand({
          type: "docx:set-paragraph-spacing",
          payload: { paragraphId, ...patch },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const toggleList = useCallback(
    async (kind: "bullet" | "ordered") => {
      const agent = agentRef.current;
      const mount = mountRef.current;
      if (!agent || !mount) return;
      const paragraphId = currentParagraphId(mount.view.state);
      if (!paragraphId) {
        pushToast("info", "Place the caret in a paragraph first.");
        return;
      }
      try {
        await agent.applyCommand({
          type: "docx:apply-list-format",
          payload: {
            paragraphId,
            format: kind === "bullet" ? "bullet" : "decimal",
            ilvl: 0,
          },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  /**
   * 9b — Word-parity list-level setter. Used by both the toolbar's
   * `ListLevelMenu` and (indirectly) the Tab/Shift+Tab handler in
   * `wordShortcutsKeymapPlugin`. Preserves the existing `numId` and
   * only bumps `ilvl`, so the list keeps its definition (bullet vs
   * numbered, restart rules, custom level text).
   *
   * Falls back to a toast when the caret isn't in a list — the menu
   * is disabled in that case, but the palette runner can still fire.
   */
  const setListLevel = useCallback(
    async (ilvl: number) => {
      const agent = agentRef.current;
      const mount = mountRef.current;
      if (!agent || !mount) return;
      const paragraphId = currentParagraphId(mount.view.state);
      if (!paragraphId) {
        pushToast("info", "Place the caret in a list paragraph first.");
        return;
      }
      const snap = agent.getSnapshot();
      const para = snap.root.body.find(
        (b): b is typeof b & { kind: "paragraph" } => b.kind === "paragraph" && b.id === paragraphId,
      );
      const numbering = para?.properties.numbering;
      if (!numbering) {
        pushToast("info", "Place the caret in a list paragraph first.");
        return;
      }
      const clamped = Math.max(0, Math.min(8, ilvl));
      try {
        await agent.applyCommand({
          type: "docx:set-paragraph-list",
          payload: { paragraphId, numId: numbering.numId, ilvl: clamped },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast],
  );

  const insertImageFromFile = useCallback(() => {
    imageInputRef.current?.click();
  }, []);

  // ── References tab — Bookmarks ──────────────────────────────────────────
  //
  // Word's References > Bookmark dialog. We dispatch the typed
  // `docx:insert-bookmark` / `docx:delete-bookmark` commands behind a
  // small modal that lists the existing bookmarks and validates the
  // identifier rules. The "Go to" affordance moves the caret to the
  // bookmark's start offset (translated from flat-text into PM
  // positions via the same walk used by Find & Replace). The dialog
  // itself is presentational — `BookmarkDialog.tsx`.
  const bookmarkRows: ReadonlyArray<BookmarkRow> = useMemo(() => {
    void uiTick;
    const a = agent;
    if (!a) return [];
    const list = listBookmarks(a.getSnapshot().root);
    return list.map((b) => ({
      name: b.name,
      paragraphId: b.paragraphId,
      startOffset: b.startOffset,
      endOffset: b.endOffset,
    }));
  }, [agent, uiTick]);

  const handleAddBookmark = useCallback(
    async (name: string) => {
      const a = agentRef.current;
      const m = mountRef.current;
      if (!a || !m) return;
      const paragraphId = currentParagraphId(m.view.state);
      if (!paragraphId) {
        pushToast("info", "Place the caret in a body paragraph first.");
        return;
      }
      // Selection range, clamped to the active paragraph. Cross-paragraph
      // bookmarks aren't supported in v1 — see the handler docstring.
      const range = pmSelectionToRange(m.view.state);
      let startOffset = 0;
      let endOffset = 0;
      if (range.start.paragraph === range.end.paragraph) {
        startOffset = Math.min(range.start.offset, range.end.offset);
        endOffset = Math.max(range.start.offset, range.end.offset);
      } else {
        // Selection spans multiple paragraphs — anchor at the caret in
        // the focused paragraph as a zero-length bookmark, matching
        // Word's fallback when "Insert > Bookmark" is invoked with a
        // multi-paragraph selection.
        const here = pmPositionToDocx(m.view.state, m.view.state.selection.from);
        startOffset = here.offset;
        endOffset = here.offset;
      }
      try {
        await a.applyCommand({
          type: "docx:insert-bookmark",
          payload: { name, paragraphId, startOffset, endOffset },
          source: "human",
        });
        pushToast("success", `Bookmark "${name}" added.`);
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast],
  );

  const handleDeleteBookmark = useCallback(
    async (name: string) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        await a.applyCommand({
          type: "docx:delete-bookmark",
          payload: { name },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast],
  );

  const handleGoToBookmark = useCallback(
    (b: BookmarkRow) => {
      const m = mountRef.current;
      if (!m) return;
      const v = m.view;
      let foundPos: number | null = null;
      v.state.doc.descendants((node, nodePos) => {
        if (foundPos !== null) return false;
        if (node.type.name === "paragraph" && node.attrs.paragraphId === b.paragraphId) {
          // PM positions are 1-based inside the paragraph; the start
          // marker sits at `nodePos`, the first text position at
          // `nodePos + 1`. The bookmark's flat offset matches the
          // same flat-text walk PM uses for `currentParagraphIndex`.
          foundPos = nodePos + 1 + b.startOffset;
          return false;
        }
        return true;
      });
      if (foundPos === null) return;
      const pos = Math.min(Math.max(foundPos, 1), v.state.doc.content.size - 1);
      const tr = v.state.tr.setSelection(TextSelection.create(v.state.doc, pos, pos));
      v.dispatch(tr.scrollIntoView());
      v.focus();
    },
    [],
  );

  const handleOpenBookmarkDialog = useCallback(() => {
    setBookmarkDialogOpen(true);
  }, []);

  // ── References tab — Table of contents ──────────────────────────────────
  //
  // Word's TOC field is a single `<w:fldSimple instr=" TOC \\o \"1-3\" "/>`
  // whose cached body is a sequence of TOC1/TOC2/… styled paragraphs
  // pointing at every heading. Implementing the field-cache plumbing is
  // multi-week work; we ship a pragmatic alternative now that mirrors
  // Word's *result* (a list of styled heading references) so users get
  // a working button today and the field-code upgrade slots in later.
  //
  // Strategy:
  //   - Walk the body for paragraphs whose style id starts with
  //     `Heading` (Heading1..Heading9). Numeric suffix → TOC level.
  //   - Synthesize one paragraph per heading, prefixed with the
  //     heading's full text. Paragraph style is `TOCN` so Word picks
  //     up the matching style from styles.xml when reopening.
  //   - "insert" splices the new paragraphs above the caret;
  //     "update" rebuilds the existing block, identified by a magic
  //     marker paragraph (`<TOC_PLACEHOLDER>` styled `TOCHeading`).
  //
  // Caveat surfaced in the inventory: round-trip preserves the
  // synthesized paragraphs but does NOT yet produce a real TOC field;
  // re-opening in Word treats them as static text.
  const handleInsertOrUpdateToc = useCallback(
    async (mode: "insert" | "update") => {
      const a = agentRef.current;
      const m = mountRef.current;
      if (!a || !m) return;
      const snap = a.getSnapshot();
      const headings = collectTocHeadings(snap);
      if (headings.length === 0) {
        pushToast("info", "No Heading 1-9 paragraphs found. Add headings to populate a TOC.");
        return;
      }
      try {
        if (mode === "update") {
          await rebuildTocBlock(a, snap, headings);
        } else {
          const insertAt = currentParagraphIndex(m.view.state);
          await insertTocBlock(a, insertAt, headings);
        }
        pushToast("success", `Table of contents ${mode === "update" ? "updated" : "inserted"} (${headings.length} entries).`);
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast],
  );

  // ── References tab — Caption ────────────────────────────────────────────
  //
  // Word's Insert > Caption dialog inserts a new paragraph styled
  // `Caption` and seeds it with a label + autonumber + free text. We
  // open a tiny `window.prompt` for the label so the button works
  // today; the dialog upgrade lands behind a follow-up plan.
  const handleInsertCaption = useCallback(async () => {
    const a = agentRef.current;
    const m = mountRef.current;
    if (!a || !m) return;
    const text = window.prompt("Caption text", "Figure 1: ");
    if (text === null) return;
    const trimmed = text.trim();
    if (trimmed.length === 0) return;
    const insertAt = currentParagraphIndex(m.view.state) + 1;
    try {
      await a.applyCommand({
        type: "docx:insert-paragraph",
        payload: { at: { paragraph: insertAt }, text: trimmed, styleId: "Caption" },
        source: "human",
      });
    } catch (err) {
      pushToast("error", err instanceof Error ? err.message : String(err));
    }
  }, [pushToast]);

  // ── References tab — Cross-reference ────────────────────────────────────
  //
  // Word's Insert > Cross-reference dialog wraps `<w:fldSimple
  // instr=" REF Bookmark1 \\h "/>` around the bookmark's resolved
  // text. Without the field-cache plumbing we produce plain text in
  // the form `→ {name}` so the visual link is at least obvious; the
  // bookmark itself is preserved for the future field upgrade.
  const handleInsertCrossReference = useCallback(async () => {
    const a = agentRef.current;
    const m = mountRef.current;
    if (!a || !m) return;
    const list = listBookmarks(a.getSnapshot().root);
    if (list.length === 0) {
      pushToast("info", "Insert a bookmark first via References > Bookmark.");
      return;
    }
    const names = list.map((b) => b.name);
    const picked = window.prompt(
      `Cross-reference to which bookmark?\n\nAvailable: ${names.join(", ")}`,
      names[0] ?? ""
    );
    if (!picked) return;
    const trimmed = picked.trim();
    if (!names.includes(trimmed)) {
      pushToast("warn", `No bookmark "${trimmed}" — insert one first.`);
      return;
    }
    const here = pmPositionToDocx(m.view.state, m.view.state.selection.from);
    try {
      await a.applyCommand({
        type: "docx:insert-text",
        payload: { at: here, text: ` → ${trimmed}` },
        source: "human",
      });
    } catch (err) {
      pushToast("error", err instanceof Error ? err.message : String(err));
    }
  }, [pushToast]);

  // ── Design tab — Session-scoped visual properties ───────────────────────
  //
  // Page color, borders, and watermark all need typed model fields and
  // serializer plumbing to round-trip into OOXML. Until those land we
  // expose them as **session-only visual effects** on the editor card
  // via CSS variables. The buttons work, the effect is visible, and
  // the inventory clearly labels the scope. A `null` value clears.
  const handleOpenPageColorPicker = useCallback(() => {
    const next = window.prompt("Page color (hex, e.g. FFF2CC). Empty clears.", "");
    if (next === null) return;
    const v = next.trim().replace(/^#/, "");
    const root = scrollEl;
    if (!root) return;
    if (v.length === 0) {
      root.style.removeProperty("--pm-page-fill");
      pushToast("info", "Page color cleared.");
    } else if (/^[0-9A-Fa-f]{6}$/.test(v)) {
      root.style.setProperty("--pm-page-fill", `#${v.toUpperCase()}`);
      pushToast("success", `Page color set to #${v.toUpperCase()} (visual only — round-trip pending).`);
    } else {
      pushToast("warn", "Color must be 6 hex digits.");
    }
  }, [scrollEl, pushToast]);

  const handleOpenPageBordersDialog = useCallback(() => {
    const root = scrollEl;
    if (!root) return;
    const current = root.style.getPropertyValue("--pm-page-border") || "none";
    const next = window.prompt(
      "Page borders. Use a CSS shorthand like '1px solid #000', or 'none' to clear.",
      current === "none" ? "1px solid #999" : current,
    );
    if (next === null) return;
    const trimmed = next.trim();
    if (trimmed.length === 0 || trimmed === "none") {
      root.style.removeProperty("--pm-page-border");
      pushToast("info", "Page borders cleared.");
    } else {
      root.style.setProperty("--pm-page-border", trimmed);
      pushToast("success", "Page borders applied (visual only — round-trip pending).");
    }
  }, [scrollEl, pushToast]);

  const handleOpenWatermarkDialog = useCallback(() => {
    const root = scrollEl;
    if (!root) return;
    const text = window.prompt("Watermark text. Empty clears.", "DRAFT");
    if (text === null) return;
    const trimmed = text.trim();
    if (trimmed.length === 0) {
      root.style.removeProperty("--pm-page-watermark");
      pushToast("info", "Watermark cleared.");
    } else {
      root.style.setProperty("--pm-page-watermark", `"${trimmed.replace(/"/g, '\\"')}"`);
      pushToast("success", `Watermark set to "${trimmed}" (visual only — round-trip pending).`);
    }
  }, [scrollEl, pushToast]);

  const handleOpenThemePicker = useCallback(() => {
    const root = scrollEl;
    if (!root) return;
    const themes: Record<string, { fill: string; accent: string }> = {
      Default: { fill: "#FFFFFF", accent: "#1F4E79" },
      Slate: { fill: "#F4F6F8", accent: "#334155" },
      Warmth: { fill: "#FFF8F0", accent: "#9A3412" },
      Forest: { fill: "#F1F8F2", accent: "#166534" },
    };
    const names = Object.keys(themes);
    const picked = window.prompt(`Document theme. One of: ${names.join(", ")}.`, "Default");
    if (!picked) return;
    const t = themes[picked];
    if (!t) {
      pushToast("warn", `Unknown theme "${picked}".`);
      return;
    }
    root.style.setProperty("--pm-page-fill", t.fill);
    root.style.setProperty("--pm-theme-accent", t.accent);
    pushToast("success", `Theme "${picked}" applied (visual only — round-trip pending).`);
  }, [scrollEl, pushToast]);

  /**
   * Open the {@link EmbeddedXlsxModal} for the embedded workbook
   * referenced by the supplied PM image attribute envelope. The
   * envelope may carry either:
   *   - `{ chart: { chartPartPath } }` for typed chart drawings, or
   *   - `{ embeddedSpreadsheet: <embeddingPartPath> }` for OLE
   *     spreadsheet leaves.
   *
   * Anything else (regular inline images, opaque drawings) is ignored
   * — those don't have an editable embedded workbook.
   */
  const openEmbeddedEditor = useCallback(
    async (envelope: unknown) => {
      const agent = agentRef.current;
      if (!agent) return;
      let chartPartPath: string | undefined;
      let embeddingPartPath: string | undefined;
      let title = "Edit data";
      if (envelope && typeof envelope === "object") {
        const env = envelope as {
          chart?: { chartPartPath?: unknown };
          embeddedSpreadsheet?: unknown;
        };
        if (env.chart && typeof env.chart.chartPartPath === "string") {
          chartPartPath = env.chart.chartPartPath;
          title = "Edit chart data";
        } else if (typeof env.embeddedSpreadsheet === "string") {
          embeddingPartPath = env.embeddedSpreadsheet;
          title = "Edit spreadsheet";
        }
      }
      if (!chartPartPath && !embeddingPartPath) return;
      try {
        const ref = resolveEmbeddedXlsxRef({
          source: { kind: "docx", agent },
          ...(chartPartPath ? { chartPartPath } : { embeddingPartPath: embeddingPartPath! }),
        });
        if (!ref) {
          pushToast("info", "This object has no embedded workbook to edit.");
          return;
        }
        const bytes = await readEmbeddedXlsxBytes({
          source: { kind: "docx", agent },
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
          title,
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const handleEmbeddedXlsxSave = useCallback(
    async (result: {
      readonly bytes: Uint8Array;
      readonly grid: ReadonlyArray<ReadonlyArray<string | number | null>>;
    }) => {
      const agent = agentRef.current;
      const ctx = editingEmbed;
      setEditingEmbed(null);
      if (!agent || !ctx) return;
      try {
        await agent.applyCommand({
          type: "docx:update-spreadsheet",
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
            await agent.applyCommand({
              type: "docx:set-chart-data",
              payload: {
                chartPartPath: ctx.chartPartPath,
                categories: chartUpdate.categories,
                series: chartUpdate.series,
              },
              source: "human",
            });
          }
        }
        pushToast("info", "Saved embedded data.");
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [editingEmbed, pushToast]
  );

  const handleXlsxPickerSubmit = useCallback(
    async (result: XlsxRangePickerResult) => {
      const agent = agentRef.current;
      const mount = mountRef.current;
      setXlsxPickerOpen(null);
      if (!agent) return;
      const idx = mount ? currentParagraphIndex(mount.view.state) : 0;
      try {
        await applyXlsxEmbed({
          target: { kind: "docx", agent, paragraphIndex: idx },
          snapshot: result.snapshot,
          mode: result.mode,
        });
        const label =
          result.mode === "live" ? "embedded spreadsheet" : result.mode === "chart" ? "chart" : "table";
        pushToast("info", `Inserted ${label} from xlsx.`);
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const insertTable = useCallback(
    async (rows: number, cols: number) => {
      const agent = agentRef.current;
      const mount = mountRef.current;
      if (!agent) return;
      const idx = mount ? currentParagraphIndex(mount.view.state) : 0;
      try {
        await agent.applyCommand({
          type: "docx:insert-table",
          payload: {
            at: { paragraph: idx + 1, run: 0, offset: 0 },
            rows,
            cols,
          },
          source: "human",
        });
        pushToast("info", `Inserted ${rows} × ${cols} table.`);
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  // B11 — Insert section break. The section break is appended at the
  // end of the paragraph the caret currently sits in (mirrors Word:
  // the typed flow position is converted into a "before the next
  // block" insertion point).
  const insertSectionBreak = useCallback(
    async (type: "nextPage" | "continuous" | "evenPage" | "oddPage") => {
      const agent = agentRef.current;
      const mount = mountRef.current;
      if (!agent) return;
      const idx = mount ? currentParagraphIndex(mount.view.state) : 0;
      try {
        await agent.applyCommand({
          type: "docx:insert-section-break",
          payload: { paragraphIndex: idx, type },
          source: "human",
        });
        pushToast("info", `Inserted ${SECTION_BREAK_LABEL[type]} section break.`);
      } catch (err) {
        if (err instanceof NotImplementedError) {
          pushToast("warn", "Not yet supported in this build.");
          return;
        }
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const handleTableInsertRow = useCallback(
    async (tableId: string, where: "top" | "bottom") => {
      const agent = agentRef.current;
      if (!agent) return;
      const snap = agent.getSnapshot();
      const table = snap.root.body.find((b) => b.kind === "table" && b.id === tableId);
      const rowsLen = table && table.kind === "table" ? table.rows.length : 0;
      const at = where === "top" ? 0 : rowsLen;
      try {
        await agent.applyCommand({
          type: "docx:insert-row",
          payload: { tableId, at },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const handleTableInsertColumn = useCallback(
    async (tableId: string, where: "start" | "end") => {
      const agent = agentRef.current;
      if (!agent) return;
      const snap = agent.getSnapshot();
      const table = snap.root.body.find((b) => b.kind === "table" && b.id === tableId);
      const colsLen = table && table.kind === "table" ? table.grid.length : 0;
      const at = where === "start" ? 0 : colsLen;
      try {
        await agent.applyCommand({
          type: "docx:insert-column",
          payload: { tableId, at },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const handleDeleteTableRow = useCallback(
    async (tableId: string, row: number) => {
      const agent = agentRef.current;
      if (!agent) return;
      try {
        await agent.applyCommand({
          type: "docx:delete-row",
          payload: { tableId, at: row },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const handleDeleteTableColumn = useCallback(
    async (tableId: string, column: number) => {
      const agent = agentRef.current;
      if (!agent) return;
      try {
        await agent.applyCommand({
          type: "docx:delete-column",
          payload: { tableId, at: column },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const handleDeleteTable = useCallback(
    async (tableId: string) => {
      const agent = agentRef.current;
      if (!agent) return;
      try {
        await agent.applyCommand({
          type: "docx:delete-table",
          payload: { tableId },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const handleSetCellShading = useCallback(
    async (tableId: string, row: number, column: number, fill: string | null) => {
      const agent = agentRef.current;
      if (!agent) return;
      try {
        await agent.applyCommand({
          type: "docx:set-cell-shading",
          payload: { tableId, row, column, fill },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const handleSetCellAlignment = useCallback(
    async (
      tableId: string,
      row: number,
      column: number,
      vAlign: "top" | "center" | "bottom" | null
    ) => {
      const agent = agentRef.current;
      if (!agent) return;
      try {
        await agent.applyCommand({
          type: "docx:set-cell-alignment",
          payload: { tableId, row, column, vAlign },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const handleSetRowHeight = useCallback(
    async (
      tableId: string,
      row: number,
      heightTwips: number | null,
      rule?: "auto" | "exact" | "atLeast"
    ) => {
      const agent = agentRef.current;
      if (!agent) return;
      try {
        await agent.applyCommand({
          type: "docx:set-row-height",
          payload: { tableId, row, heightTwips, ...(rule ? { rule } : {}) },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const handleSetColumnWidth = useCallback(
    async (tableId: string, column: number, widthTwips: number) => {
      const agent = agentRef.current;
      if (!agent) return;
      try {
        await agent.applyCommand({
          type: "docx:set-column-width",
          payload: { tableId, column, widthTwips },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const handleMergeCellsHorizontal = useCallback(
    async (tableId: string, row: number, fromColumn: number, toColumn: number) => {
      const agent = agentRef.current;
      if (!agent) return;
      try {
        await agent.applyCommand({
          type: "docx:merge-cells-horizontal",
          payload: { tableId, row, fromColumn, toColumn },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  // B6 — image manipulation. Resize commits a single
  // `docx:set-image-properties` with the rounded final dimensions;
  // alt-text edit opens the dialog seeded with the current value;
  // delete uses `docx:delete-range` against the image's enclosing
  // run so the image leaf is removed cleanly without disturbing
  // surrounding runs.
  const handleImageResize = useCallback(
    async (imageId: string, widthPx: number, heightPx: number) => {
      const agent = agentRef.current;
      if (!agent) return;
      try {
        await agent.applyCommand({
          type: "docx:set-image-properties",
          payload: { imageId, widthPx, heightPx },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const handleImageEditAlt = useCallback((info: SelectedImageInfo) => {
    setAltTextRequest({ imageId: info.imageId, initial: info.altText });
  }, []);

  const submitAltText = useCallback(
    async (imageId: string, altText: string | null) => {
      const agent = agentRef.current;
      if (!agent) return;
      try {
        await agent.applyCommand({
          type: "docx:set-image-properties",
          payload: { imageId, altText },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const handleImageDelete = useCallback(
    async (imageId: string) => {
      const agent = agentRef.current;
      if (!agent) return;
      try {
        await agent.applyCommand({
          type: "docx:delete-image",
          payload: { imageId },
          source: "human",
        });
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const handleImageFile = useCallback(
    async (file: File) => {
      const agent = agentRef.current;
      const mount = mountRef.current;
      if (!agent) {
        pushToast("warn", "Document is still loading.");
        return;
      }
      try {
        await insertImageIntoDocx(agent, file, mount?.view.state ?? null);
        pushToast("info", `Inserted ${file.name || "image"}.`);
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  // Drag-drop + paste handlers for image files. We intercept the events
  // at the host DOM level *before* ProseMirror sees them so that:
  //   - a dropped file lands as a typed `docx:insert-image` instead of
  //     PM's default "paste as text" behaviour;
  //   - a pasted screenshot (clipboard image) gets inserted at the
  //     caret instead of being silently dropped on the floor.
  // Non-file drops (text, regular HTML pastes) fall through to PM.
  useEffect(() => {
    if (!hostEl) return;
    const onDragOver = (e: DragEvent) => {
      if (!e.dataTransfer) return;
      const hasFile = Array.from(e.dataTransfer.items ?? []).some((it) => it.kind === "file");
      if (hasFile) {
        e.preventDefault();
        e.dataTransfer.dropEffect = "copy";
      }
    };
    const onDrop = (e: DragEvent) => {
      const file = pickImageFile(e.dataTransfer?.files);
      if (!file) return;
      e.preventDefault();
      e.stopPropagation();
      void handleImageFile(file);
    };
    const onPaste = (e: ClipboardEvent) => {
      // 1) Cross-format embed (XLSX → DOCX table). Gated on the
      //    NEXT_PUBLIC_OAI_EMBED flag so existing PM HTML-table
      //    paste behaviour stays the default in production.
      if (isEmbedEnabled()) {
        const raw = e.clipboardData?.getData(EMBED_MIME);
        const env = parseEnvelope(raw);
        if (env && env.payload.kind === "xlsx-range") {
          const agent = agentRef.current;
          const mount = mountRef.current;
          if (!agent) return;
          e.preventDefault();
          e.stopPropagation();
          const paragraphIndex = mount ? currentParagraphIndex(mount.view.state) : 0;
          // Alt held → embed as a live OLE Excel object instead of a
          // materialised table. Mirrors PowerPoint's "Paste Special →
          // Microsoft Excel Worksheet Object" shortcut so power users
          // can opt-in without round-tripping through a dialog. Alt
          // state is tracked separately because `ClipboardEvent`
          // doesn't expose modifier keys directly.
          const mode = isAltKeyPressed() ? "live" : "materialized";
          void (async () => {
            try {
              await applyXlsxRangeToDocx({
                agent,
                snapshot:
                  env.payload.kind === "xlsx-range"
                    ? env.payload.snapshot
                    : (() => {
                        throw new Error("unreachable");
                      })(),
                paragraphIndex: Math.max(0, paragraphIndex),
                mode,
              });
            } catch (err) {
              pushToast("error", err instanceof Error ? err.message : String(err));
            }
          })();
          return;
        }
      }
      // 2) Image paste — screenshots, drag-and-drop file pastes.
      const file = pickImageFile(e.clipboardData?.files);
      if (!file) return;
      e.preventDefault();
      void handleImageFile(file);
    };
    // Double-click on a chart drawing or embedded spreadsheet pops
    // the {@link EmbeddedXlsxModal}. We hook at the host level so PM
    // doesn't see the event first and try to start a text selection
    // inside the atom node — the dblclick on an `image` PM node is
    // otherwise just dead air. The PM image's host DOM is an `<img>`
    // (or a `[image]` placeholder span) without our `drawingJson`
    // attribute on the DOM, so we resolve the position through
    // ProseMirror and read the typed node attrs.
    const onDblClick = (e: MouseEvent) => {
      const view = mountRef.current?.view;
      if (!view) return;
      const pos = view.posAtCoords({ left: e.clientX, top: e.clientY });
      if (!pos) return;
      const node = view.state.doc.nodeAt(pos.pos);
      const raw = node?.type.name === "image" ? node.attrs.drawingJson : null;
      if (typeof raw !== "string" || raw.length === 0) return;
      let envelope: unknown;
      try {
        envelope = JSON.parse(raw);
      } catch {
        return;
      }
      const env = envelope as { chart?: unknown; embeddedSpreadsheet?: unknown };
      if (!env.chart && !env.embeddedSpreadsheet) return;
      e.preventDefault();
      e.stopPropagation();
      void openEmbeddedEditor(envelope);
    };
    hostEl.addEventListener("dragover", onDragOver);
    hostEl.addEventListener("drop", onDrop);
    hostEl.addEventListener("paste", onPaste);
    hostEl.addEventListener("dblclick", onDblClick);
    const uninstallAlt = installAltKeyTracker();
    return () => {
      hostEl.removeEventListener("dragover", onDragOver);
      hostEl.removeEventListener("drop", onDrop);
      hostEl.removeEventListener("paste", onPaste);
      hostEl.removeEventListener("dblclick", onDblClick);
      uninstallAlt();
    };
  }, [hostEl, handleImageFile, pushToast, openEmbeddedEditor]);

  const scrollToComment = useCallback(
    (commentId: string) => {
      const host = hostEl;
      if (!host) return;
      const target = host.querySelector<HTMLElement>(
        `.pm-comment-mark[data-comment-id="${cssEscape(commentId)}"]`
      );
      if (!target) {
        pushToast("info", "Comment anchor is no longer in the document.");
        return;
      }
      target.scrollIntoView({ behavior: "smooth", block: "center" });
      target.classList.add("pm-comment-flash");
      window.setTimeout(() => target.classList.remove("pm-comment-flash"), 1600);
    },
    [pushToast, hostEl]
  );

  const acceptChange = useCallback(
    async (revisionId: string) => {
      const agent = agentRef.current;
      if (!agent) return;
      try {
        await agent.applyCommand({
          type: "docx:accept-change",
          payload: { revisionId },
          source: "human",
        });
        pushToast("info", "Change accepted.");
      } catch (err) {
        if (err instanceof NotImplementedError) {
          pushToast("warn", "Not yet supported in this build.");
          return;
        }
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const rejectChange = useCallback(
    async (revisionId: string) => {
      const agent = agentRef.current;
      if (!agent) return;
      try {
        await agent.applyCommand({
          type: "docx:reject-change",
          payload: { revisionId },
          source: "human",
        });
        pushToast("info", "Change rejected.");
      } catch (err) {
        if (err instanceof NotImplementedError) {
          pushToast("warn", "Not yet supported in this build.");
          return;
        }
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const acceptAllChanges = useCallback(async () => {
    const agent = agentRef.current;
    if (!agent) return;
    try {
      await agent.applyCommand({
        type: "docx:accept-all-changes",
        payload: {},
        source: "human",
      });
      pushToast("info", "Accepted all tracked changes.");
    } catch (err) {
      if (err instanceof NotImplementedError) {
        pushToast("warn", "Not yet supported in this build.");
        return;
      }
      pushToast("error", err instanceof Error ? err.message : String(err));
    }
  }, [pushToast]);

  const rejectAllChanges = useCallback(async () => {
    const agent = agentRef.current;
    if (!agent) return;
    try {
      await agent.applyCommand({
        type: "docx:reject-all-changes",
        payload: {},
        source: "human",
      });
      pushToast("info", "Rejected all tracked changes.");
    } catch (err) {
      if (err instanceof NotImplementedError) {
        pushToast("warn", "Not yet supported in this build.");
        return;
      }
      pushToast("error", err instanceof Error ? err.message : String(err));
    }
  }, [pushToast]);

  // Derive toolbar UI state from the current PM view (re-runs on uiTick).
  void uiTick;
  const snapshot = agent?.getSnapshot() ?? null;
  const currentParaIndex = view ? currentParagraphIndex(view.state) : 0;
  const activeStyle = snapshot ? paragraphStyle(snapshot, currentParaIndex) : "Normal";
  const currentSectionTitlePg = snapshot ? sectionTitlePgAt(snapshot, currentParaIndex) : false;
  // 9b — surface the current list level so `ListLevelMenu` can both
  // disable itself when not in a list AND highlight the active level.
  // Walks the body once per uiTick (cheap; bodies are sub-1k-paragraph
  // in practice). Returns null when the caret paragraph isn't a list
  // item, distinguishing "level 0 in a list" from "no list".
  const currentListLevel = snapshot ? listLevelAt(snapshot, currentParaIndex) : null;
  // P3.1 / W3 — toolbar dropdowns fall back through the typed style
  // cascade when no direct PM mark carries the attribute. The shared
  // text-formatting provider plumbs that through ActiveTextFormat so a
  // Heading1 paragraph surfaces its inherited 16pt / Calibri.
  // Provider is built once via useState's lazy initialiser. The
  // closure stashes mountRef/agentRef and reads `.current` only at
  // event-handler time (apply / getActive), which is safe — but the
  // React Compiler can't see through the closure boundary, so we
  // silence the rule for this construction site.
  /* eslint-disable react-hooks/refs */
  const [textFormatProvider] = useState(() => createDocxFormatProvider({ mountRef, agentRef, pushToast }));
  /* eslint-enable react-hooks/refs */
  const textFormatActive = computeDocxActive(view, snapshot);
  const activeAlignment = view ? currentParagraphAlignment(view.state) : null;
  const activeParagraphIndex = view ? currentParagraphIndex(view.state) : -1;
  const activeSpacing = computeActiveSpacing(snapshot, activeParagraphIndex);
  const activeIndentLeft = computeActiveIndentLeft(snapshot, activeParagraphIndex);
  const styleOptions = paragraphStyleOptions(snapshot, activeStyle);
  const trackedChangesCount = countTrackedChanges(snapshot);
  const protectionState = useMemo(
    () => readDocumentProtection(snapshot),
    [snapshot]
  );
  // Drive the contextual ribbon tabs ("Bildtools", "Tabellentools").
  // We re-derive on every render because `uiTick` already forces one
  // on every selection change, so this is essentially `O(1)` work
  // gated to the same cadence as the rest of the toolbar state.
  const selectedImage = view ? extractSelectedImage(view) : null;
  const selectedTableId = view ? extractSelectedTableId(view) : null;
  const lastSelectedTableIdRef = useRef<string | null>(null);
  useEffect(() => {
    if (lastSelectedTableIdRef.current !== selectedTableId) {
      lastSelectedTableIdRef.current = selectedTableId;
      setActiveTableCell({ row: 0, column: 0 });
    }
  }, [selectedTableId]);
  void commentParagraphIndex;

  // The side rail hosts the comments sidebar; tracked changes are no
  // longer surfaced here — `<TrackedChangesMargin>` paints them as
  // Word-style balloons in the page's right margin instead. The rail
  // therefore only earns its 320px column when the document actually
  // carries comments. The hover overlay over insertion spans is
  // mounted separately (see `<TrackedChangesHover>` below) so its
  // mouse listeners survive the rail being hidden.
  const hasComments = snapshot ? commentThreads(snapshot).length > 0 : false;
  // Build the shared CommentsProvider once per agent. The sidebar in
  // `@officeai/ui` calls `provider.threads()` on every render and goes
  // through the live agent snapshot, so we don't need to re-create the
  // provider when comments change — only when the underlying agent
  // instance does (i.e. on document load).
  /* eslint-disable react-hooks/refs -- `pushToast` wraps a state setter
     exposed via a stable callback; the provider only invokes it from
     event handlers (sidebar clicks), never during render. */
  const commentsProvider = useMemo(
    () =>
      agent
        ? createDocxCommentsProvider({
            agent,
            onScrollTo: scrollToComment,
            onToast: pushToast,
          })
        : null,
    [agent, scrollToComment, pushToast]
  );
  /* eslint-enable react-hooks/refs */

  const hasRevisions = snapshot ? collectRevisions(snapshot).length > 0 : false;
  const showRail = hasComments;
  // Width of the right-margin balloon column, used both to reserve
  // empty space next to the page card so the balloons stay visible
  // and to size the gap between page and balloons. Mirrors the
  // BALLOON_WIDTH + BALLOON_GUTTER_PX constants in TrackedChangesUI.
  const BALLOON_RESERVED_PX = 252;

  // ── Shell adapter wiring ──
  // B9 — Outline / Navigation panel.
  //
  // We surface every paragraph whose resolved style is a Heading 1..9,
  // the localised "Überschrift" variants, or the Title style. The
  // active entry is the last heading at-or-before the current caret
  // paragraph — this mirrors Word's Navigation pane behaviour where
  // the closest ancestor heading is highlighted so the user can
  // always see where they are without scrolling the outline manually.
  const outline = useMemo<ReadonlyArray<OutlineEntry>>(() => {
    if (!snapshot) return [];
    type OutlineSeed = { id: string; level: number; text: string; paraIndex: number };
    const seeds: OutlineSeed[] = [];
    snapshot.root.body.forEach((b, idx) => {
      if (b.kind !== "paragraph") return;
      const style = paragraphStyle(snapshot, idx);
      const m = /^(Heading|berschrift)(\d)$/.exec(style);
      const isTitle = style === "Title";
      let level: number | null = null;
      if (isTitle) level = 1;
      else if (m) level = Number(m[2]);
      if (level === null) return;
      const text = b.children
        .map((node) => {
          if (node.kind !== "run") return "";
          return node.children.map((c) => (c.kind === "text" ? c.text : "")).join("");
        })
        .join("")
        .trim();
      seeds.push({ id: `p-${idx}`, level, text, paraIndex: idx });
    });

    let activeIdx = -1;
    for (let i = seeds.length - 1; i >= 0; i--) {
      if (seeds[i]!.paraIndex <= currentParaIndex) {
        activeIdx = i;
        break;
      }
    }

    return seeds.map((s, i) => ({
      id: s.id,
      level: s.level,
      text: s.text,
      active: i === activeIdx,
      onActivate: () => {
        const host = hostEl;
        if (!host) return;
        const paras = host.querySelectorAll<HTMLElement>("p, h1, h2, h3, h4, h5, h6");
        const target = paras[s.paraIndex];
        if (target) target.scrollIntoView({ behavior: "smooth", block: "center" });
      },
    }));
  }, [snapshot, hostEl, currentParaIndex]);

  const findAdapter = useMemo<FindAdapter | undefined>(() => {
    if (!view) return undefined;
    const buildRegex = (q: string, opts: FindOptions): RegExp | null => {
      try {
        const flags = opts.caseSensitive ? "g" : "gi";
        const body = opts.regex ? q : q.replace(/[-/\\^$*+?.()|[\]{}]/g, "\\$&");
        const wrapped = opts.wholeWord ? `\\b${body}\\b` : body;
        return new RegExp(wrapped, flags);
      } catch {
        return null;
      }
    };
    return {
      findAll(query, opts) {
        if (query.length === 0) return [];
        const v = mountRef.current?.view;
        if (!v) return [];
        const re = buildRegex(query, opts);
        if (!re) return [];
        const text = v.state.doc.textBetween(0, v.state.doc.content.size, "\n", " ");
        const out: FindMatch[] = [];
        let m: RegExpExecArray | null;
        let count = 0;
        while ((m = re.exec(text)) !== null && count < 5000) {
          const start = m.index;
          const end = start + m[0].length;
          const preview = text.slice(Math.max(0, start - 24), Math.min(text.length, end + 24));
          out.push({ id: `${start}:${end}`, preview });
          count += 1;
          if (m[0].length === 0) re.lastIndex += 1;
        }
        return out;
      },
      gotoMatch(match) {
        const v = mountRef.current?.view;
        if (!v) return;
        const [start, end] = match.id.split(":").map(Number);
        // Walk PM doc to convert flat-text offset to PM positions.
        const doc = v.state.doc;
        let tStart = -1;
        let tEnd = -1;
        let acc = 0;
        doc.descendants((node, pos) => {
          if (!node.isText) return true;
          const len = node.text?.length ?? 0;
          if (tStart < 0 && acc + len >= start) tStart = pos + (start - acc);
          if (tEnd < 0 && acc + len >= end) tEnd = pos + (end - acc);
          acc += len;
          return true;
        });
        if (tStart < 0 || tEnd < 0) return;
        const tr = v.state.tr.setSelection(TextSelection.create(v.state.doc, tStart, tEnd));
        v.dispatch(tr.scrollIntoView());
      },
      replaceMatch(match, replacement) {
        const v = mountRef.current?.view;
        if (!v) return;
        const positions = textOffsetsToPmRange(v.state.doc, match);
        if (!positions) return;
        const { start, end } = positions;
        // Use the schema's text node so marks at the boundary collapse the
        // way Word would: the replacement inherits the marks at the start
        // of the original match (matches Word's "carry formatting"
        // behaviour). We deliberately strip block boundaries — Replace is
        // a flat-text operation; users use Insert for structural changes.
        if (replacement.length === 0) {
          const tr = v.state.tr.delete(start, end).setSelection(
            TextSelection.create(v.state.doc.resolve(start).doc, start, start),
          );
          v.dispatch(tr.scrollIntoView());
          return;
        }
        const marks = v.state.doc.resolve(start).marks();
        const node = v.state.schema.text(replacement, marks);
        const tr = v.state.tr.replaceWith(start, end, node);
        v.dispatch(tr.scrollIntoView());
      },
      replaceAll(query, replacement, opts) {
        const v = mountRef.current?.view;
        if (!v) return 0;
        if (query.length === 0) return 0;
        const re = buildRegex(query, opts);
        if (!re) return 0;
        const text = v.state.doc.textBetween(0, v.state.doc.content.size, "\n", " ");
        // Collect every match in flat-text coordinates first; mutating the
        // doc per match would invalidate the regex's lastIndex and shift
        // every subsequent position. We compute right-to-left applies so
        // earlier replacements don't shift later positions.
        const hits: Array<{ start: number; end: number }> = [];
        let m: RegExpExecArray | null;
        let safety = 0;
        while ((m = re.exec(text)) !== null && safety < 5000) {
          hits.push({ start: m.index, end: m.index + m[0].length });
          if (m[0].length === 0) re.lastIndex += 1;
          safety += 1;
        }
        if (hits.length === 0) return 0;

        // Map flat-text offsets to PM positions in one pass for the whole
        // doc — O(N) walk vs. O(N*hits) if we re-walked per match.
        const pmPositions = mapFlatOffsetsToPmPositions(v.state.doc, hits);
        if (pmPositions.length === 0) return 0;

        // Apply from the back so earlier replacements don't shift later
        // PM positions. One transaction = one undo step.
        let tr = v.state.tr;
        for (let i = pmPositions.length - 1; i >= 0; i--) {
          const range = pmPositions[i];
          if (!range) continue;
          const { start: s, end: e } = range;
          const marks = tr.doc.resolve(s).marks();
          if (replacement.length === 0) {
            tr = tr.delete(s, e);
          } else {
            tr = tr.replaceWith(s, e, v.state.schema.text(replacement, marks));
          }
        }
        v.dispatch(tr.scrollIntoView());
        return pmPositions.length;
      },
    };
  }, [view]);

  // 9b — User-authored footnote count (excludes the standard
  // separator at id=-1 and continuation at id=0). Drives both the
  // rail-tab badge AND the "show the tab at all" gate.
  const userFootnoteCount = useMemo(() => {
    const part = agent?.getSnapshot().root.footnotesPart;
    if (!part) return 0;
    return part.footnotes.filter((f) => f.id > 0).length;
  }, [agent, uiTick]); // eslint-disable-line react-hooks/exhaustive-deps

  const renderFootnotesPanel = useCallback((): React.ReactNode => {
    return (
      <FootnotesPanel
        agent={agent}
        snapshot={agent?.getSnapshot() ?? null}
        onScrollToReference={(footnoteId) => {
          // Best-effort: scroll the first matching reference into
          // view. ProseMirror nodes carry the footnoteId on the
          // `footnoteRef` attr; we walk the DOM to avoid coupling the
          // rail to PM internals.
          const surface = mountRef.current?.view.dom;
          if (!surface) return;
          const target = surface.querySelector<HTMLElement>(
            `[data-footnote-id="${footnoteId}"]`,
          );
          if (target) target.scrollIntoView({ behavior: "smooth", block: "center" });
        }}
      />
    );
  }, [agent]);

  const renderCommentsPanel = useCallback((): React.ReactNode => {
    if (!commentsProvider) {
      return (
        <div className="p-4 text-sm text-secondary">
          No comments. Select text and click Add comment in the toolbar.
        </div>
      );
    }
    return (
      <div className="flex min-h-0 flex-1 flex-col overflow-y-auto">
        <CommentsSidebar
          provider={commentsProvider}
          onScrollTo={scrollToComment}
          {...(authorIdentity ? { authorIdentity } : {})}
        />
      </div>
    );
  }, [commentsProvider, scrollToComment, authorIdentity]);

  // Palette is generated from the central docx action catalogue (see
  // packages/docx/src/actions/catalogue.ts) so labels/sections/shortcuts
  // never drift between Cmd+K and the CLI. The runners map only carries
  // the closure-bound side effect for each id; metadata flows from the
  // catalogue.
  const paletteCommands = useMemo<ReadonlyArray<PaletteCommand>>(() => {
    const runners: PaletteRunners = {
      "docx.add-comment": { run: () => openCommentComposer() },
      "docx.page-setup": { run: () => setPageSetupOpen(true) },
      "docx.insert-image": { run: insertImageFromFile },
      "docx.insert-footnote": { run: () => void insertFootnote() },
      "docx.insert-table-3x3": { run: () => void insertTable(3, 3) },
      "docx.insert-table-2x2": { run: () => void insertTable(2, 2) },
      "docx.insert-table-from-xlsx": { run: () => setXlsxPickerOpen("materialized") },
      "docx.insert-spreadsheet-from-xlsx": { run: () => setXlsxPickerOpen("live") },
      "docx.insert-chart-from-xlsx": { run: () => setXlsxPickerOpen("chart") },
      "docx.insert-hyperlink": { run: openHyperlinkPopover },
      "docx.toggle-marks": { run: handleToggleFormattingMarks },
      "docx.bullet-list": { run: () => void toggleList("bullet") },
      "docx.ordered-list": { run: () => void toggleList("ordered") },
      "docx.align-left": { run: () => void setAlignment("left") },
      "docx.align-center": { run: () => void setAlignment("center") },
      "docx.align-right": { run: () => void setAlignment("right") },
      "docx.align-justify": { run: () => void setAlignment("justify") },
      "docx.indent-increase": { run: () => void adjustIndent(360) },
      "docx.indent-decrease": { run: () => void adjustIndent(-360) },
      "docx.section-break-next-page": { run: () => void insertSectionBreak("nextPage") },
      "docx.section-break-continuous": { run: () => void insertSectionBreak("continuous") },
      "docx.set-mode-edit": { run: () => setEditMode("edit"), enabled: editMode !== "edit" },
      "docx.set-mode-suggest": { run: () => setEditMode("suggest"), enabled: editMode !== "suggest" },
      "docx.set-mode-view": { run: () => setEditMode("view"), enabled: editMode !== "view" },
      "docx.set-protection": { run: () => setProtectDocumentOpen(true) },
      "docx.insert-bookmark": { run: () => setBookmarkDialogOpen(true) },
      "docx.delete-bookmark": {
        run: () => {
          const a = agentRef.current;
          if (!a) return;
          const list = listBookmarks(a.getSnapshot().root);
          if (list.length === 0) {
            pushToast("info", "No bookmarks to delete.");
            return;
          }
          const names = list.map((b) => b.name);
          const picked = window.prompt(
            `Delete which bookmark? Available: ${names.join(", ")}`,
            names[0] ?? ""
          );
          if (picked) void handleDeleteBookmark(picked.trim());
        },
      },
      "docx.set-cell-shading": {
        run: () => {
          if (!selectedTableId) return;
          void handleSetCellShading(
            selectedTableId,
            activeTableCell.row,
            activeTableCell.column,
            "FFF2CC"
          );
        },
        enabled: selectedTableId !== null,
      },
      "docx.set-cell-alignment": {
        run: () => {
          if (!selectedTableId) return;
          void handleSetCellAlignment(
            selectedTableId,
            activeTableCell.row,
            activeTableCell.column,
            "center"
          );
        },
        enabled: selectedTableId !== null,
      },
      "docx.set-row-height": {
        run: () => {
          if (!selectedTableId) return;
          void handleSetRowHeight(selectedTableId, activeTableCell.row, 720, "atLeast");
        },
        enabled: selectedTableId !== null,
      },
      "docx.set-column-width": {
        run: () => {
          if (!selectedTableId) return;
          void handleSetColumnWidth(selectedTableId, activeTableCell.column, 1440);
        },
        enabled: selectedTableId !== null,
      },
      "docx.merge-cells-horizontal": {
        run: () => {
          if (!selectedTableId) return;
          void handleMergeCellsHorizontal(
            selectedTableId,
            activeTableCell.row,
            activeTableCell.column,
            activeTableCell.column + 1
          );
        },
        enabled: selectedTableId !== null,
      },
      "docx.delete-row": {
        run: () => {
          if (!selectedTableId) return;
          void handleDeleteTableRow(selectedTableId, activeTableCell.row);
        },
        enabled: selectedTableId !== null,
      },
      "docx.delete-column": {
        run: () => {
          if (!selectedTableId) return;
          void handleDeleteTableColumn(selectedTableId, activeTableCell.column);
        },
        enabled: selectedTableId !== null,
      },
      "docx.delete-table": {
        run: () => {
          if (!selectedTableId) return;
          void handleDeleteTable(selectedTableId);
        },
        enabled: selectedTableId !== null,
      },
    };
    return buildPaletteFromCatalogue(docxActions, runners, t);
  }, [
    t,
    adjustIndent,
    editMode,
    handleToggleFormattingMarks,
    insertFootnote,
    insertImageFromFile,
    insertSectionBreak,
    insertTable,
    openCommentComposer,
    openHyperlinkPopover,
    setAlignment,
    toggleList,
    selectedTableId,
    activeTableCell,
    handleSetCellShading,
    handleSetCellAlignment,
    handleSetRowHeight,
    handleSetColumnWidth,
    handleMergeCellsHorizontal,
    handleDeleteTableRow,
    handleDeleteTableColumn,
    handleDeleteTable,
  ]);

  const selectionText = useMemo<string>(() => {
    if (!docInfo) return "";
    const para = activeParagraphIndex >= 0 ? `Paragraph ${activeParagraphIndex + 1}` : "";
    const stats = `${docInfo.paragraphs} paragraphs · rev ${docInfo.revision}`;
    return para ? `${para} · ${stats}` : stats;
  }, [activeParagraphIndex, docInfo]);

  const tabFallback = useStableTabId("docx");
  const realtimeRoomId = useMemo<string | null>(() => {
    if (!agentReady) return null;
    // Explicit `null` from the host disables realtime entirely (used
    // for read-only previews). A non-empty string wins over both
    // `?room=` and `?src=` so embedding hosts can pin two browsers
    // viewing the same document (keyed e.g. by S3 object key) into
    // the same room without coordinating URLs.
    if (roomOverride === null) return null;
    if (typeof roomOverride === "string" && roomOverride.length > 0) {
      return `oai/docx/host/${roomOverride}`;
    }
    if (!tabFallback && !initialSource) return null;
    return roomIdForSource({
      product: "docx",
      src: initialSource?.url,
      tabFallback,
      explicitRoom: readExplicitRoomFromUrl(),
    });
  }, [agentReady, initialSource, tabFallback, roomOverride]);
  const realtimeRoom = useRealtimeRoom({
    roomId: realtimeRoomId,
    product: "docx",
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

  // Publish a DocxCursor into the awareness payload whenever the
  // ProseMirror selection changes. `uiTick` already bumps on every
  // `selectionchange`, so it's the cheapest re-trigger we have.
  const presenceCursor = useMemo(() => {
    if (!view) return null;
    const sel = view.state.selection;
    void uiTick;
    return { product: "docx" as const, head: sel.head, anchor: sel.anchor };
  }, [view, uiTick]);
  usePublishPresence({ room: realtimeRoom.room, cursor: presenceCursor });

  // Mirror the realtime identity into the local `authorIdentity` state
  // declared near the top so memoized renderers (e.g. the comments
  // panel) can read it without forward-referencing `realtimeRoom`.
  const liveIdentity = realtimeRoom.room?.identity ?? null;
  useEffect(() => {
    if (!liveIdentity) {
      setAuthorIdentity(null);
      return;
    }
    setAuthorIdentity({
      name: liveIdentity.name,
      id: liveIdentity.id,
      color: liveIdentity.color,
    });
  }, [liveIdentity]);

  // Bind the late `submitComment` impl now that `commentsProvider`
  // is in scope. The earlier `submitComment` useCallback (declared
  // next to the composer state) just dispatches through this ref —
  // see the comment on `submitCommentRef` for why.
  useEffect(() => {
    submitCommentRef.current = (text: string) => {
      if (!composer) return;
      if (!commentsProvider) return;
      const authorName = authorIdentity?.name ?? "You";
      void commentsProvider
        .add({
          author: authorName,
          text,
          anchor: { kind: "docx-range", paragraphIndex: 0, range: composer.range },
          ...(authorIdentity?.id ? { authorId: authorIdentity.id } : {}),
          ...(authorIdentity?.color ? { authorColor: authorIdentity.color } : {}),
        })
        .catch((err) => {
          pushToast("error", err instanceof Error ? err.message : String(err));
        })
        .finally(() => setComposer(null));
    };
  }, [composer, commentsProvider, authorIdentity, pushToast]);

  const adapter = useMemo<ProductAdapter>(
    () => ({
      product: "docx",
      filename: docName,
      saveState,
      comments: { openCount: hasComments ? (docInfo?.commentThreads ?? 0) : 0, resolvedCount: 0 },
      outline,
      selectionSummary: { text: selectionText },
      canOpen: true,
      hideOpen: hideLocalFileOpen,
      canSave: agentReady,
      canExport: agentReady,
      exportFormats: DOCX_EXPORT_FORMATS,
      onOpenFile: () => void handleOpenFile(),
      onSave: () => handleSave(),
      onExport: (format, options) => handleExport(format, options),
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
      renderFootnotesPanel,
      footnoteCount: userFootnoteCount,
      onAddComment: openCommentComposer,
    }),
    [
      agent,
      agentReady,
      docInfo?.commentThreads,
      docName,
      findAdapter,
      handleExport,
      handleOpenFile,
      handleSave,
      hasComments,
      hideLocalFileOpen,
      openCommentComposer,
      outline,
      paletteCommands,
      renderCommentsPanel,
      renderFootnotesPanel,
      saveState,
      selectionText,
      shortcutsDialog,
      userFootnoteCount,
    ]
  );

  return (
    <>
      <RemotePresenceList peers={realtimeRoom.remotePeers} />
      <DocxRemoteCursorLayer view={view} host={scrollEl} peers={realtimeRoom.remotePeers} />
      <EditorShell
        adapter={adapter}
        onBack={onCloseProp}
        topBarExtras={<PresenceSlot state={realtimeRoom} />}
        toolbar={
          <Toolbar
            agentReady={agentReady}
            docInfo={docInfo}
            activeStyle={activeStyle}
            textFormatProvider={textFormatProvider}
            textFormatActive={textFormatActive}
            activeAlignment={activeAlignment}
            activeSpacing={activeSpacing}
            activeIndentLeft={activeIndentLeft}
            styleOptions={styleOptions}
            onInsertImage={insertImageFromFile}
            onInsertFootnote={() => void insertFootnote()}
            onInsertTable={(r, c) => void insertTable(r, c)}
            onInsertFromXlsx={() => setXlsxPickerOpen("materialized")}
            onSetParagraphStyle={(s) => void setParagraphStyle(s)}
            onSetAlignment={(a) => void setAlignment(a)}
            onAdjustIndent={(d) => void adjustIndent(d)}
            onSetParagraphSpacing={(patch) => void setParagraphSpacing(patch)}
            onToggleList={(k) => void toggleList(k)}
            onAddComment={openCommentComposer}
            onUnsupported={surfaceUnsupported}
            editMode={editMode}
            onSetEditMode={setEditMode}
            formattingMarksOn={formattingMarksOn}
            onToggleFormattingMarks={handleToggleFormattingMarks}
            trackedChangesCount={trackedChangesCount}
            onAcceptAllChanges={() => void acceptAllChanges()}
            onRejectAllChanges={() => void rejectAllChanges()}
            onOpenProtectDocument={() => setProtectDocumentOpen(true)}
            documentProtectionActive={protectionState.enabled}
            onInsertSectionBreak={(type) => void insertSectionBreak(type)}
            hfFocus={hfZoneFocus}
            onCloseHeaderFooter={closeHeaderFooter}
            onToggleSectionDifferentFirst={(checked) => void toggleSectionDifferentFirst(checked)}
            currentSectionTitlePg={currentSectionTitlePg}
            currentListLevel={currentListLevel}
            onSetListLevel={(ilvl) => void setListLevel(ilvl)}
            onInsertHFField={(kind) => void insertHeaderFooterField(kind)}
            onInsertHFImage={() => hfImageInputRef.current?.click()}
            selectedImage={selectedImage}
            onEditImageAlt={handleImageEditAlt}
            onDeleteImage={(id) => void handleImageDelete(id)}
            selectedTableId={selectedTableId}
            onInsertTableRow={(id, where) => void handleTableInsertRow(id, where)}
            onInsertTableColumn={(id, where) => void handleTableInsertColumn(id, where)}
            onDeleteTableRow={(id, r) => void handleDeleteTableRow(id, r)}
            onDeleteTableColumn={(id, c) => void handleDeleteTableColumn(id, c)}
            onDeleteTable={(id) => void handleDeleteTable(id)}
            activeTableCell={activeTableCell}
            onSetActiveTableCell={setActiveTableCell}
            onSetCellShading={(id, r, c, fill) => void handleSetCellShading(id, r, c, fill)}
            onSetCellAlignment={(id, r, c, v) => void handleSetCellAlignment(id, r, c, v)}
            onSetRowHeight={(id, r, h, rule) => void handleSetRowHeight(id, r, h, rule)}
            onSetColumnWidth={(id, c, w) => void handleSetColumnWidth(id, c, w)}
            onMergeCellsHorizontal={(id, r, f, t) => void handleMergeCellsHorizontal(id, r, f, t)}
            onOpenBookmarkDialog={handleOpenBookmarkDialog}
            onInsertOrUpdateToc={(mode) => void handleInsertOrUpdateToc(mode)}
            onInsertCaption={() => void handleInsertCaption()}
            onInsertCrossReference={() => void handleInsertCrossReference()}
            onOpenPageColorPicker={() => void handleOpenPageColorPicker()}
            onOpenPageBordersDialog={() => void handleOpenPageBordersDialog()}
            onOpenWatermarkDialog={() => void handleOpenWatermarkDialog()}
            onOpenThemePicker={() => void handleOpenThemePicker()}
          />
        }
        body={
          <div className="docx-editor flex min-h-0 flex-1 flex-col gap-3 p-3">
            <input
              ref={imageInputRef}
              data-testid="image-file-input"
              type="file"
              accept="image/png,image/jpeg,image/gif,image/bmp,image/webp,image/svg+xml"
              className="hidden"
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) void handleImageFile(f);
                e.target.value = "";
              }}
            />
            <input
              ref={hfImageInputRef}
              data-testid="docx-hf-image-file-input"
              type="file"
              accept="image/png,image/jpeg,image/gif,image/bmp,image/webp,image/svg+xml"
              className="hidden"
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) void handleHeaderFooterImageFile(f);
                e.target.value = "";
              }}
            />
            <div
              ref={setScrollEl}
              className="relative flex-1 overflow-auto rounded-md border border-divider bg-[var(--page-backdrop)]"
            >
              {(() => {
                // Page geometry, derived once per render from the live snapshot.
                // - 1 inch = 1440 twips = 96 CSS px → 1 CSS px = 15 twips.
                // - Width drives the white card; left/right margins drive the
                //   universal horizontal padding selector in globals.css.
                // - When the snapshot is null (initial frame), fall back to
                //   US-Letter so the visual frame has a stable initial size
                //   instead of collapsing to 0.
                const TWIPS_PER_CSS_PX = 15;
                // Phase 1 of docx-fidelity-overhaul: the wrapper uses the
                // WIDEST geometry across all sections so a doc that mixes
                // portrait + landscape (or different page sizes per
                // section) has enough horizontal room for every page. The
                // FIRST-section geometry still drives the default
                // `--pm-page-margin-*` variables on the editor card; each
                // page-block decoration overrides them per chunk via inline
                // `--pm-page-margin-left` / `--pm-page-margin-right` in
                // `page-decorations.ts`.
                // Initial-frame fallback is DIN A4 (11906 × 16838 twips =
                // 210 × 297 mm) with Word's German-locale default 2.5 cm
                // margins (1417 twips). A4 covers ~99% of installs outside
                // the US; using US-Letter as the fallback made the editor
                // briefly render as a wider letter-shaped card before the
                // real snapshot landed and snapped it back to A4. Once the
                // snapshot resolves, `documentMaxPageGeometry` takes over
                // and respects whatever the doc actually declares.
                const wrapperGeometry = snapshot
                  ? documentMaxPageGeometry(snapshot)
                  : {
                      pgSz: { w: 11906, h: 16838 },
                      pgMar: { top: 1417, right: 1417, bottom: 1417, left: 1417, header: 708, footer: 708 },
                    };
                const firstSectionGeometry = snapshot ? documentPageGeometry(snapshot) : wrapperGeometry;
                const pageWidthCssPx = wrapperGeometry.pgSz.w / TWIPS_PER_CSS_PX;
                const pageMarginLeftCssPx = firstSectionGeometry.pgMar.left / TWIPS_PER_CSS_PX;
                const pageMarginRightCssPx = firstSectionGeometry.pgMar.right / TWIPS_PER_CSS_PX;
                const cssVars = {
                  "--pm-page-width": `${pageWidthCssPx}px`,
                  "--pm-page-margin-left": `${pageMarginLeftCssPx}px`,
                  "--pm-page-margin-right": `${pageMarginRightCssPx}px`,
                } as React.CSSProperties;
                // When the document carries tracked changes, reserve a
                // right gutter inside the page wrapper so the Word-style
                // balloons rendered by `<TrackedChangesMargin>` have
                // guaranteed visual real estate alongside the page card
                // (otherwise they would overflow the scroll container's
                // visible area on narrower viewports).
                const reservedGutter = hasRevisions ? BALLOON_RESERVED_PX : 0;
                const wrapperWidthPx = pageWidthCssPx + reservedGutter;
                return (
                  <div
                    className="mx-auto py-6"
                    style={{
                      ...cssVars,
                      width: `${wrapperWidthPx * zoom}px`,
                      transform: `scale(${zoom})`,
                      transformOrigin: "top center",
                    }}
                  >
                    <div style={{ width: `${pageWidthCssPx}px` }}>
                      <PageRuler
                        snapshot={snapshot}
                        onMarginsChange={(next) => void handleRulerMargins(next)}
                        onOpenPageSetup={() => setPageSetupOpen(true)}
                      />
                      <div ref={setHostEl} className="prose-pm min-h-[60vh] outline-none" style={cssVars} />
                    </div>
                  </div>
                );
              })()}
              <TrackedChangesMargin
                snapshot={snapshot}
                editorHost={hostEl}
                scrollContainer={scrollEl}
                onAccept={acceptChange}
                onReject={rejectChange}
              />
              <TableContextToolbar
                view={view}
                host={scrollEl}
                onInsertRow={(id, w) => void handleTableInsertRow(id, w)}
                onInsertColumn={(id, w) => void handleTableInsertColumn(id, w)}
              />
              <ImageResizeOverlay
                view={view}
                host={scrollEl}
                onResize={(id, w, h) => void handleImageResize(id, w, h)}
              />
              <ImageContextToolbar
                view={view}
                host={scrollEl}
                onEditAlt={handleImageEditAlt}
                onDelete={(id) => void handleImageDelete(id)}
              />
              {composer && (
                <CommentComposer
                  selectionText={composer.selectionText}
                  anchor={composer.anchor}
                  onSubmit={(t) => void submitComment(t)}
                  onCancel={() => setComposer(null)}
                />
              )}
              {linkRequest && (
                <HyperlinkPopover
                  anchor={linkRequest.anchor}
                  initialUrl={linkRequest.initialUrl}
                  initialText={linkRequest.selectionText}
                  existingHyperlinkId={null}
                  onApply={(next) => void submitHyperlink(next)}
                  onCancel={() => setLinkRequest(null)}
                />
              )}
            </div>
            <TrackedChangesHover editorHost={hostEl} onAccept={acceptChange} onReject={rejectChange} />
          </div>
        }
        statusBarRight={
          <PageStatusBar
            view={view}
            totalPages={docInfo?.pageCount ?? 1}
            zoom={zoom}
            onZoomChange={setZoom}
          />
        }
        toasts={toasts}
        onDismissToast={dismissToast}
        onFileDrop={(file) => void handleFile(file)}
        dropExtension=".docx"
        onRenameFilename={(next) => setDocName(next)}
      />
      <KeyboardShortcutsDialog
        product="docx"
        open={shortcutsDialog.open}
        onClose={() => shortcutsDialog.setOpen(false)}
      />
      <PageSetupDialog
        open={pageSetupOpen}
        snapshot={snapshot}
        paragraphIndex={view ? currentParagraphIndex(view.state) : 0}
        onClose={() => setPageSetupOpen(false)}
        onSubmit={(next) => void applyPageSetup(next)}
      />
      <ProtectDocumentDialog
        open={protectDocumentOpen}
        current={{
          enabled: protectionState.enabled,
          ...(protectionState.edit ? { edit: protectionState.edit } : {}),
          ...(protectionState.enforce !== undefined ? { enforce: protectionState.enforce } : {}),
          ...(protectionState.formatting !== undefined ? { formatting: protectionState.formatting } : {}),
        }}
        onClose={() => setProtectDocumentOpen(false)}
        onSubmit={(payload) => void applyProtection(payload)}
      />
      <BookmarkDialog
        open={bookmarkDialogOpen}
        bookmarks={bookmarkRows}
        hasSelection={(() => {
          const v = mountRef.current?.view;
          if (!v) return false;
          const { from, to } = v.state.selection;
          return from !== to;
        })()}
        onClose={() => setBookmarkDialogOpen(false)}
        onAdd={(name) => void handleAddBookmark(name)}
        onDelete={(name) => void handleDeleteBookmark(name)}
        onGoTo={(b) => {
          handleGoToBookmark(b);
          setBookmarkDialogOpen(false);
        }}
      />
      <AltTextDialog
        open={altTextRequest !== null}
        imageId={altTextRequest?.imageId ?? null}
        initial={altTextRequest?.initial ?? ""}
        onClose={() => setAltTextRequest(null)}
        onSubmit={(id, alt) => void submitAltText(id, alt)}
      />
      <XlsxRangePickerDialog
        open={xlsxPickerOpen !== null}
        defaultMode={xlsxPickerOpen ?? "materialized"}
        onCancel={() => setXlsxPickerOpen(null)}
        onSubmit={(result) => void handleXlsxPickerSubmit(result)}
      />
      <EmbeddedXlsxModal
        open={editingEmbed !== null}
        bytes={editingEmbed?.bytes ?? null}
        title={editingEmbed?.title}
        onCancel={() => setEditingEmbed(null)}
        onSave={(r) => void handleEmbeddedXlsxSave(r)}
      />
      <GotoDialog
        open={gotoOpen}
        currentPage={
          view ? pageNumberForPos(getPageChunks(view.state), view.state, view.state.selection.from) : 1
        }
        totalPages={docInfo?.pageCount ?? 1}
        onClose={() => setGotoOpen(false)}
        onSubmit={(n) => {
          const v = mountRef.current?.view;
          if (v) gotoPage(v, n, getPageChunks(v.state));
        }}
      />
    </>
  );
}

void Button;

const SECTION_BREAK_LABEL: Record<"nextPage" | "continuous" | "evenPage" | "oddPage", string> = {
  nextPage: "next-page",
  continuous: "continuous",
  evenPage: "even-page",
  oddPage: "odd-page",
};

/** CSS.escape polyfill that is safe to call from older Safari. */
function cssEscape(value: string): string {
  if (typeof CSS !== "undefined" && typeof CSS.escape === "function") return CSS.escape(value);
  return value.replace(/[^a-zA-Z0-9_-]/g, (c) => `\\${c}`);
}

function normaliseUrl(raw: string): string {
  const v = raw.trim();
  if (!v) return v;
  if (/^[a-z][a-z0-9+\-.]*:/i.test(v)) return v;
  if (/^#/.test(v)) return v;
  return `https://${v}`;
}

/**
 * See {@link projectGridToChartData} in `PptxEditor.tsx` — same
 * convention (row 0 → series headers, col 0 → categories, interior
 * → numeric values). Duplicated here instead of being lifted into a
 * shared helper module because the function has zero React-specific
 * deps and the editors don't otherwise share an "embed glue"
 * module beyond the `lib/embed/` directory; lifting one tiny pure
 * helper into its own file would be more ceremony than it's worth.
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

function pickImageFile(files: FileList | null | undefined): File | null {
  if (!files || files.length === 0) return null;
  for (let i = 0; i < files.length; i++) {
    const f = files[i];
    const mime = (f.type || "").toLowerCase();
    if (SUPPORTED_IMAGE_MIME.has(mime)) return f;
    if (mime.startsWith("image/")) return f;
  }
  return null;
}

/**
 * Decode the supplied image file just enough to read its natural
 * width/height in CSS pixels. Used by the H/F "Insert image" toolbar
 * button so we round-trip a sensible default sizing into
 * `docx:insert-header-footer-image` instead of always shipping a
 * 96×96 box. Resolves null on any decode failure (caller falls back
 * to the safe default). The blob URL is revoked synchronously after
 * the load handler fires to avoid leaking it.
 */
function probeImageDimensions(file: File): Promise<{ width: number; height: number }> {
  return new Promise((resolve, reject) => {
    if (typeof window === "undefined") {
      reject(new Error("not in a browser"));
      return;
    }
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      URL.revokeObjectURL(url);
      resolve({ width: img.naturalWidth || 96, height: img.naturalHeight || 96 });
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("image decode failed"));
    };
    img.src = url;
  });
}

/**
 * P3.1 / W4 — derive the effective spacing for the paragraph at the
 * caret. Reads the resolved cascade so an inherited "Heading1 → 1.5
 * line" surfaces in the toolbar even when the paragraph has no direct
 * `<w:spacing>`.
 */
function computeActiveSpacing(
  snapshot: DocxSnapshot | null,
  paragraphIndex: number
): ResolvedSpacingDisplay | null {
  if (!snapshot || paragraphIndex < 0) return null;
  const block = snapshot.root.body[paragraphIndex];
  if (!block || block.kind !== "paragraph") return null;
  const resolved = resolveEffectivePpr(snapshot, paragraphIndex);
  const s = resolved.spacing;
  if (!s) return {};
  const out: ResolvedSpacingDisplay = {};
  if (s.line !== undefined) (out as { line: number }).line = s.line;
  if (s.lineRule !== undefined) {
    (out as { lineRule: ResolvedSpacingDisplay["lineRule"] }).lineRule = s.lineRule;
  }
  if (s.before !== undefined) (out as { before: number }).before = s.before;
  if (s.after !== undefined) (out as { after: number }).after = s.after;
  return out;
}

function computeActiveIndentLeft(snapshot: DocxSnapshot | null, paragraphIndex: number): number | null {
  if (!snapshot || paragraphIndex < 0) return null;
  const block = snapshot.root.body[paragraphIndex];
  if (!block || block.kind !== "paragraph") return null;
  const resolved = resolveEffectivePpr(snapshot, paragraphIndex);
  return resolved.indentation?.left ?? 0;
}

/**
 * B8 — total number of unique unresolved tracked-change wrappers
 * across the body, headers and footers. The Review menu uses this
 * count for its badge and for disabling Accept-all / Reject-all
 * when the document is clean.
 */
/**
 * Pull the currently NodeSelection-selected inline image off the
 * editor view, or return `null` when no image is selected. Mirrors
 * the observation logic in {@link ImageContextToolbar} but is
 * pure-synchronous (no React state) so the contextual ribbon tab can
 * read it during render.
 */
function extractSelectedImage(view: EditorView): SelectedImageInfo | null {
  const sel = view.state.selection;
  const node = (sel as { node?: { type: { name: string }; attrs: Record<string, unknown> } }).node;
  if (!node || node.type.name !== "image") return null;
  const runId = typeof node.attrs.runId === "string" ? node.attrs.runId : null;
  if (!runId) return null;
  const dom = view.nodeDOM(sel.from) as HTMLElement | null;
  const rect = dom?.getBoundingClientRect();
  const width = typeof node.attrs.width === "number" ? node.attrs.width : Math.round(rect?.width ?? 0);
  const height = typeof node.attrs.height === "number" ? node.attrs.height : Math.round(rect?.height ?? 0);
  const alt = typeof node.attrs.alt === "string" ? node.attrs.alt : "";
  return { imageId: runId, widthPx: width, heightPx: height, altText: alt };
}

/**
 * Pull the currently NodeSelection-selected table id off the editor
 * view. Tables surface as PM atom nodes today (cell editing is not
 * yet wired) so a NodeSelection at a `table` node is the canonical
 * "table is selected" signal.
 */
function extractSelectedTableId(view: EditorView): string | null {
  const sel = view.state.selection;
  const node = (sel as { node?: { type: { name: string }; attrs: Record<string, unknown> } }).node;
  if (!node || node.type.name !== "table") return null;
  return typeof node.attrs.tableId === "string" ? node.attrs.tableId : null;
}

interface ProtectionState {
  readonly enabled: boolean;
  readonly edit?: ProtectionEdit;
  readonly enforce?: boolean;
  readonly formatting?: boolean;
}

/**
 * Inspect the verbatim `word/settings.xml` for an active
 * `<w:documentProtection>` element. We only look at the attributes
 * — the password hash / salt / spinCount fields are write-only for
 * the dialog (they're regenerated whenever the user re-applies a
 * password) so we don't surface them in the UI.
 *
 * Treats the element as "active" only when both `w:edit` is set and
 * `w:enforcement` is explicitly `"1"` (mirrors Word's semantics:
 * elements with `enforcement="0"` are persisted but not enforced).
 */
/** Mirror `findOwningSection` from `set-section-different-first.ts` —
 * walk forward from `paragraphIndex` to the next `<w:sectPr>`, falling
 * back to the trailing implicit section. Returns the section's
 * `titlePg` flag (false when no section exists at all). */
function sectionTitlePgAt(snapshot: DocxSnapshot, paragraphIndex: number): boolean {
  const body = snapshot.root.body;
  for (let i = paragraphIndex; i < body.length; i++) {
    const block = body[i];
    if (block && block.kind === "section-break") {
      return Boolean(block.properties.titlePg);
    }
  }
  for (let i = body.length - 1; i >= 0; i--) {
    const block = body[i];
    if (block && block.kind === "section-break") {
      return Boolean(block.properties.titlePg);
    }
  }
  return false;
}

/**
 * 9b — Resolve the active list level (ilvl) for the paragraph the
 * caret currently sits in. Returns `null` when:
 *   - the paragraph index is out of range,
 *   - the body block at that index isn't a paragraph (e.g. a section
 *     break — Word treats those as non-list anyway),
 *   - the paragraph has no `<w:numPr>`.
 *
 * The "in a list at level 0" case must remain distinguishable from
 * "not in a list" so the toolbar's `ListLevelMenu` can disable
 * itself rather than highlight a level the user didn't pick.
 */
function listLevelAt(snapshot: DocxSnapshot, paragraphIndex: number): number | null {
  const block = snapshot.root.body[paragraphIndex];
  if (!block || block.kind !== "paragraph") return null;
  const numbering = block.properties.numbering;
  if (!numbering) return null;
  return numbering.ilvl;
}

function readDocumentProtection(snapshot: DocxSnapshot | null): ProtectionState {
  const empty: ProtectionState = { enabled: false };
  if (!snapshot) return empty;
  const xml = snapshot.root.settingsXml;
  if (!xml) return empty;
  const m = xml.match(/<w:documentProtection\b([^/>]*)\/?>/);
  if (!m) return empty;
  const attrs = m[1] ?? "";
  const edit = readAttr(attrs, "w:edit");
  const enforcement = readAttr(attrs, "w:enforcement");
  const formatting = readAttr(attrs, "w:formatting");
  if (!edit) return empty;
  const isProtectionEdit = (v: string): v is ProtectionEdit =>
    v === "readOnly" ||
    v === "comments" ||
    v === "trackedChanges" ||
    v === "forms" ||
    v === "none";
  const editTyped: ProtectionEdit = isProtectionEdit(edit) ? edit : "readOnly";
  const enforce = enforcement === "1" || enforcement === "true";
  return {
    enabled: enforce && editTyped !== "none",
    edit: editTyped,
    enforce,
    formatting: formatting === "1" || formatting === "true",
  };
}

function readAttr(attrs: string, name: string): string | null {
  const re = new RegExp(`${name.replace(/[.*+?^${}()|[\\]\\\\]/g, "\\$&")}="([^"]*)"`);
  const m = attrs.match(re);
  return m ? (m[1] ?? null) : null;
}

/** Convert a single `FindMatch` (id format `start:end` in flat-text offsets,
 * matching `findAdapter.findAll` above) to PM positions. Returns `null` when
 * the match falls outside any text node — defensive for edge cases like a
 * match that straddled a block boundary that has since been removed. */
function textOffsetsToPmRange(
  doc: import("prosemirror-model").Node,
  match: FindMatch,
): { start: number; end: number } | null {
  const parts = match.id.split(":");
  if (parts.length !== 2) return null;
  const flatStart = Number(parts[0]);
  const flatEnd = Number(parts[1]);
  if (!Number.isFinite(flatStart) || !Number.isFinite(flatEnd)) return null;
  const mapped = mapFlatOffsetsToPmPositions(doc, [{ start: flatStart, end: flatEnd }]);
  return mapped[0] ?? null;
}

/** Walk the PM doc once and resolve a list of flat-text `{start, end}`
 * offsets to PM `{start, end}` positions. The flat text uses the same
 * `textBetween(0, size, "\n", " ")` shape that `findAdapter.findAll` uses,
 * so block boundaries cost one character (matching the regex offsets). */
function mapFlatOffsetsToPmPositions(
  doc: import("prosemirror-model").Node,
  hits: ReadonlyArray<{ start: number; end: number }>,
): Array<{ start: number; end: number }> {
  const out: Array<{ start: number; end: number } | null> = hits.map(() => null);
  let acc = 0;
  doc.descendants((node, pos) => {
    if (!node.isText) return true;
    const len = node.text?.length ?? 0;
    const nodeStart = acc;
    const nodeEnd = acc + len;
    for (let i = 0; i < hits.length; i++) {
      const h = hits[i];
      if (!h) continue;
      const cur = out[i] ?? { start: -1, end: -1 };
      if (cur.start < 0 && h.start >= nodeStart && h.start <= nodeEnd) {
        cur.start = pos + (h.start - nodeStart);
      }
      if (cur.end < 0 && h.end >= nodeStart && h.end <= nodeEnd) {
        cur.end = pos + (h.end - nodeStart);
      }
      out[i] = cur;
    }
    acc = nodeEnd;
    return true;
  });
  return out
    .filter((p): p is { start: number; end: number } => p !== null && p.start >= 0 && p.end >= 0 && p.end >= p.start);
}

/**
 * Walk the body for `Heading1..Heading9` styled paragraphs and collect
 * them as TOC entries. Used by the References > TOC button to
 * synthesize a Word-style table of contents — see the docstring on
 * `handleInsertOrUpdateToc` for the full strategy + caveats.
 *
 * Returned entries are in document order with the same flat text Word
 * would render. Tables are walked too so headings inside table cells
 * still appear in the TOC, matching Word's default `\\u` switch.
 */
interface TocHeading {
  readonly text: string;
  readonly level: number;
  readonly paragraphId: string;
}

function collectTocHeadings(snapshot: DocxSnapshot): ReadonlyArray<TocHeading> {
  const out: TocHeading[] = [];
  const visit = (blocks: readonly BlockNode[]): void => {
    for (const block of blocks) {
      if (block.kind === "paragraph") {
        const styleId = block.properties.styleId;
        if (styleId && /^Heading[1-9]$/.test(styleId)) {
          const level = Number.parseInt(styleId.replace("Heading", ""), 10);
          const text = paragraphFlatText(block).trim();
          if (text.length > 0) {
            out.push({ text, level, paragraphId: block.id });
          }
        }
      } else if (block.kind === "table") {
        for (const row of block.rows) {
          for (const cell of row.cells) visit(cell.body);
        }
      }
    }
  };
  visit(snapshot.root.body);
  return out;
}

function paragraphFlatText(p: { children: ReadonlyArray<InlineNode> }): string {
  let buf = "";
  for (const child of p.children) {
    if (child.kind === "run") {
      for (const c of child.children) {
        if (c.kind === "text") buf += c.text;
      }
    }
  }
  return buf;
}

/**
 * Splice a fresh TOC block above `insertAt`. The block opens with a
 * heading paragraph styled `TOCHeading` (Word's standard) and adds one
 * `TOCN` paragraph per heading. The heading paragraph carries the
 * sentinel text "Inhaltsverzeichnis" so the "update" path can find it
 * later without a bespoke marker attribute.
 *
 * Each command is dispatched independently through the existing
 * insert-paragraph + set-paragraph-style commands, which means undo is
 * one tick per paragraph; bundling these into one mutation is a
 * follow-up optimisation.
 */
async function insertTocBlock(
  agent: DocxAgent,
  insertAt: number,
  headings: ReadonlyArray<TocHeading>
): Promise<void> {
  // Heading + entries in reverse, splicing each one at `insertAt` so
  // positions stay stable as we work upwards through the inserted
  // block. The serializer respects `at.paragraph` as an index in the
  // body, so no PM walk is needed here.
  const lines: Array<{ text: string; styleId: string }> = [];
  lines.push({ text: TOC_TITLE_TEXT, styleId: "TOCHeading" });
  for (const h of headings) {
    lines.push({ text: h.text, styleId: `TOC${Math.max(1, Math.min(9, h.level))}` });
  }
  // Insert from last to first so each new paragraph lands at `insertAt`.
  for (let i = lines.length - 1; i >= 0; i--) {
    const line = lines[i];
    if (!line) continue;
    await agent.applyCommand({
      type: "docx:insert-paragraph",
      payload: { at: { paragraph: insertAt }, text: line.text, styleId: line.styleId },
      source: "human",
    });
  }
}

const TOC_TITLE_TEXT = "Inhaltsverzeichnis";

/**
 * Find an existing TOC block — heuristically defined as a paragraph
 * styled `TOCHeading` whose plain text matches the sentinel — and
 * rebuild every TOC entry beneath it. Stops at the first paragraph
 * whose style isn't a TOC entry. When no block is found we fall back
 * to inserting a new one at the top of the body.
 */
async function rebuildTocBlock(
  agent: DocxAgent,
  snapshot: DocxSnapshot,
  headings: ReadonlyArray<TocHeading>
): Promise<void> {
  let headingIdx = -1;
  for (let i = 0; i < snapshot.root.body.length; i++) {
    const b = snapshot.root.body[i];
    if (
      b &&
      b.kind === "paragraph" &&
      b.properties.styleId === "TOCHeading" &&
      paragraphFlatText(b).trim() === TOC_TITLE_TEXT
    ) {
      headingIdx = i;
      break;
    }
  }
  if (headingIdx < 0) {
    await insertTocBlock(agent, 0, headings);
    return;
  }
  // Find the trailing extent of the existing TOC entries.
  let endIdx = headingIdx + 1;
  while (endIdx < snapshot.root.body.length) {
    const b = snapshot.root.body[endIdx];
    if (!b || b.kind !== "paragraph") break;
    const sty = b.properties.styleId ?? "";
    if (!/^TOC[1-9]$/.test(sty)) break;
    endIdx++;
  }
  // Replace by deleting the block (excluding the heading) bottom-up so
  // indices stay stable, then re-inserting fresh entries. We keep the
  // heading paragraph in place so the user's positioning is preserved.
  for (let i = endIdx - 1; i > headingIdx; i--) {
    await agent.applyCommand({
      type: "docx:delete-range",
      payload: {
        range: {
          start: { paragraph: i, run: 0, offset: 0 },
          end: { paragraph: i + 1, run: 0, offset: 0 },
        },
      },
      source: "human",
    });
  }
  // Insert new entries after the heading.
  for (let i = headings.length - 1; i >= 0; i--) {
    const h = headings[i];
    if (!h) continue;
    await agent.applyCommand({
      type: "docx:insert-paragraph",
      payload: {
        at: { paragraph: headingIdx + 1 },
        text: h.text,
        styleId: `TOC${Math.max(1, Math.min(9, h.level))}`,
      },
      source: "human",
    });
  }
}

function countTrackedChanges(snapshot: DocxSnapshot | null): number {
  if (!snapshot) return 0;
  const seen = new Set<string>();
  const visitInline = (node: InlineNode): void => {
    if (node.kind === "revision") {
      seen.add(node.revisionId);
      for (const child of node.children) visitInline(child);
    }
  };
  const visitBlocks = (blocks: readonly BlockNode[]): void => {
    for (const block of blocks) {
      if (block.kind === "paragraph") {
        for (const child of block.children) visitInline(child);
      } else if (block.kind === "table") {
        for (const row of block.rows) {
          for (const cell of row.cells) visitBlocks(cell.body);
        }
      }
    }
  };
  visitBlocks(snapshot.root.body);
  for (const part of snapshot.root.headersAndFooters) visitBlocks(part.body);
  return seen.size;
}

/**
 * P3.3 / W12-W13 — bottom status bar for the editor pane.
 *
 * Shows `Page X of N` based on the live caret position (resolved via
 * the page-decorations plugin's chunk array) and exposes a zoom
 * slider (50–200 %). Both are read-only against the snapshot — zoom
 * is purely a CSS transform on the editor surface, never written
 * back into the document model.
 */
function PageStatusBar(props: {
  view: EditorView | null;
  totalPages: number;
  zoom: number;
  onZoomChange: (z: number) => void;
}) {
  const { view, totalPages, zoom, onZoomChange } = props;
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState("");

  const currentPage = (() => {
    if (!view) return 1;
    const chunks = getPageChunks(view.state);
    if (chunks.length === 0) return 1;
    return pageNumberForPos(chunks, view.state, view.state.selection.from);
  })();

  const submitGoto = () => {
    if (!view) {
      setEditing(false);
      return;
    }
    const n = Number.parseInt(draft, 10);
    if (Number.isFinite(n) && n >= 1) {
      gotoPage(view, n, getPageChunks(view.state));
    }
    setEditing(false);
  };

  return (
    <div className="mt-2 flex items-center justify-between gap-3 px-1 text-xs text-secondary">
      {editing ? (
        <span className="inline-flex items-center gap-1 tabular-nums">
          <span>Page</span>
          <input
            type="number"
            min={1}
            max={totalPages}
            autoFocus
            value={draft}
            onChange={(e) => setDraft(e.currentTarget.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") submitGoto();
              if (e.key === "Escape") setEditing(false);
            }}
            onBlur={submitGoto}
            className="w-12 rounded border border-divider bg-background px-1 py-0.5 text-xs"
            data-testid="page-goto-input"
          />
          <span>of {totalPages}</span>
        </span>
      ) : (
        <button
          type="button"
          onClick={() => {
            setDraft(String(currentPage));
            setEditing(true);
          }}
          className="rounded px-1 py-0.5 tabular-nums hover:bg-hover"
          title="Go to page"
          data-testid="page-status"
        >
          Page {currentPage} of {totalPages}
        </button>
      )}
      <ZoomControl value={zoom} onChange={onZoomChange} />
    </div>
  );
}
