"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Button, cn } from "@officeai/ui";
import {
  EditorShell,
  EmptyState,
  ZoomControl,
  createToastId,
  type ExportFormat,
  type ExportOptionValues,
  type FindAdapter,
  type FindMatch,
  type FindOptions,
  type OutlineEntry,
  type PaletteCommand,
  type ProductAdapter,
  type SaveState,
  type ToastItem,
} from "@/lib/shell";
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
  type PageZoneCommitDetail,
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
import {
  activeMarkAttr,
  commentParagraphIndex,
  commentThreads,
  currentParagraphAlignment,
  currentParagraphId,
  currentParagraphIndex,
  paragraphStyle,
  paragraphStyleOptions,
  pmSelectionToRange,
} from "@/lib/format-helpers";
import { Toolbar, type AlignmentValue, type ResolvedSpacingDisplay } from "./Toolbar";
import { computeDocxActive, createDocxFormatProvider } from "./docxFormatProvider";
import { PageRuler } from "./PageRuler";
import { PageSetupDialog, type PageSetupValues } from "./PageSetupDialog";
import { TableContextToolbar } from "./TableContextToolbar";
import { ImageContextToolbar, type SelectedImageInfo } from "./ImageContextToolbar";
import { ImageResizeOverlay } from "./ImageResizeOverlay";
import { AltTextDialog } from "./AltTextDialog";
import { GotoDialog } from "./GotoDialog";
import { HyperlinkPopover } from "./HyperlinkPopover";
import { CommentsSidebar } from "./CommentsSidebar";
import { TrackedChangesHover, TrackedChangesMargin } from "./TrackedChangesUI";
import { CommentComposer } from "./CommentComposer";
import { collectRevisions } from "@/lib/format-helpers";
import { createDocxCommentsProvider } from "./docxCommentsProvider";
import { insertImageIntoDocx, SUPPORTED_IMAGE_MIME } from "@/lib/image-insert";
import {
  PresenceSlot,
  roomIdForSource,
  useCommandBroadcast,
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
export function DocxEditor({
  onBootstrapReady,
  initialSource,
  initialBlank,
}: DocxEditorProps = {}): React.ReactNode {
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
    initialSource?.name ?? (initialBlank ? "Untitled.docx" : "welcome.docx")
  );
  const [docInfo, setDocInfo] = useState<{
    paragraphs: number;
    revision: number;
    commentThreads: number;
    pageCount: number;
  } | null>(null);
  const [zoom, setZoom] = useState<number>(1);
  const [pageSetupOpen, setPageSetupOpen] = useState(false);
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
        // Three bootstrap paths, picked in priority order:
        //   1. `initialSource` — fetch a pre-existing .docx (sample
        //      files listing on the home page).
        //   2. `initialBlank` — build a truly empty document (the
        //      "New document" action on the home page).
        //   3. Default — build the synthetic welcome sample so the
        //      editor route is never empty when navigated to
        //      directly.
        const buf = initialSource
          ? await fetch(initialSource.url).then(async (res) => {
              if (!res.ok) {
                throw new Error(`Failed to load ${initialSource.name} (${res.status})`);
              }
              return res.arrayBuffer();
            })
          : initialBlank
            ? await buildBlankDocx()
            : await buildSampleDocx();
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
  }, [mountAgent, hostEl, initialSource, initialBlank, pushToast]);

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
      const wroteInPlace = await saveFileViaService(
        new Uint8Array(buf),
        docName,
        PRODUCT_FILE_TYPES.docx.primaryMime,
        fileHandleRef.current
      );
      setSaveState("saved");
      pushToast("success", wroteInPlace ? `Saved ${docName}` : `Downloaded ${docName}`);
    } catch (err) {
      setSaveState("error");
      pushToast("error", err instanceof Error ? err.message : String(err));
    }
  }, [docName, pushToast]);

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

  const submitComment = useCallback(
    async (text: string) => {
      const agent = agentRef.current;
      if (!agent || !composer) return;
      try {
        await agent.applyCommand({
          type: "docx:add-comment",
          payload: {
            range: composer.range,
            text,
            author: "You",
            initials: "Y",
          },
          source: "human",
        });
        pushToast("info", "Comment added.");
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      } finally {
        setComposer(null);
      }
    },
    [composer, pushToast]
  );

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

  const insertImageFromFile = useCallback(() => {
    imageInputRef.current?.click();
  }, []);

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
      const file = pickImageFile(e.clipboardData?.files);
      if (!file) return;
      e.preventDefault();
      void handleImageFile(file);
    };
    hostEl.addEventListener("dragover", onDragOver);
    hostEl.addEventListener("drop", onDrop);
    hostEl.addEventListener("paste", onPaste);
    return () => {
      hostEl.removeEventListener("dragover", onDragOver);
      hostEl.removeEventListener("drop", onDrop);
      hostEl.removeEventListener("paste", onPaste);
    };
  }, [hostEl, handleImageFile]);

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
      replaceMatch() {
        // PM-based replace will land in B5/B6 polish. For now, a no-op
        // keeps the find UI useful without risking model corruption.
      },
      replaceAll() {
        return 0;
      },
    };
  }, [view]);

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
        <CommentsSidebar provider={commentsProvider} onScrollTo={scrollToComment} />
      </div>
    );
  }, [commentsProvider, scrollToComment]);

  const paletteCommands = useMemo<ReadonlyArray<PaletteCommand>>(() => {
    return [
      {
        id: "docx.add-comment",
        label: "Add comment",
        section: "Collaboration",
        run: () => openCommentComposer(),
      },
      {
        id: "docx.page-setup",
        label: "Page setup…",
        section: "Layout",
        run: () => setPageSetupOpen(true),
      },
      {
        id: "docx.insert-table-3x3",
        label: "Insert table (3 × 3)",
        section: "Insert",
        run: () => void insertTable(3, 3),
      },
      {
        id: "docx.insert-table-2x2",
        label: "Insert table (2 × 2)",
        section: "Insert",
        run: () => void insertTable(2, 2),
      },
      {
        id: "docx.insert-hyperlink",
        label: "Insert hyperlink…",
        section: "Insert",
        run: () => openHyperlinkPopover(),
      },
      {
        id: "docx.toggle-marks",
        label: formattingMarksOn ? "Hide formatting marks" : "Show formatting marks",
        section: "View",
        run: () => handleToggleFormattingMarks(),
      },
      {
        id: "docx.set-mode-edit",
        label: "Switch to Editing mode",
        section: "Mode",
        run: () => setEditMode("edit"),
        enabled: editMode !== "edit",
      },
      {
        id: "docx.set-mode-suggest",
        label: "Switch to Suggesting mode",
        section: "Mode",
        run: () => setEditMode("suggest"),
        enabled: editMode !== "suggest",
      },
      {
        id: "docx.set-mode-view",
        label: "Switch to Viewing mode",
        section: "Mode",
        run: () => setEditMode("view"),
        enabled: editMode !== "view",
      },
    ];
  }, [
    editMode,
    formattingMarksOn,
    handleToggleFormattingMarks,
    insertTable,
    openCommentComposer,
    openHyperlinkPopover,
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
    if (!tabFallback && !initialSource) return null;
    return roomIdForSource({
      product: "docx",
      src: initialSource?.url,
      tabFallback,
    });
  }, [agentReady, initialSource, tabFallback]);
  const realtimeRoom = useRealtimeRoom({
    roomId: realtimeRoomId,
    product: "docx",
  });
  useCommandBroadcast({
    agent: agent as unknown as Parameters<typeof useCommandBroadcast>[0]["agent"],
    room: realtimeRoom.room,
  });

  const adapter = useMemo<ProductAdapter>(
    () => ({
      product: "docx",
      filename: docName,
      saveState,
      comments: { openCount: hasComments ? (docInfo?.commentThreads ?? 0) : 0, resolvedCount: 0 },
      outline,
      selectionSummary: { text: selectionText },
      canOpen: true,
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
      openCommentComposer,
      outline,
      paletteCommands,
      renderCommentsPanel,
      saveState,
      selectionText,
      shortcutsDialog,
    ]
  );

  return (
    <>
      <EditorShell
        adapter={adapter}
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
            onInsertTable={(r, c) => void insertTable(r, c)}
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
            onInsertSectionBreak={(type) => void insertSectionBreak(type)}
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
      <AltTextDialog
        open={altTextRequest !== null}
        imageId={altTextRequest?.imageId ?? null}
        initial={altTextRequest?.initial ?? ""}
        onClose={() => setAltTextRequest(null)}
        onSubmit={(id, alt) => void submitAltText(id, alt)}
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
