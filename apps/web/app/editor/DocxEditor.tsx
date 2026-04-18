"use client";

import { useCallback, useEffect, useRef, useState } from "react";
import { Loader2, MessageCircle, X } from "lucide-react";
import { Button, cn } from "@officeai/ui";
import {
  DocxAgent,
  chunkIntoPages,
  documentPageGeometry,
  mountDocxEditor,
  docxSchema,
  resolveEffectivePpr,
} from "@officeai/docx";
import type { DocxSnapshot, MountResult, UnsupportedTx, TextFormat } from "@officeai/docx";
import {
  getPageChunks,
  gotoPage,
  pageDecorationsPlugin,
  pageNumberForPos,
  PAGE_ZONE_EDIT_EVENT,
  type PageZoneEditDetail,
} from "@/lib/page-decorations";
import { pageKeymapPlugin } from "@/lib/page-keymap";
import { PageZoneEditor } from "./PageZoneEditor";
import type { EditorView } from "prosemirror-view";
import { NotImplementedError } from "@officeai/core";
import { buildSampleDocx } from "@/lib/sample-docx";
import {
  activeMarkAttr,
  activeMarks as computeActiveMarks,
  activeRunAttr,
  commentParagraphIndex,
  commentThreads,
  currentParagraphAlignment,
  currentParagraphId,
  currentParagraphIndex,
  discoverNumId,
  paragraphStyle,
  paragraphStyleOptions,
  pmSelectionToRange,
} from "@/lib/format-helpers";
import { Toolbar, type AlignmentValue, type ResolvedSpacingDisplay } from "./Toolbar";
import { HeaderFooterPanel } from "./HeaderFooterPanel";
import { PageRuler } from "./PageRuler";
import { CommentsSidebar } from "./CommentsSidebar";
import { TrackedChangesUI } from "./TrackedChangesUI";
import { CommentComposer } from "./CommentComposer";
import { insertImageIntoDocx, SUPPORTED_IMAGE_MIME } from "@/lib/image-insert";

interface ToastMessage {
  id: number;
  kind: "info" | "warn" | "error";
  text: string;
}

export interface DocxEditorProps {}

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
export function DocxEditor(_props: DocxEditorProps = {}): React.ReactNode {
  // The editor host DOM node is exposed via a callback ref so that
  // descendants (e.g. TrackedChangesUI's hover delegation) can read it
  // from React state during render — accessing `hostRef.current`
  // directly during render trips `react-hooks/refs`.
  const [hostEl, setHostEl] = useState<HTMLDivElement | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const imageInputRef = useRef<HTMLInputElement | null>(null);
  // `agentRef` / `mountRef` are kept in addition to the React state
  // mirrors below so that long-lived callbacks (file open, accept
  // change, …) capture a stable reference without re-binding on
  // every state change.
  const agentRef = useRef<DocxAgent | null>(null);
  const mountRef = useRef<MountResult | null>(null);
  const [agent, setAgent] = useState<DocxAgent | null>(null);
  const [view, setView] = useState<EditorView | null>(null);
  const [agentReady, setAgentReady] = useState(false);
  const [toasts, setToasts] = useState<ToastMessage[]>([]);
  const [docName, setDocName] = useState("welcome.docx");
  const [docInfo, setDocInfo] = useState<{
    paragraphs: number;
    revision: number;
    commentThreads: number;
    pageCount: number;
  } | null>(null);
  const [zoom, setZoom] = useState<number>(1);
  // Bumped to force re-derivation of toolbar state (active marks /
  // active style) without keeping a redundant copy of the snapshot.
  const [uiTick, setUiTick] = useState(0);
  const [drawerOpen, setDrawerOpen] = useState(false);

  const toastId = useRef(0);

  const pushToast = useCallback((kind: ToastMessage["kind"], text: string) => {
    const id = ++toastId.current;
    setToasts((prev) => [...prev, { id, kind, text }]);
    setTimeout(() => setToasts((prev) => prev.filter((t) => t.id !== id)), 4000);
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
      const off = agentInstance.subscribe(() => refreshState());
      host.innerHTML = "";
      const mount = mountDocxEditor(host, {
        agent: agentInstance,
        source: "human",
        onUnsupported,
        onError,
        extraPlugins: [pageDecorationsPlugin(agentInstance), pageKeymapPlugin(agentInstance)],
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
      const buf = await buildSampleDocx();
      if (cancelled) return;
      cleanup = await mountAgent(buf, hostEl);
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
  }, [mountAgent, hostEl]);

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

  const handleFile = useCallback(
    async (file: File) => {
      const buf = await file.arrayBuffer();
      setDocName(file.name);
      const host = hostEl;
      if (!host) {
        pushToast("error", "Editor not yet mounted.");
        return;
      }
      try {
        await mountAgent(buf, host);
        pushToast("info", `Opened ${file.name}`);
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [mountAgent, pushToast, hostEl]
  );

  const handleExport = useCallback(async () => {
    const agent = agentRef.current;
    if (!agent) return;
    try {
      const buf = await agent.exportFile();
      const blob = new Blob([buf], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = docName;
      a.click();
      URL.revokeObjectURL(url);
      pushToast("info", `Exported ${docName}`);
    } catch (err) {
      pushToast("error", err instanceof Error ? err.message : String(err));
    }
  }, [docName, pushToast]);

  const toggleMark = useCallback(
    (markName: "bold" | "italic" | "underline" | "strike") => {
      const mount = mountRef.current;
      if (!mount) return;
      const view = mount.view;
      const { from, to, empty } = view.state.selection;
      if (empty) {
        pushToast("info", "Select some text first.");
        return;
      }
      const schemaMarkName = markName === "strike" ? "strikethrough" : markName;
      const markType = docxSchema.marks[schemaMarkName];
      if (!markType) return;
      const has = view.state.doc.rangeHasMark(from, to, markType);
      const tx = has
        ? view.state.tr.removeMark(from, to, markType)
        : view.state.tr.addMark(from, to, markType.create());
      view.dispatch(tx);
    },
    [pushToast]
  );

  const applyFormat = useCallback(
    async (format: TextFormat) => {
      const agent = agentRef.current;
      const mount = mountRef.current;
      if (!agent || !mount) return;
      const view = mount.view;
      if (view.state.selection.empty) {
        pushToast("info", "Select some text first.");
        return;
      }
      const range = pmSelectionToRange(view.state);
      try {
        await agent.applyCommand({
          type: "docx:format-range",
          payload: { range, format },
          source: "human",
        });
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

  // P3.8 — Word-like in-page header/footer authoring. Populated when
  // the page-decorations plugin fires a `pm-page-zone-edit` event
  // (double-click on a header / footer zone). The popover commits
  // back via `docx:set-header-text` / `docx:set-footer-text` and
  // `docx:insert-page-number` so existing tests + OOXML round-trip
  // stay correct.
  const [zoneEditor, setZoneEditor] = useState<{
    detail: PageZoneEditDetail;
    rect: { left: number; top: number; bottom: number; width: number };
  } | null>(null);

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

  useEffect(() => {
    if (!hostEl) return;
    const onZoneEdit = (event: Event) => {
      const ce = event as CustomEvent<PageZoneEditDetail>;
      const detail = ce.detail;
      if (!detail) return;
      const target = event.target as HTMLElement | null;
      const rect = target?.getBoundingClientRect();
      if (!rect) return;
      setZoneEditor({
        detail,
        rect: { left: rect.left, top: rect.top, bottom: rect.bottom, width: rect.width },
      });
    };
    hostEl.addEventListener(PAGE_ZONE_EDIT_EVENT, onZoneEdit as EventListener);
    return () => hostEl.removeEventListener(PAGE_ZONE_EDIT_EVENT, onZoneEdit as EventListener);
  }, [hostEl]);

  // Mirror the active zone-edit state onto the host element via a
  // `data-zone-editing` attribute so the page-sheets CSS can dim the
  // body and label the active zone (Word's Header & Footer mode).
  useEffect(() => {
    if (!hostEl) return;
    if (zoneEditor) hostEl.setAttribute("data-zone-editing", zoneEditor.detail.slot);
    else hostEl.removeAttribute("data-zone-editing");
    return () => hostEl.removeAttribute("data-zone-editing");
  }, [hostEl, zoneEditor]);

  const submitZoneEdit = useCallback(
    async (text: string) => {
      const agent = agentRef.current;
      if (!agent || !zoneEditor) return;
      const { detail } = zoneEditor;
      if (!detail.partPath) {
        pushToast(
          "warn",
          `This document has no ${detail.slot} part. Add one in Word, then re-open the file.`
        );
        setZoneEditor(null);
        return;
      }
      const cmdType = detail.slot === "header" ? "docx:set-header-text" : "docx:set-footer-text";
      try {
        await agent.applyCommand({
          type: cmdType,
          payload: { partId: detail.partPath, paragraphIndex: 0, text },
          source: "human",
        });
        pushToast("info", `${detail.slot === "header" ? "Header" : "Footer"} updated.`);
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      } finally {
        setZoneEditor(null);
      }
    },
    [pushToast, zoneEditor]
  );

  const insertZonePageNumber = useCallback(async () => {
    const agent = agentRef.current;
    if (!agent || !zoneEditor) return;
    const { detail } = zoneEditor;
    if (!detail.partPath) {
      setZoneEditor(null);
      return;
    }
    const snap = agent.getSnapshot();
    const part = snap.root.headersAndFooters.find((p) => p.partPath === detail.partPath);
    const firstPara = part?.body[0];
    const paragraphId = firstPara && firstPara.kind === "paragraph" ? firstPara.id : null;
    if (!paragraphId) {
      pushToast("warn", `Part has no paragraph to insert into.`);
      return;
    }
    try {
      await agent.applyCommand({
        type: "docx:insert-page-number",
        payload: { paragraphId, offset: Number.MAX_SAFE_INTEGER },
        source: "human",
      });
      pushToast("info", "Page number inserted.");
    } catch (err) {
      pushToast("error", err instanceof Error ? err.message : String(err));
    } finally {
      setZoneEditor(null);
    }
  }, [pushToast, zoneEditor]);

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
      const snap = agent.getSnapshot();
      const target = discoverNumId(snap, kind);
      if (!target) {
        pushToast(
          "warn",
          `This document has no ${kind === "bullet" ? "bullet" : "numbered"} list definition. Auto-creation is not yet supported in this build.`
        );
        return;
      }
      try {
        await agent.applyCommand({
          type: "docx:set-paragraph-list",
          payload: { paragraphId, numId: target.numId, ilvl: target.ilvl },
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

  const replyComment = useCallback(
    async (parentId: string, text: string) => {
      const agent = agentRef.current;
      if (!agent) return;
      try {
        await agent.applyCommand({
          type: "docx:reply-comment",
          payload: { parentId, text, author: "You", initials: "Y" },
          source: "human",
        });
        pushToast("info", "Reply added.");
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const resolveComment = useCallback(
    async (commentId: string, resolved: boolean) => {
      const agent = agentRef.current;
      if (!agent) return;
      try {
        await agent.applyCommand({
          type: "docx:resolve-comment",
          payload: { commentId, resolved },
          source: "human",
        });
        pushToast("info", resolved ? "Comment resolved." : "Comment reopened.");
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
  );

  const deleteComment = useCallback(
    async (commentId: string) => {
      const agent = agentRef.current;
      if (!agent) return;
      try {
        await agent.applyCommand({
          type: "docx:delete-comment",
          payload: { commentId },
          source: "human",
        });
        pushToast("info", "Comment deleted.");
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [pushToast]
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

  // Derive toolbar UI state from the current PM view (re-runs on uiTick).
  void uiTick;
  const snapshot = agent?.getSnapshot() ?? null;
  const activeMarks = view ? computeActiveMarks(view.state) : new Set<string>();
  const currentParaIndex = view ? currentParagraphIndex(view.state) : 0;
  const activeStyle = snapshot ? paragraphStyle(snapshot, currentParaIndex) : "Normal";
  // P3.1 / W3 — toolbar dropdowns fall back through the typed style
  // cascade when no direct PM mark carries the attribute. Without this,
  // selecting a Heading1 paragraph showed "Size" / "Font" placeholders
  // instead of the inherited 16pt / Calibri.
  const activeFontSize = view
    ? activeRunAttr<number>(view.state, "font_size", "halfPoints", snapshot, (rPr) => rPr.fontSize)
    : undefined;
  const activeFontFamily = view
    ? activeRunAttr<string>(view.state, "font_family", "family", snapshot, (rPr) => rPr.fontFamily)
    : undefined;
  const activeColor = view
    ? activeRunAttr<string>(view.state, "color", "rgb", snapshot, (rPr) => rPr.color)
    : undefined;
  const activeHighlight = view
    ? activeRunAttr<string>(view.state, "highlight", "name", snapshot, (rPr) => rPr.highlight)
    : undefined;
  const activeAlignment = view ? currentParagraphAlignment(view.state) : null;
  const activeParagraphIndex = view ? currentParagraphIndex(view.state) : -1;
  const activeSpacing = computeActiveSpacing(snapshot, activeParagraphIndex);
  const activeIndentLeft = computeActiveIndentLeft(snapshot, activeParagraphIndex);
  const styleOptions = paragraphStyleOptions(snapshot, activeStyle);
  void commentParagraphIndex;

  return (
    <div className="docx-editor flex h-full min-h-0 flex-col gap-3 lg:grid lg:grid-cols-[minmax(0,1fr)_320px] lg:grid-rows-1 lg:gap-6">
      <section className="flex min-h-0 flex-col">
        <Toolbar
          agentReady={agentReady}
          docInfo={docInfo}
          activeStyle={activeStyle}
          activeMarks={activeMarks}
          activeFontSize={activeFontSize}
          activeFontFamily={activeFontFamily}
          activeColor={activeColor}
          activeHighlight={activeHighlight}
          activeAlignment={activeAlignment}
          activeSpacing={activeSpacing}
          activeIndentLeft={activeIndentLeft}
          styleOptions={styleOptions}
          onOpenFile={() => fileInputRef.current?.click()}
          onInsertImage={insertImageFromFile}
          onExport={() => void handleExport()}
          onSetParagraphStyle={(s) => void setParagraphStyle(s)}
          onApplyFormat={(f) => void applyFormat(f)}
          onToggleMark={toggleMark}
          onSetAlignment={(a) => void setAlignment(a)}
          onAdjustIndent={(d) => void adjustIndent(d)}
          onSetParagraphSpacing={(patch) => void setParagraphSpacing(patch)}
          onToggleList={(k) => void toggleList(k)}
          onAddComment={openCommentComposer}
          onUnsupported={surfaceUnsupported}
        />
        <input
          ref={fileInputRef}
          type="file"
          accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          className="hidden"
          onChange={(e) => {
            const f = e.target.files?.[0];
            if (f) void handleFile(f);
            e.target.value = "";
          }}
        />
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
        <HeaderFooterPanel
          agent={agent}
          snapshot={snapshot}
          onError={(msg) => pushToast("error", msg)}
          onInfo={(msg) => pushToast("info", msg)}
        />
        <div className="relative mt-3 flex-1 overflow-auto rounded-md border border-divider bg-[color-mix(in_srgb,var(--divider)_25%,var(--surface))] dark:bg-[#0e0e0e]">
          {(() => {
            // Page geometry, derived once per render from the live snapshot.
            // - 1 inch = 1440 twips = 96 CSS px → 1 CSS px = 15 twips.
            // - Width drives the white card; left/right margins drive the
            //   universal horizontal padding selector in globals.css.
            // - When the snapshot is null (initial frame), fall back to
            //   US-Letter so the visual frame has a stable initial size
            //   instead of collapsing to 0.
            const TWIPS_PER_CSS_PX = 15;
            const geometry = snapshot
              ? documentPageGeometry(snapshot)
              : {
                  pgSz: { w: 12240, h: 15840 },
                  pgMar: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720 },
                };
            const pageWidthCssPx = geometry.pgSz.w / TWIPS_PER_CSS_PX;
            const pageMarginLeftCssPx = geometry.pgMar.left / TWIPS_PER_CSS_PX;
            const pageMarginRightCssPx = geometry.pgMar.right / TWIPS_PER_CSS_PX;
            const cssVars = {
              "--pm-page-width": `${pageWidthCssPx}px`,
              "--pm-page-margin-left": `${pageMarginLeftCssPx}px`,
              "--pm-page-margin-right": `${pageMarginRightCssPx}px`,
            } as React.CSSProperties;
            return (
              <div
                className="mx-auto py-6"
                style={{
                  ...cssVars,
                  width: `${pageWidthCssPx * zoom}px`,
                  transform: `scale(${zoom})`,
                  transformOrigin: "top center",
                }}
              >
                <PageRuler snapshot={snapshot} />
                <div ref={setHostEl} className="prose-pm min-h-[60vh] outline-none" style={cssVars} />
              </div>
            );
          })()}
          {!agentReady && (
            <div className="absolute inset-0 flex items-center justify-center text-sm text-secondary">
              <Loader2 className="mr-2 animate-spin" size={14} />
              Loading…
            </div>
          )}
          {composer && (
            <CommentComposer
              selectionText={composer.selectionText}
              anchor={composer.anchor}
              onSubmit={(t) => void submitComment(t)}
              onCancel={() => setComposer(null)}
            />
          )}
          {zoneEditor && (
            <PageZoneEditor
              slot={zoneEditor.detail.slot}
              partPath={zoneEditor.detail.partPath}
              pageNumber={zoneEditor.detail.pageNumber}
              initialText={zoneEditor.detail.text}
              anchorRect={zoneEditor.rect}
              onSubmit={(t) => void submitZoneEdit(t)}
              onInsertPageNumber={() => void insertZonePageNumber()}
              onCancel={() => setZoneEditor(null)}
            />
          )}
        </div>
        <PageStatusBar view={view} totalPages={docInfo?.pageCount ?? 1} zoom={zoom} onZoomChange={setZoom} />
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
      </section>

      {/* Drawer toggle — mobile / tablet only */}
      <button
        type="button"
        onClick={() => setDrawerOpen((v) => !v)}
        aria-label="Toggle comments and agent panel"
        aria-expanded={drawerOpen}
        className="fixed bottom-6 right-6 z-40 inline-flex items-center gap-1.5 rounded-full border border-divider bg-surface px-3 py-2 text-xs font-medium text-foreground shadow-md hover:bg-hover lg:hidden"
      >
        <MessageCircle size={14} />
        Comments
      </button>

      <aside
        data-testid="editor-side-panel"
        className={cn(
          "side-panel flex min-h-0 flex-col gap-4 border-divider pt-2 lg:border-l lg:pl-6 lg:pt-0",
          // Below lg, the panel is a slide-up drawer triggered by the
          // floating button. Above lg, it sits in the grid column.
          drawerOpen
            ? "fixed inset-x-0 bottom-0 z-40 max-h-[80vh] overflow-y-auto rounded-t-2xl border-t bg-background p-4 shadow-2xl lg:static lg:max-h-none lg:border-t-0 lg:p-0 lg:shadow-none"
            : "hidden lg:flex"
        )}
      >
        <div className="flex items-center justify-between lg:hidden">
          <span className="text-sm font-medium text-foreground">Side panel</span>
          <button
            type="button"
            aria-label="Close side panel"
            onClick={() => setDrawerOpen(false)}
            className="rounded p-1 text-secondary hover:bg-hover"
          >
            <X size={14} />
          </button>
        </div>

        <CommentsSidebar
          snapshot={snapshot}
          onScrollTo={scrollToComment}
          onReply={replyComment}
          onResolve={resolveComment}
          onDelete={deleteComment}
        />

        <TrackedChangesUI
          snapshot={snapshot}
          editorHost={hostEl}
          onAccept={acceptChange}
          onReject={rejectChange}
        />
      </aside>
    </div>
  );
}

void Button;

/** CSS.escape polyfill that is safe to call from older Safari. */
function cssEscape(value: string): string {
  if (typeof CSS !== "undefined" && typeof CSS.escape === "function") return CSS.escape(value);
  return value.replace(/[^a-zA-Z0-9_-]/g, (c) => `\\${c}`);
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
      <div className="flex items-center gap-2">
        <button
          type="button"
          onClick={() => onZoomChange(Math.max(0.5, Math.round((zoom - 0.1) * 10) / 10))}
          className="rounded border border-divider px-1.5 py-0.5 hover:bg-hover"
          aria-label="Zoom out"
        >
          −
        </button>
        <span className="tabular-nums" data-testid="zoom-percent">
          {Math.round(zoom * 100)}%
        </span>
        <button
          type="button"
          onClick={() => onZoomChange(Math.min(2, Math.round((zoom + 0.1) * 10) / 10))}
          className="rounded border border-divider px-1.5 py-0.5 hover:bg-hover"
          aria-label="Zoom in"
        >
          +
        </button>
        <button
          type="button"
          onClick={() => onZoomChange(1)}
          className="rounded border border-divider px-1.5 py-0.5 hover:bg-hover"
        >
          Reset
        </button>
      </div>
    </div>
  );
}
