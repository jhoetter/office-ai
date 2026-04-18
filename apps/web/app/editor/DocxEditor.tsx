"use client";

import { useCallback, useEffect, useRef, useState } from "react";
import { Loader2, MessageCircle, X } from "lucide-react";
import { Button, cn } from "@officeai/ui";
import { DocxAgent, mountDocxEditor, docxSchema } from "@officeai/docx";
import type { MountResult, UnsupportedTx, TextFormat } from "@officeai/docx";
import type { EditorView } from "prosemirror-view";
import { NotImplementedError, type Mutation } from "@officeai/core";
import { buildSampleDocx } from "@/lib/sample-docx";
import {
  activeMarkAttr,
  activeMarks as computeActiveMarks,
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
import { Toolbar, type AlignmentValue } from "./Toolbar";
import { CommentsSidebar } from "./CommentsSidebar";
import { TrackedChangesUI } from "./TrackedChangesUI";
import { AgentPrompt, type AgentPromptDispatch } from "./AgentPrompt";
import { dispatchToLlm } from "@/lib/llm-client";
import { insertImageIntoDocx, SUPPORTED_IMAGE_MIME } from "@/lib/image-insert";

interface ToastMessage {
  id: number;
  kind: "info" | "warn" | "error";
  text: string;
}

export interface DocxEditorProps {
  /**
   * Override how the agent prompt translates a free-form prompt into bus
   * commands. Defaults to the demo "[AI] " + add-comment recipe so the
   * existing P1.1 e2e flow keeps working. W6 will inject a real LLM
   * caller here.
   */
  agentPromptDispatch?: (agent: DocxAgent) => AgentPromptDispatch;
}

/**
 * The editor surface composed from the four P1.2 / W5 panels:
 *
 *   ┌────────────────────────────────────────────┬──────────────┐
 *   │ Toolbar (style/marks/colors/align/lists)   │              │
 *   ├────────────────────────────────────────────┤   Comments   │
 *   │                                            │              │
 *   │           ProseMirror editor surface       │   Tracked    │
 *   │                                            │   changes    │
 *   │                                            │              │
 *   │                                            │   Agent      │
 *   │                                            │   prompt     │
 *   └────────────────────────────────────────────┴──────────────┘
 *
 * Below 1024px the right column hides behind a "Comments" drawer
 * button anchored bottom-right.
 */
export function DocxEditor(props: DocxEditorProps = {}): React.ReactNode {
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
  const [pending, setPending] = useState<Mutation[]>([]);
  const [toasts, setToasts] = useState<ToastMessage[]>([]);
  const [docName, setDocName] = useState("welcome.docx");
  const [docInfo, setDocInfo] = useState<{
    paragraphs: number;
    revision: number;
    commentThreads: number;
  } | null>(null);
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
        setPending([...agentInstance.getPendingMutations()]);
        const snap = agentInstance.getSnapshot();
        const paragraphs = snap.root.body.reduce((n, b) => (b.kind === "paragraph" ? n + 1 : n), 0);
        setDocInfo({
          paragraphs,
          revision: snap.revision,
          commentThreads: commentThreads(snap).length,
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

  const insertCommentDemo = useCallback(async () => {
    const agent = agentRef.current;
    const mount = mountRef.current;
    if (!agent || !mount) return;
    const view = mount.view;
    if (view.state.selection.empty) {
      pushToast("info", "Select some text to comment on.");
      return;
    }
    const range = pmSelectionToRange(view.state);
    try {
      await agent.applyCommand({
        type: "docx:add-comment",
        payload: {
          range,
          text: "Looks good?",
          author: "You",
          initials: "Y",
        },
        source: "human",
      });
      pushToast("info", "Comment added.");
    } catch (err) {
      pushToast("error", err instanceof Error ? err.message : String(err));
    }
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

  const approveAll = useCallback(() => {
    const agent = agentRef.current;
    if (!agent) return;
    pending.forEach((m) => agent.approveMutation(m.id));
  }, [pending]);

  const rejectAll = useCallback(() => {
    const agent = agentRef.current;
    if (!agent) return;
    pending.forEach((m) => agent.rejectMutation(m.id));
  }, [pending]);

  // Derive toolbar UI state from the current PM view (re-runs on uiTick).
  void uiTick;
  const snapshot = agent?.getSnapshot() ?? null;
  const activeMarks = view ? computeActiveMarks(view.state) : new Set<string>();
  const currentParaIndex = view ? currentParagraphIndex(view.state) : 0;
  const activeStyle = snapshot ? paragraphStyle(snapshot, currentParaIndex) : "Normal";
  const activeFontSize = view ? activeMarkAttr<number>(view.state, "font_size", "halfPoints") : undefined;
  const activeFontFamily = view ? activeMarkAttr<string>(view.state, "font_family", "family") : undefined;
  const activeColor = view ? activeMarkAttr<string>(view.state, "color", "rgb") : undefined;
  const activeHighlight = view ? activeMarkAttr<string>(view.state, "highlight", "name") : undefined;
  const activeAlignment = view ? currentParagraphAlignment(view.state) : null;
  const styleOptions = paragraphStyleOptions(snapshot, activeStyle);
  void commentParagraphIndex;

  // Default dispatch routes through the LLM bridge (`/api/llm`). When the
  // server has no `OPENAI_API_KEY` configured, the helper transparently
  // falls back to the same `[AI] ` + `add-comment` recipe the editor used
  // before W6, so the existing e2e flow keeps working with no env vars.
  const { agentPromptDispatch: agentPromptDispatchProp } = props;
  const promptDispatch: AgentPromptDispatch = (() => {
    if (!agent) return async () => undefined;
    if (agentPromptDispatchProp) return agentPromptDispatchProp(agent);
    return async (text: string) => {
      const result = await dispatchToLlm(text, agent);
      if (result.note) pushToast("warn", result.note);
      if (result.commands.length > 0) await agent.applyCommands(result.commands);
    };
  })();

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
          styleOptions={styleOptions}
          onOpenFile={() => fileInputRef.current?.click()}
          onInsertImage={insertImageFromFile}
          onExport={() => void handleExport()}
          onSetParagraphStyle={(s) => void setParagraphStyle(s)}
          onApplyFormat={(f) => void applyFormat(f)}
          onToggleMark={toggleMark}
          onSetAlignment={(a) => void setAlignment(a)}
          onAdjustIndent={(d) => void adjustIndent(d)}
          onToggleList={(k) => void toggleList(k)}
          onAddComment={() => void insertCommentDemo()}
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
        <div className="relative mt-3 flex-1 overflow-auto rounded-md border border-divider bg-background">
          <div
            ref={setHostEl}
            className="prose-pm mx-auto min-h-[60vh] w-full max-w-[720px] px-8 py-12 outline-none"
          />
          {!agentReady && (
            <div className="absolute inset-0 flex items-center justify-center text-sm text-secondary">
              <Loader2 className="mr-2 animate-spin" size={14} />
              Loading…
            </div>
          )}
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

        <AgentPrompt
          agent={agent}
          agentReady={agentReady}
          pending={pending}
          onApprove={(id) => agent?.approveMutation(id)}
          onReject={(id) => agent?.rejectMutation(id)}
          onApproveAll={approveAll}
          onRejectAll={rejectAll}
          onError={onError}
          dispatch={promptDispatch}
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
