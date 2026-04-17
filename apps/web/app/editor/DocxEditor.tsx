"use client";

import { useCallback, useEffect, useRef, useState } from "react";
import {
  Bold,
  Italic,
  Underline,
  Download,
  FileUp,
  MessageSquarePlus,
  Sparkles,
  Loader2,
  Check,
  X,
} from "lucide-react";
import { Button, cn } from "@officeai/ui";
import { DocxAgent, mountDocxEditor, docxSchema } from "@officeai/docx";
import type { MountResult, UnsupportedTx } from "@officeai/docx";
import type { Mutation } from "@officeai/core";
import { buildSampleDocx } from "@/lib/sample-docx";

interface ToastMessage {
  id: number;
  kind: "info" | "warn" | "error";
  text: string;
}

export function DocxEditor() {
  const hostRef = useRef<HTMLDivElement | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const agentRef = useRef<DocxAgent | null>(null);
  const mountRef = useRef<MountResult | null>(null);
  const [agentReady, setAgentReady] = useState(false);
  const [pending, setPending] = useState<Mutation[]>([]);
  const [toasts, setToasts] = useState<ToastMessage[]>([]);
  const [agentBusy, setAgentBusy] = useState(false);
  const [agentPrompt, setAgentPrompt] = useState("");
  const [docName, setDocName] = useState("welcome.docx");
  const [docInfo, setDocInfo] = useState<{ blocks: number; revision: number; comments: number } | null>(null);

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
    async (buf: ArrayBuffer) => {
      mountRef.current?.destroy();
      const agent = await DocxAgent.fromBuffer(buf);
      agentRef.current = agent;
      const refreshState = () => {
        setPending([...agent.getPendingMutations()]);
        const snap = agent.getSnapshot();
        setDocInfo({
          blocks: snap.root.body.length,
          revision: snap.revision,
          comments: snap.root.comments.length,
        });
      };
      const off = agent.subscribe(() => refreshState());
      const host = hostRef.current;
      if (!host) return () => off();
      host.innerHTML = "";
      const mount = mountDocxEditor(host, {
        agent,
        source: "human",
        onUnsupported,
        onError,
      });
      mountRef.current = mount;
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
    let cancelled = false;
    let cleanup: (() => void) | undefined;
    void (async () => {
      const buf = await buildSampleDocx();
      if (cancelled) return;
      cleanup = await mountAgent(buf);
    })();
    return () => {
      cancelled = true;
      cleanup?.();
      mountRef.current?.destroy();
      mountRef.current = null;
      agentRef.current = null;
    };
  }, [mountAgent]);

  const handleFile = useCallback(
    async (file: File) => {
      const buf = await file.arrayBuffer();
      setDocName(file.name);
      try {
        await mountAgent(buf);
        pushToast("info", `Opened ${file.name}`);
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    },
    [mountAgent, pushToast]
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

  const applyMark = useCallback(
    (markName: "bold" | "italic" | "underline") => {
      const mount = mountRef.current;
      if (!mount) return;
      const view = mount.view;
      const { from, to, empty } = view.state.selection;
      if (empty) {
        pushToast("info", "Select some text first.");
        return;
      }
      const markType = docxSchema.marks[markName];
      const tx = view.state.tr.addMark(from, to, markType.create());
      view.dispatch(tx);
    },
    [pushToast]
  );

  const insertCommentDemo = useCallback(async () => {
    const agent = agentRef.current;
    const mount = mountRef.current;
    if (!agent || !mount) return;
    const { from, to, empty } = mount.view.state.selection;
    if (empty) {
      pushToast("info", "Select some text to comment on.");
      return;
    }
    const docTextBefore = (pos: number): { paragraph: number; offset: number } => {
      let paragraphIndex = -1;
      let result: { paragraph: number; offset: number } | null = null;
      mount.view.state.doc.descendants((node, nodePos) => {
        if (result) return false;
        if (node.type.name === "paragraph") {
          paragraphIndex++;
          if (pos >= nodePos && pos <= nodePos + node.nodeSize) {
            const start = nodePos + 1;
            result = { paragraph: paragraphIndex, offset: Math.max(0, pos - start) };
            return false;
          }
        }
        return true;
      });
      return result ?? { paragraph: 0, offset: 0 };
    };
    const start = docTextBefore(from);
    const end = docTextBefore(to);
    try {
      await agent.applyCommand({
        type: "docx:add-comment",
        payload: {
          range: {
            start: { paragraph: start.paragraph, run: 0, offset: start.offset },
            end: { paragraph: end.paragraph, run: 0, offset: end.offset },
          },
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

  const runAgentPrompt = useCallback(async () => {
    const agent = agentRef.current;
    if (!agent || !agentPrompt.trim()) return;
    setAgentBusy(true);
    try {
      // Demo agent: prepends "[AI] " to the first paragraph and adds a comment.
      // Real agent integrations build the same Command list — they just
      // do it with an LLM, not a hard-coded recipe.
      await agent.applyCommands([
        {
          type: "docx:insert-text",
          payload: { at: { paragraph: 0, run: 0, offset: 0 }, text: "[AI] " },
          source: "agent",
          agentId: "demo-agent",
        },
        {
          type: "docx:add-comment",
          payload: {
            range: {
              start: { paragraph: 0, run: 0, offset: 0 },
              end: { paragraph: 0, run: 0, offset: 5 },
            },
            text: agentPrompt.trim(),
            author: "demo-agent",
            initials: "AI",
          },
          source: "agent",
          agentId: "demo-agent",
        },
      ]);
      setAgentPrompt("");
      pushToast("info", "Agent proposed 2 changes — review below.");
    } finally {
      setAgentBusy(false);
    }
  }, [agentPrompt, pushToast]);

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

  const pendingCount = pending.length;

  return (
    <div className="grid h-full min-h-0 gap-6 lg:grid-cols-[minmax(0,1fr)_300px]">
      <section className="flex min-h-0 flex-col">
        <div className="flex flex-wrap items-center gap-1 border-b border-divider pb-3">
          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            className="inline-flex items-center gap-1.5 rounded-md border border-divider bg-surface px-2.5 py-1 text-xs text-foreground hover:bg-hover"
          >
            <FileUp size={14} />
            Open .docx
          </button>
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
          <div className="mx-1 h-4 w-px bg-divider" />
          <ToolbarBtn label="Bold" onClick={() => applyMark("bold")}>
            <Bold size={14} />
          </ToolbarBtn>
          <ToolbarBtn label="Italic" onClick={() => applyMark("italic")}>
            <Italic size={14} />
          </ToolbarBtn>
          <ToolbarBtn label="Underline" onClick={() => applyMark("underline")}>
            <Underline size={14} />
          </ToolbarBtn>
          <div className="mx-1 h-4 w-px bg-divider" />
          <ToolbarBtn label="Add comment" onClick={() => void insertCommentDemo()}>
            <MessageSquarePlus size={14} />
          </ToolbarBtn>
          <div className="ml-auto flex items-center gap-3 text-xs text-secondary">
            {docInfo && (
              <span className="hidden whitespace-nowrap md:inline">
                {docInfo.blocks} blocks · rev {docInfo.revision} · {docInfo.comments} comments
              </span>
            )}
            <Button variant="accent" size="sm" onClick={() => void handleExport()}>
              <Download size={14} />
              Export
            </Button>
          </div>
        </div>
        <div className="relative mt-3 flex-1 overflow-auto rounded-md border border-divider bg-background">
          <div
            ref={hostRef}
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
              className={cn(
                "pointer-events-auto rounded-md border px-3 py-1.5 text-xs shadow-sm",
                t.kind === "info" && "border-divider bg-surface text-foreground",
                t.kind === "warn" && "border-[var(--warning)] bg-[var(--warning)]/10 text-[var(--warning)]",
                t.kind === "error" && "border-[var(--error-border)] bg-[var(--error-bg)] text-[var(--error-text)]"
              )}
            >
              {t.text}
            </div>
          ))}
        </div>
      </section>

      <aside className="flex min-h-0 flex-col gap-4 border-divider pt-2 lg:border-l lg:pl-6 lg:pt-0">
        <div>
          <div className="flex items-center gap-2 text-xs font-medium uppercase tracking-wide text-secondary">
            <Sparkles size={12} className="text-[var(--ai-violet)]" />
            Agent
          </div>
          <p className="mt-1 text-xs text-secondary">
            Ask the demo agent to propose changes. Each change goes into the pending queue
            for human review before it lands.
          </p>
          <div className="mt-3 flex flex-col gap-2">
            <textarea
              value={agentPrompt}
              onChange={(e) => setAgentPrompt(e.target.value)}
              placeholder="e.g. Make the intro more concise."
              rows={3}
              className="w-full resize-none rounded-md border border-divider bg-background px-2.5 py-2 text-xs text-foreground placeholder:text-tertiary focus:border-[var(--ai-violet)] focus:outline-none"
            />
            <Button
              variant="accent"
              size="sm"
              onClick={() => void runAgentPrompt()}
              disabled={!agentReady || agentBusy || agentPrompt.trim().length === 0}
              className="bg-[var(--ai-violet)] hover:bg-[var(--ai-violet)]/90"
            >
              {agentBusy ? <Loader2 className="animate-spin" size={14} /> : <Sparkles size={14} />}
              Propose changes
            </Button>
          </div>
        </div>

        <div>
          <div className="flex items-center justify-between">
            <div className="text-xs font-medium uppercase tracking-wide text-secondary">
              Pending mutations
            </div>
            <span className="rounded-full bg-[var(--ai-violet-light)] px-2 py-0.5 text-[10px] font-medium text-[var(--ai-violet)]">
              {pendingCount}
            </span>
          </div>
          {pendingCount === 0 ? (
            <p className="mt-2 text-xs text-secondary">No pending agent edits.</p>
          ) : (
            <>
              <ul className="mt-2 flex flex-col gap-1.5">
                {pending.map((m) => (
                  <li
                    key={m.id}
                    className="flex items-center justify-between gap-2 rounded-md border border-[var(--ai-violet-muted)] bg-[var(--ai-violet-light)] px-2.5 py-1.5 text-xs"
                  >
                    <div className="min-w-0">
                      <div className="font-medium text-foreground">{m.command.type}</div>
                      <div className="truncate text-[10px] text-secondary">{m.id}</div>
                    </div>
                    <div className="flex items-center gap-1">
                      <button
                        type="button"
                        title="Approve"
                        onClick={() => agentRef.current?.approveMutation(m.id)}
                        className="rounded p-1 text-[var(--success)] hover:bg-[var(--success)]/10"
                      >
                        <Check size={12} />
                      </button>
                      <button
                        type="button"
                        title="Reject"
                        onClick={() => agentRef.current?.rejectMutation(m.id)}
                        className="rounded p-1 text-[var(--error)] hover:bg-[var(--error)]/10"
                      >
                        <X size={12} />
                      </button>
                    </div>
                  </li>
                ))}
              </ul>
              <div className="mt-2 flex items-center gap-2">
                <Button variant="ghost" size="sm" onClick={rejectAll}>
                  Reject all
                </Button>
                <Button variant="accent" size="sm" onClick={approveAll}>
                  Approve all
                </Button>
              </div>
            </>
          )}
        </div>
      </aside>
    </div>
  );
}

function ToolbarBtn(props: { label: string; onClick: () => void; children: React.ReactNode }) {
  return (
    <button
      type="button"
      title={props.label}
      onClick={props.onClick}
      className="rounded-md p-1.5 text-secondary hover:bg-hover hover:text-foreground"
    >
      {props.children}
    </button>
  );
}
