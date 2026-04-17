"use client";

import { useState, type ReactNode } from "react";
import { Sparkles, Loader2, Check, X } from "lucide-react";
import { Button } from "@officeai/ui";
import type { DocxAgent } from "@officeai/docx";
import type { Mutation } from "@officeai/core";

/**
 * `dispatch` is the only LLM-touching seam. The default
 * `defaultAgentDispatch` reproduces the P1.1 behaviour byte-for-byte
 * (prepend "[AI] " to paragraph 0 and add a comment with the prompt
 * text) so the existing `add-comment.spec.ts` continues to pass.
 *
 * W6 will swap in a real LLM caller via this prop. The signature is
 * intentionally simple: a free-form prompt string in, void promise out;
 * the implementation is responsible for translating the prompt into a
 * sequence of bus commands via the agent it captured at construction
 * time.
 */
export type AgentPromptDispatch = (text: string) => Promise<void>;

export interface AgentPromptProps {
  agent: DocxAgent | null;
  agentReady: boolean;
  pending: ReadonlyArray<Mutation>;
  onApprove: (id: string) => void;
  onReject: (id: string) => void;
  onApproveAll: () => void;
  onRejectAll: () => void;
  /** Override the demo dispatch. Defaults to `defaultAgentDispatch(agent)`. */
  dispatch?: AgentPromptDispatch;
  onError?: (err: unknown) => void;
  onSuccess?: () => void;
}

/**
 * Build the legacy "[AI] " + add-comment dispatcher bound to `agent`.
 * Exported so callers can compose with custom LLM dispatchers without
 * losing the demo flow as a fallback.
 */
export function defaultAgentDispatch(agent: DocxAgent): AgentPromptDispatch {
  return async (text: string) => {
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
          text,
          author: "demo-agent",
          initials: "AI",
        },
        source: "agent",
        agentId: "demo-agent",
      },
    ]);
  };
}

export function AgentPrompt(props: AgentPromptProps): ReactNode {
  const [prompt, setPrompt] = useState("");
  const [busy, setBusy] = useState(false);
  const pendingCount = props.pending.length;

  const run = async () => {
    if (!props.agent || !prompt.trim()) return;
    setBusy(true);
    try {
      const dispatch = props.dispatch ?? defaultAgentDispatch(props.agent);
      await dispatch(prompt.trim());
      setPrompt("");
      props.onSuccess?.();
    } catch (err) {
      props.onError?.(err);
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="agent-prompt flex flex-col gap-3">
      <div>
        <div className="flex items-center gap-2 text-xs font-medium uppercase tracking-wide text-secondary">
          <Sparkles size={12} className="text-[var(--ai-violet)]" />
          Agent
        </div>
        <p className="mt-1 text-xs text-secondary">
          Ask the demo agent to propose changes. Each change goes into the pending queue for human review
          before it lands.
        </p>
        <div className="mt-3 flex flex-col gap-2">
          <textarea
            value={prompt}
            onChange={(e) => setPrompt(e.target.value)}
            placeholder="e.g. Make the intro more concise."
            rows={3}
            aria-label="Agent prompt"
            className="w-full resize-none rounded-md border border-divider bg-background px-2.5 py-2 text-xs text-foreground placeholder:text-tertiary focus:border-[var(--ai-violet)] focus:outline-none"
          />
          <Button
            variant="accent"
            size="sm"
            onClick={() => void run()}
            disabled={!props.agentReady || busy || prompt.trim().length === 0}
            className="bg-[var(--ai-violet)] hover:bg-[var(--ai-violet)]/90"
          >
            {busy ? <Loader2 className="animate-spin" size={14} /> : <Sparkles size={14} />}
            Propose changes
          </Button>
        </div>
      </div>

      <div>
        <div className="flex items-center justify-between">
          <div className="text-xs font-medium uppercase tracking-wide text-secondary">Pending mutations</div>
          <span className="rounded-full bg-[var(--ai-violet-light)] px-2 py-0.5 text-[10px] font-medium text-[var(--ai-violet)]">
            {pendingCount}
          </span>
        </div>
        {pendingCount === 0 ? (
          <p className="mt-2 text-xs text-secondary">No pending agent edits.</p>
        ) : (
          <>
            <ul className="mt-2 flex flex-col gap-1.5">
              {props.pending.map((m) => (
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
                      onClick={() => props.onApprove(m.id)}
                      className="rounded p-1 text-[var(--success)] hover:bg-[var(--success)]/10"
                    >
                      <Check size={12} />
                    </button>
                    <button
                      type="button"
                      title="Reject"
                      onClick={() => props.onReject(m.id)}
                      className="rounded p-1 text-[var(--error)] hover:bg-[var(--error)]/10"
                    >
                      <X size={12} />
                    </button>
                  </div>
                </li>
              ))}
            </ul>
            <div className="mt-2 flex items-center gap-2">
              <Button variant="ghost" size="sm" onClick={props.onRejectAll}>
                Reject all
              </Button>
              <Button variant="accent" size="sm" onClick={props.onApproveAll}>
                Approve all
              </Button>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
