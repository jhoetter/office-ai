"use client";

import { useState } from "react";
import { Sparkles, Loader2, Check, X } from "lucide-react";
import type { Mutation } from "@officeai/core";

export type PptxAgentDispatch = (text: string) => Promise<void>;

export interface PptxAgentPanelProps {
  readonly ready: boolean;
  readonly pending: ReadonlyArray<Mutation>;
  readonly onApprove: (id: string) => void;
  readonly onReject: (id: string) => void;
  readonly onApproveAll: () => void;
  readonly onRejectAll: () => void;
  readonly dispatch: PptxAgentDispatch;
}

export function PptxAgentPanel(props: PptxAgentPanelProps) {
  const [text, setText] = useState("");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState<string | null>(null);

  const onSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!text.trim() || !props.ready) return;
    setBusy(true);
    setErr(null);
    try {
      await props.dispatch(text);
      setText("");
    } catch (caught) {
      setErr(caught instanceof Error ? caught.message : String(caught));
    } finally {
      setBusy(false);
    }
  };

  return (
    <section className="flex min-h-0 flex-col gap-3">
      <header className="flex items-center justify-between">
        <h3 className="text-xs font-semibold uppercase tracking-wide text-secondary">
          Agent
        </h3>
        {props.pending.length > 0 ? (
          <div className="flex gap-1">
            <button
              type="button"
              onClick={props.onApproveAll}
              className="rounded border border-divider px-2 py-0.5 text-[11px] hover:bg-hover"
            >
              Approve all
            </button>
            <button
              type="button"
              onClick={props.onRejectAll}
              className="rounded border border-divider px-2 py-0.5 text-[11px] hover:bg-hover"
            >
              Reject all
            </button>
          </div>
        ) : null}
      </header>

      <form onSubmit={onSubmit} className="flex flex-col gap-2">
        <textarea
          value={text}
          onChange={(e) => setText(e.target.value)}
          placeholder='Try: "add a slide titled Q4 Plan" or "make the title red"'
          disabled={!props.ready || busy}
          rows={3}
          className="resize-y rounded-md border border-divider bg-background px-3 py-2 text-sm outline-none focus:border-[var(--ai-violet)]"
        />
        <button
          type="submit"
          disabled={!props.ready || busy || !text.trim()}
          className="inline-flex items-center justify-center gap-2 rounded-md bg-[var(--ai-violet)] px-3 py-1.5 text-xs font-medium text-white hover:opacity-90 disabled:opacity-40"
        >
          {busy ? <Loader2 className="animate-spin" size={12} /> : <Sparkles size={12} />}
          Run
        </button>
        {err ? <p className="text-xs text-[var(--error-text)]">{err}</p> : null}
      </form>

      {props.pending.length > 0 ? (
        <div className="flex flex-col gap-2 overflow-y-auto">
          <h4 className="text-[11px] font-semibold uppercase tracking-wide text-secondary">
            Pending review ({props.pending.length})
          </h4>
          <ul className="flex flex-col gap-1.5">
            {props.pending.map((m) => (
              <li
                key={m.id}
                className="flex items-center justify-between rounded border border-divider bg-surface px-2 py-1 text-xs"
              >
                <span className="truncate" title={m.command.type}>
                  {m.command.type}
                </span>
                <span className="ml-2 flex shrink-0 items-center gap-1">
                  <button
                    type="button"
                    aria-label="Approve"
                    onClick={() => props.onApprove(m.id)}
                    className="rounded p-1 text-[var(--success)] hover:bg-hover"
                  >
                    <Check size={12} />
                  </button>
                  <button
                    type="button"
                    aria-label="Reject"
                    onClick={() => props.onReject(m.id)}
                    className="rounded p-1 text-[var(--error-text)] hover:bg-hover"
                  >
                    <X size={12} />
                  </button>
                </span>
              </li>
            ))}
          </ul>
        </div>
      ) : null}
    </section>
  );
}
