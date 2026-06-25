"use client";

import { useEffect, type ReactNode } from "react";
import { AlertTriangle, CheckCircle2, Info, X, XCircle } from "../sonaloop-icons";
import { cn } from "../lib/cn";

export type ToastKind = "info" | "warn" | "error" | "success";

export interface ToastItem {
  readonly id: string;
  readonly kind: ToastKind;
  readonly text: string;
  /** Auto-dismiss after N ms. 0 = sticky. Default 4500. */
  readonly ttlMs?: number;
}

export interface ToasterProps {
  readonly toasts: ReadonlyArray<ToastItem>;
  readonly onDismiss: (id: string) => void;
}

/**
 * Bottom-right toast stack used by all three editors.
 *
 * The product owns the queue (a `useState<ToastItem[]>`); this
 * component just renders + handles auto-dismiss and the close
 * affordance. Same kinds, same animation, same placement everywhere.
 */
export function Toaster({ toasts, onDismiss }: ToasterProps): ReactNode {
  return (
    <div
      className="pointer-events-none fixed bottom-3 right-3 z-50 flex w-full max-w-[360px] flex-col gap-2"
      role="region"
      aria-label="Notifications"
      aria-live="polite"
      aria-relevant="additions"
    >
      {toasts.map((t) => (
        <ToastCard key={t.id} toast={t} onDismiss={onDismiss} />
      ))}
    </div>
  );
}

function ToastCard({
  toast,
  onDismiss,
}: {
  readonly toast: ToastItem;
  readonly onDismiss: (id: string) => void;
}): ReactNode {
  useEffect(() => {
    const ttl = toast.ttlMs ?? 4500;
    if (ttl === 0) return;
    const handle = setTimeout(() => onDismiss(toast.id), ttl);
    return () => clearTimeout(handle);
  }, [toast.id, toast.ttlMs, onDismiss]);

  const tone = TONE[toast.kind];
  return (
    <div
      role={toast.kind === "error" || toast.kind === "warn" ? "alert" : "status"}
      className={cn(
        "pointer-events-auto flex animate-fade-in items-start gap-2 rounded-md border bg-background p-2.5 text-sm shadow-md",
        tone.border
      )}
      data-testid={`toast-${toast.kind}`}
    >
      <span className={cn("mt-0.5 flex-shrink-0", tone.icon)}>{tone.glyph}</span>
      <span className="flex-1 text-foreground">{toast.text}</span>
      <button
        type="button"
        onClick={() => onDismiss(toast.id)}
        className="inline-flex h-5 w-5 flex-shrink-0 items-center justify-center rounded text-secondary hover:bg-hover hover:text-foreground"
        aria-label="Dismiss notification"
      >
        <X size={12} />
      </button>
    </div>
  );
}

const TONE: Record<ToastKind, { border: string; icon: string; glyph: ReactNode }> = {
  info: {
    border: "border-divider",
    icon: "text-secondary",
    glyph: <Info size={14} />,
  },
  warn: {
    border: "border-[color:var(--warning)]/40",
    icon: "text-[color:var(--warning)]",
    glyph: <AlertTriangle size={14} />,
  },
  error: {
    border: "border-[color:var(--error)]/40",
    icon: "text-[color:var(--error)]",
    glyph: <XCircle size={14} />,
  },
  success: {
    border: "border-[color:var(--success)]/40",
    icon: "text-[color:var(--success)]",
    glyph: <CheckCircle2 size={14} />,
  },
};

let toastSeq = 0;
export function createToastId(prefix = "t"): string {
  toastSeq += 1;
  return `${prefix}-${Date.now()}-${toastSeq}`;
}
