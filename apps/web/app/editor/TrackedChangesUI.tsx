"use client";

import { useEffect, useRef, useState, type ReactNode } from "react";
import { Check, X, GitPullRequestArrow } from "lucide-react";
import { cn } from "@officeai/ui";
import { collectRevisions } from "@/lib/format-helpers";
import type { DocxSnapshot, RevisionWrapper } from "@officeai/docx";

export interface TrackedChangesUIProps {
  snapshot: DocxSnapshot | null;
  /**
   * The PM editor host (the same div passed to mountDocxEditor). The
   * inline overlay attaches mouseover listeners to it so that hovering
   * over `.pm-revision-ins` / `.pm-revision-del` reveals a small
   * accept/reject pair anchored to the hovered span.
   */
  editorHost: HTMLElement | null;
  onAccept: (revisionId: string) => Promise<void> | void;
  onReject: (revisionId: string) => Promise<void> | void;
}

/**
 * Renders two surfaces:
 *
 *   1. A floating accept/reject pair that follows the cursor over any
 *      `.pm-revision-{ins,del}` span inside the editor host. This is the
 *      "small floating accept/reject pair on hover" called out in the
 *      acceptance criteria.
 *
 *   2. A collapsible ribbon listing every revision in document order
 *      with the same accept/reject controls — the dropdown surface for
 *      keyboard / no-mouse users (and for tests, which prefer stable
 *      selectors over hover positioning).
 *
 * If the underlying handlers throw `NotImplementedError` (W4 not yet
 * landed at runtime), the parent's `onUnsupported` toast fires via the
 * accept/reject callbacks; the change stays visible.
 */
export function TrackedChangesUI(props: TrackedChangesUIProps): ReactNode {
  const revisions = props.snapshot ? collectRevisions(props.snapshot) : [];
  return (
    <>
      <InlineHoverWidget editorHost={props.editorHost} onAccept={props.onAccept} onReject={props.onReject} />
      <ChangeListRibbon revisions={revisions} onAccept={props.onAccept} onReject={props.onReject} />
    </>
  );
}

function InlineHoverWidget(props: {
  editorHost: HTMLElement | null;
  onAccept: (id: string) => Promise<void> | void;
  onReject: (id: string) => Promise<void> | void;
}): ReactNode {
  const [hovered, setHovered] = useState<{
    revisionId: string;
    rect: DOMRect;
    type: "ins" | "del";
  } | null>(null);
  const widgetRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    const host = props.editorHost;
    if (!host) return;

    let hideTimer: ReturnType<typeof setTimeout> | null = null;
    const cancelHide = () => {
      if (hideTimer) {
        clearTimeout(hideTimer);
        hideTimer = null;
      }
    };
    const scheduleHide = () => {
      cancelHide();
      hideTimer = setTimeout(() => setHovered(null), 200);
    };

    const onOver = (e: MouseEvent) => {
      const target = e.target as HTMLElement | null;
      if (!target) return;
      const span = target.closest<HTMLElement>(".pm-revision-ins, .pm-revision-del");
      if (!span) return;
      const revisionId = span.getAttribute("data-revision-id") ?? "";
      if (!revisionId) return;
      cancelHide();
      setHovered({
        revisionId,
        rect: span.getBoundingClientRect(),
        type: span.classList.contains("pm-revision-ins") ? "ins" : "del",
      });
    };

    const onOut = (e: MouseEvent) => {
      const related = e.relatedTarget as Node | null;
      if (related && widgetRef.current?.contains(related)) return;
      scheduleHide();
    };

    host.addEventListener("mouseover", onOver);
    host.addEventListener("mouseout", onOut);
    return () => {
      cancelHide();
      host.removeEventListener("mouseover", onOver);
      host.removeEventListener("mouseout", onOut);
    };
  }, [props.editorHost]);

  if (!hovered) return null;

  const top = hovered.rect.top - 32 + window.scrollY;
  const left = hovered.rect.left + window.scrollX;

  return (
    <div
      ref={widgetRef}
      role="group"
      aria-label="Tracked change actions"
      className="tracked-change-widget pointer-events-auto fixed z-50 flex items-center gap-1 rounded-md border border-divider bg-surface px-1 py-0.5 text-xs shadow-md"
      style={{ top, left, position: "absolute" }}
      onMouseEnter={() => undefined}
    >
      <span
        aria-hidden
        className={cn(
          "px-1 text-[10px] font-medium uppercase",
          hovered.type === "ins" ? "text-[var(--success)]" : "text-[var(--error)]"
        )}
      >
        {hovered.type === "ins" ? "Insertion" : "Deletion"}
      </span>
      <button
        type="button"
        title="Accept change"
        aria-label="Accept change"
        onClick={() => void props.onAccept(hovered.revisionId)}
        className="rounded p-1 text-[var(--success)] hover:bg-[var(--success)]/10"
      >
        <Check size={12} />
      </button>
      <button
        type="button"
        title="Reject change"
        aria-label="Reject change"
        onClick={() => void props.onReject(hovered.revisionId)}
        className="rounded p-1 text-[var(--error)] hover:bg-[var(--error)]/10"
      >
        <X size={12} />
      </button>
    </div>
  );
}

function ChangeListRibbon(props: {
  revisions: ReadonlyArray<RevisionWrapper>;
  onAccept: (id: string) => Promise<void> | void;
  onReject: (id: string) => Promise<void> | void;
}): ReactNode {
  const [open, setOpen] = useState(false);
  const count = props.revisions.length;
  return (
    <div className="tracked-changes-ribbon flex flex-col gap-2" data-testid="tracked-changes-ribbon">
      <button
        type="button"
        onClick={() => setOpen((v) => !v)}
        aria-expanded={open}
        className="flex items-center justify-between rounded-md border border-divider bg-surface px-2 py-1.5 text-xs text-foreground hover:bg-hover"
      >
        <span className="flex items-center gap-1.5">
          <GitPullRequestArrow size={12} />
          Tracked changes
        </span>
        <span className="rounded-full bg-hover px-2 py-0.5 text-[10px] font-medium text-secondary">
          {count}
        </span>
      </button>
      {open && (
        <ul className="flex max-h-48 flex-col gap-1 overflow-y-auto rounded-md border border-divider bg-surface p-1 text-xs">
          {count === 0 && <li className="px-2 py-1 text-secondary">No tracked changes in this document.</li>}
          {props.revisions.map((rev) => (
            <li
              key={rev.revisionId}
              data-testid="tracked-change-row"
              data-revision-id={rev.revisionId}
              data-revision-type={rev.revisionType}
              className="flex items-center justify-between gap-2 rounded px-1.5 py-1 hover:bg-hover"
            >
              <div className="min-w-0">
                <div className="flex items-center gap-1.5 font-medium text-foreground">
                  <span
                    className={cn(
                      "rounded px-1 py-0.5 text-[10px] uppercase",
                      rev.revisionType === "ins"
                        ? "bg-[var(--success)]/10 text-[var(--success)]"
                        : "bg-[var(--error)]/10 text-[var(--error)]"
                    )}
                  >
                    {rev.revisionType === "ins" ? "Insert" : "Delete"}
                  </span>
                  <span className="truncate">{rev.author || "Unknown"}</span>
                </div>
                <div className="truncate text-[10px] text-tertiary">{rev.revisionId}</div>
              </div>
              <div className="flex shrink-0 items-center gap-0.5">
                <button
                  type="button"
                  title="Accept"
                  aria-label={`Accept change ${rev.revisionId}`}
                  onClick={() => void props.onAccept(rev.revisionId)}
                  className="rounded p-1 text-[var(--success)] hover:bg-[var(--success)]/10"
                >
                  <Check size={12} />
                </button>
                <button
                  type="button"
                  title="Reject"
                  aria-label={`Reject change ${rev.revisionId}`}
                  onClick={() => void props.onReject(rev.revisionId)}
                  className="rounded p-1 text-[var(--error)] hover:bg-[var(--error)]/10"
                >
                  <X size={12} />
                </button>
              </div>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
