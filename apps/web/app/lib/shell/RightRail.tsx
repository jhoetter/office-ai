"use client";

import { useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { cn } from "@officeai/ui";
import { LayoutTemplate, ListTree, MessageSquare, Search, Sparkles, X } from "lucide-react";
import type { OutlineEntry, ProductAdapter } from "./types";

export type RightRailTab = "comments" | "outline" | "animations" | "master";

export interface RightRailProps {
  readonly adapter: ProductAdapter;
  readonly open: boolean;
  readonly tab: RightRailTab;
  readonly onTabChange: (tab: RightRailTab) => void;
  readonly onClose: () => void;
}

/**
 * Collapsible right-hand panel shared by all three editors.
 *
 * Tabs:
 *   - `Comments` (all three products) — delegated to the per-product
 *     `renderCommentsPanel()` adapter.
 *   - `Outline` (DOCX only — adapter exposes `outline?`).
 *
 * The shell auto-opens the rail on Comments when the document has
 * any open comments. Closing it via the X collapses to a hidden
 * state; the top-bar comments icon toggles it back on.
 */
export function RightRail({ adapter, open, tab, onTabChange, onClose }: RightRailProps): ReactNode {
  if (!open) return null;
  const hasOutline = (adapter.outline?.length ?? 0) > 0 || adapter.product === "docx";
  const hasAnimations = adapter.renderAnimationsPanel != null;
  const hasMaster = adapter.renderMasterPanel != null;
  return (
    <aside
      className="flex h-full w-[320px] flex-col border-l border-divider bg-background"
      role="complementary"
      aria-label="Editor side panel"
    >
      <div className="flex items-center justify-between border-b border-divider px-2 py-1.5">
        <div className="flex items-center gap-1" role="tablist" aria-label="Side panel tabs">
          <RailTab
            label="Comments"
            active={tab === "comments"}
            onClick={() => onTabChange("comments")}
            icon={<MessageSquare size={13} />}
            badge={adapter.comments?.openCount}
            testId="rail-tab-comments"
          />
          {hasOutline ? (
            <RailTab
              label="Outline"
              active={tab === "outline"}
              onClick={() => onTabChange("outline")}
              icon={<ListTree size={13} />}
              testId="rail-tab-outline"
            />
          ) : null}
          {hasAnimations ? (
            <RailTab
              label="Animations"
              active={tab === "animations"}
              onClick={() => onTabChange("animations")}
              icon={<Sparkles size={13} />}
              testId="rail-tab-animations"
            />
          ) : null}
          {hasMaster ? (
            <RailTab
              label="Master"
              active={tab === "master"}
              onClick={() => onTabChange("master")}
              icon={<LayoutTemplate size={13} />}
              testId="rail-tab-master"
            />
          ) : null}
        </div>
        <button
          type="button"
          onClick={onClose}
          className="inline-flex h-7 w-7 items-center justify-center rounded-md text-secondary hover:bg-hover hover:text-foreground"
          aria-label="Close side panel"
          title="Close side panel"
          data-testid="rail-close"
        >
          <X size={14} />
        </button>
      </div>
      <div className="flex min-h-0 flex-1 flex-col overflow-hidden">{renderRailBody(tab, adapter)}</div>
    </aside>
  );
}

/**
 * Body dispatcher — keeps `RightRail` declarative and gives us an
 * exhaustive switch so adding a new tab forces a TS compile error
 * rather than a silent fall-through to the comments panel.
 */
function renderRailBody(tab: RightRailTab, adapter: ProductAdapter): ReactNode {
  switch (tab) {
    case "comments":
      return adapter.renderCommentsPanel ? (
        adapter.renderCommentsPanel()
      ) : (
        <div className="p-4 text-sm text-secondary">No comments provider.</div>
      );
    case "outline":
      return <OutlinePanel entries={adapter.outline ?? []} />;
    case "animations":
      return adapter.renderAnimationsPanel ? (
        adapter.renderAnimationsPanel()
      ) : (
        <div className="p-4 text-sm text-secondary">No animations provider.</div>
      );
    case "master":
      return adapter.renderMasterPanel ? (
        adapter.renderMasterPanel()
      ) : (
        <div className="p-4 text-sm text-secondary">No master provider.</div>
      );
    default: {
      const exhaustive: never = tab;
      return exhaustive;
    }
  }
}

function RailTab({
  label,
  active,
  onClick,
  icon,
  badge,
  testId,
}: {
  readonly label: string;
  readonly active: boolean;
  readonly onClick: () => void;
  readonly icon: ReactNode;
  readonly badge?: number;
  readonly testId?: string;
}): ReactNode {
  return (
    <button
      type="button"
      role="tab"
      aria-selected={active}
      onClick={onClick}
      className={cn(
        "inline-flex h-7 items-center gap-1.5 rounded-md px-2 text-xs font-medium transition-colors",
        active ? "bg-hover text-foreground" : "text-secondary hover:bg-hover hover:text-foreground"
      )}
      data-testid={testId}
    >
      {icon}
      <span>{label}</span>
      {badge && badge > 0 ? (
        <span className="ml-1 inline-flex h-4 min-w-4 items-center justify-center rounded-full bg-[var(--accent)] px-1 text-[10px] font-semibold leading-none text-white">
          {badge > 99 ? "99+" : badge}
        </span>
      ) : null}
    </button>
  );
}

/**
 * B9 — Outline / Navigation panel.
 *
 * Word's Navigation pane condensed to its essentials: a filter field
 * at the top, an indented list of headings below, the active heading
 * highlighted, and click-to-scroll. The active row auto-scrolls into
 * view as the caret moves so a long document never strands the user
 * out of context.
 *
 * Filtering is purely client-side (case-insensitive substring) — the
 * outline is small enough that re-evaluating on every keystroke is
 * cheaper than maintaining an index, and we never want filter state
 * to leak into the underlying document model.
 */
function OutlinePanel({ entries }: { readonly entries: ReadonlyArray<OutlineEntry> }): ReactNode {
  const [filter, setFilter] = useState("");
  const activeRef = useRef<HTMLButtonElement | null>(null);

  const visible = useMemo(() => {
    const q = filter.trim().toLowerCase();
    if (!q) return entries;
    return entries.filter((e) => e.text.toLowerCase().includes(q));
  }, [entries, filter]);

  useEffect(() => {
    const el = activeRef.current;
    if (!el) return;
    el.scrollIntoView({ block: "nearest" });
  }, [visible]);

  if (entries.length === 0) {
    return (
      <div className="p-4 text-sm text-secondary">
        No headings yet. Use heading styles in the toolbar to populate the outline.
      </div>
    );
  }
  return (
    <div className="flex min-h-0 flex-1 flex-col">
      <div className="border-b border-divider p-2">
        <label className="relative block">
          <Search size={12} className="absolute left-2 top-1/2 -translate-y-1/2 text-secondary" aria-hidden />
          <input
            type="search"
            value={filter}
            onChange={(e) => setFilter(e.target.value)}
            placeholder="Filter headings"
            className="w-full rounded-md border border-divider bg-surface py-1 pl-7 pr-2 text-xs text-foreground placeholder:text-secondary focus:border-[var(--accent)] focus:outline-none"
            aria-label="Filter outline"
            data-testid="outline-filter"
          />
        </label>
      </div>
      {visible.length === 0 ? (
        <div className="p-4 text-sm text-secondary">No headings match.</div>
      ) : (
        <nav className="flex flex-1 flex-col gap-0.5 overflow-y-auto p-2" aria-label="Document outline">
          {visible.map((e) => (
            <button
              key={e.id}
              ref={e.active ? activeRef : undefined}
              type="button"
              onClick={e.onActivate}
              aria-current={e.active ? "true" : undefined}
              className={cn(
                "truncate rounded-sm border-l-2 px-2 py-1 text-left text-sm text-foreground transition-colors",
                e.active ? "border-[var(--accent)] bg-hover font-medium" : "border-transparent hover:bg-hover"
              )}
              style={{ paddingLeft: `${6 + (e.level - 1) * 12}px` }}
              data-testid={`outline-${e.id}`}
              data-active={e.active ? "true" : "false"}
              data-level={e.level}
              title={e.text || `Heading ${e.level}`}
            >
              {e.text || `Heading ${e.level}`}
            </button>
          ))}
        </nav>
      )}
    </div>
  );
}

/** Hook: derives the auto-open behaviour. Opens the rail on Comments
 * if the document has any unresolved comments and the user hasn't
 * explicitly closed it. */
export function useRightRailController(adapter: ProductAdapter): {
  open: boolean;
  tab: RightRailTab;
  setOpen: (next: boolean) => void;
  setTab: (next: RightRailTab) => void;
} {
  const [open, setOpen] = useState(false);
  const [tab, setTab] = useState<RightRailTab>("comments");
  const [autoOpened, setAutoOpened] = useState(false);
  const [userClosed, setUserClosed] = useState(false);

  useEffect(() => {
    if (autoOpened || userClosed) return;
    if ((adapter.comments?.openCount ?? 0) > 0) {
      setOpen(true);
      setTab("comments");
      setAutoOpened(true);
    }
  }, [adapter.comments?.openCount, autoOpened, userClosed]);

  return {
    open,
    tab,
    setOpen: (next) => {
      setOpen(next);
      if (!next) setUserClosed(true);
      else setUserClosed(false);
    },
    setTab,
  };
}
