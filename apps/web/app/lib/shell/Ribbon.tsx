"use client";

import * as React from "react";
import { cn } from "@officeai/ui";
import type { RibbonCatalogue, RibbonContextualAccent, RibbonTab } from "./RibbonTypes";
import { autoActivationSignature, resolveActiveTabId, visibleRibbonTabs } from "./ribbon-resolve";

/**
 * Office-style two-row ribbon shared by DOCX, XLSX, and PPTX
 * editors.
 *
 *   Row 1 — tab strip:
 *     persistent tabs · | · contextual tabs                [trailing]
 *   Row 2 — group surface:
 *     [ buttons ]  |  [ buttons ]  |  [ buttons ]  ...
 *     ─Schriftart─    ─Absatz─       ─Formatvorlagen─
 *
 * Behaviour:
 *
 * - The tab strip auto-switches to a contextual tab whose
 *   `autoActivateWhen` is true. The previously pinned (persistent)
 *   tab is remembered and restored when the trigger goes false.
 * - Clicking a tab pins it for the lifetime of the current ctx
 *   "shape" (e.g. as long as the user stays in the same selection
 *   frame). When ctx changes such that auto-activate switches the
 *   tab, the pin is overridden — that's the whole point of
 *   contextual tabs.
 * - Keyboard nav: Left/Right arrows move focus across visible tabs;
 *   Enter/Space activates the focused tab (which also pins it).
 * - The active tab's group surface is always rendered as a single
 *   horizontal row. Each group's `render(ctx)` returns the buttons;
 *   the Ribbon adds the bottom-aligned group label underneath.
 *
 * Width / scrolling: the surface row uses `overflow-x-auto` so a
 * dense tab (many groups) scrolls horizontally on narrow viewports
 * instead of wrapping (and shifting the canvas down). The tab
 * strip itself uses `overflow-x-auto` for the same reason.
 */
export interface RibbonProps<C> {
  readonly catalogue: RibbonCatalogue<C>;
  readonly ctx: C;
  /**
   * Optional right-pinned area (typically actions like "Present",
   * "Edit mode", "Notes"). Mirrors `ToolbarRow.trailing` so
   * existing trailing UIs port over unchanged.
   */
  readonly trailing?: React.ReactNode;
  readonly testId?: string;
  readonly ariaLabel?: string;
}

export function Ribbon<C>(props: RibbonProps<C>): React.ReactElement {
  const { catalogue, ctx, trailing, testId, ariaLabel } = props;

  const visibleTabs = React.useMemo(() => visibleRibbonTabs(catalogue, ctx), [catalogue, ctx]);

  const [pinnedId, setPinnedId] = React.useState<string>(() => {
    const def = catalogue.tabs.find((t) => t.id === catalogue.defaultTabId);
    return def?.id ?? catalogue.tabs[0]?.id ?? "";
  });

  // Office-style "I clicked another tab while a contextual one was
  // firing": when the user explicitly activates a tab we snapshot the
  // current auto-activation signature. Subsequent renders with the
  // same signature respect the pin (so e.g. picking "Animationen"
  // while a shape is selected stays on Animationen). When the
  // selection changes (signature flips) the pin no longer overrides
  // and the next contextual tab takes over automatically.
  const [suppressedSignature, setSuppressedSignature] = React.useState<string | undefined>(undefined);
  const liveSignature = React.useMemo(() => autoActivationSignature(visibleTabs, ctx), [visibleTabs, ctx]);

  // If the auto-activation signature changes after the user pinned a
  // tab, drop the override so the new contextual tab can take over.
  React.useEffect(() => {
    if (suppressedSignature !== undefined && suppressedSignature !== liveSignature) {
      setSuppressedSignature(undefined);
    }
  }, [liveSignature, suppressedSignature]);

  const handleTabActivate = React.useCallback(
    (id: string) => {
      setPinnedId(id);
      setSuppressedSignature(liveSignature);
    },
    [liveSignature]
  );

  // If the previously pinned tab disappears (visibility flipped or
  // catalogue swapped), fall back to the catalogue default. Avoids
  // rendering an empty surface.
  React.useEffect(() => {
    if (!visibleTabs.some((t) => t.id === pinnedId)) {
      setPinnedId(catalogue.defaultTabId);
    }
  }, [visibleTabs, pinnedId, catalogue.defaultTabId]);

  const activeId = resolveActiveTabId(visibleTabs, ctx, pinnedId, suppressedSignature);
  const activeTab = visibleTabs.find((t) => t.id === activeId) ?? visibleTabs[0];

  // Keyboard navigation across tabs. The strip is a single
  // [role=tablist] so screen readers and arrow-key users get the
  // standard pattern.
  const tabRefs = React.useRef<Array<HTMLButtonElement | null>>([]);
  const onTabKey = React.useCallback(
    (e: React.KeyboardEvent<HTMLButtonElement>, index: number) => {
      if (e.key !== "ArrowLeft" && e.key !== "ArrowRight" && e.key !== "Home" && e.key !== "End") return;
      e.preventDefault();
      const last = visibleTabs.length - 1;
      let next = index;
      if (e.key === "ArrowLeft") next = index <= 0 ? last : index - 1;
      else if (e.key === "ArrowRight") next = index >= last ? 0 : index + 1;
      else if (e.key === "Home") next = 0;
      else if (e.key === "End") next = last;
      const target = tabRefs.current[next];
      target?.focus();
    },
    [visibleTabs.length]
  );

  const persistentTabs = visibleTabs.filter((t) => !t.contextual);
  const contextualTabs = visibleTabs.filter((t) => !!t.contextual);

  return (
    <div
      role="region"
      aria-label={ariaLabel}
      data-testid={testId}
      data-active-tab={activeTab?.id}
      className="flex w-full flex-col"
    >
      {/* Row 1 — tab strip */}
      <div
        role="tablist"
        aria-label={ariaLabel ? `${ariaLabel} tabs` : "Ribbon tabs"}
        className="flex h-7 min-h-7 max-h-7 w-full items-stretch border-b border-divider"
      >
        <div className="ribbon-tablist-scroll flex min-w-0 flex-1 items-center gap-0.5 overflow-x-auto px-2 [&>*]:shrink-0">
          {persistentTabs.map((t, idx) => (
            <RibbonTabButton
              key={t.id}
              tab={t}
              active={activeTab?.id === t.id}
              ref={(el) => {
                tabRefs.current[idx] = el;
              }}
              onActivate={() => handleTabActivate(t.id)}
              onKeyDown={(e) => onTabKey(e, idx)}
            />
          ))}
          {contextualTabs.length > 0 ? <span className="mx-1 h-3.5 w-px bg-divider" aria-hidden /> : null}
          {contextualTabs.map((t, i) => {
            const idx = persistentTabs.length + i;
            return (
              <RibbonTabButton
                key={t.id}
                tab={t}
                active={activeTab?.id === t.id}
                ref={(el) => {
                  tabRefs.current[idx] = el;
                }}
                onActivate={() => handleTabActivate(t.id)}
                onKeyDown={(e) => onTabKey(e, idx)}
              />
            );
          })}
        </div>
        {trailing ? (
          <div className="flex shrink-0 items-center gap-1 border-l border-divider/60 px-2">{trailing}</div>
        ) : null}
      </div>

      {/* Row 2 — group surface */}
      <div
        role="tabpanel"
        aria-labelledby={activeTab ? `ribbon-tab-${activeTab.id}` : undefined}
        data-testid="ribbon-active-surface"
        data-active-tab={activeTab?.id}
        className="ribbon-surface-scroll flex h-12 min-h-12 max-h-12 w-full items-stretch overflow-x-auto overflow-y-hidden"
      >
        {activeTab ? <RibbonSurface tab={activeTab} ctx={ctx} /> : null}
      </div>
    </div>
  );
}

interface RibbonTabButtonProps<C> {
  readonly tab: RibbonTab<C>;
  readonly active: boolean;
  readonly onActivate: () => void;
  readonly onKeyDown: (e: React.KeyboardEvent<HTMLButtonElement>) => void;
}

const RibbonTabButton = React.forwardRef(function RibbonTabButtonInner<C>(
  props: RibbonTabButtonProps<C>,
  ref: React.Ref<HTMLButtonElement>
) {
  const { tab, active, onActivate, onKeyDown } = props;
  const accentClass = tab.contextual ? CONTEXTUAL_ACCENT_CLASS[tab.contextual.accent] : "";
  return (
    <button
      ref={ref}
      type="button"
      role="tab"
      id={`ribbon-tab-${tab.id}`}
      aria-selected={active}
      tabIndex={active ? 0 : -1}
      data-testid={tab.testId ?? `ribbon-tab-${tab.id}`}
      data-tab-id={tab.id}
      data-contextual={tab.contextual ? "true" : undefined}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onActivate}
      onKeyDown={onKeyDown}
      className={cn(
        "relative inline-flex h-6 items-center gap-1 rounded-sm px-2 text-[11px] font-medium uppercase tracking-wide text-secondary transition-colors hover:bg-hover hover:text-foreground",
        active && "bg-hover text-foreground",
        active && tab.contextual && "after:absolute after:inset-x-1 after:-top-px after:h-0.5",
        active && accentClass
      )}
    >
      {tab.label}
    </button>
  );
}) as <C>(p: RibbonTabButtonProps<C> & { ref?: React.Ref<HTMLButtonElement> }) => React.ReactElement;

/**
 * CSS classes for the contextual-tab accent strip. The
 * `after:bg-...` selector colours the 2px line shown above an
 * active contextual tab, mirroring Word's coloured tab band.
 * Values reuse our theme tokens so dark mode tracks automatically.
 */
const CONTEXTUAL_ACCENT_CLASS: Readonly<Record<RibbonContextualAccent, string>> = {
  image: "after:bg-amber-400",
  picture: "after:bg-amber-400",
  table: "after:bg-emerald-500",
  shape: "after:bg-purple-500",
  chart: "after:bg-sky-500",
  hf: "after:bg-[var(--accent)]",
};

interface RibbonSurfaceProps<C> {
  readonly tab: RibbonTab<C>;
  readonly ctx: C;
}

function RibbonSurface<C>(props: RibbonSurfaceProps<C>): React.ReactElement {
  const { tab, ctx } = props;
  const visibleGroups = tab.groups.filter((g) => (g.visible ? g.visible(ctx) : true));
  return (
    <div className="flex min-w-0 flex-1 items-stretch px-2">
      {visibleGroups.map((g, i) => (
        <React.Fragment key={g.id}>
          {i > 0 ? <div className="mx-1 my-1.5 w-px bg-divider/60" aria-hidden /> : null}
          <div
            role="group"
            aria-label={g.label}
            data-testid={`ribbon-group-${g.id}`}
            className="flex flex-col items-stretch justify-between py-0.5"
          >
            <div className="flex flex-1 items-center gap-0.5 px-1">{g.render(ctx)}</div>
            <div className="px-1 text-center text-[9px] uppercase tracking-wide text-secondary/70">
              {g.label}
            </div>
          </div>
        </React.Fragment>
      ))}
    </div>
  );
}
