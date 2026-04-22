import type { ReactNode } from "react";

/**
 * Visual accent for a contextual ribbon tab. Office uses a coloured
 * top-band over the tab strip to telegraph what the contextual tab
 * relates to (e.g. yellow for picture tools, green for table tools).
 * The values here map to CSS classes in `globals.css` so the rendering
 * stays themeable.
 */
export type RibbonContextualAccent = "image" | "table" | "shape" | "picture" | "chart" | "hf";

/**
 * One group of related controls within a tab. Mirrors the bottom-
 * labelled clusters Word/Excel/PowerPoint stack horizontally on a
 * tab (e.g. "Zwischenablage", "Schriftart", "Absatz"). The
 * `render` function returns the buttons themselves; the Ribbon
 * primitive wraps them and renders the `label` underneath.
 *
 * Groups can opt out of rendering entirely via `visible(ctx)` —
 * useful when a group's only actions don't apply to the current
 * selection (e.g. table-only buttons in a "Tabelle" group when no
 * table is selected). Hidden groups also drop their separator so
 * the tab doesn't show floating dividers.
 */
export interface RibbonGroup<C> {
  readonly id: string;
  readonly label: string;
  readonly visible?: (ctx: C) => boolean;
  readonly render: (ctx: C) => ReactNode;
}

/**
 * One ribbon tab. Persistent tabs (Start, Einfügen, Layout, ...)
 * render in the leading area of the tab strip. Contextual tabs
 * (Bildtools, Tabellentools, Kopf- und Fußzeile, ...) render after
 * a separator and only appear when `visible(ctx)` returns true.
 *
 * Auto-activation: when `autoActivateWhen(ctx)` flips true the
 * Ribbon switches to that tab, remembering whatever persistent tab
 * the user had pinned. When the trigger goes false the Ribbon
 * restores the pinned tab. Manual tab clicks always pin (overriding
 * any auto-activation until the next selection-driven flip).
 */
export interface RibbonTab<C> {
  readonly id: string;
  readonly label: string;
  readonly testId?: string;
  readonly contextual?: { readonly accent: RibbonContextualAccent };
  readonly visible?: (ctx: C) => boolean;
  readonly autoActivateWhen?: (ctx: C) => boolean;
  readonly groups: ReadonlyArray<RibbonGroup<C>>;
}

/**
 * A complete ribbon catalogue for one editor (DOCX / XLSX / PPTX).
 * `defaultTabId` sets the initial pinned tab — usually `"start"`.
 */
export interface RibbonCatalogue<C> {
  readonly defaultTabId: string;
  readonly tabs: ReadonlyArray<RibbonTab<C>>;
}
