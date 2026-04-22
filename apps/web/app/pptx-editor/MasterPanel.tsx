"use client";

import * as React from "react";
import { LayoutTemplate, Palette } from "lucide-react";
import type { LayoutKindPayload, PptxSnapshot, SlideLayout, SlideMaster, Theme } from "@officeai/pptx";

/**
 * Read-only "Slide Master" panel. PowerPoint's Slide Master view is
 * editable; ours is intentionally inspect-only for Phase-1 because
 * the underlying typed model only lifts identity (`partPath`, name,
 * layouts list) without breaking apart the layout's shape tree —
 * editing master shapes round-trips through the verbatim `raw` blob,
 * so any change would have to round-trip the full XML which we don't
 * yet do. The visible panel still gives the user the inventory they
 * need to:
 *   • see which masters / layouts / themes the deck declares
 *   • change which layout new slides default to (the picker is wired
 *     to `pptx:add-slide` via the editor's New-slide button)
 *
 * Editing the master / theme palette lands behind a follow-up
 * "F1 master commands" effort — surfaced as a "Coming soon" badge so
 * users aren't surprised by the missing affordances.
 */
export interface MasterPanelProps {
  readonly snapshot: PptxSnapshot;
  readonly defaultLayoutKind: LayoutKindPayload | null;
  readonly onChangeDefaultLayout: (kind: LayoutKindPayload | null) => void;
}

export function MasterPanel({
  snapshot,
  defaultLayoutKind,
  onChangeDefaultLayout,
}: MasterPanelProps): React.ReactNode {
  const masters = React.useMemo(() => Array.from(snapshot.root.masters.values()), [snapshot.root.masters]);
  const themes = React.useMemo(() => Array.from(snapshot.root.theme.values()), [snapshot.root.theme]);
  const themeColors = snapshot.root.themeDefault;

  const allLayouts = React.useMemo(() => {
    const out: { master: SlideMaster; layout: SlideLayout }[] = [];
    for (const master of masters) {
      for (const layout of master.layouts) {
        out.push({ master, layout });
      }
    }
    return out;
  }, [masters]);

  // Aggregate distinct kinds so the picker stays compact even on
  // decks with several masters declaring the same kinds. Skip the
  // `"unknown"` sentinel — it's a parser fallback, not a layout
  // PowerPoint would actually surface to the user.
  const distinctKinds = React.useMemo<ReadonlyArray<LayoutKindPayload>>(() => {
    const set = new Set<LayoutKindPayload>();
    for (const { layout } of allLayouts) {
      if (layout.kind === "unknown") continue;
      set.add(layout.kind);
    }
    return Array.from(set);
  }, [allLayouts]);

  return (
    <div className="flex min-h-0 flex-1 flex-col overflow-y-auto">
      <Section title="New slide default">
        <p className="mb-2 text-xs text-secondary">
          The layout PowerPoint-style &ldquo;New slide&rdquo; buttons inherit. Select &quot;auto&quot; to keep
          our best-guess (title + content for typical decks).
        </p>
        <select
          className="w-full rounded-md border border-divider bg-surface px-2 py-1 text-xs text-foreground focus:border-[var(--accent)] focus:outline-none"
          value={defaultLayoutKind ?? ""}
          onChange={(e) => {
            const v = e.target.value;
            onChangeDefaultLayout(v === "" ? null : (v as LayoutKindPayload));
          }}
          data-testid="master-default-layout"
          aria-label="Default layout for new slides"
        >
          <option value="">Auto (best guess)</option>
          {distinctKinds.map((k) => (
            <option key={k} value={k}>
              {layoutKindLabel(k)}
            </option>
          ))}
        </select>
      </Section>

      <Section title="Theme">
        {themes.length === 0 ? (
          <p className="text-xs text-secondary">No theme parts.</p>
        ) : (
          <div className="space-y-2">
            {themes.map((t) => (
              <ThemeRow key={t.partPath} theme={t} />
            ))}
            <div className="rounded-md border border-divider bg-surface p-2">
              <div className="mb-1.5 text-[11px] font-medium uppercase tracking-wide text-secondary">
                Color scheme
              </div>
              <div className="flex flex-wrap gap-1">
                <Swatch label="Background 1" color={themeColors.bg1} />
                <Swatch label="Background 2" color={themeColors.bg2} />
                <Swatch label="Text 1" color={themeColors.tx1} />
                <Swatch label="Text 2" color={themeColors.tx2} />
                <Swatch label="Accent 1" color={themeColors.accent1} />
                <Swatch label="Accent 2" color={themeColors.accent2} />
                <Swatch label="Accent 3" color={themeColors.accent3} />
                <Swatch label="Accent 4" color={themeColors.accent4} />
                <Swatch label="Accent 5" color={themeColors.accent5} />
                <Swatch label="Accent 6" color={themeColors.accent6} />
              </div>
            </div>
          </div>
        )}
      </Section>

      <Section title={`Slide masters (${masters.length})`}>
        {masters.length === 0 ? (
          <p className="text-xs text-secondary">No slide masters.</p>
        ) : (
          <ul className="space-y-3">
            {masters.map((m) => (
              <MasterRow key={m.partPath} master={m} />
            ))}
          </ul>
        )}
      </Section>

      <Section title="Editing">
        <div
          data-testid="master-coming-soon"
          className="flex flex-col items-center gap-1 rounded border border-dashed border-divider px-3 py-4 text-center"
        >
          <span className="rounded bg-hover px-1.5 py-0.5 text-[10px] font-semibold uppercase tracking-wide text-secondary">
            Coming soon
          </span>
          <p className="text-[11px] text-secondary">
            Editing slide masters and theme palettes is supported in the data model but the UI ships in a follow-up. Use the CLI in the meantime.
          </p>
        </div>
      </Section>
    </div>
  );
}

function Section({
  title,
  children,
}: {
  readonly title: string;
  readonly children: React.ReactNode;
}): React.ReactNode {
  return (
    <section className="border-b border-divider px-3 py-3 last:border-b-0">
      <h3 className="mb-2 text-xs font-semibold uppercase tracking-wide text-secondary">{title}</h3>
      {children}
    </section>
  );
}

function MasterRow({ master }: { readonly master: SlideMaster }): React.ReactNode {
  const partLeaf = leafName(master.partPath);
  return (
    <li className="rounded-md border border-divider bg-surface p-2">
      <div className="mb-1 flex items-center gap-1.5 text-xs font-medium text-foreground">
        <LayoutTemplate size={12} className="text-secondary" />
        <span className="truncate" title={master.partPath}>
          {partLeaf}
        </span>
      </div>
      <div className="text-[11px] text-secondary">
        {master.layouts.length} layout{master.layouts.length === 1 ? "" : "s"}
      </div>
      {master.layouts.length > 0 ? (
        <ul className="mt-1.5 space-y-0.5 pl-3">
          {master.layouts.map((l) => (
            <li
              key={l.partPath}
              className="flex items-center justify-between gap-2 text-[11px] text-foreground"
              data-testid={`master-layout-${leafName(l.partPath)}`}
            >
              <span className="truncate" title={l.partPath}>
                {l.name ?? leafName(l.partPath)}
              </span>
              <span className="shrink-0 rounded-sm border border-divider bg-background px-1 py-0.5 text-[10px] uppercase tracking-wide text-secondary">
                {l.kind}
              </span>
            </li>
          ))}
        </ul>
      ) : null}
    </li>
  );
}

function ThemeRow({ theme }: { readonly theme: Theme }): React.ReactNode {
  return (
    <div className="flex items-center gap-2 rounded-md border border-divider bg-surface p-2 text-xs">
      <Palette size={12} className="text-secondary" />
      <span className="truncate font-medium text-foreground" title={theme.partPath}>
        {theme.name ?? leafName(theme.partPath)}
      </span>
    </div>
  );
}

function Swatch({ label, color }: { readonly label: string; readonly color: string }): React.ReactNode {
  const safe = color.startsWith("#") ? color : `#${color}`;
  return (
    <div
      title={`${label} — ${safe.toUpperCase()}`}
      className="h-6 w-6 rounded-md border border-divider"
      style={{ backgroundColor: safe }}
      aria-label={`${label} swatch ${safe}`}
    />
  );
}

function leafName(path: string): string {
  const i = path.lastIndexOf("/");
  return i === -1 ? path : path.slice(i + 1);
}

function layoutKindLabel(kind: LayoutKindPayload): string {
  switch (kind) {
    case "title":
      return "Title";
    case "titleAndContent":
      return "Title and Content";
    case "sectionHeader":
      return "Section Header";
    case "twoContent":
      return "Two Content";
    case "comparison":
      return "Comparison";
    case "titleOnly":
      return "Title Only";
    case "blank":
      return "Blank";
    case "contentWithCaption":
      return "Content with Caption";
    case "pictureWithCaption":
      return "Picture with Caption";
    case "titleSlide":
      return "Title Slide";
    case "bigNumber":
      return "Big Number";
  }
}
