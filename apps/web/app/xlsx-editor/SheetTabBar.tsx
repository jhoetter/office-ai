"use client";

import { useEffect, useRef, useState } from "react";
import type { ReactNode } from "react";
import { Plus } from "lucide-react";
import { cn } from "@officeai/ui";
import { ContextMenu, type ContextMenuItem } from "./ContextMenu";

/**
 * C9 — Sheet management UI.
 *
 * Excel-parity sheet tab strip with the affordances a daily Excel
 * user reaches for instinctively:
 *
 *   - Click a tab → activate.
 *   - Double-click a tab → inline rename (Enter commits, Esc cancels).
 *   - Right-click a tab → context menu: Rename / Duplicate (deferred) /
 *     Delete / Move Left / Move Right / Hide / Unhide.
 *   - Drag a tab horizontally to reorder; the drop target shows a
 *     thin highlight bar between neighbours.
 *   - Trailing "+" button adds a new sheet.
 *
 * The component is purely presentational — every mutation is
 * dispatched through callbacks the parent wires to the command bus
 * (so undo/redo, OOXML round-trip and dirty tracking all flow
 * through the standard pipeline).
 */
export interface SheetTabDescriptor {
  readonly id: string;
  readonly name: string;
  readonly state: "visible" | "hidden" | "veryHidden";
  /** ARGB hex (e.g. "FFCC0000") or undefined for the OS default. */
  readonly tabColor?: string;
}

export interface SheetTabPeerDot {
  readonly clientId: number;
  readonly sheetName: string;
  readonly name: string;
  readonly color: string;
}

export interface SheetTabBarProps {
  readonly sheets: ReadonlyArray<SheetTabDescriptor>;
  readonly activeName: string | null;
  readonly onActivate: (name: string) => void;
  readonly onRename: (currentName: string, nextName: string) => void;
  readonly onDelete: (name: string) => void;
  readonly onMove: (name: string, to: number) => void;
  readonly onAdd: () => void;
  readonly onSetState: (name: string, state: "visible" | "hidden") => void;
  /**
   * Set the tab color for `name`. Pass `null` to clear (Excel: "No
   * color"). Wired by the parent to `xlsx:set-sheet-tab-color`.
   */
  readonly onSetTabColor: (name: string, color: string | null) => void;
  /**
   * Realtime peers projected to a per-sheet dot. Rendered as a tiny
   * stack of colored dots after each tab whose `sheetName` matches a
   * remote peer's currently-published `XlsxSelection.sheetName`. The
   * tooltip lists the peer names so collisions still resolve.
   */
  readonly peers?: ReadonlyArray<SheetTabPeerDot>;
}

export function SheetTabBar(props: SheetTabBarProps): ReactNode {
  const {
    sheets,
    activeName,
    onActivate,
    onRename,
    onDelete,
    onMove,
    onAdd,
    onSetState,
    onSetTabColor,
    peers,
  } = props;
  const peersBySheet = new Map<string, SheetTabPeerDot[]>();
  for (const p of peers ?? []) {
    const existing = peersBySheet.get(p.sheetName);
    if (existing) existing.push(p);
    else peersBySheet.set(p.sheetName, [p]);
  }
  const [renaming, setRenaming] = useState<string | null>(null);
  const [draftName, setDraftName] = useState<string>("");
  const [ctx, setCtx] = useState<{ name: string; x: number; y: number } | null>(null);
  const [dropAt, setDropAt] = useState<number | null>(null);
  const dragName = useRef<string | null>(null);
  const inputRef = useRef<HTMLInputElement | null>(null);

  // Auto-focus + select-all the rename input the moment it mounts so
  // the user can immediately type a new name (Excel parity).
  useEffect(() => {
    if (renaming === null) return;
    const el = inputRef.current;
    if (!el) return;
    el.focus();
    el.select();
  }, [renaming]);

  const startRename = (name: string): void => {
    setRenaming(name);
    setDraftName(name);
  };

  const commitRename = (): void => {
    if (renaming === null) return;
    const next = draftName.trim();
    if (next.length > 0 && next !== renaming) {
      onRename(renaming, next);
    }
    setRenaming(null);
  };

  const cancelRename = (): void => setRenaming(null);

  const visible = sheets.filter((s) => s.state === "visible");
  const hidden = sheets.filter((s) => s.state !== "visible");

  return (
    <div
      data-testid="sheet-tabs"
      className="sheet-tabs flex items-center gap-1 overflow-x-auto rounded-md border border-divider bg-surface px-2 py-1"
    >
      <button
        type="button"
        title="New sheet"
        aria-label="New sheet"
        data-testid="sheet-tab-add"
        onMouseDown={(e) => e.preventDefault()}
        onClick={onAdd}
        className="inline-flex h-6 w-6 shrink-0 items-center justify-center rounded text-secondary hover:bg-hover hover:text-foreground"
      >
        <Plus size={14} />
      </button>

      {visible.length === 0 ? (
        <span className="text-xs text-secondary">No sheets</span>
      ) : (
        visible.map((s, idx) => {
          const isActive = s.name === activeName;
          const isRenaming = renaming === s.name;
          return (
            <div key={s.id} className="relative flex shrink-0 items-center">
              {dropAt === idx ? <div className="mx-0.5 h-5 w-0.5 rounded bg-accent" aria-hidden /> : null}
              {isRenaming ? (
                <input
                  ref={inputRef}
                  data-testid={`sheet-tab-rename-${s.name}`}
                  value={draftName}
                  onChange={(e) => setDraftName(e.target.value)}
                  onBlur={commitRename}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") {
                      e.preventDefault();
                      commitRename();
                    } else if (e.key === "Escape") {
                      e.preventDefault();
                      cancelRename();
                    }
                  }}
                  className="h-6 w-32 rounded border border-accent bg-background px-2 text-xs font-medium text-foreground outline-none"
                />
              ) : (
                <button
                  type="button"
                  data-testid={`sheet-tab-${s.name}`}
                  draggable
                  // Skip default mousedown focus on the tab button so
                  // that clicking a tab while the formula bar is in
                  // Excel-style "point mode" does not blur the bar
                  // (which would otherwise commit/cancel the in-flight
                  // edit). The drag handlers below still fire because
                  // dragstart runs after mousedown.
                  onMouseDown={(e) => e.preventDefault()}
                  onDragStart={(e) => {
                    dragName.current = s.name;
                    e.dataTransfer.effectAllowed = "move";
                    e.dataTransfer.setData("text/plain", s.name);
                  }}
                  onDragOver={(e) => {
                    if (dragName.current === null) return;
                    e.preventDefault();
                    const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
                    const before = e.clientX < rect.left + rect.width / 2;
                    setDropAt(before ? idx : idx + 1);
                  }}
                  onDragEnd={() => {
                    dragName.current = null;
                    setDropAt(null);
                  }}
                  onDrop={(e) => {
                    e.preventDefault();
                    const moving = dragName.current ?? e.dataTransfer.getData("text/plain");
                    const to = dropAt ?? idx;
                    dragName.current = null;
                    setDropAt(null);
                    if (!moving) return;
                    // Convert visible-only `to` into the real
                    // position in `sheets`, accounting for the
                    // moving tab being removed before insertion.
                    const movingIdx = sheets.findIndex((x) => x.name === moving);
                    if (movingIdx === -1 || movingIdx === to) return;
                    const target = sheets[Math.min(visible.length - 1, to)] ?? sheets[sheets.length - 1];
                    const realTo = sheets.findIndex((x) => x.id === target.id);
                    onMove(moving, realTo);
                  }}
                  onClick={() => onActivate(s.name)}
                  onDoubleClick={() => startRename(s.name)}
                  onContextMenu={(e) => {
                    e.preventDefault();
                    setCtx({ name: s.name, x: e.clientX, y: e.clientY });
                  }}
                  className={cn(
                    "inline-flex shrink-0 items-center gap-1.5 rounded px-3 py-1 text-xs font-medium transition-colors",
                    isActive
                      ? "bg-background text-foreground shadow-sm border border-divider"
                      : "text-secondary hover:text-foreground hover:bg-hover"
                  )}
                  style={
                    s.tabColor
                      ? {
                          boxShadow: `inset 0 -3px 0 0 ${argbToCss(s.tabColor)}`,
                        }
                      : undefined
                  }
                  data-testid-tab-color={s.tabColor ?? "none"}
                >
                  <span>{s.name}</span>
                  {(() => {
                    const dots = peersBySheet.get(s.name);
                    if (!dots || dots.length === 0) return null;
                    const tooltip = dots.map((d) => d.name).join(", ");
                    return (
                      <span
                        className="inline-flex items-center -space-x-0.5"
                        title={tooltip}
                        data-testid={`sheet-tab-peers-${s.name}`}
                      >
                        {dots.slice(0, 3).map((d) => (
                          <span
                            key={d.clientId}
                            aria-hidden
                            className="inline-block size-1.5 rounded-full ring-1 ring-surface"
                            style={{ backgroundColor: d.color }}
                          />
                        ))}
                        {dots.length > 3 ? (
                          <span className="ml-1 text-[10px] font-normal text-secondary">
                            +{dots.length - 3}
                          </span>
                        ) : null}
                      </span>
                    );
                  })()}
                </button>
              )}
            </div>
          );
        })
      )}

      {dropAt === visible.length ? <div className="mx-0.5 h-5 w-0.5 rounded bg-accent" aria-hidden /> : null}

      {hidden.length > 0 ? (
        <div
          className="ml-2 inline-flex items-center gap-1 border-l border-divider pl-2 text-[10px] uppercase tracking-wide text-secondary"
          title="Hidden sheets"
        >
          <span>Hidden:</span>
          {hidden.map((s) => (
            <button
              key={s.id}
              type="button"
              onClick={() => onSetState(s.name, "visible")}
              className="rounded border border-dashed border-divider px-2 py-0.5 hover:border-accent hover:text-foreground"
              title={`Show ${s.name}`}
              data-testid={`sheet-tab-hidden-${s.name}`}
            >
              {s.name}
            </button>
          ))}
        </div>
      ) : null}

      <ContextMenu
        open={ctx !== null}
        x={ctx?.x ?? 0}
        y={ctx?.y ?? 0}
        items={
          ctx
            ? buildItems(
                ctx.name,
                sheets,
                onActivate,
                startRename,
                onMove,
                onDelete,
                onSetState,
                onSetTabColor
              )
            : []
        }
        onClose={() => setCtx(null)}
        testId="sheet-tab-context-menu"
      />
    </div>
  );
}

function buildItems(
  name: string,
  sheets: ReadonlyArray<SheetTabDescriptor>,
  onActivate: (name: string) => void,
  startRename: (name: string) => void,
  onMove: (name: string, to: number) => void,
  onDelete: (name: string) => void,
  onSetState: (name: string, state: "visible" | "hidden") => void,
  onSetTabColor: (name: string, color: string | null) => void
): ContextMenuItem[] {
  const idx = sheets.findIndex((s) => s.name === name);
  const isVisible = sheets[idx]?.state === "visible";
  const visibleCount = sheets.filter((s) => s.state === "visible").length;
  const currentColor = sheets[idx]?.tabColor;
  return [
    {
      kind: "action",
      id: "activate",
      label: "Activate",
      onSelect: () => onActivate(name),
    },
    { kind: "divider", id: "d1" },
    {
      kind: "action",
      id: "rename",
      label: "Rename",
      onSelect: () => startRename(name),
    },
    {
      kind: "action",
      id: "move-left",
      label: "Move left",
      disabled: idx <= 0,
      onSelect: () => onMove(name, Math.max(0, idx - 1)),
    },
    {
      kind: "action",
      id: "move-right",
      label: "Move right",
      disabled: idx === -1 || idx >= sheets.length - 1,
      onSelect: () => onMove(name, Math.min(sheets.length - 1, idx + 1)),
    },
    { kind: "divider", id: "d2" },
    {
      kind: "custom" as const,
      id: "tab-color",
      render: (close) => (
        <div className="px-3 py-1.5">
          <div className="mb-1 text-[10px] uppercase tracking-wide text-secondary">Tab color</div>
          <div className="flex flex-wrap items-center gap-1">
            {TAB_COLOR_SWATCHES.map((s) => {
              const isCurrent = currentColor?.toUpperCase() === s.argb.toUpperCase();
              return (
                <button
                  key={s.argb}
                  type="button"
                  title={s.label}
                  data-testid={`sheet-tab-color-${s.argb}`}
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={() => {
                    onSetTabColor(name, s.argb);
                    close();
                  }}
                  className={cn(
                    "size-4 rounded-sm border",
                    isCurrent ? "border-foreground" : "border-divider hover:border-foreground"
                  )}
                  style={{ backgroundColor: argbToCss(s.argb) }}
                />
              );
            })}
            <button
              type="button"
              title="No color"
              data-testid="sheet-tab-color-clear"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => {
                onSetTabColor(name, null);
                close();
              }}
              className={cn(
                "size-4 rounded-sm border text-[8px] leading-none flex items-center justify-center bg-background",
                currentColor ? "border-divider hover:border-foreground" : "border-foreground"
              )}
            >
              ×
            </button>
          </div>
        </div>
      ),
    },
    { kind: "divider", id: "d2b" },
    isVisible
      ? {
          kind: "action" as const,
          id: "hide",
          label: "Hide",
          // Excel forbids hiding the only visible sheet.
          disabled: visibleCount <= 1,
          onSelect: () => onSetState(name, "hidden"),
        }
      : {
          kind: "action" as const,
          id: "unhide",
          label: "Unhide",
          onSelect: () => onSetState(name, "visible"),
        },
    { kind: "divider", id: "d3" },
    {
      kind: "action",
      id: "delete",
      label: "Delete",
      // Excel forbids deleting the only worksheet in a workbook.
      disabled: sheets.length <= 1,
      onSelect: () => onDelete(name),
    },
  ];
}

/**
 * Excel "Standard Colors" palette mirrored from the desktop ribbon's
 * Tab Color submenu. ARGB values are uppercase 8-char hex so they
 * round-trip cleanly through `<tabColor rgb="..."/>`.
 */
const TAB_COLOR_SWATCHES: ReadonlyArray<{ argb: string; label: string }> = [
  { argb: "FFCC0000", label: "Dark Red" },
  { argb: "FFFF0000", label: "Red" },
  { argb: "FFFFC000", label: "Orange" },
  { argb: "FFFFFF00", label: "Yellow" },
  { argb: "FF92D050", label: "Light Green" },
  { argb: "FF00B050", label: "Green" },
  { argb: "FF00B0F0", label: "Light Blue" },
  { argb: "FF0070C0", label: "Blue" },
  { argb: "FF002060", label: "Dark Blue" },
  { argb: "FF7030A0", label: "Purple" },
];

/**
 * Convert an 8-char ARGB hex (`AARRGGBB`, opacity-first per OOXML)
 * into a CSS `#RRGGBB` triple. Falls back to `currentColor` on any
 * parse glitch so the UI still renders something.
 */
function argbToCss(argb: string): string {
  const raw = argb.replace(/^#/, "").trim();
  if (raw.length === 8) return `#${raw.slice(2)}`;
  if (raw.length === 6) return `#${raw}`;
  return "currentColor";
}
