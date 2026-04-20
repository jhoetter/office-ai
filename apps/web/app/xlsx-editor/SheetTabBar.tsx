"use client";

import { useEffect, useRef, useState } from "react";
import type { ReactNode } from "react";
import { Plus } from "lucide-react";
import { cn } from "@officeai/ui";
import { ContextMenu, type ContextMenuItem } from "./ContextMenu";
import { useTranslator, type TranslateVars } from "@/lib/i18n";

type TFn = (key: string, vars?: TranslateVars) => string;

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
}

export function SheetTabBar(props: SheetTabBarProps): ReactNode {
  const { sheets, activeName, onActivate, onRename, onDelete, onMove, onAdd, onSetState } = props;
  const { t } = useTranslator();
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
        title={t("xlsx.sheetTab.newSheet")}
        aria-label={t("xlsx.sheetTab.newSheet")}
        data-testid="sheet-tab-add"
        onMouseDown={(e) => e.preventDefault()}
        onClick={onAdd}
        className="inline-flex h-6 w-6 shrink-0 items-center justify-center rounded text-secondary hover:bg-hover hover:text-foreground"
      >
        <Plus size={14} />
      </button>

      {visible.length === 0 ? (
        <span className="text-xs text-secondary">{t("xlsx.sheetTab.noSheets")}</span>
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
                    "shrink-0 rounded px-3 py-1 text-xs font-medium transition-colors",
                    isActive
                      ? "bg-background text-foreground shadow-sm border border-divider"
                      : "text-secondary hover:text-foreground hover:bg-hover"
                  )}
                >
                  {s.name}
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
          title={t("xlsx.sheetTab.hidden")}
        >
          <span>{t("xlsx.sheetTab.hiddenLabel")}</span>
          {hidden.map((s) => (
            <button
              key={s.id}
              type="button"
              onClick={() => onSetState(s.name, "visible")}
              className="rounded border border-dashed border-divider px-2 py-0.5 hover:border-accent hover:text-foreground"
              title={t("xlsx.sheetTab.showSheet", { name: s.name })}
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
        items={ctx ? buildItems(ctx.name, sheets, onActivate, startRename, onMove, onDelete, onSetState, t) : []}
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
  t: TFn
): ContextMenuItem[] {
  const idx = sheets.findIndex((s) => s.name === name);
  const isVisible = sheets[idx]?.state === "visible";
  const visibleCount = sheets.filter((s) => s.state === "visible").length;
  return [
    {
      kind: "action",
      id: "activate",
      label: t("xlsx.sheetTab.activate"),
      onSelect: () => onActivate(name),
    },
    { kind: "divider", id: "d1" },
    {
      kind: "action",
      id: "rename",
      label: t("xlsx.sheetTab.rename"),
      onSelect: () => startRename(name),
    },
    {
      kind: "action",
      id: "move-left",
      label: t("xlsx.sheetTab.moveLeft"),
      disabled: idx <= 0,
      onSelect: () => onMove(name, Math.max(0, idx - 1)),
    },
    {
      kind: "action",
      id: "move-right",
      label: t("xlsx.sheetTab.moveRight"),
      disabled: idx === -1 || idx >= sheets.length - 1,
      onSelect: () => onMove(name, Math.min(sheets.length - 1, idx + 1)),
    },
    { kind: "divider", id: "d2" },
    isVisible
      ? {
          kind: "action" as const,
          id: "hide",
          label: t("xlsx.sheetTab.hide"),
          // Excel forbids hiding the only visible sheet.
          disabled: visibleCount <= 1,
          onSelect: () => onSetState(name, "hidden"),
        }
      : {
          kind: "action" as const,
          id: "unhide",
          label: t("xlsx.sheetTab.unhide"),
          onSelect: () => onSetState(name, "visible"),
        },
    { kind: "divider", id: "d3" },
    {
      kind: "action",
      id: "delete",
      label: t("xlsx.sheetTab.delete"),
      // Excel forbids deleting the only worksheet in a workbook.
      disabled: sheets.length <= 1,
      onSelect: () => onDelete(name),
    },
  ];
}
