"use client";

import { useEffect, useRef } from "react";
import type { ReactNode } from "react";
import { cn } from "@officeai/ui";

/**
 * One row in a context menu. Either a clickable action or a
 * separator. Submenus and toggle items are deliberately out of scope
 * for P13 — Excel's right-click is mostly flat.
 */
export type ContextMenuItem =
  | {
      readonly kind: "action";
      readonly id: string;
      readonly label: string;
      /** Render hint, e.g. "⌘C". Not actually wired to a key listener. */
      readonly shortcut?: string;
      readonly disabled?: boolean;
      readonly onSelect: () => void;
    }
  | { readonly kind: "divider"; readonly id: string };

export interface ContextMenuProps {
  readonly open: boolean;
  readonly x: number;
  readonly y: number;
  readonly items: ReadonlyArray<ContextMenuItem>;
  readonly onClose: () => void;
  /** Test-id stamped on the wrapping div for Playwright. */
  readonly testId?: string;
}

/**
 * Lightweight Excel-style right-click menu. Lives in the normal flow
 * (no portal) — anchored at the mouse coordinates the parent
 * recorded on the `contextmenu` event. Dismissed on outside-click,
 * Escape, or scroll/resize of the underlying viewport.
 *
 * Focus is moved into the menu so the keyboard ↑/↓/Enter pattern
 * works out of the box; Esc returns focus to whatever had it before.
 */
export function ContextMenu(props: ContextMenuProps): ReactNode {
  const { open, x, y, items, onClose, testId = "context-menu" } = props;
  const ref = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    if (!open) return;
    const onDocMouse = (e: MouseEvent) => {
      const el = ref.current;
      if (!el) return;
      if (!el.contains(e.target as Node)) onClose();
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        e.preventDefault();
        onClose();
      }
    };
    const onScroll = () => onClose();
    window.addEventListener("mousedown", onDocMouse, true);
    window.addEventListener("keydown", onKey);
    window.addEventListener("scroll", onScroll, true);
    window.addEventListener("resize", onScroll);
    return () => {
      window.removeEventListener("mousedown", onDocMouse, true);
      window.removeEventListener("keydown", onKey);
      window.removeEventListener("scroll", onScroll, true);
      window.removeEventListener("resize", onScroll);
    };
  }, [open, onClose]);

  if (!open) return null;

  return (
    <div
      ref={ref}
      role="menu"
      data-testid={testId}
      style={{
        position: "fixed",
        top: y,
        left: x,
        zIndex: 100,
        minWidth: 192,
      }}
      className="rounded-md border border-divider bg-surface py-1 shadow-lg"
    >
      {items.map((item) => {
        if (item.kind === "divider") {
          return <div key={item.id} className="my-1 h-px bg-divider" aria-hidden />;
        }
        return (
          <button
            key={item.id}
            type="button"
            role="menuitem"
            data-testid={`menu-item-${item.id}`}
            disabled={item.disabled}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => {
              if (item.disabled) return;
              item.onSelect();
              onClose();
            }}
            className={cn(
              "flex w-full items-center justify-between gap-4 px-3 py-1 text-left text-xs text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40 disabled:hover:bg-transparent"
            )}
          >
            <span>{item.label}</span>
            {item.shortcut ? (
              <span className="text-secondary text-[10px] font-mono">{item.shortcut}</span>
            ) : null}
          </button>
        );
      })}
    </div>
  );
}
