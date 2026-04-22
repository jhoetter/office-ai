"use client";

/**
 * Shared "gallery" button for the ribbon — opens a popover containing
 * a grid of preview tiles (e.g. theme thumbnails, animation effects,
 * picture styles, transition effects).
 *
 * The trigger is sized for a single ribbon row and shows a chevron;
 * the popover uses a compact tile grid driven by the supplied items.
 *
 * Used by Phase 3 (DOCX themes / picture styles), Phase 4 (XLSX
 * chart styles / cell styles), Phase 5 (PPTX transitions /
 * animations / picture styles).
 */

import { useEffect, useRef, useState, type ReactNode } from "react";
import { ChevronDown } from "lucide-react";
import { cn } from "@officeai/ui";
import { ToolbarMenu } from "../ToolbarMenu";

export interface GalleryItem {
  readonly id: string;
  readonly label: string;
  readonly preview: ReactNode;
  readonly active?: boolean;
  readonly disabled?: boolean;
}

export interface RibbonGalleryButtonProps {
  readonly icon: ReactNode;
  readonly label: string;
  readonly testId: string;
  readonly disabled?: boolean;
  readonly items: ReadonlyArray<GalleryItem>;
  readonly onPick: (id: string) => void;
  readonly columns?: number;
  /** Optional footer rendered below the grid (e.g. "More options…"). */
  readonly footer?: ReactNode;
}

export function RibbonGalleryButton(props: RibbonGalleryButtonProps): ReactNode {
  const { icon, label, testId, disabled = false, items, onPick, columns = 4, footer } = props;
  const [open, setOpen] = useState(false);
  const triggerRef = useRef<HTMLButtonElement | null>(null);

  // Close on Esc
  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") setOpen(false);
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [open]);

  return (
    <span className="inline-flex items-center">
      <button
        ref={triggerRef}
        type="button"
        data-testid={testId}
        title={label}
        aria-label={label}
        aria-expanded={open}
        aria-haspopup="menu"
        disabled={disabled}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => setOpen((v) => !v)}
        className={cn(
          "inline-flex h-7 items-center gap-0.5 rounded px-1 text-foreground hover:bg-hover disabled:opacity-50",
          open && "bg-hover"
        )}
      >
        {icon}
        <ChevronDown size={10} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        role="menu"
        testId={`${testId}-menu`}
        className="rounded-md border border-divider bg-surface p-2 shadow-lg"
      >
        <div
          className="grid gap-1"
          style={{ gridTemplateColumns: `repeat(${columns}, minmax(0, 1fr))` }}
        >
          {items.map((item) => (
            <button
              key={item.id}
              type="button"
              role="menuitem"
              data-testid={`${testId}-item-${item.id}`}
              title={item.label}
              aria-label={item.label}
              aria-pressed={item.active}
              disabled={item.disabled}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => {
                setOpen(false);
                onPick(item.id);
              }}
              className={cn(
                "group flex h-12 w-16 flex-col items-center justify-center rounded border border-divider p-1 text-xs text-foreground hover:bg-hover hover:border-[var(--accent)] disabled:opacity-40",
                item.active && "border-[var(--accent)] bg-accent-soft text-accent"
              )}
            >
              <div className="flex h-7 w-full items-center justify-center overflow-hidden">
                {item.preview}
              </div>
              <span className="mt-0.5 line-clamp-1 text-[9px] text-secondary group-hover:text-foreground">
                {item.label}
              </span>
            </button>
          ))}
        </div>
        {footer ? (
          <>
            <div className="my-1 h-px bg-divider" />
            <div className="px-1 py-0.5 text-xs">{footer}</div>
          </>
        ) : null}
      </ToolbarMenu>
    </span>
  );
}
