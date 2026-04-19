"use client";

import type { ReactNode } from "react";
import { cn } from "@officeai/ui";

export interface ToolbarRowProps {
  /**
   * Toolbar contents — buttons, dividers, dropdowns. The row never
   * wraps; instead it scrolls horizontally when the contents exceed
   * the viewport width. This keeps the toolbar's height stable so
   * the body below never shifts when selection changes flip
   * conditional buttons.
   */
  readonly children: ReactNode;
  /**
   * Optional trailing slot (typically right-aligned actions like
   * "Present", "Edit mode", or a metadata strip). When provided it
   * is pinned to the row's right edge regardless of overflow on the
   * leading slot, so the action stays reachable.
   */
  readonly trailing?: ReactNode;
  readonly testId?: string;
  /** Accessible label for the toolbar landmark. */
  readonly ariaLabel?: string;
  /** Extra classes applied to the leading (scrollable) slot. */
  readonly leadingClassName?: string;
}

/**
 * Stable-height single-row toolbar primitive shared by DOCX, PPTX,
 * and XLSX editors.
 *
 * Why this exists: toolbar height previously varied with selection
 * state because each editor used `flex flex-wrap`. When the user
 * selected a shape (or a picture, or a single-column range), extra
 * buttons appeared, the row reflowed onto a second line, and the
 * canvas/grid jumped down a row. That made click targets move
 * underneath the cursor.
 *
 * Layout:
 *
 *   [  scrollable leading slot  →  →  →  ] [ pinned trailing slot ]
 *
 * - Fixed height (40px / `h-10`) regardless of contents.
 * - Leading slot uses `flex` (no wrap) + `overflow-x-auto`. On the
 *   rare narrow viewport the row scrolls horizontally instead of
 *   wrapping; the scrollbar is rendered thin and hidden until
 *   needed (`scrollbar-thin` style).
 * - Trailing slot stays right-aligned and never participates in the
 *   leading scroll, so primary actions like "Present" remain
 *   visible.
 *
 * Conditional buttons should still be rendered inside this row
 * (visible + disabled) rather than mounted/unmounted, so even the
 * leading slot's *width* stays stable across selection changes.
 */
export function ToolbarRow({
  children,
  trailing,
  testId,
  ariaLabel,
  leadingClassName,
}: ToolbarRowProps): ReactNode {
  return (
    <div
      role="toolbar"
      aria-label={ariaLabel}
      data-testid={testId}
      className="flex h-10 min-h-10 max-h-10 w-full items-stretch"
    >
      <div
        className={cn(
          // `[&>*]:shrink-0` keeps every direct child at its natural
          // width — so a flex parent with `overflow-x-auto` scrolls
          // horizontally instead of compressing button hit-targets
          // when many controls are present.
          "toolbar-row-scroll flex min-w-0 flex-1 items-center gap-1 overflow-x-auto overflow-y-hidden px-3 [&>*]:shrink-0",
          leadingClassName
        )}
      >
        {children}
      </div>
      {trailing ? (
        <div className="flex shrink-0 items-center gap-1 border-l border-divider/60 px-3">
          {trailing}
        </div>
      ) : null}
    </div>
  );
}
