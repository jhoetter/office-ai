"use client";

import type { ComponentType, ReactNode } from "react";
import { FileSpreadsheet, FileText, Presentation } from "lucide-react";
import { InlineSpinner } from "./InlineSpinner";
import type { ProductKind } from "./types";

/**
 * Shared loading affordance for every editor product (Word, Excel,
 * PowerPoint).
 *
 * Two variants cover the two moments where the app needs to say
 * "wait":
 *
 *  - `fill` — the editor isn't mounted yet. Used by Next.js dynamic
 *    `loading` fallbacks where there's nothing on screen but the
 *    `<main>` shell, so we just fill it. When `product` is passed
 *    we render an "elevated" splash with the product's Lucide icon
 *    inside an accent ring — same badge convention as `EmptyState`,
 *    so the very first paint already feels like the right product.
 *  - `overlay` — the editor chrome is already painted and we're
 *    parsing the document/workbook/deck. Anchored with
 *    `absolute inset-0` over a `relative` container so the toolbars
 *    stay visible and clickable while the canvas hydrates. The
 *    surrounding chrome already conveys product identity, so this
 *    variant intentionally stays minimal — a small spinner and
 *    a label, nothing more.
 *
 * Keep the implementation deliberately small. The point is sameness:
 * every product should hit the same "loading" beat, with the same
 * spinner (`InlineSpinner`) the rest of the shell uses for save and
 * export.
 */
export interface LoadingScreenProps {
  readonly variant?: "fill" | "overlay";
  /** When set, the `fill` variant renders an elevated badge with
   * this product's icon and an accent ring spinning around it. The
   * default label also becomes product-specific. Ignored by
   * `overlay` because the editor chrome already shows the product. */
  readonly product?: ProductKind;
  /** Defaults to "Loading…" (or "Loading <product>…" when
   * `product` is set). Pass a more specific label when the generic
   * one would be confusing — e.g. when the editor is mounted but a
   * different artefact is being prepared. */
  readonly label?: string;
  readonly testId?: string;
}

const PRODUCT_ICON: Record<ProductKind, ComponentType<{ size?: number }>> = {
  docx: FileText,
  xlsx: FileSpreadsheet,
  pptx: Presentation,
};

const PRODUCT_LABEL: Record<ProductKind, string> = {
  docx: "Loading Word document…",
  xlsx: "Loading Excel workbook…",
  pptx: "Loading PowerPoint presentation…",
};

export function LoadingScreen({
  variant = "fill",
  product,
  label,
  testId = "loading-screen",
}: LoadingScreenProps): ReactNode {
  const resolvedLabel = label ?? (product ? PRODUCT_LABEL[product] : "Loading…");

  if (variant === "fill" && product) {
    const Icon = PRODUCT_ICON[product];
    return (
      <div
        className="flex h-full w-full flex-col items-center justify-center gap-3 px-6 text-center"
        data-testid={testId}
        role="status"
        aria-live="polite"
      >
        <div className="relative flex h-12 w-12 items-center justify-center rounded-full bg-hover text-secondary">
          <Icon size={22} />
          {/* A 2 px ring with a single accent segment, rotating
           * around the product badge. `border-transparent` keeps the
           * other three segments invisible; the inline style sets
           * the top segment to the accent so we don't depend on
           * Tailwind arbitrary-value parsing for CSS variables. */}
          <span
            className="pointer-events-none absolute -inset-1 rounded-full border-2 border-transparent animate-spin"
            style={{ borderTopColor: "var(--accent)" }}
            aria-hidden="true"
          />
        </div>
        <p className="text-sm text-secondary">{resolvedLabel}</p>
      </div>
    );
  }

  const containerClass =
    variant === "overlay"
      ? "absolute inset-0 flex items-center justify-center text-sm text-secondary"
      : "flex h-full w-full items-center justify-center text-sm text-secondary";

  return (
    <div className={containerClass} data-testid={testId} role="status" aria-live="polite">
      <InlineSpinner size={14} className="mr-2" />
      {resolvedLabel}
    </div>
  );
}
