"use client";

import type { ComponentType, ReactNode } from "react";
import { FileSpreadsheet, FileText, Presentation } from "lucide-react";
import { InlineSpinner } from "./InlineSpinner";
import type { ProductKind } from "./types";

/**
 * Shared loading affordance for every editor product (Word, Excel,
 * PowerPoint).
 *
 * Three variants cover the moments where the app needs to say "wait":
 *
 *  - `fill` — the editor isn't mounted yet. Used by Next.js dynamic
 *    `loading` fallbacks where there's nothing on screen but the
 *    `<main>` shell, so we just fill it. When `product` is passed
 *    we render an "elevated" splash with the product's Lucide icon
 *    inside an accent ring — same badge convention as `EmptyState`,
 *    so the very first paint already feels like the right product.
 *  - `splash` — the editor *is* mounted (chrome included) but the
 *    document/workbook/deck is still being parsed. Renders the same
 *    elevated layout as `fill + product` but with `position: fixed`
 *    and a solid `bg-background`, so the splash covers toolbars,
 *    status bar, and sidebars on top of the chrome. The user sees
 *    one continuous splash from page navigation through agent
 *    bootstrap, regardless of whether the JS chunk was cached.
 *    `splash` requires `product` (the elevated layout is the point).
 *  - `overlay` — the older, minimal `Loader2 + label` overlay,
 *    anchored with `absolute inset-0` over a `relative` container.
 *    Kept for narrow cases where you really do want toolbars to
 *    stay visible behind the loader; bootstrap should use `splash`.
 *
 * Keep the implementation deliberately small. The point is sameness:
 * every product should hit the same "loading" beat, with the same
 * spinner (`InlineSpinner`) the rest of the shell uses for save and
 * export.
 */
export interface LoadingScreenProps {
  readonly variant?: "fill" | "overlay" | "splash";
  /** When set, the `fill` variant renders an elevated badge with
   * this product's icon and an accent ring spinning around it, and
   * the default label becomes product-specific. Required by `splash`
   * (the elevated layout is the whole point of that variant).
   * Ignored by `overlay` because the editor chrome already shows
   * the product. */
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

  if (variant === "splash" && product) {
    return (
      <div
        className="fixed inset-0 z-50 flex flex-col items-center justify-center gap-3 bg-background px-6 text-center"
        data-testid={testId}
        role="status"
        aria-live="polite"
      >
        <ProductBadge product={product} />
        <p className="text-sm text-secondary">{resolvedLabel}</p>
      </div>
    );
  }

  if (variant === "fill" && product) {
    return (
      <div
        className="flex h-full w-full flex-col items-center justify-center gap-3 px-6 text-center"
        data-testid={testId}
        role="status"
        aria-live="polite"
      >
        <ProductBadge product={product} />
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

/**
 * Shared elevated badge used by both `fill + product` and `splash`:
 * a 48 px circular badge with the product's Lucide icon and a thin
 * 2 px accent ring rotating around it. `border-transparent` keeps
 * the other three segments invisible; the inline style sets the top
 * segment to the accent so we don't depend on Tailwind arbitrary-
 * value parsing for CSS variables.
 */
function ProductBadge({ product }: { readonly product: ProductKind }): ReactNode {
  const Icon = PRODUCT_ICON[product];
  return (
    <div className="relative flex h-12 w-12 items-center justify-center rounded-full bg-hover text-secondary">
      <Icon size={22} />
      <span
        className="pointer-events-none absolute -inset-1 rounded-full border-2 border-transparent animate-spin"
        style={{ borderTopColor: "var(--accent)" }}
        aria-hidden="true"
      />
    </div>
  );
}
