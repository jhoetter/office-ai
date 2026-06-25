"use client";

import { useEffect, useState, type ComponentType, type ReactNode } from "react";
import { BookOpen, FileSpreadsheet, FileText, Presentation } from "@officeai/ui/sonaloop-icons";
import { InlineSpinner } from "./InlineSpinner";
import { useTranslator } from "@/lib/i18n";
import type { ProductKind } from "./types";

const SPLASH_FADE_OUT_MS = 280;
const SPIN_DURATION_MS = 1000;

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
 *  - `splash` — the bootstrap loader. Intended to be mounted at the
 *    editor *page* level (sibling to the dynamically-imported editor
 *    component), NOT inside the editor itself, so the badge `<span>`
 *    is created exactly once when the page hydrates and stays in the
 *    DOM through the `next/dynamic` resolution, the editor mount,
 *    and the agent-ready handoff. Renders the same elevated layout
 *    as `fill + product` but with `position: fixed` and a solid
 *    `bg-background`, so the splash covers toolbars, status bar, and
 *    sidebars on top of the chrome. Once the editor reports ready,
 *    flip `show` to `false` to fade the splash out over ~280 ms,
 *    unveiling the editor underneath. `splash` requires `product`
 *    (the elevated layout is the point). See e.g.
 *    `apps/web/app/editor/page.tsx` for the canonical wiring.
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
  /** Defaults to "Loading editor…" — the same wording the page-
   * level dynamic fallback uses, so the text doesn't flicker when
   * the splash hands off from `fill` to `splash` mid-bootstrap.
   * Pass a more specific label only when the generic one would be
   * confusing (e.g. a different artefact is being prepared). */
  readonly label?: string;
  readonly testId?: string;
  /** Splash-only. While `true` the overlay is rendered at full
   * opacity; flipping to `false` triggers a ~280 ms fade-out and
   * then unmounts the splash, "unveiling" the editor underneath.
   * Defaults to `true` (instant mount, never auto-dismisses). The
   * `fill` and `overlay` variants ignore this — they mount and
   * unmount instantly via the parent's conditional. */
  readonly show?: boolean;
}

const PRODUCT_ICON: Record<ProductKind, ComponentType<{ size?: number }>> = {
  docx: FileText,
  xlsx: FileSpreadsheet,
  pptx: Presentation,
  pdf: BookOpen,
};

export function LoadingScreen({
  variant = "fill",
  product,
  label,
  testId = "loading-screen",
  show = true,
}: LoadingScreenProps): ReactNode {
  const { t } = useTranslator();
  const fallbackLabel = product ? t(`loading.${product}`) : t("common.loading");
  const resolvedLabel = label ?? fallbackLabel;

  if (variant === "splash" && product) {
    return (
      <SplashOverlay show={show} testId={testId}>
        <ProductBadge product={product} />
        <p className="text-sm text-secondary">{resolvedLabel}</p>
      </SplashOverlay>
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
 *
 * The accent ring's `animationDelay` is anchored to the wall clock
 * (negative so it counts as "already running") so that if the badge
 * ever remounts (e.g. the `fill` variant is reused elsewhere
 * alongside a `splash`), the freshly-mounted ring picks up the same
 * rotation angle the old one was at instead of snapping back to 0°.
 *
 * The delay is intentionally set *after* mount: SSR'd HTML must be
 * deterministic, and `Date.now()` differs between the server render
 * and client hydration, which would trip React's hydration check
 * (the inline `style` would mismatch). We render with no delay on
 * the server and the first client paint, then sync the wall-clock
 * delay in an effect — visually indistinguishable for a fresh mount,
 * and only the remount-mid-bootstrap edge case relies on it anyway.
 */
function ProductBadge({ product }: { readonly product: ProductKind }): ReactNode {
  const Icon = PRODUCT_ICON[product];
  const [animationDelay, setAnimationDelay] = useState<string | undefined>(undefined);
  useEffect(() => {
    setAnimationDelay(`${-((Date.now() % SPIN_DURATION_MS) / 1000)}s`);
  }, []);
  return (
    <div className="relative flex h-12 w-12 items-center justify-center rounded-full bg-hover text-secondary">
      <Icon size={22} />
      <span
        className="pointer-events-none absolute -inset-1 rounded-full border-2 border-transparent animate-spin"
        style={{ borderTopColor: "var(--accent)", animationDelay }}
        aria-hidden="true"
      />
    </div>
  );
}

/**
 * Fade-out wrapper for the `splash` variant. Mounts opaque (no
 * fade-in — the splash is intended to be the very first thing the
 * page paints, so there's nothing to fade in from) and, when `show`
 * flips false, transitions opacity to 0 over `SPLASH_FADE_OUT_MS`
 * before unmounting. Because
 * the overlay is `position: fixed; z-50` with a solid `bg-background`,
 * the editor underneath is gradually unveiled through the fading
 * panel — the "unveiling" beat the bootstrap is supposed to feel like.
 */
function SplashOverlay({
  show,
  testId,
  children,
}: {
  readonly show: boolean;
  readonly testId: string;
  readonly children: ReactNode;
}): ReactNode {
  const [mounted, setMounted] = useState(show);
  const [opacity, setOpacity] = useState(1);

  useEffect(() => {
    if (show) {
      setMounted(true);
      setOpacity(1);
      return;
    }
    setOpacity(0);
    const t = setTimeout(() => setMounted(false), SPLASH_FADE_OUT_MS);
    return () => clearTimeout(t);
  }, [show]);

  if (!mounted) return null;

  return (
    <div
      className="fixed inset-0 z-50 flex flex-col items-center justify-center gap-3 bg-background px-6 text-center"
      data-testid={testId}
      role="status"
      aria-live="polite"
      aria-hidden={!show}
      style={{
        opacity,
        transition: `opacity ${SPLASH_FADE_OUT_MS}ms ease-out`,
        pointerEvents: show ? undefined : "none",
      }}
    >
      {children}
    </div>
  );
}
