import type { RibbonCatalogue, RibbonTab } from "./RibbonTypes";

/**
 * Pure helpers for the Ribbon's tab resolution. Extracted so the
 * decision logic is unit-testable without React / DOM stubs.
 */

export function visibleRibbonTabs<C>(catalogue: RibbonCatalogue<C>, ctx: C): ReadonlyArray<RibbonTab<C>> {
  return catalogue.tabs.filter((t) => (t.visible ? t.visible(ctx) : true));
}

/**
 * Compute the currently-active tab id. Returns the first contextual
 * tab whose `autoActivateWhen(ctx)` fires; otherwise falls back to
 * the user-pinned tab. If the pinned tab is no longer visible the
 * caller is expected to repair the pin (the React Ribbon does this
 * via an effect); meanwhile we return the first visible tab so the
 * surface never renders empty.
 *
 * `suppressedAutoSignature` is Office's "I clicked another tab while
 * a contextual one was firing" affordance: the caller passes the
 * `autoActivationSignature(visibleTabs, ctx)` value captured at the
 * moment the user pinned a tab. As long as the live signature
 * matches, the user's choice wins over auto-activation. Once the
 * signature changes (selection moved to a different shape kind, or
 * cleared and a new selection started), the pin no longer overrides
 * and the next selection's contextual tab takes over again — which
 * is what users expect from PowerPoint / Word.
 */
export function resolveActiveTabId<C>(
  visibleTabs: ReadonlyArray<RibbonTab<C>>,
  ctx: C,
  pinnedId: string,
  suppressedAutoSignature?: string
): string {
  const autoTabs = visibleTabs.filter((t) => t.contextual && t.autoActivateWhen?.(ctx));
  const liveSignature = autoTabs.map((t) => t.id).join("|");
  const userOverride =
    suppressedAutoSignature !== undefined &&
    suppressedAutoSignature === liveSignature &&
    visibleTabs.some((t) => t.id === pinnedId);
  if (autoTabs.length > 0 && !userOverride) return autoTabs[0].id;
  if (visibleTabs.some((t) => t.id === pinnedId)) return pinnedId;
  return visibleTabs[0]?.id ?? "";
}

/**
 * Stable string identifying the set of contextual tabs whose
 * `autoActivateWhen` is firing for `ctx`. Used by the Ribbon to
 * detect when the user's "I want to stay on this tab" override is
 * still applicable: same signature → same selection → keep the pin;
 * different signature → selection changed → drop the override.
 */
export function autoActivationSignature<C>(visibleTabs: ReadonlyArray<RibbonTab<C>>, ctx: C): string {
  const ids: string[] = [];
  for (const t of visibleTabs) {
    if (t.contextual && t.autoActivateWhen?.(ctx)) ids.push(t.id);
  }
  return ids.join("|");
}
