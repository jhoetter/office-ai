/**
 * Hook + helper for resolving a catalogue action by id at the call
 * site. Toolbar buttons / context menus / keyboard hint badges all
 * pull their label, description, shortcut, and icon hint from the
 * shared `ActionDescriptor` instead of re-declaring them.
 *
 * `useAction(catalogue, id)` is the React-friendly wrapper; it
 * memoises the descriptor lookup. `getAction(catalogue, id)` is the
 * pure version, useful inside other hooks or non-React code.
 *
 * Both throw if the id is missing from the catalogue. We prefer a
 * loud failure over silent rendering of a button with no label —
 * typos are caught the first time the component renders, not later
 * via "why is my Cmd+K entry missing?" debugging.
 */

import { useMemo } from "react";
import type { ActionDescriptor } from "@officeai/core";
import { useTranslator } from "../i18n";
import { translateAction } from "./translateAction";

export interface ResolvedAction {
  readonly id: string;
  readonly label: string;
  readonly description: string;
  readonly section: string;
  readonly shortcut?: string;
  readonly icon?: string;
}

export function getAction(
  catalogue: ReadonlyArray<ActionDescriptor>,
  id: string
): ResolvedAction {
  for (const a of catalogue) {
    if (a.id !== id) continue;
    return {
      id: a.id,
      label: a.label,
      description: a.description,
      section: a.section,
      ...(a.shortcut ? { shortcut: a.shortcut } : {}),
      ...(a.icon ? { icon: a.icon } : {}),
    };
  }
  throw new Error(
    `useAction/getAction: no catalogue entry for id "${id}". Add it to the format's actions/catalogue.ts or fix the typo.`
  );
}

export function useAction(
  catalogue: ReadonlyArray<ActionDescriptor>,
  id: string
): ResolvedAction {
  const { t } = useTranslator();
  return useMemo(() => {
    const base = getAction(catalogue, id);
    const descriptor = catalogue.find((a) => a.id === id);
    if (!descriptor) return base;
    const strings = translateAction(descriptor, t);
    return {
      ...base,
      label: strings.label,
      description: strings.description,
      section: strings.section,
    };
  }, [catalogue, id, t]);
}
