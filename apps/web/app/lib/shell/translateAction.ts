/**
 * Translate an `ActionDescriptor` into the active locale.
 *
 * Convention-based key derivation keeps the catalogue files
 * (`packages/<format>/src/actions/catalogue.ts`) free of i18n bookkeeping
 * and the message catalogues free of duplicated metadata. For an action
 * with `id: "docx.insert-image"` and `section: "Insert"` the resolver
 * looks up:
 *
 *   - `actions.docx.insert-image.label`
 *   - `actions.docx.insert-image.description`
 *   - `actions.docx.sections.Insert`
 *
 * If a key is missing the existing `useTranslator` returns the key
 * verbatim — the resolver detects that echo and falls back to the
 * catalogue's English string. This keeps the CLI 1:1 with the catalogue
 * source and prevents half-translated palettes from breaking the UI
 * during translation rollouts.
 */

import type { ActionDescriptor } from "@officeai/core";
import type { TranslateVars } from "../i18n/useTranslator";

export type TranslateFn = (key: string, vars?: TranslateVars) => string;

export interface TranslatedActionStrings {
  readonly label: string;
  readonly description: string;
  readonly section: string;
}

export function translateAction(action: ActionDescriptor, t: TranslateFn): TranslatedActionStrings {
  const format = formatOf(action.id);
  const localId = action.id.startsWith(`${format}.`) ? action.id.slice(format.length + 1) : action.id;
  const labelKey = `actions.${format}.${localId}.label`;
  const descKey = `actions.${format}.${localId}.description`;
  const sectionKey = `actions.${format}.sections.${action.section}`;
  return {
    label: orFallback(t(labelKey), labelKey, action.label),
    description: orFallback(t(descKey), descKey, action.description),
    section: orFallback(t(sectionKey), sectionKey, action.section),
  };
}

function formatOf(id: string): string {
  const dot = id.indexOf(".");
  return dot < 0 ? id : id.slice(0, dot);
}

function orFallback(value: string, key: string, fallback: string): string {
  // `useTranslator` returns the key verbatim when no translation
  // exists. That's the signal to fall back to the catalogue's
  // English source string.
  return value === key ? fallback : value;
}
