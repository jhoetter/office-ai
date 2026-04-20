"use client";

import { useCallback } from "react";
import { useI18n } from "./I18nProvider";
import type { MessageNode, Messages } from "./types";

export type TranslateVars = Readonly<Record<string, string | number>>;

/**
 * Resolve a dotted path inside a nested message catalogue. Returns
 * the raw string if found, or `undefined` so callers can decide
 * whether to fall back.
 */
function lookup(messages: Messages, key: string): string | undefined {
  const parts = key.split(".");
  let cursor: MessageNode | undefined = messages as MessageNode;
  for (const part of parts) {
    if (typeof cursor !== "object" || cursor === null) return undefined;
    cursor = (cursor as Record<string, MessageNode>)[part];
    if (cursor === undefined) return undefined;
  }
  return typeof cursor === "string" ? cursor : undefined;
}

/**
 * Substitute ICU-light `{name}` placeholders. We deliberately do
 * not implement the full ICU plural/select grammar — none of our
 * catalogues need it yet, and shipping a 30 KB MessageFormat
 * compiler would be wasteful. If we hit a real plural case we'll
 * graduate to `intl-messageformat` for that one key.
 */
function format(template: string, vars: TranslateVars | undefined): string {
  if (!vars) return template;
  return template.replace(/\{(\w+)\}/g, (_, name: string) => {
    const value = vars[name];
    return value === undefined ? `{${name}}` : String(value);
  });
}

export function useTranslator(): {
  readonly t: (key: string, vars?: TranslateVars) => string;
  readonly locale: ReturnType<typeof useI18n>["locale"];
} {
  const { messages, locale } = useI18n();
  const t = useCallback(
    (key: string, vars?: TranslateVars) => {
      const raw = lookup(messages, key);
      // Falling back to the key itself is intentional: it makes
      // missing translations visible during development without
      // breaking the UI in production.
      if (raw === undefined) return key;
      return format(raw, vars);
    },
    [messages]
  );
  return { t, locale };
}
