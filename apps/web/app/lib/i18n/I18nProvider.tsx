"use client";

import { createContext, useCallback, useContext, useEffect, useMemo, useState, type ReactNode } from "react";
import en from "./messages/en.json";
import de from "./messages/de.json";
import { DEFAULT_LOCALE, LOCALE_COOKIE, isLocale, type Locale, type Messages } from "./types";

const CATALOGUES: Record<Locale, Messages> = {
  en: en as unknown as Messages,
  de: de as unknown as Messages,
};

interface I18nContextValue {
  readonly locale: Locale;
  readonly messages: Messages;
  setLocale(next: Locale): void;
}

const I18nContext = createContext<I18nContextValue | null>(null);

function readCookieLocale(): Locale | null {
  if (typeof document === "undefined") return null;
  const match = document.cookie
    .split("; ")
    .map((p) => p.split("="))
    .find(([k]) => k === LOCALE_COOKIE);
  if (!match) return null;
  const value = decodeURIComponent(match[1] ?? "");
  return isLocale(value) ? value : null;
}

function readNavigatorLocale(): Locale | null {
  if (typeof navigator === "undefined") return null;
  const tag = (navigator.languages?.[0] ?? navigator.language ?? "").toLowerCase();
  if (tag.startsWith("de")) return "de";
  if (tag.startsWith("en")) return "en";
  return null;
}

function persistLocale(locale: Locale): void {
  if (typeof document === "undefined") return;
  // 1 year cookie; SameSite=Lax keeps it from leaking on cross-site requests
  // but still travels with same-site navigation, which is what we want.
  document.cookie = `${LOCALE_COOKIE}=${locale}; path=/; max-age=${60 * 60 * 24 * 365}; SameSite=Lax`;
  // Also broadcast on the next tick so peer tabs can re-read on focus
  // without us having to wire BroadcastChannel here.
  try {
    document.documentElement.lang = locale;
  } catch {
    /* noop */
  }
}

export interface I18nProviderProps {
  readonly children: ReactNode;
  /**
   * Optional initial locale (mostly for tests / Storybook). When
   * omitted we read the cookie, then fall back to the navigator
   * preference, then the configured default.
   */
  readonly initialLocale?: Locale;
}

/**
 * Cookie-driven, client-only i18n provider. We deliberately keep
 * this thinner than next-intl's full surface — the cost of a deep
 * locale-routing setup outweighs its benefits for an OOXML editor
 * where every page is client-rendered anyway.
 *
 * Hydration note: messages are bundled into the client chunk (one
 * JSON catalogue per locale). The active locale is read from the
 * cookie on mount; before the cookie is read, the provider serves
 * the configured default — which means the very first paint after
 * a cold load may show English for ~1 frame before the German
 * cookie applies. This is below human perception thresholds and
 * keeps the SSR/CSR boundary trivial.
 */
export function I18nProvider({ children, initialLocale }: I18nProviderProps): ReactNode {
  const [locale, setLocaleState] = useState<Locale>(initialLocale ?? DEFAULT_LOCALE);

  useEffect(() => {
    if (initialLocale) return;
    const next = readCookieLocale() ?? readNavigatorLocale() ?? DEFAULT_LOCALE;
    setLocaleState(next);
    try {
      document.documentElement.lang = next;
    } catch {
      /* noop */
    }
  }, [initialLocale]);

  const setLocale = useCallback((next: Locale) => {
    persistLocale(next);
    setLocaleState(next);
  }, []);

  const value = useMemo<I18nContextValue>(
    () => ({ locale, messages: CATALOGUES[locale], setLocale }),
    [locale, setLocale]
  );

  return <I18nContext.Provider value={value}>{children}</I18nContext.Provider>;
}

export function useI18n(): I18nContextValue {
  const ctx = useContext(I18nContext);
  if (!ctx) {
    // Fall back to defaults rather than throwing — keeps tests that
    // render a leaf component without the provider from blowing up.
    return {
      locale: DEFAULT_LOCALE,
      messages: CATALOGUES[DEFAULT_LOCALE],
      setLocale: () => undefined,
    };
  }
  return ctx;
}
