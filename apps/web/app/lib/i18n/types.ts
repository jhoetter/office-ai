/**
 * Lightweight i18n types. We intentionally don't use next-intl's
 * generated types because our message catalogues live as plain
 * JSON next to the consumer code — keeping the surface tiny lets
 * us swap engines later without touching every call-site.
 */
export type Locale = "en" | "de";

export const SUPPORTED_LOCALES: ReadonlyArray<Locale> = ["en", "de"] as const;
export const DEFAULT_LOCALE: Locale = "en";
export const LOCALE_COOKIE = "officeai.locale";

/**
 * The dictionary shape is `{ [namespace]: { [key]: string | nested } }`.
 * `string` may contain `{name}` ICU-style placeholders which `t(key, vars)`
 * substitutes at format time.
 */
export type MessageNode = string | { readonly [key: string]: MessageNode };
export type Messages = Readonly<Record<string, MessageNode>>;

export function isLocale(value: unknown): value is Locale {
  return typeof value === "string" && SUPPORTED_LOCALES.includes(value as Locale);
}
