"use client";

import { useI18n } from "./I18nProvider";
import type { Locale } from "./types";

const OPTIONS: ReadonlyArray<{ readonly value: Locale; readonly label: string; readonly title: string }> = [
  { value: "en", label: "EN", title: "English" },
  { value: "de", label: "DE", title: "Deutsch" },
];

interface LocaleToggleProps {
  readonly className?: string;
  /** Compact (single button cycling) for cramped toolbars. */
  readonly compact?: boolean;
}

/**
 * Mirrors the visual treatment of `ThemeToggle` so the two
 * controls feel like a pair when placed side-by-side in the
 * editor top-bar. Switching is purely client-side; the cookie
 * write happens inside the provider.
 */
export function LocaleToggle({ className, compact }: LocaleToggleProps) {
  const { locale, setLocale } = useI18n();

  if (compact) {
    const currentIdx = OPTIONS.findIndex((o) => o.value === locale);
    const current = OPTIONS[currentIdx >= 0 ? currentIdx : 0];
    const next = OPTIONS[(OPTIONS.indexOf(current) + 1) % OPTIONS.length];
    return (
      <button
        type="button"
        onClick={() => setLocale(next.value)}
        className={
          "flex h-7 min-w-7 items-center justify-center rounded-md px-1.5 text-[11px] font-medium text-secondary hover:bg-hover hover:text-foreground transition-colors duration-150 " +
          (className ?? "")
        }
        title={`${current.title} → ${next.title}`}
        aria-label={`Switch language (currently ${current.title})`}
      >
        {current.label}
      </button>
    );
  }

  return (
    <div
      role="group"
      aria-label="Language"
      className={"inline-flex items-center rounded-md bg-hover p-0.5 gap-0.5 " + (className ?? "")}
    >
      {OPTIONS.map(({ value, label, title }) => {
        const active = locale === value;
        return (
          <button
            key={value}
            type="button"
            onClick={() => setLocale(value)}
            aria-pressed={active}
            title={title}
            className={
              "flex h-6 items-center justify-center rounded px-2 text-[11px] font-medium transition-colors duration-150 " +
              (active ? "bg-background text-foreground shadow-sm" : "text-secondary hover:text-foreground")
            }
          >
            {label}
          </button>
        );
      })}
    </div>
  );
}

export type { LocaleToggleProps };
