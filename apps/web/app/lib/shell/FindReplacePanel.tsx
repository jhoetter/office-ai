"use client";

import { useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { ChevronDown, ChevronUp, Replace, Search, X } from "@officeai/ui/sonaloop-icons";
import { cn } from "@officeai/ui";
import { useTranslator } from "@/lib/i18n";
import type { FindAdapter, FindMatch, FindOptions } from "./types";

export interface FindReplacePanelProps {
  readonly adapter: FindAdapter | undefined;
  readonly open: boolean;
  readonly mode: "find" | "replace";
  readonly onClose: () => void;
}

/**
 * Floating Find / Replace panel docked top-right of the document.
 *
 * The product owns iteration through a typed `FindAdapter`:
 *   - DOCX walks ProseMirror text positions
 *   - XLSX scans cell values + formulas
 *   - PPTX walks slide / shape text
 *
 * Cmd+F opens find. Cmd+Alt+F opens with replace exposed.
 * Enter advances; Shift+Enter goes back; Esc closes.
 */
export function FindReplacePanel({ adapter, open, mode, onClose }: FindReplacePanelProps): ReactNode {
  const { t } = useTranslator();
  const [query, setQuery] = useState("");
  const [replacement, setReplacement] = useState("");
  const [showReplace, setShowReplace] = useState(mode === "replace");
  const [opts, setOpts] = useState<FindOptions>({
    caseSensitive: false,
    wholeWord: false,
    regex: false,
  });
  const [activeIdx, setActiveIdx] = useState(0);
  const findInputRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    setShowReplace(mode === "replace");
  }, [mode]);

  useEffect(() => {
    if (!open) return;
    requestAnimationFrame(() => findInputRef.current?.focus());
  }, [open]);

  const matches = useMemo(() => {
    if (!open || !adapter || query.length === 0) return [];
    try {
      return adapter.findAll(query, opts);
    } catch (e) {
      console.warn("[find] adapter.findAll threw", e);
      return [];
    }
  }, [adapter, open, query, opts]);

  useEffect(() => {
    setActiveIdx(0);
  }, [query, matches.length]);

  useEffect(() => {
    if (!open || !adapter || matches.length === 0) return;
    const m = matches[activeIdx];
    if (m) adapter.gotoMatch(m);
  }, [open, adapter, matches, activeIdx]);

  if (!open || !adapter) return null;

  const current: FindMatch | null = matches[activeIdx] ?? null;
  const total = matches.length;

  const next = () => setActiveIdx((i) => (total === 0 ? 0 : (i + 1) % total));
  const prev = () => setActiveIdx((i) => (total === 0 ? 0 : (i - 1 + total) % total));

  const replaceCurrent = async () => {
    if (!current) return;
    await adapter.replaceMatch(current, replacement);
    // Adapter mutates the document; matches will refresh on next render.
  };

  const replaceAll = async () => {
    if (query.length === 0) return;
    await adapter.replaceAll(query, replacement, opts);
  };

  return (
    <div
      role="dialog"
      aria-label={t("common.findAndReplace")}
      className="absolute right-3 top-3 z-20 w-[320px] rounded-md border border-divider bg-background p-2 shadow-lg"
      data-testid="find-replace"
    >
      <div className="flex items-center gap-1">
        <button
          type="button"
          onClick={() => setShowReplace((v) => !v)}
          className="inline-flex h-7 w-7 items-center justify-center rounded-md text-secondary hover:bg-hover hover:text-foreground"
          title={showReplace ? t("common.hideReplace") : t("common.showReplace")}
          aria-label={showReplace ? t("common.hideReplace") : t("common.showReplace")}
          data-testid="find-toggle-replace"
        >
          {showReplace ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
        </button>
        <div className="relative flex-1">
          <Search
            size={12}
            className="pointer-events-none absolute left-2 top-1/2 -translate-y-1/2 text-tertiary"
          />
          <input
            ref={findInputRef}
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") {
                e.preventDefault();
                if (e.shiftKey) prev();
                else next();
              } else if (e.key === "Escape") {
                e.preventDefault();
                onClose();
              }
            }}
            placeholder={t("common.findPlaceholder")}
            className="h-7 w-full rounded-md border border-divider bg-background pl-6 pr-12 text-sm text-foreground placeholder:text-tertiary focus:border-[var(--accent)] focus:outline-none"
            aria-label={t("common.find")}
            data-testid="find-input"
          />
          <span
            className="pointer-events-none absolute right-2 top-1/2 -translate-y-1/2 text-[11px] tabular-nums text-tertiary"
            aria-live="polite"
          >
            {query.length === 0 ? "" : total === 0 ? "0/0" : `${activeIdx + 1}/${total}`}
          </span>
        </div>
        <button
          type="button"
          onClick={prev}
          className="inline-flex h-7 w-7 items-center justify-center rounded-md text-secondary hover:bg-hover hover:text-foreground"
          aria-label={t("common.previousMatch")}
          title={`${t("common.previousMatch")} (Shift+Enter)`}
          data-testid="find-prev"
        >
          <ChevronUp size={14} />
        </button>
        <button
          type="button"
          onClick={next}
          className="inline-flex h-7 w-7 items-center justify-center rounded-md text-secondary hover:bg-hover hover:text-foreground"
          aria-label={t("common.nextMatch")}
          title={`${t("common.nextMatch")} (Enter)`}
          data-testid="find-next"
        >
          <ChevronDown size={14} />
        </button>
        <button
          type="button"
          onClick={onClose}
          className="inline-flex h-7 w-7 items-center justify-center rounded-md text-secondary hover:bg-hover hover:text-foreground"
          aria-label={t("common.closeFind")}
          title={`${t("common.close")} (Esc)`}
          data-testid="find-close"
        >
          <X size={14} />
        </button>
      </div>

      {showReplace ? (
        <div className="mt-1.5 flex items-center gap-1">
          <span className="inline-flex h-7 w-7 items-center justify-center text-tertiary">
            <Replace size={12} />
          </span>
          <input
            value={replacement}
            onChange={(e) => setReplacement(e.target.value)}
            placeholder={t("common.replacePlaceholder")}
            className="h-7 flex-1 rounded-md border border-divider bg-background px-2 text-sm text-foreground placeholder:text-tertiary focus:border-[var(--accent)] focus:outline-none"
            aria-label={t("common.replacePlaceholder")}
            data-testid="replace-input"
          />
          <button
            type="button"
            onClick={() => void replaceCurrent()}
            disabled={!current}
            className={cn(
              "inline-flex h-7 items-center rounded-md border border-divider px-2 text-xs text-foreground hover:bg-hover",
              !current && "pointer-events-none opacity-40"
            )}
            data-testid="replace-one"
          >
            {t("common.replace")}
          </button>
          <button
            type="button"
            onClick={() => void replaceAll()}
            disabled={query.length === 0}
            className={cn(
              "inline-flex h-7 items-center rounded-md border border-divider px-2 text-xs text-foreground hover:bg-hover",
              query.length === 0 && "pointer-events-none opacity-40"
            )}
            data-testid="replace-all"
          >
            {t("common.replaceAll")}
          </button>
        </div>
      ) : null}

      <div className="mt-1.5 flex items-center gap-2 px-1 text-[11px] text-secondary">
        <FindOptToggle
          label="Aa"
          tip={t("common.caseSensitive")}
          testId="case-sensitive"
          active={opts.caseSensitive}
          onClick={() => setOpts((o) => ({ ...o, caseSensitive: !o.caseSensitive }))}
        />
        <FindOptToggle
          label="ab"
          tip={t("common.wholeWord")}
          testId="whole-word"
          active={opts.wholeWord}
          onClick={() => setOpts((o) => ({ ...o, wholeWord: !o.wholeWord }))}
        />
        <FindOptToggle
          label=".*"
          tip={t("common.regex")}
          testId="regex"
          active={opts.regex}
          onClick={() => setOpts((o) => ({ ...o, regex: !o.regex }))}
        />
      </div>
    </div>
  );
}

function FindOptToggle({
  label,
  tip,
  testId,
  active,
  onClick,
}: {
  readonly label: string;
  readonly tip: string;
  readonly testId: string;
  readonly active: boolean;
  readonly onClick: () => void;
}): ReactNode {
  return (
    <button
      type="button"
      onClick={onClick}
      title={tip}
      aria-label={tip}
      aria-pressed={active}
      className={cn(
        "inline-flex h-5 w-5 items-center justify-center rounded text-[10px] font-medium",
        active ? "bg-[var(--accent)] text-white" : "text-secondary hover:bg-hover"
      )}
      data-testid={`find-opt-${testId}`}
    >
      {label}
    </button>
  );
}
