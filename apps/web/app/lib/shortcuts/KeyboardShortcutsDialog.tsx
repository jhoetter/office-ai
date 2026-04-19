"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { Search, X } from "lucide-react";
import { useFocusTrap } from "@officeai/ui";
import {
  SHORTCUT_CATALOGS,
  type ShortcutCatalog,
  type ShortcutEntry,
  type ShortcutKey,
  type ShortcutProduct,
} from "./catalog";

export interface KeyboardShortcutsDialogProps {
  readonly product: ShortcutProduct;
  readonly open: boolean;
  readonly onClose: () => void;
}

/**
 * Searchable, OS-aware "Keyboard shortcuts" help dialog.
 *
 * Catalog data lives in {@link SHORTCUT_CATALOGS} so each product's
 * dialog stays in sync with what its surface actually does. The
 * dialog itself is presentational: backdrop click + Esc both close,
 * the search input filters by label/keys, and the list groups by
 * `category` in first-seen order.
 *
 * `Mod` renders as ⌘ on macOS and Ctrl on Windows/Linux; `Alt` as ⌥
 * / Alt; `Shift` as ⇧ / Shift. Detection happens once on mount via
 * `navigator.platform` so server-render emits a sane default
 * (`Ctrl`); the client effect upgrades it after hydration if needed.
 */
export function KeyboardShortcutsDialog(props: KeyboardShortcutsDialogProps) {
  const { product, open, onClose } = props;
  const catalog = SHORTCUT_CATALOGS[product];
  const [query, setQuery] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);
  const panelRef = useRef<HTMLDivElement>(null);
  const isMac = useIsMac();

  useEffect(() => {
    if (!open) return;
    setQuery("");
  }, [open]);

  useFocusTrap(panelRef, {
    enabled: open,
    initialFocusRef: inputRef,
    onEscape: onClose,
  });

  const filtered = useMemo(() => filterCatalog(catalog, query), [catalog, query]);

  if (!open) return null;

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="shortcuts-dialog-title"
      data-shortcuts-dialog
      className="fixed inset-0 z-[60] flex items-center justify-center bg-black/40 px-4 py-6"
      onClick={(e) => {
        // Backdrop click closes; clicks inside the panel stop bubbling.
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        ref={panelRef}
        tabIndex={-1}
        className="flex max-h-[85vh] w-full max-w-3xl flex-col rounded-lg border border-divider bg-surface shadow-2xl outline-none"
        onClick={(e) => e.stopPropagation()}
      >
        <header className="flex items-center justify-between gap-3 border-b border-divider px-5 py-3">
          <div>
            <h2 id="shortcuts-dialog-title" className="text-base font-semibold">
              Keyboard shortcuts
            </h2>
            <p className="text-xs text-secondary">{catalog.title}</p>
          </div>
          <button
            type="button"
            aria-label="Close"
            onClick={onClose}
            className="rounded p-1 text-secondary transition-colors hover:bg-surface-2 hover:text-default"
          >
            <X size={16} />
          </button>
        </header>

        <div className="border-b border-divider px-5 py-3">
          <label className="flex items-center gap-2 rounded-md border border-divider bg-surface-2 px-2 py-1.5">
            <Search size={14} className="text-secondary" />
            <input
              ref={inputRef}
              type="text"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="Search shortcuts…"
              className="w-full bg-transparent text-sm outline-none placeholder:text-secondary"
              data-shortcuts-search
            />
          </label>
        </div>

        <div className="flex-1 overflow-y-auto px-5 py-4">
          {filtered.length === 0 ? (
            <p className="py-8 text-center text-sm text-secondary">
              No shortcuts match &ldquo;{query}&rdquo;.
            </p>
          ) : (
            filtered.map((group) => (
              <section key={group.category} className="mb-5 last:mb-0">
                <h3 className="mb-2 text-xs font-semibold uppercase tracking-wide text-secondary">
                  {group.category}
                </h3>
                <ul className="flex flex-col gap-1">
                  {group.entries.map((entry) => (
                    <li
                      // Some labels are intentionally shared across
                      // multiple bindings (e.g. Backspace + Delete both
                      // map to "Delete selected shape(s)") — include the
                      // key combo in the React key so siblings stay
                      // unique.
                      key={`${group.category}-${entry.label}-${entry.keys.join("+")}`}
                      className={`flex items-center justify-between gap-3 rounded px-2 py-1.5 ${
                        entry.status === "planned" ? "text-secondary opacity-60" : "text-default"
                      }`}
                      data-shortcut-label={entry.label}
                    >
                      <span className="text-sm">
                        {entry.label}
                        {entry.status === "planned" ? (
                          <span className="ml-2 inline-flex items-center rounded bg-surface-2 px-1.5 py-0.5 text-[10px] uppercase tracking-wide text-secondary">
                            Soon
                          </span>
                        ) : null}
                      </span>
                      <KeyCombo keys={entry.keys} isMac={isMac} />
                    </li>
                  ))}
                </ul>
              </section>
            ))
          )}
        </div>

        <footer className="border-t border-divider px-5 py-2 text-xs text-secondary">
          Press <Kbd>Esc</Kbd> to close · Press <Kbd>{isMac ? "⌘" : "Ctrl"}</Kbd>
          <Kbd>/</Kbd> to open from anywhere
        </footer>
      </div>
    </div>
  );
}

interface FilteredGroup {
  readonly category: string;
  readonly entries: ReadonlyArray<ShortcutEntry>;
}

function filterCatalog(catalog: ShortcutCatalog, query: string): ReadonlyArray<FilteredGroup> {
  const q = query.trim().toLowerCase();
  const groups = new Map<string, ShortcutEntry[]>();
  for (const entry of catalog.entries) {
    if (q.length > 0) {
      const haystack = `${entry.label} ${entry.keys.join(" ")}`.toLowerCase();
      if (!haystack.includes(q)) continue;
    }
    let bucket = groups.get(entry.category);
    if (!bucket) {
      bucket = [];
      groups.set(entry.category, bucket);
    }
    bucket.push(entry);
  }
  return Array.from(groups.entries()).map(([category, entries]) => ({ category, entries }));
}

function KeyCombo({ keys, isMac }: { readonly keys: ReadonlyArray<ShortcutKey>; readonly isMac: boolean }) {
  return (
    <span className="flex items-center gap-1">
      {keys.map((key, i) => (
        <Kbd key={`${key}-${i}`}>{renderKey(key, isMac)}</Kbd>
      ))}
    </span>
  );
}

function Kbd({ children }: { readonly children: React.ReactNode }) {
  return (
    <kbd className="inline-flex min-w-[1.6rem] items-center justify-center rounded border border-divider bg-surface-2 px-1.5 py-0.5 font-mono text-[11px] text-default">
      {children}
    </kbd>
  );
}

function renderKey(key: ShortcutKey, isMac: boolean): string {
  switch (key) {
    case "Mod":
      return isMac ? "⌘" : "Ctrl";
    case "Shift":
      return isMac ? "⇧" : "Shift";
    case "Alt":
      return isMac ? "⌥" : "Alt";
    case "Enter":
      return "↵";
    case "Tab":
      return "Tab";
    case "Esc":
      return "Esc";
    case "Backspace":
      return "⌫";
    case "Delete":
      return "Del";
    case "Space":
      return "Space";
    case "ArrowUp":
      return "↑";
    case "ArrowDown":
      return "↓";
    case "ArrowLeft":
      return "←";
    case "ArrowRight":
      return "→";
    case "Home":
      return "Home";
    case "End":
      return "End";
    case "PageUp":
      return "PgUp";
    case "PageDown":
      return "PgDn";
    case "F2":
      return "F2";
    case "F4":
      return "F4";
    default:
      return key;
  }
}

function useIsMac(): boolean {
  const [isMac, setIsMac] = useState(false);
  useEffect(() => {
    if (typeof navigator === "undefined") return;
    // `navigator.platform` is deprecated in spec, but every browser
    // still ships it and it's the only synchronous source of truth.
    // `userAgentData.platform` is async + Chromium-only, so it's not
    // worth the extra round-trip for a presentational hint.
    const platform = navigator.platform || "";
    setIsMac(/Mac|iPod|iPhone|iPad/.test(platform));
  }, []);
  return isMac;
}
