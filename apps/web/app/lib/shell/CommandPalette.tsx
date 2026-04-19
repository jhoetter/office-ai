"use client";

import { useEffect, useMemo, useRef, useState, type KeyboardEvent, type ReactNode } from "react";
import { cn } from "@officeai/ui";
import type { PaletteCommand } from "./types";

export interface CommandPaletteProps {
  readonly open: boolean;
  readonly onClose: () => void;
  readonly commands: ReadonlyArray<PaletteCommand>;
}

const RECENT_KEY = "officeai.palette.recent";
const RECENT_LIMIT = 6;

/**
 * Spotlight-style command palette opened with `Cmd+K`.
 *
 * Indexes the active product's typed commands plus a curated UI
 * list (Open / Save / Insert Image / …). Implements:
 *   - fuzzy substring + word-boundary scoring
 *   - arrow / enter / esc nav
 *   - "recent commands first" (localStorage, scoped to the host
 *     domain — no cross-product bleed because the IDs include the
 *     product prefix from the adapter side).
 */
export function CommandPalette({ open, onClose, commands }: CommandPaletteProps): ReactNode {
  const [query, setQuery] = useState("");
  const [active, setActive] = useState(0);
  const inputRef = useRef<HTMLInputElement | null>(null);
  const listRef = useRef<HTMLDivElement | null>(null);

  const recents = useMemo(() => loadRecents(), [open]);

  const filtered = useMemo(() => {
    if (!open) return [];
    const enabled = commands.filter((c) => c.enabled !== false);
    if (query.trim().length === 0) {
      const recentSet = new Set(recents);
      const recentItems = recents
        .map((id) => enabled.find((c) => c.id === id))
        .filter((c): c is PaletteCommand => Boolean(c));
      const others = enabled.filter((c) => !recentSet.has(c.id));
      return [...recentItems, ...others].slice(0, 60);
    }
    const q = query.toLowerCase();
    return enabled
      .map((c) => ({ c, score: scoreMatch(c, q) }))
      .filter((r) => r.score > 0)
      .sort((a, b) => b.score - a.score)
      .slice(0, 60)
      .map((r) => r.c);
  }, [open, query, commands, recents]);

  useEffect(() => {
    if (!open) return;
    setQuery("");
    setActive(0);
    requestAnimationFrame(() => inputRef.current?.focus());
  }, [open]);

  useEffect(() => {
    setActive(0);
  }, [query]);

  useEffect(() => {
    if (!open) return;
    const onKey = (e: globalThis.KeyboardEvent) => {
      if (e.key === "Escape") {
        e.preventDefault();
        onClose();
      }
    };
    document.addEventListener("keydown", onKey);
    return () => document.removeEventListener("keydown", onKey);
  }, [open, onClose]);

  if (!open) return null;

  const onInputKey = (e: KeyboardEvent<HTMLInputElement>) => {
    if (e.key === "ArrowDown") {
      e.preventDefault();
      setActive((i) => Math.min(filtered.length - 1, i + 1));
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      setActive((i) => Math.max(0, i - 1));
    } else if (e.key === "Enter") {
      e.preventDefault();
      const cmd = filtered[active];
      if (cmd) {
        rememberRecent(cmd.id);
        onClose();
        void cmd.run();
      }
    }
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label="Command palette"
      className="fixed inset-0 z-40 flex items-start justify-center bg-black/30 px-4 pt-[18vh]"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
      data-testid="command-palette"
    >
      <div className="w-full max-w-[560px] overflow-hidden rounded-lg border border-divider bg-background shadow-xl">
        <div className="border-b border-divider p-2">
          <input
            ref={inputRef}
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            onKeyDown={onInputKey}
            placeholder="Type a command…"
            className="h-8 w-full rounded-md bg-background px-2 text-sm text-foreground outline-none placeholder:text-tertiary"
            data-testid="command-palette-input"
            aria-label="Search commands"
          />
        </div>
        <div ref={listRef} className="max-h-[50vh] overflow-y-auto p-1">
          {filtered.length === 0 ? (
            <div className="px-3 py-6 text-center text-sm text-secondary">No commands match.</div>
          ) : (
            filtered.map((c, i) => (
              <button
                key={c.id}
                type="button"
                onClick={() => {
                  rememberRecent(c.id);
                  onClose();
                  void c.run();
                }}
                onMouseEnter={() => setActive(i)}
                className={cn(
                  "flex w-full items-center gap-2 rounded-md px-2 py-1.5 text-left text-sm",
                  active === i ? "bg-hover text-foreground" : "text-foreground hover:bg-hover"
                )}
                data-testid={`palette-cmd-${c.id}`}
              >
                <div className="flex min-w-0 flex-1 flex-col">
                  <span className="truncate">{c.label}</span>
                  {c.hint ? <span className="truncate text-xs text-secondary">{c.hint}</span> : null}
                </div>
                {c.shortcut ? <span className="text-xs tabular-nums text-tertiary">{c.shortcut}</span> : null}
              </button>
            ))
          )}
        </div>
      </div>
    </div>
  );
}

function scoreMatch(cmd: PaletteCommand, q: string): number {
  const label = cmd.label.toLowerCase();
  const id = cmd.id.toLowerCase();
  const hint = (cmd.hint ?? "").toLowerCase();
  if (label === q) return 1000;
  if (label.startsWith(q)) return 500;
  if (label.includes(` ${q}`)) return 250;
  if (label.includes(q)) return 100;
  if (id.includes(q)) return 50;
  if (hint.includes(q)) return 25;
  // Acronym match: first letters of words.
  const initials = label
    .split(/\s+/)
    .map((w) => w[0] ?? "")
    .join("");
  if (initials.startsWith(q)) return 75;
  return 0;
}

function loadRecents(): string[] {
  if (typeof window === "undefined") return [];
  try {
    const raw = window.localStorage.getItem(RECENT_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed.filter((x): x is string => typeof x === "string") : [];
  } catch {
    return [];
  }
}

function rememberRecent(id: string): void {
  if (typeof window === "undefined") return;
  try {
    const current = loadRecents();
    const next = [id, ...current.filter((x) => x !== id)].slice(0, RECENT_LIMIT);
    window.localStorage.setItem(RECENT_KEY, JSON.stringify(next));
  } catch {
    // localStorage may be unavailable (private mode); silently ignore.
  }
}
