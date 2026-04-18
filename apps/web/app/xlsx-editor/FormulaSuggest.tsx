"use client";

import type { ReactNode } from "react";
import { listRegisteredFunctions, type RegisteredFunctionInfo } from "@officeai/xlsx";

const ALL_FUNCTIONS: ReadonlyArray<RegisteredFunctionInfo> = listRegisteredFunctions();
const MAX_SUGGESTIONS = 8;

/**
 * Walk left from `caret` and return the partial function-name token
 * the caret is sitting on, plus its [from..to] span inside `draft`.
 *
 * Returns `null` when:
 *   - the draft isn't a formula (`=…`),
 *   - the prefix is empty (caret right after a delimiter), or
 *   - the prefix doesn't start with a letter (e.g. cell ref `B5`).
 *
 * Excel triggers autocomplete only when the prefix is preceded by
 * `=`, `(`, `,`, an arithmetic / comparison operator, `:`, `&`, or
 * whitespace.
 */
export function getActiveFunctionPrefix(
  draft: string,
  caret: number
): { prefix: string; from: number; to: number } | null {
  if (!draft.startsWith("=")) return null;
  if (caret > draft.length) caret = draft.length;
  let i = caret;
  while (i > 0 && /[A-Za-z0-9.]/.test(draft[i - 1])) i--;
  const prefix = draft.slice(i, caret);
  if (prefix.length === 0) return null;
  const before = i === 0 ? "=" : draft[i - 1];
  if (!/[=(,+\-*/^&<>%\s:]/.test(before)) return null;
  if (!/^[A-Za-z]/.test(prefix)) return null;
  return { prefix: prefix.toUpperCase(), from: i, to: caret };
}

export function getSuggestions(
  draft: string,
  caret: number
): {
  readonly matches: ReadonlyArray<RegisteredFunctionInfo>;
  readonly active: { prefix: string; from: number; to: number } | null;
} {
  const active = getActiveFunctionPrefix(draft, caret);
  if (!active) return { matches: [], active: null };
  const { prefix } = active;
  const matches = ALL_FUNCTIONS.filter((f) => f.name.startsWith(prefix)).slice(0, MAX_SUGGESTIONS);
  return { matches, active };
}

export interface FormulaSuggestProps {
  readonly matches: ReadonlyArray<RegisteredFunctionInfo>;
  readonly highlight: number;
  /** Mouse-pick: forwards to the parent's accept handler. */
  readonly onPick: (info: RegisteredFunctionInfo) => void;
  readonly onHighlight: (idx: number) => void;
}

/**
 * Excel-style function-name autocomplete popover. Presentational —
 * the parent (`XlsxEditor`) owns the matches, highlight index, and
 * accept logic so Tab / Enter / Up / Down can be wired through the
 * formula input's existing keydown handler.
 */
export function FormulaSuggest(props: FormulaSuggestProps): ReactNode {
  const { matches, highlight, onPick, onHighlight } = props;
  if (matches.length === 0) return null;
  return (
    <div
      data-testid="formula-suggest"
      role="listbox"
      aria-label="Formula suggestions"
      style={{
        position: "absolute",
        top: "100%",
        left: 0,
        marginTop: 4,
        minWidth: 220,
        maxHeight: 240,
        overflowY: "auto",
        background: "var(--surface)",
        border: "1px solid var(--divider)",
        borderRadius: 6,
        boxShadow: "0 4px 16px rgba(0,0,0,0.08)",
        zIndex: 50,
        fontSize: 12,
      }}
    >
      {matches.map((info, idx) => {
        const isActive = idx === highlight;
        return (
          <div
            key={info.name}
            data-testid={`formula-suggest-row-${info.name}`}
            role="option"
            aria-selected={isActive}
            onMouseDown={(e) => {
              // Prevent input blur so caret/selection survive the click.
              e.preventDefault();
              onPick(info);
            }}
            onMouseEnter={() => onHighlight(idx)}
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              padding: "4px 8px",
              cursor: "pointer",
              background: isActive ? "var(--ai-violet-light)" : "transparent",
              color: "var(--foreground)",
            }}
          >
            <span style={{ fontFamily: "var(--font-mono, monospace)" }}>{info.name}</span>
            <span style={{ color: "var(--secondary)", fontSize: 10 }}>{info.category}</span>
          </div>
        );
      })}
    </div>
  );
}

/**
 * Insert `NAME(` at `[span.from..span.to]` in `draft` and park the
 * caret inside the parens. Returned `caret` is the new offset.
 */
export function applySuggestion(
  draft: string,
  info: RegisteredFunctionInfo,
  span: { from: number; to: number }
): { next: string; caret: number } {
  const insertion = `${info.name}(`;
  const next = draft.slice(0, span.from) + insertion + draft.slice(span.to);
  const caret = span.from + insertion.length;
  return { next, caret };
}
