"use client";

import type { CSSProperties, ReactNode } from "react";
import type { DisplayToken } from "@officeai/xlsx";

interface FormulaHighlightProps {
  /**
   * The string currently visible in the formula bar `<input>` —
   * usually `formulaDraft` while focused, otherwise the resolved
   * cell formula.
   */
  readonly value: string;
  /**
   * Tokenised view of `value`, produced by `tokenizeForDisplay` in
   * the parent. Passed in so the parent can also feed the same
   * tokens to the grid for matching ref borders.
   */
  readonly tokens: ReadonlyArray<DisplayToken>;
  /**
   * Map from `token.refKey` to the colour assigned by
   * `assignRefColors`. Keys missing from the map render in the
   * default text colour.
   */
  readonly refColors: ReadonlyMap<string, string>;
  /**
   * Horizontal scroll offset of the underlying input, so the
   * overlay tracks long formulas. Caller wires `onScroll` on the
   * input to keep this in sync.
   */
  readonly scrollLeft: number;
}

/**
 * Read-only overlay that paints colored ref tokens behind a
 * transparent-text formula bar input. The input remains the
 * focusable surface (caret, selection, IME), so this layer is
 * `pointer-events: none` and aria-hidden.
 *
 * Render contract:
 *   - The overlay must use the **same font / line-height / padding**
 *     as the underlying input, so each character lines up under the
 *     transparent caret.
 *   - Spans cover `value` exactly (`tokens` from
 *     `tokenizeForDisplay` already enforce this) — character offsets
 *     stay 1:1 between layers.
 */
export function FormulaHighlight({ value, tokens, refColors, scrollLeft }: FormulaHighlightProps): ReactNode {
  const style: CSSProperties = {
    position: "absolute",
    inset: 0,
    overflow: "hidden",
    pointerEvents: "none",
    whiteSpace: "pre",
    fontFamily: "inherit",
    fontSize: "inherit",
    lineHeight: "inherit",
    padding: "0.25rem",
    color: "var(--foreground)",
    transform: `translateX(${-scrollLeft}px)`,
  };

  return (
    <div aria-hidden style={style} data-testid="formula-highlight">
      {tokens.length === 0
        ? // Render a non-breaking space so the overlay still has the
          // same baseline as the input even when the value is empty.
          "\u00a0"
        : tokens.map((t, i) => renderToken(t, i, refColors, value))}
    </div>
  );
}

function renderToken(
  t: DisplayToken,
  i: number,
  refColors: ReadonlyMap<string, string>,
  value: string
): ReactNode {
  const slice = value.slice(t.start, t.end);
  const color = t.refKey ? refColors.get(t.refKey) : undefined;
  const baseStyle: CSSProperties = {
    fontWeight: t.kind === "function" ? 600 : undefined,
    color: color ?? colorForKind(t.kind),
  };
  return (
    <span key={i} style={baseStyle} data-token-kind={t.kind}>
      {slice}
    </span>
  );
}

function colorForKind(kind: DisplayToken["kind"]): string | undefined {
  switch (kind) {
    case "function":
      return "var(--ai-violet)";
    case "string":
      return "var(--success)";
    case "number":
      return "var(--info)";
    case "operator":
    case "punct":
      return "var(--secondary)";
    case "error":
      return "var(--error)";
    case "ref":
    case "range":
    case "text":
    default:
      return undefined;
  }
}
