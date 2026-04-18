/**
 * Forgiving formula tokenizer for *display* purposes only.
 *
 * Unlike `lex()`, this scanner never throws — it walks the string
 * left-to-right and emits coarse-grained spans suitable for syntax
 * highlighting and Excel-style coloured-reference UX:
 *
 *   - cell references (`A1`, `$A$1`, `Sheet2!A1`) → kind: "ref"
 *   - ranges (`A1:B2`, `Sheet2!A1:B2`) → kind: "range"
 *   - function names followed by `(` → kind: "function"
 *   - quoted strings → kind: "string"
 *   - numeric literals → kind: "number"
 *   - operators (`+ - * / ^ & = < > <= >= <>`) → kind: "operator"
 *   - braces / parens / commas / colons → kind: "punct"
 *   - error literals (`#REF!`, `#DIV/0!`, …) → kind: "error"
 *   - everything else (whitespace, partial input) → kind: "text"
 *
 * Ref / range tokens carry a `refKey` — a normalised form
 * (uppercased, `$` stripped) — so the caller can assign a stable
 * colour to "A1" / "$A$1" / "a1" alike. Sheet-qualified refs
 * include the sheet name in the key (`SHEET2!A1`).
 *
 * Used by:
 *   - the formula bar overlay (Phase 12a) to render colored ref text
 *   - the grid (Phase 12c) to draw matching coloured borders on the
 *     cells the formula points at
 *
 * Spec: this is presentational; no entry in `formula-engine.md`.
 * The strict `lex()` remains the single source of truth for the
 * evaluator.
 */
export type DisplayTokenKind =
  | "ref"
  | "range"
  | "function"
  | "string"
  | "number"
  | "operator"
  | "punct"
  | "error"
  | "text";

export interface DisplayToken {
  /** 0-based offset into the original source string. */
  readonly start: number;
  /** Exclusive end offset. */
  readonly end: number;
  /** Verbatim slice of the original source (`source.slice(start, end)`). */
  readonly text: string;
  readonly kind: DisplayTokenKind;
  /**
   * Normalised reference key (uppercase, `$`-stripped). Set only for
   * `ref` and `range` tokens. Sheet-qualified refs prepend the
   * uppercased sheet name + `!`.
   */
  readonly refKey?: string;
  /**
   * For `ref` tokens, the resolved 0-based `{ row, col }` of the
   * cell. For `range` tokens, the inclusive `{ r1, c1, r2, c2 }`
   * rectangle. Sheet name (if present) is in `sheet`.
   */
  readonly target?: RefTarget;
}

export type RefTarget =
  | { kind: "ref"; sheet?: string; row: number; col: number }
  | { kind: "range"; sheet?: string; r1: number; c1: number; r2: number; c2: number };

const REF_PATTERN = /^(?:'((?:[^']|'')+)'!|([A-Za-z_][A-Za-z0-9_.]*)!)?(\$?[A-Za-z]{1,3}\$?\d{1,7})(?::(\$?[A-Za-z]{1,3}\$?\d{1,7}))?/;
const NUMBER_PATTERN = /^\d+(?:\.\d+)?(?:[eE][+-]?\d+)?/;
const FUNCTION_PATTERN = /^([A-Za-z_][A-Za-z0-9_.]*)\s*\(/;
const NAME_PATTERN = /^[A-Za-z_][A-Za-z0-9_.]*/;
const ERROR_PATTERN = /^#(?:REF|NAME\?|VALUE|NUM|N\/A|DIV\/0|NULL)!/;
const TWO_CHAR_OPS = new Set(["<>", "<=", ">="]);
const ONE_CHAR_OPS = new Set(["+", "-", "*", "/", "^", "&", "=", "<", ">"]);
const PUNCT = new Set(["(", ")", "{", "}", ",", ";", ":", "%"]);

/**
 * Tokenise a formula source string for display.
 *
 * The leading `=` (if any) is emitted as a single `text` token so
 * downstream renderers can preserve it verbatim without worrying
 * about offsets.
 *
 * Never throws. Anything the scanner cannot classify is bucketed
 * into a contiguous `text` token so the output spans cover the
 * input exactly (`tokens[i].end === tokens[i+1].start`).
 */
export function tokenizeForDisplay(source: string): ReadonlyArray<DisplayToken> {
  if (source.length === 0) return [];
  const out: DisplayToken[] = [];
  let i = 0;
  if (source[0] === "=") {
    out.push({ start: 0, end: 1, text: "=", kind: "text" });
    i = 1;
  }

  let textStart = -1;
  const flushText = (end: number): void => {
    if (textStart >= 0 && end > textStart) {
      out.push({
        start: textStart,
        end,
        text: source.slice(textStart, end),
        kind: "text",
      });
    }
    textStart = -1;
  };
  const startText = (at: number): void => {
    if (textStart < 0) textStart = at;
  };

  while (i < source.length) {
    const rest = source.slice(i);
    const c = source[i]!;

    // Whitespace runs into the current text bucket so spacing is
    // preserved for the overlay renderer.
    if (c === " " || c === "\t" || c === "\n" || c === "\r") {
      startText(i);
      i++;
      continue;
    }

    // Quoted string. Scan until closing quote (with `""` escape).
    if (c === '"') {
      flushText(i);
      let j = i + 1;
      while (j < source.length) {
        if (source[j] === '"') {
          if (source[j + 1] === '"') {
            j += 2;
            continue;
          }
          j++;
          break;
        }
        j++;
      }
      out.push({ start: i, end: j, text: source.slice(i, j), kind: "string" });
      i = j;
      continue;
    }

    // Error literal.
    const errMatch = ERROR_PATTERN.exec(rest);
    if (errMatch) {
      flushText(i);
      const end = i + errMatch[0].length;
      out.push({ start: i, end, text: errMatch[0], kind: "error" });
      i = end;
      continue;
    }

    // Reference or range (incl. sheet-qualified).
    const refMatch = REF_PATTERN.exec(rest);
    if (refMatch && (/[A-Za-z]/.test(c) || c === "'" || c === "$")) {
      const consumed = refMatch[0];
      const sheet = (refMatch[1] ?? refMatch[2])?.replace(/''/g, "'");
      const a1Start = refMatch[3]!;
      const a1End = refMatch[4];
      const target = a1End ? buildRangeTarget(sheet, a1Start, a1End) : buildRefTarget(sheet, a1Start);
      // If the address is malformed (e.g. column out of bounds for
      // our 16384 col / 1048576 row clamp) treat the slice as text.
      if (target) {
        flushText(i);
        const end = i + consumed.length;
        const refKey = computeRefKey(sheet, a1Start, a1End);
        out.push({
          start: i,
          end,
          text: consumed,
          kind: a1End ? "range" : "ref",
          refKey,
          target,
        });
        i = end;
        continue;
      }
    }

    // Function name followed by `(` — peek for the paren.
    const fnMatch = FUNCTION_PATTERN.exec(rest);
    if (fnMatch) {
      flushText(i);
      const end = i + fnMatch[1]!.length;
      out.push({ start: i, end, text: fnMatch[1]!, kind: "function" });
      i = end;
      continue;
    }

    // Number literal.
    if (c === "." || (c >= "0" && c <= "9")) {
      const numMatch = NUMBER_PATTERN.exec(rest);
      if (numMatch) {
        flushText(i);
        const end = i + numMatch[0].length;
        out.push({ start: i, end, text: numMatch[0], kind: "number" });
        i = end;
        continue;
      }
    }

    // Two-char operator.
    const two = rest.slice(0, 2);
    if (TWO_CHAR_OPS.has(two)) {
      flushText(i);
      out.push({ start: i, end: i + 2, text: two, kind: "operator" });
      i += 2;
      continue;
    }

    // One-char operator.
    if (ONE_CHAR_OPS.has(c)) {
      flushText(i);
      out.push({ start: i, end: i + 1, text: c, kind: "operator" });
      i++;
      continue;
    }

    // Punctuation.
    if (PUNCT.has(c)) {
      flushText(i);
      out.push({ start: i, end: i + 1, text: c, kind: "punct" });
      i++;
      continue;
    }

    // A bare name (TRUE, FALSE, defined name) — emit as text.
    const nm = NAME_PATTERN.exec(rest);
    if (nm) {
      flushText(i);
      const end = i + nm[0].length;
      out.push({ start: i, end, text: nm[0], kind: "text" });
      i = end;
      continue;
    }

    // Anything else → bucket into text.
    startText(i);
    i++;
  }
  flushText(source.length);
  return out;
}

/**
 * Default cycling palette used to colour distinct references in the
 * formula bar / grid. Picked to be distinguishable on both light
 * and dark themes; tweak via the optional argument to
 * `assignRefColors`.
 *
 * Order chosen to match Excel's "trace precedents" colour-cycle
 * intuition (cool → warm → green → magenta) so two adjacent refs
 * never collide.
 */
export const DEFAULT_REF_COLORS: ReadonlyArray<string> = [
  "#1f77b4", //  blue
  "#d62728", //  red
  "#2ca02c", //  green
  "#9467bd", //  violet
  "#ff7f0e", //  orange
  "#17becf", //  teal
  "#e377c2", //  pink
  "#8c564b", //  brown
];

/**
 * Walk a token list and assign a stable colour per unique `refKey`.
 * The same key (e.g. `A1` referenced twice in `=A1+A1`) gets the
 * same colour; distinct keys cycle through the palette.
 */
export function assignRefColors(
  tokens: ReadonlyArray<DisplayToken>,
  palette: ReadonlyArray<string> = DEFAULT_REF_COLORS
): ReadonlyMap<string, string> {
  const out = new Map<string, string>();
  let next = 0;
  for (const t of tokens) {
    if (!t.refKey) continue;
    if (out.has(t.refKey)) continue;
    out.set(t.refKey, palette[next % palette.length]!);
    next++;
  }
  return out;
}

function buildRefTarget(sheet: string | undefined, a1: string): RefTarget | null {
  const addr = parseA1Address(a1);
  if (!addr) return null;
  return { kind: "ref", sheet, row: addr.row, col: addr.col };
}

function buildRangeTarget(sheet: string | undefined, a1Start: string, a1End: string): RefTarget | null {
  const a = parseA1Address(a1Start);
  const b = parseA1Address(a1End);
  if (!a || !b) return null;
  return {
    kind: "range",
    sheet,
    r1: Math.min(a.row, b.row),
    c1: Math.min(a.col, b.col),
    r2: Math.max(a.row, b.row),
    c2: Math.max(a.col, b.col),
  };
}

function computeRefKey(sheet: string | undefined, a1Start: string, a1End: string | undefined): string {
  const start = a1Start.replace(/\$/g, "").toUpperCase();
  const end = a1End ? a1End.replace(/\$/g, "").toUpperCase() : undefined;
  const head = sheet ? `${sheet.toUpperCase()}!` : "";
  return end ? `${head}${start}:${end}` : `${head}${start}`;
}

function parseA1Address(a1: string): { row: number; col: number } | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(a1);
  if (!m) return null;
  const col = colLettersToIndex(m[1]!);
  const row = Number(m[2]) - 1;
  if (col < 0 || col >= 16384 || row < 0 || row >= 1048576) return null;
  return { row, col };
}

function colLettersToIndex(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n - 1;
}
