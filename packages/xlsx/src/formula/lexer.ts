import { FormulaParseError } from "./ast.js";
import { parseErrorLiteral } from "./errors.js";
import type { Token, TokenType } from "./tokens.js";

/**
 * Single forward-scan lexer.
 *
 * Spec: `spec/xlsx/formula-engine.md` §10. The lexer accepts an
 * optional leading `=` and emits `EOF` at end. References are
 * recognised by the regex catalogue in §10.1; the parser turns the
 * lexeme into a typed `CellRef`/`RangeRef` via `references.ts`.
 */
export function lex(source: string): ReadonlyArray<Token> {
  const text = source.startsWith("=") ? source.slice(1) : source;
  const offset = source.startsWith("=") ? 1 : 0;
  const out: Token[] = [];
  let i = 0;
  const len = text.length;

  while (i < len) {
    const c = text[i];

    if (c === " " || c === "\t" || c === "\n" || c === "\r") {
      i++;
      continue;
    }

    const start = i + offset;

    if (c === "(") {
      out.push(tok("LPAREN", "(", start, start + 1));
      i++;
      continue;
    }
    if (c === ")") {
      out.push(tok("RPAREN", ")", start, start + 1));
      i++;
      continue;
    }
    if (c === "{") {
      out.push(tok("LBRACE", "{", start, start + 1));
      i++;
      continue;
    }
    if (c === "}") {
      out.push(tok("RBRACE", "}", start, start + 1));
      i++;
      continue;
    }
    if (c === ",") {
      out.push(tok("COMMA", ",", start, start + 1));
      i++;
      continue;
    }
    if (c === ":") {
      out.push(tok("COLON", ":", start, start + 1));
      i++;
      continue;
    }
    if (c === ";") {
      out.push(tok("SEMICOLON", ";", start, start + 1));
      i++;
      continue;
    }
    if (c === "%") {
      out.push(tok("PERCENT", "%", start, start + 1));
      i++;
      continue;
    }

    // Two-char comparison operators
    if (c === "<") {
      const next = text[i + 1];
      if (next === ">") {
        out.push(tokOp("<>", start, start + 2));
        i += 2;
        continue;
      }
      if (next === "=") {
        out.push(tokOp("<=", start, start + 2));
        i += 2;
        continue;
      }
      out.push(tokOp("<", start, start + 1));
      i++;
      continue;
    }
    if (c === ">") {
      const next = text[i + 1];
      if (next === "=") {
        out.push(tokOp(">=", start, start + 2));
        i += 2;
        continue;
      }
      out.push(tokOp(">", start, start + 1));
      i++;
      continue;
    }

    if (c === "=" || c === "+" || c === "-" || c === "*" || c === "/" || c === "^" || c === "&") {
      out.push(tokOp(c, start, start + 1));
      i++;
      continue;
    }

    if (c === '"') {
      // string literal — doubled "" is a single literal "
      let j = i + 1;
      let buf = "";
      let terminated = false;
      while (j < len) {
        if (text[j] === '"') {
          if (text[j + 1] === '"') {
            buf += '"';
            j += 2;
            continue;
          }
          terminated = true;
          j++;
          break;
        }
        buf += text[j];
        j++;
      }
      if (!terminated) {
        throw new FormulaParseError(
          "unterminated-string",
          `Unterminated string literal starting at ${start}`,
          { start, end: j + offset }
        );
      }
      out.push({
        type: "STRING",
        text: text.slice(i, j),
        start,
        end: j + offset,
        value: buf,
      });
      i = j;
      continue;
    }

    if (c === "#") {
      // error literal
      let j = i + 1;
      // include the longest known error literal char set
      while (j < len && /[A-Z0-9_/!?]/.test(text[j])) j++;
      const lit = text.slice(i, j);
      const err = parseErrorLiteral(lit);
      if (!err) {
        throw new FormulaParseError("unexpected-token", `Unknown error literal "${lit}"`, {
          start,
          end: j + offset,
        });
      }
      out.push({
        type: "ERROR",
        text: lit,
        start,
        end: j + offset,
        value: err,
      });
      i = j;
      continue;
    }

    // `$`-prefixed cell or range, e.g. `$A$1` or `$A$1:$B$5`.
    if (c === "$") {
      const refMatch = readRefCandidate(text, i);
      if (refMatch) {
        const isRange = refMatch.text.includes(":");
        out.push({
          type: isRange ? "RANGE_REF" : "REF",
          text: refMatch.text,
          start,
          end: start + refMatch.text.length,
        });
        i += refMatch.text.length;
        continue;
      }
      throw new FormulaParseError("unexpected-token", `Unexpected "$" at ${start}`, {
        start,
        end: start + 1,
      });
    }

    if (isDigit(c)) {
      // Whole-row range like `3:5` takes precedence over a NUMBER literal.
      const rowRange = ROW_RANGE.exec(text.slice(i));
      if (rowRange) {
        out.push({
          type: "RANGE_REF",
          text: rowRange[0],
          start,
          end: start + rowRange[0].length,
        });
        i += rowRange[0].length;
        continue;
      }
    }

    if (isDigit(c) || (c === "." && isDigit(text[i + 1]))) {
      const m = NUMBER_RE.exec(text.slice(i));
      if (!m) {
        throw new FormulaParseError("invalid-number", `Invalid number at ${start}`, {
          start,
          end: start + 1,
        });
      }
      const lit = m[0];
      const n = Number(lit);
      if (!Number.isFinite(n)) {
        throw new FormulaParseError("invalid-number", `Invalid number "${lit}"`, {
          start,
          end: start + lit.length,
        });
      }
      out.push({
        type: "NUMBER",
        text: lit,
        start,
        end: start + lit.length,
        value: n,
      });
      i += lit.length;
      continue;
    }

    if (c === "'" || isIdentStart(c)) {
      const refMatch = readRefCandidate(text, i);
      if (refMatch) {
        const isRange = refMatch.text.includes(":");
        out.push({
          type: isRange ? "RANGE_REF" : "REF",
          text: refMatch.text,
          start,
          end: start + refMatch.text.length,
        });
        i += refMatch.text.length;
        continue;
      }

      // identifier (function or NAME or BOOL)
      const idMatch = IDENT_RE.exec(text.slice(i));
      if (idMatch) {
        const id = idMatch[0];
        const upper = id.toUpperCase();
        const next = text[i + id.length];
        if (next === "(") {
          out.push({
            type: "FUNCTION",
            text: id,
            start,
            end: start + id.length,
            value: upper,
          });
          i += id.length;
          continue;
        }
        if (upper === "TRUE" || upper === "FALSE") {
          out.push({
            type: "BOOL",
            text: id,
            start,
            end: start + id.length,
            value: upper === "TRUE",
          });
          i += id.length;
          continue;
        }
        out.push({
          type: "NAME",
          text: id,
          start,
          end: start + id.length,
          value: id,
        });
        i += id.length;
        continue;
      }
    }

    throw new FormulaParseError("unexpected-token", `Unexpected character "${c}" at ${start}`, {
      start,
      end: start + 1,
    });
  }

  out.push({ type: "EOF", text: "", start: source.length, end: source.length });
  return out;
}

function tok(type: TokenType, lit: string, start: number, end: number): Token {
  return { type, text: lit, start, end };
}

function tokOp(op: string, start: number, end: number): Token {
  return { type: "OPERATOR", text: op, start, end, value: op };
}

function isDigit(c: string | undefined): boolean {
  return c !== undefined && c >= "0" && c <= "9";
}

function isIdentStart(c: string): boolean {
  return (c >= "A" && c <= "Z") || (c >= "a" && c <= "z") || c === "_";
}

const NUMBER_RE = /^(?:\d+\.\d*|\.\d+|\d+)(?:[eE][+-]?\d+)?/;
const IDENT_RE = /^[A-Za-z_][A-Za-z0-9_.]*/;

// Reference recognition (§10.1) — quoted-sheet | bare-sheet | A1 | range | whole-col | whole-row
const SHEET_QUOTED = /^'(?:[^']|'')*'!/;
const SHEET_BARE = /^[A-Za-z_][A-Za-z0-9_.]*!/;
const CELL = /^\$?[A-Z]{1,3}\$?\d+/i;
const CELL_RANGE = /^\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+/i;
const COL_RANGE = /^\$?[A-Z]{1,3}:\$?[A-Z]{1,3}/i;
const ROW_RANGE = /^\$?\d+:\$?\d+/i;

function readRefCandidate(text: string, start: number): { text: string } | undefined {
  const slice = text.slice(start);
  let prefixLen = 0;

  const q = SHEET_QUOTED.exec(slice);
  if (q) {
    prefixLen = q[0].length;
  } else {
    const b = SHEET_BARE.exec(slice);
    if (b) {
      // disambiguate: don't swallow function-call identifier that just happens to look like SheetName!
      // BARE!IDENT(   ← here BARE!IDENT is a sheet-prefixed ref only if IDENT starts with $? letter+digit.
      // We just attempt the body match; if the body doesn't parse as cell/range, we *don't* take the prefix.
      const bodySlice = slice.slice(b[0].length);
      if (
        CELL_RANGE.test(bodySlice) ||
        CELL.test(bodySlice) ||
        COL_RANGE.test(bodySlice) ||
        ROW_RANGE.test(bodySlice)
      ) {
        prefixLen = b[0].length;
      }
    }
  }

  const bodySlice = slice.slice(prefixLen);
  const bodyMatch: RegExpExecArray | null =
    CELL_RANGE.exec(bodySlice) ??
    COL_RANGE.exec(bodySlice) ??
    ROW_RANGE.exec(bodySlice) ??
    CELL.exec(bodySlice);

  if (!bodyMatch) return undefined;
  const total = prefixLen + bodyMatch[0].length;

  // Reject if the next character continues an identifier — that means
  // we matched a partial token (e.g. "A1FOO" should not be a ref).
  const nextCh = slice[total];
  if (nextCh && /[A-Za-z0-9_.]/.test(nextCh)) {
    // exception: if the bare match is just CELL and next is `(`, it's not a ref either — but
    // CELL won't match identifiers, so this only triggers for things like "A1B" which isn't a valid ref.
    if (prefixLen === 0) return undefined;
    return undefined;
  }

  return { text: slice.slice(0, total) };
}
