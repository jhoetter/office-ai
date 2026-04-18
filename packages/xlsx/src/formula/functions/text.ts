import { ErrorKinds } from "../errors.js";
import { arity, type MutableFunctionRegistry } from "../function-registry.js";
import {
  bool,
  err,
  isError,
  num,
  str,
  toNumber,
  toString,
  type ErrorValue,
  type StringValue,
  type Value,
} from "../values.js";

/**
 * Text-category functions (P0). Spec: `spec/xlsx/formula-engine.md`
 * §16.4.
 *
 * Each impl coerces inputs through `toString` / `toNumber`, short-
 * circuits on errors, and never throws. `CONCATENATE` rejects ranges
 * (Excel parity); `CONCAT` flattens them row-major. `TEXT` is a P0
 * minimal implementation — see the inline comment on the registration
 * for the supported format hints.
 */
export function registerText(reg: MutableFunctionRegistry): void {
  reg.register({
    name: "CONCATENATE",
    arity: arity(1, 255),
    fn: (args) => {
      let out = "";
      for (const a of args) {
        if (isError(a)) return a;
        if (a.kind === "r") {
          // Excel `CONCATENATE` is scalar-only; collapse a 1×1 range
          // to its scalar, otherwise reject. (`CONCAT` is the modern
          // range-aware variant.)
          if (a.v.length === 1 && a.v[0].length === 1) {
            const s = toString(a.v[0][0]);
            if (isError(s)) return s;
            out += s.v;
            continue;
          }
          return err(ErrorKinds.VALUE);
        }
        const s = toString(a);
        if (isError(s)) return s;
        out += s.v;
      }
      return str(out);
    },
  });

  reg.register({
    name: "CONCAT",
    arity: arity(1, 255),
    fn: (args) => {
      let out = "";
      for (const a of args) {
        const r = appendFlattened(a, (s) => {
          out += s;
        });
        if (r) return r;
      }
      return str(out);
    },
  });

  reg.register({
    name: "TEXTJOIN",
    arity: arity(3, 254),
    fn: (args) => {
      const delimVal = toString(args[0]);
      if (isError(delimVal)) return delimVal;
      const ignoreEmpty = toBooleanLoose(args[1]);
      if (isError(ignoreEmpty)) return ignoreEmpty;
      const parts: string[] = [];
      for (let i = 2; i < args.length; i++) {
        const r = appendFlattened(args[i], (s) => {
          if (ignoreEmpty.v && s === "") return;
          parts.push(s);
        });
        if (r) return r;
      }
      const joined = parts.join(delimVal.v);
      if (joined.length > MAX_TEXT_LEN) return err(ErrorKinds.VALUE);
      return str(joined);
    },
  });

  reg.register({
    name: "LEFT",
    arity: arity(1, 2),
    fn: (args) => {
      const s = toString(args[0]);
      if (isError(s)) return s;
      const n = args.length === 2 ? toIntFloor(args[1]) : { kind: "n" as const, v: 1 };
      if (isError(n)) return n;
      if (n.v < 0) return err(ErrorKinds.VALUE);
      return str(s.v.slice(0, n.v));
    },
  });

  reg.register({
    name: "RIGHT",
    arity: arity(1, 2),
    fn: (args) => {
      const s = toString(args[0]);
      if (isError(s)) return s;
      const n = args.length === 2 ? toIntFloor(args[1]) : { kind: "n" as const, v: 1 };
      if (isError(n)) return n;
      if (n.v < 0) return err(ErrorKinds.VALUE);
      if (n.v === 0) return str("");
      if (n.v >= s.v.length) return str(s.v);
      return str(s.v.slice(s.v.length - n.v));
    },
  });

  reg.register({
    name: "MID",
    arity: arity(3, 3),
    fn: (args) => {
      const s = toString(args[0]);
      if (isError(s)) return s;
      const start = toIntFloor(args[1]);
      if (isError(start)) return start;
      const len = toIntFloor(args[2]);
      if (isError(len)) return len;
      if (start.v < 1) return err(ErrorKinds.VALUE);
      if (len.v < 0) return err(ErrorKinds.VALUE);
      if (start.v > s.v.length) return str("");
      return str(s.v.slice(start.v - 1, start.v - 1 + len.v));
    },
  });

  reg.register({
    name: "LEN",
    arity: arity(1, 1),
    fn: (args) => {
      const s = toString(args[0]);
      if (isError(s)) return s;
      return num(s.v.length);
    },
  });

  reg.register({
    name: "TRIM",
    arity: arity(1, 1),
    fn: (args) => {
      const s = toString(args[0]);
      if (isError(s)) return s;
      // Excel `TRIM` collapses internal runs of ASCII spaces only and
      // strips leading/trailing whitespace. We treat any whitespace as
      // a separator (close enough for P0).
      const trimmed = s.v.trim().replace(/\s+/g, " ");
      return str(trimmed);
    },
  });

  reg.register({
    name: "UPPER",
    arity: arity(1, 1),
    fn: (args) => {
      const s = toString(args[0]);
      if (isError(s)) return s;
      return str(s.v.toUpperCase());
    },
  });

  reg.register({
    name: "LOWER",
    arity: arity(1, 1),
    fn: (args) => {
      const s = toString(args[0]);
      if (isError(s)) return s;
      return str(s.v.toLowerCase());
    },
  });

  reg.register({
    name: "PROPER",
    arity: arity(1, 1),
    fn: (args) => {
      const s = toString(args[0]);
      if (isError(s)) return s;
      return str(toProperCase(s.v));
    },
  });

  reg.register({
    name: "FIND",
    arity: arity(2, 3),
    fn: (args) => {
      const needle = toString(args[0]);
      if (isError(needle)) return needle;
      const haystack = toString(args[1]);
      if (isError(haystack)) return haystack;
      const start = args.length === 3 ? toIntFloor(args[2]) : { kind: "n" as const, v: 1 };
      if (isError(start)) return start;
      if (start.v < 1) return err(ErrorKinds.VALUE);
      if (start.v > haystack.v.length && needle.v !== "") return err(ErrorKinds.VALUE);
      const idx = haystack.v.indexOf(needle.v, start.v - 1);
      if (idx === -1) return err(ErrorKinds.VALUE);
      return num(idx + 1);
    },
  });

  reg.register({
    name: "SEARCH",
    arity: arity(2, 3),
    fn: (args) => {
      const needle = toString(args[0]);
      if (isError(needle)) return needle;
      const haystack = toString(args[1]);
      if (isError(haystack)) return haystack;
      const start = args.length === 3 ? toIntFloor(args[2]) : { kind: "n" as const, v: 1 };
      if (isError(start)) return start;
      if (start.v < 1) return err(ErrorKinds.VALUE);
      const re = wildcardToRegex(needle.v);
      const idx = haystack.v
        .toLowerCase()
        .slice(start.v - 1)
        .search(re);
      if (idx === -1) return err(ErrorKinds.VALUE);
      return num(idx + start.v);
    },
  });

  reg.register({
    name: "SUBSTITUTE",
    arity: arity(3, 4),
    fn: (args) => {
      const text = toString(args[0]);
      if (isError(text)) return text;
      const oldText = toString(args[1]);
      if (isError(oldText)) return oldText;
      const newText = toString(args[2]);
      if (isError(newText)) return newText;
      if (oldText.v === "") return text;
      if (args.length === 4) {
        const nth = toIntFloor(args[3]);
        if (isError(nth)) return nth;
        if (nth.v < 1) return err(ErrorKinds.VALUE);
        return str(substituteNth(text.v, oldText.v, newText.v, nth.v));
      }
      return str(text.v.split(oldText.v).join(newText.v));
    },
  });

  reg.register({
    name: "REPLACE",
    arity: arity(4, 4),
    fn: (args) => {
      const text = toString(args[0]);
      if (isError(text)) return text;
      const start = toIntFloor(args[1]);
      if (isError(start)) return start;
      const length = toIntFloor(args[2]);
      if (isError(length)) return length;
      const newText = toString(args[3]);
      if (isError(newText)) return newText;
      if (start.v < 1) return err(ErrorKinds.VALUE);
      if (length.v < 0) return err(ErrorKinds.VALUE);
      const i = start.v - 1;
      const before = text.v.slice(0, i);
      const after = text.v.slice(i + length.v);
      return str(before + newText.v + after);
    },
  });

  reg.register({
    name: "REPT",
    arity: arity(2, 2),
    fn: (args) => {
      const s = toString(args[0]);
      if (isError(s)) return s;
      const n = toIntFloor(args[1]);
      if (isError(n)) return n;
      if (n.v < 0) return err(ErrorKinds.VALUE);
      if (n.v === 0 || s.v === "") return str("");
      const total = s.v.length * n.v;
      if (total > MAX_TEXT_LEN) return err(ErrorKinds.VALUE);
      return str(s.v.repeat(n.v));
    },
  });

  reg.register({
    name: "TEXT",
    arity: arity(2, 2),
    // P0 minimal: we recognise a small set of common Excel format
    // hints (fixed decimals, thousands separators, percent, dollar).
    // Anything else falls through to the default `toString` rendering
    // — never an error. The full number-format engine is §16.7.
    fn: (args) => {
      const fmt = toString(args[1]);
      if (isError(fmt)) return fmt;
      if (isError(args[0])) return args[0];
      const numv = toNumber(args[0]);
      if (isError(numv)) {
        const s = toString(args[0]);
        if (isError(s)) return s;
        return s;
      }
      return str(formatNumberMinimal(numv.v, fmt.v, args[0]));
    },
  });

  reg.register({
    name: "VALUE",
    arity: arity(1, 1),
    fn: (args) => {
      if (isError(args[0])) return args[0];
      const s = toString(args[0]);
      if (isError(s)) return s;
      const trimmed = s.v.trim();
      if (trimmed === "") return num(0);
      const n = Number(trimmed);
      if (!Number.isFinite(n)) return err(ErrorKinds.VALUE);
      return num(n);
    },
  });

  reg.register({
    name: "NUMBERVALUE",
    arity: arity(1, 3),
    fn: (args) => {
      if (isError(args[0])) return args[0];
      const s = toString(args[0]);
      if (isError(s)) return s;
      const decSep = args.length >= 2 ? toString(args[1]) : str(".");
      if (isError(decSep)) return decSep;
      const groupSep = args.length >= 3 ? toString(args[2]) : str(",");
      if (isError(groupSep)) return groupSep;
      // Single-character separators only — Excel parity.
      if (decSep.v.length < 1 || groupSep.v.length < 1) return err(ErrorKinds.VALUE);
      const dec = decSep.v.charAt(0);
      const grp = groupSep.v.charAt(0);
      let cleaned = "";
      for (const ch of s.v) {
        if (ch === grp) continue;
        if (ch === dec) {
          cleaned += ".";
          continue;
        }
        cleaned += ch;
      }
      const trimmed = cleaned.trim();
      if (trimmed === "") return err(ErrorKinds.VALUE);
      const n = Number(trimmed);
      if (!Number.isFinite(n)) return err(ErrorKinds.VALUE);
      return num(n);
    },
  });

  reg.register({
    name: "CHAR",
    arity: arity(1, 1),
    fn: (args) => {
      const n = toIntFloor(args[0]);
      if (isError(n)) return n;
      if (n.v < 1 || n.v > 255) return err(ErrorKinds.VALUE);
      return str(String.fromCharCode(n.v));
    },
  });

  reg.register({
    name: "CODE",
    arity: arity(1, 1),
    fn: (args) => {
      const s = toString(args[0]);
      if (isError(s)) return s;
      if (s.v.length === 0) return err(ErrorKinds.VALUE);
      return num(s.v.charCodeAt(0));
    },
  });

  reg.register({
    name: "EXACT",
    arity: arity(2, 2),
    fn: (args) => {
      const a = toString(args[0]);
      if (isError(a)) return a;
      const b = toString(args[1]);
      if (isError(b)) return b;
      return bool(a.v === b.v);
    },
  });

  reg.register({
    name: "T",
    arity: arity(1, 1),
    fn: (args) => {
      const v = args[0];
      switch (v.kind) {
        case "s":
          return v;
        case "e":
          return v;
        case "n":
        case "b":
          return str("");
        case "r": {
          if (v.v.length === 1 && v.v[0].length === 1) {
            const inner = v.v[0][0];
            switch (inner.kind) {
              case "s":
                return inner;
              case "e":
                return inner;
              case "n":
              case "b":
              case "r":
                return str("");
            }
          }
          return str("");
        }
      }
    },
  });
}

// ── Helpers ────────────────────────────────────────────────────────────────

const MAX_TEXT_LEN = 32_767;

/**
 * Append all scalar string forms of `v` to the sink. Returns an
 * `ErrorValue` if any element fails coercion (so the caller can
 * propagate).
 */
function appendFlattened(v: Value, push: (s: string) => void): ErrorValue | undefined {
  if (isError(v)) return v;
  if (v.kind === "r") {
    for (const row of v.v) {
      for (const cell of row) {
        if (isError(cell)) return cell;
        const s = toString(cell);
        if (isError(s)) return s;
        push(s.v);
      }
    }
    return undefined;
  }
  const s = toString(v);
  if (isError(s)) return s;
  push(s.v);
  return undefined;
}

function toIntFloor(v: Value): { kind: "n"; v: number } | ErrorValue {
  const n = toNumber(v);
  if (isError(n)) return n;
  return { kind: "n", v: Math.trunc(n.v) };
}

/**
 * Loose boolean coercion mirroring `coerceCondition` in `logic.ts`:
 * numbers ≠ 0 are truthy, blanks/0 are falsy, strings go through
 * `toBoolean` (only "TRUE" / "FALSE").
 */
function toBooleanLoose(v: Value): { kind: "b"; v: boolean } | ErrorValue {
  switch (v.kind) {
    case "b":
      return v;
    case "n":
      return { kind: "b", v: v.v !== 0 };
    case "s": {
      const upper = v.v.toUpperCase();
      if (upper === "TRUE") return { kind: "b", v: true };
      if (upper === "FALSE") return { kind: "b", v: false };
      return err(ErrorKinds.VALUE);
    }
    case "e":
      return v;
    case "r": {
      if (v.v.length === 1 && v.v[0].length === 1) return toBooleanLoose(v.v[0][0]);
      return err(ErrorKinds.VALUE);
    }
  }
}

function toProperCase(s: string): string {
  let out = "";
  let prevAlpha = false;
  for (const ch of s) {
    const isAlpha = /[A-Za-z]/.test(ch);
    if (isAlpha) {
      out += prevAlpha ? ch.toLowerCase() : ch.toUpperCase();
    } else {
      out += ch;
    }
    prevAlpha = isAlpha;
  }
  return out;
}

/**
 * Translate an Excel wildcard pattern into a case-insensitive RegExp.
 * `?` matches one char, `*` matches any run, `~?` / `~*` / `~~`
 * escape the literal. Other regex meta-chars are escaped.
 */
function wildcardToRegex(pattern: string): RegExp {
  let out = "";
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === "~" && i + 1 < pattern.length) {
      const next = pattern[i + 1];
      if (next === "?" || next === "*" || next === "~") {
        out += escapeRegex(next);
        i++;
        continue;
      }
      out += escapeRegex(ch);
      continue;
    }
    if (ch === "?") {
      out += ".";
      continue;
    }
    if (ch === "*") {
      out += ".*";
      continue;
    }
    out += escapeRegex(ch);
  }
  return new RegExp(out, "i");
}

function escapeRegex(ch: string): string {
  return ch.replace(/[\\^$.*+?()[\]{}|]/g, "\\$&");
}

function substituteNth(text: string, oldText: string, newText: string, nth: number): string {
  let from = 0;
  let occurrence = 0;
  while (from <= text.length) {
    const idx = text.indexOf(oldText, from);
    if (idx === -1) return text;
    occurrence++;
    if (occurrence === nth) {
      return text.slice(0, idx) + newText + text.slice(idx + oldText.length);
    }
    from = idx + oldText.length;
  }
  return text;
}

/**
 * P0 minimal `TEXT` formatter. Recognises:
 *
 *   - plain digit blocks `0`, `0.0`, `0.00` → fixed decimals
 *   - `#,##0`, `#,##0.00` → thousands separator (comma) + decimals
 *   - `0%`, `0.0%` → percent
 *   - `$0.00`, `$#,##0.00` → leading dollar sign
 *
 * Anything else falls back to the default `toString` rendering of the
 * original value. The full number-format engine (§16.7) ships later.
 */
function formatNumberMinimal(value: number, fmt: string, original: Value): string {
  if (fmt === "") return formatFallback(original);
  let body = fmt;
  let prefix = "";
  if (body.startsWith("$")) {
    prefix = "$";
    body = body.slice(1);
  }
  let percent = false;
  if (body.endsWith("%")) {
    percent = true;
    body = body.slice(0, -1);
  }
  const useGroup = body.startsWith("#,##0");
  const digits = useGroup ? body.slice("#,##0".length) : body.startsWith("0") ? body.slice(1) : null;
  if (digits === null) return formatFallback(original);
  let decimals = 0;
  if (digits !== "") {
    if (!digits.startsWith(".")) return formatFallback(original);
    const tail = digits.slice(1);
    if (tail === "" || !/^0+$/.test(tail)) return formatFallback(original);
    decimals = tail.length;
  }
  let n = percent ? value * 100 : value;
  const negative = n < 0;
  if (negative) n = -n;
  const fixed = n.toFixed(decimals);
  const parts = fixed.split(".") as [string, string | undefined];
  let intPart = parts[0];
  const fracPart = parts[1];
  if (useGroup) intPart = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  let out = intPart;
  if (fracPart !== undefined) out += "." + fracPart;
  if (percent) out += "%";
  if (prefix !== "") out = prefix + out;
  if (negative) out = "-" + out;
  return out;
}

function formatFallback(v: Value): string {
  const s = toString(v);
  return isErrorString(s) ? "" : s.v;
}

function isErrorString(s: StringValue | ErrorValue): s is ErrorValue {
  return s.kind === "e";
}
