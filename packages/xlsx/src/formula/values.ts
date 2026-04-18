import { ErrorKinds, Errors, type CellError, type CellErrorKind } from "./errors.js";

/**
 * Runtime value type used at every layer above the lexer.
 *
 * Discriminated union with single-character `kind` tags for compact
 * debugger output and cheap dispatch (per
 * `analysis-univer-formula.md` §11 item 1).
 *
 * Spec: `spec/xlsx/formula-engine.md` §7.
 */
export type Value = NumberValue | StringValue | BoolValue | ErrorValue | RangeValue;

export interface NumberValue {
  readonly kind: "n";
  readonly v: number;
}
export interface StringValue {
  readonly kind: "s";
  readonly v: string;
}
export interface BoolValue {
  readonly kind: "b";
  readonly v: boolean;
}
export interface ErrorValue {
  readonly kind: "e";
  readonly v: CellError;
}
export interface RangeValue {
  readonly kind: "r";
  readonly v: Range2D;
}

export type Range2D = ReadonlyArray<ReadonlyArray<Value>>;

/** Sentinel for an empty cell. */
export const Blank: Value = { kind: "n", v: 0 };

export function num(v: number): NumberValue {
  return { kind: "n", v };
}
export function str(v: string): StringValue {
  return { kind: "s", v };
}
export function bool(v: boolean): BoolValue {
  return { kind: "b", v };
}
export function err(e: CellError | CellErrorKind): ErrorValue {
  return typeof e === "string" ? { kind: "e", v: Errors[e] } : { kind: "e", v: e };
}
export function rangeVal(v: Range2D): RangeValue {
  return { kind: "r", v };
}

export function isError(v: Value): v is ErrorValue {
  return v.kind === "e";
}

// ── Coercion ──────────────────────────────────────────────────────────────

/**
 * Coerce a Value to a number. Errors propagate; non-coercible strings
 * yield `#VALUE!`. Per §7.1.
 */
export function toNumber(v: Value): NumberValue | ErrorValue {
  switch (v.kind) {
    case "n":
      return v;
    case "b":
      return num(v.v ? 1 : 0);
    case "s": {
      const trimmed = v.v.trim();
      if (trimmed === "") return num(0);
      const n = Number(trimmed);
      if (!Number.isFinite(n)) return err(ErrorKinds.VALUE);
      return num(n);
    }
    case "e":
      return v;
    case "r":
      return collapseRange(v, toNumber);
  }
}

export function toString(v: Value): StringValue | ErrorValue {
  switch (v.kind) {
    case "s":
      return v;
    case "n":
      return str(numberToExcelString(v.v));
    case "b":
      return str(v.v ? "TRUE" : "FALSE");
    case "e":
      return v;
    case "r":
      return collapseRange(v, toString);
  }
}

export function toBoolean(v: Value): BoolValue | ErrorValue {
  switch (v.kind) {
    case "b":
      return v;
    case "n":
      return bool(v.v !== 0);
    case "s": {
      const upper = v.v.toUpperCase();
      if (upper === "TRUE") return bool(true);
      if (upper === "FALSE") return bool(false);
      return err(ErrorKinds.VALUE);
    }
    case "e":
      return v;
    case "r":
      return collapseRange(v, toBoolean);
  }
}

function collapseRange<T extends Value>(r: RangeValue, fn: (v: Value) => T | ErrorValue): T | ErrorValue {
  if (r.v.length === 1 && r.v[0].length === 1) return fn(r.v[0][0]);
  return err(ErrorKinds.VALUE);
}

function numberToExcelString(n: number): string {
  if (Number.isInteger(n)) return n.toString();
  // Excel general format: up to 15 significant digits.
  return Number(n.toPrecision(15)).toString();
}

// ── Comparison ────────────────────────────────────────────────────────────

/**
 * Excel cross-type comparison (`number < string < bool`). Same-type
 * uses standard ordering; strings are case-insensitive. Per §7.2.
 */
export function compare(a: Value, b: Value): -1 | 0 | 1 | ErrorValue {
  if (a.kind === "e") return a;
  if (b.kind === "e") return b;
  if (a.kind === "r" || b.kind === "r") {
    const ca = collapseScalar(a);
    const cb = collapseScalar(b);
    if (isError(ca)) return ca;
    if (isError(cb)) return cb;
    return compare(ca, cb);
  }
  const ra = typeRank(a);
  const rb = typeRank(b);
  if (ra !== rb) return ra < rb ? -1 : 1;
  if (a.kind === "n" && b.kind === "n") {
    if (a.v === b.v) return 0;
    return a.v < b.v ? -1 : 1;
  }
  if (a.kind === "s" && b.kind === "s") {
    const sa = a.v.toUpperCase();
    const sb = b.v.toUpperCase();
    if (sa === sb) return 0;
    return sa < sb ? -1 : 1;
  }
  if (a.kind === "b" && b.kind === "b") {
    if (a.v === b.v) return 0;
    return !a.v ? -1 : 1;
  }
  return 0;
}

function typeRank(v: Value): number {
  switch (v.kind) {
    case "n":
      return 0;
    case "s":
      return 1;
    case "b":
      return 2;
    case "e":
      return 3;
    case "r":
      return 4;
  }
}

function collapseScalar(v: Value): Value {
  if (v.kind === "r" && v.v.length === 1 && v.v[0].length === 1) return v.v[0][0];
  return v;
}

// ── Arithmetic helpers ────────────────────────────────────────────────────

export function add(a: Value, b: Value): Value {
  if (a.kind === "e") return a;
  if (b.kind === "e") return b;
  const na = toNumber(a);
  if (isError(na)) return na;
  const nb = toNumber(b);
  if (isError(nb)) return nb;
  return num(na.v + nb.v);
}

export function sub(a: Value, b: Value): Value {
  if (a.kind === "e") return a;
  if (b.kind === "e") return b;
  const na = toNumber(a);
  if (isError(na)) return na;
  const nb = toNumber(b);
  if (isError(nb)) return nb;
  return num(na.v - nb.v);
}

export function mul(a: Value, b: Value): Value {
  if (a.kind === "e") return a;
  if (b.kind === "e") return b;
  const na = toNumber(a);
  if (isError(na)) return na;
  const nb = toNumber(b);
  if (isError(nb)) return nb;
  return num(na.v * nb.v);
}

export function div(a: Value, b: Value): Value {
  if (a.kind === "e") return a;
  if (b.kind === "e") return b;
  const na = toNumber(a);
  if (isError(na)) return na;
  const nb = toNumber(b);
  if (isError(nb)) return nb;
  if (nb.v === 0) return err(ErrorKinds.DIV0);
  return num(na.v / nb.v);
}

export function pow(a: Value, b: Value): Value {
  if (a.kind === "e") return a;
  if (b.kind === "e") return b;
  const na = toNumber(a);
  if (isError(na)) return na;
  const nb = toNumber(b);
  if (isError(nb)) return nb;
  const r = Math.pow(na.v, nb.v);
  if (!Number.isFinite(r)) return err(ErrorKinds.NUM);
  return num(r);
}

export function neg(a: Value): Value {
  if (a.kind === "e") return a;
  const na = toNumber(a);
  if (isError(na)) return na;
  return num(-na.v);
}

export function pct(a: Value): Value {
  if (a.kind === "e") return a;
  const na = toNumber(a);
  if (isError(na)) return na;
  return num(na.v / 100);
}

export function concat(a: Value, b: Value): Value {
  if (a.kind === "e") return a;
  if (b.kind === "e") return b;
  const sa = toString(a);
  if (isError(sa)) return sa;
  const sb = toString(b);
  if (isError(sb)) return sb;
  return str(sa.v + sb.v);
}

// ── Comparison operators (return BoolValue, propagate errors) ─────────────

function cmpToValue(c: -1 | 0 | 1 | ErrorValue, want: (n: -1 | 0 | 1) => boolean): Value {
  if (typeof c !== "number") return c;
  return bool(want(c));
}

export function eq(a: Value, b: Value): Value {
  return cmpToValue(compare(a, b), (n) => n === 0);
}
export function neq(a: Value, b: Value): Value {
  return cmpToValue(compare(a, b), (n) => n !== 0);
}
export function lt(a: Value, b: Value): Value {
  return cmpToValue(compare(a, b), (n) => n < 0);
}
export function gt(a: Value, b: Value): Value {
  return cmpToValue(compare(a, b), (n) => n > 0);
}
export function lte(a: Value, b: Value): Value {
  return cmpToValue(compare(a, b), (n) => n <= 0);
}
export function gte(a: Value, b: Value): Value {
  return cmpToValue(compare(a, b), (n) => n >= 0);
}
