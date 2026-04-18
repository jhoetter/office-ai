import { ErrorKinds } from "../errors.js";
import { arity, type EagerFn, type EvalContext, type MutableFunctionRegistry } from "../function-registry.js";
import {
  err,
  isError,
  num,
  toNumber,
  type ErrorValue,
  type NumberValue,
  type Range2D,
  type Value,
} from "../values.js";

/**
 * P0 Math / Statistics functions.
 *
 * Spec: `spec/xlsx/formula-engine.md` §16.1.
 *
 * Each function is implemented as a small `EagerFn`. Coercion rules
 * follow §7.1 (scalar args) and the SUM-family range conventions
 * documented inline in `walkSumNumeric` (range strings silently
 * skipped, range bools coerced to 0/1, errors short-circuit).
 */

export function registerMath(reg: MutableFunctionRegistry): void {
  reg.register({ name: "SUM", arity: arity(1, 255), fn: SUM });
  reg.register({ name: "AVERAGE", arity: arity(1, 255), fn: AVERAGE });
  reg.register({ name: "COUNT", arity: arity(1, 255), fn: COUNT });
  reg.register({ name: "COUNTA", arity: arity(1, 255), fn: COUNTA });
  reg.register({ name: "COUNTBLANK", arity: arity(1, 1), fn: COUNTBLANK });
  reg.register({ name: "MIN", arity: arity(1, 255), fn: MIN });
  reg.register({ name: "MAX", arity: arity(1, 255), fn: MAX });
  reg.register({ name: "SUMIF", arity: arity(2, 3), fn: SUMIF });
  reg.register({ name: "SUMIFS", arity: arity(3, 255), fn: SUMIFS });
  reg.register({ name: "COUNTIF", arity: arity(2, 2), fn: COUNTIF });
  reg.register({ name: "COUNTIFS", arity: arity(2, 254), fn: COUNTIFS });
  reg.register({ name: "AVERAGEIF", arity: arity(2, 3), fn: AVERAGEIF });
  reg.register({ name: "AVERAGEIFS", arity: arity(3, 255), fn: AVERAGEIFS });
  reg.register({ name: "ROUND", arity: arity(2, 2), fn: ROUND });
  reg.register({ name: "ROUNDUP", arity: arity(2, 2), fn: ROUNDUP });
  reg.register({ name: "ROUNDDOWN", arity: arity(2, 2), fn: ROUNDDOWN });
  reg.register({ name: "INT", arity: arity(1, 1), fn: INT });
  reg.register({ name: "ABS", arity: arity(1, 1), fn: ABS });
  reg.register({ name: "MOD", arity: arity(2, 2), fn: MOD });
  reg.register({ name: "POWER", arity: arity(2, 2), fn: POWER });
  reg.register({ name: "SQRT", arity: arity(1, 1), fn: SQRT });
  reg.register({ name: "CEILING", arity: arity(1, 2), fn: CEILING });
  reg.register({ name: "FLOOR", arity: arity(1, 2), fn: FLOOR });
  reg.register({ name: "RAND", arity: arity(0, 0), volatile: true, fn: RAND });
  reg.register({
    name: "RANDBETWEEN",
    arity: arity(2, 2),
    volatile: true,
    fn: RANDBETWEEN,
  });
  reg.register({ name: "LARGE", arity: arity(2, 2), fn: LARGE });
  reg.register({ name: "SMALL", arity: arity(2, 2), fn: SMALL });
  reg.register({ name: "RANK", arity: arity(2, 3), fn: RANK });
  reg.register({ name: "MEDIAN", arity: arity(1, 255), fn: MEDIAN });
  reg.register({ name: "STDEV", arity: arity(1, 255), fn: STDEV });
  reg.register({ name: "VAR", arity: arity(1, 255), fn: VAR });
  reg.register({ name: "PRODUCT", arity: arity(1, 255), fn: PRODUCT });
  reg.register({ name: "SUMPRODUCT", arity: arity(1, 255), fn: SUMPRODUCT });
}

// ── Aggregations ──────────────────────────────────────────────────────────

const SUM: EagerFn = (args) => {
  let total = 0;
  const e = walkSumNumeric(args, (n) => {
    total += n;
  });
  if (e) return e;
  return num(total);
};

const AVERAGE: EagerFn = (args) => {
  let total = 0;
  let count = 0;
  const e = walkSumNumeric(args, (n) => {
    total += n;
    count++;
  });
  if (e) return e;
  if (count === 0) return err(ErrorKinds.DIV0);
  return num(total / count);
};

const PRODUCT: EagerFn = (args) => {
  let acc = 1;
  let count = 0;
  const e = walkSumNumeric(args, (n) => {
    acc *= n;
    count++;
  });
  if (e) return e;
  if (count === 0) return num(0);
  return acc === 0 ? num(0) : num(acc);
};

const MIN: EagerFn = (args) => {
  let best = Infinity;
  let count = 0;
  const e = walkSumNumeric(args, (n) => {
    if (n < best) best = n;
    count++;
  });
  if (e) return e;
  if (count === 0) return num(0);
  return num(best);
};

const MAX: EagerFn = (args) => {
  let best = -Infinity;
  let count = 0;
  const e = walkSumNumeric(args, (n) => {
    if (n > best) best = n;
    count++;
  });
  if (e) return e;
  if (count === 0) return num(0);
  return num(best);
};

// ── Counts ────────────────────────────────────────────────────────────────

const COUNT: EagerFn = (args) => {
  let count = 0;
  for (const a of args) {
    switch (a.kind) {
      case "e":
        return a;
      case "n":
        count++;
        break;
      case "b":
        count++;
        break;
      case "s": {
        const n = toNumber(a);
        if (!isError(n)) count++;
        break;
      }
      case "r": {
        for (const row of a.v) {
          for (const cell of row) {
            if (cell.kind === "e") return cell;
            if (cell.kind === "n") count++;
          }
        }
        break;
      }
    }
  }
  return num(count);
};

const COUNTA: EagerFn = (args) => {
  let count = 0;
  for (const a of args) {
    switch (a.kind) {
      case "e":
        return a;
      case "n":
      case "b":
        count++;
        break;
      case "s":
        // Excel counts empty strings produced by formulas as "non-blank".
        count++;
        break;
      case "r": {
        for (const row of a.v) {
          for (const cell of row) {
            if (cell.kind === "s" && cell.v === "") continue;
            count++;
          }
        }
        break;
      }
    }
  }
  return num(count);
};

const COUNTBLANK: EagerFn = (args) => {
  // Count cells whose current value is the empty string. Note: the
  // P0 host represents truly-blank cells as `num(0)`, so this can
  // only detect formula-blanks (`""`) — full-fidelity blank
  // detection is a §16.1 follow-up tracked in the build log.
  const a = args[0];
  switch (a.kind) {
    case "e":
      return a;
    case "n":
      return num(0);
    case "b":
      return num(0);
    case "s":
      return num(a.v === "" ? 1 : 0);
    case "r": {
      let count = 0;
      for (const row of a.v) {
        for (const cell of row) {
          if (cell.kind === "s" && cell.v === "") count++;
        }
      }
      return num(count);
    }
  }
};

// ── *IF / *IFS family ─────────────────────────────────────────────────────

const SUMIF: EagerFn = (args) => {
  const range = expectRange(args[0]);
  const sumRange = args.length === 3 ? expectRange(args[2]) : range;
  if (!sameShape(range, sumRange)) return err(ErrorKinds.VALUE);
  const c = parseCriteria(args[1]);
  if (isParsedError(c)) return c;
  let total = 0;
  for (let r = 0; r < range.length; r++) {
    for (let cIdx = 0; cIdx < range[r].length; cIdx++) {
      const cell = range[r][cIdx];
      if (cell.kind === "e") return cell;
      if (matchesCriteria(c, cell)) {
        const target = sumRange[r][cIdx];
        if (target.kind === "e") return target;
        if (target.kind === "n") total += target.v;
        else if (target.kind === "b") total += target.v ? 1 : 0;
      }
    }
  }
  return num(total);
};

const SUMIFS: EagerFn = (args) => {
  if ((args.length - 1) % 2 !== 0) return err(ErrorKinds.VALUE);
  const sumRange = expectRange(args[0]);
  const pairs = parseIfsPairs(args.slice(1), sumRange);
  if (isParsedError(pairs)) return pairs;
  let total = 0;
  const rows = sumRange.length;
  const cols = sumRange[0]?.length ?? 0;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      if (!allMatch(pairs, r, c)) continue;
      const cell = sumRange[r][c];
      if (cell.kind === "e") return cell;
      if (cell.kind === "n") total += cell.v;
      else if (cell.kind === "b") total += cell.v ? 1 : 0;
    }
  }
  return num(total);
};

const COUNTIF: EagerFn = (args) => {
  const range = expectRange(args[0]);
  const c = parseCriteria(args[1]);
  if (isParsedError(c)) return c;
  let count = 0;
  for (const row of range) {
    for (const cell of row) {
      if (cell.kind === "e") return cell;
      if (matchesCriteria(c, cell)) count++;
    }
  }
  return num(count);
};

const COUNTIFS: EagerFn = (args) => {
  if (args.length % 2 !== 0) return err(ErrorKinds.VALUE);
  const first = expectRange(args[0]);
  const pairs = parseIfsPairs(args, first);
  if (isParsedError(pairs)) return pairs;
  let count = 0;
  const rows = first.length;
  const cols = first[0]?.length ?? 0;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      if (allMatch(pairs, r, c)) count++;
    }
  }
  return num(count);
};

const AVERAGEIF: EagerFn = (args) => {
  const range = expectRange(args[0]);
  const sumRange = args.length === 3 ? expectRange(args[2]) : range;
  if (!sameShape(range, sumRange)) return err(ErrorKinds.VALUE);
  const c = parseCriteria(args[1]);
  if (isParsedError(c)) return c;
  let total = 0;
  let count = 0;
  for (let r = 0; r < range.length; r++) {
    for (let cIdx = 0; cIdx < range[r].length; cIdx++) {
      const cell = range[r][cIdx];
      if (cell.kind === "e") return cell;
      if (!matchesCriteria(c, cell)) continue;
      const target = sumRange[r][cIdx];
      if (target.kind === "e") return target;
      if (target.kind === "n") {
        total += target.v;
        count++;
      } else if (target.kind === "b") {
        total += target.v ? 1 : 0;
        count++;
      }
    }
  }
  if (count === 0) return err(ErrorKinds.DIV0);
  return num(total / count);
};

const AVERAGEIFS: EagerFn = (args) => {
  if ((args.length - 1) % 2 !== 0) return err(ErrorKinds.VALUE);
  const avgRange = expectRange(args[0]);
  const pairs = parseIfsPairs(args.slice(1), avgRange);
  if (isParsedError(pairs)) return pairs;
  let total = 0;
  let count = 0;
  const rows = avgRange.length;
  const cols = avgRange[0]?.length ?? 0;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      if (!allMatch(pairs, r, c)) continue;
      const cell = avgRange[r][c];
      if (cell.kind === "e") return cell;
      if (cell.kind === "n") {
        total += cell.v;
        count++;
      } else if (cell.kind === "b") {
        total += cell.v ? 1 : 0;
        count++;
      }
    }
  }
  if (count === 0) return err(ErrorKinds.DIV0);
  return num(total / count);
};

// ── Rounding ──────────────────────────────────────────────────────────────

const ROUND: EagerFn = (args) => roundOp(args, roundHalfAwayFromZero);
const ROUNDUP: EagerFn = (args) => roundOp(args, roundAwayFromZero);
const ROUNDDOWN: EagerFn = (args) => roundOp(args, roundTowardZero);

const INT: EagerFn = (args) => {
  const x = toScalarNumber(args[0]);
  if (isError(x)) return x;
  return num(Math.floor(x.v));
};

const ABS: EagerFn = (args) => {
  const x = toScalarNumber(args[0]);
  if (isError(x)) return x;
  return num(Math.abs(x.v));
};

const MOD: EagerFn = (args) => {
  const a = toScalarNumber(args[0]);
  if (isError(a)) return a;
  const b = toScalarNumber(args[1]);
  if (isError(b)) return b;
  if (b.v === 0) return err(ErrorKinds.DIV0);
  // Excel: result has same sign as divisor (n - d * INT(n/d)).
  const r = a.v - b.v * Math.floor(a.v / b.v);
  return num(r);
};

const POWER: EagerFn = (args) => {
  const base = toScalarNumber(args[0]);
  if (isError(base)) return base;
  const exp = toScalarNumber(args[1]);
  if (isError(exp)) return exp;
  const r = Math.pow(base.v, exp.v);
  if (!Number.isFinite(r) || Number.isNaN(r)) return err(ErrorKinds.NUM);
  return num(r);
};

const SQRT: EagerFn = (args) => {
  const x = toScalarNumber(args[0]);
  if (isError(x)) return x;
  if (x.v < 0) return err(ErrorKinds.NUM);
  return num(Math.sqrt(x.v));
};

const CEILING: EagerFn = (args) => {
  const x = toScalarNumber(args[0]);
  if (isError(x)) return x;
  const sigArg = args.length === 2 ? toScalarNumber(args[1]) : num(1);
  if (isError(sigArg)) return sigArg;
  const sig = sigArg.v;
  if (sig === 0) return num(0);
  if (x.v === 0) return num(0);
  // Modern semantics: round magnitude up to nearest |sig|, preserve sign of x.
  const mag = Math.ceil(Math.abs(x.v) / Math.abs(sig)) * Math.abs(sig);
  return num(Math.sign(x.v) * mag);
};

const FLOOR: EagerFn = (args) => {
  const x = toScalarNumber(args[0]);
  if (isError(x)) return x;
  const sigArg = args.length === 2 ? toScalarNumber(args[1]) : num(1);
  if (isError(sigArg)) return sigArg;
  const sig = sigArg.v;
  if (sig === 0) return err(ErrorKinds.DIV0);
  // Modern semantics: round magnitude down to nearest |sig|, preserve sign of x.
  const mag = Math.floor(Math.abs(x.v) / Math.abs(sig)) * Math.abs(sig);
  return num(Math.sign(x.v) * mag);
};

// ── Random ────────────────────────────────────────────────────────────────

const RAND: EagerFn = (_args, ctx: EvalContext) => num(ctx.random());

const RANDBETWEEN: EagerFn = (args, ctx: EvalContext) => {
  const a = toScalarNumber(args[0]);
  if (isError(a)) return a;
  const b = toScalarNumber(args[1]);
  if (isError(b)) return b;
  const lo = Math.ceil(a.v);
  const hi = Math.floor(b.v);
  if (lo > hi) return err(ErrorKinds.NUM);
  return num(Math.floor(ctx.random() * (hi - lo + 1)) + lo);
};

// ── Order statistics ──────────────────────────────────────────────────────

const LARGE: EagerFn = (args) => {
  const data = collectRangeNumbers(args[0]);
  if (!Array.isArray(data)) return data;
  const k = toScalarNumber(args[1]);
  if (isError(k)) return k;
  const ki = Math.trunc(k.v);
  if (ki < 1 || ki > data.length) return err(ErrorKinds.NUM);
  const sorted = [...data].sort((x, y) => y - x);
  return num(sorted[ki - 1]);
};

const SMALL: EagerFn = (args) => {
  const data = collectRangeNumbers(args[0]);
  if (!Array.isArray(data)) return data;
  const k = toScalarNumber(args[1]);
  if (isError(k)) return k;
  const ki = Math.trunc(k.v);
  if (ki < 1 || ki > data.length) return err(ErrorKinds.NUM);
  const sorted = [...data].sort((x, y) => x - y);
  return num(sorted[ki - 1]);
};

const RANK: EagerFn = (args) => {
  const v = toScalarNumber(args[0]);
  if (isError(v)) return v;
  const data = collectRangeNumbers(args[1]);
  if (!Array.isArray(data)) return data;
  let order = 0;
  if (args.length === 3) {
    const o = toScalarNumber(args[2]);
    if (isError(o)) return o;
    order = o.v;
  }
  const ascending = order !== 0;
  const sorted = [...data].sort((x, y) => (ascending ? x - y : y - x));
  const idx = sorted.indexOf(v.v);
  if (idx < 0) return err(ErrorKinds.NA);
  return num(idx + 1);
};

const MEDIAN: EagerFn = (args) => {
  const data: number[] = [];
  const e = walkSumNumeric(args, (n) => {
    data.push(n);
  });
  if (e) return e;
  if (data.length === 0) return err(ErrorKinds.NUM);
  data.sort((a, b) => a - b);
  const mid = Math.floor(data.length / 2);
  if (data.length % 2 === 1) return num(data[mid]);
  return num((data[mid - 1] + data[mid]) / 2);
};

const STDEV: EagerFn = (args) => {
  const v = sampleVariance(args);
  if (isError(v)) return v;
  return num(Math.sqrt(v.v));
};

const VAR: EagerFn = (args) => sampleVariance(args);

// ── SUMPRODUCT ────────────────────────────────────────────────────────────

const SUMPRODUCT: EagerFn = (args) => {
  const ranges: Range2D[] = [];
  for (const a of args) {
    if (a.kind === "e") return a;
    ranges.push(a.kind === "r" ? a.v : [[a]]);
  }
  if (ranges.length === 0) return num(0);
  const rows = ranges[0].length;
  const cols = ranges[0][0]?.length ?? 0;
  for (const r of ranges) {
    if (r.length !== rows) return err(ErrorKinds.VALUE);
    for (const row of r) {
      if (row.length !== cols) return err(ErrorKinds.VALUE);
    }
  }
  let total = 0;
  for (let i = 0; i < rows; i++) {
    for (let j = 0; j < cols; j++) {
      let prod = 1;
      for (const r of ranges) {
        const cell = r[i][j];
        if (cell.kind === "e") return cell;
        // SUMPRODUCT treats text as 0 (Excel parity).
        if (cell.kind === "n") prod *= cell.v;
        else if (cell.kind === "b") prod *= cell.v ? 1 : 0;
        else {
          prod = 0;
          break;
        }
      }
      total += prod;
    }
  }
  return num(total);
};

// ── Helpers ───────────────────────────────────────────────────────────────

/**
 * Visit each numeric value in an arg list following the SUM-family
 * rules:
 *
 * - scalar number / boolean → coerce
 * - scalar string → coerce per §7.1 (`#VALUE!` if non-numeric)
 * - range cell number → include
 * - range cell boolean → coerce 0/1
 * - range cell string → silently skip
 * - any error (scalar or in range) → short-circuit return
 */
function walkSumNumeric(args: ReadonlyArray<Value>, visit: (n: number) => void): ErrorValue | undefined {
  for (const a of args) {
    switch (a.kind) {
      case "e":
        return a;
      case "n":
        visit(a.v);
        break;
      case "b":
        visit(a.v ? 1 : 0);
        break;
      case "s": {
        const n = toNumber(a);
        if (isError(n)) return n;
        visit(n.v);
        break;
      }
      case "r": {
        for (const row of a.v) {
          for (const cell of row) {
            if (cell.kind === "e") return cell;
            if (cell.kind === "n") visit(cell.v);
            else if (cell.kind === "b") visit(cell.v ? 1 : 0);
            // cell.kind === "s" or nested "r": skipped
          }
        }
        break;
      }
    }
  }
  return undefined;
}

function toScalarNumber(v: Value): NumberValue | ErrorValue {
  return toNumber(v);
}

function expectRange(v: Value): Range2D {
  if (v.kind === "r") return v.v;
  return [[v]];
}

function sameShape(a: Range2D, b: Range2D): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (a[i].length !== b[i].length) return false;
  }
  return true;
}

interface IfsPair {
  range: Range2D;
  criteria: ParsedCriteria;
}

function parseIfsPairs(args: ReadonlyArray<Value>, shape: Range2D): IfsPair[] | ErrorValue {
  const out: IfsPair[] = [];
  for (let i = 0; i + 1 < args.length; i += 2) {
    const range = expectRange(args[i]);
    if (!sameShape(range, shape)) return err(ErrorKinds.VALUE);
    const c = parseCriteria(args[i + 1]);
    if (isParsedError(c)) return c;
    out.push({ range, criteria: c });
  }
  return out;
}

function allMatch(pairs: ReadonlyArray<IfsPair>, r: number, c: number): boolean {
  for (const p of pairs) {
    const cell = p.range[r][c];
    if (cell.kind === "e") return false;
    if (!matchesCriteria(p.criteria, cell)) return false;
  }
  return true;
}

function isParsedError<T>(x: T | ErrorValue): x is ErrorValue {
  return typeof x === "object" && x !== null && (x as { kind?: string }).kind === "e";
}

// ── Criteria parsing for *IF / *IFS ───────────────────────────────────────

type CriteriaOp = "=" | "<>" | ">" | "<" | ">=" | "<=";

interface ParsedCriteria {
  readonly op: CriteriaOp;
  readonly text: string;
  readonly numeric?: number;
  readonly regex?: RegExp;
}

function parseCriteria(v: Value): ParsedCriteria | ErrorValue {
  switch (v.kind) {
    case "e":
      return v;
    case "n":
      return { op: "=", text: String(v.v), numeric: v.v };
    case "b":
      return { op: "=", text: v.v ? "TRUE" : "FALSE" };
    case "r": {
      if (v.v.length === 1 && v.v[0].length === 1) return parseCriteria(v.v[0][0]);
      return err(ErrorKinds.VALUE);
    }
    case "s": {
      const s = v.v;
      let op: CriteriaOp = "=";
      let rest = s;
      if (s.startsWith("<>") || s.startsWith(">=") || s.startsWith("<=")) {
        op = s.slice(0, 2) as CriteriaOp;
        rest = s.slice(2);
      } else if (s.startsWith("=") || s.startsWith(">") || s.startsWith("<")) {
        op = s[0] as CriteriaOp;
        rest = s.slice(1);
      }
      const trimmed = rest.trim();
      const result: { -readonly [K in keyof ParsedCriteria]: ParsedCriteria[K] } = {
        op,
        text: rest,
      };
      if (trimmed !== "") {
        const n = Number(trimmed);
        if (Number.isFinite(n)) result.numeric = n;
      }
      if ((op === "=" || op === "<>") && (rest.includes("*") || rest.includes("?") || rest.includes("~"))) {
        result.regex = buildWildcardRegex(rest);
      }
      return result;
    }
  }
}

function matchesCriteria(c: ParsedCriteria, v: Value): boolean {
  if (v.kind === "e") return false;
  if (v.kind === "r") return false;
  if (c.regex) {
    const text = stringForCriteria(v);
    const m = c.regex.test(text);
    return c.op === "=" ? m : !m;
  }
  if (c.numeric !== undefined) {
    const n = numericForCriteria(v);
    if (n === undefined) {
      // Numeric criteria against non-numeric cell: only `<>` matches.
      return c.op === "<>";
    }
    switch (c.op) {
      case "=":
        return n === c.numeric;
      case "<>":
        return n !== c.numeric;
      case ">":
        return n > c.numeric;
      case "<":
        return n < c.numeric;
      case ">=":
        return n >= c.numeric;
      case "<=":
        return n <= c.numeric;
    }
  }
  // Text comparison.
  const a = stringForCriteria(v).toUpperCase();
  const b = c.text.toUpperCase();
  switch (c.op) {
    case "=":
      return a === b;
    case "<>":
      return a !== b;
    case ">":
      return a > b;
    case "<":
      return a < b;
    case ">=":
      return a >= b;
    case "<=":
      return a <= b;
  }
}

function stringForCriteria(v: Value): string {
  switch (v.kind) {
    case "s":
      return v.v;
    case "n":
      return String(v.v);
    case "b":
      return v.v ? "TRUE" : "FALSE";
    case "e":
      return "";
    case "r":
      return "";
  }
}

function numericForCriteria(v: Value): number | undefined {
  switch (v.kind) {
    case "n":
      return v.v;
    case "b":
      return v.v ? 1 : 0;
    case "s": {
      const t = v.v.trim();
      if (t === "") return undefined;
      const n = Number(t);
      return Number.isFinite(n) ? n : undefined;
    }
    case "e":
      return undefined;
    case "r":
      return undefined;
  }
}

function buildWildcardRegex(pattern: string): RegExp {
  let re = "";
  for (let i = 0; i < pattern.length; i++) {
    const c = pattern[i];
    if (c === "~" && i + 1 < pattern.length) {
      const next = pattern[i + 1];
      if (next === "*" || next === "?" || next === "~") {
        re += escapeRegex(next);
        i++;
        continue;
      }
    }
    if (c === "*") re += ".*";
    else if (c === "?") re += ".";
    else re += escapeRegex(c);
  }
  return new RegExp("^" + re + "$", "i");
}

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// ── Rounding helpers ──────────────────────────────────────────────────────

function roundOp(args: ReadonlyArray<Value>, fn: (x: number, factor: number) => number): Value {
  const x = toScalarNumber(args[0]);
  if (isError(x)) return x;
  const d = toScalarNumber(args[1]);
  if (isError(d)) return d;
  const digits = Math.trunc(d.v);
  const factor = Math.pow(10, digits);
  return num(fn(x.v, factor));
}

function roundHalfAwayFromZero(x: number, factor: number): number {
  return (Math.sign(x) * Math.round(Math.abs(x) * factor)) / factor;
}

function roundAwayFromZero(x: number, factor: number): number {
  return (Math.sign(x) * Math.ceil(Math.abs(x) * factor)) / factor;
}

function roundTowardZero(x: number, factor: number): number {
  return (Math.sign(x) * Math.floor(Math.abs(x) * factor)) / factor;
}

// ── LARGE / SMALL / RANK helpers ──────────────────────────────────────────

function collectRangeNumbers(v: Value): number[] | ErrorValue {
  const out: number[] = [];
  const e = walkSumNumeric([v], (n) => {
    out.push(n);
  });
  if (e) return e;
  return out;
}

// ── STDEV / VAR ───────────────────────────────────────────────────────────

function sampleVariance(args: ReadonlyArray<Value>): NumberValue | ErrorValue {
  const data: number[] = [];
  const e = walkSumNumeric(args, (n) => {
    data.push(n);
  });
  if (e) return e;
  if (data.length < 2) return err(ErrorKinds.DIV0);
  const mean = data.reduce((a, b) => a + b, 0) / data.length;
  let s = 0;
  for (const n of data) {
    const d = n - mean;
    s += d * d;
  }
  return num(s / (data.length - 1));
}
