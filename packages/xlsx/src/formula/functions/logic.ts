import type { AstNode } from "../ast.js";
import { ErrorKinds } from "../errors.js";
import { arity, type LazyEvalAccess, type MutableFunctionRegistry } from "../function-registry.js";
import { bool, err, isError, toBoolean, type BoolValue, type ErrorValue, type Value } from "../values.js";

/**
 * Logic-category functions (P0). Spec: `spec/xlsx/formula-engine.md`
 * §16.2.
 *
 * Lazy-arg functions (`IF`, `IFS`, `IFERROR`, `IFNA`, `SWITCH`)
 * declare `lazyArgs: true` so the impl drives evaluation order and
 * skips un-chosen branches — Excel parity for short-circuit semantics
 * and to keep side-effects (volatile fns) to the active path.
 */
export function registerLogic(reg: MutableFunctionRegistry): void {
  reg.register({
    name: "IF",
    arity: arity(2, 3),
    lazyArgs: true,
    fn: (args, lazy) => {
      const cond = lazy.evaluate(args[0]);
      if (isError(cond)) return cond;
      const truthy = coerceCondition(cond);
      if (isError(truthy)) return truthy;
      if (truthy.v) return lazy.evaluate(args[1]);
      return args.length === 3 ? lazy.evaluate(args[2]) : bool(false);
    },
  });

  reg.register({
    name: "IFS",
    arity: arity(2, 254),
    lazyArgs: true,
    fn: (args, lazy) => {
      // Pairs of (condition, value). Excel returns #N/A when no
      // condition matches; an odd trailing arg is **not** a default
      // (that is `SWITCH` semantics).
      if (args.length % 2 !== 0) return err(ErrorKinds.NA);
      for (let i = 0; i < args.length; i += 2) {
        const cond = lazy.evaluate(args[i]);
        if (isError(cond)) return cond;
        const truthy = coerceCondition(cond);
        if (isError(truthy)) return truthy;
        if (truthy.v) return lazy.evaluate(args[i + 1]);
      }
      return err(ErrorKinds.NA);
    },
  });

  reg.register({
    name: "AND",
    arity: arity(1, 255),
    fn: (args) => {
      // Excel evaluates every arg eagerly, but logically short-circuits
      // on the first FALSE for the return value.
      let sawAny = false;
      for (const a of args) {
        const b = collectBools(a);
        if (!Array.isArray(b)) return b;
        for (const v of b) {
          sawAny = true;
          if (!v) return bool(false);
        }
      }
      if (!sawAny) return err(ErrorKinds.VALUE);
      return bool(true);
    },
  });

  reg.register({
    name: "OR",
    arity: arity(1, 255),
    fn: (args) => {
      let sawAny = false;
      let result = false;
      for (const a of args) {
        const b = collectBools(a);
        if (!Array.isArray(b)) return b;
        for (const v of b) {
          sawAny = true;
          if (v) result = true;
        }
      }
      if (!sawAny) return err(ErrorKinds.VALUE);
      return bool(result);
    },
  });

  reg.register({
    name: "NOT",
    arity: arity(1, 1),
    fn: (args) => {
      const b = toBoolean(args[0]);
      if (isError(b)) return b;
      return bool(!b.v);
    },
  });

  reg.register({
    name: "XOR",
    arity: arity(1, 255),
    fn: (args) => {
      let trues = 0;
      let sawAny = false;
      for (const a of args) {
        const b = collectBools(a);
        if (!Array.isArray(b)) return b;
        for (const v of b) {
          sawAny = true;
          if (v) trues++;
        }
      }
      if (!sawAny) return err(ErrorKinds.VALUE);
      return bool(trues % 2 === 1);
    },
  });

  reg.register({
    name: "IFERROR",
    arity: arity(2, 2),
    lazyArgs: true,
    fn: (args, lazy) => {
      const v = lazy.evaluate(args[0]);
      if (isError(v)) return lazy.evaluate(args[1]);
      return v;
    },
  });

  reg.register({
    name: "IFNA",
    arity: arity(2, 2),
    lazyArgs: true,
    fn: (args, lazy) => {
      const v = lazy.evaluate(args[0]);
      if (isError(v) && v.v.kind === ErrorKinds.NA) return lazy.evaluate(args[1]);
      return v;
    },
  });

  reg.register({
    name: "SWITCH",
    arity: arity(3, 254),
    lazyArgs: true,
    fn: (args, lazy) => evaluateSwitch(args, lazy),
  });

  reg.register({
    name: "TRUE",
    arity: arity(0, 0),
    fn: () => bool(true),
  });

  reg.register({
    name: "FALSE",
    arity: arity(0, 0),
    fn: () => bool(false),
  });
}

// ── Helpers ────────────────────────────────────────────────────────────────

/**
 * Coerce an `IF` / `IFS` condition to a boolean. Numbers ≠ 0 are
 * truthy, 0 is falsy, strings go through `toBoolean` (only "TRUE" /
 * "FALSE" accepted; everything else is `#VALUE!`). Errors propagate.
 */
function coerceCondition(v: Value): BoolValue | ErrorValue {
  switch (v.kind) {
    case "b":
      return v;
    case "n":
      return bool(v.v !== 0);
    case "s":
      return toBoolean(v);
    case "e":
      return v;
    case "r":
      return toBoolean(v);
  }
}

/**
 * Flatten an arg into a sequence of boolean values for `AND` / `OR` /
 * `XOR`. Numbers coerce by `≠ 0`; strings are skipped (Excel parity:
 * text inside a range is ignored, but a scalar text arg coerces and
 * may yield `#VALUE!`). Errors propagate.
 */
function collectBools(v: Value): boolean[] | ErrorValue {
  const out: boolean[] = [];
  switch (v.kind) {
    case "e":
      return v;
    case "b":
      out.push(v.v);
      return out;
    case "n":
      out.push(v.v !== 0);
      return out;
    case "s": {
      const b = toBoolean(v);
      if (isError(b)) return b;
      out.push(b.v);
      return out;
    }
    case "r": {
      for (const row of v.v) {
        for (const cell of row) {
          switch (cell.kind) {
            case "e":
              return cell;
            case "b":
              out.push(cell.v);
              break;
            case "n":
              out.push(cell.v !== 0);
              break;
            case "s":
              // Per Excel, strings inside a range arg are ignored.
              break;
            case "r":
              // Nested ranges should not occur post-evaluation, but be
              // defensive: ignore.
              break;
          }
        }
      }
      return out;
    }
  }
}

/**
 * SWITCH(expr, c1, v1, c2, v2, …, [default]).
 *
 * If the trailing arg count is odd, the last arg is the default; if
 * omitted and no condition matches, return `#N/A`. Comparison uses
 * Excel scalar equality (case-insensitive for strings).
 */
function evaluateSwitch(args: ReadonlyArray<AstNode>, lazy: LazyEvalAccess): Value {
  const expr = lazy.evaluate(args[0]);
  if (isError(expr)) return expr;
  const remaining = args.length - 1;
  const pairs = Math.floor(remaining / 2);
  const hasDefault = remaining % 2 === 1;
  for (let i = 0; i < pairs; i++) {
    const candidate = lazy.evaluate(args[1 + i * 2]);
    if (isError(candidate)) return candidate;
    if (scalarEquals(expr, candidate)) return lazy.evaluate(args[1 + i * 2 + 1]);
  }
  if (hasDefault) return lazy.evaluate(args[args.length - 1]);
  return err(ErrorKinds.NA);
}

/**
 * Excel SWITCH equality: same-kind scalar compare; numbers compare by
 * `===`, strings case-insensitively, booleans by value. Anything
 * else (including range or cross-kind) returns false.
 */
function scalarEquals(a: Value, b: Value): boolean {
  if (a.kind !== b.kind) return false;
  switch (a.kind) {
    case "n":
      return b.kind === "n" && a.v === b.v;
    case "s":
      return b.kind === "s" && a.v.toUpperCase() === b.v.toUpperCase();
    case "b":
      return b.kind === "b" && a.v === b.v;
    case "e":
      return false;
    case "r":
      return false;
  }
}
