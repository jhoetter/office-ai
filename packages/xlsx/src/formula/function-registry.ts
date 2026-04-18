import type { AstNode } from "./ast.js";
import type { CellRef, RangeRef } from "./references.js";
import type { Range2D, Value } from "./values.js";

/**
 * Function-registry surface.
 *
 * Spec: `spec/xlsx/formula-engine.md` §13 (evaluator dispatches to
 * `registry.get(name)`) and §16 (function inventory).
 *
 * The registry is the single mutable surface in the formula engine —
 * functions register on construction and never change at runtime in
 * P0. The registry is intentionally tiny: it holds an `arity`
 * descriptor and an `eval` function; everything else (coercion, error
 * propagation, range collapsing) lives in `values.ts` so each function
 * impl stays linear.
 *
 * Lazy-arg functions (`IF`, `IFS`, `SWITCH`, `IFERROR`, `IFNA`)
 * declare `lazyArgs: true` and receive un-evaluated AST nodes plus
 * the eval context. The function impl drives evaluation order
 * itself — see §13.1 for the rationale.
 */

/**
 * Minimal accessors a lazy-arg function needs to drive its own
 * evaluation. Provided by `evaluator.ts`; we declare it here so the
 * registry has no import cycle on the evaluator.
 */
export interface LazyEvalAccess {
  evaluate(node: AstNode): Value;
  ctx: EvalContext;
}

/**
 * The minimum surface the evaluator + functions need from the host.
 * A real workbook implementation supplies a richer object; the
 * formula engine only ever touches these three accessors plus the
 * volatility-helper hooks.
 */
export interface EvalContext {
  /** Cell-data accessor; returns `Blank` for empty cells. */
  getCell(ref: CellRef): Value;
  /** Range materialisation; returns a 2-D array of Values. */
  getRange(ref: RangeRef): Range2D;
  /** Defined-name resolver (post-parse fallback). */
  resolveName(name: string): RangeRef | CellRef | undefined;
  /** Anchor for relative-time semantics (volatile fns). */
  now(): number;
  /** RNG for RAND / RANDBETWEEN; deterministic in tests. */
  random(): number;
  /** Anchor of the formula being evaluated (for ROW/COLUMN with no args). */
  anchor: CellRef;
  /** The function registry (passed through so functions can call each other). */
  registry: FunctionRegistry;
}

/** Arity descriptor; `accepts` is the only thing the dispatcher calls. */
export interface Arity {
  readonly min: number;
  readonly max: number;
  accepts(n: number): boolean;
}

export function arity(min: number, max: number): Arity {
  return {
    min,
    max,
    accepts(n: number): boolean {
      return n >= min && n <= max;
    },
  };
}

/**
 * Eager function: receives evaluated `Value` args and returns a
 * `Value`. Most functions are eager.
 */
export type EagerFn = (args: ReadonlyArray<Value>, ctx: EvalContext) => Value;

/**
 * Lazy function: receives un-evaluated AST nodes and an evaluator
 * accessor. Used by `IF`, `IFS`, `SWITCH`, `IFERROR`, `IFNA`. The
 * function impl is responsible for short-circuit evaluation.
 */
export type LazyFn = (args: ReadonlyArray<AstNode>, lazy: LazyEvalAccess) => Value;

export interface EagerEntry {
  readonly name: string;
  readonly arity: Arity;
  readonly volatile?: boolean;
  readonly lazyArgs?: false;
  readonly fn: EagerFn;
}

export interface LazyEntry {
  readonly name: string;
  readonly arity: Arity;
  readonly volatile?: boolean;
  readonly lazyArgs: true;
  readonly fn: LazyFn;
}

export type FunctionEntry = EagerEntry | LazyEntry;

export interface FunctionRegistry {
  /** Look up a function (case-insensitive name). */
  get(name: string): FunctionEntry | undefined;
  /** Whether the name is registered. */
  has(name: string): boolean;
  /** Iterate all entries (for diagnostics / docs). */
  entries(): IterableIterator<FunctionEntry>;
  /** Names of all volatile functions (used at parse time for `formula.volatile`). */
  volatileNames(): ReadonlySet<string>;
}

/** Mutable internal surface used by the registry builder. */
export interface MutableFunctionRegistry extends FunctionRegistry {
  register(entry: FunctionEntry): void;
}

export function createRegistry(): MutableFunctionRegistry {
  const map = new Map<string, FunctionEntry>();
  const volatile = new Set<string>();
  return {
    register(entry: FunctionEntry): void {
      const key = entry.name.toUpperCase();
      map.set(key, entry);
      if (entry.volatile) volatile.add(key);
    },
    get(name: string): FunctionEntry | undefined {
      return map.get(name.toUpperCase());
    },
    has(name: string): boolean {
      return map.has(name.toUpperCase());
    },
    entries(): IterableIterator<FunctionEntry> {
      return map.values();
    },
    volatileNames(): ReadonlySet<string> {
      return volatile;
    },
  };
}
