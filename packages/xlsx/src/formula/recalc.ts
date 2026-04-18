import type { Formula } from "./ast.js";
import { createDepGraph, cycleError, refForKey, type DepGraph } from "./dependency-graph.js";
import { evaluate } from "./evaluator.js";
import { type EvalContext, type FunctionRegistry } from "./function-registry.js";
import { parse } from "./parser.js";
import { cellRefKey, type CellKey, type CellRef, type RangeRef } from "./references.js";
import { type Range2D, type Value } from "./values.js";

/**
 * Recalc orchestrator — the public surface every command handler that
 * touches cells calls into.
 *
 * Spec: `spec/xlsx/formula-engine.md` §15.
 *
 * The engine is **passive**: it owns the dependency graph and the
 * cached values, but a host (the workbook model) is responsible for
 * supplying cell/range data via `EngineHost.read*` callbacks. This
 * keeps the engine deployable both inside the headless `XlsxAgent` and
 * inside the browser editor without any DOM dependency.
 */
export interface EngineHost {
  /**
   * Read a cell's *raw* value (i.e. the value the user typed or the
   * cached value of a non-formula cell). Formulas are looked up via
   * the engine's own cache, not the host.
   */
  readCell(ref: CellRef): Value;
  /** Materialise a range; the host decides how to bound whole-col/row. */
  readRange(ref: RangeRef): Range2D;
  /** Defined-name resolver (post-parse fallback). */
  resolveName?(name: string): RangeRef | CellRef | undefined;
  /** Optional clock; defaults to `Date.now`. */
  now?(): number;
  /** Optional RNG; defaults to `Math.random`. */
  random?(): number;
}

export interface RecalcResult {
  values: Map<CellKey, Value>;
  cycles: ReadonlyArray<ReadonlyArray<CellKey>>;
  elapsedMs: number;
}

export interface FormulaEngine {
  parse(text: string, anchor: CellRef): Formula;
  addCell(ref: CellRef, formula: Formula | null, value: Value | null): void;
  removeCell(ref: CellRef): void;
  onCellChanged(ref: CellRef): void;
  recalc(): RecalcResult;
  getCachedValue(ref: CellRef): Value | undefined;
  recalcAll(): RecalcResult;
  readonly registry: FunctionRegistry;
  readonly graph: DepGraph;
}

export interface EngineOptions {
  readonly registry: FunctionRegistry;
  readonly host: EngineHost;
}

export function createFormulaEngine(opts: EngineOptions): FormulaEngine {
  const graph = createDepGraph();
  const registry = opts.registry;
  const host = opts.host;
  const now = host.now ?? (() => Date.now());
  const random = host.random ?? (() => Math.random());

  function buildContext(anchor: CellRef): EvalContext {
    return {
      getCell(ref: CellRef): Value {
        // Formula cells are served from the engine's cache; non-formula
        // cells fall through to the host. This is the contract that
        // enforces dependency ordering — by the time a downstream cell
        // evaluates, its precedents already wrote their cached values.
        const cell = graph.getCell(cellRefKey(ref));
        if (cell) return cell.cachedValue;
        return host.readCell(ref);
      },
      getRange(ref: RangeRef): Range2D {
        // For ranges, we ask the host to materialise — the host is
        // responsible for stitching cached formula values + raw cell
        // values into a uniform 2-D array.
        return host.readRange(ref);
      },
      resolveName(name: string) {
        return host.resolveName?.(name);
      },
      now,
      random,
      anchor,
      registry,
    };
  }

  function recalcCore(): RecalcResult {
    const start = nowMs();
    const { order, cycles } = graph.drainTopological();
    const values = new Map<CellKey, Value>();

    for (const scc of cycles) {
      const cycleVal = cycleError(scc);
      for (const k of scc) {
        graph.setCachedValue(k, cycleVal);
        values.set(k, cycleVal);
      }
    }

    for (const k of order) {
      const cell = graph.getCell(k);
      if (!cell) continue;
      if (cell.kind === "formula") {
        const ref = refForKey(k);
        const ctx = buildContext(ref);
        const v = evaluate(cell.formula.ast, ctx);
        graph.setCachedValue(k, v);
        values.set(k, v);
      } else {
        // Value cell that was marked dirty (e.g. via `onCellChanged`)
        // — re-read from the host to refresh the cache.
        const ref = refForKey(k);
        const v = host.readCell(ref);
        graph.setCachedValue(k, v);
        values.set(k, v);
      }
    }

    return { values, cycles, elapsedMs: nowMs() - start };
  }

  return {
    parse(text: string, anchor: CellRef): Formula {
      return parse(text, { anchor });
    },
    addCell(ref: CellRef, formula: Formula | null, value: Value | null): void {
      graph.addCell(ref, formula, value);
    },
    removeCell(ref: CellRef): void {
      graph.removeCell(ref);
    },
    onCellChanged(ref: CellRef): void {
      graph.markDirty(ref);
    },
    recalc(): RecalcResult {
      return recalcCore();
    },
    getCachedValue(ref: CellRef): Value | undefined {
      return graph.getCachedValue(ref);
    },
    recalcAll(): RecalcResult {
      // Force every known cell into the dirty set, then drain.
      for (const key of graph.allKeys()) {
        const ref = refForKey(key);
        graph.markDirty(ref);
      }
      return recalcCore();
    },
    registry,
    graph,
  };
}

function nowMs(): number {
  return typeof performance !== "undefined" && typeof performance.now === "function"
    ? performance.now()
    : Date.now();
}
