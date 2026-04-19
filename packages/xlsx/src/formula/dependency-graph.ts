import type { Formula } from "./ast.js";
import { refWithCycle } from "./errors.js";
import {
  cellRefKey,
  makeCellKey,
  parseCellKey,
  type CellKey,
  type CellRef,
  type RangeRef,
} from "./references.js";
import { err, type Value } from "./values.js";

/**
 * Forward dependency graph + range-overlay index used to drive
 * incremental recalc.
 *
 * Spec: `spec/xlsx/formula-engine.md` §14.
 *
 * v1 implementation deliberately favours clarity over micro-perf:
 *   - Cell ↔ cell edges live in two `Map<CellKey, Set<CellKey>>`
 *     (`forward` = "who reads me", `reverse` = "who do I read").
 *   - Range dependencies live in a per-sheet flat list of
 *     `{ rangeRef, dependentKey }` and are resolved at
 *     `markDirty`/`markRangeDirty` time by linear scan. This is
 *     fine for the §17 perf budget on the 10k-formula fixture; we
 *     swap in an interval/R-tree if profiling demands it.
 *   - Cycle detection uses Tarjan's SCC over the residual subgraph
 *     after Kahn's topological pop.
 */

export interface FormulaCell {
  readonly kind: "formula";
  readonly formula: Formula;
  cachedValue: Value;
}

export interface ValueCell {
  readonly kind: "value";
  cachedValue: Value;
}

export type AnyCell = FormulaCell | ValueCell;

export interface DepGraph {
  /** Add or replace the cell at `ref`. Pass `formula` xor `value`. */
  addCell(ref: CellRef, formula: Formula | null, value: Value | null): void;
  /** Remove a cell from the graph (also drops its edges). */
  removeCell(ref: CellRef): void;
  /** Mark a single cell — and every cell that transitively depends on it — dirty. */
  markDirty(ref: CellRef): void;
  /** Bulk dirty for an inserted/deleted range. */
  markRangeDirty(ref: RangeRef): void;
  /** Drain the dirty set; returns cells in safe topological order plus any cycles. */
  drainTopological(): { order: ReadonlyArray<CellKey>; cycles: ReadonlyArray<ReadonlyArray<CellKey>> };
  /** Read the cached value (post-recalc). */
  getCachedValue(ref: CellRef): Value | undefined;
  /** Read the FormulaCell for a key (used by recalc). Returns undefined for value cells / missing. */
  getCell(key: CellKey): AnyCell | undefined;
  /** Write the cached value for a key (called by recalc after evaluation). */
  setCachedValue(key: CellKey, value: Value): void;
  /** Iterate every cell key (for `recalcAll`). */
  allKeys(): ReadonlyArray<CellKey>;
  /** Names of volatile cell keys (force-dirty every recalc). */
  volatileKeys(): ReadonlySet<CellKey>;
  /**
   * Collect every cell key that transitively depends on `ref` —
   * direct readers via forward edges plus any formula whose range
   * dependency includes `ref`. Read-only counterpart to
   * `markDirty`: useful for command handlers that want to
   * propagate cached-value updates to dependents only, without
   * touching unrelated formula cells (preserving OOXML
   * round-trip identity for parts that didn't actually change).
   */
  collectDependents(ref: CellRef): ReadonlySet<CellKey>;
  /** Total cell count. */
  readonly size: number;
}

interface RangeDep {
  readonly range: RangeRef;
  readonly dependentKey: CellKey;
}

interface State {
  readonly cells: Map<CellKey, AnyCell>;
  /** Forward edges: cell key → cells that *read* it. */
  readonly forward: Map<CellKey, Set<CellKey>>;
  /** Reverse edges: cell key → cells that it reads. */
  readonly reverse: Map<CellKey, Set<CellKey>>;
  /** Per-sheet flat list of range-dependency entries. */
  readonly rangeIndex: Map<string, RangeDep[]>;
  readonly volatile: Set<CellKey>;
  readonly dirty: Set<CellKey>;
}

export function createDepGraph(): DepGraph {
  const state: State = {
    cells: new Map(),
    forward: new Map(),
    reverse: new Map(),
    rangeIndex: new Map(),
    volatile: new Set(),
    dirty: new Set(),
  };

  function ensureForward(key: CellKey): Set<CellKey> {
    let set = state.forward.get(key);
    if (!set) {
      set = new Set();
      state.forward.set(key, set);
    }
    return set;
  }

  function ensureReverse(key: CellKey): Set<CellKey> {
    let set = state.reverse.get(key);
    if (!set) {
      set = new Set();
      state.reverse.set(key, set);
    }
    return set;
  }

  function ensureRangeBucket(sheet: string): RangeDep[] {
    let arr = state.rangeIndex.get(sheet);
    if (!arr) {
      arr = [];
      state.rangeIndex.set(sheet, arr);
    }
    return arr;
  }

  function dropEdges(key: CellKey): void {
    const reads = state.reverse.get(key);
    if (reads) {
      for (const dep of reads) {
        const back = state.forward.get(dep);
        if (back) {
          back.delete(key);
          if (back.size === 0) state.forward.delete(dep);
        }
      }
      state.reverse.delete(key);
    }
    // Drop range entries whose dependent is `key`.
    for (const [sheet, arr] of state.rangeIndex) {
      const next = arr.filter((entry) => entry.dependentKey !== key);
      if (next.length === 0) state.rangeIndex.delete(sheet);
      else state.rangeIndex.set(sheet, next);
    }
    state.volatile.delete(key);
  }

  function addCell(ref: CellRef, formula: Formula | null, value: Value | null): void {
    const key = cellRefKey(ref);
    dropEdges(key);
    if (formula) {
      state.cells.set(key, { kind: "formula", formula, cachedValue: value ?? blank() });
      for (const dep of formula.dependencies) {
        if ("row" in dep) {
          // CellRef
          const depKey = cellRefKey(dep);
          ensureForward(depKey).add(key);
          ensureReverse(key).add(depKey);
        } else {
          // RangeRef
          ensureRangeBucket(dep.sheet).push({ range: dep, dependentKey: key });
        }
      }
      if (formula.volatile) state.volatile.add(key);
    } else {
      state.cells.set(key, { kind: "value", cachedValue: value ?? blank() });
    }
    state.dirty.add(key);
  }

  function removeCell(ref: CellRef): void {
    const key = cellRefKey(ref);
    dropEdges(key);
    state.cells.delete(key);
    state.dirty.delete(key);
    // Anyone who *depended on* this cell still has a forward edge pointing
    // at `key`. We keep those edges (they survive in `forward[key]`) so
    // that when the cell comes back to life later, the dependent recalcs.
    // The forward set is consulted at markDirty time; an absent cell just
    // means the dependents read `Blank` until the cell is re-added.
    markDownstreamDirty(key, new Set());
  }

  function markDownstreamDirty(key: CellKey, seen: Set<CellKey>): void {
    if (seen.has(key)) return;
    seen.add(key);
    const fwd = state.forward.get(key);
    if (fwd) {
      for (const downstream of fwd) {
        state.dirty.add(downstream);
        markDownstreamDirty(downstream, seen);
      }
    }
    // Also dirty any cell whose range-dependency includes this cell.
    const parsed = parseCellKey(key);
    const bucket = state.rangeIndex.get(parsed.sheet);
    if (bucket) {
      for (const entry of bucket) {
        if (rangeContains(entry.range, parsed.row, parsed.col)) {
          state.dirty.add(entry.dependentKey);
          markDownstreamDirty(entry.dependentKey, seen);
        }
      }
    }
  }

  function markDirty(ref: CellRef): void {
    const key = cellRefKey(ref);
    if (state.dirty.has(key) && hasOnlyDownstreamProcessed(key)) return;
    state.dirty.add(key);
    markDownstreamDirty(key, new Set());
  }

  // Helper to keep the dirty-add idempotent without losing the
  // downstream walk (we always want to expand transitive closure).
  function hasOnlyDownstreamProcessed(_key: CellKey): boolean {
    return false;
  }

  function markRangeDirty(ref: RangeRef): void {
    for (let r = ref.r0; r <= ref.r1; r++) {
      for (let c = ref.c0; c <= ref.c1; c++) {
        markDirty({ sheet: ref.sheet, row: r, col: c, abs: 0 });
      }
    }
  }

  function drainTopological(): {
    order: ReadonlyArray<CellKey>;
    cycles: ReadonlyArray<ReadonlyArray<CellKey>>;
  } {
    // Build the affected set: everything currently dirty plus volatile.
    const affected = new Set<CellKey>();
    for (const k of state.dirty) affected.add(k);
    for (const k of state.volatile) affected.add(k);
    state.dirty.clear();

    // Compute restricted in-degree (only edges within `affected` count).
    const inDeg = new Map<CellKey, number>();
    for (const k of affected) {
      const reads = state.reverse.get(k);
      let n = 0;
      if (reads) {
        for (const dep of reads) if (affected.has(dep)) n++;
      }
      // Range-deps: a cell k that reads a range overlapping any
      // affected cell adds those overlaps to its in-degree.
      const cell = state.cells.get(k);
      if (cell && cell.kind === "formula") {
        for (const dep of cell.formula.dependencies) {
          if ("r0" in dep) {
            for (const a of affected) {
              const parsed = parseCellKey(a);
              if (parsed.sheet === dep.sheet && rangeContains(dep, parsed.row, parsed.col) && a !== k) {
                n++;
              }
            }
          }
        }
      }
      inDeg.set(k, n);
    }

    const ready: CellKey[] = [];
    for (const [k, n] of inDeg) if (n === 0) ready.push(k);
    const order: CellKey[] = [];
    while (ready.length > 0) {
      const k = ready.pop()!;
      order.push(k);
      const fwd = state.forward.get(k);
      if (fwd) {
        for (const d of fwd) {
          if (!affected.has(d)) continue;
          const next = (inDeg.get(d) ?? 0) - 1;
          inDeg.set(d, next);
          if (next === 0) ready.push(d);
        }
      }
      // Range dependents: anyone whose range covers `k`.
      const parsed = parseCellKey(k);
      const bucket = state.rangeIndex.get(parsed.sheet);
      if (bucket) {
        for (const entry of bucket) {
          if (entry.dependentKey === k) continue;
          if (!affected.has(entry.dependentKey)) continue;
          if (rangeContains(entry.range, parsed.row, parsed.col)) {
            const next = (inDeg.get(entry.dependentKey) ?? 0) - 1;
            inDeg.set(entry.dependentKey, next);
            if (next === 0) ready.push(entry.dependentKey);
          }
        }
      }
    }

    let cycles: ReadonlyArray<ReadonlyArray<CellKey>> = [];
    if (order.length < affected.size) {
      const remaining = new Set<CellKey>();
      for (const k of affected) if (!order.includes(k)) remaining.add(k);
      cycles = tarjanSCC(remaining, state);
    }
    return { order, cycles };
  }

  function getCachedValue(ref: CellRef): Value | undefined {
    return state.cells.get(cellRefKey(ref))?.cachedValue;
  }

  function getCell(key: CellKey): AnyCell | undefined {
    return state.cells.get(key);
  }

  function setCachedValue(key: CellKey, value: Value): void {
    const cell = state.cells.get(key);
    if (cell) cell.cachedValue = value;
  }

  function allKeys(): ReadonlyArray<CellKey> {
    return Array.from(state.cells.keys());
  }

  function volatileKeys(): ReadonlySet<CellKey> {
    return state.volatile;
  }

  function collectDependents(ref: CellRef): ReadonlySet<CellKey> {
    const seedKey = cellRefKey(ref);
    const out = new Set<CellKey>();
    function walk(key: CellKey): void {
      const fwd = state.forward.get(key);
      if (fwd) {
        for (const downstream of fwd) {
          if (out.has(downstream)) continue;
          out.add(downstream);
          walk(downstream);
        }
      }
      const parsed = parseCellKey(key);
      const bucket = state.rangeIndex.get(parsed.sheet);
      if (bucket) {
        for (const entry of bucket) {
          if (rangeContains(entry.range, parsed.row, parsed.col)) {
            if (out.has(entry.dependentKey)) continue;
            out.add(entry.dependentKey);
            walk(entry.dependentKey);
          }
        }
      }
    }
    walk(seedKey);
    return out;
  }

  return {
    addCell,
    removeCell,
    markDirty,
    markRangeDirty,
    drainTopological,
    getCachedValue,
    getCell,
    setCachedValue,
    allKeys,
    volatileKeys,
    collectDependents,
    get size() {
      return state.cells.size;
    },
  };
}

// ── Helpers ────────────────────────────────────────────────────────────────

function blank(): Value {
  return { kind: "n", v: 0 };
}

function rangeContains(r: RangeRef, row: number, col: number): boolean {
  return row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1;
}

/**
 * Tarjan's strongly-connected components, restricted to the residual
 * subgraph that survived the Kahn topological pop. Each SCC of size
 * >= 1 with at least one cycle edge is reported as a cycle; isolated
 * single-cell SCCs only count if they self-loop.
 */
function tarjanSCC(set: Set<CellKey>, state: State): ReadonlyArray<ReadonlyArray<CellKey>> {
  const index = new Map<CellKey, number>();
  const lowlink = new Map<CellKey, number>();
  const onStack = new Set<CellKey>();
  const stack: CellKey[] = [];
  const out: CellKey[][] = [];
  let counter = 0;

  function neighbours(k: CellKey): CellKey[] {
    const result: CellKey[] = [];
    const fwd = state.forward.get(k);
    if (fwd) for (const d of fwd) if (set.has(d)) result.push(d);
    const parsed = parseCellKey(k);
    const bucket = state.rangeIndex.get(parsed.sheet);
    if (bucket) {
      for (const entry of bucket) {
        if (entry.dependentKey === k) continue;
        if (!set.has(entry.dependentKey)) continue;
        if (rangeContains(entry.range, parsed.row, parsed.col)) {
          result.push(entry.dependentKey);
        }
      }
    }
    return result;
  }

  function strongconnect(v: CellKey): void {
    index.set(v, counter);
    lowlink.set(v, counter);
    counter++;
    stack.push(v);
    onStack.add(v);
    for (const w of neighbours(v)) {
      if (!index.has(w)) {
        strongconnect(w);
        lowlink.set(v, Math.min(lowlink.get(v)!, lowlink.get(w)!));
      } else if (onStack.has(w)) {
        lowlink.set(v, Math.min(lowlink.get(v)!, index.get(w)!));
      }
    }
    if (lowlink.get(v) === index.get(v)) {
      const scc: CellKey[] = [];
      while (true) {
        const w = stack.pop()!;
        onStack.delete(w);
        scc.push(w);
        if (w === v) break;
      }
      // Only emit if it actually contains a cycle: size > 1 or self-loop.
      if (scc.length > 1 || neighbours(v).includes(v)) out.push(scc);
    }
  }

  for (const v of set) if (!index.has(v)) strongconnect(v);
  return out;
}

/** Convenience helper used by recalc: build a `#REF!` cycle error for an SCC. */
export function cycleError(scc: ReadonlyArray<CellKey>): Value {
  return err(refWithCycle(scc));
}

/** Convenience: parse a CellKey back into a CellRef (abs flag is always 0). */
export function refForKey(key: CellKey): CellRef {
  const { sheet, row, col } = parseCellKey(key);
  return { sheet, row, col, abs: 0 };
}

export { makeCellKey };
