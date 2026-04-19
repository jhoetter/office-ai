/**
 * C1 — Real grid bounds with lazy index-backed dimensions.
 *
 * Excel's worksheet bounds are 1,048,576 rows × 16,384 columns. We
 * cannot precompute a length-1M prefix-sum array on every render
 * (≈10 ms of GC + alloc per change). Instead, we observe that
 * column / row sizes are uniform unless the user explicitly
 * resized something. The sparse override maps on `Sheet`
 * (`columnWidths`, `rowHeights`) typically carry < 100 entries.
 *
 * Strategy:
 *   1. Build a sorted array of override indices.
 *   2. Build a parallel array of cumulative deltas (sum of
 *      `override - default` up to and including each entry).
 *   3. `xAt(c)` = `c * defaultWidth + cumulativeDelta(c)` →
 *      O(log n) binary search where `n` = number of overrides.
 *   4. `colAtX(x)` reverses the same math via binary search over a
 *      derived "X at each override" array.
 *
 * Both forward and inverse lookups stay below 20 µs even for the
 * full 1M-row sheet, which keeps scrolling and selection
 * interactions cheap.
 */

export interface GridDims {
  readonly totalRows: number;
  readonly totalCols: number;
  readonly defaultColWidth: number;
  readonly defaultRowHeight: number;
  readonly totalWidth: number;
  readonly totalHeight: number;
  /** Pixel X at the left edge of column `c` (`0 ≤ c ≤ totalCols`). */
  xAt(c: number): number;
  /** Pixel Y at the top edge of row `r` (`0 ≤ r ≤ totalRows`). */
  yAt(r: number): number;
  /** Width of column `c`. */
  colWidth(c: number): number;
  /** Height of row `r`. */
  rowHeight(r: number): number;
  /** Reverse lookup: column index whose left edge is the largest ≤ `x`. */
  colAtX(x: number): number;
  /** Reverse lookup: row index whose top edge is the largest ≤ `y`. */
  rowAtY(y: number): number;
}

interface AxisIndex {
  readonly defaultSize: number;
  readonly total: number;
  /** Sorted override indices. */
  readonly idxs: ReadonlyArray<number>;
  /** Override sizes parallel to `idxs`. */
  readonly sizes: ReadonlyArray<number>;
  /** Cumulative `(size - defaultSize)` up to and including each idx. */
  readonly cumDelta: ReadonlyArray<number>;
  /** Pixel position at each override's left/top edge (cached). */
  readonly cumPos: ReadonlyArray<number>;
}

function buildAxisIndex(
  overrides: ReadonlyMap<number, number>,
  defaultSize: number,
  total: number
): AxisIndex {
  if (overrides.size === 0) {
    return {
      defaultSize,
      total,
      idxs: [],
      sizes: [],
      cumDelta: [],
      cumPos: [],
    };
  }
  const idxs: number[] = [];
  for (const k of overrides.keys()) {
    if (k >= 0 && k < total) idxs.push(k);
  }
  idxs.sort((a, b) => a - b);
  const sizes: number[] = new Array(idxs.length);
  const cumDelta: number[] = new Array(idxs.length);
  const cumPos: number[] = new Array(idxs.length);
  let runningDelta = 0;
  for (let i = 0; i < idxs.length; i++) {
    const idx = idxs[i]!;
    const size = Math.max(0, overrides.get(idx) ?? defaultSize);
    sizes[i] = size;
    runningDelta += size - defaultSize;
    cumDelta[i] = runningDelta;
    // Pixel position at the *left/top* edge of this override row/col:
    // base position from default + accumulated delta from prior overrides.
    cumPos[i] = idx * defaultSize + (i === 0 ? 0 : cumDelta[i - 1]!);
  }
  return { defaultSize, total, idxs, sizes, cumDelta, cumPos };
}

/** Largest i in `arr` with `arr[i] <= target`, or -1 if none. */
function upperBoundLE(arr: ReadonlyArray<number>, target: number): number {
  let lo = 0;
  let hi = arr.length - 1;
  let ans = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1;
    if (arr[mid]! <= target) {
      ans = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  return ans;
}

function axisPosAt(axis: AxisIndex, idx: number): number {
  if (idx <= 0) return 0;
  if (idx >= axis.total) {
    return axis.total * axis.defaultSize + (axis.cumDelta[axis.cumDelta.length - 1] ?? 0);
  }
  if (axis.idxs.length === 0) return idx * axis.defaultSize;
  // Find largest override < idx (strict): then base = idx * defaultSize +
  // cumulative delta accumulated through that override. The override at
  // position `idx` itself does NOT contribute (we want the *left* edge).
  const i = upperBoundLE(axis.idxs, idx - 1);
  const baseDelta = i >= 0 ? axis.cumDelta[i]! : 0;
  return idx * axis.defaultSize + baseDelta;
}

function axisSizeAt(axis: AxisIndex, idx: number): number {
  if (idx < 0 || idx >= axis.total) return axis.defaultSize;
  if (axis.idxs.length === 0) return axis.defaultSize;
  const i = upperBoundLE(axis.idxs, idx);
  if (i >= 0 && axis.idxs[i] === idx) return axis.sizes[i]!;
  return axis.defaultSize;
}

function axisIndexAtPos(axis: AxisIndex, pos: number): number {
  if (pos <= 0) return 0;
  if (axis.idxs.length === 0) {
    const idx = Math.floor(pos / axis.defaultSize);
    return Math.max(0, Math.min(axis.total - 1, idx));
  }
  // Find override whose pixel position is the largest ≤ `pos`.
  const i = upperBoundLE(axis.cumPos, pos);
  const baseIdx = i >= 0 ? axis.idxs[i]! : 0;
  const basePos = i >= 0 ? axis.cumPos[i]! : 0;
  // From `baseIdx` onward, the next override's position tells us how
  // far the uniform stretch runs. If `pos` lands *inside* the override
  // row itself we return `baseIdx`; otherwise we step uniformly from
  // `baseIdx + 1` using `defaultSize`.
  if (i >= 0) {
    const overrideEnd = basePos + axis.sizes[i]!;
    if (pos < overrideEnd) return baseIdx;
    const past = pos - overrideEnd;
    const stepped = baseIdx + 1 + Math.floor(past / axis.defaultSize);
    return Math.max(0, Math.min(axis.total - 1, stepped));
  }
  const idx = Math.floor(pos / axis.defaultSize);
  return Math.max(0, Math.min(axis.total - 1, idx));
}

export interface BuildGridDimsArgs {
  readonly columnWidths: ReadonlyMap<number, number>;
  readonly rowHeights: ReadonlyMap<number, number>;
  readonly defaultColWidth: number;
  readonly defaultRowHeight: number;
  readonly totalRows: number;
  readonly totalCols: number;
}

export function buildGridDims(args: BuildGridDimsArgs): GridDims {
  const colAxis = buildAxisIndex(args.columnWidths, args.defaultColWidth, args.totalCols);
  const rowAxis = buildAxisIndex(args.rowHeights, args.defaultRowHeight, args.totalRows);
  const totalWidth = axisPosAt(colAxis, args.totalCols);
  const totalHeight = axisPosAt(rowAxis, args.totalRows);
  return {
    totalRows: args.totalRows,
    totalCols: args.totalCols,
    defaultColWidth: args.defaultColWidth,
    defaultRowHeight: args.defaultRowHeight,
    totalWidth,
    totalHeight,
    xAt: (c) => axisPosAt(colAxis, c),
    yAt: (r) => axisPosAt(rowAxis, r),
    colWidth: (c) => axisSizeAt(colAxis, c),
    rowHeight: (r) => axisSizeAt(rowAxis, r),
    colAtX: (x) => axisIndexAtPos(colAxis, x),
    rowAtY: (y) => axisIndexAtPos(rowAxis, y),
  };
}

/**
 * Array-like view over a {@link GridDims} axis. Exposes `obj[i]`
 * (delegating to `xAt(i)` / `yAt(i)`) plus a `length` so existing
 * code that treated the prefix-sum result as a `ReadonlyArray<number>`
 * continues to work without a sweeping refactor.
 *
 * The returned value is intentionally **not** a real array — iterating
 * with `forEach` / `map` etc. is not supported. Code that needs full
 * iteration should use `dims.totalRows` / `dims.totalCols` and the
 * direct `xAt` / `yAt` getters instead.
 */
export type AxisLookup = {
  readonly length: number;
  readonly [index: number]: number;
};

export function colXsView(dims: GridDims): AxisLookup {
  const length = dims.totalCols + 1;
  return new Proxy(
    { length },
    {
      get(target, prop) {
        if (prop === "length") return length;
        if (typeof prop === "string") {
          const n = Number(prop);
          if (Number.isInteger(n) && n >= 0 && n <= dims.totalCols) {
            return dims.xAt(n);
          }
        }
        return undefined;
      },
    }
  ) as unknown as AxisLookup;
}

export function rowYsView(dims: GridDims): AxisLookup {
  const length = dims.totalRows + 1;
  return new Proxy(
    { length },
    {
      get(target, prop) {
        if (prop === "length") return length;
        if (typeof prop === "string") {
          const n = Number(prop);
          if (Number.isInteger(n) && n >= 0 && n <= dims.totalRows) {
            return dims.yAt(n);
          }
        }
        return undefined;
      },
    }
  ) as unknown as AxisLookup;
}

/**
 * C1 — Index-backed used range. Iterates the sparse `cells` map once
 * to derive `[r1, c1, r2, c2]` bounds inclusive. Excel's `Cmd+End`
 * lands on this rectangle's bottom-right corner; navigation
 * shortcuts use it to jump to the data edge instead of stepping
 * through a million empty rows.
 */
export interface UsedRange {
  readonly r1: number;
  readonly c1: number;
  readonly r2: number;
  readonly c2: number;
}

export function computeUsedRange(cells: ReadonlyMap<string, unknown>): UsedRange | null {
  if (cells.size === 0) return null;
  let r1 = Infinity;
  let c1 = Infinity;
  let r2 = -Infinity;
  let c2 = -Infinity;
  for (const key of cells.keys()) {
    const sep = key.indexOf(":");
    if (sep < 0) continue;
    const r = Number.parseInt(key.slice(0, sep), 10);
    const c = Number.parseInt(key.slice(sep + 1), 10);
    if (!Number.isFinite(r) || !Number.isFinite(c)) continue;
    if (r < r1) r1 = r;
    if (c < c1) c1 = c;
    if (r > r2) r2 = r;
    if (c > c2) c2 = c;
  }
  if (!Number.isFinite(r1) || !Number.isFinite(c1)) return null;
  return { r1, c1, r2, c2 };
}
