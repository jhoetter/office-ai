import { cellKey } from "../model/refs.js";
import { flattenCellXf } from "../model/style-mutate.js";
import type { StyleColor, StyleTable } from "../model/style-table.js";
import type {
  AutoFilter,
  Cell,
  CellValue,
  CustomFilterOp,
  DynamicFilterType,
  FilterColumn,
  Sheet,
} from "../model/types.js";

/**
 * Recompute the set of hidden rows for a sheet given an AutoFilter.
 *
 * Excel evaluates per-column criteria with AND across columns; an
 * empty `columns` map (or no AutoFilter at all) hides nothing.
 *
 * Only rows in the filter range below the header row are considered.
 * Rows outside the range stay visible (matches Excel — collapsing
 * outside the filter range is what `<row hidden=…>` driven by row
 * hide/group does, which is a separate model surface).
 *
 * Excel's filter does NOT auto-recompute when cells are edited — the
 * user has to "Reapply" — so handlers only call this on filter
 * mutations, not value mutations.
 */
export function recomputeHiddenRows(
  sheet: Sheet,
  styles: StyleTable,
  autoFilter: AutoFilter | undefined
): Set<number> {
  const out = new Set<number>();
  if (!autoFilter) return out;
  const { range, columns } = autoFilter;
  if (columns.size === 0) return out;

  // Pre-compute Top-N thresholds per column.
  const top10Cache = new Map<number, ReadonlySet<number>>();
  for (const [colId, fc] of columns) {
    if (fc.kind !== "top10") continue;
    const absCol = range.c1 + colId;
    const nums: number[] = [];
    for (let r = range.r1 + 1; r <= range.r2; r++) {
      const cell = sheet.cells.get(cellKey(r, absCol));
      const v = cell?.value;
      if (typeof v === "number" && Number.isFinite(v)) nums.push(v);
    }
    nums.sort((a, b) => (fc.top ? b - a : a - b));
    let take: number;
    if (fc.percent) {
      take = Math.max(0, Math.floor((nums.length * fc.n) / 100));
    } else {
      take = Math.max(0, Math.min(nums.length, fc.n));
    }
    const allowed = new Set<number>();
    for (let i = 0; i < take; i++) {
      const v = nums[i];
      if (v !== undefined) allowed.add(v);
    }
    top10Cache.set(colId, allowed);
  }

  for (let r = range.r1 + 1; r <= range.r2; r++) {
    let keep = true;
    for (const [colId, fc] of columns) {
      const absCol = range.c1 + colId;
      const cell = sheet.cells.get(cellKey(r, absCol));
      const matched = matchFilter(fc, cell, colId, styles, top10Cache);
      if (!matched) {
        keep = false;
        break;
      }
    }
    if (!keep) out.add(r);
  }
  return out;
}

function matchFilter(
  fc: FilterColumn,
  cell: Cell | undefined,
  colId: number,
  styles: StyleTable,
  top10Cache: ReadonlyMap<number, ReadonlySet<number>>
): boolean {
  switch (fc.kind) {
    case "values":
      return matchValues(fc.values, fc.blank, cell, styles);
    case "custom":
      return matchCustom(fc, cell);
    case "top10": {
      const allowed = top10Cache.get(colId);
      if (!allowed) return true;
      const v = cell?.value;
      return typeof v === "number" && allowed.has(v);
    }
    case "dynamic":
      return matchDynamic(fc.type, cell);
    case "color":
      return matchColor(fc, cell, styles);
  }
}

function matchValues(
  values: ReadonlySet<string>,
  blank: boolean,
  cell: Cell | undefined,
  styles: StyleTable
): boolean {
  if (cell === undefined || cell.value === null || cell.value === "") return blank;
  const eff = flattenCellXf(styles, cell.styleId);
  const formatted = formatCellValueForFilter(cell.value, eff.numFmtId);
  return values.has(formatted);
}

function matchCustom(fc: Extract<FilterColumn, { kind: "custom" }>, cell: Cell | undefined): boolean {
  const a = matchCustomOp(fc.op1, cell);
  if (!fc.op2) return a;
  return fc.combine === "and" ? a && matchCustomOp(fc.op2, cell) : a || matchCustomOp(fc.op2, cell);
}

function matchCustomOp(op: CustomFilterOp, cell: Cell | undefined): boolean {
  const v = cell?.value ?? null;
  // Numeric comparison path when both sides are numbers.
  const numericOp = op.val !== "" && Number.isFinite(Number(op.val));
  if (numericOp && typeof v === "number") {
    const target = Number(op.val);
    switch (op.operator) {
      case "equal":
        return v === target;
      case "notEqual":
        return v !== target;
      case "greaterThan":
        return v > target;
      case "greaterThanOrEqual":
        return v >= target;
      case "lessThan":
        return v < target;
      case "lessThanOrEqual":
        return v <= target;
    }
  }
  // String comparison path; supports `*` and `?` wildcards for equal /
  // notEqual (matches Excel parity).
  const left = stringifyForCompare(v);
  if (op.operator === "equal" || op.operator === "notEqual") {
    const re = wildcardToRegExp(op.val);
    const matched = re.test(left);
    return op.operator === "equal" ? matched : !matched;
  }
  // Lexicographic compare for greater/less ops on strings.
  const cmp = left.localeCompare(op.val);
  switch (op.operator) {
    case "greaterThan":
      return cmp > 0;
    case "greaterThanOrEqual":
      return cmp >= 0;
    case "lessThan":
      return cmp < 0;
    case "lessThanOrEqual":
      return cmp <= 0;
  }
}

function matchDynamic(type: DynamicFilterType, cell: Cell | undefined): boolean {
  const v = cell?.value;
  if (typeof v !== "number" || !Number.isFinite(v)) return false;
  const date = excelSerialToDate(v);
  if (!date) return false;
  const today = startOfDay(new Date());
  const cellDay = startOfDay(date);

  // Month filters: M1..M12
  const monthMatch = /^M(\d{1,2})$/.exec(type);
  if (monthMatch) {
    const m = Number(monthMatch[1]);
    return cellDay.getMonth() + 1 === m;
  }
  // Quarter filters: Q1..Q4
  const qMatch = /^Q([1-4])$/.exec(type);
  if (qMatch) {
    const q = Number(qMatch[1]);
    const cellQ = Math.floor(cellDay.getMonth() / 3) + 1;
    return cellQ === q;
  }

  switch (type) {
    case "today":
      return sameDay(cellDay, today);
    case "yesterday":
      return sameDay(cellDay, addDays(today, -1));
    case "tomorrow":
      return sameDay(cellDay, addDays(today, 1));
    case "thisWeek":
      return inRange(cellDay, startOfWeek(today), addDays(startOfWeek(today), 7));
    case "lastWeek":
      return inRange(cellDay, addDays(startOfWeek(today), -7), startOfWeek(today));
    case "nextWeek":
      return inRange(cellDay, addDays(startOfWeek(today), 7), addDays(startOfWeek(today), 14));
    case "thisMonth":
      return cellDay.getFullYear() === today.getFullYear() && cellDay.getMonth() === today.getMonth();
    case "lastMonth": {
      const ref = addMonths(today, -1);
      return cellDay.getFullYear() === ref.getFullYear() && cellDay.getMonth() === ref.getMonth();
    }
    case "nextMonth": {
      const ref = addMonths(today, 1);
      return cellDay.getFullYear() === ref.getFullYear() && cellDay.getMonth() === ref.getMonth();
    }
    case "thisQuarter":
      return inSameQuarter(cellDay, today);
    case "lastQuarter":
      return inSameQuarter(cellDay, addMonths(today, -3));
    case "nextQuarter":
      return inSameQuarter(cellDay, addMonths(today, 3));
    case "thisYear":
      return cellDay.getFullYear() === today.getFullYear();
    case "lastYear":
      return cellDay.getFullYear() === today.getFullYear() - 1;
    case "nextYear":
      return cellDay.getFullYear() === today.getFullYear() + 1;
    case "yearToDate":
      return cellDay.getFullYear() === today.getFullYear() && cellDay.getTime() <= today.getTime();
    default:
      return false;
  }
}

function matchColor(
  fc: Extract<FilterColumn, { kind: "color" }>,
  cell: Cell | undefined,
  styles: StyleTable
): boolean {
  if (!cell) return false;
  const eff = flattenCellXf(styles, cell.styleId);
  const color = fc.isCellColor
    ? eff.fill.kind === "pattern"
      ? eff.fill.fgColor
      : undefined
    : eff.font.color;
  return colorArgbEquals(color, fc.argb);
}

function colorArgbEquals(color: StyleColor | undefined, argb: string): boolean {
  if (color === undefined) return false;
  const a = (color.rgb ?? "").toUpperCase();
  const b = argb.toUpperCase();
  if (!a) return false;
  // Allow callers to pass either 6-char RGB or 8-char ARGB. We compare
  // on the trailing 6 hex chars when lengths differ.
  const tailA = a.length === 8 ? a.slice(2) : a;
  const tailB = b.length === 8 ? b.slice(2) : b;
  return tailA === tailB;
}

/**
 * Mirror of the apps/web `formatCellValue` so the value-checklist a
 * filter records and the rendered cell text are spelled the same way.
 * Only the handful of presets exposed by the toolbar are modelled
 * here; anything else falls back to the raw stringification.
 */
export function formatCellValueForFilter(value: CellValue, numFmtId: number): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value;
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "number") return formatNumber(value, numFmtId);
  switch (value.kind) {
    case "error":
      return value.code;
  }
}

function stringifyForCompare(v: CellValue | null): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v;
  if (typeof v === "number") return String(v);
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  switch (v.kind) {
    case "error":
      return v.code;
  }
}

function wildcardToRegExp(pattern: string): RegExp {
  let body = "";
  for (const ch of pattern) {
    if (ch === "*") body += ".*";
    else if (ch === "?") body += ".";
    else body += ch.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }
  return new RegExp(`^${body}$`, "i");
}

function formatNumber(n: number, numFmtId: number): string {
  switch (numFmtId) {
    case 0:
      return String(n);
    case 1:
      return n.toFixed(0);
    case 2:
      return n.toFixed(2);
    case 3:
      return Math.round(n).toLocaleString("en-US");
    case 4:
      return n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    case 9:
      return `${Math.round(n * 100)}%`;
    case 10:
      return `${(n * 100).toFixed(2)}%`;
    case 14:
      return excelSerialToIso(n) ?? String(n);
    default:
      return String(n);
  }
}

function excelSerialToDate(serial: number): Date | null {
  if (!Number.isFinite(serial) || serial < 1) return null;
  const adjusted = serial < 60 ? serial + 1 : serial;
  const ms = (adjusted - 25569) * 86_400_000;
  const d = new Date(ms);
  if (Number.isNaN(d.getTime())) return null;
  return d;
}

function excelSerialToIso(serial: number): string | null {
  const d = excelSerialToDate(serial);
  if (!d) return null;
  const yyyy = d.getUTCFullYear().toString().padStart(4, "0");
  const mm = (d.getUTCMonth() + 1).toString().padStart(2, "0");
  const dd = d.getUTCDate().toString().padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function startOfDay(d: Date): Date {
  const out = new Date(d.getTime());
  out.setHours(0, 0, 0, 0);
  return out;
}

function sameDay(a: Date, b: Date): boolean {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

function addDays(d: Date, n: number): Date {
  const out = new Date(d.getTime());
  out.setDate(out.getDate() + n);
  return out;
}

function addMonths(d: Date, n: number): Date {
  const out = new Date(d.getTime());
  out.setMonth(out.getMonth() + n);
  return out;
}

function inRange(d: Date, startInclusive: Date, endExclusive: Date): boolean {
  const t = d.getTime();
  return t >= startInclusive.getTime() && t < endExclusive.getTime();
}

function startOfWeek(d: Date): Date {
  const out = startOfDay(d);
  // Excel's "this week" uses Sunday as the first day of the week.
  const dow = out.getDay();
  out.setDate(out.getDate() - dow);
  return out;
}

function inSameQuarter(a: Date, b: Date): boolean {
  if (a.getFullYear() !== b.getFullYear()) return false;
  return Math.floor(a.getMonth() / 3) === Math.floor(b.getMonth() / 3);
}
