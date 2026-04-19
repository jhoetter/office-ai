import type { Cell, ConditionalFormat, ConditionalFormatOverlay, Sheet } from "../model/types.js";
import { parseRange } from "../model/refs.js";

/**
 * C10 — Render-time evaluator for conditional-formatting rules.
 *
 * Walks every typed rule on a sheet and returns a sparse map of
 * cell-key → overlay describing the visual override. Rules layer
 * earliest-first, so later rules win conflicts (Excel parity).
 *
 * The evaluator is deliberately format-agnostic: it works off the
 * typed cell value (number, string, boolean, error, null) and
 * never touches the OOXML wire format. Style application is the
 * Grid's job — this module just decides what overlay to apply.
 */

export interface CellRef {
  readonly row: number;
  readonly col: number;
}

export type OverlayMap = ReadonlyMap<string, ConditionalFormatOverlay>;

function cellKey(r: number, c: number): string {
  return `${r}:${c}`;
}

interface ParsedArea {
  readonly r1: number;
  readonly c1: number;
  readonly r2: number;
  readonly c2: number;
}

function parseAreas(range: string): ReadonlyArray<ParsedArea> {
  const out: ParsedArea[] = [];
  for (const piece of range.split(",")) {
    const trimmed = piece.trim();
    if (!trimmed) continue;
    try {
      const r = parseRange(trimmed);
      out.push({
        r1: Math.min(r.start.row, r.end.row),
        c1: Math.min(r.start.col, r.end.col),
        r2: Math.max(r.start.row, r.end.row),
        c2: Math.max(r.start.col, r.end.col),
      });
    } catch {
      // Ignore malformed range fragments — the editor validates
      // before dispatch, so this only triggers on hand-edited
      // model state.
    }
  }
  return out;
}

function* iterateArea(areas: ReadonlyArray<ParsedArea>): Generator<CellRef> {
  for (const a of areas) {
    for (let r = a.r1; r <= a.r2; r++) {
      for (let c = a.c1; c <= a.c2; c++) {
        yield { row: r, col: c };
      }
    }
  }
}

function asNumber(v: Cell | undefined): number | null {
  if (!v) return null;
  if (typeof v.value === "number") return v.value;
  if (typeof v.value === "boolean") return v.value ? 1 : 0;
  if (typeof v.value === "string") {
    const n = Number(v.value);
    return Number.isFinite(n) ? n : null;
  }
  return null;
}

function asString(v: Cell | undefined): string | null {
  if (!v || v.value === null || v.value === undefined) return null;
  if (typeof v.value === "string") return v.value;
  if (typeof v.value === "number" || typeof v.value === "boolean") return String(v.value);
  return null;
}

function setOverlay(
  out: Map<string, ConditionalFormatOverlay>,
  key: string,
  overlay: ConditionalFormatOverlay
): void {
  const existing = out.get(key);
  out.set(key, existing ? { ...existing, ...overlay } : overlay);
}

function evalCellIs(
  rule: Extract<ConditionalFormat, { kind: "cellIs" }>,
  cells: ReadonlyMap<string, Cell>,
  out: Map<string, ConditionalFormatOverlay>
): void {
  const areas = parseAreas(rule.range);
  for (const ref of iterateArea(areas)) {
    const cell = cells.get(cellKey(ref.row, ref.col));
    const v = asNumber(cell);
    if (v === null) continue;
    let match = false;
    switch (rule.op) {
      case "gt":
        match = v > rule.value;
        break;
      case "ge":
        match = v >= rule.value;
        break;
      case "lt":
        match = v < rule.value;
        break;
      case "le":
        match = v <= rule.value;
        break;
      case "eq":
        match = v === rule.value;
        break;
      case "ne":
        match = v !== rule.value;
        break;
      case "between":
        match = rule.value2 !== undefined && v >= rule.value && v <= rule.value2;
        break;
      case "notBetween":
        match = rule.value2 !== undefined && (v < rule.value || v > rule.value2);
        break;
      default: {
        const _exhaustive: never = rule.op;
        void _exhaustive;
      }
    }
    if (match) setOverlay(out, cellKey(ref.row, ref.col), rule.overlay);
  }
}

function evalContainsText(
  rule: Extract<ConditionalFormat, { kind: "containsText" }>,
  cells: ReadonlyMap<string, Cell>,
  out: Map<string, ConditionalFormatOverlay>
): void {
  const areas = parseAreas(rule.range);
  const needle = rule.text.toLowerCase();
  for (const ref of iterateArea(areas)) {
    const cell = cells.get(cellKey(ref.row, ref.col));
    const s = asString(cell);
    if (s === null) continue;
    const has = s.toLowerCase().includes(needle);
    const match = rule.contains ? has : !has;
    if (match) setOverlay(out, cellKey(ref.row, ref.col), rule.overlay);
  }
}

function evalTop10(
  rule: Extract<ConditionalFormat, { kind: "top10" }>,
  cells: ReadonlyMap<string, Cell>,
  out: Map<string, ConditionalFormatOverlay>
): void {
  const areas = parseAreas(rule.range);
  const values: { ref: CellRef; v: number }[] = [];
  for (const ref of iterateArea(areas)) {
    const cell = cells.get(cellKey(ref.row, ref.col));
    const v = asNumber(cell);
    if (v === null) continue;
    values.push({ ref, v });
  }
  if (values.length === 0) return;
  values.sort((a, b) => (rule.bottom ? a.v - b.v : b.v - a.v));
  let cutoff = rule.rank;
  if (rule.percent) {
    cutoff = Math.max(1, Math.floor((values.length * rule.rank) / 100));
  }
  cutoff = Math.min(cutoff, values.length);
  for (let i = 0; i < cutoff; i++) {
    setOverlay(out, cellKey(values[i].ref.row, values[i].ref.col), rule.overlay);
  }
}

function evalDuplicate(
  rule: Extract<ConditionalFormat, { kind: "duplicate" }>,
  cells: ReadonlyMap<string, Cell>,
  out: Map<string, ConditionalFormatOverlay>
): void {
  const areas = parseAreas(rule.range);
  const counts = new Map<string, number>();
  const refsByKey = new Map<string, CellRef[]>();
  for (const ref of iterateArea(areas)) {
    const cell = cells.get(cellKey(ref.row, ref.col));
    if (!cell || cell.value === null || cell.value === undefined) continue;
    const k = String(cell.value);
    counts.set(k, (counts.get(k) ?? 0) + 1);
    const arr = refsByKey.get(k) ?? [];
    arr.push(ref);
    refsByKey.set(k, arr);
  }
  for (const [k, n] of counts.entries()) {
    const isDup = n > 1;
    const match = rule.unique ? !isDup : isDup;
    if (!match) continue;
    for (const ref of refsByKey.get(k) ?? []) {
      setOverlay(out, cellKey(ref.row, ref.col), rule.overlay);
    }
  }
}

function lerpColor(a: string, b: string, t: number): string {
  const ar = parseInt(a.slice(0, 2), 16);
  const ag = parseInt(a.slice(2, 4), 16);
  const ab = parseInt(a.slice(4, 6), 16);
  const br = parseInt(b.slice(0, 2), 16);
  const bg = parseInt(b.slice(2, 4), 16);
  const bb = parseInt(b.slice(4, 6), 16);
  const r = Math.round(ar + (br - ar) * t);
  const g = Math.round(ag + (bg - ag) * t);
  const bl = Math.round(ab + (bb - ab) * t);
  const hex = (n: number): string => n.toString(16).padStart(2, "0").toUpperCase();
  return `${hex(r)}${hex(g)}${hex(bl)}`;
}

function evalColorScale(
  rule: Extract<ConditionalFormat, { kind: "colorScale" }>,
  cells: ReadonlyMap<string, Cell>,
  out: Map<string, ConditionalFormatOverlay>
): void {
  const areas = parseAreas(rule.range);
  const values: { ref: CellRef; v: number }[] = [];
  for (const ref of iterateArea(areas)) {
    const cell = cells.get(cellKey(ref.row, ref.col));
    const v = asNumber(cell);
    if (v === null) continue;
    values.push({ ref, v });
  }
  if (values.length === 0) return;
  let lo = values[0].v;
  let hi = values[0].v;
  for (const e of values) {
    if (e.v < lo) lo = e.v;
    if (e.v > hi) hi = e.v;
  }
  const span = hi - lo;
  for (const e of values) {
    const t = span === 0 ? 0.5 : (e.v - lo) / span;
    let fill: string;
    if (rule.midColor !== undefined) {
      fill =
        t < 0.5
          ? lerpColor(rule.minColor, rule.midColor, t * 2)
          : lerpColor(rule.midColor, rule.maxColor, (t - 0.5) * 2);
    } else {
      fill = lerpColor(rule.minColor, rule.maxColor, t);
    }
    setOverlay(out, cellKey(e.ref.row, e.ref.col), { fill });
  }
}

function evalDataBar(
  rule: Extract<ConditionalFormat, { kind: "dataBar" }>,
  cells: ReadonlyMap<string, Cell>,
  out: Map<string, ConditionalFormatOverlay>
): void {
  const areas = parseAreas(rule.range);
  const values: { ref: CellRef; v: number }[] = [];
  for (const ref of iterateArea(areas)) {
    const cell = cells.get(cellKey(ref.row, ref.col));
    const v = asNumber(cell);
    if (v === null) continue;
    values.push({ ref, v });
  }
  if (values.length === 0) return;
  let lo = Math.min(0, values[0].v);
  let hi = values[0].v;
  for (const e of values) {
    if (e.v < lo) lo = Math.min(lo, e.v);
    if (e.v > hi) hi = e.v;
  }
  const span = hi - lo;
  for (const e of values) {
    const fraction = span === 0 ? 1 : (e.v - lo) / span;
    setOverlay(out, cellKey(e.ref.row, e.ref.col), {
      barColor: rule.color,
      barFraction: Math.max(0, Math.min(1, fraction)),
    });
  }
}

/**
 * Evaluate every typed CF rule on a sheet and return the resulting
 * per-cell overlay map.
 */
export function evaluateConditionalFormats(sheet: Sheet): OverlayMap {
  if (sheet.conditionalFormats.length === 0) return new Map();
  const out = new Map<string, ConditionalFormatOverlay>();
  for (const rule of sheet.conditionalFormats) {
    switch (rule.kind) {
      case "cellIs":
        evalCellIs(rule, sheet.cells, out);
        break;
      case "containsText":
        evalContainsText(rule, sheet.cells, out);
        break;
      case "top10":
        evalTop10(rule, sheet.cells, out);
        break;
      case "duplicate":
        evalDuplicate(rule, sheet.cells, out);
        break;
      case "colorScale":
        evalColorScale(rule, sheet.cells, out);
        break;
      case "dataBar":
        evalDataBar(rule, sheet.cells, out);
        break;
      default: {
        const _exhaustive: never = rule;
        void _exhaustive;
      }
    }
  }
  return out;
}
