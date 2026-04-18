import { CommandError, type CommandHandler, type DiffChange } from "@officeai/core";
import { cellKey, formatA1 } from "../model/refs.js";
import type { Cell, Sheet, XlsxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { TextToColumnsPayload } from "./payloads.js";
import { parseCellRef, parseRangeRef, resolveSheet } from "./validation.js";

/**
 * `xlsx:text-to-columns` — Excel's Data → Text to Columns.
 *
 * Walks each row of `range`, takes the value of the left-most cell
 * (the "primary" column), and splits it on `delimiter`. The split
 * pieces are written across the row starting at `destination`
 * (default: the source cell's own column).
 *
 * Spec: `spec/xlsx/agent-commands.md` §16.
 *
 * Conventions:
 *   - Source values that are not strings are coerced via `String(...)`.
 *   - Empty source rows are no-ops (no clearing).
 *   - `treatConsecutiveAsOne: true` collapses runs of the delimiter
 *     (matches the "Treat consecutive delimiters as one" toggle).
 *   - Each split piece is parsed back into a typed value:
 *       - `"true" / "false"` → boolean
 *       - numeric literal → number
 *       - everything else → string
 *   - The handler clears any leftover cells past the new split width
 *     in the destination strip (so re-running with fewer pieces does
 *     not leave stale data behind).
 */
export const textToColumnsHandler: CommandHandler<TextToColumnsPayload, XlsxSnapshot> = {
  type: "xlsx:text-to-columns",
  apply(snapshot, payload) {
    if (typeof payload.delimiter !== "string" || payload.delimiter.length === 0) {
      throw new CommandError("invalid-payload", "text-to-columns requires a non-empty delimiter");
    }
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const source = parseRangeRef(payload.range);
    const destination = payload.destination
      ? parseCellRef(payload.destination)
      : { row: source.start.row, col: source.start.col };

    const rows = source.end.row - source.start.row + 1;
    const sourceCol = source.start.col;

    const cells = new Map(sheet.cells);
    const changes: DiffChange[] = [];

    let maxPieces = 0;
    const splitRows: string[][] = [];
    for (let dr = 0; dr < rows; dr++) {
      const r = source.start.row + dr;
      const cell = cells.get(cellKey(r, sourceCol));
      const raw = cell?.value;
      if (raw === null || raw === undefined) {
        splitRows.push([]);
        continue;
      }
      const text = typeof raw === "string" ? raw : String(typeof raw === "object" && raw && "kind" in raw ? raw.code : raw);
      const pieces = splitText(text, payload.delimiter, payload.treatConsecutiveAsOne === true);
      splitRows.push(pieces);
      if (pieces.length > maxPieces) maxPieces = pieces.length;
    }

    if (destination.row + rows > 1048576 || destination.col + maxPieces > 16384) {
      throw new CommandError(
        "out-of-bounds",
        `text-to-columns of ${rows}×${maxPieces} at ${formatA1(destination)} extends past the sheet edge`
      );
    }

    // Determine width of the strip we may need to clear (max of any
    // existing data in the destination row that previously came from
    // a wider split).
    let clearWidth = maxPieces;
    for (let dr = 0; dr < rows; dr++) {
      const r = destination.row + dr;
      // Cap the look-ahead so we never run past the sheet.
      const probeMax = Math.min(16384 - destination.col, 1024);
      for (let dc = clearWidth; dc < probeMax; dc++) {
        if (!cells.has(cellKey(r, destination.col + dc))) break;
        clearWidth = dc + 1;
      }
    }

    for (let dr = 0; dr < rows; dr++) {
      const pieces = splitRows[dr]!;
      const r = destination.row + dr;
      for (let dc = 0; dc < clearWidth; dc++) {
        const c = destination.col + dc;
        const key = cellKey(r, c);
        const before = cells.get(key);
        if (dc < pieces.length) {
          const value = coerceLiteral(pieces[dc]!);
          const next: Cell = {
            row: r,
            col: c,
            value,
            ...(before?.styleId !== undefined ? { styleId: before.styleId } : {}),
          };
          if (cellsEqual(before, next)) continue;
          cells.set(key, next);
          changes.push(diffChange(sheet, r, c, before?.value ?? null, next.value));
        } else if (before) {
          // Beyond the split count → clear.
          if (before.formula || before.value !== null) {
            cells.delete(key);
            changes.push(diffChange(sheet, r, c, before.value ?? null, null));
          }
        }
      }
    }

    if (changes.length === 0) {
      const next = evolveSnapshot(snapshot, snapshot.root, {});
      return {
        next,
        diff: buildDiff(snapshot.revision, next.revision, [
          {
            kind: "node-updated",
            nodeId: sheet.id,
            path: ["sheets", sheet.index],
            field: "noop",
            summary: `text-to-columns at ${formatA1(destination)}: no changes`,
          },
        ]),
      };
    }

    const nextSheet: Sheet = { ...sheet, cells };
    const nextWorkbook = replaceSheet(snapshot.root, nextSheet);
    const next = evolveSnapshot(snapshot, nextWorkbook, { sheets: [sheet.partPath] });
    return { next, diff: buildDiff(snapshot.revision, next.revision, changes) };
  },
};

/**
 * Split `text` on `delimiter`. With `consecutiveAsOne`, runs of the
 * delimiter are collapsed into a single split point. Multi-char
 * delimiters are supported. We do not honour quoted segments here —
 * external CSV-style quoting belongs in `delimitedToSnapshot`.
 */
function splitText(text: string, delimiter: string, consecutiveAsOne: boolean): string[] {
  if (text === "") return [];
  const parts = text.split(delimiter);
  if (!consecutiveAsOne) return parts;
  // Treat any run of the delimiter as a single delimiter. That drops
  // every empty field, including ones at the start/end caused by
  // leading/trailing delimiters (Excel's "Treat consecutive
  // delimiters as one" behaviour).
  return parts.filter((p) => p !== "");
}

function coerceLiteral(raw: string): Cell["value"] {
  const t = raw.trim();
  if (t === "") return raw === "" ? null : raw;
  if (/^-?\d+(?:\.\d+)?$/.test(t)) {
    const n = Number(t);
    if (Number.isFinite(n)) return n;
  }
  const lower = t.toLowerCase();
  if (lower === "true") return true;
  if (lower === "false") return false;
  return raw;
}

function cellsEqual(a: Cell | undefined, b: Cell): boolean {
  if (!a) return false;
  if ((a.styleId ?? -1) !== (b.styleId ?? -1)) return false;
  if ((a.formula?.text ?? null) !== (b.formula?.text ?? null)) return false;
  return cellValuesEqual(a.value, b.value);
}

function cellValuesEqual(a: Cell["value"], b: Cell["value"]): boolean {
  if (a === b) return true;
  if (a === null || b === null) return false;
  if (typeof a === "object" && typeof b === "object" && "kind" in a && "kind" in b) {
    return a.kind === b.kind && a.code === b.code;
  }
  return false;
}

function diffChange(
  sheet: Sheet,
  row: number,
  col: number,
  beforeVal: Cell["value"],
  afterVal: Cell["value"]
): DiffChange {
  return {
    kind: "node-updated",
    nodeId: sheet.id,
    path: ["sheets", sheet.index, "cells", `${sheet.name}!${formatA1({ row, col })}`],
    field: "value",
    summary: `${formatA1({ row, col })}: ${formatVal(beforeVal)} → ${formatVal(afterVal)}`,
    meta: { before: { value: beforeVal }, after: { value: afterVal } },
  };
}

function formatVal(v: Cell["value"]): string {
  if (v === null || v === undefined) return "∅";
  if (typeof v === "string") return JSON.stringify(v);
  if (typeof v === "object" && v && "kind" in v) return v.code;
  return String(v);
}
