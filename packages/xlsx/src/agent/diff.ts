import type { DiffChange, DocumentDiff } from "@officeai/core";
import { formatA1 } from "../model/refs.js";
import type { Cell, MergedCell, Sheet, XlsxSnapshot } from "../model/types.js";

/**
 * Structural diff between two XLSX snapshots.
 *
 * The diff is at sheet + cell + merge granularity. Sheets are matched
 * by stable sheet id; cells by `${row}:${col}` within their sheet;
 * merges by exact rectangle. Anything finer (style table, comments,
 * conditional formatting) lands once those models are typed in
 * Phase 7+.
 *
 * Spec contract: `spec/shared/agent-api.md` §getDiff.
 */
export function diffXlsxSnapshots(from: XlsxSnapshot, to: XlsxSnapshot): DocumentDiff {
  const changes: DiffChange[] = [];

  const fromById = new Map(from.root.sheets.map((s) => [s.sheetId, s]));
  const toById = new Map(to.root.sheets.map((s) => [s.sheetId, s]));

  for (const [id, beforeSheet] of fromById) {
    const afterSheet = toById.get(id);
    if (!afterSheet) {
      changes.push({
        kind: "node-deleted",
        nodeId: beforeSheet.id,
        path: ["sheets", beforeSheet.index, beforeSheet.name],
        summary: `removed sheet "${beforeSheet.name}"`,
      });
      continue;
    }
    diffSheet(beforeSheet, afterSheet, changes);
  }

  for (const [id, afterSheet] of toById) {
    if (fromById.has(id)) continue;
    changes.push({
      kind: "node-inserted",
      nodeId: afterSheet.id,
      path: ["sheets", afterSheet.index, afterSheet.name],
      summary: `added sheet "${afterSheet.name}"`,
    });
  }

  return {
    format: "xlsx",
    fromRevision: from.revision,
    toRevision: to.revision,
    changes,
  };
}

function diffSheet(before: Sheet, after: Sheet, out: DiffChange[]): void {
  if (before.name !== after.name) {
    out.push({
      kind: "node-updated",
      nodeId: after.id,
      path: ["sheets", after.index, "name"],
      field: "name",
      summary: `renamed "${before.name}" → "${after.name}"`,
    });
  }
  if (before.index !== after.index) {
    out.push({
      kind: "node-moved",
      nodeId: after.id,
      from: ["sheets", before.index],
      to: ["sheets", after.index],
      summary: `moved "${after.name}" from index ${before.index} to ${after.index}`,
    });
  }

  for (const [key, beforeCell] of before.cells) {
    const afterCell = after.cells.get(key);
    if (!afterCell) {
      out.push(cellDiff("node-deleted", after, beforeCell, undefined));
      continue;
    }
    if (!cellsEqual(beforeCell, afterCell)) {
      out.push(cellDiff("node-updated", after, beforeCell, afterCell));
    }
  }
  for (const [key, afterCell] of after.cells) {
    if (before.cells.has(key)) continue;
    out.push(cellDiff("node-inserted", after, undefined, afterCell));
  }

  const beforeMerges = new Set(before.merges.map(mergeKey));
  const afterMerges = new Set(after.merges.map(mergeKey));
  for (const m of before.merges) {
    if (!afterMerges.has(mergeKey(m))) {
      out.push({
        kind: "node-deleted",
        nodeId: after.id,
        path: ["sheets", after.index, "merges", mergeKey(m)],
        summary: `removed merge ${mergeKey(m)}`,
      });
    }
  }
  for (const m of after.merges) {
    if (!beforeMerges.has(mergeKey(m))) {
      out.push({
        kind: "node-inserted",
        nodeId: after.id,
        path: ["sheets", after.index, "merges", mergeKey(m)],
        summary: `added merge ${mergeKey(m)}`,
      });
    }
  }
}

function cellDiff(
  kind: "node-inserted" | "node-deleted" | "node-updated",
  sheet: Sheet,
  before: Cell | undefined,
  after: Cell | undefined
): DiffChange {
  const ref = formatA1({
    row: (after ?? before)!.row,
    col: (after ?? before)!.col,
  });
  const path: ReadonlyArray<string | number> = ["sheets", sheet.index, "cells", ref];
  if (kind === "node-updated") {
    return {
      kind: "node-updated",
      nodeId: sheet.id,
      path,
      field: "value",
      summary: `${ref}: ${describe(before?.value ?? null)} → ${describe(after?.value ?? null)}`,
    };
  }
  if (kind === "node-inserted") {
    return {
      kind: "node-inserted",
      nodeId: sheet.id,
      path,
      summary: `${ref} ← ${describe(after?.value ?? null)}`,
    };
  }
  return {
    kind: "node-deleted",
    nodeId: sheet.id,
    path,
    summary: `${ref}: cleared (was ${describe(before?.value ?? null)})`,
  };
}

function cellsEqual(a: Cell, b: Cell): boolean {
  if ((a.formula?.text ?? null) !== (b.formula?.text ?? null)) return false;
  return cellValueEqual(a.value, b.value);
}

function cellValueEqual(a: Cell["value"], b: Cell["value"]): boolean {
  if (a === b) return true;
  if (a === null || b === null) return false;
  if (typeof a === "object" && typeof b === "object") {
    if ("kind" in a && "kind" in b) return a.code === b.code;
  }
  return false;
}

function mergeKey(m: MergedCell): string {
  return `${formatA1({ row: m.r1, col: m.c1 })}:${formatA1({ row: m.r2, col: m.c2 })}`;
}

function describe(v: Cell["value"]): string {
  if (v === null) return "∅";
  if (typeof v === "string") return JSON.stringify(v);
  if (typeof v === "object" && "kind" in v) return v.code;
  return String(v);
}
