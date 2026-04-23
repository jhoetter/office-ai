import { CommandError, type CommandHandler } from "@officeai/core";
import type { DefinedName, XlsxSnapshot, XlsxWorkbook } from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import type { SetPrintAreaPayload, SetPrintTitlesPayload } from "./payloads.js";
import { resolveSheet, parseRangeRef, parseCellRef } from "./validation.js";

const PRINT_AREA = "_xlnm.Print_Area";
const PRINT_TITLES = "_xlnm.Print_Titles";

/**
 * `xlsx:set-print-area` — write / clear the per-sheet
 * `_xlnm.Print_Area` defined name. Mirrors Page Layout → Print Area.
 *
 * The defined name is sheet-scoped; Excel writes it with a
 * `localSheetId` matching the worksheet's 0-based index. The
 * `refersTo` expression is `'<Sheet>'!$A$1:$D$20` (sheet name quoted,
 * absolute references). Pass `clear: true` to remove the entry.
 */
export const setPrintAreaHandler: CommandHandler<SetPrintAreaPayload, XlsxSnapshot> = {
  type: "xlsx:set-print-area",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const list = snapshot.root.definedNames;
    const idx = list.findIndex((d) => d.name === PRINT_AREA && d.scope === sheet.name);

    if (payload.clear || !payload.range) {
      if (payload.clear) {
        if (idx === -1) {
          return { next: snapshot, diff: buildDiff(snapshot.revision, snapshot.revision, []) };
        }
        return commit(snapshot, removeAt(list, idx), `clear print area on ${sheet.name}`);
      }
      throw new CommandError(
        "invalid-payload",
        "set-print-area: must supply --range, or --clear to remove the existing area."
      );
    }

    parseRangeRef(payload.range); // throws on bad input
    const refersTo = `'${escapeSheetName(sheet.name)}'!${absolutise(payload.range)}`;
    const entry: DefinedName = {
      id: `dn-${PRINT_AREA}-${sheet.name}`,
      name: PRINT_AREA,
      scope: sheet.name,
      refersTo,
    };
    const nextList = idx === -1 ? [...list, entry] : list.map((d, i) => (i === idx ? entry : d));
    if (idx !== -1 && list[idx]?.refersTo === refersTo) {
      return { next: snapshot, diff: buildDiff(snapshot.revision, snapshot.revision, []) };
    }
    return commit(snapshot, nextList, `set print area on ${sheet.name} = ${payload.range}`);
  },
};

/**
 * `xlsx:set-print-titles` — write / clear the per-sheet
 * `_xlnm.Print_Titles` defined name. Mirrors Page Layout → Print
 * Titles. Excel uses a comma-separated `refersTo` of two ranges:
 *
 *   `'Sheet1'!$1:$1,'Sheet1'!$A:$B`
 *
 * Either rows, columns, or both may be supplied.
 */
export const setPrintTitlesHandler: CommandHandler<SetPrintTitlesPayload, XlsxSnapshot> = {
  type: "xlsx:set-print-titles",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const list = snapshot.root.definedNames;
    const idx = list.findIndex((d) => d.name === PRINT_TITLES && d.scope === sheet.name);

    if (payload.clear || (!payload.rows && !payload.cols)) {
      if (payload.clear) {
        if (idx === -1) {
          return { next: snapshot, diff: buildDiff(snapshot.revision, snapshot.revision, []) };
        }
        return commit(snapshot, removeAt(list, idx), `clear print titles on ${sheet.name}`);
      }
      throw new CommandError(
        "invalid-payload",
        "set-print-titles: must supply --rows or --cols (or both), or --clear."
      );
    }

    const parts: string[] = [];
    const sheetRef = `'${escapeSheetName(sheet.name)}'`;
    if (payload.rows) {
      parts.push(`${sheetRef}!${absolutiseRowRange(payload.rows)}`);
    }
    if (payload.cols) {
      parts.push(`${sheetRef}!${absolutiseColRange(payload.cols)}`);
    }
    const refersTo = parts.join(",");

    const entry: DefinedName = {
      id: `dn-${PRINT_TITLES}-${sheet.name}`,
      name: PRINT_TITLES,
      scope: sheet.name,
      refersTo,
    };
    const nextList = idx === -1 ? [...list, entry] : list.map((d, i) => (i === idx ? entry : d));
    if (idx !== -1 && list[idx]?.refersTo === refersTo) {
      return { next: snapshot, diff: buildDiff(snapshot.revision, snapshot.revision, []) };
    }
    return commit(snapshot, nextList, `set print titles on ${sheet.name}`);
  },
};

function commit(
  snapshot: XlsxSnapshot,
  definedNames: ReadonlyArray<DefinedName>,
  summary: string
): { next: XlsxSnapshot; diff: ReturnType<typeof buildDiff> } {
  const next: XlsxWorkbook = { ...snapshot.root, definedNames };
  const evolved = evolveSnapshot(snapshot, next, { workbook: true });
  return {
    next: evolved,
    diff: buildDiff(snapshot.revision, evolved.revision, [
      {
        kind: "node-updated",
        nodeId: snapshot.root.id,
        path: ["definedNames"],
        field: "definedNames",
        summary,
      },
    ]),
  };
}

function removeAt(list: ReadonlyArray<DefinedName>, idx: number): ReadonlyArray<DefinedName> {
  return list.filter((_, i) => i !== idx);
}

function escapeSheetName(name: string): string {
  return name.replace(/'/g, "''");
}

function absolutise(range: string): string {
  // Convert "A1:D20" → "$A$1:$D$20"
  const parts = range.split(":");
  return parts.map((p) => absolutiseCell(p)).join(":");
}

function absolutiseCell(ref: string): string {
  parseCellRef(ref); // validates
  const m = /^([A-Z]+)(\d+)$/.exec(ref.toUpperCase());
  if (!m) return ref;
  return `$${m[1]}$${m[2]}`;
}

function absolutiseRowRange(rows: string): string {
  // Accepts "1", "1:1", "$1:$1" → "$1:$1"
  const cleaned = rows.replace(/\$/g, "");
  const m = /^(\d+)(?::(\d+))?$/.exec(cleaned);
  if (!m) {
    throw new CommandError("invalid-payload", `set-print-titles: invalid row range "${rows}"`);
  }
  const a = m[1];
  const b = m[2] ?? a;
  return `$${a}:$${b}`;
}

function absolutiseColRange(cols: string): string {
  const cleaned = cols.replace(/\$/g, "").toUpperCase();
  const m = /^([A-Z]+)(?::([A-Z]+))?$/.exec(cleaned);
  if (!m) {
    throw new CommandError("invalid-payload", `set-print-titles: invalid column range "${cols}"`);
  }
  const a = m[1];
  const b = m[2] ?? a;
  return `$${a}:$${b}`;
}
