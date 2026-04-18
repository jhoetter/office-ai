import { CommandError, type CommandHandler, type DiffChange } from "@officeai/core";
import { cellKey, formatA1 } from "../model/refs.js";
import {
  flattenCellXf,
  internNumberFormat,
  internStyle,
  type EffectiveStyle,
} from "../model/style-mutate.js";
import type {
  StyleAlignment,
  StyleBorder,
  StyleBorderSide,
  StyleColor,
  StyleFill,
  StyleFont,
  StyleTable,
} from "../model/style-table.js";
import type { Cell, Sheet, XlsxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { CellFormatBorderSide, CellFormatPatch, SetCellFormatPayload } from "./payloads.js";
import { parseRangeRef, resolveSheet } from "./validation.js";

/**
 * `xlsx:set-cell-format` — apply an agent-facing `CellFormat` patch to
 * every cell in `range` and re-intern through the typed style table.
 *
 * Spec: `spec/xlsx/agent-commands.md` §4 + `document-model.md` §6.
 *
 * Pipeline:
 *   1. Resolve sheet + range; precheck malformed colour / numberFormat.
 *   2. For each cell in the range:
 *      - Resolve current effective style from `cell.styleId` (or the
 *        default xf).
 *      - Apply the patch.
 *      - Intern via `internStyle` (content-hash dedupes; reuses
 *        existing `xfId`s where possible).
 *      - Set `cell.styleId = newXfId`. Blank cells get materialised
 *        with `value: null` so the styleId has somewhere to live.
 *   3. Mark `dirty.styles = true` and the sheet path dirty.
 *   4. Diff: one `node-updated{ field: "format" }` per cell that
 *      changed, plus `style-added` entries via `meta` for any new
 *      font/fill/border/xf indices.
 *
 * The "small constant" property in the spec — bulk-bolding 1000 cells
 * only allocates one new xfId per pre-existing distinct style — falls
 * out of step 2's content-hash intern accumulator.
 */
export const setCellFormatHandler: CommandHandler<SetCellFormatPayload, XlsxSnapshot> = {
  type: "xlsx:set-cell-format",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const range = parseRangeRef(payload.range);
    validatePatch(payload.format);

    let table = snapshot.root.styles;
    const newCells = new Map(sheet.cells);
    const changes: DiffChange[] = [];
    const addedFontIds = new Set<number>();
    const addedFillIds = new Set<number>();
    const addedBorderIds = new Set<number>();
    const addedXfIds = new Set<number>();
    const addedNumFmtIds = new Set<number>();

    for (let r = range.start.row; r <= range.end.row; r++) {
      for (let c = range.start.col; c <= range.end.col; c++) {
        const key = cellKey(r, c);
        const existing = sheet.cells.get(key);
        const beforeXfId = existing?.styleId;
        const beforeEffective = flattenCellXf(table, beforeXfId);

        const patchResult = applyPatch(table, beforeEffective, payload.format);
        table = patchResult.table;
        if (patchResult.addedNumFmtId !== undefined) addedNumFmtIds.add(patchResult.addedNumFmtId);

        const interned = internStyle(table, patchResult.effective);
        table = interned.table;
        if (interned.addedFontId !== undefined) addedFontIds.add(interned.addedFontId);
        if (interned.addedFillId !== undefined) addedFillIds.add(interned.addedFillId);
        if (interned.addedBorderId !== undefined) addedBorderIds.add(interned.addedBorderId);
        if (interned.addedXfId !== undefined) addedXfIds.add(interned.addedXfId);

        if (existing && existing.styleId === interned.xfId) continue;

        const nextCell: Cell = existing
          ? { ...existing, styleId: interned.xfId }
          : { row: r, col: c, value: null, styleId: interned.xfId };
        newCells.set(key, nextCell);

        const a1 = formatA1({ row: r, col: c });
        changes.push({
          kind: "node-updated",
          nodeId: sheet.id,
          path: ["sheets", sheet.index, "cells", `${sheet.name}!${a1}`],
          field: "format",
          summary: `${sheet.name}!${a1}: styleId ${beforeXfId ?? 0} → ${interned.xfId}`,
          meta: {
            before: { styleId: beforeXfId ?? null },
            after: { styleId: interned.xfId },
          },
        });
      }
    }

    if (changes.length === 0) {
      const noopNext = evolveSnapshot(snapshot, snapshot.root, {});
      return { next: noopNext, diff: buildDiff(snapshot.revision, noopNext.revision, []) };
    }

    if (addedFontIds.size > 0 || addedFillIds.size > 0 || addedBorderIds.size > 0 || addedXfIds.size > 0) {
      changes.push({
        kind: "node-updated",
        nodeId: snapshot.root.id,
        path: ["styles"],
        field: "style-added",
        summary:
          `style-table grew: ` +
          `+${addedFontIds.size} font, +${addedFillIds.size} fill, ` +
          `+${addedBorderIds.size} border, +${addedXfIds.size} xf`,
        meta: {
          addedFontIds: Array.from(addedFontIds),
          addedFillIds: Array.from(addedFillIds),
          addedBorderIds: Array.from(addedBorderIds),
          addedXfIds: Array.from(addedXfIds),
          addedNumFmtIds: Array.from(addedNumFmtIds),
        },
      });
    }

    const nextSheet: Sheet = { ...sheet, cells: newCells };
    const workbookWithSheet = replaceSheet(snapshot.root, nextSheet);
    const nextWorkbook = { ...workbookWithSheet, styles: table };
    const next = evolveSnapshot(snapshot, nextWorkbook, {
      sheets: [sheet.partPath],
      styles: true,
    });

    return { next, diff: buildDiff(snapshot.revision, next.revision, changes) };
  },
};

/* ── patch validation + application ─────────────────────────────────────── */

const RGB_RE = /^[0-9A-Fa-f]{6}$/;

function validatePatch(patch: CellFormatPatch): void {
  if (patch.font?.color !== undefined) assertColor("font.color", patch.font.color);
  if (patch.fill?.color !== undefined) assertColor("fill.color", patch.fill.color);
  for (const side of ["top", "right", "bottom", "left"] as const) {
    const s = patch.border?.[side];
    if (s?.color !== undefined) assertColor(`border.${side}.color`, s.color);
  }
}

function assertColor(field: string, value: string): void {
  if (!RGB_RE.test(value)) {
    throw new CommandError(
      "invalid-format",
      `${field} must be a 6-hex-digit RRGGBB without '#'; got ${JSON.stringify(value)}`
    );
  }
}

interface PatchResult {
  readonly table: StyleTable;
  readonly effective: EffectiveStyle;
  readonly addedNumFmtId?: number;
}

function applyPatch(table: StyleTable, current: EffectiveStyle, patch: CellFormatPatch): PatchResult {
  let nextTable = table;
  let addedNumFmtId: number | undefined;

  let numFmtId = current.numFmtId;
  if (patch.numberFormat !== undefined) {
    const resolved = resolveNumberFormat(nextTable, patch.numberFormat);
    if (resolved.added) {
      addedNumFmtId = resolved.numFmtId;
      nextTable = resolved.table;
    }
    numFmtId = resolved.numFmtId;
  }

  const font = mergeFont(current.font, patch.font);
  const fill = mergeFill(current.fill, patch.fill);
  const border = mergeBorder(current.border, patch.border);
  const alignment = mergeAlignment(current.alignment, patch.alignment);

  const effective: EffectiveStyle = {
    numFmtId,
    font,
    fill,
    border,
    ...(alignment ? { alignment } : {}),
    ...(current.protection ? { protection: current.protection } : {}),
  };
  return addedNumFmtId !== undefined
    ? { table: nextTable, effective, addedNumFmtId }
    : { table: nextTable, effective };
}

function resolveNumberFormat(
  table: StyleTable,
  raw: string
): { table: StyleTable; numFmtId: number; added: boolean } {
  const trimmed = raw.trim();
  const asInt = Number.parseInt(trimmed, 10);
  if (Number.isFinite(asInt) && String(asInt) === trimmed) {
    if (asInt < 0) {
      throw new CommandError("invalid-format", `numberFormat id must be ≥ 0; got ${trimmed}`);
    }
    if (asInt <= 49) {
      return { table, numFmtId: asInt, added: false };
    }
    if (asInt < 164) {
      if (!table.numFmts.has(asInt)) {
        throw new CommandError(
          "unknown-style-id",
          `numberFormat id ${asInt} is in the reserved 50..163 range and not registered in this workbook; use a built-in 0..49 id or a custom format string`
        );
      }
      return { table, numFmtId: asInt, added: false };
    }
    if (table.numFmts.has(asInt)) return { table, numFmtId: asInt, added: false };
    throw new CommandError(
      "unknown-style-id",
      `numberFormat id ${asInt} not registered in this workbook; pass the format string instead to register it`
    );
  }
  return internNumberFormat(table, trimmed);
}

function mergeFont(current: StyleFont, patch: CellFormatPatch["font"]): StyleFont {
  if (!patch) return current;
  const next: {
    name?: string;
    size?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: string | true;
    strike?: boolean;
    color?: StyleColor;
    family?: number;
    scheme?: string;
    charset?: number;
    opaqueExtras: ReadonlyArray<string>;
  } = { ...current, opaqueExtras: current.opaqueExtras };
  if (patch.family !== undefined) next.name = patch.family;
  if (patch.size !== undefined) next.size = patch.size;
  if (patch.bold !== undefined) next.bold = patch.bold;
  if (patch.italic !== undefined) next.italic = patch.italic;
  if (patch.underline !== undefined) next.underline = patch.underline ? true : undefined;
  if (patch.strike !== undefined) next.strike = patch.strike;
  if (patch.color !== undefined) next.color = { rgb: `FF${patch.color.toUpperCase()}` };
  return next;
}

function mergeFill(current: StyleFill, patch: CellFormatPatch["fill"]): StyleFill {
  if (!patch) return current;
  const pattern = patch.pattern ?? (patch.color !== undefined ? "solid" : (current.patternType ?? "none"));
  const fill: { kind: "pattern"; patternType?: string; fgColor?: StyleColor; bgColor?: StyleColor } = {
    kind: "pattern",
    patternType: pattern,
  };
  if (patch.color !== undefined) {
    const rgb: StyleColor = { rgb: `FF${patch.color.toUpperCase()}` };
    fill.fgColor = rgb;
    fill.bgColor = { indexed: 64 };
  } else if (current.kind === "pattern") {
    if (current.fgColor) fill.fgColor = current.fgColor;
    if (current.bgColor) fill.bgColor = current.bgColor;
  }
  return fill;
}

function mergeBorder(current: StyleBorder, patch: CellFormatPatch["border"]): StyleBorder {
  if (!patch) return current;
  const next: {
    left?: StyleBorderSide;
    right?: StyleBorderSide;
    top?: StyleBorderSide;
    bottom?: StyleBorderSide;
    diagonal?: StyleBorderSide;
    diagonalUp?: boolean;
    diagonalDown?: boolean;
    outline?: boolean;
    opaqueExtras: ReadonlyArray<string>;
  } = { ...current, opaqueExtras: current.opaqueExtras };
  if (patch.top) next.top = mergeBorderSide(current.top, patch.top);
  if (patch.right) next.right = mergeBorderSide(current.right, patch.right);
  if (patch.bottom) next.bottom = mergeBorderSide(current.bottom, patch.bottom);
  if (patch.left) next.left = mergeBorderSide(current.left, patch.left);
  return next;
}

function mergeBorderSide(current: StyleBorderSide | undefined, patch: CellFormatBorderSide): StyleBorderSide {
  const side: { style?: string; color?: StyleColor } = { ...(current ?? {}) };
  if (patch.style !== undefined) {
    if (patch.style === "none") {
      delete side.style;
      delete side.color;
    } else {
      side.style = patch.style;
    }
  }
  if (patch.color !== undefined) side.color = { rgb: `FF${patch.color.toUpperCase()}` };
  return side;
}

function mergeAlignment(
  current: StyleAlignment | undefined,
  patch: CellFormatPatch["alignment"]
): StyleAlignment | undefined {
  if (!patch) return current;
  const next: {
    horizontal?: string;
    vertical?: string;
    wrapText?: boolean;
    textRotation?: number;
    indent?: number;
    shrinkToFit?: boolean;
    readingOrder?: number;
    justifyLastLine?: boolean;
    relativeIndent?: number;
  } = { ...(current ?? {}) };
  if (patch.horizontal !== undefined) next.horizontal = patch.horizontal;
  if (patch.vertical !== undefined) next.vertical = mapVerticalAgentToOoxml(patch.vertical);
  if (patch.wrapText !== undefined) next.wrapText = patch.wrapText;
  if (patch.indent !== undefined) next.indent = patch.indent;
  return next;
}

function mapVerticalAgentToOoxml(v: "top" | "middle" | "bottom"): string {
  switch (v) {
    case "top":
      return "top";
    case "middle":
      return "center";
    case "bottom":
      return "bottom";
  }
}
