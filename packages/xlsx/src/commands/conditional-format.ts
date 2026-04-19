import { CommandError, type CommandHandler } from "@officeai/core";
import type { Sheet, XlsxSnapshot, XlsxWorkbook } from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import type {
  AddConditionalFormatPayload,
  RemoveConditionalFormatPayload,
  ClearConditionalFormatsPayload,
} from "./payloads.js";

/**
 * C10 — Conditional Formatting commands.
 *
 * The model carries typed rules in `Sheet.conditionalFormats`. The
 * Grid evaluates them at render time and overlays
 * fill/font/data-bar visuals on top of the base style. Typed rules
 * are NOT yet emitted to OOXML on serialize — opening + saving
 * preserves any pre-existing `<conditionalFormatting>` blocks via
 * `Sheet.opaqueConditionalFormats`, but new typed rules will be
 * dropped by the serializer until the dxf-emission pass lands.
 */

function findSheet(snapshot: XlsxSnapshot, name: string): Sheet {
  const sheet = snapshot.root.sheets.find((s) => s.name === name);
  if (!sheet) {
    throw new CommandError(
      "unknown-sheet",
      `Sheet "${name}" not found; available: ${snapshot.root.sheets.map((s) => s.name).join(", ")}`
    );
  }
  return sheet;
}

function replaceSheet(snapshot: XlsxSnapshot, next: Sheet): XlsxWorkbook {
  const sheets = snapshot.root.sheets.slice();
  sheets[next.index] = next;
  return { ...snapshot.root, sheets };
}

export const addConditionalFormatHandler: CommandHandler<AddConditionalFormatPayload, XlsxSnapshot> = {
  type: "xlsx:add-conditional-format",
  apply(snapshot, payload) {
    const sheet = findSheet(snapshot, payload.sheet);
    if (!payload.rule.id || payload.rule.id.length === 0) {
      throw new CommandError("invalid-position", "Conditional format rule requires an id");
    }
    if (sheet.conditionalFormats.some((r) => r.id === payload.rule.id)) {
      throw new CommandError(
        "invalid-position",
        `Conditional format rule "${payload.rule.id}" already exists on sheet "${sheet.name}"`
      );
    }
    const next: Sheet = {
      ...sheet,
      conditionalFormats: [...sheet.conditionalFormats, payload.rule],
    };
    const evolved = evolveSnapshot(snapshot, replaceSheet(snapshot, next), {
      sheets: [sheet.partPath],
    });
    return {
      next: evolved,
      diff: buildDiff(snapshot.revision, evolved.revision, [
        {
          kind: "node-inserted",
          nodeId: sheet.id,
          path: ["sheets", sheet.index, "conditionalFormats"],
          summary: `add conditional format ${payload.rule.id} on ${sheet.name}`,
        },
      ]),
    };
  },
};

export const removeConditionalFormatHandler: CommandHandler<RemoveConditionalFormatPayload, XlsxSnapshot> = {
  type: "xlsx:remove-conditional-format",
  apply(snapshot, payload) {
    const sheet = findSheet(snapshot, payload.sheet);
    if (!sheet.conditionalFormats.some((r) => r.id === payload.id)) {
      throw new CommandError(
        "invalid-position",
        `Conditional format rule "${payload.id}" not found on sheet "${sheet.name}"`
      );
    }
    const next: Sheet = {
      ...sheet,
      conditionalFormats: sheet.conditionalFormats.filter((r) => r.id !== payload.id),
    };
    const evolved = evolveSnapshot(snapshot, replaceSheet(snapshot, next), {
      sheets: [sheet.partPath],
    });
    return {
      next: evolved,
      diff: buildDiff(snapshot.revision, evolved.revision, [
        {
          kind: "node-deleted",
          nodeId: sheet.id,
          path: ["sheets", sheet.index, "conditionalFormats", payload.id],
          summary: `remove conditional format ${payload.id} on ${sheet.name}`,
        },
      ]),
    };
  },
};

export const clearConditionalFormatsHandler: CommandHandler<ClearConditionalFormatsPayload, XlsxSnapshot> = {
  type: "xlsx:clear-conditional-formats",
  apply(snapshot, payload) {
    const sheet = findSheet(snapshot, payload.sheet);
    if (sheet.conditionalFormats.length === 0) {
      const evolved = evolveSnapshot(snapshot, snapshot.root, {});
      return {
        next: evolved,
        diff: buildDiff(snapshot.revision, evolved.revision, []),
      };
    }
    const next: Sheet = { ...sheet, conditionalFormats: [] };
    const evolved = evolveSnapshot(snapshot, replaceSheet(snapshot, next), {
      sheets: [sheet.partPath],
    });
    return {
      next: evolved,
      diff: buildDiff(snapshot.revision, evolved.revision, [
        {
          kind: "node-deleted",
          nodeId: sheet.id,
          path: ["sheets", sheet.index, "conditionalFormats"],
          summary: `clear conditional formats on ${sheet.name}`,
        },
      ]),
    };
  },
};
