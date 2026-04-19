import { CommandError, type CommandHandler } from "@officeai/core";
import type { Sheet, XlsxSnapshot, XlsxWorkbook } from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import type {
  AddDataValidationPayload,
  RemoveDataValidationPayload,
  ClearDataValidationsPayload,
} from "./payloads.js";

/**
 * C11 — Data Validation commands.
 *
 * The model currently types one validation kind: `list` (in-cell
 * dropdown picker). The serializer re-emits typed lists alongside
 * any opaque non-list rules captured at parse time so opening +
 * saving never loses validations the user didn't touch.
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

export const addDataValidationHandler: CommandHandler<AddDataValidationPayload, XlsxSnapshot> = {
  type: "xlsx:add-data-validation",
  apply(snapshot, payload) {
    const sheet = findSheet(snapshot, payload.sheet);
    if (!payload.rule.id || payload.rule.id.length === 0) {
      throw new CommandError("invalid-position", "Data validation rule requires an id");
    }
    if (sheet.dataValidations.some((r) => r.id === payload.rule.id)) {
      throw new CommandError(
        "invalid-position",
        `Data validation rule "${payload.rule.id}" already exists on sheet "${sheet.name}"`
      );
    }
    if (payload.rule.kind === "list" && payload.rule.source.trim() === "") {
      throw new CommandError("invalid-position", "Data validation list requires a non-empty source");
    }
    const next: Sheet = {
      ...sheet,
      dataValidations: [...sheet.dataValidations, payload.rule],
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
          path: ["sheets", sheet.index, "dataValidations"],
          summary: `add data validation ${payload.rule.id} on ${sheet.name}`,
        },
      ]),
    };
  },
};

export const removeDataValidationHandler: CommandHandler<RemoveDataValidationPayload, XlsxSnapshot> = {
  type: "xlsx:remove-data-validation",
  apply(snapshot, payload) {
    const sheet = findSheet(snapshot, payload.sheet);
    if (!sheet.dataValidations.some((r) => r.id === payload.id)) {
      throw new CommandError(
        "invalid-position",
        `Data validation rule "${payload.id}" not found on sheet "${sheet.name}"`
      );
    }
    const next: Sheet = {
      ...sheet,
      dataValidations: sheet.dataValidations.filter((r) => r.id !== payload.id),
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
          path: ["sheets", sheet.index, "dataValidations", payload.id],
          summary: `remove data validation ${payload.id} on ${sheet.name}`,
        },
      ]),
    };
  },
};

export const clearDataValidationsHandler: CommandHandler<ClearDataValidationsPayload, XlsxSnapshot> = {
  type: "xlsx:clear-data-validations",
  apply(snapshot, payload) {
    const sheet = findSheet(snapshot, payload.sheet);
    if (sheet.dataValidations.length === 0 && !sheet.opaqueDataValidations) {
      const evolved = evolveSnapshot(snapshot, snapshot.root, {});
      return {
        next: evolved,
        diff: buildDiff(snapshot.revision, evolved.revision, []),
      };
    }
    // Strip both typed and opaque so "Remove all rules" feels final.
    const next: Sheet = (() => {
      const { opaqueDataValidations: _omit, ...rest } = sheet;
      void _omit;
      return { ...rest, dataValidations: [] };
    })();
    const evolved = evolveSnapshot(snapshot, replaceSheet(snapshot, next), {
      sheets: [sheet.partPath],
    });
    return {
      next: evolved,
      diff: buildDiff(snapshot.revision, evolved.revision, [
        {
          kind: "node-deleted",
          nodeId: sheet.id,
          path: ["sheets", sheet.index, "dataValidations"],
          summary: `clear data validations on ${sheet.name}`,
        },
      ]),
    };
  },
};
