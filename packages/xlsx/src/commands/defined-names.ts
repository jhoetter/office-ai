import { CommandError, type CommandHandler } from "@officeai/core";
import type { DefinedName, XlsxSnapshot, XlsxWorkbook } from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import type {
  AddDefinedNamePayload,
  RemoveDefinedNamePayload,
  UpdateDefinedNamePayload,
} from "./payloads.js";

/**
 * C12 — Defined-name (named range) command handlers.
 *
 * The handlers mutate {@link XlsxWorkbook.definedNames} and mark
 * `dirty.workbook` so the serializer re-emits the
 * `<definedNames>` block. Validation mirrors Excel's Name Manager:
 *
 *   - Names start with a letter or `_`, contain only letters,
 *     digits, dot, underscore (no spaces, no operators).
 *   - Names are not pure cell references (`A1`, `R1C1`).
 *   - Within a single scope, names are unique.
 *   - Reserved names (`Print_Area`, `Print_Titles`, `Criteria`,
 *     `Database`, `Extract`) are accepted but not minted by our UI.
 */

const NAME_RE = /^[A-Za-z_][A-Za-z0-9_.]*$/;
const A1_RE = /^[A-Z]{1,3}\$?\d{1,7}$/;
const R1C1_RE = /^R\d+C\d+$/i;

function validateName(name: string): void {
  if (!name) {
    throw new CommandError("invalid-name", "Defined name cannot be empty");
  }
  if (name.length > 255) {
    throw new CommandError("invalid-name", `Defined name "${name}" exceeds 255 characters`);
  }
  if (!NAME_RE.test(name)) {
    throw new CommandError(
      "invalid-name",
      `"${name}" is not a valid defined name (must start with a letter or _ and contain only letters, digits, dot, underscore)`
    );
  }
  if (A1_RE.test(name) || R1C1_RE.test(name)) {
    throw new CommandError(
      "invalid-name",
      `"${name}" looks like a cell reference and cannot be used as a defined name`
    );
  }
}

function findIndex(list: ReadonlyArray<DefinedName>, name: string, scope: string | undefined): number {
  return list.findIndex((d) => d.name === name && (d.scope ?? undefined) === scope);
}

function makeId(name: string, scope: string | undefined): string {
  return `dn-${name}-${scope ?? "wb"}`;
}

export const addDefinedNameHandler: CommandHandler<AddDefinedNamePayload, XlsxSnapshot> = {
  type: "xlsx:add-defined-name",
  apply(snapshot, payload) {
    validateName(payload.name);
    if (!payload.refersTo || !payload.refersTo.trim()) {
      throw new CommandError("invalid-name", "Defined name needs a refersTo expression");
    }
    if (payload.scope && !snapshot.root.sheets.some((s) => s.name === payload.scope)) {
      throw new CommandError("unknown-sheet", `Unknown sheet scope "${payload.scope}"`);
    }
    const list = snapshot.root.definedNames;
    if (findIndex(list, payload.name, payload.scope) !== -1) {
      throw new CommandError(
        "duplicate-name",
        `A defined name "${payload.name}" already exists in this scope`
      );
    }
    const entry: DefinedName = {
      id: makeId(payload.name, payload.scope),
      name: payload.name,
      refersTo: payload.refersTo.trim(),
      ...(payload.scope ? { scope: payload.scope } : {}),
      ...(payload.comment ? { comment: payload.comment } : {}),
    };
    const definedNames = [...list, entry];
    const next: XlsxWorkbook = { ...snapshot.root, definedNames };
    const evolved = evolveSnapshot(snapshot, next, { workbook: true });
    return {
      next: evolved,
      diff: buildDiff(snapshot.revision, evolved.revision, [
        {
          kind: "node-updated",
          nodeId: snapshot.root.id,
          path: ["definedNames", entry.id],
          field: "definedNames",
          summary: `add defined name ${entry.name}${entry.scope ? ` (${entry.scope})` : ""}`,
        },
      ]),
    };
  },
};

export const updateDefinedNameHandler: CommandHandler<UpdateDefinedNamePayload, XlsxSnapshot> = {
  type: "xlsx:update-defined-name",
  apply(snapshot, payload) {
    const list = snapshot.root.definedNames;
    const idx = findIndex(list, payload.name, payload.scope);
    if (idx === -1) {
      throw new CommandError(
        "unknown-name",
        `No defined name "${payload.name}" in ${payload.scope ?? "workbook"} scope`
      );
    }
    const prev = list[idx]!;
    const nextName = payload.nextName ?? prev.name;
    if (nextName !== prev.name) {
      validateName(nextName);
      if (findIndex(list, nextName, prev.scope) !== -1) {
        throw new CommandError("duplicate-name", `A defined name "${nextName}" already exists in this scope`);
      }
    }
    const refersTo = payload.refersTo !== undefined ? payload.refersTo.trim() : prev.refersTo;
    if (!refersTo) {
      throw new CommandError("invalid-name", "Defined name needs a refersTo expression");
    }
    const updated: DefinedName = {
      ...prev,
      id: makeId(nextName, prev.scope),
      name: nextName,
      refersTo,
      ...(payload.comment !== undefined ? { comment: payload.comment } : {}),
    };
    const definedNames = list.slice();
    definedNames[idx] = updated;
    const next: XlsxWorkbook = { ...snapshot.root, definedNames };
    const evolved = evolveSnapshot(snapshot, next, { workbook: true });
    return {
      next: evolved,
      diff: buildDiff(snapshot.revision, evolved.revision, [
        {
          kind: "node-updated",
          nodeId: snapshot.root.id,
          path: ["definedNames", updated.id],
          field: "definedNames",
          summary: `update defined name ${prev.name} → ${updated.name}`,
        },
      ]),
    };
  },
};

export const removeDefinedNameHandler: CommandHandler<RemoveDefinedNamePayload, XlsxSnapshot> = {
  type: "xlsx:remove-defined-name",
  apply(snapshot, payload) {
    const list = snapshot.root.definedNames;
    const idx = findIndex(list, payload.name, payload.scope);
    if (idx === -1) {
      throw new CommandError(
        "unknown-name",
        `No defined name "${payload.name}" in ${payload.scope ?? "workbook"} scope`
      );
    }
    const prev = list[idx]!;
    const definedNames = list.slice();
    definedNames.splice(idx, 1);
    const next: XlsxWorkbook = { ...snapshot.root, definedNames };
    const evolved = evolveSnapshot(snapshot, next, { workbook: true });
    return {
      next: evolved,
      diff: buildDiff(snapshot.revision, evolved.revision, [
        {
          kind: "node-updated",
          nodeId: snapshot.root.id,
          path: ["definedNames", prev.id],
          field: "definedNames",
          summary: `remove defined name ${prev.name}`,
        },
      ]),
    };
  },
};
