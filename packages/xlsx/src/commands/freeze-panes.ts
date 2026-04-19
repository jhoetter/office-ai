import { CommandError, type CommandHandler } from "@officeai/core";
import type { Sheet, XlsxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { FreezePanesPayload, UnfreezePanesPayload } from "./payloads.js";
import { resolveSheet } from "./validation.js";

const MAX_ROWS = 1_048_576;
const MAX_COLS = 16_384;

/**
 * `xlsx:freeze-panes` (C3) — toggle / update Excel's frozen-pane
 * configuration for a sheet. Setting both axes to `0` is equivalent
 * to dispatching `xlsx:unfreeze-panes`.
 *
 * The handler keeps the model the single source of truth — the
 * serializer rebuilds `<sheetView><pane state="frozen"/></sheetView>`
 * from `Sheet.freeze` on every save, so Excel restores the frozen
 * state on open.
 */
export const freezePanesHandler: CommandHandler<FreezePanesPayload, XlsxSnapshot> = {
  type: "xlsx:freeze-panes",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    if (!Number.isInteger(payload.rows) || payload.rows < 0 || payload.rows > MAX_ROWS) {
      throw new CommandError(
        "invalid-range",
        `rows must be an integer in [0, ${MAX_ROWS}]; got ${payload.rows}`
      );
    }
    if (!Number.isInteger(payload.cols) || payload.cols < 0 || payload.cols > MAX_COLS) {
      throw new CommandError(
        "invalid-range",
        `cols must be an integer in [0, ${MAX_COLS}]; got ${payload.cols}`
      );
    }

    return applyFreeze(snapshot, sheet, payload.rows, payload.cols);
  },
};

/**
 * `xlsx:unfreeze-panes` — drop any frozen-pane configuration on the
 * sheet. No-ops cleanly when nothing is frozen.
 */
export const unfreezePanesHandler: CommandHandler<UnfreezePanesPayload, XlsxSnapshot> = {
  type: "xlsx:unfreeze-panes",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    return applyFreeze(snapshot, sheet, 0, 0);
  },
};

function applyFreeze(snapshot: XlsxSnapshot, sheet: Sheet, rows: number, cols: number) {
  const before = sheet.freeze ?? { rows: 0, cols: 0 };
  const next = rows === 0 && cols === 0 ? undefined : { rows, cols };
  if (before.rows === (next?.rows ?? 0) && before.cols === (next?.cols ?? 0)) {
    const noopNext = evolveSnapshot(snapshot, snapshot.root, {});
    return { next: noopNext, diff: buildDiff(snapshot.revision, noopNext.revision, []) };
  }

  const nextSheet: Sheet = next
    ? { ...sheet, freeze: next }
    : (() => {
        const { freeze: _drop, ...rest } = sheet;
        return rest;
      })();
  const nextWorkbook = replaceSheet(snapshot.root, nextSheet);
  const nextSnap = evolveSnapshot(snapshot, nextWorkbook, { sheets: [sheet.partPath] });
  const beforeLabel = before.rows === 0 && before.cols === 0 ? "none" : `${before.rows}r×${before.cols}c`;
  const afterLabel = next ? `${next.rows}r×${next.cols}c` : "none";
  return {
    next: nextSnap,
    diff: buildDiff(snapshot.revision, nextSnap.revision, [
      {
        kind: "node-updated",
        nodeId: sheet.id,
        path: ["sheets", sheet.index, "freeze"],
        field: "freeze",
        summary: `${sheet.name} freeze: ${beforeLabel} → ${afterLabel}`,
        meta: { before: before.rows === 0 && before.cols === 0 ? null : before, after: next ?? null },
      },
    ]),
  };
}
