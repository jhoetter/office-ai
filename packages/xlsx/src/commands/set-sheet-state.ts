import { CommandError, type CommandHandler } from "@officeai/core";
import type { Sheet, XlsxSnapshot, XlsxWorkbook } from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import type { SetSheetStatePayload } from "./payloads.js";

/**
 * `xlsx:set-sheet-state` — toggle a worksheet between
 * `visible` / `hidden` / `veryHidden`.
 *
 * Excel parity:
 *   - At least one sheet must remain `visible` at all times.
 *   - `veryHidden` sheets are not exposed via Excel's "Unhide…" menu;
 *     they can only be revived by macros (or, here, by a future Name
 *     Manager / VBA-style hatch). The command itself accepts the
 *     transition though.
 *
 * Round-trip:
 *   - The serializer already emits `state="hidden"` /
 *     `state="veryHidden"` on `<sheet>` from `Sheet.state` when it
 *     differs from `"visible"`, so we just flip the typed field and
 *     mark `dirty.workbook`.
 */
export const setSheetStateHandler: CommandHandler<SetSheetStatePayload, XlsxSnapshot> = {
  type: "xlsx:set-sheet-state",
  apply(snapshot, payload) {
    const sheet = snapshot.root.sheets.find((s) => s.name === payload.name);
    if (!sheet) {
      throw new CommandError(
        "unknown-sheet",
        `Sheet "${payload.name}" not found; available: ${snapshot.root.sheets.map((s) => s.name).join(", ")}`
      );
    }

    if (sheet.state === payload.state) {
      const next = evolveSnapshot(snapshot, snapshot.root, {});
      return {
        next,
        diff: buildDiff(snapshot.revision, next.revision, [
          {
            kind: "node-updated",
            nodeId: sheet.id,
            path: ["sheets", sheet.index, "state"],
            field: "state",
            summary: `set sheet state ${sheet.name} (no-op, already ${payload.state})`,
          },
        ]),
      };
    }

    if (payload.state !== "visible") {
      const visibleCount = snapshot.root.sheets.filter((s) => s.state === "visible").length;
      if (sheet.state === "visible" && visibleCount <= 1) {
        throw new CommandError("invalid-position", `Cannot hide the only visible sheet ("${sheet.name}")`);
      }
    }

    const updated: Sheet = { ...sheet, state: payload.state };
    const sheets = snapshot.root.sheets.slice();
    sheets[sheet.index] = updated;

    const nextWorkbook: XlsxWorkbook = { ...snapshot.root, sheets };
    const next = evolveSnapshot(snapshot, nextWorkbook, { workbook: true });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, [
        {
          kind: "node-updated",
          nodeId: sheet.id,
          path: ["sheets", sheet.index, "state"],
          field: "state",
          summary: `set sheet state ${sheet.name}: ${sheet.state} → ${payload.state}`,
        },
      ]),
    };
  },
};
