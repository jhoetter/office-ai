import { CommandError, type CommandHandler } from "@officeai/core";
import type { Sheet, XlsxSnapshot, XlsxWorkbook } from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import type { SetSheetTabColorPayload } from "./payloads.js";

const ARGB_RE = /^[0-9A-Fa-f]{6}$|^[0-9A-Fa-f]{8}$/;

/**
 * `xlsx:set-sheet-tab-color` — paint or clear a worksheet's tab
 * color. Mirrors Excel's "Tab Color" submenu on the sheet-tab
 * right-click menu.
 *
 * The color is normalized to uppercase 8-char ARGB hex (e.g.
 * `"FFCC0000"`). Passing `null` clears the color. The change marks
 * the workbook dirty so the serializer re-emits the worksheet XML;
 * `injectTabColor` in the serializer rewrites the `<sheetPr>` block.
 */
export const setSheetTabColorHandler: CommandHandler<SetSheetTabColorPayload, XlsxSnapshot> = {
  type: "xlsx:set-sheet-tab-color",
  apply(snapshot, payload) {
    const sheet = snapshot.root.sheets.find((s) => s.name === payload.name);
    if (!sheet) {
      throw new CommandError(
        "unknown-sheet",
        `Sheet "${payload.name}" not found; available: ${snapshot.root.sheets.map((s) => s.name).join(", ")}`
      );
    }

    let normalized: string | undefined;
    if (payload.color !== null) {
      const raw = payload.color.trim().replace(/^#/, "");
      if (!ARGB_RE.test(raw)) {
        throw new CommandError(
          "invalid-position",
          `Invalid tab color "${payload.color}"; expected 6- or 8-char hex (e.g. CC0000 or FFCC0000)`
        );
      }
      normalized = (raw.length === 6 ? `FF${raw}` : raw).toUpperCase();
    }

    const previous = sheet.tabColor;
    if (previous === normalized) {
      const next = evolveSnapshot(snapshot, snapshot.root, {});
      return {
        next,
        diff: buildDiff(snapshot.revision, next.revision, [
          {
            kind: "node-updated",
            nodeId: sheet.id,
            path: ["sheets", sheet.index, "tabColor"],
            field: "tabColor",
            summary: `set sheet tab color ${sheet.name} (no-op, already ${normalized ?? "none"})`,
          },
        ]),
      };
    }

    const updated: Sheet = normalized
      ? { ...sheet, tabColor: normalized }
      : (() => {
          const { tabColor: _omit, ...rest } = sheet;
          return rest as Sheet;
        })();
    const sheets = snapshot.root.sheets.slice();
    sheets[sheet.index] = updated;

    const nextWorkbook: XlsxWorkbook = { ...snapshot.root, sheets };
    const next = evolveSnapshot(snapshot, nextWorkbook, {
      workbook: true,
      sheets: [sheet.partPath],
    });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, [
        {
          kind: "node-updated",
          nodeId: sheet.id,
          path: ["sheets", sheet.index, "tabColor"],
          field: "tabColor",
          summary: `set sheet tab color ${sheet.name}: ${previous ?? "none"} → ${normalized ?? "none"}`,
        },
      ]),
    };
  },
};
