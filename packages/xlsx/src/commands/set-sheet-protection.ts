import { CommandError, type CommandHandler } from "@officeai/core";
import type { Sheet, XlsxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { SetSheetProtectionPayload } from "./payloads.js";
import { resolveSheet } from "./validation.js";

/**
 * `xlsx:set-sheet-protection` — toggle / configure the
 * `<sheetProtection .../>` element on a worksheet. Mirrors Excel's
 * "Protect Sheet" dialog under the Review tab.
 *
 * The element is built from the typed payload and stamped into
 * `sheet.sheetProtectionXml`; the existing serializer already injects
 * that opaque field after `<sheetViews>`, so we get round-trip and
 * preservation of every other sheet element for free.
 *
 * No password hashing is performed — agents pass `passwordHash` /
 * `algorithmName` / `saltValue` / `spinCount` directly when they
 * already have a hash to write (mirroring the OOXML attributes); the
 * `password` field on the payload is rejected to avoid producing an
 * insecure plaintext-equivalent. Use `enabled: false` to drop the
 * `<sheetProtection>` element entirely (matches Excel's "Unprotect
 * Sheet" with no password).
 *
 * Spec: ECMA-376 Part 1, §18.3.1.85 (sheetProtection).
 */
export const setSheetProtectionHandler: CommandHandler<SetSheetProtectionPayload, XlsxSnapshot> = {
  type: "xlsx:set-sheet-protection",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);

    if (!payload.enabled) {
      if (!sheet.sheetProtectionXml) {
        return {
          next: snapshot,
          diff: buildDiff(snapshot.revision, snapshot.revision, []),
        };
      }
      const next: Sheet = { ...sheet, sheetProtectionXml: undefined };
      const root = replaceSheet(snapshot.root, next);
      const evolved = evolveSnapshot(snapshot, root, { sheets: [sheet.partPath] });
      return {
        next: evolved,
        diff: buildDiff(snapshot.revision, evolved.revision, [
          {
            kind: "node-updated",
            nodeId: sheet.id,
            path: ["sheets", sheet.index, "protection"],
            field: "protection",
            summary: "disabled",
          },
        ]),
      };
    }

    const xml = buildSheetProtectionXml(payload);
    if (xml === sheet.sheetProtectionXml) {
      return { next: snapshot, diff: buildDiff(snapshot.revision, snapshot.revision, []) };
    }
    const next: Sheet = { ...sheet, sheetProtectionXml: xml };
    const root = replaceSheet(snapshot.root, next);
    const evolved = evolveSnapshot(snapshot, root, { sheets: [sheet.partPath] });

    return {
      next: evolved,
      diff: buildDiff(snapshot.revision, evolved.revision, [
        {
          kind: "node-updated",
          nodeId: sheet.id,
          path: ["sheets", sheet.index, "protection"],
          field: "protection",
          summary: "enabled",
        },
      ]),
    };
  },
};

function buildSheetProtectionXml(payload: SetSheetProtectionPayload): string {
  if ("password" in payload && payload.password !== undefined) {
    throw new CommandError(
      "invalid-payload",
      "Plaintext passwords are not accepted; pass a precomputed passwordHash + algorithmName + saltValue + spinCount instead."
    );
  }

  const attrs: string[] = [`sheet="1"`];
  if (payload.passwordHash) {
    attrs.push(`password="${escapeXmlAttr(payload.passwordHash)}"`);
  }
  if (payload.algorithmName) {
    attrs.push(`algorithmName="${escapeXmlAttr(payload.algorithmName)}"`);
    if (payload.hashValue) attrs.push(`hashValue="${escapeXmlAttr(payload.hashValue)}"`);
    if (payload.saltValue) attrs.push(`saltValue="${escapeXmlAttr(payload.saltValue)}"`);
    if (payload.spinCount !== undefined) {
      attrs.push(`spinCount="${Math.max(0, Math.floor(payload.spinCount))}"`);
    }
  }
  for (const [flag, value] of [
    ["objects", payload.objects],
    ["scenarios", payload.scenarios],
    ["formatCells", payload.formatCells],
    ["formatColumns", payload.formatColumns],
    ["formatRows", payload.formatRows],
    ["insertColumns", payload.insertColumns],
    ["insertRows", payload.insertRows],
    ["insertHyperlinks", payload.insertHyperlinks],
    ["deleteColumns", payload.deleteColumns],
    ["deleteRows", payload.deleteRows],
    ["selectLockedCells", payload.selectLockedCells],
    ["sort", payload.sort],
    ["autoFilter", payload.autoFilter],
    ["pivotTables", payload.pivotTables],
    ["selectUnlockedCells", payload.selectUnlockedCells],
  ] as const) {
    if (value === undefined) continue;
    attrs.push(`${flag}="${value ? "1" : "0"}"`);
  }
  return `<sheetProtection ${attrs.join(" ")}/>`;
}

function escapeXmlAttr(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
