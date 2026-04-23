import { CommandError, type CommandHandler } from "@officeai/core";
import type { XlsxSnapshot, XlsxWorkbook } from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import type { SetWorkbookProtectionPayload } from "./payloads.js";

/**
 * `xlsx:set-workbook-protection` — toggle / configure the
 * `<workbookProtection .../>` element on the workbook root. Mirrors
 * Excel's "Protect Workbook" command on the Review tab.
 *
 * Same security stance as `set-sheet-protection`: this handler does
 * NOT hash plaintext passwords. Pass precomputed
 * `workbookPasswordHash` / `algorithmName` / `saltValue` /
 * `spinCount` if you want those attributes written; pass `enabled:
 * false` to drop the element entirely.
 *
 * Spec: ECMA-376 Part 1, §18.2.29 (workbookProtection).
 */
export const setWorkbookProtectionHandler: CommandHandler<SetWorkbookProtectionPayload, XlsxSnapshot> = {
  type: "xlsx:set-workbook-protection",
  apply(snapshot, payload) {
    const current = snapshot.root.workbookProtectionXml;

    if (!payload.enabled) {
      if (!current) {
        return { next: snapshot, diff: buildDiff(snapshot.revision, snapshot.revision, []) };
      }
      const root: XlsxWorkbook = { ...snapshot.root, workbookProtectionXml: undefined };
      const evolved = evolveSnapshot(snapshot, root, { workbook: true });
      return {
        next: evolved,
        diff: buildDiff(snapshot.revision, evolved.revision, [
          {
            kind: "node-updated",
            nodeId: snapshot.root.id,
            path: ["workbook", "protection"],
            field: "protection",
            summary: "disabled",
          },
        ]),
      };
    }

    const xml = buildWorkbookProtectionXml(payload);
    if (xml === current) {
      return { next: snapshot, diff: buildDiff(snapshot.revision, snapshot.revision, []) };
    }
    const root: XlsxWorkbook = { ...snapshot.root, workbookProtectionXml: xml };
    const evolved = evolveSnapshot(snapshot, root, { workbook: true });

    return {
      next: evolved,
      diff: buildDiff(snapshot.revision, evolved.revision, [
        {
          kind: "node-updated",
          nodeId: snapshot.root.id,
          path: ["workbook", "protection"],
          field: "protection",
          summary: "enabled",
        },
      ]),
    };
  },
};

function buildWorkbookProtectionXml(payload: SetWorkbookProtectionPayload): string {
  if ("password" in payload && payload.password !== undefined) {
    throw new CommandError(
      "invalid-payload",
      "Plaintext passwords are not accepted; pass a precomputed workbookPasswordHash + algorithmName + saltValue + spinCount instead."
    );
  }

  const attrs: string[] = [];
  if (payload.workbookPasswordHash) {
    attrs.push(`workbookPassword="${escapeXmlAttr(payload.workbookPasswordHash)}"`);
  }
  if (payload.algorithmName) {
    attrs.push(`workbookAlgorithmName="${escapeXmlAttr(payload.algorithmName)}"`);
    if (payload.hashValue) attrs.push(`workbookHashValue="${escapeXmlAttr(payload.hashValue)}"`);
    if (payload.saltValue) attrs.push(`workbookSaltValue="${escapeXmlAttr(payload.saltValue)}"`);
    if (payload.spinCount !== undefined) {
      attrs.push(`workbookSpinCount="${Math.max(0, Math.floor(payload.spinCount))}"`);
    }
  }
  if (payload.lockStructure !== undefined) {
    attrs.push(`lockStructure="${payload.lockStructure ? "1" : "0"}"`);
  }
  if (payload.lockWindows !== undefined) {
    attrs.push(`lockWindows="${payload.lockWindows ? "1" : "0"}"`);
  }
  if (payload.lockRevision !== undefined) {
    attrs.push(`lockRevision="${payload.lockRevision ? "1" : "0"}"`);
  }
  if (attrs.length === 0) attrs.push(`lockStructure="1"`);
  return `<workbookProtection ${attrs.join(" ")}/>`;
}

function escapeXmlAttr(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
