import type { CommandHandler } from "@officeai/core";
import type { Sheet, XlsxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { SetPrintOptionsPayload } from "./payloads.js";
import { resolveSheet } from "./validation.js";

/**
 * `xlsx:set-print-options` — patch the worksheet's `<printOptions>`
 * element. Mirrors Page Layout → Sheet Options (gridlines, headings,
 * centering on page).
 *
 * Patches preserve untouched attributes. Pass `clear: true` to drop
 * the element entirely.
 */
export const setPrintOptionsHandler: CommandHandler<SetPrintOptionsPayload, XlsxSnapshot> = {
  type: "xlsx:set-print-options",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);

    if (payload.clear) {
      if (!sheet.printOptionsXml) {
        return { next: snapshot, diff: buildDiff(snapshot.revision, snapshot.revision, []) };
      }
      const next: Sheet = { ...sheet, printOptionsXml: undefined };
      return commit(snapshot, sheet, next, "cleared");
    }

    const attrs = parseAttrs(sheet.printOptionsXml);
    setBool(attrs, "horizontalCentered", payload.horizontalCentered);
    setBool(attrs, "verticalCentered", payload.verticalCentered);
    setBool(attrs, "headings", payload.headings);
    setBool(attrs, "gridLines", payload.gridLines);
    setBool(attrs, "gridLinesSet", payload.gridLinesSet);

    const xml = attrs.size === 0 ? undefined : buildElement(attrs);
    if ((xml ?? "") === (sheet.printOptionsXml ?? "")) {
      return { next: snapshot, diff: buildDiff(snapshot.revision, snapshot.revision, []) };
    }
    const next: Sheet = { ...sheet, printOptionsXml: xml };
    return commit(snapshot, sheet, next, summarise(payload));
  },
};

function commit(
  snapshot: XlsxSnapshot,
  before: Sheet,
  after: Sheet,
  summary: string
): { next: XlsxSnapshot; diff: ReturnType<typeof buildDiff> } {
  const root = replaceSheet(snapshot.root, after);
  const evolved = evolveSnapshot(snapshot, root, { sheets: [before.partPath] });
  return {
    next: evolved,
    diff: buildDiff(snapshot.revision, evolved.revision, [
      {
        kind: "node-updated",
        nodeId: before.id,
        path: ["sheets", before.index, "printOptions"],
        field: "printOptions",
        summary,
      },
    ]),
  };
}

function parseAttrs(xml: string | undefined): Map<string, string> {
  const out = new Map<string, string>();
  if (!xml) return out;
  const attrRe = /\b([A-Za-z]+)\s*=\s*"([^"]*)"/g;
  let m: RegExpExecArray | null;
  while ((m = attrRe.exec(xml)) !== null) {
    out.set(m[1], m[2]);
  }
  return out;
}

function setBool(attrs: Map<string, string>, name: string, value: boolean | undefined): void {
  if (value === undefined) return;
  if (value) attrs.set(name, "1");
  else attrs.delete(name);
}

function buildElement(attrs: Map<string, string>): string {
  const out: string[] = [];
  for (const [k, v] of attrs) out.push(`${k}="${v}"`);
  return `<printOptions ${out.join(" ")}/>`;
}

function summarise(p: SetPrintOptionsPayload): string {
  const parts: string[] = [];
  if (p.horizontalCentered !== undefined) parts.push(`hc=${p.horizontalCentered}`);
  if (p.verticalCentered !== undefined) parts.push(`vc=${p.verticalCentered}`);
  if (p.headings !== undefined) parts.push(`hd=${p.headings}`);
  if (p.gridLines !== undefined) parts.push(`gl=${p.gridLines}`);
  return parts.join(",") || "patched";
}
