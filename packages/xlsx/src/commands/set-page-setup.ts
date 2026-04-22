import { CommandError, type CommandHandler } from "@officeai/core";
import type { Sheet, XlsxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { SetPageSetupPayload } from "./payloads.js";
import { resolveSheet } from "./validation.js";

/**
 * `xlsx:set-page-setup` — patch the worksheet's `<pageSetup>` element
 * (Page Layout tab → Page Setup group). Mirrors Excel's orientation,
 * paper-size, scaling, and fit-to-pages knobs.
 *
 * Strategy: parse the existing `pageSetupXml` opaque blob into an attr
 * map, apply the supplied patch, and re-serialize. Untouched
 * attributes round-trip verbatim. Pass `clear: true` to drop the
 * element entirely (revert to Excel defaults).
 */
export const setPageSetupHandler: CommandHandler<SetPageSetupPayload, XlsxSnapshot> = {
  type: "xlsx:set-page-setup",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);

    if (payload.clear) {
      if (!sheet.pageSetupXml) {
        return { next: snapshot, diff: buildDiff(snapshot.revision, snapshot.revision, []) };
      }
      const next: Sheet = { ...sheet, pageSetupXml: undefined };
      return commit(snapshot, sheet, next, "cleared");
    }

    if (
      payload.scale !== undefined &&
      (payload.scale < 10 || payload.scale > 400 || !Number.isFinite(payload.scale))
    ) {
      throw new CommandError(
        "invalid-payload",
        `set-page-setup: scale must be 10–400 (got ${payload.scale}).`,
      );
    }

    const attrs = parseAttrs(sheet.pageSetupXml);
    applyPatch(attrs, payload);

    const xml = buildElement(attrs);
    if (xml === sheet.pageSetupXml) {
      return { next: snapshot, diff: buildDiff(snapshot.revision, snapshot.revision, []) };
    }
    const next: Sheet = { ...sheet, pageSetupXml: xml };
    return commit(snapshot, sheet, next, summarise(payload));
  },
};

function commit(
  snapshot: XlsxSnapshot,
  before: Sheet,
  after: Sheet,
  summary: string,
): { next: XlsxSnapshot; diff: ReturnType<typeof buildDiff> } {
  const root = replaceSheet(snapshot.root, after);
  const evolved = evolveSnapshot(snapshot, root, { sheets: [before.partPath] });
  return {
    next: evolved,
    diff: buildDiff(snapshot.revision, evolved.revision, [
      {
        kind: "node-updated",
        nodeId: before.id,
        path: ["sheets", before.index, "pageSetup"],
        field: "pageSetup",
        summary,
      },
    ]),
  };
}

function parseAttrs(xml: string | undefined): Map<string, string> {
  const out = new Map<string, string>();
  if (!xml) return out;
  // Match attributes inside the (self-closing) <pageSetup .../> blob.
  const attrRe = /\b([A-Za-z][A-Za-z0-9]*)\s*=\s*"([^"]*)"/g;
  let m: RegExpExecArray | null;
  while ((m = attrRe.exec(xml)) !== null) {
    out.set(m[1], m[2]);
  }
  return out;
}

function applyPatch(attrs: Map<string, string>, p: SetPageSetupPayload): void {
  if (p.orientation !== undefined) attrs.set("orientation", p.orientation);
  if (p.paperSize !== undefined) attrs.set("paperSize", String(Math.floor(p.paperSize)));
  if (p.scale !== undefined) attrs.set("scale", String(Math.floor(p.scale)));
  if (p.fitToWidth !== undefined) attrs.set("fitToWidth", String(Math.max(0, Math.floor(p.fitToWidth))));
  if (p.fitToHeight !== undefined) attrs.set("fitToHeight", String(Math.max(0, Math.floor(p.fitToHeight))));
  if (p.firstPageNumber !== undefined) attrs.set("firstPageNumber", String(Math.floor(p.firstPageNumber)));
  if (p.useFirstPageNumber !== undefined) attrs.set("useFirstPageNumber", p.useFirstPageNumber ? "1" : "0");
  if (p.horizontalDpi !== undefined) attrs.set("horizontalDpi", String(Math.max(0, Math.floor(p.horizontalDpi))));
  if (p.verticalDpi !== undefined) attrs.set("verticalDpi", String(Math.max(0, Math.floor(p.verticalDpi))));
  if (p.blackAndWhite !== undefined) attrs.set("blackAndWhite", p.blackAndWhite ? "1" : "0");
  if (p.draft !== undefined) attrs.set("draft", p.draft ? "1" : "0");
  if (p.cellComments !== undefined) attrs.set("cellComments", p.cellComments);
  if (p.errors !== undefined) attrs.set("errors", p.errors);
}

function buildElement(attrs: Map<string, string>): string {
  if (attrs.size === 0) {
    return `<pageSetup/>`;
  }
  const out: string[] = [];
  for (const [k, v] of attrs) {
    out.push(`${k}="${escapeXmlAttr(v)}"`);
  }
  return `<pageSetup ${out.join(" ")}/>`;
}

function escapeXmlAttr(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function summarise(p: SetPageSetupPayload): string {
  const parts: string[] = [];
  if (p.orientation) parts.push(p.orientation);
  if (p.paperSize !== undefined) parts.push(`paper=${p.paperSize}`);
  if (p.scale !== undefined) parts.push(`scale=${p.scale}`);
  if (p.fitToWidth !== undefined || p.fitToHeight !== undefined) {
    parts.push(`fit=${p.fitToWidth ?? "?"}×${p.fitToHeight ?? "?"}`);
  }
  return parts.join(",") || "patched";
}
