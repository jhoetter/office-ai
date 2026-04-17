import { ooxml } from "@officeai/core";
import type { BlockNode, DocxSnapshot, HeaderFooterPart } from "../model/types.js";
import { opaqueToEntry } from "../parser/xml-helpers.js";
import { DocxSerializeError } from "./errors.js";

/**
 * Serializer for `word/header*.xml` / `word/footer*.xml` parts.
 *
 * Byte-preservation contract:
 * - Untouched parts (those whose `partPath` is NOT in
 *   `snapshot.dirty.headersAndFooters`) are left alone — the cloned container
 *   already carries the original bytes, so the round-trip is byte-identical.
 * - Touched parts are re-serialized from the typed `HeaderFooterPart` model.
 *   `serializeBlock` is injected from the main serializer so we share the
 *   same paragraph emission code (and thereby the same OOXML quirks).
 */

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export function serializeHeaderFooterParts(
  container: ooxml.OoxmlContainer,
  snapshot: DocxSnapshot,
  serializeBlock: (block: BlockNode) => unknown
): void {
  if (snapshot.dirty.headersAndFooters.size === 0) return;
  const partsByPath = new Map<string, HeaderFooterPart>();
  for (const p of snapshot.root.headersAndFooters) {
    partsByPath.set(p.partPath, p);
  }
  for (const partPath of snapshot.dirty.headersAndFooters) {
    const part = partsByPath.get(partPath);
    if (!part) {
      // The part is dirty but no longer in the model — should not happen in
      // P1 (we don't support adding/removing header parts yet), but be loud
      // if it does so the bug surfaces in tests rather than corrupted output.
      throw new DocxSerializeError(
        "header-footer-missing",
        `header/footer part "${partPath}" marked dirty but missing from snapshot`
      );
    }
    if (!container.has(partPath)) {
      throw new DocxSerializeError(
        "header-footer-missing",
        `header/footer part "${partPath}" missing from container`
      );
    }
    const xml = serializeHeaderFooterXml(part, serializeBlock);
    container.writeText(partPath, xml);
  }
}

function serializeHeaderFooterXml(
  part: HeaderFooterPart,
  serializeBlock: (block: BlockNode) => unknown
): string {
  const tag = part.kind === "header" ? "w:hdr" : "w:ftr";
  const children: unknown[] = part.body.map((b) => serializeBlock(b));
  const entry: Record<string, unknown> = { [tag]: children };
  const rootAttrs: Record<string, string> = {};
  for (const [k, v] of Object.entries(part.rootAttrs)) {
    rootAttrs[`@_${k}`] = v;
  }
  if (Object.keys(rootAttrs).length > 0) {
    entry[":@"] = rootAttrs;
  }
  void opaqueToEntry; // re-export opaqueToEntry use is via serializeBlock
  return ooxml.serializeXml([entry], { xmlDeclaration: XML_DECL });
}
