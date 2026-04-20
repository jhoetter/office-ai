/**
 * DOCX embedded-binary serializer.
 *
 * Walks `dirty.embeddings` and (re)writes each embedded `.xlsx` /
 * OLE-blob into the container, and registers the matching content-type
 * override in `[Content_Types].xml` so Office recognises the part.
 *
 * The owning relationships (the `<Relationship Type=".../oleObject">`
 * entries that point at the embedded part from `word/document.xml.rels`)
 * are produced eagerly by the `docx:insert-spreadsheet` command and
 * round-trip through the existing `serializeRelationshipsParts` pass —
 * they're not touched here.
 */

import { ooxml } from "@officeai/core";
import { buildEmbeddedXlsx } from "@officeai/xlsx";
import type { DocxSnapshot } from "../model/types.js";
import { DocxSerializeError } from "./errors.js";

export async function serializeEmbeddingParts(
  container: ooxml.OoxmlContainer,
  snapshot: DocxSnapshot
): Promise<void> {
  if (snapshot.dirty.embeddings.size === 0) return;
  const contentTypes = ooxml.ContentTypes.load(container);
  let touchedContentTypes = false;

  for (const partPath of snapshot.dirty.embeddings) {
    const part = snapshot.root.embeddings.get(partPath);
    if (!part) {
      throw new DocxSerializeError(
        "missing-embedding",
        `dirty.embeddings references missing part ${partPath}`
      );
    }
    let bytes = part.bytes;
    if (!bytes && part.pendingGrid) {
      const built = await buildEmbeddedXlsx(part.pendingGrid, {
        sheetName: part.pendingSheetName ?? "Sheet1",
      });
      bytes = built.bytes;
    }
    if (!bytes) {
      throw new DocxSerializeError(
        "missing-embedding-bytes",
        `embedded part ${partPath} has neither bytes nor pendingGrid`
      );
    }
    if (container.has(partPath)) {
      container.writeBytes(partPath, bytes);
    } else {
      container.addPart(partPath, bytes);
    }
    const overrideName = partPath.startsWith("/") ? partPath : `/${partPath}`;
    if (!contentTypes.hasOverride(overrideName)) {
      contentTypes.addOverride(overrideName, part.contentType);
      touchedContentTypes = true;
    }
  }

  if (touchedContentTypes) contentTypes.writeBack(container);
}
