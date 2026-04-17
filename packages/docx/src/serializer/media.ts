import type { ooxml } from "@officeai/core";
import type { DocxSnapshot } from "../model/types.js";
import { DocxSerializeError } from "./errors.js";

/**
 * Write back media parts that have been added or changed since load.
 * Untouched parts are not in `snapshot.dirty.media` and so are left alone
 * — the cloned container already carries their original bytes, which
 * preserves byte-equality on round-trip.
 *
 * For each dirty media part:
 * - If the container does not have it, add it.
 * - If the container already has it, overwrite it. (We don't expect this
 *   path in P1.3 / W8 — `docx:insert-image` always picks fresh part
 *   paths — but it costs nothing and keeps the contract simple.)
 *
 * If a media path is in `dirty.media` but missing from the typed map we
 * surface a hard error: that combination would mean the snapshot was
 * mutated without going through the supported helpers and the package
 * would be inconsistent.
 */
export function serializeMediaParts(container: ooxml.OoxmlContainer, snapshot: DocxSnapshot): void {
  if (snapshot.dirty.media.size === 0) return;
  for (const partPath of snapshot.dirty.media) {
    const media = snapshot.root.media.get(partPath);
    if (!media) {
      throw new DocxSerializeError(
        "media-missing",
        `media part "${partPath}" marked dirty but missing from snapshot.root.media`
      );
    }
    if (container.has(partPath)) {
      container.writeBytes(partPath, media.bytes);
    } else {
      container.addPart(partPath, media.bytes);
    }
  }
}
