import type { XlsxSnapshot } from "../model/types.js";
import { XlsxSerializeError } from "./errors.js";

/**
 * Serialize an `XlsxSnapshot` back to bytes.
 *
 * Phase 4 contract:
 *   - Untouched parts: re-emitted verbatim from `OoxmlContainer.parts`.
 *     The container guarantees byte-content preservation (zip-archive
 *     bytes may differ; part-content bytes are byte-identical).
 *   - Dirty parts: re-emitted from the typed model. Phase 4 ships no
 *     command handlers, so no dirty flags are ever set; any flag the
 *     caller hand-sets is currently a no-op (with a typed error
 *     surface so Phase 5 can wire in real serialization without
 *     changing the public signature).
 *
 * Phase 5 fills in the per-part re-emission paths for dirty sheets,
 * sharedStrings, styles, and rels.
 */
export async function serializeXlsx(snapshot: XlsxSnapshot): Promise<ArrayBuffer> {
  const container = snapshot.container.clone();

  const dirty = snapshot.dirty;
  const anyDirty =
    dirty.workbook ||
    dirty.sharedStrings ||
    dirty.styles ||
    dirty.contentTypes ||
    dirty.rels ||
    dirty.sheets.size > 0 ||
    dirty.comments.size > 0 ||
    dirty.threadedComments.size > 0 ||
    dirty.sheetRels.size > 0;

  if (anyDirty) {
    throw new XlsxSerializeError(
      "container-failed",
      "Phase 4 serializer cannot re-emit dirty parts; command handlers ship in Phase 5"
    );
  }

  try {
    return await container.serialize();
  } catch (err) {
    throw new XlsxSerializeError("container-failed", "Failed to re-emit OOXML container", { cause: err });
  }
}
