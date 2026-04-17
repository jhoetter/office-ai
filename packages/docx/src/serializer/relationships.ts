import { ooxml } from "@officeai/core";
import type { DocxSnapshot } from "../model/types.js";
import { DocxSerializeError } from "./errors.js";

/**
 * Re-emit the rels parts that any mutation has touched (currently only
 * `docx:insert-image`, which adds an `image` relationship to
 * `word/_rels/document.xml.rels`). Untouched rels parts are left alone in
 * the cloned container so they round-trip byte-identical.
 *
 * The typed map on `DocxDocument.relationships` is the single source of
 * truth for any part listed in `snapshot.dirty.relationships`. We rebuild
 * the rels XML from scratch via `RelationshipGraph.writeBack` — that path
 * already produces canonical Word-compatible XML.
 */
export function serializeRelationshipsParts(container: ooxml.OoxmlContainer, snapshot: DocxSnapshot): void {
  if (snapshot.dirty.relationships.size === 0) return;
  for (const ownerPath of snapshot.dirty.relationships) {
    const rels = snapshot.root.relationships.get(ownerPath);
    if (!rels) {
      throw new DocxSerializeError(
        "rels-missing",
        `relationships for "${ownerPath}" marked dirty but missing from snapshot`
      );
    }
    const relsPath = ownerPath === "" ? "_rels/.rels" : ooxml.RelationshipGraph.relsPathFor(ownerPath);
    const graph = new ooxml.RelationshipGraph(
      relsPath,
      rels.map((r) => ({ ...r }))
    );
    graph.writeBack(container);
  }
}
