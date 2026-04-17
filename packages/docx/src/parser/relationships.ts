import { ooxml } from "@officeai/core";
import type { Relationship } from "../model/types.js";

/**
 * Parse every `*.rels` part in the package into a typed map.
 *
 * Keying convention (mirrored on the way out): each entry is keyed by the
 * **owning** part path, NOT the rels part path. So
 * `word/_rels/document.xml.rels` is keyed by `"word/document.xml"`, and the
 * package-level `_rels/.rels` is keyed by the empty string `""`. This is
 * the same convention `RelationshipGraph.relsPathFor()` inverts.
 *
 * Why not just use `RelationshipGraph` directly from the typed model?
 * Because the typed model is the source of truth for whichever rels parts
 * a mutation touches; the rest must round-trip byte-identical, which
 * means we don't even re-emit them. Holding them in the typed map keeps
 * `DocxDocument.relationships` complete (so dirty flags can flip
 * individual parts without losing siblings) without committing to
 * re-emitting parts we never modified.
 */
export function parseRelationshipsParts(
  container: ooxml.OoxmlContainer
): Map<string, ReadonlyArray<Relationship>> {
  const out = new Map<string, ReadonlyArray<Relationship>>();
  for (const partPath of container.parts.keys()) {
    if (!isRelsPath(partPath)) continue;
    const ownerPath = ownerPathFor(partPath);
    const graph = ooxml.RelationshipGraph.loadFor(container, ownerPath);
    out.set(
      ownerPath,
      graph.relationships.map((r) => ({ ...r }))
    );
  }
  return out;
}

/**
 * Convert a rels-part path (e.g. `word/_rels/document.xml.rels`) to the
 * path of the part it describes (`word/document.xml`). The package-level
 * rels (`_rels/.rels`) describes the package itself; we represent that
 * with the empty string `""`.
 */
export function ownerPathFor(relsPath: string): string {
  if (relsPath === "_rels/.rels") return "";
  const slash = relsPath.lastIndexOf("/");
  const dir = slash >= 0 ? relsPath.slice(0, slash) : "";
  const fname = slash >= 0 ? relsPath.slice(slash + 1) : relsPath;
  if (!fname.endsWith(".rels")) return relsPath;
  const ownerName = fname.slice(0, -".rels".length);
  if (!dir.endsWith("/_rels")) return relsPath;
  const ownerDir = dir.slice(0, -"/_rels".length);
  return ownerDir.length > 0 ? `${ownerDir}/${ownerName}` : ownerName;
}

function isRelsPath(partPath: string): boolean {
  return partPath.endsWith(".rels") || partPath === "_rels/.rels";
}
