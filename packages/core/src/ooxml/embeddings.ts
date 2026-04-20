import { OoxmlContainer } from "./container.js";
import { ContentTypes } from "./content-types.js";
import { RelationshipGraph } from "./relationships.js";

/**
 * Shared OOXML embedding helpers used by docx + pptx serializers when
 * authoring chart-with-embedded-xlsx, OLE Excel objects, or any other
 * "package contains another package" scenario.
 *
 * Each helper does three things atomically:
 *   1. Add (or replace) the binary part inside the container.
 *   2. Register a `<Default Extension="…">` or `<Override PartName="…">`
 *      in `[Content_Types].xml` so the package validates.
 *   3. Add a relationship from the owning part to the new embedded part.
 *
 * Relationship type constants mirror the strings Microsoft Word /
 * PowerPoint write — deviating from these breaks "Edit Data" round-trip.
 */

export const REL_TYPE_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
export const REL_TYPE_PACKAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
export const REL_TYPE_OLE_OBJECT =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
export const REL_TYPE_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

export const CT_DRAWINGML_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
export const CT_SPREADSHEETML_SHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

export interface AddPackagePartArgs {
  /** OOXML container being mutated. */
  readonly container: OoxmlContainer;
  /** Pre-loaded content-types proxy (caller must `writeBack` after batching). */
  readonly contentTypes: ContentTypes;
  /** Path to the `*.rels` graph for the OWNER part (chart, slide, document, …). */
  readonly ownerRels: RelationshipGraph;
  /** Package-absolute path of the new embedded part (e.g. `ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx`). */
  readonly partPath: string;
  /** Bytes of the embedded part. */
  readonly bytes: Uint8Array;
  /** Override content type for the new part (e.g. `CT_SPREADSHEETML_SHEET`). */
  readonly contentType: string;
  /**
   * Relationship target as it should appear in the rels XML, RELATIVE to
   * the owner's directory. For example, when the owner is
   * `ppt/charts/chart1.xml` and the embedded part lives at
   * `ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx`, the target should
   * be `../embeddings/Microsoft_Excel_Worksheet1.xlsx`.
   */
  readonly relTarget: string;
  /** Relationship type. Use one of the `REL_TYPE_*` constants. */
  readonly relType: string;
  /** Optional explicit relationship id; otherwise `ownerRels.mintId()`. */
  readonly relId?: string;
}

export interface AddPackagePartResult {
  readonly relId: string;
  readonly partPath: string;
}

/**
 * Add a binary embedded part (xlsx package, OLE blob, image, …) to the
 * container, register its content-type override, and link it to the
 * owner via a relationship. Idempotent: if the same `relTarget` + type
 * is already wired from the owner, the existing relationship is reused
 * and only the part bytes are refreshed.
 */
export function addEmbeddedPart(args: AddPackagePartArgs): AddPackagePartResult {
  const { container, contentTypes, ownerRels, partPath, bytes, contentType, relTarget, relType } = args;

  if (container.has(partPath)) {
    container.writeBytes(partPath, bytes);
  } else {
    container.addPart(partPath, bytes);
  }

  const overridePartName = partPath.startsWith("/") ? partPath : `/${partPath}`;
  if (!contentTypes.hasOverride(overridePartName)) {
    contentTypes.addOverride(overridePartName, contentType);
  }

  const existing = ownerRels.relationships.find((r) => r.type === relType && r.target === relTarget);
  if (existing) {
    return { relId: existing.id, partPath };
  }
  const rel = ownerRels.add({
    ...(args.relId ? { id: args.relId } : {}),
    type: relType,
    target: relTarget,
  });
  return { relId: rel.id, partPath };
}

/**
 * Compute a package-absolute target path's relative form for use in
 * relationship `Target` attributes. Both inputs MUST be package-absolute
 * (no leading `/`).
 *
 * Example:
 *   relativeTarget("ppt/charts/chart1.xml", "ppt/embeddings/foo.xlsx")
 *     → "../embeddings/foo.xlsx"
 */
export function relativeTarget(ownerPartPath: string, targetPartPath: string): string {
  const owner = ownerPartPath.replace(/^\//, "").split("/");
  const target = targetPartPath.replace(/^\//, "").split("/");
  owner.pop();
  let i = 0;
  while (i < owner.length && i < target.length - 1 && owner[i] === target[i]) i++;
  const ups = owner.length - i;
  const rest = target.slice(i);
  const parts: string[] = [];
  for (let k = 0; k < ups; k++) parts.push("..");
  for (const r of rest) parts.push(r);
  return parts.join("/");
}

/**
 * Add a `<Default Extension="…">` entry if missing. Use for image
 * extensions (png/jpeg/…) that ride alongside OLE preview shapes; OLE
 * preview images don't need per-part overrides because the default
 * already covers them.
 */
export function ensureDefaultContentType(
  contentTypes: ContentTypes,
  extension: string,
  contentType: string
): void {
  if (!contentTypes.hasDefault(extension)) {
    contentTypes.addDefault(extension, contentType);
  }
}
