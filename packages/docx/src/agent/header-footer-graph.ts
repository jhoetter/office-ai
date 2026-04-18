import type {
  DocxSnapshot,
  HeaderFooterPart,
  HeaderFooterRef,
  Relationship,
  SectionBreak,
} from "../model/types.js";

/**
 * Resolved header/footer parts for a single section, keyed by the
 * three OOXML `w:type` slots (`default`, `first`, `even`).
 *
 * Lookup order at render time:
 *   - `titlePg === true` → use `first` for page 1 of the section,
 *     falling back to `default` when `first` is not set.
 *   - With `<w:settings><w:evenAndOddHeaders/>` (deferred to P4) →
 *     even-numbered pages use `even`, odd pages use `default`.
 *   - Otherwise every page uses `default`.
 */
export interface ResolvedHeaderFooter {
  readonly default?: HeaderFooterPart;
  readonly first?: HeaderFooterPart;
  readonly even?: HeaderFooterPart;
}

export interface ResolvedSectionHeaderFooters {
  readonly headers: ResolvedHeaderFooter;
  readonly footers: ResolvedHeaderFooter;
}

/**
 * Resolve the typed header/footer parts referenced by the section
 * break at `body[sectionIndex]`.
 *
 * Walks `SectionBreak.properties.headerRefs[].relationshipId` through
 * `relationships.get("word/document.xml")` to a `Relationship.target`
 * (e.g. `"header1.xml"`), then looks up the matching
 * {@link HeaderFooterPart} in `document.headersAndFooters`.
 *
 * Returns empty {@link ResolvedHeaderFooter} records when the index is
 * out of range, the body element is not a section break, the
 * relationship is missing, or the target part is not loaded — every
 * downstream caller already handles "no header/footer for this slot",
 * so undefined-tolerance is the contract here.
 */
export function resolveHeaderFooterParts(
  snapshot: DocxSnapshot,
  sectionIndex: number
): ResolvedSectionHeaderFooters {
  const empty: ResolvedSectionHeaderFooters = { headers: {}, footers: {} };
  const block = snapshot.root.body[sectionIndex];
  if (!block || block.kind !== "section-break") return empty;

  const docRels = snapshot.root.relationships.get("word/document.xml") ?? [];
  const relsById = new Map<string, Relationship>(docRels.map((r) => [r.id, r]));
  const partByPath = new Map<string, HeaderFooterPart>(
    snapshot.root.headersAndFooters.map((p) => [p.partPath, p])
  );

  return {
    headers: collectSlots(block, "headerRefs", relsById, partByPath, "header"),
    footers: collectSlots(block, "footerRefs", relsById, partByPath, "footer"),
  };
}

function collectSlots(
  section: SectionBreak,
  slot: "headerRefs" | "footerRefs",
  relsById: ReadonlyMap<string, Relationship>,
  partByPath: ReadonlyMap<string, HeaderFooterPart>,
  expectedKind: "header" | "footer"
): ResolvedHeaderFooter {
  const out: { -readonly [K in keyof ResolvedHeaderFooter]: ResolvedHeaderFooter[K] } = {};
  const refs: ReadonlyArray<HeaderFooterRef> = section.properties[slot];
  for (const ref of refs) {
    const rel = relsById.get(ref.relationshipId);
    if (!rel) continue;
    const path = resolveTargetPath(rel.target);
    const part = partByPath.get(path);
    if (!part || part.kind !== expectedKind) continue;
    out[ref.type] = part;
  }
  return out;
}

/**
 * `<w:headerReference w:type="default" r:id="rId4"/>` resolves through
 * `word/_rels/document.xml.rels` to a `Target` like `"header1.xml"`.
 * Header / footer parts live under `word/`, so we prepend that prefix
 * unless the relationship target is already absolute.
 */
function resolveTargetPath(target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  if (target.startsWith("word/")) return target;
  return `word/${target}`;
}
