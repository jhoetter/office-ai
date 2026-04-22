import { CommandError, type CommandHandler, type IdMinter } from "@officeai/core";
import type {
  BlockNode,
  DocxDocument,
  DocxSnapshot,
  HeaderFooterPart,
  HeaderFooterRef,
  Paragraph,
  Relationship,
  SectionBreak,
  SectionProperties,
} from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import type { CreateHeaderFooterPartPayload } from "./payloads.js";

const HEADER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
const FOOTER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
const DOC_RELS_KEY = "word/document.xml";

/**
 * Mint a fresh, empty header or footer part and attach it to the
 * trailing implicit section as a `default` (or `first` / `even`) slot.
 *
 * See {@link CreateHeaderFooterPartPayload} for full semantics. The
 * command is idempotent: when the trailing section already declares
 * a slot of the requested kind/target, it returns a same-revision
 * no-op diff so callers can wire it to a "double-click empty zone"
 * affordance and not worry about repeated invocations.
 */
export const createHeaderFooterPartHandler: CommandHandler<CreateHeaderFooterPartPayload, DocxSnapshot> = {
  type: "docx:create-header-footer-part",
  apply(snapshot, payload, ctx) {
    const slot = payload.slot;
    if (slot !== "header" && slot !== "footer") {
      throw new CommandError("invalid-payload", `slot must be "header" or "footer" (got ${String(slot)})`);
    }
    const target = payload.target ?? "default";
    if (target !== "default" && target !== "first" && target !== "even") {
      throw new CommandError(
        "invalid-payload",
        `target must be "default" | "first" | "even" (got ${String(target)})`
      );
    }

    const located = findTrailingSection(snapshot);
    if (!located) {
      throw new CommandError("unknown-target", `document has no <w:sectPr> to attach a ${slot} reference to`);
    }

    // Idempotency: if the section already binds the requested slot/target
    // to an existing part, return a no-op diff so double-firing the
    // command (e.g. on every focus) is safe.
    const existing = findExistingPart(snapshot, located.section, slot, target);
    if (existing) {
      return {
        next: snapshot,
        diff: buildDiff(snapshot.revision, snapshot.revision, {
          kind: "node-updated",
          nodeId: existing.id,
          path: ["headersAndFooters"],
          field: "noop",
          summary: `${slot} (${target}) part already exists at ${existing.partPath}`,
        }),
      };
    }

    const partPath = mintPartPath(snapshot, slot);
    const docRels = snapshot.root.relationships.get(DOC_RELS_KEY) ?? [];
    const relId = mintRelId(docRels);
    const relTarget = relativeTargetFromMain(partPath);
    const relType = slot === "header" ? HEADER_REL_TYPE : FOOTER_REL_TYPE;
    const newRel: Relationship = { id: relId, type: relType, target: relTarget };
    const nextDocRels: ReadonlyArray<Relationship> = [...docRels, newRel];

    const newRef: HeaderFooterRef = { type: target, relationshipId: relId };
    const nextSectionProps: SectionProperties =
      slot === "header"
        ? { ...located.section.properties, headerRefs: [...located.section.properties.headerRefs, newRef] }
        : { ...located.section.properties, footerRefs: [...located.section.properties.footerRefs, newRef] };
    const updatedSection: SectionBreak = {
      ...located.section,
      properties: nextSectionProps,
      raw: undefined,
    };
    const newBody: BlockNode[] = snapshot.root.body.slice();
    newBody[located.index] = updatedSection;

    const emptyParagraph = makeEmptyParagraph(ctx.mintNodeId);
    const newPart: HeaderFooterPart = {
      kind: slot,
      id: partPath,
      partPath,
      target,
      rootAttrs: defaultRootAttrs(slot),
      body: [emptyParagraph],
    };
    const nextParts: ReadonlyArray<HeaderFooterPart> = [...snapshot.root.headersAndFooters, newPart];

    const nextRels = new Map(snapshot.root.relationships);
    nextRels.set(DOC_RELS_KEY, nextDocRels);

    const nextDoc: DocxDocument = {
      ...snapshot.root,
      body: newBody,
      headersAndFooters: nextParts,
      relationships: nextRels,
    };

    const nextRelsDirty = withAddition(snapshot.dirty.relationships, DOC_RELS_KEY);
    const nextHfDirty = withAddition(snapshot.dirty.headersAndFooters, partPath);

    const next = evolveSnapshot(snapshot, nextDoc, {
      body: true,
      relationships: nextRelsDirty,
      headersAndFooters: nextHfDirty,
      contentTypes: true,
    });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-inserted",
        nodeId: newPart.id,
        path: ["headersAndFooters", nextParts.length - 1],
        summary: `+${slot} part ${partPath} (rel ${relId}, target ${target})`,
      }),
    };
  },
};

interface LocatedSection {
  readonly index: number;
  readonly section: SectionBreak;
}

/**
 * Find the trailing implicit `<w:sectPr>` block (Word stores it as the
 * last section break in the body). Mirrors the lookup used by
 * `set-page-setup` / `set-section-different-first` so all
 * section-targeting commands agree on which section they mutate when
 * the caller doesn't pin a paragraph.
 */
function findTrailingSection(snapshot: DocxSnapshot): LocatedSection | null {
  const body = snapshot.root.body;
  for (let i = body.length - 1; i >= 0; i--) {
    const block = body[i];
    if (block.kind === "section-break") return { index: i, section: block };
  }
  return null;
}

function findExistingPart(
  snapshot: DocxSnapshot,
  section: SectionBreak,
  slot: "header" | "footer",
  target: HeaderFooterRef["type"]
): HeaderFooterPart | undefined {
  const refs = slot === "header" ? section.properties.headerRefs : section.properties.footerRefs;
  const ref = refs.find((r) => r.type === target);
  if (!ref) return undefined;
  const docRels = snapshot.root.relationships.get(DOC_RELS_KEY) ?? [];
  const rel = docRels.find((r) => r.id === ref.relationshipId);
  if (!rel) return undefined;
  const partPath = `word/${rel.target.replace(/^\.\//, "")}`;
  return snapshot.root.headersAndFooters.find((p) => p.partPath === partPath && p.kind === slot);
}

function mintPartPath(snapshot: DocxSnapshot, slot: "header" | "footer"): string {
  const taken = new Set<string>();
  for (const p of snapshot.root.headersAndFooters) taken.add(p.partPath);
  let i = 1;
  while (taken.has(`word/${slot}${i}.xml`)) i++;
  return `word/${slot}${i}.xml`;
}

function mintRelId(rels: ReadonlyArray<Relationship>): string {
  const taken = new Set(rels.map((r) => r.id));
  let i = rels.length + 1;
  while (taken.has(`rId${i}`)) i++;
  return `rId${i}`;
}

/**
 * Build the relationship `Target` value for a part referenced from
 * `word/document.xml`. The rels file lives in `word/_rels/` so paths
 * are written relative to `word/` (e.g. `header1.xml`, not
 * `word/header1.xml`).
 */
function relativeTargetFromMain(partPath: string): string {
  return partPath.startsWith("word/") ? partPath.slice("word/".length) : partPath;
}

function makeEmptyParagraph(mintNodeId: IdMinter): Paragraph {
  return {
    kind: "paragraph",
    id: mintNodeId(),
    properties: {},
    children: [],
  };
}

/**
 * Default namespace declarations for a fresh `<w:hdr>` / `<w:ftr>`
 * root. We mirror what Word emits (the `w` prefix is mandatory; `r`
 * is included so future page-number / image authoring can splice in
 * `r:id` attributes without re-declaring the namespace).
 */
function defaultRootAttrs(slot: "header" | "footer"): Readonly<Record<string, string>> {
  void slot;
  return {
    "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  };
}

function withAddition(prev: ReadonlySet<string>, member: string): ReadonlySet<string> {
  const next = new Set(prev);
  next.add(member);
  return next;
}
