import { CommandError, type CommandHandler } from "@officeai/core";
import type { DocxSnapshot, SectionBreak, SectionProperties } from "../model/types.js";
import { buildDiff, evolveSnapshot, insertBlock } from "./helpers.js";
import type { InsertSectionBreakPayload } from "./payloads.js";

/**
 * P3.4 / W17 — splice a new {@link SectionBreak} into the body at
 * `paragraphIndex`. The new break inherits geometry (page size,
 * margins, columns, header/footer refs) from the *next* section
 * forward, so the visual page chrome above the new break stays
 * unchanged. The break's `<w:type>` defaults to `"nextPage"`,
 * matching Word's "Insert → Section Break → Next Page" UI.
 *
 * Inserting at `paragraphIndex === body.length` appends just before
 * the trailing implicit section. The trailing section, if it exists,
 * is preserved verbatim — it always ends the body.
 */
export const insertSectionBreakHandler: CommandHandler<InsertSectionBreakPayload, DocxSnapshot> = {
  type: "docx:insert-section-break",
  apply(snapshot, payload, ctx) {
    const { paragraphIndex } = payload;
    const type = payload.type ?? "nextPage";
    if (!Number.isInteger(paragraphIndex) || paragraphIndex < 0) {
      throw new CommandError(
        "invalid-payload",
        `paragraphIndex must be a non-negative integer (got ${paragraphIndex})`
      );
    }
    if (paragraphIndex > snapshot.root.body.length) {
      throw new CommandError(
        "invalid-payload",
        `paragraphIndex ${paragraphIndex} out of range (body has ${snapshot.root.body.length} blocks)`
      );
    }
    if (type !== "nextPage" && type !== "continuous" && type !== "evenPage" && type !== "oddPage") {
      throw new CommandError("invalid-payload", `unknown section type "${type}"`);
    }

    const inheritedProps = inheritFromNextSection(snapshot, paragraphIndex);
    const newSection: SectionBreak = {
      kind: "section-break",
      id: ctx.mintNodeId(),
      properties: { ...inheritedProps, sectionType: type },
    };

    const nextDoc = insertBlock(snapshot.root, paragraphIndex, newSection);
    const next = evolveSnapshot(snapshot, nextDoc, { body: true });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: snapshot.root.id,
        path: ["body", paragraphIndex],
        field: "section-break",
        summary: `inserted ${type} section break at index ${paragraphIndex}`,
      }),
    };
  },
};

/**
 * Find the first {@link SectionBreak} at or after `paragraphIndex`
 * and return a structural clone of its properties so the new break
 * inherits the same page chrome. When no forward section exists the
 * caller still gets a sensible default (an empty `headerRefs` /
 * `footerRefs` skeleton) — that path is unreachable for documents
 * that came through the parser (every doc ends with a sectPr) but
 * matters for synthesised snapshots in tests.
 */
function inheritFromNextSection(snapshot: DocxSnapshot, paragraphIndex: number): SectionProperties {
  const body = snapshot.root.body;
  for (let i = paragraphIndex; i < body.length; i++) {
    const b = body[i];
    if (b.kind === "section-break") {
      return cloneSectionProperties(b.properties);
    }
  }
  // Walk backward — there might be a previous section we can copy
  // from (e.g. the trailing implicit section sits at index N-1).
  for (let i = body.length - 1; i >= 0; i--) {
    const b = body[i];
    if (b.kind === "section-break") {
      return cloneSectionProperties(b.properties);
    }
  }
  return { headerRefs: [], footerRefs: [] };
}

function cloneSectionProperties(p: SectionProperties): SectionProperties {
  return {
    ...(p.pgSz ? { pgSz: { ...p.pgSz } } : {}),
    ...(p.pgMar ? { pgMar: { ...p.pgMar } } : {}),
    ...(p.cols ? { cols: { ...p.cols } } : {}),
    headerRefs: p.headerRefs.map((r) => ({ ...r })),
    footerRefs: p.footerRefs.map((r) => ({ ...r })),
    ...(p.titlePg ? { titlePg: p.titlePg } : {}),
    ...(p.sectionType ? { sectionType: p.sectionType } : {}),
    ...(p.opaqueProps ? { opaqueProps: p.opaqueProps.map((o) => ({ ...o })) } : {}),
  };
}
