import { CommandError, type CommandHandler } from "@officeai/core";
import type {
  BlockNode,
  DocxSnapshot,
  SectionBreak,
  SectionProperties,
} from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import type { SetSectionDifferentFirstPayload } from "./payloads.js";

/**
 * P3.4 / W16 — toggle `<w:titlePg/>` on the section that contains a
 * given paragraph index. Walks forward from `paragraphIndex` to the
 * next {@link SectionBreak} block; if none exists, falls back to the
 * trailing implicit section at the end of `body`.
 *
 * Mutating the typed `SectionProperties` drops the section's `raw`
 * cache so the serializer rebuilds the `<w:sectPr>` from
 * {@link serializeSectionProperties}. Untouched sections still
 * round-trip byte-identical.
 */
export const setSectionDifferentFirstHandler: CommandHandler<
  SetSectionDifferentFirstPayload,
  DocxSnapshot
> = {
  type: "docx:set-section-different-first",
  apply(snapshot, payload) {
    const { paragraphIndex, enabled } = payload;
    if (!Number.isInteger(paragraphIndex) || paragraphIndex < 0) {
      throw new CommandError(
        "invalid-payload",
        `paragraphIndex must be a non-negative integer (got ${paragraphIndex})`
      );
    }
    if (typeof enabled !== "boolean") {
      throw new CommandError("invalid-payload", `enabled must be a boolean (got ${typeof enabled})`);
    }

    const located = findOwningSection(snapshot, paragraphIndex);
    if (!located) {
      throw new CommandError(
        "unknown-target",
        `no section found at or after paragraph index ${paragraphIndex} (body has ${snapshot.root.body.length} blocks)`
      );
    }

    const currentTitlePg = located.section.properties.titlePg ?? false;
    if (currentTitlePg === enabled) {
      return {
        next: snapshot,
        diff: buildDiff(snapshot.revision, snapshot.revision, {
          kind: "node-updated",
          nodeId: located.section.id,
          path: ["body", located.index],
          field: "titlePg",
          summary: `no-op (already ${enabled})`,
        }),
      };
    }

    const nextProps = withTitlePg(located.section.properties, enabled);
    const updatedSection: SectionBreak = {
      ...located.section,
      properties: nextProps,
      // Drop raw — the typed model is the source of truth now.
      raw: undefined,
    };

    const newBody: BlockNode[] = snapshot.root.body.slice();
    newBody[located.index] = updatedSection;
    const next = evolveSnapshot(snapshot, { ...snapshot.root, body: newBody }, { body: true });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: located.section.id,
        path: ["body", located.index],
        field: "titlePg",
        summary: `titlePg := ${enabled}`,
      }),
    };
  },
};

interface LocatedSection {
  readonly index: number;
  readonly section: SectionBreak;
}

function findOwningSection(snapshot: DocxSnapshot, paragraphIndex: number): LocatedSection | null {
  const body = snapshot.root.body;
  // Walk forward from paragraphIndex looking for the next section break.
  for (let i = paragraphIndex; i < body.length; i++) {
    const block = body[i];
    if (block.kind === "section-break") {
      return { index: i, section: block };
    }
  }
  // Fall back to the last section-break in the body (trailing implicit
  // section). If no section-break exists at all, the document has no
  // typed sections to mutate — surface as an unknown target so the
  // caller knows nothing happened.
  for (let i = body.length - 1; i >= 0; i--) {
    const block = body[i];
    if (block.kind === "section-break") {
      return { index: i, section: block };
    }
  }
  return null;
}

function withTitlePg(props: SectionProperties, enabled: boolean): SectionProperties {
  if (enabled) {
    return { ...props, titlePg: true };
  }
  // Disabling: drop the field entirely so the serializer omits
  // `<w:titlePg/>` (Word's `false` is "absent" rather than
  // `<w:titlePg w:val="false"/>`).
  const next: { -readonly [K in keyof SectionProperties]: SectionProperties[K] } = { ...props };
  delete next.titlePg;
  return next;
}
