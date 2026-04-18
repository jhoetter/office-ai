import { CommandError, type CommandHandler } from "@officeai/core";
import type { DocxSnapshot, Paragraph, ParagraphProperties } from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import { locateParagraph } from "./set-paragraph-list.js";
import type { SetParagraphAlignmentPayload } from "./payloads.js";

const ALLOWED = new Set<string>(["left", "center", "right", "justify"]);

/**
 * Set (or clear) a paragraph's `<w:jc>` alignment. Like
 * `docx:set-paragraph-list`, the paragraph is identified by its stable
 * `NodeId` so the handler can reach paragraphs nested inside table
 * cells. Only `dirty.body` is set; the styles part is not touched.
 *
 * Setting the alignment to `null` clears the typed field AND drops any
 * stale `<w:jc>` opaque carrier left over from parse, so the
 * serializer round-trips back to "no alignment override".
 */
export const setParagraphAlignmentHandler: CommandHandler<SetParagraphAlignmentPayload, DocxSnapshot> = {
  type: "docx:set-paragraph-alignment",
  apply(snapshot, payload) {
    const { paragraphId, alignment } = payload;
    if (!paragraphId) {
      throw new CommandError("invalid-payload", "paragraphId is required");
    }
    if (alignment !== null && !ALLOWED.has(alignment)) {
      throw new CommandError(
        "invalid-payload",
        `alignment must be left|center|right|justify|null (got ${String(alignment)})`
      );
    }

    const located = locateParagraph(snapshot.root, paragraphId);
    if (!located) {
      throw new CommandError("unknown-target", `no paragraph with id "${paragraphId}"`);
    }

    const previous = located.paragraph.properties.alignment ?? null;
    if (previous === alignment) {
      // No-op writes still bump revision in evolveSnapshot, so short-circuit
      // to keep the dirty surface minimal — agents that "set to current
      // value" should not pay a re-serialize tax.
      return {
        next: snapshot,
        diff: buildDiff(snapshot.revision, snapshot.revision, {
          kind: "node-updated",
          nodeId: located.paragraph.id,
          path: located.path,
          field: "alignment",
          summary: `alignment unchanged (${previous ?? "(none)"})`,
        }),
      };
    }

    const updatedProps = applyAlignment(located.paragraph.properties, alignment);
    const updatedParagraph: Paragraph = { ...located.paragraph, properties: updatedProps };
    const nextDoc = located.replace(updatedParagraph);
    const next = evolveSnapshot(snapshot, nextDoc, { body: true });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: located.paragraph.id,
        path: located.path,
        field: "alignment",
        summary: `${previous ?? "(none)"} → ${alignment ?? "(none)"}`,
      }),
    };
  },
};

function applyAlignment(
  props: ParagraphProperties,
  alignment: "left" | "center" | "right" | "justify" | null
): ParagraphProperties {
  // Drop any `<w:jc>` opaque carrier so the typed field becomes the
  // single source of truth on re-serialize.
  const opaqueProps = props.opaqueProps?.filter((o) => o.tag !== "w:jc");
  const next: ParagraphProperties =
    alignment === null ? { ...props, alignment: undefined } : { ...props, alignment };
  const withOpaque: ParagraphProperties =
    opaqueProps && opaqueProps.length > 0 ? { ...next, opaqueProps } : { ...next, opaqueProps: undefined };
  return stripUndefinedAlignmentAndOpaque(withOpaque);
}

function stripUndefinedAlignmentAndOpaque(props: ParagraphProperties): ParagraphProperties {
  // Structural-equality round-trip: omit `alignment` / `opaqueProps`
  // entirely when they are undefined.
  const out: Record<string, unknown> = { ...props };
  if (out.alignment === undefined) delete out.alignment;
  if (out.opaqueProps === undefined) delete out.opaqueProps;
  return out as ParagraphProperties;
}
