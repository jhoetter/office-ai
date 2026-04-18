import { CommandError, type CommandHandler } from "@officeai/core";
import type { DocxSnapshot, Paragraph, ParagraphProperties } from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import { locateParagraph } from "./set-paragraph-list.js";
import type { RemoveParagraphListPayload } from "./payloads.js";

/**
 * Strip a paragraph's list reference, demoting it back to a normal
 * paragraph. Strict on no-op: if the paragraph is NOT currently a list
 * item, the command rejects with `not-applicable` (rather than
 * silently succeeding) so callers always know whether the document
 * actually changed. The choice is documented in the W10 build-log
 * entry — a future "make this idempotent" mode can dispatch through a
 * different command if we need it.
 */
export const removeParagraphListHandler: CommandHandler<RemoveParagraphListPayload, DocxSnapshot> = {
  type: "docx:remove-paragraph-list",
  apply(snapshot, payload) {
    const { paragraphId } = payload;
    if (!paragraphId) {
      throw new CommandError("invalid-payload", "paragraphId is required");
    }
    const located = locateParagraph(snapshot.root, paragraphId);
    if (!located) {
      throw new CommandError("unknown-target", `no paragraph with id "${paragraphId}"`);
    }
    if (!located.paragraph.properties.numbering) {
      throw new CommandError("not-applicable", `paragraph "${paragraphId}" is not currently a list item`);
    }

    const updatedProps = stripNumbering(located.paragraph.properties);
    const updatedParagraph: Paragraph = { ...located.paragraph, properties: updatedProps };
    const nextDoc = located.replace(updatedParagraph);
    const next = evolveSnapshot(snapshot, nextDoc, { body: true });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: located.paragraph.id,
        path: located.path,
        field: "numbering",
        summary: "remove list",
      }),
    };
  },
};

function stripNumbering(props: ParagraphProperties): ParagraphProperties {
  const { numbering: _drop, opaqueProps, ...rest } = props;
  void _drop;
  const filtered = opaqueProps?.filter((o) => o.tag !== "w:numPr");
  if (filtered && filtered.length > 0) {
    return { ...rest, opaqueProps: filtered };
  }
  return rest;
}
