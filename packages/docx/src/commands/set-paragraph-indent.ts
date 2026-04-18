import { CommandError, type CommandHandler } from "@officeai/core";
import type { DocxSnapshot, Paragraph, ParagraphProperties } from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import { locateParagraph } from "./set-paragraph-list.js";
import type { SetParagraphIndentPayload } from "./payloads.js";

const MAX_TWIPS = 31680; // OOXML caps `<w:ind>` values at ~22 inches.

/**
 * Increment / decrement a paragraph's `indentation.left` by `deltaTwips`
 * and clamp the result into the OOXML legal range. Mirrors the standard
 * Word toolbar behaviour where the indent / outdent buttons step ±360
 * twips (¼ inch).
 *
 * Why a delta and not an absolute value: the toolbar buttons are
 * stateless, the caret may sit on a paragraph that already has a
 * non-zero indent (set in another tool, in the styles part, or by a
 * previous outdent click), and "step by N twips from where you are" is
 * the only sane definition that composes.
 */
export const setParagraphIndentHandler: CommandHandler<SetParagraphIndentPayload, DocxSnapshot> = {
  type: "docx:set-paragraph-indent",
  apply(snapshot, payload) {
    const { paragraphId, deltaTwips } = payload;
    if (!paragraphId) {
      throw new CommandError("invalid-payload", "paragraphId is required");
    }
    if (!Number.isFinite(deltaTwips) || !Number.isInteger(deltaTwips)) {
      throw new CommandError(
        "invalid-payload",
        `deltaTwips must be a finite integer (got ${String(deltaTwips)})`
      );
    }

    const located = locateParagraph(snapshot.root, paragraphId);
    if (!located) {
      throw new CommandError("unknown-target", `no paragraph with id "${paragraphId}"`);
    }

    const previousLeft = located.paragraph.properties.indentation?.left ?? 0;
    const nextLeft = clamp(previousLeft + deltaTwips, 0, MAX_TWIPS);
    if (nextLeft === previousLeft) {
      return {
        next: snapshot,
        diff: buildDiff(snapshot.revision, snapshot.revision, {
          kind: "node-updated",
          nodeId: located.paragraph.id,
          path: located.path,
          field: "indentation.left",
          summary: `indent unchanged (${previousLeft} twips)`,
        }),
      };
    }

    const updatedProps = applyIndentLeft(located.paragraph.properties, nextLeft);
    const updatedParagraph: Paragraph = { ...located.paragraph, properties: updatedProps };
    const nextDoc = located.replace(updatedParagraph);
    const next = evolveSnapshot(snapshot, nextDoc, { body: true });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: located.paragraph.id,
        path: located.path,
        field: "indentation.left",
        summary: `${previousLeft} → ${nextLeft} twips`,
      }),
    };
  },
};

function clamp(value: number, lo: number, hi: number): number {
  if (value < lo) return lo;
  if (value > hi) return hi;
  return value;
}

function applyIndentLeft(props: ParagraphProperties, leftTwips: number): ParagraphProperties {
  // Drop any stale `<w:ind>` opaque carrier so the typed field is the
  // single source of truth on re-serialize.
  const opaqueProps = props.opaqueProps?.filter((o) => o.tag !== "w:ind");
  const baseIndent = props.indentation ?? {};
  const nextIndentation =
    leftTwips === 0 && baseIndent.right === undefined && baseIndent.firstLine === undefined && baseIndent.hanging === undefined
      ? undefined
      : { ...baseIndent, left: leftTwips };
  const next: ParagraphProperties =
    nextIndentation === undefined
      ? { ...props, indentation: undefined }
      : { ...props, indentation: nextIndentation };
  const withOpaque: ParagraphProperties =
    opaqueProps && opaqueProps.length > 0
      ? { ...next, opaqueProps }
      : { ...next, opaqueProps: undefined };
  return stripUndefined(withOpaque);
}

function stripUndefined(props: ParagraphProperties): ParagraphProperties {
  const out: Record<string, unknown> = { ...props };
  if (out.indentation === undefined) delete out.indentation;
  if (out.opaqueProps === undefined) delete out.opaqueProps;
  return out as ParagraphProperties;
}
