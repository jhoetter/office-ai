import { CommandError, type CommandHandler } from "@officeai/core";
import type { NodeId } from "@officeai/core";
import type { DocxSnapshot, Paragraph, ParagraphProperties } from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import { locateParagraph } from "./set-paragraph-list.js";

/**
 * Payload for `docx:set-paragraph-spacing`. Mirrors the OOXML
 * `<w:spacing>` element. All fields are optional; the handler merges the
 * non-undefined ones into the paragraph's existing spacing record.
 *
 * Pass `null` to clear a previously-set field. Pass `undefined` to leave
 * it as-is.
 */
export interface SetParagraphSpacingPayload {
  paragraphId: NodeId;
  /** Twips before the paragraph (`<w:spacing w:before>`). */
  before?: number | null;
  /** Twips after the paragraph (`<w:spacing w:after>`). */
  after?: number | null;
  /** Line spacing value (`<w:spacing w:line>`). Interpreted per `lineRule`. */
  line?: number | null;
  /** `auto` (twentieths of a line), `exact` (twips), `atLeast` (twips). */
  lineRule?: "auto" | "exact" | "atLeast" | null;
}

export const setParagraphSpacingHandler: CommandHandler<SetParagraphSpacingPayload, DocxSnapshot> = {
  type: "docx:set-paragraph-spacing",
  apply(snapshot, payload) {
    const { paragraphId } = payload;
    if (!paragraphId) {
      throw new CommandError("invalid-payload", "paragraphId is required");
    }
    validateNumericFields(payload);

    const located = locateParagraph(snapshot.root, paragraphId);
    if (!located) {
      throw new CommandError("unknown-target", `no paragraph with id "${paragraphId}"`);
    }

    const updatedProps = applySpacing(located.paragraph.properties, payload);
    if (spacingEqual(located.paragraph.properties.spacing, updatedProps.spacing)) {
      return { next: snapshot, diff: buildDiff(snapshot.revision, snapshot.revision) };
    }
    const updatedParagraph: Paragraph = { ...located.paragraph, properties: updatedProps };
    const nextDoc = located.replace(updatedParagraph);
    const next = evolveSnapshot(snapshot, nextDoc, { body: true });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: located.paragraph.id,
        path: located.path,
        field: "spacing",
        summary: summariseSpacing(payload),
      }),
    };
  },
};

function validateNumericFields(p: SetParagraphSpacingPayload): void {
  for (const [name, val] of Object.entries({ before: p.before, after: p.after, line: p.line })) {
    if (val === undefined || val === null) continue;
    if (!Number.isFinite(val) || !Number.isInteger(val) || val < 0) {
      throw new CommandError(
        "invalid-payload",
        `${name} must be a non-negative integer in twips (got ${String(val)})`
      );
    }
  }
}

function applySpacing(props: ParagraphProperties, payload: SetParagraphSpacingPayload): ParagraphProperties {
  const opaqueProps = props.opaqueProps?.filter((o) => o.tag !== "w:spacing");
  const baseSpacing = props.spacing ?? {};
  const merged: NonNullable<ParagraphProperties["spacing"]> = { ...baseSpacing };

  if (payload.before !== undefined) {
    if (payload.before === null) delete (merged as { before?: number }).before;
    else (merged as { before: number }).before = payload.before;
  }
  if (payload.after !== undefined) {
    if (payload.after === null) delete (merged as { after?: number }).after;
    else (merged as { after: number }).after = payload.after;
  }
  if (payload.line !== undefined) {
    if (payload.line === null) delete (merged as { line?: number }).line;
    else (merged as { line: number }).line = payload.line;
  }
  if (payload.lineRule !== undefined) {
    if (payload.lineRule === null) delete (merged as { lineRule?: string }).lineRule;
    else (merged as { lineRule: typeof payload.lineRule }).lineRule = payload.lineRule;
  }

  const hasAnySpacing = Object.keys(merged).length > 0;
  const next: ParagraphProperties = hasAnySpacing
    ? { ...props, spacing: merged }
    : { ...props, spacing: undefined };
  const withOpaque: ParagraphProperties =
    opaqueProps && opaqueProps.length > 0 ? { ...next, opaqueProps } : { ...next, opaqueProps: undefined };
  return stripUndefined(withOpaque);
}

function stripUndefined(props: ParagraphProperties): ParagraphProperties {
  const out: Record<string, unknown> = { ...props };
  if (out.spacing === undefined) delete out.spacing;
  if (out.opaqueProps === undefined) delete out.opaqueProps;
  return out as ParagraphProperties;
}

function spacingEqual(
  a: ParagraphProperties["spacing"],
  b: ParagraphProperties["spacing"]
): boolean {
  if (a === b) return true;
  if (!a || !b) return !a && !b;
  return (
    a.before === b.before && a.after === b.after && a.line === b.line && a.lineRule === b.lineRule
  );
}

function summariseSpacing(p: SetParagraphSpacingPayload): string {
  const parts: string[] = [];
  if (p.before !== undefined) parts.push(`before=${p.before === null ? "—" : p.before}`);
  if (p.after !== undefined) parts.push(`after=${p.after === null ? "—" : p.after}`);
  if (p.line !== undefined) parts.push(`line=${p.line === null ? "—" : p.line}`);
  if (p.lineRule !== undefined) parts.push(`lineRule=${p.lineRule === null ? "—" : p.lineRule}`);
  return parts.join(", ") || "no-op";
}
