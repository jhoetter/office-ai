import { CommandError, type CommandHandler } from "@officeai/core";
import type {
  AbstractNum,
  DocxDocument,
  DocxSnapshot,
  NumInstance,
  NumberingDefinitions,
  NumberingLevel,
  Paragraph,
  ParagraphProperties,
} from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import { locateParagraph } from "./set-paragraph-list.js";
import type { ApplyListFormatPayload } from "./payloads.js";

/**
 * B7 — toolbar-friendly list applier.
 *
 * Takes a user-level intent (`bullet` / `decimal`), resolves it to a
 * concrete `<w:num>` instance (creating one if needed), points the
 * paragraph at it, and dirties:
 *
 *   - `body` — the paragraph's `<w:numPr>` pointer changed.
 *   - `numbering` — only when we actually mutated the typed
 *     `NumberingDefinitions` (auto-mint path); reusing an existing
 *     instance keeps the part byte-identical.
 *
 * On the auto-mint path, the serializer's `ensureNumberingPart`
 * helper picks up the dirty flag and registers the part with the
 * package on save. The result is a "bullet button just works"
 * experience even when the original `.docx` has no numbering part
 * at all.
 *
 * Numbering shape we synthesise (matches the minimal set Word ships
 * for new lists):
 *
 *   - bullet  → `<w:multiLevelType w:val="hybridMultilevel">` with
 *               levels 0..2 alternating •, ○, ▪
 *   - decimal → `<w:multiLevelType w:val="multilevel">` with
 *               levels 0..2 as `1.`, `a.`, `i.`
 */
export const applyListFormatHandler: CommandHandler<ApplyListFormatPayload, DocxSnapshot> = {
  type: "docx:apply-list-format",
  apply(snapshot, payload) {
    const { paragraphId, format } = payload;
    if (!paragraphId) {
      throw new CommandError("invalid-payload", "paragraphId is required");
    }
    if (format !== "bullet" && format !== "decimal") {
      throw new CommandError("invalid-payload", `unsupported list format "${format}"`);
    }
    const ilvl = payload.ilvl ?? 0;
    if (!Number.isInteger(ilvl) || ilvl < 0) {
      throw new CommandError("invalid-payload", `ilvl must be a non-negative integer (got ${ilvl})`);
    }

    const located = locateParagraph(snapshot.root, paragraphId);
    if (!located) {
      throw new CommandError("unknown-target", `no paragraph with id "${paragraphId}"`);
    }

    const { numbering: nextDefs, numId, minted } = ensureNumberingForFormat(snapshot.root.numbering, format);

    const updatedProps = applyNumberingToProps(located.paragraph.properties, numId, ilvl);
    const updatedParagraph: Paragraph = { ...located.paragraph, properties: updatedProps };
    let nextDoc: DocxDocument = located.replace(updatedParagraph);
    if (nextDefs !== snapshot.root.numbering) {
      nextDoc = { ...nextDoc, numbering: nextDefs };
    }

    const next = evolveSnapshot(snapshot, nextDoc, {
      body: true,
      ...(minted ? { numbering: true } : {}),
    });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: located.paragraph.id,
        path: located.path,
        field: "numbering",
        summary: minted
          ? `apply ${format} list (auto-mint numId=${numId})`
          : `apply ${format} list (numId=${numId})`,
      }),
    };
  },
};

interface EnsureResult {
  readonly numbering: NumberingDefinitions | undefined;
  readonly numId: number;
  readonly minted: boolean;
}

/**
 * Find an existing `<w:num>` whose abstract definition matches the
 * requested format at level 0. If no compatible instance exists,
 * mint a fresh abstract + num pair using ids that don't clash with
 * anything already in the part.
 */
function ensureNumberingForFormat(
  defs: NumberingDefinitions | undefined,
  format: "bullet" | "decimal"
): EnsureResult {
  if (defs) {
    for (const num of defs.nums.values()) {
      const abs = defs.abstractNums.get(num.abstractNumId);
      if (!abs) continue;
      const lvl0 = abs.levels.find((l) => l.ilvl === 0);
      if (!lvl0) continue;
      if (lvl0.numFmt && matchesFormat(lvl0.numFmt, format)) {
        return { numbering: defs, numId: num.numId, minted: false };
      }
    }
  }

  const nextAbstractId = mintAbstractId(defs);
  const nextNumId = mintNumId(defs);
  const abstractNum = buildAbstractNum(nextAbstractId, format);
  const numInstance: NumInstance = { numId: nextNumId, abstractNumId: nextAbstractId };

  const abstractNums = new Map(defs?.abstractNums ?? []);
  abstractNums.set(nextAbstractId, abstractNum);
  const nums = new Map(defs?.nums ?? []);
  nums.set(nextNumId, numInstance);

  return {
    numbering: { abstractNums, nums },
    numId: nextNumId,
    minted: true,
  };
}

function matchesFormat(numFmt: string, format: "bullet" | "decimal"): boolean {
  const v = numFmt.toLowerCase();
  if (format === "bullet") return v === "bullet";
  return (
    v === "decimal" ||
    v === "decimalzero" ||
    v === "lowerroman" ||
    v === "upperroman" ||
    v === "lowerletter" ||
    v === "upperletter"
  );
}

function mintAbstractId(defs: NumberingDefinitions | undefined): string {
  let max = -1;
  if (defs) {
    for (const id of defs.abstractNums.keys()) {
      const n = Number.parseInt(id, 10);
      if (Number.isFinite(n) && n > max) max = n;
    }
  }
  return String(max + 1);
}

function mintNumId(defs: NumberingDefinitions | undefined): number {
  let max = 0;
  if (defs) {
    for (const id of defs.nums.keys()) {
      if (id > max) max = id;
    }
  }
  return max + 1;
}

function buildAbstractNum(id: string, format: "bullet" | "decimal"): AbstractNum {
  if (format === "bullet") {
    const levels: NumberingLevel[] = [
      { ilvl: 0, numFmt: "bullet", lvlText: "\u2022", start: 1 },
      { ilvl: 1, numFmt: "bullet", lvlText: "\u25E6", start: 1 },
      { ilvl: 2, numFmt: "bullet", lvlText: "\u25AA", start: 1 },
    ];
    return { id, multiLevelType: "hybridMultilevel", levels };
  }
  const levels: NumberingLevel[] = [
    { ilvl: 0, numFmt: "decimal", lvlText: "%1.", start: 1 },
    { ilvl: 1, numFmt: "lowerLetter", lvlText: "%2.", start: 1 },
    { ilvl: 2, numFmt: "lowerRoman", lvlText: "%3.", start: 1 },
  ];
  return { id, multiLevelType: "multilevel", levels };
}

function applyNumberingToProps(props: ParagraphProperties, numId: number, ilvl: number): ParagraphProperties {
  const opaqueProps = props.opaqueProps?.filter((o) => o.tag !== "w:numPr");
  const next: ParagraphProperties = {
    ...props,
    numbering: { numId, ilvl },
    ...(opaqueProps && opaqueProps.length > 0 ? { opaqueProps } : { opaqueProps: undefined }),
  };
  if (next.opaqueProps === undefined) {
    const { opaqueProps: _drop, ...rest } = next;
    void _drop;
    return rest;
  }
  return next;
}
