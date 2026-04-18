import type {
  DocxSnapshot,
  Paragraph,
  ParagraphProperties,
  Run,
  RunProperties,
  StyleDefinition,
  StylesPart,
} from "../model/types.js";

/**
 * Style cascade resolver (P3.1 / W2). See spec/docx/style-cascade.md.
 *
 * Walks `docDefaults.rPrDefault` → the paragraph's `styleId` chain
 * (basedOn) → the paragraph's own `pPr.rPr` → the run's `rPr`. Returns
 * the absolute, fully-merged `RunProperties` / `ParagraphProperties`
 * that the toolbar should display.
 *
 * Pure read-only function. Cycle-safe (depth-capped + visited set).
 */

const MAX_BASED_ON_DEPTH = 16;

export function resolveEffectiveRpr(
  snapshot: DocxSnapshot,
  paragraphIndex: number,
  runIndex?: number
): RunProperties {
  const paragraph = paragraphAt(snapshot, paragraphIndex);
  const styles = snapshot.root.styles;

  let acc: RunProperties = {};
  if (styles?.docDefaults.rPrDefault) acc = mergeRpr(acc, styles.docDefaults.rPrDefault);

  if (paragraph) {
    const styleId = paragraph.properties.styleId;
    if (styleId && styles) {
      acc = mergeRpr(acc, collectStyleRpr(styles, styleId));
    }
    if (runIndex !== undefined) {
      const run = runAt(paragraph, runIndex);
      if (run) acc = mergeRpr(acc, run.properties);
    }
  }

  return acc;
}

export function resolveEffectivePpr(
  snapshot: DocxSnapshot,
  paragraphIndex: number
): ParagraphProperties {
  const paragraph = paragraphAt(snapshot, paragraphIndex);
  const styles = snapshot.root.styles;

  let acc: ParagraphProperties = {};
  if (styles?.docDefaults.pPrDefault) acc = mergePpr(acc, styles.docDefaults.pPrDefault);

  if (paragraph) {
    const styleId = paragraph.properties.styleId;
    if (styleId && styles) {
      acc = mergePpr(acc, collectStylePpr(styles, styleId));
    }
    acc = mergePpr(acc, paragraph.properties);
  }

  return acc;
}

function paragraphAt(snapshot: DocxSnapshot, paragraphIndex: number): Paragraph | null {
  const block = snapshot.root.body[paragraphIndex];
  if (!block || block.kind !== "paragraph") return null;
  return block;
}

function runAt(paragraph: Paragraph, runIndex: number): Run | null {
  let i = 0;
  for (const child of paragraph.children) {
    if (child.kind === "run") {
      if (i === runIndex) return child;
      i++;
    }
  }
  return null;
}

/**
 * Collapse the basedOn chain root → leaf into a single `RunProperties`.
 * Root applies first, leaf last (so leaf wins on conflicts), matching
 * Word's cascade.
 */
function collectStyleRpr(styles: StylesPart, styleId: string): RunProperties {
  const chain = walkBasedOnChain(styles, styleId);
  let acc: RunProperties = {};
  for (const def of chain) {
    if (def.rPr) acc = mergeRpr(acc, def.rPr);
  }
  return acc;
}

function collectStylePpr(styles: StylesPart, styleId: string): ParagraphProperties {
  const chain = walkBasedOnChain(styles, styleId);
  let acc: ParagraphProperties = {};
  for (const def of chain) {
    if (def.pPr) acc = mergePpr(acc, def.pPr);
  }
  return acc;
}

/**
 * Returns the basedOn chain in root-to-leaf order. A linked style chain
 * `Heading1 → basedOn=Normal → basedOn=undefined` returns
 * `[Normal, Heading1]`. Cycles and missing parents terminate the walk.
 */
function walkBasedOnChain(styles: StylesPart, leafId: string): ReadonlyArray<StyleDefinition> {
  const visited = new Set<string>();
  const reverse: StyleDefinition[] = [];
  let current: string | undefined = leafId;
  let depth = 0;
  while (current && !visited.has(current) && depth < MAX_BASED_ON_DEPTH) {
    visited.add(current);
    const def: StyleDefinition | undefined = styles.styles.get(current);
    if (!def) break;
    reverse.push(def);
    current = def.basedOn;
    depth++;
  }
  reverse.reverse();
  return reverse;
}

/**
 * Merge two `RunProperties`. Right wins on conflicts. `opaqueProps` are
 * concatenated rather than overridden — the cascade resolver doesn't
 * try to dedupe by tag here because the toolbar doesn't read them; a
 * future "render from typed model" path will need a smarter merge.
 */
function mergeRpr(a: RunProperties, b: RunProperties): RunProperties {
  const out: { -readonly [K in keyof RunProperties]: RunProperties[K] } = { ...a };
  if (b.bold !== undefined) out.bold = b.bold;
  if (b.italic !== undefined) out.italic = b.italic;
  if (b.underline !== undefined) out.underline = b.underline;
  if (b.strike !== undefined) out.strike = b.strike;
  if (b.fontFamily !== undefined) out.fontFamily = b.fontFamily;
  if (b.fontSize !== undefined) out.fontSize = b.fontSize;
  if (b.color !== undefined) out.color = b.color;
  if (b.highlight !== undefined) out.highlight = b.highlight;
  return out;
}

function mergePpr(a: ParagraphProperties, b: ParagraphProperties): ParagraphProperties {
  const out: { -readonly [K in keyof ParagraphProperties]: ParagraphProperties[K] } = { ...a };
  if (b.styleId !== undefined) out.styleId = b.styleId;
  if (b.alignment !== undefined) out.alignment = b.alignment;
  if (b.indentation !== undefined) out.indentation = { ...(a.indentation ?? {}), ...b.indentation };
  if (b.spacing !== undefined) out.spacing = { ...(a.spacing ?? {}), ...b.spacing };
  if (b.numbering !== undefined) out.numbering = b.numbering;
  return out;
}
