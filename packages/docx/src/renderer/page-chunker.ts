import type {
  BlockNode,
  DocxSnapshot,
  PageMargins,
  PageSize,
  Paragraph,
  RunChild,
  SectionBreak,
  SectionProperties,
} from "../model/types.js";

/**
 * Page chunker (P3.3 / W9).
 *
 * Pure function — no DOM access. Splits the body of a {@link DocxSnapshot}
 * into a sequence of {@link PageChunk} ranges, one per visible Word
 * page. The renderer wraps each chunk's blocks in a `<PageFrame>` to
 * draw the page chrome at the right geometry.
 *
 * Algorithm:
 *   1. Walk `snapshot.root.body` and group blocks by section. A
 *      section ends at every `SectionBreak` block (per OOXML, the
 *      `<w:sectPr>` carries the properties of the section it
 *      terminates). The last section uses the geometry of the final
 *      `SectionBreak` if one exists; otherwise we synthesize a
 *      US-letter default so the renderer can still draw a frame.
 *   2. Within each section, walk the blocks in order. Page boundaries
 *      come from three signals, in priority order:
 *
 *      - **Hard breaks** — a paragraph contains a `PageBreakLeaf`
 *        (`<w:br w:type="page"/>`). The break flushes the current
 *        page; the paragraph itself stays on the new page so its
 *        post-break content is visible.
 *      - **Hint breaks** — a paragraph contains a
 *        `LastRenderedPageBreakLeaf` (`<w:lastRenderedPageBreak/>`).
 *        Word writes these at positions where it broke pagination on
 *        the last save. We honor them only when the caller did not
 *        provide a `measure` function — they're a cheap heuristic for
 *        Node-side tests and the initial render before measurement
 *        completes.
 *      - **Measured breaks** — when `measure(blockIndex)` is
 *        provided, the chunker tracks accumulated content height and
 *        flushes the current page when the next block would overflow
 *        the section's content area. The browser-side renderer wires
 *        this to `getBoundingClientRect()` on the rendered PM blocks.
 *   3. Section breaks themselves do not flush — every `SectionBreak`
 *      block is the LAST block on its page. The next block (in the
 *      next section) starts a fresh page implicitly because section
 *      geometry may differ.
 *
 * Page numbers are 1-based and global across the whole document.
 * Per-section restart (via `<w:pgNumType w:start>`) is a P4 polish
 * item; P3 ships global numbering only.
 */
export interface PageGeometry {
  readonly pgSz: PageSize;
  readonly pgMar: PageMargins;
}

export interface PageChunk {
  /** Index of the first body block on this page (inclusive). */
  readonly startBlock: number;
  /** Index just after the last body block on this page (exclusive). */
  readonly endBlock: number;
  /**
   * Index into `body` of the `SectionBreak` whose properties drive
   * this page's geometry. Pages within the same section share one
   * `sectionIndex`. When the body has no terminating `SectionBreak`
   * the synthetic last section uses `body.length` so callers can
   * detect the synthesized case.
   */
  readonly sectionIndex: number;
  /** 1-based page number across the whole document. */
  readonly pageNumber: number;
  /**
   * 1-based page number within the current section. P3 just resets
   * this to 1 at every section boundary; honoring `<w:pgNumType
   * w:start>` is a P4 polish item.
   */
  readonly pageWithinSection: number;
  /** Effective page size + margins for this page. */
  readonly geometry: PageGeometry;
}

/**
 * Optional per-block height measurer (CSS pixels). The browser-side
 * renderer wires this to the rendered PM blocks; Node-side tests
 * leave it `undefined` and rely on hard / hint breaks only.
 */
export type Measure = (blockIndex: number) => number;

/** US Letter @ 1" margins, 0.5" header/footer. Used when sectPr is absent. */
const DEFAULT_GEOMETRY: PageGeometry = {
  pgSz: { w: 12240, h: 15840 },
  pgMar: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720 },
};

interface SectionGroup {
  /** Index of the SectionBreak that owns this group, or `body.length` for the synthetic tail. */
  readonly sectionIndex: number;
  /** Block indices belonging to this section (non-section-break blocks). */
  readonly blocks: ReadonlyArray<number>;
  readonly geometry: PageGeometry;
}

export function chunkIntoPages(snapshot: DocxSnapshot, measure?: Measure): ReadonlyArray<PageChunk> {
  const body = snapshot.root.body;
  const sections = groupBySection(body);
  const out: PageChunk[] = [];
  let pageNumber = 1;

  for (const section of sections) {
    const contentHeightTwips = computeContentHeight(section.geometry);
    let pageWithinSection = 1;
    let currentBlocks: number[] = [];
    let accumulated = 0;

    const flush = (): void => {
      if (currentBlocks.length === 0) return;
      out.push(buildChunk(currentBlocks, section, pageNumber, pageWithinSection));
      pageNumber += 1;
      pageWithinSection += 1;
      currentBlocks = [];
      accumulated = 0;
    };

    for (const blockIndex of section.blocks) {
      const block = body[blockIndex];
      const breakSignal = classifyBreakSignal(block, measure !== undefined);

      if (breakSignal === "hard-before" && currentBlocks.length > 0) {
        flush();
      }

      if (measure) {
        const h = measure(blockIndex);
        if (currentBlocks.length > 0 && accumulated + h > contentHeightTwips) {
          flush();
        }
        accumulated += h;
      } else if (breakSignal === "hint" && currentBlocks.length > 0) {
        flush();
      }

      currentBlocks.push(blockIndex);
    }

    if (currentBlocks.length > 0) {
      out.push(buildChunk(currentBlocks, section, pageNumber, pageWithinSection));
      pageNumber += 1;
    } else if (out.length === 0 || out[out.length - 1].sectionIndex !== section.sectionIndex) {
      // Empty section (e.g. trailing sectPr with no preceding content).
      // Still emit one frame so the renderer draws the section's
      // geometry; otherwise navigation/page count would skip it.
      out.push(buildChunk([], section, pageNumber, 1));
      pageNumber += 1;
    }
  }

  if (out.length === 0) {
    // Empty document — emit a single blank page so the renderer has
    // something to draw.
    out.push({
      startBlock: 0,
      endBlock: 0,
      sectionIndex: body.length,
      pageNumber: 1,
      pageWithinSection: 1,
      geometry: DEFAULT_GEOMETRY,
    });
  }

  return out;
}

function buildChunk(
  blocks: ReadonlyArray<number>,
  section: SectionGroup,
  pageNumber: number,
  pageWithinSection: number
): PageChunk {
  const startBlock = blocks.length > 0 ? blocks[0] : section.sectionIndex;
  const endBlock = blocks.length > 0 ? blocks[blocks.length - 1] + 1 : section.sectionIndex;
  return {
    startBlock,
    endBlock,
    sectionIndex: section.sectionIndex,
    pageNumber,
    pageWithinSection,
    geometry: section.geometry,
  };
}

function groupBySection(body: ReadonlyArray<BlockNode>): ReadonlyArray<SectionGroup> {
  const groups: SectionGroup[] = [];
  let current: number[] = [];
  for (let i = 0; i < body.length; i++) {
    const block = body[i];
    if (block.kind === "section-break") {
      groups.push({
        sectionIndex: i,
        blocks: current,
        geometry: geometryFromSection(block),
      });
      current = [];
    } else {
      current.push(i);
    }
  }
  if (current.length > 0) {
    groups.push({
      sectionIndex: body.length,
      blocks: current,
      geometry: DEFAULT_GEOMETRY,
    });
  }
  return groups;
}

function geometryFromSection(section: SectionBreak): PageGeometry {
  return geometryFromProperties(section.properties);
}

/**
 * Resolve a {@link SectionProperties} to a complete {@link PageGeometry},
 * filling missing fields with US-Letter defaults so downstream code
 * never has to handle `undefined` page sizes / margins.
 *
 * Exposed as a stable export because the React layer also needs to
 * compute geometry for the *current* page when rendering a
 * `<PageFrame>`.
 */
export function geometryFromProperties(props: SectionProperties): PageGeometry {
  return {
    pgSz: props.pgSz ?? DEFAULT_GEOMETRY.pgSz,
    pgMar: props.pgMar ?? DEFAULT_GEOMETRY.pgMar,
  };
}

/**
 * Resolve the document's *nominal* page geometry — i.e. the geometry
 * the editor card should render at when the document has not yet been
 * chunked.
 *
 * Strategy: Word stores `<w:sectPr>` at the END of the section it
 * describes, so the *first* `SectionBreak` block is the source of truth
 * for the first section's geometry. When the body has no section break
 * at all (a fresh doc, or a doc whose `sectPr` is implicit), fall back
 * to the US-Letter default so the renderer always has a frame.
 *
 * Pure / no DOM access — safe to call in the React render path.
 */
export function documentPageGeometry(snapshot: DocxSnapshot): PageGeometry {
  for (const block of snapshot.root.body) {
    if (block.kind === "section-break") {
      return geometryFromSection(block);
    }
  }
  return DEFAULT_GEOMETRY;
}

type BreakSignal = "none" | "hard-before" | "hint";

function classifyBreakSignal(block: BlockNode, measureProvided: boolean): BreakSignal {
  if (block.kind !== "paragraph") return "none";
  for (const child of paragraphRunChildren(block)) {
    if (child.kind === "page-break") return "hard-before";
    if (!measureProvided && child.kind === "last-rendered-page-break") return "hint";
  }
  return "none";
}

function* paragraphRunChildren(p: Paragraph): IterableIterator<RunChild> {
  for (const inline of p.children) {
    if (inline.kind !== "run") continue;
    for (const child of inline.children) {
      yield child;
    }
  }
}

function computeContentHeight(geometry: PageGeometry): number {
  return Math.max(1, geometry.pgSz.h - geometry.pgMar.top - geometry.pgMar.bottom);
}
