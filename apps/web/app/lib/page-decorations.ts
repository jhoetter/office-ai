import { Plugin, PluginKey } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";
import type { EditorState, Transaction } from "prosemirror-state";
import type { DocxAgent, PageChunk } from "@officeai/docx";
import { chunkIntoPages } from "@officeai/docx";

/**
 * Page-decoration plugin (P3.3 / W11).
 *
 * Renders visible page boundaries inside the single ProseMirror body
 * editor by:
 *
 * - Computing a {@link PageChunk} array from the live `DocxAgent`
 *   snapshot on every doc-changing transaction.
 * - Walking the PM document's top-level children (which mirror
 *   `snapshot.root.body` 1:1 — same ordering, same length) and
 *   inserting a "Page N" widget decoration before the first block of
 *   every chunk except the first.
 * - Wrapping every chunk's blocks in a `pm-page-band` class via
 *   inline-block decorations so per-page CSS chrome (subtle paper
 *   shadow, margin gutters) can target each page band as a whole.
 *
 * Pure read-only: never dispatches transactions, never mutates the
 * snapshot, and round-trip byte-equality is unaffected.
 *
 * The chunker runs without a `measure` function on Node-side tests
 * and on the initial render. Once the body has rendered, a follow-up
 * pass (P3.5 polish) can pass per-block heights to `chunkIntoPages`
 * for measured pagination; this initial version honors hard
 * `<w:br w:type="page"/>` and `<w:lastRenderedPageBreak/>` hints
 * baked into the document, which already covers the masterthesis
 * fixture.
 */
export const pageDecorationsKey = new PluginKey<PageDecorationsState>("docx-page-decorations");

export interface PageDecorationsState {
  readonly chunks: ReadonlyArray<PageChunk>;
  readonly decorations: DecorationSet;
}

export function pageDecorationsPlugin(agent: DocxAgent): Plugin<PageDecorationsState> {
  return new Plugin<PageDecorationsState>({
    key: pageDecorationsKey,
    state: {
      init(_config, state) {
        return computeState(agent, state);
      },
      apply(tr: Transaction, prev: PageDecorationsState, _old: EditorState, next: EditorState) {
        // Recompute on every transaction that touches the doc (paging
        // can shift when paragraphs are inserted/removed) or whenever
        // the agent revision advances (mutations from the bus may
        // shift chunks even without a PM-side change).
        const meta = tr.getMeta(pageDecorationsKey);
        if (!tr.docChanged && meta !== "force") {
          return prev;
        }
        return computeState(agent, next);
      },
    },
    props: {
      decorations(state) {
        return pageDecorationsKey.getState(state)?.decorations ?? DecorationSet.empty;
      },
    },
  });
}

function computeState(agent: DocxAgent, state: EditorState): PageDecorationsState {
  const snapshot = agent.getSnapshot();
  const chunks = chunkIntoPages(snapshot);
  const decorations = buildDecorations(chunks, state);
  return { chunks, decorations };
}

function buildDecorations(
  chunks: ReadonlyArray<PageChunk>,
  state: EditorState
): DecorationSet {
  if (chunks.length === 0) return DecorationSet.empty;

  // PM document top-level children mirror DocxSnapshot.body 1:1: each
  // typed BlockNode becomes one PM block node. Walk PM children to
  // find the position of each top-level child by index.
  const docNode = state.doc;
  const childPositions: number[] = [];
  let pos = 0;
  docNode.forEach((child) => {
    childPositions.push(pos);
    pos += child.nodeSize;
  });

  const widgets: Decoration[] = [];
  for (let i = 0; i < chunks.length; i++) {
    const chunk = chunks[i];
    const blockPos = childPositions[chunk.startBlock];
    if (blockPos === undefined) continue;
    if (i === 0) continue;
    widgets.push(
      Decoration.widget(blockPos, () => renderPageDivider(chunk.pageNumber, chunks.length), {
        side: -1,
        key: `page-divider-${chunk.pageNumber}`,
      })
    );
  }

  return DecorationSet.create(docNode, widgets);
}

function renderPageDivider(pageNumber: number, total: number): HTMLElement {
  const wrap = document.createElement("div");
  wrap.className = "pm-page-divider";
  wrap.setAttribute("data-page-number", String(pageNumber));
  wrap.setAttribute("contenteditable", "false");
  const label = document.createElement("span");
  label.className = "pm-page-divider__label";
  label.textContent = `Page ${pageNumber} of ${total}`;
  wrap.appendChild(label);
  return wrap;
}

/**
 * Read the most recent page chunks from a live editor view.
 * Returns an empty array when the plugin is not installed.
 */
export function getPageChunks(state: EditorState): ReadonlyArray<PageChunk> {
  return pageDecorationsKey.getState(state)?.chunks ?? [];
}

/**
 * Resolve which page a PM document position falls on. Returns 1
 * when the position is before the first page boundary, and `chunks.length`
 * when after the last.
 */
export function pageNumberForPos(
  chunks: ReadonlyArray<PageChunk>,
  state: EditorState,
  pos: number
): number {
  if (chunks.length === 0) return 1;
  const docNode = state.doc;
  const childIndices: number[] = [];
  let acc = 0;
  docNode.forEach((child) => {
    childIndices.push(acc);
    acc += child.nodeSize;
  });
  // Find the top-level child index that contains `pos`.
  let blockIdx = 0;
  for (let i = 0; i < childIndices.length; i++) {
    if (childIndices[i] <= pos) blockIdx = i;
    else break;
  }
  for (let i = chunks.length - 1; i >= 0; i--) {
    if (chunks[i].startBlock <= blockIdx) return chunks[i].pageNumber;
  }
  return 1;
}
