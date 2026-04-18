import { Plugin, PluginKey, TextSelection } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";
import type { EditorState, Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import type {
  DocxAgent,
  DocxSnapshot,
  HeaderFooterPart,
  PageChunk,
} from "@officeai/docx";
import { chunkIntoPages, resolveHeaderFooterParts } from "@officeai/docx";

/**
 * Page-decoration plugin (P3.3 / W11, extended in P3.8).
 *
 * Renders Word-flavoured "sheets of paper" inside the single
 * ProseMirror body editor by mixing three decoration kinds:
 *
 * 1. **Per-block node decorations** add `pm-page-block` (always),
 *    `pm-page-first` (on the first block of a chunk) and
 *    `pm-page-last` (on the last block) classes so CSS can paint
 *    each chunk as a discrete white card on a grey backdrop.
 * 2. **Page-edge widgets** are inserted at chunk boundaries and
 *    carry a *footer zone* (closing the previous page) + a
 *    visible grey gap + a *header zone* (opening the next page).
 *    They double as the visual page-break Word users expect.
 * 3. **Page-cap widgets** at the very start and very end of the
 *    document carry the first page's header zone and the last
 *    page's footer zone respectively, so every page has visible
 *    chrome on top and bottom.
 *
 * Header/footer zones surface the resolved
 * {@link HeaderFooterPart} text for the section that owns each
 * page. Double-clicking a zone fires a CustomEvent
 * (`pm-page-zone-edit`) on the editor host so the React shell can
 * pop a popover bound to `docx:set-header-text` / `set-footer-text`.
 *
 * Pure read-only inside PM: never dispatches transactions, never
 * mutates the snapshot, and the OOXML round-trip byte-equality is
 * unaffected.
 */
export const pageDecorationsKey = new PluginKey<PageDecorationsState>("docx-page-decorations");

export interface PageDecorationsState {
  readonly chunks: ReadonlyArray<PageChunk>;
  readonly decorations: DecorationSet;
}

/** CustomEvent payload dispatched on double-click of a page header / footer zone. */
export interface PageZoneEditDetail {
  readonly slot: "header" | "footer";
  readonly partPath: string | null;
  readonly target: "default" | "first" | "even" | null;
  readonly pageNumber: number;
  readonly text: string;
}

export const PAGE_ZONE_EDIT_EVENT = "pm-page-zone-edit";

export function pageDecorationsPlugin(agent: DocxAgent): Plugin<PageDecorationsState> {
  return new Plugin<PageDecorationsState>({
    key: pageDecorationsKey,
    state: {
      init(_config, state) {
        return computeState(agent, state);
      },
      apply(tr: Transaction, prev: PageDecorationsState, _old: EditorState, next: EditorState) {
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
  const decorations = buildDecorations(snapshot, chunks, state);
  return { chunks, decorations };
}

function buildDecorations(
  snapshot: DocxSnapshot,
  chunks: ReadonlyArray<PageChunk>,
  state: EditorState
): DecorationSet {
  if (chunks.length === 0) return DecorationSet.empty;

  // PM document top-level children mirror DocxSnapshot.body 1:1.
  const docNode = state.doc;
  const childPositions: number[] = [];
  const childSizes: number[] = [];
  let pos = 0;
  docNode.forEach((child) => {
    childPositions.push(pos);
    childSizes.push(child.nodeSize);
    pos += child.nodeSize;
  });
  const docEnd = pos;

  const decos: Decoration[] = [];

  for (let i = 0; i < chunks.length; i++) {
    const chunk = chunks[i];
    const resolved = resolveHeaderFooterParts(snapshot, chunk.sectionIndex);
    const headerPart = pickHeaderFooter(resolved.headers, chunk.pageWithinSection);
    const footerPart = pickHeaderFooter(resolved.footers, chunk.pageWithinSection);

    // Per-block node decorations: every block in the chunk gets the
    // base class; the chunk's first / last blocks pick up the
    // edge classes so CSS can round their corners and add the
    // sheet-of-paper drop shadow.
    for (let b = chunk.startBlock; b < chunk.endBlock; b++) {
      const start = childPositions[b];
      if (start === undefined) continue;
      const end = start + childSizes[b];
      const classes = ["pm-page-block"];
      if (b === chunk.startBlock) classes.push("pm-page-first");
      if (b === chunk.endBlock - 1) classes.push("pm-page-last");
      decos.push(
        Decoration.node(start, end, {
          class: classes.join(" "),
          "data-page-number": String(chunk.pageNumber),
        })
      );
    }

    // Page cap (header zone) for page 1: widget at the very start
    // of the document so the first sheet has visible header chrome.
    if (i === 0) {
      decos.push(
        Decoration.widget(0, () => renderHeaderCap(chunk, headerPart), {
          side: -1,
          key: `page-cap-top-${chunk.pageNumber}`,
        })
      );
    }

    // Page-edge widget between this chunk and the next chunk. Owns
    // the footer of *this* page + the gap + the header of the next
    // page so the visual page break tells a complete story.
    if (i < chunks.length - 1) {
      const nextChunk = chunks[i + 1];
      const nextResolved = resolveHeaderFooterParts(snapshot, nextChunk.sectionIndex);
      const nextHeaderPart = pickHeaderFooter(nextResolved.headers, nextChunk.pageWithinSection);
      const insertAt = childPositions[nextChunk.startBlock];
      if (insertAt === undefined) continue;
      decos.push(
        Decoration.widget(
          insertAt,
          () => renderPageEdge(chunk, nextChunk, footerPart, nextHeaderPart),
          {
            side: -1,
            key: `page-edge-${chunk.pageNumber}->${nextChunk.pageNumber}`,
          }
        )
      );
    }

    // Page cap (footer zone) for the last page: widget at the very
    // end of the document.
    if (i === chunks.length - 1) {
      decos.push(
        Decoration.widget(docEnd, () => renderFooterCap(chunk, footerPart), {
          side: 1,
          key: `page-cap-bottom-${chunk.pageNumber}`,
        })
      );
    }
  }

  return DecorationSet.create(docNode, decos);
}

function pickHeaderFooter(
  resolved: { default?: HeaderFooterPart; first?: HeaderFooterPart; even?: HeaderFooterPart },
  pageWithinSection: number
): HeaderFooterPart | undefined {
  // Word's lookup: page 1 of a section uses `first` when present
  // (gated by `<w:titlePg/>` upstream — `resolveHeaderFooterParts`
  // only populates `first` when the section break carries it), then
  // falls back to `default`. Even-page parts are not yet honoured
  // (deferred to P4 alongside `<w:evenAndOddHeaders/>`).
  if (pageWithinSection === 1 && resolved.first) return resolved.first;
  return resolved.default ?? resolved.first ?? resolved.even;
}

function renderHeaderCap(chunk: PageChunk, headerPart: HeaderFooterPart | undefined): HTMLElement {
  const cap = document.createElement("div");
  cap.className = "pm-page-cap pm-page-cap-top";
  cap.setAttribute("contenteditable", "false");
  cap.appendChild(renderHeaderZone(chunk, headerPart));
  return cap;
}

function renderFooterCap(chunk: PageChunk, footerPart: HeaderFooterPart | undefined): HTMLElement {
  const cap = document.createElement("div");
  cap.className = "pm-page-cap pm-page-cap-bottom";
  cap.setAttribute("contenteditable", "false");
  cap.appendChild(renderFooterZone(chunk, footerPart));
  return cap;
}

function renderPageEdge(
  prev: PageChunk,
  next: PageChunk,
  prevFooter: HeaderFooterPart | undefined,
  nextHeader: HeaderFooterPart | undefined
): HTMLElement {
  const edge = document.createElement("div");
  edge.className = "pm-page-edge";
  edge.setAttribute("contenteditable", "false");
  edge.appendChild(renderFooterZone(prev, prevFooter));

  const gap = document.createElement("div");
  gap.className = "pm-page-gap";
  const label = document.createElement("span");
  label.className = "pm-page-gap-label";
  label.textContent = `Page ${next.pageNumber}`;
  gap.appendChild(label);
  edge.appendChild(gap);

  edge.appendChild(renderHeaderZone(next, nextHeader));
  return edge;
}

function renderHeaderZone(chunk: PageChunk, part: HeaderFooterPart | undefined): HTMLElement {
  return renderZone("header", chunk, part);
}

function renderFooterZone(chunk: PageChunk, part: HeaderFooterPart | undefined): HTMLElement {
  return renderZone("footer", chunk, part);
}

function renderZone(
  slot: "header" | "footer",
  chunk: PageChunk,
  part: HeaderFooterPart | undefined
): HTMLElement {
  const zone = document.createElement("div");
  zone.className = `pm-page-zone pm-page-zone-${slot}`;
  zone.setAttribute("contenteditable", "false");
  zone.setAttribute("data-page-zone", slot);
  zone.setAttribute("data-page-number", String(chunk.pageNumber));
  zone.setAttribute("data-zone-label", slot);
  if (part) {
    zone.setAttribute("data-part-path", part.partPath);
    zone.setAttribute("data-part-target", part.target);
  }
  zone.title = `Double-click to edit ${slot}`;

  const text = part ? extractZoneText(part) : "";
  const inner = document.createElement("div");
  inner.className = "pm-page-zone-content";
  if (text.length === 0) {
    inner.classList.add("pm-page-zone-empty");
    inner.textContent = part
      ? `Click to add ${slot} text`
      : `Double-click to add a ${slot}`;
  } else {
    inner.textContent = text;
  }
  zone.appendChild(inner);

  zone.addEventListener("dblclick", (ev) => {
    ev.preventDefault();
    ev.stopPropagation();
    const detail: PageZoneEditDetail = {
      slot,
      partPath: part?.partPath ?? null,
      target: part?.target ?? null,
      pageNumber: chunk.pageNumber,
      text,
    };
    zone.dispatchEvent(
      new CustomEvent<PageZoneEditDetail>(PAGE_ZONE_EDIT_EVENT, {
        detail,
        bubbles: true,
      })
    );
  });

  return zone;
}

function extractZoneText(part: HeaderFooterPart): string {
  // Surface the first non-empty paragraph's text. Page-number
  // fields render as their canonical name (`PAGE`, `NUMPAGES`)
  // wrapped in brackets so the user sees what's there even though
  // the rendered value is dynamic.
  for (const block of part.body) {
    if (block.kind !== "paragraph") continue;
    let acc = "";
    for (const inline of block.children) {
      if (inline.kind !== "run") continue;
      for (const c of inline.children) {
        if (c.kind === "text") acc += c.text;
        else if (c.kind === "page-number-field") acc += `[${c.field}]`;
      }
    }
    if (acc.trim().length > 0) return acc;
  }
  return "";
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

/**
 * P3.5 / W19 — move the caret to the start of the requested page and
 * scroll it into view. Returns `true` when the jump succeeded.
 */
export function gotoPage(
  view: EditorView,
  pageNumber: number,
  chunks: ReadonlyArray<PageChunk>
): boolean {
  if (chunks.length === 0) return false;
  const clampedPage = Math.max(1, Math.min(pageNumber, chunks.length));
  const chunk = chunks.find((c) => c.pageNumber === clampedPage);
  if (!chunk) return false;
  const docNode = view.state.doc;
  let pos = 0;
  let i = 0;
  let target: number | null = null;
  docNode.forEach((child) => {
    if (i === chunk.startBlock) target = pos;
    pos += child.nodeSize;
    i++;
  });
  if (target === null) return false;
  const $pos = view.state.doc.resolve(target + 1);
  const tr = view.state.tr.setSelection(TextSelection.between($pos, $pos)).scrollIntoView();
  view.dispatch(tr);
  view.focus();
  return true;
}

/**
 * P3.5 / W21 — move the caret to the start of the next (or previous)
 * page chunk relative to the caret's current position. Returns
 * `true` when the move succeeded; `false` when the caret was already
 * on the first/last page (caller falls through to PM's default
 * PageUp / PageDown handling).
 */
export function movePageRelative(view: EditorView, direction: 1 | -1): boolean {
  const chunks = getPageChunks(view.state);
  if (chunks.length === 0) return false;
  const current = pageNumberForPos(chunks, view.state, view.state.selection.from);
  const next = current + direction;
  if (next < 1 || next > chunks.length) return false;
  return gotoPage(view, next, chunks);
}
