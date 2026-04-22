import { Plugin, PluginKey, TextSelection } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";
import type { EditorState, Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import type {
  BlockNode,
  DocxAgent,
  DocxSnapshot,
  Footnote,
  HeaderFooterPart,
  InlineImageDrawing,
  InlineNode,
  Measure,
  PageChunk,
  PageNumberFieldLeaf,
  Paragraph,
  Run,
  RunChild,
  TextLeaf,
} from "@officeai/docx";
import { chunkIntoPages, resolveHeaderFooterParts } from "@officeai/docx";

/**
 * CSS pixels → OOXML twips. 1 inch = 1440 twips = 96 CSS pixels, so
 * one CSS px is 15 twips. We measure rendered block heights in CSS
 * pixels (`getBoundingClientRect().height`) and feed them to the
 * chunker which speaks twips. Centralised so the conversion factor
 * lives in one place — drop or change it in lockstep with the
 * page-card sizing math in `DocxEditor.tsx` / `globals.css`.
 */
export const TWIPS_PER_CSS_PX = 15;

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
 * page. The zones are rendered with `contenteditable=true` so the
 * user clicks straight into them — Word-style "in-place" header /
 * footer authoring. On blur (or Enter) the zone fires a
 * `pm-page-zone-commit` CustomEvent on the editor host, which the
 * React shell routes to `docx:set-header-text` /
 * `docx:set-footer-text`.
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

/** CustomEvent payload dispatched when a header/footer zone commits an edit. */
export interface PageZoneCommitDetail {
  readonly slot: "header" | "footer";
  readonly partPath: string | null;
  readonly target: "default" | "first" | "even" | null;
  readonly pageNumber: number;
  /**
   * Plain-text fall-back. DocxEditor still wires this into the
   * legacy `docx:set-header-text` / `docx:set-footer-text`
   * commands when {@link blocks} is null (single-paragraph text-only
   * edit, no tokens to preserve).
   */
  readonly text: string;
  /**
   * Rich block-level body produced by the in-place authoring
   * surface. Non-null when the part has multi-paragraph content,
   * page-number-field tokens, or inline images that we don't want
   * the legacy plain-text command to flatten on commit. DocxEditor
   * dispatches `docx:set-header-footer-blocks` with this body when
   * present.
   */
  readonly blocks: ReadonlyArray<BlockNode> | null;
}

export const PAGE_ZONE_COMMIT_EVENT = "pm-page-zone-commit";

/**
 * CustomEvent payload dispatched when the user enters or leaves an
 * editable header/footer zone. Powers the contextual "Header & Footer"
 * toolbar cluster in the DOCX editor (Word's "Close header and footer"
 * affordance). `slot` is `null` when focus moves back to the body.
 */
export interface PageZoneFocusDetail {
  readonly slot: "header" | "footer" | null;
  readonly target: "default" | "first" | "even" | null;
  readonly partPath: string | null;
  readonly pageNumber: number | null;
}

export const PAGE_ZONE_FOCUS_EVENT = "pm-page-zone-focus";

/**
 * CustomEvent payload dispatched when the user double-clicks an empty
 * header / footer zone (one whose section has no `<w:headerReference>` /
 * `<w:footerReference>` of the requested kind). DocxEditor catches this
 * and dispatches `docx:create-header-footer-part` so the next render
 * shows an editable zone the caret can land in (Word's
 * "double-click-the-header" affordance).
 */
export interface PageZoneMintDetail {
  readonly slot: "header" | "footer";
  readonly pageNumber: number;
}

export const PAGE_ZONE_MINT_EVENT = "pm-page-zone-mint";

export function pageDecorationsPlugin(agent: DocxAgent): Plugin<PageDecorationsState> {
  // Per-block measured heights, in twips, indexed by `body` block index.
  // Lives in plugin closure rather than plugin state because it is a
  // *measurement cache* — the source of truth is the DOM, the cache is a
  // memo that lets us decide when to re-chunk without thrashing layout.
  // PM state is for things that participate in transactions; this does
  // not.
  let heightCache: number[] = [];
  // RAF handle so we never schedule more than one pending measurement
  // pass per frame (typing into a long doc can fire many transactions
  // back-to-back).
  let scheduledRaf: number | null = null;

  const measure: Measure = (blockIndex) => heightCache[blockIndex] ?? 0;

  return new Plugin<PageDecorationsState>({
    key: pageDecorationsKey,
    state: {
      init(_config, state) {
        return computeState(agent, state, measure);
      },
      apply(tr: Transaction, prev: PageDecorationsState, _old: EditorState, next: EditorState) {
        const meta = tr.getMeta(pageDecorationsKey);
        if (!tr.docChanged && meta !== "force") {
          return prev;
        }
        return computeState(agent, next, measure);
      },
    },
    props: {
      decorations(state) {
        return pageDecorationsKey.getState(state)?.decorations ?? DecorationSet.empty;
      },
    },
    view(view) {
      const schedule = (): void => {
        if (scheduledRaf !== null) return;
        scheduledRaf = requestAnimationFrame(() => {
          scheduledRaf = null;
          remeasureAndForce(view, heightCache, (next) => {
            heightCache = next;
          });
        });
      };
      // Initial measurement on mount: PM has just laid out the doc, so
      // by next frame `getBoundingClientRect()` is meaningful.
      schedule();
      // Re-chunk whenever the agent snapshot changes.
      //
      // The funnel in `mountDocxEditor` applies PM transactions
      // optimistically *before* mirroring them through the bus, so the
      // page-decorations `apply()` runs against a stale snapshot (PM
      // doc has the new paragraph, but `agent.getSnapshot()` is still
      // pre-Enter). Without this hook the chunker would never re-run
      // for that mutation: the funnel-suppress branch in mount.ts
      // returns without dispatching any tx, and remeasureAndForce
      // can't detect the missing block because the new paragraph never
      // got its `pm-page-block` class in the first place. Result: the
      // freshly inserted paragraph rendered AFTER the page filler,
      // visually pinned to the bottom of the page.
      //
      // Subscribing to the agent gives us a guaranteed signal that the
      // model just changed; we dispatch a `force` meta to re-run
      // computeState with the up-to-date snapshot, which stamps the
      // class on the new block. The follow-up RAF then measures it
      // and the filler shrinks accordingly.
      const unsubscribe = agent.subscribe(() => {
        const tr = view.state.tr.setMeta(pageDecorationsKey, "force");
        view.dispatch(tr);
        schedule();
      });
      return {
        update(updatedView, prevState) {
          // Re-measure whenever the doc changed OR the viewport / DOM
          // structure changed (selection-only transactions don't, so we
          // gate on `docChanged` to avoid pointless work).
          if (updatedView.state.doc !== prevState.doc) schedule();
        },
        destroy() {
          unsubscribe();
          if (scheduledRaf !== null) {
            cancelAnimationFrame(scheduledRaf);
            scheduledRaf = null;
          }
        },
      };
    },
  });
}

function computeState(agent: DocxAgent, state: EditorState, measure: Measure): PageDecorationsState {
  const snapshot = agent.getSnapshot();
  // The chunker tolerates a measure() that returns 0 for blocks it
  // hasn't seen yet (initial render, before the first measurement
  // pass). When every block reports 0, no measured break can fire and
  // we degrade gracefully to "honour hard + hint breaks only" — the
  // pre-measurement behaviour.
  const chunks = chunkIntoPages(snapshot, measure);
  const decorations = buildDecorations(snapshot, chunks, state, measure);
  return { chunks, decorations };
}

/**
 * Walk the editor's top-level DOM children, read each block's CSS-px
 * height via `getBoundingClientRect()`, convert to twips, and if the
 * resulting array differs from `prevCache` dispatch a force-recompute
 * meta so the plugin's `apply` re-runs the chunker against the new
 * heights.
 *
 * The dispatched transaction touches no doc content — it is a
 * meta-only signal. PM still calls `view.update`, which will trigger
 * another `schedule()`, which will read the same DOM heights, find
 * them unchanged, and short-circuit. One extra cycle per real change.
 */
function remeasureAndForce(
  view: EditorView,
  prevCache: ReadonlyArray<number>,
  commit: (next: number[]) => void
): void {
  const editorRoot = view.dom as HTMLElement;
  // PM mounts the editable doc as direct children of `view.dom`; the
  // page-decoration widgets (caps, edges) are also direct children and
  // must NOT be measured into a chunk's content height — they live in
  // the page's *margin* visually. We identify body blocks by the
  // `pm-page-block` class our own decorations stamp on every body
  // child.
  const blocks = editorRoot.querySelectorAll<HTMLElement>(":scope > .pm-page-block");
  const next: number[] = [];
  let changed = blocks.length !== prevCache.length;
  blocks.forEach((el) => {
    const idxAttr = el.getAttribute("data-block-index");
    const blockIndex = idxAttr !== null ? Number.parseInt(idxAttr, 10) : Number.NaN;
    if (!Number.isFinite(blockIndex)) return;
    const h = el.getBoundingClientRect().height;
    const twips = Math.round(h * TWIPS_PER_CSS_PX);
    next[blockIndex] = twips;
    if (prevCache[blockIndex] !== twips) changed = true;
  });
  if (!changed) return;
  commit(next);
  const tr = view.state.tr.setMeta(pageDecorationsKey, "force");
  view.dispatch(tr);
}

function buildDecorations(
  snapshot: DocxSnapshot,
  chunks: ReadonlyArray<PageChunk>,
  state: EditorState,
  measure: Measure
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
    //
    // Phase 1 of docx-fidelity-overhaul: each block also carries the
    // chunk's geometry as inline CSS variables so a doc that mixes
    // portrait and landscape (or different margin profiles) per
    // section renders each page block at its OWN margins instead of
    // the document's first-section margins. The wrapper width itself
    // is still uniform (Phase 5 will introduce per-section sheets).
    const marginLeftCssPx = chunk.geometry.pgMar.left / TWIPS_PER_CSS_PX;
    const marginRightCssPx = chunk.geometry.pgMar.right / TWIPS_PER_CSS_PX;
    const pageWidthCssPx = chunk.geometry.pgSz.w / TWIPS_PER_CSS_PX;
    const blockStyle =
      `--pm-page-margin-left:${marginLeftCssPx}px;` +
      `--pm-page-margin-right:${marginRightCssPx}px;` +
      `--pm-page-width:${pageWidthCssPx}px`;
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
          "data-section-index": String(chunk.sectionIndex),
          // Stamped so the measurement pass in remeasureAndForce can
          // correlate a measured DOM rect back to its body index — PM
          // does not expose top-level child indices on `view.dom`'s
          // DOM children, and counting siblings is fragile because of
          // interleaved widget decorations (caps, edges).
          "data-block-index": String(b),
          style: blockStyle,
        })
      );
    }

    // Word-faithful page chrome (no card sandwich): every chunk has
    // its OWN header band attached at the chunk's start position and
    // its OWN footer band attached at the chunk's end position. The
    // bands paint white sheet + side borders so they read as a
    // continuation of the same piece of paper as the body blocks
    // between them. Between two adjacent chunks we attach a thin
    // transparent `page-break` strip that lets the desk colour show
    // through — the only visible separator between two stacked
    // sheets, with no rounded corners or drop shadow.
    //
    // Side ordering at a position shared between two adjacent chunks
    // (`childPositions[chunk.endBlock] === childPositions[next.startBlock]`):
    //   side: -3  filler  (closes prev chunk's content area)
    //   side: -2  footer  (prev chunk)
    //   side: -1  break   (transparent strip between sheets)
    //   side:  0  header  (next chunk)
    // The first chunk's header sits at position 0 with side: -1.
    // The last chunk's footer sits at docEnd with side: -2 and owns
    // the footnote lane (until per-page footnote distribution
    // lands).

    // Header band for THIS chunk.
    if (i === 0) {
      decos.push(
        Decoration.widget(0, () => renderHeaderBand(chunk, headerPart, "first"), {
          side: -1,
          key: `page-header-band-${chunk.pageNumber}`,
        })
      );
    } else {
      const insertAt = childPositions[chunk.startBlock];
      if (insertAt !== undefined) {
        decos.push(
          Decoration.widget(insertAt, () => renderHeaderBand(chunk, headerPart, "mid"), {
            side: 0,
            key: `page-header-band-${chunk.pageNumber}`,
          })
        );
      }
    }

    // Page filler — pads the page's body region down to the full
    // content area (`pgSz.h - pgMar.top - pgMar.bottom`) so every
    // sheet renders at the document's page height even when the
    // chunk's content is short. Word does this implicitly because it
    // paints fixed-size pages; we have to do it explicitly because PM
    // is a flow editor that hugs content height by default.
    //
    // Sized in CSS pixels (twips ÷ TWIPS_PER_CSS_PX). Sum block
    // heights from the live measurement cache; on the very first
    // mount (cache empty) every block reports 0 and the filler paints
    // the full content area — the next RAF re-runs `apply` with real
    // measurements and the filler shrinks to its true size.
    const contentAreaTwips = Math.max(
      0,
      chunk.geometry.pgSz.h - chunk.geometry.pgMar.top - chunk.geometry.pgMar.bottom
    );
    let measuredTwips = 0;
    for (let b = chunk.startBlock; b < chunk.endBlock; b++) {
      measuredTwips += measure(b);
    }
    const fillerTwips = Math.max(0, contentAreaTwips - measuredTwips);
    const fillerCssPx = fillerTwips / TWIPS_PER_CSS_PX;
    const fillerInsertAt = chunk.endBlock < childPositions.length ? childPositions[chunk.endBlock] : docEnd;
    decos.push(
      Decoration.widget(fillerInsertAt, () => renderPageFiller(fillerCssPx, chunk.pageNumber), {
        side: -3,
        key: `page-filler-${chunk.pageNumber}-${Math.round(fillerCssPx)}`,
      })
    );

    // Footer band for THIS chunk. The last chunk hosts the footnote
    // lane until per-page footnote distribution lands.
    const isLastChunk = i === chunks.length - 1;
    const laneFootnotes = isLastChunk ? collectAuthoredFootnotes(snapshot) : [];
    const laneKey = laneFootnotes.map((fn) => fn.id).join(",");
    decos.push(
      Decoration.widget(
        fillerInsertAt,
        () => renderFooterBand(chunk, footerPart, isLastChunk ? "last" : "mid", laneFootnotes),
        {
          side: -2,
          key: `page-footer-band-${chunk.pageNumber}-fn:${laneKey}`,
        }
      )
    );

    // Page break between THIS chunk and the next: a thin transparent
    // strip showing the desk colour through. No rounded corners, no
    // shadow — the paired top/bottom borders on the bands above and
    // below already imply the page edges.
    if (!isLastChunk) {
      decos.push(
        Decoration.widget(fillerInsertAt, () => renderPageBreak(chunk.pageNumber + 1), {
          side: -1,
          key: `page-break-${chunk.pageNumber}->${chunk.pageNumber + 1}`,
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

/**
 * Word-faithful header band. Lives inside the chunk's own top
 * margin and reads as a continuation of the same white sheet as
 * the body blocks below it. `position` is `"first"` for the first
 * page (the band is the top edge of the document) and `"mid"` for
 * every following chunk (the page-break above it provides the
 * visual separation). The band carries the chunk's geometry so
 * its padding aligns with the page margins.
 */
function renderHeaderBand(
  chunk: PageChunk,
  headerPart: HeaderFooterPart | undefined,
  position: "first" | "mid"
): HTMLElement {
  const band = document.createElement("div");
  band.className = "pm-page-header-band";
  band.dataset.bandPosition = position;
  band.setAttribute("contenteditable", "false");
  band.setAttribute("data-page-number", String(chunk.pageNumber));
  applyChunkGeometryStyle(band, chunk);
  band.appendChild(renderHeaderZone(chunk, headerPart));
  return band;
}

/**
 * Word-faithful footer band. Mirrors the header band: lives
 * inside the chunk's own bottom margin, paints the same white
 * sheet, and aligns its padding with the page margins. The last
 * chunk's band hosts the footnote lane (until per-page footnote
 * distribution ships) above the footer zone proper.
 */
function renderFooterBand(
  chunk: PageChunk,
  footerPart: HeaderFooterPart | undefined,
  position: "mid" | "last",
  laneFootnotes: ReadonlyArray<Footnote> = []
): HTMLElement {
  const band = document.createElement("div");
  band.className = "pm-page-footer-band";
  band.dataset.bandPosition = position;
  band.setAttribute("contenteditable", "false");
  band.setAttribute("data-page-number", String(chunk.pageNumber));
  applyChunkGeometryStyle(band, chunk);
  if (laneFootnotes.length > 0) {
    band.appendChild(renderFootnoteLane(chunk, laneFootnotes));
  }
  band.appendChild(renderFooterZone(chunk, footerPart));
  return band;
}

/**
 * Thin transparent strip between two adjacent page chunks. Lets
 * the desk colour show through so two stacked sheets read as
 * separate pieces of paper, without imposing any chrome of its
 * own (no rounded corners, no shadow). The chunk number stamped
 * in `data-page-number` is the INCOMING page so screen readers
 * announce "page N break" when they hit it.
 */
function renderPageBreak(incomingPageNumber: number): HTMLElement {
  const br = document.createElement("div");
  br.className = "pm-page-break";
  br.setAttribute("contenteditable", "false");
  br.setAttribute("data-page-number", String(incomingPageNumber));
  return br;
}

/**
 * Bottom-of-page footnote lane (F1-UI Phase 1). Renders each
 * authored footnote as a `[N] {plain text}` row above the footer
 * zone, separated from the footer by a thin top rule. Read-only —
 * Phase 2 will introduce in-place editing once we host an editable
 * surface for footnote bodies.
 */
function renderFootnoteLane(chunk: PageChunk, footnotes: ReadonlyArray<Footnote>): HTMLElement {
  const lane = document.createElement("div");
  lane.className = "pm-page-zone pm-page-zone-footnotes";
  lane.setAttribute("contenteditable", "false");
  lane.setAttribute("data-page-zone", "footnotes");
  lane.setAttribute("data-page-number", String(chunk.pageNumber));
  applyChunkGeometryStyle(lane, chunk);

  const inner = document.createElement("div");
  inner.className = "pm-page-zone-content pm-page-footnote-lane";
  inner.setAttribute("data-page-zone-content", "footnotes");

  for (const fn of footnotes) {
    const row = document.createElement("div");
    row.className = "pm-page-footnote-row";
    row.setAttribute("data-footnote-id", String(fn.id));

    const marker = document.createElement("sup");
    marker.className = "pm-page-footnote-marker";
    marker.textContent = `[${fn.id}]`;

    const body = document.createElement("span");
    body.className = "pm-page-footnote-body";
    body.textContent = extractFootnoteText(fn);

    row.appendChild(marker);
    row.appendChild(document.createTextNode(" "));
    row.appendChild(body);
    inner.appendChild(row);
  }

  lane.appendChild(inner);
  return lane;
}

/**
 * Filter a snapshot's footnotes part down to *authored* footnotes —
 * the standard Word `separator` / `continuationSeparator` /
 * `continuationNotice` entries (typically `w:id` -1 / 0 / 1) carry
 * layout glyphs, never user content, and must never appear in the
 * lane.
 */
function collectAuthoredFootnotes(snapshot: DocxSnapshot): ReadonlyArray<Footnote> {
  const part = snapshot.root.footnotesPart;
  if (!part) return [];
  return part.footnotes.filter((fn) => fn.type === "normal");
}

/**
 * Extract a flat-text preview of a footnote body. Mirrors the
 * conventions used by `extractZoneText`: concatenate every
 * paragraph's run text, preserve page-number fields as `[FIELD]`,
 * collapse whitespace, and join multiple paragraphs with " ¶ " so
 * the lane reads naturally on a single line.
 */
function extractFootnoteText(footnote: Footnote): string {
  const parts: string[] = [];
  for (const block of footnote.body) {
    if (block.kind !== "paragraph") continue;
    let acc = "";
    for (const inline of block.children) {
      if (inline.kind !== "run") continue;
      for (const c of inline.children) {
        if (c.kind === "text") acc += c.text;
        else if (c.kind === "page-number-field") acc += `[${c.field}]`;
      }
    }
    const trimmed = acc.replace(/\s+/g, " ").trim();
    if (trimmed.length > 0) parts.push(trimmed);
  }
  return parts.join(" ¶ ");
}

/**
 * Padding widget rendered at the bottom of every page chunk so the
 * visible sheet always reaches the document's full content area
 * height (`pgSz.h - pgMar.top - pgMar.bottom`). When the chunk's
 * measured content already fills the page, this widget collapses to
 * zero. Marked `contenteditable="false"` so the cursor cannot land
 * inside the dead space.
 */
function renderPageFiller(heightCssPx: number, pageNumber: number): HTMLElement {
  const filler = document.createElement("div");
  filler.className = "pm-page-filler";
  filler.setAttribute("contenteditable", "false");
  filler.setAttribute("data-page-number", String(pageNumber));
  filler.style.height = `${Math.max(0, heightCssPx)}px`;
  return filler;
}

function renderHeaderZone(chunk: PageChunk, part: HeaderFooterPart | undefined): HTMLElement {
  const el = renderZone("header", chunk, part);
  applyChunkGeometryStyle(el, chunk);
  return el;
}

function renderFooterZone(chunk: PageChunk, part: HeaderFooterPart | undefined): HTMLElement {
  const el = renderZone("footer", chunk, part);
  applyChunkGeometryStyle(el, chunk);
  return el;
}

/**
 * Stamp a header/footer zone with the chunk's per-section margins so
 * the zone's `padding-left` / `padding-right` (which read
 * `--pm-page-margin-left` / `--pm-page-margin-right` in `globals.css`)
 * align with the body block's margins for the same chunk.
 */
function applyChunkGeometryStyle(el: HTMLElement, chunk: PageChunk): void {
  const left = chunk.geometry.pgMar.left / TWIPS_PER_CSS_PX;
  const right = chunk.geometry.pgMar.right / TWIPS_PER_CSS_PX;
  el.style.setProperty("--pm-page-margin-left", `${left}px`);
  el.style.setProperty("--pm-page-margin-right", `${right}px`);
}

function renderZone(
  slot: "header" | "footer",
  chunk: PageChunk,
  part: HeaderFooterPart | undefined
): HTMLElement {
  const zone = document.createElement("div");
  zone.className = `pm-page-zone pm-page-zone-${slot}`;
  // The wrapper stays read-only (PM widget contract) but the inner
  // content opts back in to contenteditable when there's an actual
  // header/footer part to write into. Documents without parts get
  // an inert hint until B7 (auto-mint numbering & header parts) lands.
  zone.setAttribute("contenteditable", "false");
  zone.setAttribute("data-page-zone", slot);
  zone.setAttribute("data-page-number", String(chunk.pageNumber));
  zone.setAttribute("data-zone-label", slot);
  if (part) {
    zone.setAttribute("data-part-path", part.partPath);
    zone.setAttribute("data-part-target", part.target);
  }

  const inner = document.createElement("div");
  inner.className = "pm-page-zone-content";
  inner.setAttribute("data-page-zone-content", slot);

  if (!part) {
    // Word's empty header/footer is invisible until the user hovers
    // or double-clicks into the zone. We mirror that: the
    // placeholder text is exposed via a data attribute and only
    // painted by CSS on hover. A double-click fires
    // `PAGE_ZONE_MINT_EVENT` so the editor can dispatch
    // `docx:create-header-footer-part` and the next render swaps
    // this no-part shell for an editable zone (Word's
    // "double-click-the-header" affordance).
    inner.classList.add("pm-page-zone-empty", "pm-page-zone-no-part");
    inner.dataset.emptyText = `Doppelklicken zum Bearbeiten der ${slot === "header" ? "Kopfzeile" : "Fußzeile"}`;
    zone.title = `Doppelklick: ${slot === "header" ? "Kopfzeile" : "Fußzeile"} hinzufügen`;
    const onMintDblClick = (ev: MouseEvent): void => {
      ev.preventDefault();
      ev.stopPropagation();
      const detail: PageZoneMintDetail = { slot, pageNumber: chunk.pageNumber };
      zone.dispatchEvent(
        new CustomEvent<PageZoneMintDetail>(PAGE_ZONE_MINT_EVENT, {
          detail,
          bubbles: true,
        })
      );
    };
    zone.addEventListener("dblclick", onMintDblClick);
    inner.addEventListener("dblclick", onMintDblClick);
    zone.appendChild(inner);
    return zone;
  }

  // Build a leaf-id index over the original blocks so the commit
  // path can recover the typed `PageNumberFieldLeaf` and
  // `InlineImageDrawing` instances by their stable id when the
  // user moves a token around inside the contenteditable.
  const leafIndex = buildLeafIndex(part);
  const originalSerialized = JSON.stringify(part.body);
  inner.dataset.original = originalSerialized;

  // Word's in-place authoring: the inner content is contenteditable.
  // PM keeps the outer widget read-only and skips selection mapping
  // for descendants of `contenteditable=false`, but keystrokes inside
  // `contenteditable=true` islands still work in every browser we
  // support. We stop key/wheel propagation so PM's keymap doesn't
  // interpret typing here as a doc edit.
  inner.setAttribute("contenteditable", "true");
  inner.spellcheck = true;

  const isEmptyPart = part.body.every((b) => b.kind === "paragraph" && paragraphIsBlank(b));
  if (isEmptyPart) {
    inner.classList.add("pm-page-zone-empty");
    inner.dataset.placeholder = `Click to add ${slot} text`;
  } else {
    populateZoneContent(inner, part);
  }
  zone.title = `Click to edit ${slot}`;

  const commit = (): void => {
    const blocks = serializeZoneToBlocks(inner, leafIndex);
    const serialized = JSON.stringify(blocks);
    if (serialized === inner.dataset.original) return;
    inner.dataset.original = serialized;
    const detail: PageZoneCommitDetail = {
      slot,
      partPath: part.partPath,
      target: part.target,
      pageNumber: chunk.pageNumber,
      text: blocksToPlainText(blocks),
      blocks,
    };
    zone.dispatchEvent(
      new CustomEvent<PageZoneCommitDetail>(PAGE_ZONE_COMMIT_EVENT, {
        detail,
        bubbles: true,
      })
    );
  };

  inner.addEventListener("focus", () => {
    inner.classList.remove("pm-page-zone-empty");
    if ((inner.textContent ?? "").length === 0 && inner.children.length === 0) {
      // Seed an editable empty paragraph so the caret has a home.
      const line = document.createElement("div");
      line.className = "pm-page-zone-line";
      line.appendChild(document.createElement("br"));
      inner.appendChild(line);
    }
    zone.classList.add("pm-page-zone-focused");
    const detail: PageZoneFocusDetail = {
      slot,
      target: part.target,
      partPath: part.partPath,
      pageNumber: chunk.pageNumber,
    };
    zone.dispatchEvent(
      new CustomEvent<PageZoneFocusDetail>(PAGE_ZONE_FOCUS_EVENT, {
        detail,
        bubbles: true,
      })
    );
  });
  inner.addEventListener("blur", () => {
    const value = inner.textContent ?? "";
    if (value.length === 0 && inner.querySelector(".pm-hf-token") === null) {
      inner.classList.add("pm-page-zone-empty");
    }
    zone.classList.remove("pm-page-zone-focused");
    commit();
    const detail: PageZoneFocusDetail = {
      slot: null,
      target: null,
      partPath: null,
      pageNumber: null,
    };
    zone.dispatchEvent(
      new CustomEvent<PageZoneFocusDetail>(PAGE_ZONE_FOCUS_EVENT, {
        detail,
        bubbles: true,
      })
    );
  });
  inner.addEventListener("keydown", (ev) => {
    if (ev.key === "Escape") {
      ev.preventDefault();
      ev.stopPropagation();
      // Restore from the serialized original snapshot.
      inner.innerHTML = "";
      const original = part;
      populateZoneContent(inner, original);
      inner.dataset.original = JSON.stringify(original.body);
      inner.blur();
      return;
    }
    // Allow Enter to create a new paragraph line — Word does the
    // same. We let the browser handle the actual splitting via the
    // default contenteditable behaviour (browsers insert <div> or
    // <br>; both serialise sanely in `serializeZoneToBlocks`).
    // Stop PM's keymap from interpreting our typing as a doc edit.
    ev.stopPropagation();
  });
  inner.addEventListener("input", (ev) => {
    ev.stopPropagation();
  });

  zone.appendChild(inner);
  return zone;
}

/* ── Rich H/F rendering helpers ────────────────────────────────── */

interface LeafIndex {
  fields: Map<string, PageNumberFieldLeaf>;
  images: Map<string, InlineImageDrawing>;
}

function buildLeafIndex(part: HeaderFooterPart): LeafIndex {
  const fields = new Map<string, PageNumberFieldLeaf>();
  const images = new Map<string, InlineImageDrawing>();
  for (const block of part.body) {
    if (block.kind !== "paragraph") continue;
    for (const inline of block.children) {
      if (inline.kind !== "run") continue;
      for (const c of inline.children) {
        if (c.kind === "page-number-field") fields.set(c.id, c);
        else if (c.kind === "drawing" && c.subkind === "inline-image") images.set(c.id, c);
      }
    }
  }
  return { fields, images };
}

function paragraphIsBlank(p: Paragraph): boolean {
  for (const inline of p.children) {
    if (inline.kind !== "run") continue;
    for (const c of inline.children) {
      if (c.kind === "text" && c.text.length > 0) return false;
      if (c.kind === "page-number-field") return false;
      if (c.kind === "drawing") return false;
    }
  }
  return true;
}

function populateZoneContent(host: HTMLElement, part: HeaderFooterPart): void {
  for (const block of part.body) {
    if (block.kind !== "paragraph") continue;
    host.appendChild(paragraphToLine(block, part));
  }
  if (host.children.length === 0) {
    const line = document.createElement("div");
    line.className = "pm-page-zone-line";
    line.appendChild(document.createElement("br"));
    host.appendChild(line);
  }
}

function paragraphToLine(p: Paragraph, part: HeaderFooterPart): HTMLDivElement {
  const line = document.createElement("div");
  line.className = "pm-page-zone-line";
  line.dataset.paragraphId = p.id;
  let hasContent = false;
  for (const inline of p.children) {
    if (inline.kind !== "run") continue;
    for (const c of inline.children) {
      if (c.kind === "text") {
        if (c.text.length > 0) {
          line.appendChild(document.createTextNode(c.text));
          hasContent = true;
        }
      } else if (c.kind === "page-number-field") {
        line.appendChild(renderFieldToken(c));
        hasContent = true;
      } else if (c.kind === "drawing" && c.subkind === "inline-image") {
        line.appendChild(renderImageToken(c, part));
        hasContent = true;
      }
    }
  }
  if (!hasContent) line.appendChild(document.createElement("br"));
  return line;
}

function renderFieldToken(leaf: PageNumberFieldLeaf): HTMLSpanElement {
  const span = document.createElement("span");
  span.className = "pm-hf-token pm-hf-field";
  span.contentEditable = "false";
  span.dataset.tokenKind = "page-number-field";
  span.dataset.leafId = leaf.id;
  span.dataset.field = leaf.field;
  span.textContent = `[${leaf.field}]`;
  span.title = leaf.field === "PAGE" ? "Page number field" : "Total pages field";
  return span;
}

function renderImageToken(leaf: InlineImageDrawing, part: HeaderFooterPart): HTMLElement {
  // We don't have access to the snapshot's media here without
  // threading it through; the user-visible result is the alt text
  // surrounded by a chip. The actual image bytes are still
  // preserved on commit via the leaf-id round-trip.
  const span = document.createElement("span");
  span.className = "pm-hf-token pm-hf-image";
  span.contentEditable = "false";
  span.dataset.tokenKind = "inline-image";
  span.dataset.leafId = leaf.id;
  span.dataset.relId = leaf.relId;
  span.dataset.partPath = part.partPath;
  span.dataset.cx = String(leaf.cx);
  span.dataset.cy = String(leaf.cy);
  const label = leaf.descr ?? leaf.name ?? "image";
  span.textContent = `[${label}]`;
  span.title = `${leaf.name ?? "Image"} (${Math.round(leaf.cx / 9525)} × ${Math.round(leaf.cy / 9525)} px)`;
  return span;
}

/**
 * Walk the contenteditable surface and rebuild a `BlockNode[]` for
 * the part's body. Each top-level child element (or text run) is
 * one paragraph; inside a paragraph, text nodes become
 * {@link TextLeaf}s, `.pm-hf-field` spans become
 * {@link PageNumberFieldLeaf}s (recovered by `data-leaf-id` from
 * the index when present), and `.pm-hf-image` spans become
 * {@link InlineImageDrawing}s (also recovered by id).
 */
function serializeZoneToBlocks(host: HTMLElement, leafIndex: LeafIndex): BlockNode[] {
  const lines = collectLines(host);
  const blocks: BlockNode[] = [];
  for (const line of lines) {
    const para = lineToParagraph(line, leafIndex);
    blocks.push(para);
  }
  if (blocks.length === 0) {
    blocks.push({
      kind: "paragraph",
      id: synthId("para"),
      properties: {},
      children: [],
    });
  }
  return blocks;
}

function collectLines(host: HTMLElement): HTMLElement[] {
  const lines: HTMLElement[] = [];
  // The browser may produce `<div>`s, `<p>`s, `<br>`s mixed at the
  // top level. We treat each top-level Element child as its own
  // line; text nodes and inline elements floating loose are merged
  // into the previous line (or a fresh first line).
  let current: HTMLElement | null = null;
  for (const node of Array.from(host.childNodes)) {
    if (node.nodeType === Node.ELEMENT_NODE) {
      const el = node as HTMLElement;
      const tag = el.tagName.toLowerCase();
      if (tag === "br") {
        if (current === null) {
          const line = document.createElement("div");
          lines.push(line);
          current = line;
        }
        // BR ends the current line.
        current = null;
      } else if (tag === "div" || tag === "p") {
        lines.push(el);
        current = null;
      } else {
        if (current === null) {
          current = document.createElement("div");
          lines.push(current);
        }
        current.appendChild(el.cloneNode(true));
      }
    } else if (node.nodeType === Node.TEXT_NODE) {
      if (current === null) {
        current = document.createElement("div");
        lines.push(current);
      }
      current.appendChild(node.cloneNode(true));
    }
  }
  return lines;
}

function lineToParagraph(line: HTMLElement, leafIndex: LeafIndex): Paragraph {
  const id = line.dataset?.paragraphId ?? synthId("para");
  const children: InlineNode[] = [];
  let textBuffer: string[] = [];
  const flushText = (): void => {
    if (textBuffer.length === 0) return;
    const joined = textBuffer.join("");
    if (joined.length === 0) {
      textBuffer = [];
      return;
    }
    const leaf: TextLeaf = {
      kind: "text",
      id: synthId("text"),
      text: joined,
      xmlSpacePreserve: /^\s|\s$/.test(joined),
    };
    const run: Run = {
      kind: "run",
      id: synthId("run"),
      properties: {},
      children: [leaf],
    };
    children.push(run);
    textBuffer = [];
  };
  const visit = (node: Node): void => {
    if (node.nodeType === Node.TEXT_NODE) {
      textBuffer.push(node.textContent ?? "");
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return;
    const el = node as HTMLElement;
    const kind = el.dataset?.tokenKind;
    if (kind === "page-number-field") {
      flushText();
      const id = el.dataset.leafId ?? "";
      const original = leafIndex.fields.get(id);
      const field = (el.dataset.field === "NUMPAGES" ? "NUMPAGES" : "PAGE") as "PAGE" | "NUMPAGES";
      const fieldLeaf: PageNumberFieldLeaf = original ?? {
        kind: "page-number-field",
        id: synthId("field"),
        field,
        instr: ` ${field} \\* MERGEFORMAT `,
      };
      const fieldRun: Run = {
        kind: "run",
        id: synthId("run"),
        properties: {},
        children: [fieldLeaf as RunChild],
      };
      children.push(fieldRun);
      return;
    }
    if (kind === "inline-image") {
      flushText();
      const id = el.dataset.leafId ?? "";
      const original = leafIndex.images.get(id);
      if (original) {
        const imgRun: Run = {
          kind: "run",
          id: synthId("run"),
          properties: {},
          children: [original as RunChild],
        };
        children.push(imgRun);
      }
      return;
    }
    if (el.tagName.toLowerCase() === "br") {
      // mid-line BR — collapse to a space so the paragraph's
      // flat-text length stays in sync with what the user sees.
      textBuffer.push(" ");
      return;
    }
    for (const c of Array.from(el.childNodes)) visit(c);
  };
  for (const c of Array.from(line.childNodes)) visit(c);
  flushText();
  return {
    kind: "paragraph",
    id,
    properties: {},
    children,
  };
}

function blocksToPlainText(blocks: ReadonlyArray<BlockNode>): string {
  for (const b of blocks) {
    if (b.kind !== "paragraph") continue;
    let acc = "";
    for (const inline of b.children) {
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

let synthCounter = 0;
function synthId(prefix: string): string {
  synthCounter += 1;
  return `${prefix}-mint-${Date.now().toString(36)}-${synthCounter.toString(36)}`;
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
export function pageNumberForPos(chunks: ReadonlyArray<PageChunk>, state: EditorState, pos: number): number {
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
export function gotoPage(view: EditorView, pageNumber: number, chunks: ReadonlyArray<PageChunk>): boolean {
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
