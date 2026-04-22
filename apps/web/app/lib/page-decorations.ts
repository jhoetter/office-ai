import { Plugin, PluginKey, TextSelection } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";
import type { EditorState, Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import type { DocxAgent, DocxSnapshot, Footnote, HeaderFooterPart, Measure, PageChunk } from "@officeai/docx";
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
  readonly text: string;
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
    // The filler attaches to the position right after the chunk's
    // last block — that's `childPositions[chunk.endBlock]` for any
    // non-final chunk, and `docEnd` for the very last chunk. Use a
    // more-negative `side` than the page-edge / cap-bottom widgets
    // (which use `side: -1` / `side: 1`) so the filler always renders
    // *above* the page break chrome that closes the page.
    const fillerInsertAt = chunk.endBlock < childPositions.length ? childPositions[chunk.endBlock] : docEnd;
    decos.push(
      Decoration.widget(fillerInsertAt, () => renderPageFiller(fillerCssPx, chunk.pageNumber), {
        side: -2,
        key: `page-filler-${chunk.pageNumber}-${Math.round(fillerCssPx)}`,
      })
    );

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
        Decoration.widget(insertAt, () => renderPageEdge(chunk, nextChunk, footerPart, nextHeaderPart), {
          side: -1,
          key: `page-edge-${chunk.pageNumber}->${nextChunk.pageNumber}`,
        })
      );
    }

    // Page cap (footer zone) for the last page: widget at the very
    // end of the document.
    //
    // F1-UI Phase 1 — footnote lane. Until per-page footnote
    // assignment is modelled (requires line-break / page-break
    // detection that doesn't exist yet) we render the entire
    // `footnotesPart.footnotes` collection on the LAST page, above
    // its footer zone. Future phases will swap `laneFootnotes` for
    // a per-chunk lookup keyed off footnote-reference positions in
    // the chunk's blocks.
    if (i === chunks.length - 1) {
      const laneFootnotes = collectAuthoredFootnotes(snapshot);
      const laneKey = laneFootnotes.map((fn) => fn.id).join(",");
      decos.push(
        Decoration.widget(docEnd, () => renderFooterCap(chunk, footerPart, laneFootnotes), {
          side: 1,
          key: `page-cap-bottom-${chunk.pageNumber}-fn:${laneKey}`,
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

function renderFooterCap(
  chunk: PageChunk,
  footerPart: HeaderFooterPart | undefined,
  laneFootnotes: ReadonlyArray<Footnote> = []
): HTMLElement {
  const cap = document.createElement("div");
  cap.className = "pm-page-cap pm-page-cap-bottom";
  cap.setAttribute("contenteditable", "false");
  if (laneFootnotes.length > 0) {
    cap.appendChild(renderFootnoteLane(chunk, laneFootnotes));
  }
  cap.appendChild(renderFooterZone(chunk, footerPart));
  return cap;
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
  // Word does not paint a "Page N" banner inside the gap between
  // sheets — the active page is surfaced in the status bar instead
  // (see PageStatusBar in DocxEditor.tsx). The gap stays as a pure
  // visual break so the user reads two stacked sheets, but the
  // chrome stays out of the way.
  gap.setAttribute("data-page-number", String(next.pageNumber));
  edge.appendChild(gap);

  edge.appendChild(renderHeaderZone(next, nextHeader));
  return edge;
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

  const text = part ? extractZoneText(part) : "";
  const inner = document.createElement("div");
  inner.className = "pm-page-zone-content";
  inner.setAttribute("data-page-zone-content", slot);

  if (!part) {
    // Word's empty header/footer is invisible until the user hovers
    // or double-clicks into the zone — there's no permanent banner.
    // We mirror that: the placeholder text is exposed via a data
    // attribute and only painted by CSS on hover so the page surface
    // stays uncluttered when no part exists.
    inner.classList.add("pm-page-zone-empty", "pm-page-zone-no-part");
    inner.dataset.emptyText = `No ${slot} for this section`;
    zone.title = `Documents without a ${slot} part are read-only here.`;
    zone.appendChild(inner);
    return zone;
  }

  // Word's in-place authoring: the inner content is contenteditable.
  // PM keeps the outer widget read-only and skips selection mapping
  // for descendants of `contenteditable=false`, but keystrokes inside
  // `contenteditable=true` islands still work in every browser we
  // support. We stop key/wheel propagation so PM's keymap doesn't
  // interpret typing here as a doc edit.
  inner.setAttribute("contenteditable", "true");
  inner.spellcheck = true;
  inner.dataset.original = text;
  if (text.length === 0) {
    inner.classList.add("pm-page-zone-empty");
    inner.dataset.placeholder = `Click to add ${slot} text`;
  } else {
    inner.textContent = text;
  }
  zone.title = `Click to edit ${slot}`;

  const commit = (rawText: string): void => {
    const next = rawText
      .replace(/\s+\n/g, "\n")
      .replace(/\u00a0/g, " ")
      .trimEnd();
    if (next === inner.dataset.original) return;
    inner.dataset.original = next;
    const detail: PageZoneCommitDetail = {
      slot,
      partPath: part.partPath,
      target: part.target,
      pageNumber: chunk.pageNumber,
      text: next,
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
    if (value.length === 0) {
      inner.classList.add("pm-page-zone-empty");
    }
    zone.classList.remove("pm-page-zone-focused");
    commit(value);
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
    if (ev.key === "Enter") {
      ev.preventDefault();
      ev.stopPropagation();
      inner.blur();
      return;
    }
    if (ev.key === "Escape") {
      ev.preventDefault();
      ev.stopPropagation();
      inner.textContent = inner.dataset.original ?? "";
      inner.blur();
      return;
    }
    // Stop PM's keymap from interpreting our typing as a doc edit.
    ev.stopPropagation();
  });
  inner.addEventListener("input", (ev) => {
    ev.stopPropagation();
  });

  zone.appendChild(inner);
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
