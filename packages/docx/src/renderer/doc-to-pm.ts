import { type Mark, type Node as PMNode } from "prosemirror-model";
import { docxSchema } from "./schema.js";
import { classifyOpaqueTag, extractOpaqueText } from "../model/opaque-classification.js";
import type {
  BlockNode,
  BorderSide,
  DocxDocument,
  DocxSnapshot,
  Hyperlink,
  InlineImageDrawing,
  InlineNode,
  MediaPart,
  OpaqueXml,
  Paragraph,
  Relationship,
  Run,
  RunChild,
  RunProperties,
  Table,
  TableBorders,
  TableCell,
  TableRow,
} from "../model/types.js";
import type {
  RenderableTable,
  RenderableTableBlock,
  RenderableTableBorderSide,
  RenderableTableBorders,
  RenderableTableCell,
  RenderableTableCellProps,
  RenderableTableProps,
  RenderableTableRow,
  RenderableTableRun,
} from "./schema.js";

const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

interface ResolvedImage {
  dataUrl: string;
  mimeType: string;
}

/**
 * Maps `<a:blip r:embed="rIdN"/>` style references to a base64 `data:`
 * URL of the underlying media bytes. Built once per `docToPM` call from
 * `snapshot.root.relationships` + `snapshot.root.media` and consumed by
 * `pushRunChild` so the editor renders real `<img>` elements instead of
 * the legacy `[image]` chip.
 */
type MediaResolver = (relId: string) => ResolvedImage | null;

/**
 * Module-local during a single synchronous `docToPM` invocation. Set at
 * the entry point and cleared in a `finally` so an exception inside the
 * walker can never leak the resolver into a subsequent call. We use a
 * module-scoped slot rather than threading the resolver through every
 * helper because the renderer is intentionally a "free function" tree
 * walker — adding a `ctx` arg to each of the ~20 helpers below would be
 * pure overhead for a transient render-time concern.
 */
let activeResolver: MediaResolver | null = null;

function imageStub(leaf: InlineImageDrawing): unknown {
  return {
    kind: "inline-image",
    relId: leaf.relId,
    cx: leaf.cx,
    cy: leaf.cy,
    docPrId: leaf.docPrId,
    name: leaf.name,
    ...(leaf.descr !== undefined ? { descr: leaf.descr } : {}),
  };
}

/**
 * Project a `DocxSnapshot` into a ProseMirror document. Reverse direction
 * lives in `pm-to-doc.ts`. See spec/docx/renderer.md.
 *
 * Run boundaries are preserved by NOT coalescing adjacent text fragments
 * with identical marks; this keeps reverse mapping deterministic for
 * unchanged documents.
 */
export function docToPM(snapshot: DocxSnapshot): PMNode {
  activeResolver = buildMediaResolver(snapshot.root);
  try {
    const blocks = snapshot.root.body.map((b) => blockToPM(b)).filter((n): n is PMNode => n !== null);
    return docxSchema.nodes.doc.create(null, blocks);
  } finally {
    activeResolver = null;
  }
}

function buildMediaResolver(doc: DocxDocument): MediaResolver {
  const docRels: ReadonlyArray<Relationship> = doc.relationships.get("word/document.xml") ?? [];
  const cache = new Map<string, ResolvedImage>();
  return (relId: string) => {
    const cached = cache.get(relId);
    if (cached) return cached;
    const rel = docRels.find((r) => r.id === relId && r.type === IMAGE_REL_TYPE);
    if (!rel) return null;
    const partPath = rel.target.startsWith("word/") ? rel.target : `word/${rel.target}`;
    const part = doc.media.get(partPath);
    if (!part) return null;
    const resolved: ResolvedImage = {
      dataUrl: bytesToDataUrl(part),
      mimeType: part.mimeType,
    };
    cache.set(relId, resolved);
    return resolved;
  };
}

function bytesToDataUrl(part: MediaPart): string {
  return `data:${part.mimeType};base64,${bytesToBase64(part.bytes)}`;
}

function bytesToBase64(bytes: Uint8Array): string {
  // Browser-safe; Node 18+ also supports `btoa` via globalThis.
  let bin = "";
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  if (typeof btoa === "function") return btoa(bin);
  // Fallback for older Node test envs without `btoa`.
  return Buffer.from(bin, "binary").toString("base64");
}

function emuToPx(emu: number): number {
  return Math.round(emu / 9525);
}

function blockToPM(block: BlockNode): PMNode | null {
  switch (block.kind) {
    case "paragraph":
      return paragraphToPM(block);
    case "table":
      return docxSchema.nodes.table.create({
        tableId: block.id,
        rawJson: encode(block.raw),
        tableJson: encode(tableToRenderable(block)),
      });
    case "section-break":
      return docxSchema.nodes.section_break.create({
        blockId: block.id,
        rawJson: block.raw ? encode(block.raw) : encode(null),
      });
    case "opaque-block":
      return opaqueBlockToPM(block.id, block.raw, block.children);
    case "wrapper-marker":
      return docxSchema.nodes.wrapper_marker.create({
        blockId: block.id,
        wrapperId: block.wrapperId,
        side: block.side,
        tag: block.wrapperRaw.tag,
      });
    default: {
      const _exhaustive: never = block;
      void _exhaustive;
      return null;
    }
  }
}

/**
 * Project an `OpaqueBlock` into a PM node. We classify the carrier's tag
 * via `classifyOpaqueTag` so structural metadata becomes invisible and
 * content-wrappers (SDT, simple fields, MC fallback) surface their inner
 * text instead of an opaque chip.
 *
 * The model still carries the full subtree byte-for-byte; only the *display*
 * is smartened, so round-trip integrity is unaffected.
 */
function opaqueBlockToPM(
  blockId: string,
  raw: OpaqueXml,
  children?: ReadonlyArray<BlockNode>
): PMNode | null {
  if (children && children.length > 0) {
    return docxSchema.nodes.opaque_block_wrapper.create({
      blockId,
      rawJson: encode(raw),
      tag: raw.tag,
      contentJson: encode({ blocks: children.map(blockToWrapperBlock).filter((b) => b !== null) }),
    });
  }
  const display = classifyOpaqueTag(raw.tag);
  if (display === "metadata") {
    // A metadata-only block (extremely rare at the body level, but possible
    // for things like an isolated `<w:permStart>`) cannot be dropped because
    // PM's `doc` schema requires `block+` and an empty doc is invalid. Fall
    // back to the placeholder so the user at least sees that something
    // structural lives here.
    return docxSchema.nodes.opaque_block.create({
      blockId,
      rawJson: encode(raw),
      tag: raw.tag,
      previewText: null,
    });
  }
  if (display === "content-wrapper") {
    const text = extractOpaqueText(raw);
    return docxSchema.nodes.opaque_block.create({
      blockId,
      rawJson: encode(raw),
      tag: raw.tag,
      previewText: text.length > 0 ? text : null,
    });
  }
  return docxSchema.nodes.opaque_block.create({
    blockId,
    rawJson: encode(raw),
    tag: raw.tag,
    previewText: null,
  });
}

function paragraphToPM(p: Paragraph): PMNode {
  const inlines: PMNode[] = [];
  const activeCommentIds: string[] = [];
  for (const node of p.children) {
    inlinesForChild(node, inlines, activeCommentIds);
  }
  return docxSchema.nodes.paragraph.create(
    {
      paragraphId: p.id,
      styleId: p.properties.styleId ?? null,
      alignment: p.properties.alignment ?? null,
      propsJson: encode(p.properties),
    },
    inlines
  );
}

function inlinesForChild(node: InlineNode, out: PMNode[], activeCommentIds: string[]): void {
  switch (node.kind) {
    case "run":
      pushRun(node, out, activeCommentIds, []);
      return;
    case "hyperlink":
      pushHyperlink(node, out, activeCommentIds);
      return;
    case "comment-range-start":
      activeCommentIds.push(node.commentId);
      return;
    case "comment-range-end": {
      const idx = activeCommentIds.lastIndexOf(node.commentId);
      if (idx >= 0) activeCommentIds.splice(idx, 1);
      return;
    }
    case "comment-reference":
      // Anchor is the rendered side panel concern; nothing to push inline here.
      return;
    case "revision":
      for (const child of node.children) {
        const wrap = (n: PMNode): PMNode =>
          n.mark(
            n.marks.concat(
              docxSchema.marks.revision_mark.create({
                revisionType: node.revisionType,
                author: node.author,
                date: node.date,
                revisionId: node.revisionId,
              })
            )
          );
        const tmp: PMNode[] = [];
        inlinesForChild(child, tmp, activeCommentIds);
        for (const t of tmp) out.push(wrap(t));
      }
      return;
    case "opaque-inline": {
      const pmNode = opaqueInlineToPM(node.id, node.raw, [], node.children);
      if (pmNode) out.push(pmNode);
      return;
    }
    default: {
      const _exhaustive: never = node;
      void _exhaustive;
    }
  }
}

function pushRun(run: Run, out: PMNode[], activeCommentIds: string[], extraMarks: Mark[]): void {
  const baseMarks = runMarks(run.properties);
  const commentMarks = activeCommentIds.map((id) => docxSchema.marks.comment_mark.create({ commentId: id }));
  const marks = baseMarks.concat(commentMarks).concat(extraMarks);
  for (const child of run.children) {
    pushRunChild(child, out, marks);
  }
}

function pushHyperlink(link: Hyperlink, out: PMNode[], activeCommentIds: string[]): void {
  const linkMark = docxSchema.marks.hyperlink.create({
    relationshipId: link.relationshipId ?? null,
    anchor: link.anchor ?? null,
    hyperlinkId: link.id,
  });
  for (const child of link.children) {
    pushRun(child, out, activeCommentIds, [linkMark]);
  }
}

function pushRunChild(child: RunChild, out: PMNode[], marks: Mark[]): void {
  switch (child.kind) {
    case "text":
      if (child.text.length === 0) return;
      out.push(docxSchema.text(child.text, marks));
      return;
    case "break":
      out.push(docxSchema.nodes.hard_break.create({ breakType: child.breakType ?? null }, null, marks));
      return;
    case "page-break":
      // Render as a typed hard_break with breakType="page" so the
      // existing PM <br> projection still applies; the page-frame
      // chunker (P3.3) reads the typed leaf directly off the snapshot
      // without going through PM.
      out.push(docxSchema.nodes.hard_break.create({ breakType: "page" }, null, marks));
      return;
    case "last-rendered-page-break":
      // Layout hint only; never visible. Drop on render — the snapshot
      // model still carries the leaf so it round-trips on save.
      return;
    case "page-number-field":
      // Render as plain text — the cached display value (`"3"`) when
      // present, otherwise a sentinel `#` so the user sees something
      // where the field will resolve. Live page numbers come from the
      // page-decorations plugin which substitutes the correct page
      // index per visible widget; this branch is what the editor
      // surface sees inside header/footer previews and in the body if
      // a page number ever ends up there.
      out.push(
        docxSchema.text(child.cachedText && child.cachedText.length > 0 ? child.cachedText : "#", marks)
      );
      return;
    case "tab":
      out.push(docxSchema.nodes.tab.create(null, null, marks));
      return;
    case "drawing": {
      // The renderer just needs *something* JSON-serializable to round-trip
      // through the editor's image node attribute. For drawings parsed from
      // a real file we always have the cached `raw` subtree; for typed
      // inline-images produced by `docx:insert-image` we synthesize a
      // metadata stub so the editor still gets the relationship id and
      // dimensions.
      // Charts get a structured envelope so the editor can pop the
      // "Edit Data" modal on double-click without re-parsing the
      // entire `<w:drawing>` subtree. Other drawings continue to
      // ride on the raw subtree (or a synthesized stub for typed
      // inline images) verbatim.
      const drawingMeta =
        child.subkind === "chart"
          ? { chart: { chartPartPath: child.chartPartPath } }
          : child.subkind === "inline-image"
            ? (child.raw ?? imageStub(child))
            : child.raw;
      const attrs: Record<string, unknown> = { runId: child.id, drawingJson: encode(drawingMeta) };
      // Inline images: resolve the relationship to a `data:` URL so the
      // schema's `image.toDOM` can emit a real `<img>` tag. Drawings
      // without a resolvable relId fall back to the placeholder chip.
      if (child.subkind === "inline-image") {
        const resolved = activeResolver ? activeResolver(child.relId) : null;
        if (resolved) {
          attrs.dataUrl = resolved.dataUrl;
          attrs.width = emuToPx(child.cx);
          attrs.height = emuToPx(child.cy);
          if (typeof child.descr === "string") attrs.alt = child.descr;
          else if (typeof child.name === "string") attrs.alt = child.name;
        }
      }
      out.push(docxSchema.nodes.image.create(attrs, null, marks));
      return;
    }
    case "embedded-spreadsheet": {
      // Render OLE-embedded Excel as a placeholder image chip carrying
      // its part path so the editor can wire the double-click "Edit
      // Data" flow back to the embedded workbook bytes. The chip is a
      // standalone PM image node so cursor placement / deletion behave
      // like other drawings.
      const attrs: Record<string, unknown> = {
        runId: child.id,
        drawingJson: encode({ embeddedSpreadsheet: child.embeddingPartPath }),
        width: 320,
        height: 220,
        alt: `Embedded spreadsheet (${child.progId})`,
      };
      out.push(docxSchema.nodes.image.create(attrs, null, marks));
      return;
    }
    case "opaque": {
      const pmNode = opaqueInlineToPM(child.id, child.raw, marks);
      if (pmNode) out.push(pmNode);
      return;
    }
    default: {
      const _exhaustive: never = child;
      void _exhaustive;
    }
  }
}

/**
 * Project an opaque inline (carried at the paragraph level OR as a run child)
 * into a PM node, applying the same display classification as the block-level
 * variant. Metadata-only carriers (bookmarks, field characters,
 * `lastRenderedPageBreak`, …) return `null` so the caller emits nothing.
 */
function opaqueInlineToPM(
  inlineId: string,
  raw: OpaqueXml,
  marks: Mark[],
  children?: ReadonlyArray<InlineNode>
): PMNode | null {
  if (children && children.length > 0) {
    const text = children.map(inlinePlainText).join("");
    return docxSchema.nodes.opaque_inline_wrapper.create(
      {
        inlineId,
        rawJson: encode(raw),
        tag: raw.tag,
        contentJson: encode({ text }),
      },
      null,
      marks
    );
  }
  const display = classifyOpaqueTag(raw.tag);
  if (display === "metadata") return null;
  const previewText = display === "content-wrapper" ? extractOpaqueText(raw) : "";
  return docxSchema.nodes.opaque_inline.create(
    {
      inlineId,
      rawJson: encode(raw),
      tag: raw.tag,
      previewText: previewText.length > 0 ? previewText : null,
    },
    null,
    marks
  );
}

/**
 * Project a block child of an unwrapped content-wrapper carrier into the
 * lightweight `WrapperContentBlock` shape consumed by the wrapper node's
 * `toDOM`. Tables and section breaks render as their tag chip; nested
 * opaque-blocks render as their previewText if available.
 */
function blockToWrapperBlock(
  block: BlockNode
): { kind: "paragraph"; text: string; styleId?: string; alignment?: string } | null {
  switch (block.kind) {
    case "paragraph": {
      const out: { kind: "paragraph"; text: string; styleId?: string; alignment?: string } = {
        kind: "paragraph",
        text: paragraphPlainText(block),
      };
      if (block.properties.styleId) out.styleId = block.properties.styleId;
      if (block.properties.alignment) out.alignment = block.properties.alignment;
      return out;
    }
    case "table":
      return { kind: "paragraph", text: "[table]" };
    case "section-break":
      return null;
    case "opaque-block": {
      if (block.children && block.children.length > 0) {
        const flat = block.children.map(blockToWrapperBlock).filter((b) => b !== null);
        const text = flat.map((b) => b!.text).join("\n");
        return { kind: "paragraph", text };
      }
      const t = extractOpaqueText(block.raw);
      return { kind: "paragraph", text: t.length > 0 ? t : `[${block.raw.tag}]` };
    }
    case "wrapper-marker":
      return null;
    default: {
      const _exhaustive: never = block;
      void _exhaustive;
      return null;
    }
  }
}

function runMarks(props: RunProperties): Mark[] {
  const marks: Mark[] = [];
  if (props.bold) marks.push(docxSchema.marks.bold.create());
  if (props.italic) marks.push(docxSchema.marks.italic.create());
  if (props.underline) marks.push(docxSchema.marks.underline.create({ value: props.underline }));
  if (props.strike) marks.push(docxSchema.marks.strikethrough.create());
  if (props.fontFamily) marks.push(docxSchema.marks.font_family.create({ family: props.fontFamily }));
  if (typeof props.fontSize === "number")
    marks.push(docxSchema.marks.font_size.create({ halfPoints: props.fontSize }));
  if (props.color) marks.push(docxSchema.marks.color.create({ rgb: props.color }));
  if (props.highlight) marks.push(docxSchema.marks.highlight.create({ name: props.highlight }));
  return marks;
}

/**
 * Project a typed `Table` into the `RenderableTable` shape consumed by
 * the `table` node's `toDOM`. Phase 2 of the docx-fidelity overhaul
 * upgraded this projection from "plain text per cell" to a structural
 * one that carries the typed run marks (bold, italic, color, font, …)
 * plus per-cell shading / borders / vAlign and per-table border / layout
 * / cell-margin metadata. Cells stay read-only — editability is Phase 6.
 *
 * Nested tables are still flattened (rendering nested table chrome
 * inside an atom node would double the read-only scope without much
 * payoff today). They round-trip via `Table.raw` exactly like before.
 */
function tableToRenderable(table: Table): RenderableTable {
  const out: { -readonly [K in keyof RenderableTable]: RenderableTable[K] } = {
    rows: table.rows.map(rowToRenderable),
  };
  const gridCols = table.grid.map((g) => g.w ?? 0).filter((w): w is number => Number.isFinite(w));
  if (gridCols.length > 0) out.gridCols = gridCols;
  const props = tablePropsToRenderable(table);
  if (props) out.props = props;
  return out;
}

function tablePropsToRenderable(table: Table): RenderableTableProps | undefined {
  const p = table.properties;
  const out: { -readonly [K in keyof RenderableTableProps]: RenderableTableProps[K] } = {};
  if (p.tblBorders) {
    const b = bordersToRenderable(p.tblBorders);
    if (b) out.borders = b;
  }
  if (p.tblCellMar) {
    if (p.tblCellMar.top !== undefined) out.padTop = p.tblCellMar.top;
    if (p.tblCellMar.right !== undefined) out.padRight = p.tblCellMar.right;
    if (p.tblCellMar.bottom !== undefined) out.padBottom = p.tblCellMar.bottom;
    if (p.tblCellMar.left !== undefined) out.padLeft = p.tblCellMar.left;
  }
  if (p.tblLayout) out.layout = p.tblLayout;
  if (p.width) {
    out.widthType = p.width.type;
    if (p.width.type === "dxa") out.widthTw = p.width.value;
    else if (p.width.type === "pct") out.widthPct = p.width.value;
  }
  if (p.jc) out.jc = p.jc;
  if (p.tblInd && p.tblInd.type === "dxa") out.indentTw = p.tblInd.value;
  return Object.keys(out).length > 0 ? out : undefined;
}

function bordersToRenderable(b: TableBorders): RenderableTableBorders | undefined {
  const out: { -readonly [K in keyof RenderableTableBorders]: RenderableTableBorders[K] } = {};
  if (b.top) out.top = borderSideToRenderable(b.top);
  if (b.left) out.left = borderSideToRenderable(b.left);
  if (b.bottom) out.bottom = borderSideToRenderable(b.bottom);
  if (b.right) out.right = borderSideToRenderable(b.right);
  if (b.insideH) out.insideH = borderSideToRenderable(b.insideH);
  if (b.insideV) out.insideV = borderSideToRenderable(b.insideV);
  return Object.keys(out).length > 0 ? out : undefined;
}

function borderSideToRenderable(s: BorderSide): RenderableTableBorderSide {
  const out: { -readonly [K in keyof RenderableTableBorderSide]: RenderableTableBorderSide[K] } = {
    style: s.style,
  };
  if (s.size !== undefined) out.size = s.size;
  if (s.color !== undefined) out.color = s.color;
  return out;
}

function rowToRenderable(row: TableRow): RenderableTableRow {
  return {
    header: row.properties.header === true,
    cells: row.cells.map(cellToRenderable),
  };
}

function cellToRenderable(cell: TableCell): RenderableTableCell {
  const out: { -readonly [K in keyof RenderableTableCell]: RenderableTableCell[K] } = {
    gridSpan: cell.properties.gridSpan ?? 1,
    vMerge: cell.properties.vMerge ?? null,
    blocks: cell.body.flatMap((block) => {
      if (block.kind === "paragraph") return [paragraphToRenderable(block)];
      if (block.kind === "table") {
        // Flatten nested tables to a single paragraph showing the joined
        // cell text. Rare in the wild and a deliberate compromise.
        return [
          {
            kind: "paragraph" as const,
            runs: [
              {
                text: block.rows
                  .map((r) =>
                    r.cells.map((c) => c.body.map(extractBlockText).filter(Boolean).join(" ")).join(" | ")
                  )
                  .join("\n"),
              },
            ],
          } satisfies RenderableTableBlock,
        ];
      }
      return [];
    }),
  };
  const props = cellPropsToRenderable(cell);
  if (props) out.props = props;
  return out;
}

function cellPropsToRenderable(cell: TableCell): RenderableTableCellProps | undefined {
  const p = cell.properties;
  const out: { -readonly [K in keyof RenderableTableCellProps]: RenderableTableCellProps[K] } = {};
  if (p.shd?.fill && p.shd.fill !== "auto") out.shadingFill = p.shd.fill;
  if (p.tcBorders) {
    const b = bordersToRenderable(p.tcBorders);
    if (b) out.borders = b;
  }
  if (p.vAlign) out.vAlign = p.vAlign;
  if (p.tcW && p.tcW.type === "dxa") out.widthTw = p.tcW.value;
  if (p.tcMar) {
    if (p.tcMar.top !== undefined) out.padTop = p.tcMar.top;
    if (p.tcMar.right !== undefined) out.padRight = p.tcMar.right;
    if (p.tcMar.bottom !== undefined) out.padBottom = p.tcMar.bottom;
    if (p.tcMar.left !== undefined) out.padLeft = p.tcMar.left;
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

function paragraphToRenderable(p: Paragraph): RenderableTableBlock {
  const out: { -readonly [K in keyof RenderableTableBlock]: RenderableTableBlock[K] } = {
    kind: "paragraph",
    runs: collectRenderableRuns(p),
  };
  if (p.properties.styleId) out.styleId = p.properties.styleId;
  if (p.properties.alignment) out.alignment = p.properties.alignment;
  return out;
}

/**
 * Walk a paragraph's inline children and collect their visible text plus
 * the typed run marks. Hyperlinks contribute their child runs (the
 * hyperlink mark itself is dropped — table cells are read-only and we
 * already render the underlying typography). Comment ranges, bookmarks,
 * and other metadata-only carriers contribute nothing. Tabs become a
 * literal `\t` so CSS `white-space: pre-wrap` renders them sensibly.
 */
function collectRenderableRuns(p: Paragraph): RenderableTableRun[] {
  const runs: RenderableTableRun[] = [];
  for (const child of p.children) {
    appendRenderableRuns(child, runs);
  }
  return runs;
}

function appendRenderableRuns(node: InlineNode, out: RenderableTableRun[]): void {
  switch (node.kind) {
    case "run": {
      const text = runPlainText(node);
      if (text.length === 0) return;
      out.push(runToRenderable(text, node.properties));
      return;
    }
    case "hyperlink":
      for (const child of node.children) appendRenderableRuns(child, out);
      return;
    case "revision":
      // Show inserted text; suppress deleted text (matches what the main
      // body renderer does for run-level revision marks: the typed
      // `revision_mark` would carry display semantics, which we don't
      // surface inside cells in this phase).
      if (node.revisionType === "del") return;
      for (const child of node.children) appendRenderableRuns(child, out);
      return;
    case "comment-range-start":
    case "comment-range-end":
    case "comment-reference":
      return;
    case "opaque-inline":
      if (node.children && node.children.length > 0) {
        for (const child of node.children) appendRenderableRuns(child, out);
      }
      return;
    default: {
      const _exhaustive: never = node;
      void _exhaustive;
    }
  }
}

function runToRenderable(text: string, props: RunProperties): RenderableTableRun {
  const out: { -readonly [K in keyof RenderableTableRun]: RenderableTableRun[K] } = { text };
  if (props.bold) out.bold = true;
  if (props.italic) out.italic = true;
  if (props.underline !== undefined) out.underline = props.underline;
  if (props.strike) out.strike = true;
  if (props.fontFamily) out.fontFamily = props.fontFamily;
  if (typeof props.fontSize === "number") out.fontSize = props.fontSize;
  if (props.color) out.color = props.color;
  if (props.highlight) out.highlight = props.highlight;
  return out;
}

function paragraphPlainText(p: Paragraph): string {
  let s = "";
  for (const child of p.children) {
    s += inlinePlainText(child);
  }
  return s;
}

function inlinePlainText(node: InlineNode): string {
  switch (node.kind) {
    case "run":
      return runPlainText(node);
    case "hyperlink":
      return node.children.map(runPlainText).join("");
    case "revision":
      return node.children.map(inlinePlainText).join("");
    case "comment-range-start":
    case "comment-range-end":
    case "comment-reference":
      return "";
    case "opaque-inline":
      if (node.children && node.children.length > 0) {
        return node.children.map(inlinePlainText).join("");
      }
      return "";
    default: {
      const _exhaustive: never = node;
      void _exhaustive;
      return "";
    }
  }
}

function runPlainText(run: Run): string {
  let s = "";
  for (const c of run.children) {
    if (c.kind === "text") s += c.text;
    else if (c.kind === "tab") s += "\t";
    else if (c.kind === "break") s += "\n";
  }
  return s;
}

function extractBlockText(b: BlockNode): string {
  if (b.kind === "paragraph") return paragraphPlainText(b);
  return "";
}

function encode(value: unknown): string {
  try {
    return JSON.stringify(value);
  } catch {
    return "null";
  }
}
