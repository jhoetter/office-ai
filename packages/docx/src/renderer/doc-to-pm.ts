import { type Mark, type Node as PMNode } from "prosemirror-model";
import { docxSchema } from "./schema.js";
import { classifyOpaqueTag, extractOpaqueText } from "../model/opaque-classification.js";
import type {
  BlockNode,
  DocxSnapshot,
  Hyperlink,
  InlineImageDrawing,
  InlineNode,
  OpaqueXml,
  Paragraph,
  Run,
  RunChild,
  RunProperties,
  Table,
  TableCell,
  TableRow,
} from "../model/types.js";
import type { RenderableTable, RenderableTableCell, RenderableTableRow } from "./schema.js";

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
  const blocks = snapshot.root.body.map((b) => blockToPM(b)).filter((n): n is PMNode => n !== null);
  return docxSchema.nodes.doc.create(null, blocks);
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
      return docxSchema.nodes.section_break.create({ blockId: block.id, rawJson: encode(block.raw) });
    case "opaque-block":
      return opaqueBlockToPM(block.id, block.raw, block.children);
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
      const drawingMeta = child.subkind === "inline-image" ? (child.raw ?? imageStub(child)) : child.raw;
      out.push(
        docxSchema.nodes.image.create({ runId: child.id, drawingJson: encode(drawingMeta) }, null, marks)
      );
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
 * the `table` node's `toDOM`. The renderer flattens cell paragraphs to
 * plain text on purpose: cells are read-only in this iteration, so the
 * extra fidelity (marks, runs, nested tables) would be wasted and would
 * complicate paragraph indexing in `transactionToCommands`.
 *
 * Nested tables are surfaced by joining their cell text with " | " to
 * preserve some structural cue while keeping the projection flat.
 */
function tableToRenderable(table: Table): RenderableTable {
  return { rows: table.rows.map(rowToRenderable) };
}

function rowToRenderable(row: TableRow): RenderableTableRow {
  return {
    header: row.properties.header === true,
    cells: row.cells.map(cellToRenderable),
  };
}

function cellToRenderable(cell: TableCell): RenderableTableCell {
  return {
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
            text: block.rows
              .map((r) =>
                r.cells.map((c) => c.body.map(extractBlockText).filter(Boolean).join(" ")).join(" | ")
              )
              .join("\n"),
          },
        ];
      }
      return [];
    }),
  };
}

function paragraphToRenderable(
  p: Paragraph
): RenderableTable["rows"][number]["cells"][number]["blocks"][number] {
  const out: { kind: "paragraph"; text: string; styleId?: string; alignment?: string } = {
    kind: "paragraph",
    text: paragraphPlainText(p),
  };
  if (p.properties.styleId) out.styleId = p.properties.styleId;
  if (p.properties.alignment) out.alignment = p.properties.alignment;
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
