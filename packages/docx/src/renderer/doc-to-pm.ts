import { type Mark, type Node as PMNode } from "prosemirror-model";
import { docxSchema } from "./schema.js";
import type {
  BlockNode,
  DocxSnapshot,
  Hyperlink,
  InlineImageDrawing,
  InlineNode,
  Paragraph,
  Run,
  RunChild,
  RunProperties,
} from "../model/types.js";

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
      return docxSchema.nodes.table.create({ tableId: block.id, rawJson: encode(block.raw) });
    case "section-break":
      return docxSchema.nodes.section_break.create({ blockId: block.id, rawJson: encode(block.raw) });
    case "opaque-block":
      return docxSchema.nodes.opaque_block.create({
        blockId: block.id,
        rawJson: encode(block.raw),
        tag: block.raw.tag,
      });
    default: {
      const _exhaustive: never = block;
      void _exhaustive;
      return null;
    }
  }
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
    case "opaque-inline":
      out.push(
        docxSchema.nodes.opaque_inline.create({
          inlineId: node.id,
          rawJson: encode(node.raw),
        })
      );
      return;
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
    case "opaque":
      out.push(
        docxSchema.nodes.opaque_inline.create({ inlineId: child.id, rawJson: encode(child.raw) }, null, marks)
      );
      return;
    default: {
      const _exhaustive: never = child;
      void _exhaustive;
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

function encode(value: unknown): string {
  try {
    return JSON.stringify(value);
  } catch {
    return "null";
  }
}
