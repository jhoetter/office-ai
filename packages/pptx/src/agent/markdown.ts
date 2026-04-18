import type {
  PptxSnapshot,
  Shape,
  Slide,
  TableShape,
  TextShape,
  TextParagraph,
} from "../model/types.js";

/**
 * Render a presentation snapshot to Markdown for human + agent inspection.
 * One section per slide, with text frames as nested blockquotes annotated
 * with shape ids and EMU bounding boxes (so an agent can pinpoint a shape
 * for set-position / set-text without reading XML).
 */
export function snapshotToMarkdown(snap: PptxSnapshot): string {
  const lines: string[] = [];
  lines.push(`# Presentation`);
  lines.push("");
  lines.push(
    `Slide size: ${snap.root.slideSize.cxEmu} × ${snap.root.slideSize.cyEmu} EMU`
  );
  lines.push(`Slides: ${snap.root.slides.length}`);
  lines.push("");
  for (let i = 0; i < snap.root.slides.length; i++) {
    lines.push(...renderSlide(i + 1, snap.root.slides[i]));
    lines.push("");
  }
  return lines.join("\n");
}

function renderSlide(num: number, slide: Slide): string[] {
  const out: string[] = [];
  out.push(`## Slide ${num} — \`${slide.partPath}\` (slideId=${slide.slideId})`);
  out.push("");
  for (const sh of slide.shapes) {
    out.push(...renderShape(sh, 0));
  }
  return out;
}

function renderShape(sh: Shape, depth: number): string[] {
  const indent = "  ".repeat(depth);
  const out: string[] = [];
  const bbox = bboxString(sh);
  switch (sh.kind) {
    case "text":
      out.push(`${indent}- **text** \`${sh.id}\` cNvPr=${sh.cNvPrId} ${bbox}`);
      out.push(...renderTextBody(sh, depth + 1));
      break;
    case "pic":
      out.push(
        `${indent}- **picture** \`${sh.id}\` cNvPr=${sh.cNvPrId} media=\`${sh.mediaPartPath || "?"}\` ${bbox}`
      );
      break;
    case "group":
      out.push(`${indent}- **group** \`${sh.id}\` cNvPr=${sh.cNvPrId} ${bbox}`);
      for (const c of sh.children) out.push(...renderShape(c, depth + 1));
      break;
    case "table":
      out.push(
        `${indent}- **table** \`${sh.id}\` cNvPr=${sh.cNvPrId} ${sh.rows.length}×${sh.columnWidths.length} ${bbox}`
      );
      out.push(...renderTable(sh, depth + 1));
      break;
    case "opaque":
      out.push(`${indent}- **opaque** \`${sh.id}\` <${sh.tag}> ${bbox}`);
      break;
  }
  return out;
}

function renderTextBody(shape: TextShape, depth: number): string[] {
  const indent = "  ".repeat(depth);
  return shape.txBody.paragraphs.map((p) => `${indent}> ${paragraphText(p) || "(empty)"}`);
}

function renderTable(table: TableShape, depth: number): string[] {
  const indent = "  ".repeat(depth);
  const out: string[] = [];
  for (let r = 0; r < table.rows.length; r++) {
    const row = table.rows[r];
    const cells = row.cells.map((c) => {
      const txt = c.txBody.paragraphs.map((p) => paragraphText(p)).join(" / ").trim();
      return txt.length > 0 ? txt : "(empty)";
    });
    out.push(`${indent}| ${cells.join(" | ")} |`);
    if (r === 0) {
      out.push(`${indent}|${row.cells.map(() => "---").join("|")}|`);
    }
  }
  return out;
}

export function paragraphText(p: TextParagraph): string {
  return p.runs.map((r) => (r.isLineBreak ? "\n" : r.text)).join("");
}

function bboxString(sh: Shape): string {
  const p = sh.position;
  const s = sh.size;
  if (!p && !s) return "";
  const px = p ? `(${p.xEmu}, ${p.yEmu})` : "(?, ?)";
  const sx = s ? `${s.cxEmu}×${s.cyEmu}` : "?×?";
  return `@ ${px} ${sx}`;
}
