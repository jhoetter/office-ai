import type {
  ChartPart,
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
  lines.push(`Slide size: ${snap.root.slideSize.cxEmu} × ${snap.root.slideSize.cyEmu} EMU`);
  lines.push(`Slides: ${snap.root.slides.length}`);
  lines.push("");
  for (let i = 0; i < snap.root.slides.length; i++) {
    lines.push(...renderSlide(i + 1, snap.root.slides[i], snap.root.charts));
    lines.push("");
  }
  return lines.join("\n");
}

function renderSlide(num: number, slide: Slide, charts: ReadonlyMap<string, ChartPart>): string[] {
  const out: string[] = [];
  out.push(`## Slide ${num} — \`${slide.partPath}\` (slideId=${slide.slideId})`);
  out.push("");
  if (slide.transition) {
    const speed = slide.transition.speed ? ` (${slide.transition.speed})` : "";
    out.push(`- _transition_: **${slide.transition.kind}**${speed}`);
  }
  if (slide.animations.length > 0) {
    out.push(`- _animations_:`);
    for (const a of slide.animations) {
      const dur = a.durationMs !== undefined ? ` ${a.durationMs}ms` : "";
      const dir = a.direction ? ` ${a.direction}` : "";
      const trig = a.trigger && a.trigger !== "onClick" ? ` (${a.trigger})` : "";
      out.push(
        `  - \`${a.id}\` ${a.order + 1}. **${a.category}/${a.preset}**${dir}${dur}${trig} → cNvPr=${a.targetCNvPrId}`
      );
    }
  }
  if (slide.transition || slide.animations.length > 0) out.push("");
  for (const sh of slide.shapes) {
    out.push(...renderShape(sh, 0, charts));
  }
  return out;
}

function renderShape(sh: Shape, depth: number, charts: ReadonlyMap<string, ChartPart>): string[] {
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
      for (const c of sh.children) out.push(...renderShape(c, depth + 1, charts));
      break;
    case "table":
      out.push(
        `${indent}- **table** \`${sh.id}\` cNvPr=${sh.cNvPrId} ${sh.rows.length}×${sh.columnWidths.length} ${bbox}`
      );
      out.push(...renderTable(sh, depth + 1));
      break;
    case "chart": {
      const part = charts.get(sh.chartPartPath);
      const t = part?.chartType ?? "?";
      out.push(
        `${indent}- **chart** \`${sh.id}\` cNvPr=${sh.cNvPrId} type=${t} part=\`${sh.chartPartPath || "?"}\` ${bbox}`
      );
      if (part) out.push(...renderChartPart(part, depth + 1));
      break;
    }
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
      const txt = c.txBody.paragraphs
        .map((p) => paragraphText(p))
        .join(" / ")
        .trim();
      return txt.length > 0 ? txt : "(empty)";
    });
    out.push(`${indent}| ${cells.join(" | ")} |`);
    if (r === 0) {
      out.push(`${indent}|${row.cells.map(() => "---").join("|")}|`);
    }
  }
  return out;
}

function renderChartPart(part: ChartPart, depth: number): string[] {
  const indent = "  ".repeat(depth);
  const out: string[] = [];
  if (part.title) out.push(`${indent}> title: ${part.title}`);
  if (part.categories.length > 0) {
    out.push(`${indent}> categories: ${part.categories.join(", ")}`);
  }
  for (const ser of part.series) {
    const name = ser.name ? `${ser.name}: ` : "";
    out.push(`${indent}> ${name}[${ser.values.join(", ")}]`);
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
