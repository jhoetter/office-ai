/**
 * `office-agent pptx …` subcommand group.
 *
 * Mirrors the docx subcommand surface: `inspect`, `read`, `search`,
 * `apply`, `diff` plus one subcommand per typed P0/P1 PPTX command.
 * Lives in its own module so cli.ts doesn't balloon as more formats
 * land. The shared helpers (mutation summary, JSON normalization, IO
 * stream wrapper) come from cli.ts via re-exports below — the goal is a
 * single CLI binary, not a separate one per format.
 */

import { readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import type { Command as CommanderCommand } from "commander";
import { Option } from "commander";
import {
  PptxAgent,
  snapshotToMarkdown,
  type PptxSearchResult,
  type PptxSnapshot,
  type Shape,
  type Slide,
  type TextShape,
} from "@officeai/pptx";
import { deterministicIdMinter } from "@officeai/core";
import type { Command as BusCommand, CommandLite } from "@officeai/core";

// ── IO stream type re-export so cli.ts can pass through its own ────────
export interface IO {
  stdout: NodeJS.WritableStream;
  stderr: NodeJS.WritableStream;
}

const MIME_BY_EXT: Readonly<Record<string, string>> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  tif: "image/tiff",
  tiff: "image/tiff",
  webp: "image/webp",
  svg: "image/svg+xml",
};

/**
 * Register every `pptx` subcommand under the supplied parent. The parent
 * is typically the `pptx` Command created in cli.ts, but tests can pass
 * any Commander `Command` to register them in isolation.
 */
export function registerPptxSubcommands(pptx: CommanderCommand, io: IO): void {
  pptx
    .command("inspect")
    .description("Print a structural summary (slides, shapes, masters, layouts) as JSON.")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(async (opts: { file: string; pretty: boolean }) => {
      const agent = await loadAgent(opts.file);
      io.stdout.write(stringifyJson(inspectSnapshot(agent.getSnapshot()), opts.pretty) + "\n");
    });

  pptx
    .command("read")
    .description("Read a PPTX file as Markdown, structured JSON, or plain text.")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .addOption(
      new Option("--format <fmt>", "Output format")
        .choices(["markdown", "json", "text"])
        .default("markdown")
    )
    .addOption(new Option("--slide <n>", "Restrict to a single 0-based slide index").argParser(parseIntArg))
    .option("--pretty", "Pretty-print JSON output (only with --format json)", false)
    .action(
      async (opts: {
        file: string;
        format: "markdown" | "json" | "text";
        slide?: number;
        pretty: boolean;
      }) => {
        const agent = await loadAgent(opts.file);
        const snap = agent.getSnapshot();
        const range = sliceRange(snap, opts.slide);
        switch (opts.format) {
          case "markdown":
            io.stdout.write(renderMarkdown(snap, range) + "\n");
            return;
          case "text":
            io.stdout.write(renderPlainText(snap, range) + "\n");
            return;
          case "json":
            io.stdout.write(
              stringifyJson(snapshotToJsonProjection(snap, range), opts.pretty) + "\n"
            );
            return;
          default: {
            const _exhaustive: never = opts.format;
            void _exhaustive;
          }
        }
      }
    );

  pptx
    .command("search")
    .description("Search every slide's text content; emits matches as JSON.")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("-q, --query <text>", "Search query")
    .option("--case-sensitive", "Case-sensitive search", false)
    .option("--regex", "Treat the query as a regular expression", false)
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        query: string;
        caseSensitive: boolean;
        regex: boolean;
        pretty: boolean;
      }) => {
        const agent = await loadAgent(opts.file);
        const results: PptxSearchResult[] = agent.search({
          query: opts.query,
          caseSensitive: opts.caseSensitive,
          regex: opts.regex,
        });
        io.stdout.write(stringifyJson(results, opts.pretty) + "\n");
      }
    );

  // ── Slide lifecycle ────────────────────────────────────────────────
  pptx
    .command("add-slide")
    .description("Append (or insert) a blank slide and write the result.")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .addOption(new Option("--at <n>", "0-based insert position (default: append)").argParser(parseIntArg))
    .addOption(new Option("--layout <partPath>", "Layout part path to clone placeholders from"))
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: { file: string; at?: number; layout?: string; out?: string; pretty: boolean }) => {
        await dispatchAndWrite(opts, io, [
          {
            type: "pptx:add-slide",
            payload: {
              ...(opts.at !== undefined ? { at: opts.at } : {}),
              ...(opts.layout ? { layoutPartPath: opts.layout } : {}),
            },
          },
        ]);
      }
    );

  pptx
    .command("delete-slide")
    .description("Delete a slide by 0-based index and write the result.")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(async (opts: { file: string; slide: number; out?: string; pretty: boolean }) => {
      await dispatchAndWrite(opts, io, [
        { type: "pptx:delete-slide", payload: { slideIndex: opts.slide } },
      ]);
    });

  pptx
    .command("duplicate-slide")
    .description("Deep-clone a slide and insert it after the source.")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index to duplicate", parseIntArg)
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(async (opts: { file: string; slide: number; out?: string; pretty: boolean }) => {
      await dispatchAndWrite(opts, io, [
        { type: "pptx:duplicate-slide", payload: { slideIndex: opts.slide } },
      ]);
    });

  pptx
    .command("move-slide")
    .description("Reorder a slide from --from to --to (both 0-based).")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--from <n>", "Source 0-based slide index", parseIntArg)
    .requiredOption("--to <n>", "Target 0-based slide index", parseIntArg)
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: { file: string; from: number; to: number; out?: string; pretty: boolean }) => {
        await dispatchAndWrite(opts, io, [
          { type: "pptx:move-slide", payload: { from: opts.from, to: opts.to } },
        ]);
      }
    );

  // ── Shape edits ────────────────────────────────────────────────────
  pptx
    .command("set-text")
    .description("Replace a TextShape's plain text content (\\n separates paragraphs).")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .requiredOption("--shape <id>", "Shape NodeId (as exposed by `pptx inspect`)")
    .requiredOption("--text <text>", "Replacement text")
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        slide: number;
        shape: string;
        text: string;
        out?: string;
        pretty: boolean;
      }) => {
        await dispatchAndWrite(opts, io, [
          {
            type: "pptx:set-text",
            payload: { slideIndex: opts.slide, shapeId: opts.shape, text: opts.text },
          },
        ]);
      }
    );

  pptx
    .command("set-position")
    .description("Set a shape's <a:off> position in EMU.")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .requiredOption("--shape <id>", "Shape NodeId")
    .requiredOption("--x <emu>", "X in EMU", parseIntArg)
    .requiredOption("--y <emu>", "Y in EMU", parseIntArg)
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        slide: number;
        shape: string;
        x: number;
        y: number;
        out?: string;
        pretty: boolean;
      }) => {
        await dispatchAndWrite(opts, io, [
          {
            type: "pptx:set-position",
            payload: { slideIndex: opts.slide, shapeId: opts.shape, x: opts.x, y: opts.y },
          },
        ]);
      }
    );

  pptx
    .command("set-size")
    .description("Set a shape's <a:ext> width/height in EMU.")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .requiredOption("--shape <id>", "Shape NodeId")
    .requiredOption("--width <emu>", "Width in EMU (>0)", parseIntArg)
    .requiredOption("--height <emu>", "Height in EMU (>0)", parseIntArg)
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        slide: number;
        shape: string;
        width: number;
        height: number;
        out?: string;
        pretty: boolean;
      }) => {
        await dispatchAndWrite(opts, io, [
          {
            type: "pptx:set-size",
            payload: {
              slideIndex: opts.slide,
              shapeId: opts.shape,
              width: opts.width,
              height: opts.height,
            },
          },
        ]);
      }
    );

  pptx
    .command("format-text")
    .description("Apply formatting (bold/italic/underline/strike/color/font/size) to a text range.")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .requiredOption("--shape <id>", "Shape NodeId")
    .requiredOption("--paragraph <n>", "0-based paragraph index inside txBody", parseIntArg)
    .requiredOption("--start <n>", "Inclusive char offset within the paragraph", parseIntArg)
    .requiredOption("--end <n>", "Exclusive char offset within the paragraph", parseIntArg)
    .option("--bold", "Toggle bold on", false)
    .option("--no-bold", "Toggle bold off")
    .option("--italic", "Toggle italic on", false)
    .option("--no-italic", "Toggle italic off")
    .option("--underline", "Toggle underline on", false)
    .option("--no-underline", "Toggle underline off")
    .option("--strike", "Toggle strike on", false)
    .option("--no-strike", "Toggle strike off")
    .option("--font <family>", "Font family (e.g. Calibri)")
    .option("--font-size <hundredths>", "Font size in hundredths of a point", parseIntArg)
    .option("--color <hex>", "Hex color without leading # (e.g. FF0000)")
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        slide: number;
        shape: string;
        paragraph: number;
        start: number;
        end: number;
        bold?: boolean;
        italic?: boolean;
        underline?: boolean;
        strike?: boolean;
        font?: string;
        fontSize?: number;
        color?: string;
        out?: string;
        pretty: boolean;
      }) => {
        const format: Record<string, unknown> = {};
        if (opts.bold !== undefined) format.bold = opts.bold;
        if (opts.italic !== undefined) format.italic = opts.italic;
        if (opts.underline !== undefined) format.underline = opts.underline;
        if (opts.strike !== undefined) format.strike = opts.strike;
        if (opts.font) format.fontFamily = opts.font;
        if (opts.fontSize !== undefined) format.fontSizeHundredths = opts.fontSize;
        if (opts.color) format.color = opts.color;
        await dispatchAndWrite(opts, io, [
          {
            type: "pptx:format-text",
            payload: {
              slideIndex: opts.slide,
              shapeId: opts.shape,
              range: { paragraph: opts.paragraph, start: opts.start, end: opts.end },
              format,
            },
          },
        ]);
      }
    );

  pptx
    .command("insert-image")
    .description("Insert an image (PNG/JPEG/etc.) on a slide. SHA-256 dedup'd against existing media.")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .requiredOption("--image <path>", "Path to the image file on disk")
    .requiredOption("--x <emu>", "X in EMU", parseIntArg)
    .requiredOption("--y <emu>", "Y in EMU", parseIntArg)
    .requiredOption("--width <emu>", "Width in EMU (>0)", parseIntArg)
    .requiredOption("--height <emu>", "Height in EMU (>0)", parseIntArg)
    .option("--mime <type>", "Override MIME type (defaults inferred from extension)")
    .option("--alt <text>", "Alt text for accessibility")
    .option("--name <name>", "Override <p:cNvPr name>")
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        slide: number;
        image: string;
        x: number;
        y: number;
        width: number;
        height: number;
        mime?: string;
        alt?: string;
        name?: string;
        out?: string;
        pretty: boolean;
      }) => {
        const buf = await readFile(resolve(opts.image));
        const mimeType = opts.mime ?? inferMime(opts.image);
        if (!mimeType) {
          throw new Error(
            `Could not infer MIME from "${opts.image}". Pass --mime image/png (or similar).`
          );
        }
        await dispatchAndWrite(opts, io, [
          {
            type: "pptx:insert-image",
            payload: {
              slideIndex: opts.slide,
              data: new Uint8Array(buf),
              mimeType,
              x: opts.x,
              y: opts.y,
              width: opts.width,
              height: opts.height,
              ...(opts.alt ? { altText: opts.alt } : {}),
              ...(opts.name ? { name: opts.name } : {}),
            },
          },
        ]);
      }
    );

  pptx
    .command("add-text-box")
    .description("Append a fresh text box on a slide.")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .requiredOption("--text <text>", "Initial text content")
    .requiredOption("--x <emu>", "X in EMU", parseIntArg)
    .requiredOption("--y <emu>", "Y in EMU", parseIntArg)
    .requiredOption("--width <emu>", "Width in EMU (>0)", parseIntArg)
    .requiredOption("--height <emu>", "Height in EMU (>0)", parseIntArg)
    .option("--name <name>", "Override <p:cNvPr name>")
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        slide: number;
        text: string;
        x: number;
        y: number;
        width: number;
        height: number;
        name?: string;
        out?: string;
        pretty: boolean;
      }) => {
        await dispatchAndWrite(opts, io, [
          {
            type: "pptx:add-text-box",
            payload: {
              slideIndex: opts.slide,
              text: opts.text,
              x: opts.x,
              y: opts.y,
              width: opts.width,
              height: opts.height,
              ...(opts.name ? { name: opts.name } : {}),
            },
          },
        ]);
      }
    );

  pptx
    .command("apply")
    .description("Apply a JSON command file (single command or { commands: [...] }) and write the result.")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("-c, --commands <path>", "Path to a JSON file containing one or more commands")
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(async (opts: { file: string; commands: string; out?: string; pretty: boolean }) => {
      const raw = await readFile(resolve(opts.commands), "utf8");
      const data: unknown = JSON.parse(raw);
      const cmds = normalizeCommands(data);
      await dispatchAndWrite(opts, io, cmds);
    });

  pptx
    .command("diff")
    .description("Compute a structural diff between two PPTX files.")
    .requiredOption("--before <path>", "Path to the baseline .pptx file")
    .requiredOption("--after <path>", "Path to the modified .pptx file")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(async (opts: { before: string; after: string; pretty: boolean }) => {
      const before = await loadAgent(opts.before);
      const after = await loadAgent(opts.after);
      io.stdout.write(
        stringifyJson(diffSnapshots(before.getSnapshot(), after.getSnapshot()), opts.pretty) + "\n"
      );
    });
}

// ── Helpers (also used by mcp.ts) ────────────────────────────────────────

export async function loadAgent(input: string): Promise<PptxAgent> {
  const buf = await readFile(resolve(input));
  // Tests need stable shape NodeIds across multiple CLI invocations
  // against the same fixture. The default UUID minter changes them on
  // every parse, which makes ID-based commands (set-text/format-text/
  // set-position/set-size) fail across invocations. Opt-in env var.
  if (process.env.OFFICEAI_DETERMINISTIC_IDS === "1") {
    return PptxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter("n") });
  }
  return PptxAgent.fromBuffer(buf);
}

/**
 * Lightweight, structural snapshot summary intended for `pptx inspect`
 * and the `pptx_inspect` MCP tool. Counts only — heavy projection lives
 * in `snapshotToJsonProjection`.
 */
export interface PptxSnapshotSummary {
  format: "pptx";
  revision: number;
  slideSize: { cxEmu: number; cyEmu: number };
  slides: number;
  layouts: number;
  masters: number;
  mediaParts: number;
  parts: string[];
  shapeCounts: { text: number; pic: number; group: number; opaque: number };
}

export function inspectSnapshot(snap: PptxSnapshot): PptxSnapshotSummary {
  const counts = { text: 0, pic: 0, group: 0, opaque: 0 };
  for (const slide of snap.root.slides) walkShapes(slide.shapes, (s) => bumpShapeKind(counts, s));
  const parts = Array.from(snap.container.parts.keys()).sort();
  return {
    format: "pptx",
    revision: snap.revision,
    slideSize: { cxEmu: snap.root.slideSize.cxEmu, cyEmu: snap.root.slideSize.cyEmu },
    slides: snap.root.slides.length,
    layouts: snap.root.layouts.size,
    masters: snap.root.masters.size,
    mediaParts: snap.root.media.size,
    parts,
    shapeCounts: counts,
  };
}

/**
 * JSON projection of a snapshot — slide index + part path + flat
 * per-shape text. Designed to be small enough to pipe to `jq` while
 * still carrying enough structure for an LLM to address shapes.
 */
export interface PptxSnapshotProjection {
  format: "pptx";
  revision: number;
  slides: ReadonlyArray<{
    index: number;
    id: string;
    partPath: string;
    slideId: number;
    shapes: ReadonlyArray<{
      id: string;
      kind: Shape["kind"];
      cNvPrId?: number;
      name?: string;
      bbox?: { x: number; y: number; cx: number; cy: number };
      text?: string;
    }>;
  }>;
}

export function snapshotToJsonProjection(
  snap: PptxSnapshot,
  range?: { startSlide: number; endSlide: number }
): PptxSnapshotProjection {
  const slides = snap.root.slides;
  const lo = range ? Math.max(0, range.startSlide) : 0;
  const hi = range ? Math.min(slides.length, range.endSlide) : slides.length;
  const out: Array<PptxSnapshotProjection["slides"][number]> = [];
  for (let i = lo; i < hi; i++) {
    const s = slides[i];
    out.push({
      index: i,
      id: s.id,
      partPath: s.partPath,
      slideId: s.slideId,
      shapes: s.shapes.map((sh) => projectShape(sh)),
    });
  }
  return { format: "pptx", revision: snap.revision, slides: out };
}

function projectShape(sh: Shape): PptxSnapshotProjection["slides"][number]["shapes"][number] {
  const base: {
    id: string;
    kind: Shape["kind"];
    cNvPrId?: number;
    name?: string;
    bbox?: { x: number; y: number; cx: number; cy: number };
    text?: string;
  } = { id: sh.id, kind: sh.kind };
  if ("cNvPrId" in sh) base.cNvPrId = sh.cNvPrId;
  if ("name" in sh && sh.name) base.name = sh.name;
  if (sh.position && sh.size) {
    base.bbox = {
      x: sh.position.xEmu,
      y: sh.position.yEmu,
      cx: sh.size.cxEmu,
      cy: sh.size.cyEmu,
    };
  }
  if (sh.kind === "text") base.text = textShapePlainText(sh);
  return base;
}

function textShapePlainText(s: TextShape): string {
  return s.txBody.paragraphs
    .map((p) => p.runs.map((r) => (r.isLineBreak ? "\n" : r.text)).join(""))
    .join("\n");
}

/**
 * Lightweight diff for `pptx diff` and `pptx_diff`. Compares slide
 * count, per-slide shape counts and per-shape text. Not byte-level —
 * the command bus is the canonical source for byte diffs.
 */
export interface PptxDiffSummary {
  format: "pptx";
  slides: { added: number; removed: number; modified: number };
  shapes: { added: number; removed: number; modified: number };
  changes: ReadonlyArray<{
    kind:
      | "slide-added"
      | "slide-removed"
      | "slide-modified"
      | "shape-added"
      | "shape-removed"
      | "shape-modified";
    slideIndex?: number;
    shapeId?: string;
    summary: string;
  }>;
}

export function diffSnapshots(before: PptxSnapshot, after: PptxSnapshot): PptxDiffSummary {
  const changes: Array<PptxDiffSummary["changes"][number]> = [];
  const sa = { added: 0, removed: 0, modified: 0 };
  const shp = { added: 0, removed: 0, modified: 0 };
  const max = Math.max(before.root.slides.length, after.root.slides.length);
  for (let i = 0; i < max; i++) {
    const b = before.root.slides[i];
    const a = after.root.slides[i];
    if (!b && a) {
      sa.added++;
      changes.push({ kind: "slide-added", slideIndex: i, summary: `+slide [${i}] ${a.partPath}` });
      continue;
    }
    if (b && !a) {
      sa.removed++;
      changes.push({ kind: "slide-removed", slideIndex: i, summary: `-slide [${i}] ${b.partPath}` });
      continue;
    }
    if (!b || !a) continue;
    const bShapes = new Map(flattenShapes(b).map((s) => [s.id, s] as const));
    const aShapes = new Map(flattenShapes(a).map((s) => [s.id, s] as const));
    let touched = false;
    for (const [id, ash] of aShapes) {
      const bsh = bShapes.get(id);
      if (!bsh) {
        shp.added++;
        touched = true;
        changes.push({
          kind: "shape-added",
          slideIndex: i,
          shapeId: id,
          summary: `+shape ${ash.kind} ${id}`,
        });
        continue;
      }
      if (shapeChanged(bsh, ash)) {
        shp.modified++;
        touched = true;
        changes.push({
          kind: "shape-modified",
          slideIndex: i,
          shapeId: id,
          summary: `~shape ${ash.kind} ${id}`,
        });
      }
    }
    for (const [id, bsh] of bShapes) {
      if (!aShapes.has(id)) {
        shp.removed++;
        touched = true;
        changes.push({
          kind: "shape-removed",
          slideIndex: i,
          shapeId: id,
          summary: `-shape ${bsh.kind} ${id}`,
        });
      }
    }
    if (touched) {
      sa.modified++;
      changes.push({ kind: "slide-modified", slideIndex: i, summary: `~slide [${i}]` });
    }
  }
  return { format: "pptx", slides: sa, shapes: shp, changes };
}

function shapeChanged(b: Shape, a: Shape): boolean {
  if (b.kind !== a.kind) return true;
  const bbb = b.position && b.size ? `${b.position.xEmu},${b.position.yEmu},${b.size.cxEmu},${b.size.cyEmu}` : "";
  const abb = a.position && a.size ? `${a.position.xEmu},${a.position.yEmu},${a.size.cxEmu},${a.size.cyEmu}` : "";
  if (bbb !== abb) return true;
  if (b.kind === "text" && a.kind === "text") {
    if (textShapePlainText(b) !== textShapePlainText(a)) return true;
  }
  return false;
}

function flattenShapes(slide: Slide): Shape[] {
  const out: Shape[] = [];
  walkShapes(slide.shapes, (s) => out.push(s));
  return out;
}

function walkShapes(shapes: ReadonlyArray<Shape>, visit: (s: Shape) => void): void {
  for (const s of shapes) {
    visit(s);
    if (s.kind === "group") walkShapes(s.children, visit);
  }
}

function bumpShapeKind(
  counts: { text: number; pic: number; group: number; opaque: number },
  s: Shape
): void {
  switch (s.kind) {
    case "text":
      counts.text++;
      return;
    case "pic":
      counts.pic++;
      return;
    case "group":
      counts.group++;
      return;
    case "opaque":
      counts.opaque++;
      return;
    default: {
      const _exhaustive: never = s;
      void _exhaustive;
    }
  }
}

function renderMarkdown(snap: PptxSnapshot, range?: { startSlide: number; endSlide: number }): string {
  if (!range) return snapshotToMarkdown(snap);
  const lo = Math.max(0, range.startSlide);
  const hi = Math.min(snap.root.slides.length, range.endSlide);
  const lines: string[] = [];
  lines.push("# Presentation");
  lines.push("");
  for (let i = lo; i < hi; i++) {
    const s = snap.root.slides[i];
    lines.push(`## Slide ${i + 1} — \`${s.partPath}\` (slideId=${s.slideId})`);
    lines.push("");
    for (const sh of s.shapes) {
      if (sh.kind === "text") lines.push(`> ${textShapePlainText(sh).replaceAll("\n", " · ")}`);
    }
    lines.push("");
  }
  return lines.join("\n");
}

function renderPlainText(snap: PptxSnapshot, range?: { startSlide: number; endSlide: number }): string {
  const lo = range ? Math.max(0, range.startSlide) : 0;
  const hi = range ? Math.min(snap.root.slides.length, range.endSlide) : snap.root.slides.length;
  const lines: string[] = [];
  for (let i = lo; i < hi; i++) {
    const s = snap.root.slides[i];
    walkShapes(s.shapes, (sh) => {
      if (sh.kind === "text") {
        const t = textShapePlainText(sh);
        if (t.length > 0) lines.push(t);
      }
    });
  }
  return lines.join("\n");
}

function sliceRange(
  snap: PptxSnapshot,
  slide: number | undefined
): { startSlide: number; endSlide: number } | undefined {
  if (slide === undefined) return undefined;
  if (slide < 0 || slide >= snap.root.slides.length) {
    throw new Error(
      `--slide ${slide} out of range (presentation has ${snap.root.slides.length} slides)`
    );
  }
  return { startSlide: slide, endSlide: slide + 1 };
}

function inferMime(path: string): string | null {
  const m = /\.([a-zA-Z0-9]+)$/.exec(path);
  if (!m) return null;
  return MIME_BY_EXT[m[1].toLowerCase()] ?? null;
}

function parseIntArg(v: string): number {
  const n = Number(v);
  if (!Number.isFinite(n)) throw new Error(`expected a number, got "${v}"`);
  return Math.trunc(n);
}

function stringifyJson(value: unknown, pretty: boolean): string {
  return pretty ? JSON.stringify(value, null, 2) : JSON.stringify(value);
}

export function normalizeCommands(
  data: unknown
): Array<{ type: string; payload: unknown; source?: "human" | "agent" | "system"; agentId?: string }> {
  const list = Array.isArray(data)
    ? data
    : isObj(data) && Array.isArray((data as { commands?: unknown[] }).commands)
      ? (data as { commands: unknown[] }).commands
      : [data];
  return list.map((c) => {
    if (!isObj(c) || typeof c.type !== "string") {
      throw new Error("each command must be an object with a string `type`");
    }
    return {
      type: c.type as string,
      payload: c.payload,
      ...(c.source ? { source: c.source as "human" | "agent" | "system" } : { source: "agent" as const }),
      ...(typeof c.agentId === "string" ? { agentId: c.agentId } : { agentId: "office-agent-cli" }),
    };
  });
}

function isObj(v: unknown): v is Record<string, unknown> {
  return typeof v === "object" && v !== null;
}

async function dispatchAndWrite(
  opts: { file: string; out?: string; pretty: boolean },
  io: IO,
  commands: ReadonlyArray<BusCommand | CommandLite>
): Promise<void> {
  const agent = await loadAgent(opts.file);
  const muts = await agent.applyCommands(commands);
  agent.getPendingMutations().forEach((m) => agent.approveMutation(m.id));
  const target = opts.out ?? opts.file;
  await writeFile(resolve(target), Buffer.from(await agent.exportFile()));
  io.stdout.write(
    stringifyJson(
      {
        wrote: target,
        mutations: muts.map((m) => ({
          id: m.id,
          type: m.command.type,
          status: m.status,
          ...(m.rejection ? { rejection: m.rejection } : {}),
        })),
      },
      opts.pretty
    ) + "\n"
  );
}

