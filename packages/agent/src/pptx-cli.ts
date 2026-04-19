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
import type { Command as BusCommand, CommandLite, Mutation } from "@officeai/core";
import {
  CliError,
  isEmuUnit,
  readStdinToString,
  toEmu,
  useDeterministicIds,
  type EmuUnit,
} from "./cli-shared.js";

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
    .command("create")
    .description("Create a brand-new blank .pptx file at --out (one empty slide on a default Blank layout).")
    .requiredOption("--out <path>", "Path to write the new .pptx file")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(async (opts: { out: string; pretty: boolean }) => {
      const agent = await PptxAgent.empty();
      await writeFile(resolve(opts.out), Buffer.from(await agent.exportFile()));
      io.stdout.write(stringifyJson({ wrote: opts.out, format: "pptx" }, opts.pretty) + "\n");
    });

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
      new Option("--format <fmt>", "Output format").choices(["markdown", "json", "text"]).default("markdown")
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
            io.stdout.write(stringifyJson(snapshotToJsonProjection(snap, range), opts.pretty) + "\n");
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
    .action(async (opts: { file: string; at?: number; layout?: string; out?: string; pretty: boolean }) => {
      await dispatchAndWrite(opts, io, [
        {
          type: "pptx:add-slide",
          payload: {
            ...(opts.at !== undefined ? { at: opts.at } : {}),
            ...(opts.layout ? { layoutPartPath: opts.layout } : {}),
          },
        },
      ]);
    });

  pptx
    .command("delete-slide")
    .description("Delete a slide by 0-based index and write the result.")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(async (opts: { file: string; slide: number; out?: string; pretty: boolean }) => {
      await dispatchAndWrite(opts, io, [{ type: "pptx:delete-slide", payload: { slideIndex: opts.slide } }]);
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
    .action(async (opts: { file: string; from: number; to: number; out?: string; pretty: boolean }) => {
      await dispatchAndWrite(opts, io, [
        { type: "pptx:move-slide", payload: { from: opts.from, to: opts.to } },
      ]);
    });

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
    .command("set-title")
    .description(
      "Replace the title placeholder text on a slide. Looks up the first text shape with placeholder.type ∈ {title, ctrTitle}."
    )
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .requiredOption("--text <text>", "Replacement title text")
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(async (opts: { file: string; slide: number; text: string; out?: string; pretty: boolean }) => {
      const agent = await loadAgent(opts.file);
      const shapeId = findPlaceholderId(agent.getSnapshot(), opts.slide, ["title", "ctrTitle"]);
      if (!shapeId) {
        throw new CliError(
          65,
          `pptx set-title: no title/ctrTitle placeholder on slide ${opts.slide}. Decks created via 'pptx create' or 'add-slide' have no placeholders; use 'pptx add-text-box' for new titles, or 'pptx read --format json' to discover shape ids and 'pptx set-text --shape-id …' to overwrite an existing free shape.`
        );
      }
      await dispatchAndWrite(opts, io, [
        {
          type: "pptx:set-text",
          payload: { slideIndex: opts.slide, shapeId, text: opts.text },
        },
      ]);
    });

  pptx
    .command("set-body")
    .description(
      "Replace the body placeholder text on a slide. Looks up the first text shape with placeholder.type ∈ {body, subTitle}; \\n separates paragraphs/bullets."
    )
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .requiredOption("--text <text>", "Replacement body text (\\n separates paragraphs)")
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(async (opts: { file: string; slide: number; text: string; out?: string; pretty: boolean }) => {
      const agent = await loadAgent(opts.file);
      const shapeId = findPlaceholderId(agent.getSnapshot(), opts.slide, ["body", "subTitle"]);
      if (!shapeId) {
        throw new CliError(
          65,
          `pptx set-body: no body/subTitle placeholder on slide ${opts.slide}. Decks created via 'pptx create' or 'add-slide' have no placeholders; use 'pptx add-text-box' for new bodies, or 'pptx read --format json' to discover shape ids and 'pptx set-text --shape-id …' to overwrite an existing free shape.`
        );
      }
      await dispatchAndWrite(opts, io, [
        {
          type: "pptx:set-text",
          payload: { slideIndex: opts.slide, shapeId, text: opts.text },
        },
      ]);
    });

  pptx
    .command("set-position")
    .description(
      "Set a shape's <a:off> position. Pass --x/--y in --unit (default emu); 1in=914400emu, 1px=9525emu, 1pt=12700emu, 1cm=360000emu."
    )
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .requiredOption("--shape <id>", "Shape NodeId")
    .requiredOption("--x <n>", "X position in --unit", parseFloatArg)
    .requiredOption("--y <n>", "Y position in --unit", parseFloatArg)
    .addOption(
      new Option("--unit <unit>", "Unit for --x/--y").choices(["emu", "px", "in", "cm", "pt"]).default("emu")
    )
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        slide: number;
        shape: string;
        x: number;
        y: number;
        unit: string;
        out?: string;
        pretty: boolean;
      }) => {
        const unit = parseEmuUnit(opts.unit);
        await dispatchAndWrite(opts, io, [
          {
            type: "pptx:set-position",
            payload: {
              slideIndex: opts.slide,
              shapeId: opts.shape,
              x: toEmu(opts.x, unit),
              y: toEmu(opts.y, unit),
            },
          },
        ]);
      }
    );

  pptx
    .command("set-size")
    .description("Set a shape's <a:ext> width/height. Pass --width/--height in --unit (default emu).")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .requiredOption("--shape <id>", "Shape NodeId")
    .requiredOption("--width <n>", "Width in --unit (>0)", parseFloatArg)
    .requiredOption("--height <n>", "Height in --unit (>0)", parseFloatArg)
    .addOption(
      new Option("--unit <unit>", "Unit for --width/--height")
        .choices(["emu", "px", "in", "cm", "pt"])
        .default("emu")
    )
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        slide: number;
        shape: string;
        width: number;
        height: number;
        unit: string;
        out?: string;
        pretty: boolean;
      }) => {
        const unit = parseEmuUnit(opts.unit);
        await dispatchAndWrite(opts, io, [
          {
            type: "pptx:set-size",
            payload: {
              slideIndex: opts.slide,
              shapeId: opts.shape,
              width: toEmu(opts.width, unit),
              height: toEmu(opts.height, unit),
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
    .command("set-paragraph-alignment")
    .description(
      "Set (or clear with --clear) the horizontal alignment on one or every paragraph of a text shape."
    )
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .requiredOption("--shape <id>", "Shape NodeId")
    .addOption(
      new Option("--alignment <value>", "Horizontal alignment to apply").choices([
        "left",
        "center",
        "right",
        "justify",
      ])
    )
    .option(
      "--paragraph <n>",
      "0-based paragraph index inside the shape (repeat to target many; omit to apply shape-wide)",
      collectInts,
      [] as number[]
    )
    .option("--clear", "Clear any existing alignment override on the targeted paragraphs", false)
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        slide: number;
        shape: string;
        alignment?: string;
        paragraph: number[];
        clear: boolean;
        out?: string;
        pretty: boolean;
      }) => {
        if (!opts.clear && !opts.alignment) {
          throw new CliError(
            64,
            "set-paragraph-alignment: pass --alignment <left|center|right|justify> or --clear"
          );
        }
        if (opts.clear && opts.alignment) {
          throw new CliError(64, "set-paragraph-alignment: --alignment and --clear are mutually exclusive");
        }
        const alignment = opts.clear ? null : (opts.alignment as "left" | "center" | "right" | "justify");
        const payload: Record<string, unknown> = {
          slideIndex: opts.slide,
          shapeId: opts.shape,
          alignment,
        };
        if (opts.paragraph.length > 0) payload.paragraphs = opts.paragraph;
        await dispatchAndWrite(opts, io, [
          { type: "pptx:set-paragraph-alignment", payload },
        ]);
      }
    );

  pptx
    .command("set-text-anchor")
    .description("Set (or clear with --clear) the vertical text anchor on a text shape.")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .requiredOption("--shape <id>", "Shape NodeId")
    .addOption(
      new Option("--anchor <value>", "Vertical anchor to apply").choices(["top", "middle", "bottom"])
    )
    .option("--clear", "Clear any existing anchor override on the shape", false)
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        slide: number;
        shape: string;
        anchor?: string;
        clear: boolean;
        out?: string;
        pretty: boolean;
      }) => {
        if (!opts.clear && !opts.anchor) {
          throw new CliError(64, "set-text-anchor: pass --anchor <top|middle|bottom> or --clear");
        }
        if (opts.clear && opts.anchor) {
          throw new CliError(64, "set-text-anchor: --anchor and --clear are mutually exclusive");
        }
        const anchor = opts.clear ? null : (opts.anchor as "top" | "middle" | "bottom");
        await dispatchAndWrite(opts, io, [
          {
            type: "pptx:set-text-anchor",
            payload: { slideIndex: opts.slide, shapeId: opts.shape, anchor },
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
          throw new Error(`Could not infer MIME from "${opts.image}". Pass --mime image/png (or similar).`);
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
    .description("Append a fresh text box on a slide. Coordinates are in --unit (default emu).")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .requiredOption("--text <text>", "Initial text content")
    .requiredOption("--x <n>", "X in --unit", parseFloatArg)
    .requiredOption("--y <n>", "Y in --unit", parseFloatArg)
    .requiredOption("--width <n>", "Width in --unit (>0)", parseFloatArg)
    .requiredOption("--height <n>", "Height in --unit (>0)", parseFloatArg)
    .addOption(
      new Option("--unit <unit>", "Unit for --x/--y/--width/--height")
        .choices(["emu", "px", "in", "cm", "pt"])
        .default("emu")
    )
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
        unit: string;
        name?: string;
        out?: string;
        pretty: boolean;
      }) => {
        const unit = parseEmuUnit(opts.unit);
        await dispatchAndWrite(opts, io, [
          {
            type: "pptx:add-text-box",
            payload: {
              slideIndex: opts.slide,
              text: opts.text,
              x: toEmu(opts.x, unit),
              y: toEmu(opts.y, unit),
              width: toEmu(opts.width, unit),
              height: toEmu(opts.height, unit),
              ...(opts.name ? { name: opts.name } : {}),
            },
          },
        ]);
      }
    );

  // ── Tables (F2) ─────────────────────────────────────────────────────
  pptx
    .command("table-set-cell-text")
    .description("Set the text of a single cell in a typed TableShape.")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .requiredOption("--shape <id>", "Table shape NodeId")
    .requiredOption("--row <n>", "0-based row index", parseIntArg)
    .requiredOption("--column <n>", "0-based column index", parseIntArg)
    .requiredOption("--text <text>", "Replacement text (\\n splits paragraphs)")
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        slide: number;
        shape: string;
        row: number;
        column: number;
        text: string;
        out?: string;
        pretty: boolean;
      }) => {
        await dispatchAndWrite(opts, io, [
          {
            type: "pptx:table-set-cell-text",
            payload: {
              slideIndex: opts.slide,
              shapeId: opts.shape,
              row: opts.row,
              column: opts.column,
              text: opts.text,
            },
          },
        ]);
      }
    );

  pptx
    .command("table-add-row")
    .description("Insert a new row into a typed TableShape.")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .requiredOption("--shape <id>", "Table shape NodeId")
    .addOption(new Option("--at <n>", "0-based insert position (default: append)").argParser(parseIntArg))
    .addOption(
      new Option("--height <emu>", "Row height in EMU (default: median of existing rows)").argParser(
        parseIntArg
      )
    )
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        slide: number;
        shape: string;
        at?: number;
        height?: number;
        out?: string;
        pretty: boolean;
      }) => {
        await dispatchAndWrite(opts, io, [
          {
            type: "pptx:table-add-row",
            payload: {
              slideIndex: opts.slide,
              shapeId: opts.shape,
              ...(opts.at !== undefined ? { at: opts.at } : {}),
              ...(opts.height !== undefined ? { height: opts.height } : {}),
            },
          },
        ]);
      }
    );

  pptx
    .command("table-delete-row")
    .description("Delete a row from a typed TableShape.")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .requiredOption("--shape <id>", "Table shape NodeId")
    .requiredOption("--row <n>", "0-based row index", parseIntArg)
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        slide: number;
        shape: string;
        row: number;
        out?: string;
        pretty: boolean;
      }) => {
        await dispatchAndWrite(opts, io, [
          {
            type: "pptx:table-delete-row",
            payload: { slideIndex: opts.slide, shapeId: opts.shape, row: opts.row },
          },
        ]);
      }
    );

  pptx
    .command("table-add-column")
    .description("Insert a new column into a typed TableShape.")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .requiredOption("--shape <id>", "Table shape NodeId")
    .addOption(new Option("--at <n>", "0-based insert position (default: append)").argParser(parseIntArg))
    .addOption(
      new Option("--width <emu>", "Column width in EMU (default: average of existing columns)").argParser(
        parseIntArg
      )
    )
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        slide: number;
        shape: string;
        at?: number;
        width?: number;
        out?: string;
        pretty: boolean;
      }) => {
        await dispatchAndWrite(opts, io, [
          {
            type: "pptx:table-add-column",
            payload: {
              slideIndex: opts.slide,
              shapeId: opts.shape,
              ...(opts.at !== undefined ? { at: opts.at } : {}),
              ...(opts.width !== undefined ? { width: opts.width } : {}),
            },
          },
        ]);
      }
    );

  pptx
    .command("table-delete-column")
    .description("Delete a column from a typed TableShape.")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .requiredOption("--shape <id>", "Table shape NodeId")
    .requiredOption("--column <n>", "0-based column index", parseIntArg)
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        slide: number;
        shape: string;
        column: number;
        out?: string;
        pretty: boolean;
      }) => {
        await dispatchAndWrite(opts, io, [
          {
            type: "pptx:table-delete-column",
            payload: { slideIndex: opts.slide, shapeId: opts.shape, column: opts.column },
          },
        ]);
      }
    );

  // ── Charts (F3) ─────────────────────────────────────────────────────
  pptx
    .command("chart-set-title")
    .description("Set or remove a typed ChartShape's title text. Pass --remove to clear it.")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .requiredOption("--shape <id>", "Chart shape NodeId")
    .option("--title <text>", "Replacement title text")
    .option("--remove", "Remove the title", false)
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        slide: number;
        shape: string;
        title?: string;
        remove?: boolean;
        out?: string;
        pretty: boolean;
      }) => {
        if (!opts.remove && opts.title === undefined) {
          throw new Error("Pass either --title <text> or --remove");
        }
        await dispatchAndWrite(opts, io, [
          {
            type: "pptx:set-chart-title",
            payload: {
              slideIndex: opts.slide,
              shapeId: opts.shape,
              title: opts.remove ? null : (opts.title ?? null),
            },
          },
        ]);
      }
    );

  pptx
    .command("chart-set-type")
    .description("Change a typed ChartShape's chart type (bar|line|area|pie).")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .requiredOption("--shape <id>", "Chart shape NodeId")
    .addOption(
      new Option("--type <kind>", "Target chart type")
        .choices(["bar", "line", "area", "pie"])
        .makeOptionMandatory(true)
    )
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        slide: number;
        shape: string;
        type: "bar" | "line" | "area" | "pie";
        out?: string;
        pretty: boolean;
      }) => {
        await dispatchAndWrite(opts, io, [
          {
            type: "pptx:set-chart-type",
            payload: {
              slideIndex: opts.slide,
              shapeId: opts.shape,
              chartType: opts.type,
            },
          },
        ]);
      }
    );

  pptx
    .command("chart-set-data")
    .description("Replace a typed ChartShape's categories + series data from a JSON file.")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .requiredOption("--shape <id>", "Chart shape NodeId")
    .requiredOption(
      "--data <path>",
      'JSON file: { "categories": ["Q1",…], "series": [{ "name": "Revenue", "values": [1,2,3] }, …] }'
    )
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        slide: number;
        shape: string;
        data: string;
        out?: string;
        pretty: boolean;
      }) => {
        const raw = await readFile(resolve(opts.data), "utf8");
        const parsed: unknown = JSON.parse(raw);
        const { categories, series } = parseChartData(parsed);
        await dispatchAndWrite(opts, io, [
          {
            type: "pptx:set-chart-data",
            payload: {
              slideIndex: opts.slide,
              shapeId: opts.shape,
              categories,
              series,
            },
          },
        ]);
      }
    );

  // ── Animations (F4) ─────────────────────────────────────────────────
  pptx
    .command("set-slide-transition")
    .description("Replace a slide's typed transition. Pass --kind none to remove it.")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .addOption(
      new Option("--kind <kind>", "Transition kind")
        .choices(["none", "fade", "push", "wipe", "split", "cut"])
        .makeOptionMandatory(true)
    )
    .addOption(new Option("--speed <speed>", "Transition speed").choices(["slow", "med", "fast"]))
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        slide: number;
        kind: "none" | "fade" | "push" | "wipe" | "split" | "cut";
        speed?: "slow" | "med" | "fast";
        out?: string;
        pretty: boolean;
      }) => {
        await dispatchAndWrite(opts, io, [
          {
            type: "pptx:set-slide-transition",
            payload: {
              slideIndex: opts.slide,
              kind: opts.kind,
              ...(opts.speed ? { speed: opts.speed } : {}),
            },
          },
        ]);
      }
    );

  pptx
    .command("add-shape-animation")
    .description("Append (or insert) a typed entrance animation on a shape.")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .requiredOption("--shape <id>", "Shape NodeId to animate")
    .addOption(
      new Option("--effect <effect>", "Entrance effect")
        .choices(["appear", "fade", "fly-in", "wipe"])
        .makeOptionMandatory(true)
    )
    .option("--at <n>", "Insert position in the entrance sequence", parseIntArg)
    .option("--duration-ms <n>", "Effect duration in milliseconds", parseIntArg)
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        slide: number;
        shape: string;
        effect: "appear" | "fade" | "fly-in" | "wipe";
        at?: number;
        durationMs?: number;
        out?: string;
        pretty: boolean;
      }) => {
        await dispatchAndWrite(opts, io, [
          {
            type: "pptx:add-shape-animation",
            payload: {
              slideIndex: opts.slide,
              shapeId: opts.shape,
              effect: opts.effect,
              ...(opts.at !== undefined ? { at: opts.at } : {}),
              ...(opts.durationMs !== undefined ? { durationMs: opts.durationMs } : {}),
            },
          },
        ]);
      }
    );

  pptx
    .command("remove-shape-animation")
    .description("Remove a typed entrance animation by its animation NodeId.")
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .requiredOption("--animation <id>", "Animation NodeId returned by inspect/read")
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: { file: string; slide: number; animation: string; out?: string; pretty: boolean }) => {
        await dispatchAndWrite(opts, io, [
          {
            type: "pptx:remove-shape-animation",
            payload: { slideIndex: opts.slide, animationId: opts.animation },
          },
        ]);
      }
    );

  pptx
    .command("reorder-shape-animations")
    .description(
      "Atomically reorder a slide's typed animations. --order accepts a comma-separated list of NodeIds."
    )
    .requiredOption("--file <path>", "Path to a .pptx file")
    .requiredOption("--slide <n>", "0-based slide index", parseIntArg)
    .requiredOption("--order <ids>", "Comma-separated NodeIds in the new order")
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(async (opts: { file: string; slide: number; order: string; out?: string; pretty: boolean }) => {
      const order = opts.order
        .split(",")
        .map((s) => s.trim())
        .filter((s) => s.length > 0);
      if (order.length === 0) throw new Error("--order requires at least one id");
      await dispatchAndWrite(opts, io, [
        {
          type: "pptx:reorder-shape-animations",
          payload: { slideIndex: opts.slide, order },
        },
      ]);
    });

  pptx
    .command("apply")
    .description(
      "Apply a JSON command file (single command or { commands: [...] }) and write the result. Pass -c/--commands to read from disk or --from-stdin to read JSON piped on stdin (mutually exclusive)."
    )
    .requiredOption("--file <path>", "Path to a .pptx file")
    .option("-c, --commands <path>", "Path to a JSON file containing one or more commands")
    .option("--from-stdin", "Read the JSON command body from stdin instead of -c <path>", false)
    .option("--out <path>", "Path to write the resulting .pptx file (defaults to --file)")
    .option("--pretty", "Pretty-print JSON output", false)
    .action(
      async (opts: {
        file: string;
        commands?: string;
        fromStdin: boolean;
        out?: string;
        pretty: boolean;
      }) => {
        if (opts.fromStdin === Boolean(opts.commands)) {
          throw new CliError(64, "pptx apply: pass exactly one of -c/--commands <path> or --from-stdin");
        }
        const raw = opts.fromStdin
          ? await readStdinToString()
          : await readFile(resolve(opts.commands as string), "utf8");
        const data: unknown = JSON.parse(raw);
        const cmds = normalizeCommands(data);
        await dispatchAndWrite(opts, io, cmds);
      }
    );

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
  // Tests + scripted CLI flows need stable shape NodeIds across
  // multiple invocations against the same fixture. The default UUID
  // minter changes them on every parse, which makes ID-based commands
  // (set-text/format-text/set-position/set-size) fail across runs.
  // Toggle via the root `--deterministic-ids` flag (preferred) or the
  // legacy `OFFICEAI_DETERMINISTIC_IDS=1` env var.
  if (useDeterministicIds()) {
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
  shapeCounts: {
    text: number;
    pic: number;
    group: number;
    table: number;
    chart: number;
    connector: number;
    opaque: number;
  };
  /** F4: total typed entrance animations across all slides. */
  animations: number;
  /** F4: number of slides carrying a typed transition. */
  transitions: number;
}

export function inspectSnapshot(snap: PptxSnapshot): PptxSnapshotSummary {
  const counts = { text: 0, pic: 0, group: 0, table: 0, chart: 0, connector: 0, opaque: 0 };
  for (const slide of snap.root.slides) walkShapes(slide.shapes, (s) => bumpShapeKind(counts, s));
  let animations = 0;
  let transitions = 0;
  for (const slide of snap.root.slides) {
    animations += slide.animations.length;
    if (slide.transition) transitions++;
  }
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
    animations,
    transitions,
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
    transition?: { kind: string; speed?: string };
    animations?: ReadonlyArray<{
      id: string;
      effect: string;
      targetCNvPrId: number;
      durationMs?: number;
      order: number;
    }>;
    shapes: ReadonlyArray<{
      id: string;
      kind: Shape["kind"];
      cNvPrId?: number;
      name?: string;
      bbox?: { x: number; y: number; cx: number; cy: number };
      text?: string;
      table?: {
        rows: number;
        columns: number;
        cells: ReadonlyArray<ReadonlyArray<string>>;
      };
      chart?: {
        partPath: string;
        chartType: string;
        title?: string;
        categories: ReadonlyArray<string>;
        series: ReadonlyArray<{ name?: string; values: ReadonlyArray<number> }>;
      };
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
      ...(s.transition
        ? {
            transition: {
              kind: s.transition.kind,
              ...(s.transition.speed ? { speed: s.transition.speed } : {}),
            },
          }
        : {}),
      ...(s.animations.length > 0
        ? {
            animations: s.animations.map((a) => ({
              id: a.id,
              effect: a.effect,
              targetCNvPrId: a.targetCNvPrId,
              ...(a.durationMs !== undefined ? { durationMs: a.durationMs } : {}),
              order: a.order,
            })),
          }
        : {}),
      shapes: s.shapes.map((sh) => projectShape(sh, snap)),
    });
  }
  return { format: "pptx", revision: snap.revision, slides: out };
}

function projectShape(
  sh: Shape,
  snap: PptxSnapshot
): PptxSnapshotProjection["slides"][number]["shapes"][number] {
  const base: {
    id: string;
    kind: Shape["kind"];
    cNvPrId?: number;
    name?: string;
    bbox?: { x: number; y: number; cx: number; cy: number };
    text?: string;
    table?: { rows: number; columns: number; cells: ReadonlyArray<ReadonlyArray<string>> };
    chart?: {
      partPath: string;
      chartType: string;
      title?: string;
      categories: ReadonlyArray<string>;
      series: ReadonlyArray<{ name?: string; values: ReadonlyArray<number> }>;
    };
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
  if (sh.kind === "table") {
    base.table = {
      rows: sh.rows.length,
      columns: sh.columnWidths.length,
      cells: sh.rows.map((row) =>
        row.cells.map((cell) =>
          cell.txBody.paragraphs
            .map((p) => p.runs.map((r) => (r.isLineBreak ? "\n" : r.text)).join(""))
            .join("\n")
        )
      ),
    };
  }
  if (sh.kind === "chart") {
    const part = snap.root.charts.get(sh.chartPartPath);
    base.chart = {
      partPath: sh.chartPartPath,
      chartType: part?.chartType ?? "unsupported",
      ...(part?.title !== undefined ? { title: part.title } : {}),
      categories: part?.categories ?? [],
      series: (part?.series ?? []).map((s) => ({
        ...(s.name !== undefined ? { name: s.name } : {}),
        values: s.values,
      })),
    };
  }
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
  const bbb =
    b.position && b.size ? `${b.position.xEmu},${b.position.yEmu},${b.size.cxEmu},${b.size.cyEmu}` : "";
  const abb =
    a.position && a.size ? `${a.position.xEmu},${a.position.yEmu},${a.size.cxEmu},${a.size.cyEmu}` : "";
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
  counts: {
    text: number;
    pic: number;
    group: number;
    table: number;
    chart: number;
    connector: number;
    opaque: number;
  },
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
    case "table":
      counts.table++;
      return;
    case "chart":
      counts.chart++;
      return;
    case "connector":
      counts.connector++;
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
    throw new Error(`--slide ${slide} out of range (presentation has ${snap.root.slides.length} slides)`);
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

function parseFloatArg(v: string): number {
  const n = Number(v);
  if (!Number.isFinite(n)) throw new Error(`expected a number, got "${v}"`);
  return n;
}

/**
 * Commander variadic-option collector for repeated `--paragraph <n>`
 * flags. Each invocation appends one parsed integer to the running
 * list, so callers can write `--paragraph 0 --paragraph 2` to target
 * a paragraph subset.
 */
function collectInts(value: string, previous: number[]): number[] {
  return [...previous, parseIntArg(value)];
}

/**
 * Walk a slide's shapes (recursing into groups) and return the first
 * `TextShape` whose `placeholder.type` matches one of `wantedTypes`.
 * Used by `pptx set-title` / `pptx set-body` so agents don't have to
 * round-trip through `pptx inspect` to discover the placeholder
 * shape id.
 */
function findPlaceholderId(
  snap: PptxSnapshot,
  slideIndex: number,
  wantedTypes: ReadonlyArray<string>
): string | undefined {
  const slide = snap.root.slides[slideIndex];
  if (!slide) return undefined;
  let found: string | undefined;
  walkShapes(slide.shapes, (s) => {
    if (found) return;
    if (s.kind !== "text") return;
    const ph = (s as TextShape).placeholder;
    if (!ph?.type) return;
    if (wantedTypes.includes(ph.type)) found = s.id;
  });
  return found;
}

function parseEmuUnit(unit: string): EmuUnit {
  if (!isEmuUnit(unit)) {
    throw new CliError(64, `--unit must be one of emu|px|in|cm|pt; got "${unit}"`);
  }
  return unit;
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

function parseChartData(v: unknown): {
  categories: ReadonlyArray<string>;
  series: ReadonlyArray<{ name?: string; values: ReadonlyArray<number> }>;
} {
  if (!isObj(v)) throw new Error("chart data must be a JSON object");
  const cats = (v as { categories?: unknown }).categories;
  const sers = (v as { series?: unknown }).series;
  if (!Array.isArray(cats) || cats.some((c) => typeof c !== "string")) {
    throw new Error("chart data must have a `categories: string[]` field");
  }
  if (!Array.isArray(sers)) {
    throw new Error("chart data must have a `series: { name?, values: number[] }[]` field");
  }
  const series = sers.map((s, i) => {
    if (!isObj(s)) throw new Error(`series[${i}] must be an object`);
    const name = (s as { name?: unknown }).name;
    const values = (s as { values?: unknown }).values;
    if (!Array.isArray(values) || values.some((n) => typeof n !== "number")) {
      throw new Error(`series[${i}].values must be number[]`);
    }
    if (values.length !== cats.length) {
      throw new Error(
        `series[${i}].values length (${values.length}) must match categories length (${cats.length})`
      );
    }
    return {
      ...(typeof name === "string" ? { name } : {}),
      values: values as number[],
    };
  });
  return { categories: cats as string[], series };
}

async function dispatchAndWrite(
  opts: { file: string; out?: string; pretty: boolean },
  io: IO,
  commands: ReadonlyArray<BusCommand | CommandLite>
): Promise<void> {
  const agent = await loadAgent(opts.file);
  const muts = await agent.applyCommands(commands);
  const ids = agent.getPendingMutations().map((m) => m.id);
  for (const id of ids) agent.approveMutation(id);
  const target = opts.out ?? opts.file;
  await writeFile(resolve(target), Buffer.from(await agent.exportFile()));
  const post = agent.getSnapshot();
  io.stdout.write(
    stringifyJson(
      {
        wrote: target,
        mutations: muts.map((m) => pptxMutationSummary(m, post, true)),
      },
      opts.pretty
    ) + "\n"
  );
  const rejected = muts.filter((m) => m.status === "rejected");
  if (rejected.length > 0) {
    throw new CliError(
      2,
      `pptx: ${rejected.length}/${muts.length} mutation(s) rejected; first failure: ${rejected[0].command.type} → ${rejected[0].rejection?.code ?? "unknown"} (${rejected[0].rejection?.message ?? "no message"})`
    );
  }
}

/**
 * Per-mutation envelope used by every PPTX write command. Mirrors
 * `mutationLineSummary` in cli.ts: surfaces freshly minted node ids
 * and added OPC parts so callers can chain commands without an
 * intermediate `pptx inspect` round-trip. Special-cases
 * `pptx:add-slide` / `pptx:duplicate-slide`: the diff only carries
 * the slide id, so we read the post-dispatch snapshot to enumerate
 * the slide's placeholder shapes (id + cNvPrId + name + kind) so
 * downstream calls can target them directly.
 */
function pptxMutationSummary(
  m: Mutation<PptxSnapshot>,
  post: PptxSnapshot,
  bulkApproved: boolean
): {
  id: string;
  type: string;
  status: string;
  inserted?: ReadonlyArray<{ nodeId: string; path: string }>;
  addedParts?: ReadonlyArray<string>;
  shapes?: ReadonlyArray<{ id: string; kind: string; cNvPrId?: number; name?: string }>;
  rejection?: { code: string; message: string };
} {
  const inserted: Array<{ nodeId: string; path: string }> = [];
  const addedParts: string[] = [];
  for (const c of m.diff.changes) {
    if (c.kind === "node-inserted") {
      inserted.push({ nodeId: c.nodeId, path: c.path.join("/") });
    } else if (c.kind === "part-added") {
      addedParts.push(c.path.join("/"));
    }
  }
  let shapes: Array<{ id: string; kind: string; cNvPrId?: number; name?: string }> | undefined;
  if (m.command.type === "pptx:add-slide" || m.command.type === "pptx:duplicate-slide") {
    const newSlide = post.root.slides.find(
      (s) => s.id === m.diff.changes.find((c) => c.kind === "node-inserted")?.nodeId
    );
    if (newSlide) {
      shapes = [];
      walkShapes(newSlide.shapes, (s) => {
        shapes!.push({
          id: s.id,
          kind: s.kind,
          ...(s.cNvPrId !== undefined ? { cNvPrId: s.cNvPrId } : {}),
          ...(s.name ? { name: s.name } : {}),
        });
      });
    }
  }
  const status = m.status === "rejected" ? "rejected" : bulkApproved ? "approved" : m.status;
  return {
    id: m.id,
    type: m.command.type,
    status,
    ...(inserted.length > 0 ? { inserted } : {}),
    ...(addedParts.length > 0 ? { addedParts } : {}),
    ...(shapes && shapes.length > 0 ? { shapes } : {}),
    ...(m.rejection ? { rejection: m.rejection } : {}),
  };
}
