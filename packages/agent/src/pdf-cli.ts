/**
 * `office-agent pdf …` subcommand group.
 *
 * Mirrors the docx / xlsx / pptx subcommand surfaces, but built on the
 * pdf-* package family:
 *
 *   • read-* commands load a `PdfAgent`, project its snapshot, and emit
 *     a versioned JSON envelope to stdout
 *     (`{ "schema": "office-agent/pdf-<verb>@1", … }`).
 *   • mutate commands run the relevant pdf-edit / pdf-annotations /
 *     pdf-forms / pdf-ocr helper directly on the input bytes (the
 *     PdfAgent's command bus only covers the page-rotation / reorder
 *     subset today; the page-level mutators live in pdf-edit) and
 *     write the result to `--out`. They emit a single JSON summary on
 *     stdout: `{ schema, in, out, bytes, summary }`.
 *
 * Error envelope: every failure is caught by `pdfAction` which writes
 * `{ "error": "<code>", "message": "<msg>" }` to stderr and returns
 * exit code 1 (via a silent `CliError`). The intent is that LLM
 * callers can `JSON.parse(stderr)` on non-zero exit to recover a
 * structured failure.
 */

import { mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import type { Command as CommanderCommand } from "commander";
import { PdfAgent } from "@officeai/pdf";
import type { PdfAnnotation, PdfFormField, PdfOutlineNode, PdfSnapshot } from "@officeai/pdf";
import {
  addPageNumbers,
  addWatermark,
  cropPages,
  deletePages,
  extractPages,
  mergePdfs,
  reorderPages,
  rotatePages,
  setMetadata,
  splitPdf,
} from "@officeai/pdf-edit";
import { fillForm, flattenForm, listFormFields, resetForm } from "@officeai/pdf-forms";
import type { ListedFormField } from "@officeai/pdf-forms";
import { addTextLayer } from "@officeai/pdf-ocr";
import { CliError, type IO } from "./cli-shared.js";

// ── Public entry point ────────────────────────────────────────────────────

/**
 * Register every `pdf` subcommand under the supplied parent. The parent
 * is typically the `pdf` Command created in cli.ts; tests may pass any
 * Commander `Command` to register the surface in isolation.
 */
export function registerPdfSubcommands(pdf: CommanderCommand, io: IO): void {
  registerReadCommands(pdf, io);
  registerMutateCommands(pdf, io);
}

// ── Helpers (also used by mcp.ts) ────────────────────────────────────────

export type PdfMetadataSummary = {
  readonly schema: "office-agent/pdf-read-metadata@1";
  readonly title?: string;
  readonly author?: string;
  readonly subject?: string;
  readonly keywords?: string;
  readonly creator?: string;
  readonly producer?: string;
  readonly creationDate?: string;
  readonly modificationDate?: string;
  readonly pageCount: number;
  readonly signatureCount: number;
  readonly engine: "pdfjs" | "pdfium";
  readonly pdfVersion?: string;
  readonly linearized: boolean;
  readonly encryption: { hasUserPassword: boolean; hasOwnerPassword: boolean };
};

export function projectMetadata(snapshot: PdfSnapshot): PdfMetadataSummary {
  const md = snapshot.root.metadata;
  return {
    schema: "office-agent/pdf-read-metadata@1",
    ...(md.title !== undefined ? { title: md.title } : {}),
    ...(md.author !== undefined ? { author: md.author } : {}),
    ...(md.subject !== undefined ? { subject: md.subject } : {}),
    ...(md.keywords !== undefined ? { keywords: md.keywords } : {}),
    ...(md.creator !== undefined ? { creator: md.creator } : {}),
    ...(md.producer !== undefined ? { producer: md.producer } : {}),
    ...(md.creationDate !== undefined ? { creationDate: md.creationDate } : {}),
    ...(md.modificationDate !== undefined ? { modificationDate: md.modificationDate } : {}),
    pageCount: snapshot.root.pages.length,
    signatureCount: snapshot.root.signatureCount,
    engine: snapshot.root.engineKind,
    ...(md.pdfVersion !== undefined ? { pdfVersion: md.pdfVersion } : {}),
    linearized: md.linearized === true,
    encryption: md.encryption ?? { hasUserPassword: false, hasOwnerPassword: false },
  };
}

export type PdfPageProjection = {
  readonly schema: "office-agent/pdf-read-page@1";
  readonly page: {
    readonly pageNumber: number;
    readonly id: string;
    readonly width: number;
    readonly height: number;
    readonly rotation: number;
    readonly hasTextLayer: boolean;
    readonly text: string;
  };
  readonly annotations: ReadonlyArray<PdfAnnotation>;
  readonly formFields: ReadonlyArray<PdfFormField>;
};

export function projectPage(snapshot: PdfSnapshot, pageNumber: number): PdfPageProjection {
  const total = snapshot.root.pages.length;
  if (!Number.isInteger(pageNumber) || pageNumber < 1 || pageNumber > total) {
    throw new PdfError(
      "page-out-of-range",
      `--page ${pageNumber} out of range (document has ${total} page${total === 1 ? "" : "s"})`
    );
  }
  const page = snapshot.root.pages[pageNumber - 1];
  return {
    schema: "office-agent/pdf-read-page@1",
    page: {
      pageNumber: page.pageNumber,
      id: page.id,
      width: page.width,
      height: page.height,
      rotation: page.rotation,
      hasTextLayer: page.hasTextLayer,
      text: page.text,
    },
    annotations: snapshot.root.annotations.filter((a) => a.pageNumber === pageNumber),
    formFields: snapshot.root.formFields.filter((f) => f.pageNumber === pageNumber),
  };
}

export type PdfOutlineProjection = {
  readonly schema: "office-agent/pdf-read-outline@1";
  readonly outline: ReadonlyArray<PdfOutlineNode>;
};

export function projectOutline(snapshot: PdfSnapshot): PdfOutlineProjection {
  return {
    schema: "office-agent/pdf-read-outline@1",
    outline: snapshot.root.outline,
  };
}

export type PdfAnnotationsProjection = {
  readonly schema: "office-agent/pdf-read-annotations@1";
  readonly annotations: ReadonlyArray<PdfAnnotation>;
};

export function projectAnnotations(
  snapshot: PdfSnapshot,
  filter?: { readonly page?: number }
): PdfAnnotationsProjection {
  const all = snapshot.root.annotations;
  const annotations = filter?.page !== undefined ? all.filter((a) => a.pageNumber === filter.page) : all;
  return { schema: "office-agent/pdf-read-annotations@1", annotations };
}

export type PdfFormFieldsProjection = {
  readonly schema: "office-agent/pdf-list-form-fields@1";
  readonly fields: ReadonlyArray<ListedFormField & { readonly pageNumber?: number }>;
};

/**
 * Project form fields by joining the pdf-forms reader (canonical
 * value/options) with the snapshot's per-field pageNumber so callers
 * can decide where each widget lives without a second round-trip.
 */
export async function projectFormFields(
  snapshot: PdfSnapshot,
  bytes: Uint8Array
): Promise<PdfFormFieldsProjection> {
  const fromForms = await listFormFields(bytes);
  const pageByName = new Map(snapshot.root.formFields.map((f) => [f.name, f.pageNumber] as const));
  return {
    schema: "office-agent/pdf-list-form-fields@1",
    fields: fromForms.map((f) => ({
      ...f,
      ...(pageByName.has(f.name) ? { pageNumber: pageByName.get(f.name) } : {}),
    })),
  };
}

export type PdfSearchProjection = {
  readonly schema: "office-agent/pdf-search-text@1";
  readonly query: string;
  readonly regex: boolean;
  readonly caseSensitive: boolean;
  readonly results: ReadonlyArray<{
    readonly page: number;
    readonly start: number;
    readonly end: number;
    readonly preview: string;
    readonly match: string;
  }>;
};

export function projectSearch(
  agent: PdfAgent,
  spec: { query: string; regex?: boolean; caseSensitive?: boolean }
): PdfSearchProjection {
  const hits = agent.search({
    query: spec.query,
    regex: spec.regex === true,
    caseSensitive: spec.caseSensitive === true,
  });
  return {
    schema: "office-agent/pdf-search-text@1",
    query: spec.query,
    regex: spec.regex === true,
    caseSensitive: spec.caseSensitive === true,
    results: hits.map((h) => ({
      page: h.pageNumber,
      start: h.start,
      end: h.end,
      preview: h.preview,
      match: h.match,
    })),
  };
}

// ── Read commands ────────────────────────────────────────────────────────

function registerReadCommands(pdf: CommanderCommand, io: IO): void {
  pdf
    .command("read-metadata <file>")
    .description(
      "Print the PDF's metadata, page/signature counts, encryption flags, engine, version, and linearization status as JSON. Schema: office-agent/pdf-read-metadata@1."
    )
    .action(async (file: string) => {
      await pdfAction(io, async () => {
        const bytes = await readPdfBytes(file);
        const agent = await PdfAgent.fromBuffer(bytes);
        writeJson(io, projectMetadata(agent.getSnapshot()));
      });
    });

  pdf
    .command("read-page <file>")
    .description(
      "Print a single page's projection (size, rotation, text, annotations, form fields). Schema: office-agent/pdf-read-page@1."
    )
    .requiredOption("--page <n>", "1-indexed page number to project", parseIntArg)
    .action(async (file: string, opts: { page: number }) => {
      await pdfAction(io, async () => {
        const bytes = await readPdfBytes(file);
        const agent = await PdfAgent.fromBuffer(bytes);
        writeJson(io, projectPage(agent.getSnapshot(), opts.page));
      });
    });

  pdf
    .command("read-outline <file>")
    .description(
      "Print the recursive outline tree (each entry carries its title, optional pageNumber, and children). Schema: office-agent/pdf-read-outline@1."
    )
    .action(async (file: string) => {
      await pdfAction(io, async () => {
        const bytes = await readPdfBytes(file);
        const agent = await PdfAgent.fromBuffer(bytes);
        writeJson(io, projectOutline(agent.getSnapshot()));
      });
    });

  pdf
    .command("read-annotations <file>")
    .description(
      "Print every annotation in the document as a flat JSON array, optionally filtered to one page. Schema: office-agent/pdf-read-annotations@1."
    )
    .option("--page <n>", "Restrict to a single 1-indexed page", parseIntArg)
    .action(async (file: string, opts: { page?: number }) => {
      await pdfAction(io, async () => {
        const bytes = await readPdfBytes(file);
        const agent = await PdfAgent.fromBuffer(bytes);
        writeJson(
          io,
          projectAnnotations(agent.getSnapshot(), opts.page !== undefined ? { page: opts.page } : undefined)
        );
      });
    });

  pdf
    .command("list-form-fields <file>")
    .description(
      "List AcroForm fields with name, type, value, options, readOnly, and required. Schema: office-agent/pdf-list-form-fields@1."
    )
    .action(async (file: string) => {
      await pdfAction(io, async () => {
        const bytes = await readPdfBytes(file);
        const agent = await PdfAgent.fromBuffer(bytes);
        writeJson(io, await projectFormFields(agent.getSnapshot(), bytes));
      });
    });

  pdf
    .command("search-text <file> <query>")
    .description(
      "Search every page's text layer for the query and emit per-page hits ({ page, start, end, preview, match }). Schema: office-agent/pdf-search-text@1."
    )
    .option("--regex", "Treat <query> as a regular expression", false)
    .option("--case", "Case-sensitive search (default: case-insensitive)", false)
    .action(async (file: string, query: string, opts: { regex: boolean; case: boolean }) => {
      await pdfAction(io, async () => {
        const bytes = await readPdfBytes(file);
        const agent = await PdfAgent.fromBuffer(bytes);
        writeJson(
          io,
          projectSearch(agent, {
            query,
            regex: opts.regex,
            caseSensitive: opts.case,
          })
        );
      });
    });

  pdf
    .command("export-markdown <file>")
    .description(
      "Render the PDF as Markdown via PdfAgent.toMarkdown(). Without --out the markdown is written verbatim to stdout; with --out a JSON summary is emitted (Schema: office-agent/pdf-export-markdown@1) and the markdown lands at the given path."
    )
    .option("--out <path>", "Optional output path; when omitted the markdown is written to stdout")
    .action(async (file: string, opts: { out?: string }) => {
      await pdfAction(io, async () => {
        const bytes = await readPdfBytes(file);
        const agent = await PdfAgent.fromBuffer(bytes);
        const md = agent.toMarkdown();
        if (opts.out) {
          const target = resolve(opts.out);
          await ensureDir(target);
          const buf = Buffer.from(md, "utf8");
          await writeFile(target, buf);
          writeJson(io, {
            schema: "office-agent/pdf-export-markdown@1",
            in: file,
            out: opts.out,
            bytes: buf.byteLength,
            summary: `wrote ${buf.byteLength} bytes of markdown`,
          });
        } else {
          io.stdout.write(md);
          if (!md.endsWith("\n")) io.stdout.write("\n");
        }
      });
    });
}

// ── Mutate commands ──────────────────────────────────────────────────────

function registerMutateCommands(pdf: CommanderCommand, io: IO): void {
  pdf
    .command("rotate-pages <file>")
    .description(
      "Rotate selected pages by --delta (multiple of 90). Schema: office-agent/pdf-rotate-pages@1."
    )
    .requiredOption("--pages <list>", "Comma-separated 1-indexed page numbers (e.g. 1,2,5)")
    .requiredOption("--delta <degrees>", "Rotation delta in degrees (±90/180/270)", parseIntArg)
    .requiredOption("--out <path>", "Output PDF path")
    .action(async (file: string, opts: { pages: string; delta: number; out: string }) => {
      await pdfAction(io, async () => {
        const pages = parseIntList(opts.pages, "--pages");
        if (![90, 180, 270, -90, -180, -270].includes(opts.delta)) {
          throw new PdfError("invalid-delta", `--delta must be ±90/±180/±270 (got ${opts.delta})`);
        }
        const bytes = await readPdfBytes(file);
        const out = await rotatePages(bytes, {
          pages,
          delta: opts.delta as 90 | 180 | 270 | -90 | -180 | -270,
        });
        await writeMutationOutput(io, "office-agent/pdf-rotate-pages@1", file, opts.out, out, {
          summary: `rotated ${pages.length} page${pages.length === 1 ? "" : "s"} by ${opts.delta}°`,
        });
      });
    });

  pdf
    .command("reorder-pages <file>")
    .description("Reorder pages to a new permutation of 1..N. Schema: office-agent/pdf-reorder-pages@1.")
    .requiredOption("--order <list>", "Comma-separated 1-indexed permutation (e.g. 3,1,2)")
    .requiredOption("--out <path>", "Output PDF path")
    .action(async (file: string, opts: { order: string; out: string }) => {
      await pdfAction(io, async () => {
        const order = parseIntList(opts.order, "--order");
        const bytes = await readPdfBytes(file);
        const out = await reorderPages(bytes, { order });
        await writeMutationOutput(io, "office-agent/pdf-reorder-pages@1", file, opts.out, out, {
          summary: `reordered ${order.length} pages`,
        });
      });
    });

  pdf
    .command("delete-pages <file>")
    .description("Drop the selected pages. Schema: office-agent/pdf-delete-pages@1.")
    .requiredOption("--pages <list>", "Comma-separated 1-indexed page numbers (e.g. 2,4)")
    .requiredOption("--out <path>", "Output PDF path")
    .action(async (file: string, opts: { pages: string; out: string }) => {
      await pdfAction(io, async () => {
        const pages = parseIntList(opts.pages, "--pages");
        const bytes = await readPdfBytes(file);
        const out = await deletePages(bytes, { pages });
        await writeMutationOutput(io, "office-agent/pdf-delete-pages@1", file, opts.out, out, {
          summary: `deleted ${pages.length} page${pages.length === 1 ? "" : "s"}`,
        });
      });
    });

  pdf
    .command("merge <a> <b> [more...]")
    .description("Concatenate two-or-more PDFs into a single document. Schema: office-agent/pdf-merge@1.")
    .requiredOption("--out <path>", "Output PDF path")
    .action(async (a: string, b: string, more: string[], opts: { out: string }) => {
      await pdfAction(io, async () => {
        const inputs = [a, b, ...more];
        const buffers = await Promise.all(inputs.map((p) => readPdfBytes(p)));
        const out = await mergePdfs({ inputs: buffers });
        const target = resolve(opts.out);
        await ensureDir(target);
        const buf = Buffer.from(out);
        await writeFile(target, buf);
        writeJson(io, {
          schema: "office-agent/pdf-merge@1",
          inputs,
          out: opts.out,
          bytes: buf.byteLength,
          summary: `merged ${inputs.length} PDFs`,
        });
      });
    });

  pdf
    .command("split <file>")
    .description(
      "Split the PDF before page --at into two parts. Writes <prefix>-001.pdf and <prefix>-002.pdf. Schema: office-agent/pdf-split@1."
    )
    .requiredOption("--at <n>", "1-indexed split point; pages 1..at-1 land in part 1", parseIntArg)
    .requiredOption("--out-prefix <path>", "Output path prefix (e.g. ./parts/section)")
    .action(async (file: string, opts: { at: number; outPrefix: string }) => {
      await pdfAction(io, async () => {
        const bytes = await readPdfBytes(file);
        const agent = await PdfAgent.fromBuffer(bytes);
        const total = agent.getSnapshot().root.pages.length;
        if (opts.at < 2 || opts.at > total) {
          throw new PdfError(
            "invalid-split-point",
            `--at ${opts.at} must be in 2..${total} (split happens before page --at)`
          );
        }
        const result = await splitPdf(bytes, {
          strategy: {
            kind: "ranges",
            ranges: [
              [1, opts.at - 1],
              [opts.at, total],
            ],
          },
        });
        const outputs: Array<{ out: string; pages: ReadonlyArray<number>; bytes: number }> = [];
        for (let i = 0; i < result.parts.length; i++) {
          const part = result.parts[i];
          const target = `${opts.outPrefix}-${String(i + 1).padStart(3, "0")}.pdf`;
          const resolved = resolve(target);
          await ensureDir(resolved);
          const buf = Buffer.from(part.bytes);
          await writeFile(resolved, buf);
          outputs.push({ out: target, pages: part.pages, bytes: buf.byteLength });
        }
        writeJson(io, {
          schema: "office-agent/pdf-split@1",
          in: file,
          outputs,
          summary: `split into ${outputs.length} parts at page ${opts.at}`,
        });
      });
    });

  pdf
    .command("extract-pages <file>")
    .description(
      "Build a fresh PDF containing only the selected pages. Schema: office-agent/pdf-extract-pages@1."
    )
    .requiredOption("--pages <list>", "Comma-separated 1-indexed page numbers")
    .requiredOption("--out <path>", "Output PDF path")
    .action(async (file: string, opts: { pages: string; out: string }) => {
      await pdfAction(io, async () => {
        const pages = parseIntList(opts.pages, "--pages");
        const bytes = await readPdfBytes(file);
        const out = await extractPages(bytes, { pages });
        await writeMutationOutput(io, "office-agent/pdf-extract-pages@1", file, opts.out, out, {
          summary: `extracted ${pages.length} page${pages.length === 1 ? "" : "s"}`,
        });
      });
    });

  pdf
    .command("set-metadata <file>")
    .description(
      "Patch document Info dictionary fields (title/author/subject/keywords/creator/producer). Schema: office-agent/pdf-set-metadata@1."
    )
    .option("--title <text>", "Document title")
    .option("--author <name>", "Document author")
    .option("--subject <text>", "Document subject")
    .option("--keywords <text>", "Comma-separated keywords")
    .option("--creator <name>", "Creator application")
    .option("--producer <name>", "Producer application")
    .requiredOption("--out <path>", "Output PDF path")
    .action(
      async (
        file: string,
        opts: {
          title?: string;
          author?: string;
          subject?: string;
          keywords?: string;
          creator?: string;
          producer?: string;
          out: string;
        }
      ) => {
        await pdfAction(io, async () => {
          const patch: Record<string, string> = {};
          if (opts.title !== undefined) patch.title = opts.title;
          if (opts.author !== undefined) patch.author = opts.author;
          if (opts.subject !== undefined) patch.subject = opts.subject;
          if (opts.keywords !== undefined) patch.keywords = opts.keywords;
          if (opts.creator !== undefined) patch.creator = opts.creator;
          if (opts.producer !== undefined) patch.producer = opts.producer;
          if (Object.keys(patch).length === 0) {
            throw new PdfError(
              "no-metadata-fields",
              "set-metadata: pass at least one of --title/--author/--subject/--keywords/--creator/--producer"
            );
          }
          const bytes = await readPdfBytes(file);
          const out = await setMetadata(bytes, patch);
          await writeMutationOutput(io, "office-agent/pdf-set-metadata@1", file, opts.out, out, {
            summary: `set ${Object.keys(patch).join(", ")}`,
          });
        });
      }
    );

  pdf
    .command("add-watermark <file>")
    .description(
      "Stamp a centered text watermark on every page (or --pages subset). Schema: office-agent/pdf-add-watermark@1."
    )
    .requiredOption("--text <text>", "Watermark text")
    .option("--opacity <n>", "Opacity in [0,1] (default 0.18)", parseFloatArg)
    .option("--rotation <deg>", "Rotation in degrees (default 30)", parseFloatArg)
    .option("--font-size <n>", "Font size in points (default 60)", parseFloatArg)
    .option("--pages <list>", "Comma-separated 1-indexed pages (default: all)")
    .requiredOption("--out <path>", "Output PDF path")
    .action(
      async (
        file: string,
        opts: {
          text: string;
          opacity?: number;
          rotation?: number;
          fontSize?: number;
          pages?: string;
          out: string;
        }
      ) => {
        await pdfAction(io, async () => {
          const bytes = await readPdfBytes(file);
          const watermark: Parameters<typeof addWatermark>[1] = {
            text: opts.text,
            ...(opts.opacity !== undefined ? { opacity: opts.opacity } : {}),
            ...(opts.rotation !== undefined ? { rotate: opts.rotation } : {}),
            ...(opts.fontSize !== undefined ? { fontSize: opts.fontSize } : {}),
            ...(opts.pages ? { pages: parseIntList(opts.pages, "--pages") } : {}),
          };
          const out = await addWatermark(bytes, watermark);
          await writeMutationOutput(io, "office-agent/pdf-add-watermark@1", file, opts.out, out, {
            summary: `stamped "${opts.text}" watermark`,
          });
        });
      }
    );

  pdf
    .command("add-page-numbers <file>")
    .description(
      "Draw page numbers on every page (or starting at --start). Schema: office-agent/pdf-add-page-numbers@1."
    )
    .option("--start <n>", "First page that should display a number (default 1)", parseIntArg)
    .option(
      "--position <pos>",
      "top-left | top-center | top-right | bottom-left | bottom-center | bottom-right (default bottom-center)"
    )
    .option("--font-size <n>", "Font size in points (default 10)", parseFloatArg)
    .option("--margin <n>", "Margin from page edge in PDF user units (default 24)", parseFloatArg)
    .option(
      "--format <fmt>",
      'Format string with {page} and {total} placeholders (default "{page} / {total}")'
    )
    .requiredOption("--out <path>", "Output PDF path")
    .action(
      async (
        file: string,
        opts: {
          start?: number;
          position?: string;
          fontSize?: number;
          margin?: number;
          format?: string;
          out: string;
        }
      ) => {
        await pdfAction(io, async () => {
          const position = opts.position ? parsePagePosition(opts.position) : undefined;
          const bytes = await readPdfBytes(file);
          const out = await addPageNumbers(bytes, {
            ...(opts.start !== undefined ? { startAt: opts.start } : {}),
            ...(position ? { position } : {}),
            ...(opts.fontSize !== undefined ? { fontSize: opts.fontSize } : {}),
            ...(opts.margin !== undefined ? { margin: opts.margin } : {}),
            ...(opts.format ? { format: opts.format } : {}),
          });
          await writeMutationOutput(io, "office-agent/pdf-add-page-numbers@1", file, opts.out, out, {
            summary: `added page numbers (position=${position ?? "bottom-center"})`,
          });
        });
      }
    );

  pdf
    .command("crop-pages <file>")
    .description(
      "Crop selected pages to a rectangle expressed in PDF user units (lower-left origin). Schema: office-agent/pdf-crop-pages@1."
    )
    .requiredOption("--pages <list>", "Comma-separated 1-indexed pages (or 'all')")
    .requiredOption("--rect <x1,y1,x2,y2>", 'Crop rectangle in PDF user units, e.g. "50,50,562,742"')
    .requiredOption("--out <path>", "Output PDF path")
    .action(async (file: string, opts: { pages: string; rect: string; out: string }) => {
      await pdfAction(io, async () => {
        const rect = parseRect(opts.rect);
        const pages = opts.pages === "all" ? "all" : parseIntList(opts.pages, "--pages");
        const bytes = await readPdfBytes(file);
        const out = await cropPagesToRect(bytes, pages, rect);
        await writeMutationOutput(io, "office-agent/pdf-crop-pages@1", file, opts.out, out, {
          summary: `cropped ${pages === "all" ? "all" : pages.length} pages to [${rect.join(",")}]`,
        });
      });
    });

  pdf
    .command("fill-form <file>")
    .description(
      "Fill AcroForm fields from a JSON object (string|boolean|string[] values). Pass --values @path.json to read from disk. Schema: office-agent/pdf-fill-form@1."
    )
    .requiredOption(
      "--values <json>",
      'JSON object literal or @path.json reference (e.g. \'{"first.name":"Ada"}\')'
    )
    .requiredOption("--out <path>", "Output PDF path")
    .action(async (file: string, opts: { values: string; out: string }) => {
      await pdfAction(io, async () => {
        const values = await readJsonValueArg(opts.values);
        if (typeof values !== "object" || values === null || Array.isArray(values)) {
          throw new PdfError(
            "invalid-values",
            "--values must decode to a JSON object mapping field name → string|boolean|string[]"
          );
        }
        const bytes = await readPdfBytes(file);
        const out = await fillForm(bytes, {
          values: values as Record<string, string | boolean | ReadonlyArray<string>>,
        });
        await writeMutationOutput(io, "office-agent/pdf-fill-form@1", file, opts.out, out, {
          summary: `filled ${Object.keys(values).length} form field${Object.keys(values).length === 1 ? "" : "s"}`,
        });
      });
    });

  pdf
    .command("flatten-form <file>")
    .description(
      "Flatten every form widget into static page content (no longer fillable). Schema: office-agent/pdf-flatten-form@1."
    )
    .requiredOption("--out <path>", "Output PDF path")
    .action(async (file: string, opts: { out: string }) => {
      await pdfAction(io, async () => {
        const bytes = await readPdfBytes(file);
        const out = await flattenForm(bytes);
        await writeMutationOutput(io, "office-agent/pdf-flatten-form@1", file, opts.out, out, {
          summary: "flattened form fields",
        });
      });
    });

  pdf
    .command("reset-form <file>")
    .description("Reset every form field to its empty/default value. Schema: office-agent/pdf-reset-form@1.")
    .requiredOption("--out <path>", "Output PDF path")
    .action(async (file: string, opts: { out: string }) => {
      await pdfAction(io, async () => {
        const bytes = await readPdfBytes(file);
        const out = await resetForm(bytes);
        await writeMutationOutput(io, "office-agent/pdf-reset-form@1", file, opts.out, out, {
          summary: "reset form fields to defaults",
        });
      });
    });

  pdf
    .command("add-text-layer <file>")
    .description(
      "Add an invisible OCR text layer on top of every page. Requires the optional `tesseract.js` peer dependency (install with `pnpm add tesseract.js`). Schema: office-agent/pdf-add-text-layer@1."
    )
    .option("--lang <code>", "Tesseract language code(s), e.g. 'eng' or 'deu+eng' (default eng+deu)")
    .option("--dpi <n>", "Render DPI used to rasterize each page (default 200)", parseIntArg)
    .requiredOption("--out <path>", "Output PDF path")
    .action(async (file: string, opts: { lang?: string; dpi?: number; out: string }) => {
      await pdfAction(io, async () => {
        const bytes = await readPdfBytes(file);
        const out = await addTextLayer(bytes, {
          ...(opts.lang ? { lang: opts.lang } : {}),
          ...(opts.dpi !== undefined ? { dpi: opts.dpi } : {}),
        });
        await writeMutationOutput(io, "office-agent/pdf-add-text-layer@1", file, opts.out, out, {
          summary: `added OCR text layer (lang=${opts.lang ?? "eng+deu"})`,
        });
      });
    });
}

// ── Internal helpers ─────────────────────────────────────────────────────

/**
 * Custom error type carrying a stable string `code` so the JSON
 * envelope on stderr can be machine-routed (`{ "error": code,
 * "message": msg }`). Caught and re-emitted by `pdfAction`; never
 * surfaces to the user as an exception.
 */
class PdfError extends Error {
  constructor(
    public readonly errorCode: string,
    message: string
  ) {
    super(message);
  }
}

/**
 * Wrap every PDF subcommand body so any thrown error becomes a
 * structured `{ "error": "<code>", "message": "<msg>" }` envelope on
 * stderr plus a silent CliError(1). LLM callers can `JSON.parse(stderr)`
 * on non-zero exit and immediately recover the failure shape.
 */
async function pdfAction(io: IO, fn: () => Promise<void>): Promise<void> {
  try {
    await fn();
  } catch (err) {
    const code =
      err instanceof PdfError
        ? err.errorCode
        : err instanceof Error && (err as { code?: string }).code === "ENOENT"
          ? "file-not-found"
          : "pdf-error";
    const message = err instanceof Error ? err.message : String(err);
    io.stderr.write(JSON.stringify({ error: code, message }) + "\n");
    throw new CliError(1, message, { silent: true });
  }
}

function writeJson(io: IO, value: unknown): void {
  io.stdout.write(JSON.stringify(value) + "\n");
}

async function readPdfBytes(path: string): Promise<Uint8Array> {
  try {
    const buf = await readFile(resolve(path));
    return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
  } catch (err) {
    if (err instanceof Error && (err as { code?: string }).code === "ENOENT") {
      throw new PdfError("file-not-found", `file not found: ${path}`);
    }
    throw err;
  }
}

async function ensureDir(targetFile: string): Promise<void> {
  await mkdir(dirname(targetFile), { recursive: true });
}

async function writeMutationOutput(
  io: IO,
  schema: string,
  input: string,
  outPath: string,
  bytes: Uint8Array,
  extra: { summary: string }
): Promise<void> {
  const target = resolve(outPath);
  await ensureDir(target);
  const buf = Buffer.from(bytes);
  await writeFile(target, buf);
  writeJson(io, {
    schema,
    in: input,
    out: outPath,
    bytes: buf.byteLength,
    summary: extra.summary,
  });
}

function parseIntList(raw: string, flag: string): number[] {
  const items = raw
    .split(",")
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
  if (items.length === 0) {
    throw new PdfError("invalid-list", `${flag}: expected a non-empty comma-separated list`);
  }
  const out: number[] = [];
  for (const item of items) {
    const n = Number.parseInt(item, 10);
    if (!Number.isFinite(n)) {
      throw new PdfError("invalid-list", `${flag}: "${item}" is not an integer`);
    }
    out.push(n);
  }
  return out;
}

function parseRect(raw: string): readonly [number, number, number, number] {
  const items = raw
    .split(",")
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
  if (items.length !== 4) {
    throw new PdfError("invalid-rect", `--rect must be "x1,y1,x2,y2"; got "${raw}"`);
  }
  const nums = items.map((s) => Number.parseFloat(s));
  if (nums.some((n) => !Number.isFinite(n))) {
    throw new PdfError("invalid-rect", `--rect contains a non-numeric component: "${raw}"`);
  }
  const [x1, y1, x2, y2] = nums as [number, number, number, number];
  if (x2 <= x1 || y2 <= y1) {
    throw new PdfError("invalid-rect", `--rect requires x2>x1 and y2>y1; got "${raw}"`);
  }
  return [x1, y1, x2, y2] as const;
}

function parsePagePosition(
  raw: string
): "top-left" | "top-center" | "top-right" | "bottom-left" | "bottom-center" | "bottom-right" {
  switch (raw) {
    case "top-left":
    case "top-center":
    case "top-right":
    case "bottom-left":
    case "bottom-center":
    case "bottom-right":
      return raw;
    default:
      throw new PdfError(
        "invalid-position",
        `--position must be top-left|top-center|top-right|bottom-left|bottom-center|bottom-right (got "${raw}")`
      );
  }
}

function parseIntArg(value: string): number {
  const n = Number.parseInt(value, 10);
  if (!Number.isFinite(n)) throw new PdfError("invalid-int", `expected integer, got "${value}"`);
  return n;
}

function parseFloatArg(value: string): number {
  const n = Number.parseFloat(value);
  if (!Number.isFinite(n)) throw new PdfError("invalid-float", `expected number, got "${value}"`);
  return n;
}

async function readJsonValueArg(raw: string): Promise<unknown> {
  if (raw.startsWith("@")) {
    const path = raw.slice(1);
    if (path.length === 0) {
      throw new PdfError("invalid-values", "--values @<path>: path is empty");
    }
    let text: string;
    try {
      text = await readFile(resolve(path), "utf8");
    } catch (err) {
      if (err instanceof Error && (err as { code?: string }).code === "ENOENT") {
        throw new PdfError("file-not-found", `--values @${path}: file not found`);
      }
      throw err;
    }
    try {
      return JSON.parse(text) as unknown;
    } catch (err) {
      throw new PdfError("invalid-values", `--values @${path}: invalid JSON (${(err as Error).message})`);
    }
  }
  try {
    return JSON.parse(raw) as unknown;
  } catch (err) {
    throw new PdfError("invalid-values", `--values: invalid JSON (${(err as Error).message})`);
  }
}

/**
 * Crop the requested pages to an absolute rectangle (PDF user units).
 *
 * `pdf-edit/cropPages` only accepts symmetric `margins` from the
 * existing crop box. To honour the CLI spec's per-page bounding-box
 * input we walk the parsed snapshot for each requested page, derive
 * the four-sided margins from the page's media size, and apply the
 * crop one page at a time. This costs an extra `loadPdf` round-trip
 * per page but keeps the agent package free of a direct pdf-lib
 * dependency (pdf-lib is already pinned by every pdf-* workspace
 * package, so transitive resolution is stable).
 */
async function cropPagesToRect(
  bytes: Uint8Array,
  pages: ReadonlyArray<number> | "all",
  rect: readonly [number, number, number, number]
): Promise<Uint8Array> {
  const agent = await PdfAgent.fromBuffer(bytes);
  const snapPages = agent.getSnapshot().root.pages;
  const total = snapPages.length;
  const target = pages === "all" ? Array.from({ length: total }, (_, i) => i + 1) : pages;
  for (const p of target) {
    if (p < 1 || p > total) {
      throw new PdfError(
        "page-out-of-range",
        `crop-pages: page ${p} out of range (document has ${total} page${total === 1 ? "" : "s"})`
      );
    }
  }
  const [x1, y1, x2, y2] = rect;
  let current = bytes;
  for (const p of target) {
    const page = snapPages[p - 1];
    const { width, height } = page;
    if (x1 < 0 || y1 < 0 || x2 > width || y2 > height) {
      throw new PdfError(
        "rect-out-of-page",
        `crop-pages: rect [${rect.join(",")}] exceeds page ${p} size (${width}×${height})`
      );
    }
    current = await cropPages(current, {
      pages: [p],
      margins: [x1, height - y2, width - x2, y1] as const,
    });
  }
  return current;
}
