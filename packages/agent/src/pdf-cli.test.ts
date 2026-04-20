/**
 * pdf-cli integration tests.
 *
 * Strategy: build synthetic PDF fixtures in-memory via pdf-lib, write
 * them to a temp dir, then drive the CLI via `runCli` (the same path
 * `office-agent` binary uses) so JSON-envelope behaviour is end-to-end.
 *
 * Covers ≥ 4 read commands and ≥ 3 mutate commands (well over) plus
 * error-envelope round-tripping.
 */
import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { PDFDocument, StandardFonts } from "pdf-lib";
import { runCli } from "./cli.js";

beforeAll(() => {
  process.env.OFFICEAI_DETERMINISTIC_IDS = "1";
});
afterAll(() => {
  delete process.env.OFFICEAI_DETERMINISTIC_IDS;
});

class CapturedStream {
  chunks: string[] = [];
  write(s: string | Uint8Array): boolean {
    this.chunks.push(typeof s === "string" ? s : Buffer.from(s).toString("utf8"));
    return true;
  }
  text(): string {
    return this.chunks.join("");
  }
}

function makeIO() {
  const stdout = new CapturedStream();
  const stderr = new CapturedStream();
  return {
    io: {
      stdout: stdout as unknown as NodeJS.WritableStream,
      stderr: stderr as unknown as NodeJS.WritableStream,
    },
    stdout,
    stderr,
  };
}

const TMP = mkdtempSync(join(tmpdir(), "office-agent-pdf-cli-"));

async function buildPdf(opts: {
  pages: number;
  label?: string;
  title?: string;
  author?: string;
  body?: (pageIndex: number) => string;
}): Promise<Uint8Array> {
  const pdf = await PDFDocument.create();
  if (opts.title) pdf.setTitle(opts.title);
  if (opts.author) pdf.setAuthor(opts.author);
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  for (let i = 1; i <= opts.pages; i++) {
    const page = pdf.addPage([612, 792]);
    const text = opts.body ? opts.body(i - 1) : `${opts.label ?? "Page"} ${i}`;
    page.drawText(text, { x: 50, y: 720, size: 18, font });
  }
  return pdf.save();
}

async function buildAcroFormPdf(): Promise<Uint8Array> {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const page = pdf.addPage([612, 792]);
  page.drawText("Form fixture", { x: 50, y: 740, size: 18, font });
  const form = pdf.getForm();
  const tf = form.createTextField("first.name");
  tf.setText("");
  tf.addToPage(page, { x: 50, y: 700, width: 200, height: 24 });
  const cb = form.createCheckBox("agree");
  cb.addToPage(page, { x: 50, y: 660, width: 16, height: 16 });
  return pdf.save();
}

async function writeFixture(name: string, bytes: Uint8Array): Promise<string> {
  const path = join(TMP, name);
  writeFileSync(path, Buffer.from(bytes));
  return path;
}

describe("office-agent pdf — create command", () => {
  it("create writes a fresh blank PDF and emits a versioned summary", async () => {
    const out = join(TMP, "blank.pdf");
    const { io, stdout } = makeIO();
    const code = await runCli(["pdf", "create", "--out", out], io);
    expect(code).toBe(0);
    const summary = JSON.parse(stdout.text());
    expect(summary.schema).toBe("office-agent/pdf-create@1");
    expect(summary.format).toBe("pdf");
    expect(summary.wrote).toBe(out);
    expect(summary.bytes).toBeGreaterThan(0);
    const written = readFileSync(out);
    expect(written.byteLength).toBe(summary.bytes);
    expect(written.subarray(0, 4).toString("ascii")).toBe("%PDF");
  });
});

describe("office-agent pdf — read commands", () => {
  it("read-metadata returns a versioned envelope", async () => {
    const path = await writeFixture("meta.pdf", await buildPdf({ pages: 3, title: "Hello", author: "Ada" }));
    const { io, stdout } = makeIO();
    const code = await runCli(["pdf", "read-metadata", path], io);
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(parsed.schema).toBe("office-agent/pdf-read-metadata@1");
    expect(parsed.title).toBe("Hello");
    expect(parsed.author).toBe("Ada");
    expect(parsed.pageCount).toBe(3);
    expect(typeof parsed.engine).toBe("string");
    expect(typeof parsed.linearized).toBe("boolean");
    expect(parsed.encryption).toEqual({ hasUserPassword: false, hasOwnerPassword: false });
  });

  it("read-page projects a single page with size + text", async () => {
    const path = await writeFixture(
      "pages.pdf",
      await buildPdf({ pages: 2, body: (i) => `Body ${i}-XYZ`, label: "Body" })
    );
    const { io, stdout } = makeIO();
    const code = await runCli(["pdf", "read-page", path, "--page", "2"], io);
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(parsed.schema).toBe("office-agent/pdf-read-page@1");
    expect(parsed.page.pageNumber).toBe(2);
    expect(parsed.page.width).toBe(612);
    expect(parsed.page.height).toBe(792);
    expect(parsed.page.rotation).toBe(0);
    expect(parsed.page.text.length).toBeGreaterThan(0);
    expect(Array.isArray(parsed.annotations)).toBe(true);
    expect(Array.isArray(parsed.formFields)).toBe(true);
  });

  it("read-page out of range emits a structured error envelope on stderr", async () => {
    const path = await writeFixture("oob.pdf", await buildPdf({ pages: 1 }));
    const { io, stdout, stderr } = makeIO();
    const code = await runCli(["pdf", "read-page", path, "--page", "99"], io);
    expect(code).toBe(1);
    expect(stdout.text()).toBe("");
    const lines = stderr.text().trim().split("\n").filter(Boolean);
    const last = JSON.parse(lines[lines.length - 1]);
    expect(last.error).toBe("page-out-of-range");
    expect(last.message).toMatch(/--page 99/);
  });

  it("search-text returns per-page hits", async () => {
    const path = await writeFixture(
      "search.pdf",
      await buildPdf({ pages: 3, body: (i) => `alpha ${i % 2 === 0 ? "BETA" : "gamma"}` })
    );
    const { io, stdout } = makeIO();
    const code = await runCli(["pdf", "search-text", path, "BETA"], io);
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(parsed.schema).toBe("office-agent/pdf-search-text@1");
    expect(parsed.results.length).toBeGreaterThan(0);
    for (const hit of parsed.results) {
      expect(typeof hit.page).toBe("number");
      expect(typeof hit.preview).toBe("string");
      expect(hit.match.toLowerCase()).toBe("beta");
    }
  });

  it("read-outline emits an outline envelope (empty array allowed)", async () => {
    const path = await writeFixture("outline.pdf", await buildPdf({ pages: 1 }));
    const { io, stdout } = makeIO();
    const code = await runCli(["pdf", "read-outline", path], io);
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(parsed.schema).toBe("office-agent/pdf-read-outline@1");
    expect(Array.isArray(parsed.outline)).toBe(true);
  });

  it("list-form-fields enumerates AcroForm widgets", async () => {
    const path = await writeFixture("form.pdf", await buildAcroFormPdf());
    const { io, stdout } = makeIO();
    const code = await runCli(["pdf", "list-form-fields", path], io);
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(parsed.schema).toBe("office-agent/pdf-list-form-fields@1");
    const byName = Object.fromEntries(
      (parsed.fields as Array<{ name: string; type: string }>).map((f) => [f.name, f])
    );
    expect(byName["first.name"].type).toBe("text");
    expect(byName["agree"].type).toBe("checkbox");
  });
});

describe("office-agent pdf — mutate commands", () => {
  it("rotate-pages writes a rotated PDF and emits a summary envelope", async () => {
    const path = await writeFixture("rot-in.pdf", await buildPdf({ pages: 2 }));
    const out = join(TMP, "rot-out.pdf");
    const { io, stdout } = makeIO();
    const code = await runCli(
      ["pdf", "rotate-pages", path, "--pages", "1", "--delta", "90", "--out", out],
      io
    );
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(parsed.schema).toBe("office-agent/pdf-rotate-pages@1");
    expect(parsed.out).toBe(out);
    expect(parsed.bytes).toBeGreaterThan(0);
    expect(parsed.summary).toMatch(/rotated 1 page by 90°/);
    const reopened = await PDFDocument.load(readFileSync(out));
    expect(reopened.getPage(0).getRotation().angle).toBe(90);
    expect(reopened.getPage(1).getRotation().angle).toBe(0);
  });

  it("delete-pages drops the requested pages", async () => {
    const path = await writeFixture("del-in.pdf", await buildPdf({ pages: 4 }));
    const out = join(TMP, "del-out.pdf");
    const { io, stdout } = makeIO();
    const code = await runCli(["pdf", "delete-pages", path, "--pages", "2,4", "--out", out], io);
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(parsed.schema).toBe("office-agent/pdf-delete-pages@1");
    const reopened = await PDFDocument.load(readFileSync(out));
    expect(reopened.getPageCount()).toBe(2);
  });

  it("merge concatenates two PDFs", async () => {
    const a = await writeFixture("merge-a.pdf", await buildPdf({ pages: 2, label: "A" }));
    const b = await writeFixture("merge-b.pdf", await buildPdf({ pages: 3, label: "B" }));
    const out = join(TMP, "merge-out.pdf");
    const { io, stdout } = makeIO();
    const code = await runCli(["pdf", "merge", a, b, "--out", out], io);
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(parsed.schema).toBe("office-agent/pdf-merge@1");
    expect(parsed.summary).toMatch(/merged 2 PDFs/);
    const reopened = await PDFDocument.load(readFileSync(out));
    expect(reopened.getPageCount()).toBe(5);
  });

  it("split slices a PDF before --at into two parts", async () => {
    const path = await writeFixture("split-in.pdf", await buildPdf({ pages: 4 }));
    const prefix = join(TMP, "split-out");
    const { io, stdout } = makeIO();
    const code = await runCli(["pdf", "split", path, "--at", "3", "--out-prefix", prefix], io);
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(parsed.schema).toBe("office-agent/pdf-split@1");
    expect(parsed.outputs.length).toBe(2);
    const part1 = await PDFDocument.load(readFileSync(`${prefix}-001.pdf`));
    const part2 = await PDFDocument.load(readFileSync(`${prefix}-002.pdf`));
    expect(part1.getPageCount()).toBe(2);
    expect(part2.getPageCount()).toBe(2);
  });

  it("set-metadata patches the Info dictionary", async () => {
    const path = await writeFixture("meta-in.pdf", await buildPdf({ pages: 1 }));
    const out = join(TMP, "meta-out.pdf");
    const { io, stdout } = makeIO();
    const code = await runCli(
      ["pdf", "set-metadata", path, "--title", "Updated", "--author", "Linus", "--out", out],
      io
    );
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(parsed.schema).toBe("office-agent/pdf-set-metadata@1");
    const reopened = await PDFDocument.load(readFileSync(out));
    expect(reopened.getTitle()).toBe("Updated");
    expect(reopened.getAuthor()).toBe("Linus");
  });

  it("fill-form accepts a JSON values blob and writes a filled PDF", async () => {
    const path = await writeFixture("fill-in.pdf", await buildAcroFormPdf());
    const out = join(TMP, "fill-out.pdf");
    const { io, stdout } = makeIO();
    const code = await runCli(
      [
        "pdf",
        "fill-form",
        path,
        "--values",
        JSON.stringify({ "first.name": "Ada", agree: true }),
        "--out",
        out,
      ],
      io
    );
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(parsed.schema).toBe("office-agent/pdf-fill-form@1");
    expect(parsed.summary).toMatch(/filled 2 form fields/);
    // Re-list fields from the saved file via the CLI to confirm the
    // values round-tripped.
    const { io: io2, stdout: stdout2 } = makeIO();
    const code2 = await runCli(["pdf", "list-form-fields", out], io2);
    expect(code2).toBe(0);
    const parsed2 = JSON.parse(stdout2.text());
    const byName = Object.fromEntries(
      (parsed2.fields as Array<{ name: string; value?: unknown }>).map((f) => [f.name, f])
    );
    expect(byName["first.name"].value).toBe("Ada");
    expect(byName["agree"].value).toBe(true);
  });

  it("flatten-form drops every form field", async () => {
    const path = await writeFixture("flat-in.pdf", await buildAcroFormPdf());
    const out = join(TMP, "flat-out.pdf");
    const { io, stdout } = makeIO();
    const code = await runCli(["pdf", "flatten-form", path, "--out", out], io);
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(parsed.schema).toBe("office-agent/pdf-flatten-form@1");
    const { io: io2, stdout: stdout2 } = makeIO();
    const code2 = await runCli(["pdf", "list-form-fields", out], io2);
    expect(code2).toBe(0);
    const fields = JSON.parse(stdout2.text()).fields as ReadonlyArray<unknown>;
    expect(fields.length).toBe(0);
  });
});
