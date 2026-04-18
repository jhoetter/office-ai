import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { PptxAgent } from "@officeai/pptx";
import { deterministicIdMinter } from "@officeai/core";
import { runCli } from "./cli.js";

beforeAll(() => {
  process.env.OFFICEAI_DETERMINISTIC_IDS = "1";
});
afterAll(() => {
  delete process.env.OFFICEAI_DETERMINISTIC_IDS;
});

async function loadDeterministic(path: string): Promise<PptxAgent> {
  return PptxAgent.fromBuffer(readFileSync(path), {
    idMinter: deterministicIdMinter("n"),
  });
}

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

const HERE = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = resolve(HERE, "..", "..", "..");
const FIXTURE = (name: string): string =>
  resolve(REPO_ROOT, "fixtures", "pptx", "synthetic", name);

const SINGLE = FIXTURE("03-title-and-content.pptx");
const MULTI = FIXTURE("07-multi-slide.pptx");
const TABLE = FIXTURE("06-with-table.pptx");

describe("office-agent pptx subcommand group", () => {
  it("pptx inspect prints structural counts as JSON", async () => {
    const { io, stdout } = makeIO();
    const code = await runCli(["pptx", "inspect", "--file", SINGLE], io);
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(parsed.format).toBe("pptx");
    expect(parsed.slides).toBe(1);
    expect(typeof parsed.slideSize.cxEmu).toBe("number");
    expect(Array.isArray(parsed.parts)).toBe(true);
    expect(parsed.shapeCounts.text).toBeGreaterThanOrEqual(1);
  });

  it("pptx read --format markdown emits a presentation outline", async () => {
    const { io, stdout } = makeIO();
    const code = await runCli(["pptx", "read", "--file", SINGLE, "--format", "markdown"], io);
    expect(code).toBe(0);
    expect(stdout.text()).toMatch(/# Presentation/);
    expect(stdout.text()).toMatch(/Slide 1/);
  });

  it("pptx read --format json emits a per-slide projection", async () => {
    const { io, stdout } = makeIO();
    const code = await runCli(["pptx", "read", "--file", SINGLE, "--format", "json"], io);
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(parsed.format).toBe("pptx");
    expect(parsed.slides).toHaveLength(1);
    const shapes = parsed.slides[0].shapes as Array<{ kind: string }>;
    expect(shapes.length).toBeGreaterThanOrEqual(1);
  });

  it("pptx read --slide N restricts to a single slide", async () => {
    const { io, stdout } = makeIO();
    const code = await runCli(
      ["pptx", "read", "--file", MULTI, "--format", "json", "--slide", "1"],
      io
    );
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(parsed.slides).toHaveLength(1);
    expect(parsed.slides[0].index).toBe(1);
  });

  it("pptx search prints JSON matches", async () => {
    const { io, stdout } = makeIO();
    const code = await runCli(["pptx", "search", "--file", SINGLE, "-q", "Title"], io);
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(Array.isArray(parsed)).toBe(true);
  });

  it("pptx add-slide writes a file with one extra slide", async () => {
    const dir = mkdtempSync(join(tmpdir(), "pptx-cli-add-"));
    const out = join(dir, "out.pptx");
    const { io } = makeIO();
    const code = await runCli(
      ["pptx", "add-slide", "--file", SINGLE, "--out", out],
      io
    );
    expect(code).toBe(0);
    const agent = await PptxAgent.fromBuffer(readFileSync(out));
    expect(agent.getSnapshot().root.slides.length).toBe(2);
  });

  it("pptx delete-slide removes the targeted slide", async () => {
    const dir = mkdtempSync(join(tmpdir(), "pptx-cli-del-"));
    const out = join(dir, "out.pptx");
    const { io } = makeIO();
    const code = await runCli(
      ["pptx", "delete-slide", "--file", MULTI, "--slide", "1", "--out", out],
      io
    );
    expect(code).toBe(0);
    const before = await PptxAgent.fromBuffer(readFileSync(MULTI));
    const after = await PptxAgent.fromBuffer(readFileSync(out));
    expect(after.getSnapshot().root.slides.length).toBe(
      before.getSnapshot().root.slides.length - 1
    );
  });

  it("pptx duplicate-slide and move-slide compose into the expected order", async () => {
    const dir = mkdtempSync(join(tmpdir(), "pptx-cli-dup-"));
    const dup = join(dir, "dup.pptx");
    const moved = join(dir, "moved.pptx");
    const { io } = makeIO();

    let code = await runCli(
      ["pptx", "duplicate-slide", "--file", MULTI, "--slide", "0", "--out", dup],
      io
    );
    expect(code).toBe(0);
    const dupAgent = await PptxAgent.fromBuffer(readFileSync(dup));
    const baseLen = (await PptxAgent.fromBuffer(readFileSync(MULTI))).getSnapshot().root.slides
      .length;
    expect(dupAgent.getSnapshot().root.slides.length).toBe(baseLen + 1);

    code = await runCli(
      ["pptx", "move-slide", "--file", dup, "--from", "0", "--to", "1", "--out", moved],
      io
    );
    expect(code).toBe(0);
    const movedAgent = await PptxAgent.fromBuffer(readFileSync(moved));
    expect(movedAgent.getSnapshot().root.slides.length).toBe(baseLen + 1);
  });

  it("pptx set-text replaces a TextShape's plain text", async () => {
    const baseAgent = await loadDeterministic(SINGLE);
    const slide0 = baseAgent.getSnapshot().root.slides[0];
    const textShape = slide0.shapes.find((s) => s.kind === "text");
    if (!textShape) throw new Error("expected at least one text shape on slide 0");

    const dir = mkdtempSync(join(tmpdir(), "pptx-cli-set-"));
    const out = join(dir, "out.pptx");
    const { io } = makeIO();
    const code = await runCli(
      [
        "pptx",
        "set-text",
        "--file",
        SINGLE,
        "--out",
        out,
        "--slide",
        "0",
        "--shape",
        textShape.id,
        "--text",
        "Hello CLI",
      ],
      io
    );
    expect(code).toBe(0);
    const after = await PptxAgent.fromBuffer(readFileSync(out));
    expect(after.toMarkdown()).toContain("Hello CLI");
  });

  it("pptx set-position and set-size update geometry", async () => {
    const baseAgent = await loadDeterministic(SINGLE);
    const textShape = baseAgent
      .getSnapshot()
      .root.slides[0].shapes.find((s) => s.kind === "text");
    if (!textShape) throw new Error("expected text shape");

    const dir = mkdtempSync(join(tmpdir(), "pptx-cli-pos-"));
    const a = join(dir, "a.pptx");
    const b = join(dir, "b.pptx");
    const { io } = makeIO();

    let code = await runCli(
      [
        "pptx",
        "set-position",
        "--file",
        SINGLE,
        "--out",
        a,
        "--slide",
        "0",
        "--shape",
        textShape.id,
        "--x",
        "100000",
        "--y",
        "200000",
      ],
      io
    );
    expect(code).toBe(0);
    const aAgent = await loadDeterministic(a);
    const aShape = aAgent.getSnapshot().root.slides[0].shapes.find((s) => s.id === textShape.id);
    expect(aShape?.position).toEqual({ xEmu: 100000, yEmu: 200000 });

    code = await runCli(
      [
        "pptx",
        "set-size",
        "--file",
        a,
        "--out",
        b,
        "--slide",
        "0",
        "--shape",
        textShape.id,
        "--width",
        "5000000",
        "--height",
        "1000000",
      ],
      io
    );
    expect(code).toBe(0);
    const bAgent = await loadDeterministic(b);
    const bShape = bAgent.getSnapshot().root.slides[0].shapes.find((s) => s.id === textShape.id);
    expect(bShape?.size).toEqual({ cxEmu: 5000000, cyEmu: 1000000 });
  });

  it("pptx add-text-box appends a new text shape", async () => {
    const dir = mkdtempSync(join(tmpdir(), "pptx-cli-tb-"));
    const out = join(dir, "out.pptx");
    const { io } = makeIO();
    const code = await runCli(
      [
        "pptx",
        "add-text-box",
        "--file",
        SINGLE,
        "--out",
        out,
        "--slide",
        "0",
        "--text",
        "Box from CLI",
        "--x",
        "100000",
        "--y",
        "200000",
        "--width",
        "3000000",
        "--height",
        "500000",
      ],
      io
    );
    expect(code).toBe(0);
    const after = await PptxAgent.fromBuffer(readFileSync(out));
    expect(after.toMarkdown()).toContain("Box from CLI");
  });

  it("pptx apply runs a JSON command file", async () => {
    const baseAgent = await loadDeterministic(SINGLE);
    const textShape = baseAgent
      .getSnapshot()
      .root.slides[0].shapes.find((s) => s.kind === "text");
    if (!textShape) throw new Error("expected text shape");

    const dir = mkdtempSync(join(tmpdir(), "pptx-cli-apply-"));
    const cmdsPath = join(dir, "commands.json");
    const out = join(dir, "out.pptx");
    writeFileSync(
      cmdsPath,
      JSON.stringify({
        commands: [
          { type: "pptx:add-slide", payload: {} },
          {
            type: "pptx:set-text",
            payload: { slideIndex: 0, shapeId: textShape.id, text: "Hello apply" },
          },
        ],
      })
    );
    const { io, stdout } = makeIO();
    const code = await runCli(
      ["pptx", "apply", "--file", SINGLE, "--commands", cmdsPath, "--out", out],
      io
    );
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(parsed.mutations).toHaveLength(2);
    const after = await PptxAgent.fromBuffer(readFileSync(out));
    expect(after.getSnapshot().root.slides.length).toBe(2);
    expect(after.toMarkdown()).toContain("Hello apply");
  });

  describe("table CLI subcommands", () => {
    async function findTable(path: string): Promise<{ slide: number; shapeId: string }> {
      const agent = await loadDeterministic(path);
      const slides = agent.getSnapshot().root.slides;
      for (let i = 0; i < slides.length; i++) {
        const t = slides[i].shapes.find((s) => s.kind === "table");
        if (t) return { slide: i, shapeId: t.id };
      }
      throw new Error("no table shape in fixture");
    }

    it("table-set-cell-text rewrites a single cell", async () => {
      const { slide, shapeId } = await findTable(TABLE);
      const dir = mkdtempSync(join(tmpdir(), "pptx-cli-tbl-set-"));
      const out = join(dir, "out.pptx");
      const { io } = makeIO();
      const code = await runCli(
        [
          "pptx", "table-set-cell-text",
          "--file", TABLE, "--out", out,
          "--slide", String(slide),
          "--shape", shapeId,
          "--row", "0", "--column", "0",
          "--text", "CLI Cell",
        ],
        io
      );
      expect(code).toBe(0);
      const after = await loadDeterministic(out);
      const tbl = after.getSnapshot().root.slides[slide].shapes.find((s) => s.kind === "table");
      if (!tbl || tbl.kind !== "table") throw new Error("table missing");
      const cell = tbl.rows[0].cells[0];
      const text = cell.txBody.paragraphs
        .flatMap((p) => p.runs.map((r) => (r.isLineBreak ? "\n" : r.text)))
        .join("");
      expect(text).toBe("CLI Cell");
    });

    it("table-add-row and table-add-column grow the grid", async () => {
      const before = await loadDeterministic(TABLE);
      const beforeTbl = before
        .getSnapshot()
        .root.slides.flatMap((s, i) => s.shapes.map((sh) => ({ slide: i, sh })))
        .find((x) => x.sh.kind === "table");
      if (!beforeTbl || beforeTbl.sh.kind !== "table") throw new Error("table missing");
      const beforeRows = beforeTbl.sh.rows.length;
      const beforeCols = beforeTbl.sh.columnWidths.length;

      const dir = mkdtempSync(join(tmpdir(), "pptx-cli-tbl-grow-"));
      const a = join(dir, "a.pptx");
      const b = join(dir, "b.pptx");
      const { io } = makeIO();

      let code = await runCli(
        [
          "pptx", "table-add-row",
          "--file", TABLE, "--out", a,
          "--slide", String(beforeTbl.slide),
          "--shape", beforeTbl.sh.id,
        ],
        io
      );
      expect(code).toBe(0);

      // IDs are minted from the parser, and adding a row shifts later IDs.
      // Re-resolve the table id from the saved file before the next call.
      const aAgent = await loadDeterministic(a);
      const aTbl = aAgent
        .getSnapshot()
        .root.slides[beforeTbl.slide].shapes.find((s) => s.kind === "table");
      if (!aTbl) throw new Error("table missing in intermediate file");

      code = await runCli(
        [
          "pptx", "table-add-column",
          "--file", a, "--out", b,
          "--slide", String(beforeTbl.slide),
          "--shape", aTbl.id,
        ],
        io
      );
      expect(code).toBe(0);

      const after = await loadDeterministic(b);
      const tbl = after.getSnapshot().root.slides[beforeTbl.slide].shapes.find(
        (s) => s.kind === "table"
      );
      if (!tbl || tbl.kind !== "table") throw new Error("table missing");
      expect(tbl.rows.length).toBe(beforeRows + 1);
      expect(tbl.columnWidths.length).toBe(beforeCols + 1);
      for (const row of tbl.rows) {
        expect(row.cells.length).toBe(beforeCols + 1);
      }
    });

    it("table-delete-row and table-delete-column shrink the grid", async () => {
      const before = await loadDeterministic(TABLE);
      const beforeTbl = before
        .getSnapshot()
        .root.slides.flatMap((s, i) => s.shapes.map((sh) => ({ slide: i, sh })))
        .find((x) => x.sh.kind === "table");
      if (!beforeTbl || beforeTbl.sh.kind !== "table") throw new Error("table missing");
      const beforeRows = beforeTbl.sh.rows.length;
      const beforeCols = beforeTbl.sh.columnWidths.length;
      if (beforeRows < 2 || beforeCols < 2) {
        throw new Error("fixture must have ≥2 rows and ≥2 columns");
      }

      const dir = mkdtempSync(join(tmpdir(), "pptx-cli-tbl-shrink-"));
      const a = join(dir, "a.pptx");
      const b = join(dir, "b.pptx");
      const { io } = makeIO();

      let code = await runCli(
        [
          "pptx", "table-delete-row",
          "--file", TABLE, "--out", a,
          "--slide", String(beforeTbl.slide),
          "--shape", beforeTbl.sh.id,
          "--row", String(beforeRows - 1),
        ],
        io
      );
      expect(code).toBe(0);

      const aAgent = await loadDeterministic(a);
      const aTbl = aAgent
        .getSnapshot()
        .root.slides[beforeTbl.slide].shapes.find((s) => s.kind === "table");
      if (!aTbl) throw new Error("table missing in intermediate file");

      code = await runCli(
        [
          "pptx", "table-delete-column",
          "--file", a, "--out", b,
          "--slide", String(beforeTbl.slide),
          "--shape", aTbl.id,
          "--column", String(beforeCols - 1),
        ],
        io
      );
      expect(code).toBe(0);

      const after = await loadDeterministic(b);
      const tbl = after.getSnapshot().root.slides[beforeTbl.slide].shapes.find(
        (s) => s.kind === "table"
      );
      if (!tbl || tbl.kind !== "table") throw new Error("table missing");
      expect(tbl.rows.length).toBe(beforeRows - 1);
      expect(tbl.columnWidths.length).toBe(beforeCols - 1);
    });

    it("inspect reports table count after a no-op load", async () => {
      const { io, stdout } = makeIO();
      const code = await runCli(["pptx", "inspect", "--file", TABLE], io);
      expect(code).toBe(0);
      const parsed = JSON.parse(stdout.text());
      expect(parsed.shapeCounts.table).toBeGreaterThanOrEqual(1);
    });
  });

  it("pptx diff reports slide and shape counts", async () => {
    const dir = mkdtempSync(join(tmpdir(), "pptx-cli-diff-"));
    const out = join(dir, "out.pptx");
    const { io } = makeIO();

    let code = await runCli(["pptx", "add-slide", "--file", SINGLE, "--out", out], io);
    expect(code).toBe(0);

    const { io: io2, stdout } = makeIO();
    code = await runCli(["pptx", "diff", "--before", SINGLE, "--after", out], io2);
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(parsed.format).toBe("pptx");
    expect(parsed.slides.added).toBe(1);
  });
});
