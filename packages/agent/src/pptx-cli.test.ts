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
const FIXTURE = (name: string): string => resolve(REPO_ROOT, "fixtures", "pptx", "synthetic", name);

const SINGLE = FIXTURE("03-title-and-content.pptx");
const MULTI = FIXTURE("07-multi-slide.pptx");
const TABLE = FIXTURE("06-with-table.pptx");
const CHART = FIXTURE("09-with-chart.pptx");
const ANIM = FIXTURE("10-with-anim.pptx");

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
    const code = await runCli(["pptx", "read", "--file", MULTI, "--format", "json", "--slide", "1"], io);
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
    const code = await runCli(["pptx", "add-slide", "--file", SINGLE, "--out", out], io);
    expect(code).toBe(0);
    const agent = await PptxAgent.fromBuffer(readFileSync(out));
    expect(agent.getSnapshot().root.slides.length).toBe(2);
  });

  it("pptx delete-slide removes the targeted slide", async () => {
    const dir = mkdtempSync(join(tmpdir(), "pptx-cli-del-"));
    const out = join(dir, "out.pptx");
    const { io } = makeIO();
    const code = await runCli(["pptx", "delete-slide", "--file", MULTI, "--slide", "1", "--out", out], io);
    expect(code).toBe(0);
    const before = await PptxAgent.fromBuffer(readFileSync(MULTI));
    const after = await PptxAgent.fromBuffer(readFileSync(out));
    expect(after.getSnapshot().root.slides.length).toBe(before.getSnapshot().root.slides.length - 1);
  });

  it("pptx duplicate-slide and move-slide compose into the expected order", async () => {
    const dir = mkdtempSync(join(tmpdir(), "pptx-cli-dup-"));
    const dup = join(dir, "dup.pptx");
    const moved = join(dir, "moved.pptx");
    const { io } = makeIO();

    let code = await runCli(["pptx", "duplicate-slide", "--file", MULTI, "--slide", "0", "--out", dup], io);
    expect(code).toBe(0);
    const dupAgent = await PptxAgent.fromBuffer(readFileSync(dup));
    const baseLen = (await PptxAgent.fromBuffer(readFileSync(MULTI))).getSnapshot().root.slides.length;
    expect(dupAgent.getSnapshot().root.slides.length).toBe(baseLen + 1);

    code = await runCli(
      ["pptx", "move-slide", "--file", dup, "--from", "0", "--to", "1", "--out", moved],
      io
    );
    expect(code).toBe(0);
    const movedAgent = await PptxAgent.fromBuffer(readFileSync(moved));
    expect(movedAgent.getSnapshot().root.slides.length).toBe(baseLen + 1);
  });

  it('pptx set-paragraph-alignment writes algn="ctr" into the slide XML', async () => {
    const baseAgent = await loadDeterministic(SINGLE);
    const textShape = baseAgent.getSnapshot().root.slides[0].shapes.find((s) => s.kind === "text");
    if (!textShape) throw new Error("expected text shape");

    const dir = mkdtempSync(join(tmpdir(), "pptx-cli-align-"));
    const out = join(dir, "out.pptx");
    const { io } = makeIO();
    const code = await runCli(
      [
        "pptx",
        "set-paragraph-alignment",
        "--file",
        SINGLE,
        "--out",
        out,
        "--slide",
        "0",
        "--shape",
        textShape.id,
        "--alignment",
        "center",
        "--paragraph",
        "0",
      ],
      io
    );
    expect(code).toBe(0);
    const after = await loadDeterministic(out);
    const ts = after
      .getSnapshot()
      .root.slides[0].shapes.find((s): s is import("@officeai/pptx").TextShape => s.id === textShape.id);
    expect(ts?.txBody.paragraphs[0].properties.alignment).toBe("center");
  });

  it("pptx set-paragraph-alignment --clear strips the override", async () => {
    const baseAgent = await loadDeterministic(SINGLE);
    const textShape = baseAgent.getSnapshot().root.slides[0].shapes.find((s) => s.kind === "text");
    if (!textShape) throw new Error("expected text shape");

    const dir = mkdtempSync(join(tmpdir(), "pptx-cli-align-clear-"));
    const aligned = join(dir, "aligned.pptx");
    const cleared = join(dir, "cleared.pptx");
    const { io } = makeIO();

    let code = await runCli(
      [
        "pptx",
        "set-paragraph-alignment",
        "--file",
        SINGLE,
        "--out",
        aligned,
        "--slide",
        "0",
        "--shape",
        textShape.id,
        "--alignment",
        "right",
      ],
      io
    );
    expect(code).toBe(0);

    code = await runCli(
      [
        "pptx",
        "set-paragraph-alignment",
        "--file",
        aligned,
        "--out",
        cleared,
        "--slide",
        "0",
        "--shape",
        textShape.id,
        "--clear",
      ],
      io
    );
    expect(code).toBe(0);
    const after = await loadDeterministic(cleared);
    const ts = after
      .getSnapshot()
      .root.slides[0].shapes.find((s): s is import("@officeai/pptx").TextShape => s.id === textShape.id);
    for (const p of ts?.txBody.paragraphs ?? []) {
      expect(p.properties.alignment).toBeUndefined();
    }
  });

  it("pptx set-text-anchor sets and clears the bodyPr anchor", async () => {
    const baseAgent = await loadDeterministic(SINGLE);
    const textShape = baseAgent.getSnapshot().root.slides[0].shapes.find((s) => s.kind === "text");
    if (!textShape) throw new Error("expected text shape");

    const dir = mkdtempSync(join(tmpdir(), "pptx-cli-anchor-"));
    const out = join(dir, "out.pptx");
    const cleared = join(dir, "cleared.pptx");
    const { io } = makeIO();

    let code = await runCli(
      [
        "pptx",
        "set-text-anchor",
        "--file",
        SINGLE,
        "--out",
        out,
        "--slide",
        "0",
        "--shape",
        textShape.id,
        "--anchor",
        "middle",
      ],
      io
    );
    expect(code).toBe(0);
    const after = await loadDeterministic(out);
    const ts = after
      .getSnapshot()
      .root.slides[0].shapes.find((s): s is import("@officeai/pptx").TextShape => s.id === textShape.id);
    expect(ts?.txBody.bodyPrRaw?.attrs.anchor).toBe("ctr");

    code = await runCli(
      [
        "pptx",
        "set-text-anchor",
        "--file",
        out,
        "--out",
        cleared,
        "--slide",
        "0",
        "--shape",
        textShape.id,
        "--clear",
      ],
      io
    );
    expect(code).toBe(0);
    const cleanedAgent = await loadDeterministic(cleared);
    const ts2 = cleanedAgent
      .getSnapshot()
      .root.slides[0].shapes.find((s): s is import("@officeai/pptx").TextShape => s.id === textShape.id);
    expect(ts2?.txBody.bodyPrRaw?.attrs.anchor).toBeUndefined();
  });

  it("pptx set-paragraph-alignment rejects --alignment combined with --clear", async () => {
    const baseAgent = await loadDeterministic(SINGLE);
    const textShape = baseAgent.getSnapshot().root.slides[0].shapes.find((s) => s.kind === "text");
    if (!textShape) throw new Error("expected text shape");

    const { io, stderr } = makeIO();
    const code = await runCli(
      [
        "pptx",
        "set-paragraph-alignment",
        "--file",
        SINGLE,
        "--slide",
        "0",
        "--shape",
        textShape.id,
        "--alignment",
        "left",
        "--clear",
      ],
      io
    );
    expect(code).not.toBe(0);
    expect(stderr.text()).toMatch(/mutually exclusive/);
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
    const textShape = baseAgent.getSnapshot().root.slides[0].shapes.find((s) => s.kind === "text");
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
    const textShape = baseAgent.getSnapshot().root.slides[0].shapes.find((s) => s.kind === "text");
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
    const code = await runCli(["pptx", "apply", "--file", SINGLE, "--commands", cmdsPath, "--out", out], io);
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
          "pptx",
          "table-set-cell-text",
          "--file",
          TABLE,
          "--out",
          out,
          "--slide",
          String(slide),
          "--shape",
          shapeId,
          "--row",
          "0",
          "--column",
          "0",
          "--text",
          "CLI Cell",
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
          "pptx",
          "table-add-row",
          "--file",
          TABLE,
          "--out",
          a,
          "--slide",
          String(beforeTbl.slide),
          "--shape",
          beforeTbl.sh.id,
        ],
        io
      );
      expect(code).toBe(0);

      // IDs are minted from the parser, and adding a row shifts later IDs.
      // Re-resolve the table id from the saved file before the next call.
      const aAgent = await loadDeterministic(a);
      const aTbl = aAgent.getSnapshot().root.slides[beforeTbl.slide].shapes.find((s) => s.kind === "table");
      if (!aTbl) throw new Error("table missing in intermediate file");

      code = await runCli(
        [
          "pptx",
          "table-add-column",
          "--file",
          a,
          "--out",
          b,
          "--slide",
          String(beforeTbl.slide),
          "--shape",
          aTbl.id,
        ],
        io
      );
      expect(code).toBe(0);

      const after = await loadDeterministic(b);
      const tbl = after.getSnapshot().root.slides[beforeTbl.slide].shapes.find((s) => s.kind === "table");
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
          "pptx",
          "table-delete-row",
          "--file",
          TABLE,
          "--out",
          a,
          "--slide",
          String(beforeTbl.slide),
          "--shape",
          beforeTbl.sh.id,
          "--row",
          String(beforeRows - 1),
        ],
        io
      );
      expect(code).toBe(0);

      const aAgent = await loadDeterministic(a);
      const aTbl = aAgent.getSnapshot().root.slides[beforeTbl.slide].shapes.find((s) => s.kind === "table");
      if (!aTbl) throw new Error("table missing in intermediate file");

      code = await runCli(
        [
          "pptx",
          "table-delete-column",
          "--file",
          a,
          "--out",
          b,
          "--slide",
          String(beforeTbl.slide),
          "--shape",
          aTbl.id,
          "--column",
          String(beforeCols - 1),
        ],
        io
      );
      expect(code).toBe(0);

      const after = await loadDeterministic(b);
      const tbl = after.getSnapshot().root.slides[beforeTbl.slide].shapes.find((s) => s.kind === "table");
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

  describe("chart CLI subcommands", () => {
    async function findChart(path: string): Promise<{ slide: number; shapeId: string }> {
      const agent = await loadDeterministic(path);
      const slides = agent.getSnapshot().root.slides;
      for (let i = 0; i < slides.length; i++) {
        const c = slides[i].shapes.find((s) => s.kind === "chart");
        if (c) return { slide: i, shapeId: c.id };
      }
      throw new Error("no chart shape in fixture");
    }

    it("inspect reports chart count and read --format json projects chart data", async () => {
      const { io: io1, stdout: out1 } = makeIO();
      let code = await runCli(["pptx", "inspect", "--file", CHART], io1);
      expect(code).toBe(0);
      const inspected = JSON.parse(out1.text());
      expect(inspected.shapeCounts.chart).toBeGreaterThanOrEqual(1);

      const { io: io2, stdout: out2 } = makeIO();
      code = await runCli(["pptx", "read", "--file", CHART, "--format", "json"], io2);
      expect(code).toBe(0);
      const projection = JSON.parse(out2.text());
      const chart = projection.slides
        .flatMap((s: { shapes: Array<{ kind: string; chart?: unknown }> }) => s.shapes)
        .find((sh: { kind: string }) => sh.kind === "chart");
      expect(chart).toBeDefined();
      expect(chart.chart.partPath).toMatch(/^ppt\/charts\/chart\d+\.xml$/);
      expect(["bar", "line", "area", "pie", "unsupported"]).toContain(chart.chart.chartType);
      expect(Array.isArray(chart.chart.series)).toBe(true);
    });

    it("chart-set-title rewrites the title and chart-set-type swaps the chart kind", async () => {
      const { slide, shapeId } = await findChart(CHART);
      const dir = mkdtempSync(join(tmpdir(), "pptx-cli-chart-"));
      const a = join(dir, "a.pptx");
      const b = join(dir, "b.pptx");
      const { io } = makeIO();

      let code = await runCli(
        [
          "pptx",
          "chart-set-title",
          "--file",
          CHART,
          "--out",
          a,
          "--slide",
          String(slide),
          "--shape",
          shapeId,
          "--title",
          "Quarterly Revenue",
        ],
        io
      );
      expect(code).toBe(0);
      const aAgent = await loadDeterministic(a);
      const aChart = aAgent.getSnapshot().root.slides[slide].shapes.find((s) => s.kind === "chart");
      if (!aChart || aChart.kind !== "chart") throw new Error("chart missing");
      const aPart = aAgent.getSnapshot().root.charts.get(aChart.chartPartPath);
      expect(aPart?.title).toBe("Quarterly Revenue");

      code = await runCli(
        [
          "pptx",
          "chart-set-type",
          "--file",
          a,
          "--out",
          b,
          "--slide",
          String(slide),
          "--shape",
          aChart.id,
          "--type",
          "line",
        ],
        io
      );
      expect(code).toBe(0);
      const bAgent = await loadDeterministic(b);
      const bChart = bAgent.getSnapshot().root.slides[slide].shapes.find((s) => s.kind === "chart");
      if (!bChart || bChart.kind !== "chart") throw new Error("chart missing");
      const bPart = bAgent.getSnapshot().root.charts.get(bChart.chartPartPath);
      expect(bPart?.chartType).toBe("line");
    });

    it("chart-set-data replaces categories + series from a JSON file", async () => {
      const { slide, shapeId } = await findChart(CHART);
      const dir = mkdtempSync(join(tmpdir(), "pptx-cli-chart-data-"));
      const dataPath = join(dir, "data.json");
      const out = join(dir, "out.pptx");
      writeFileSync(
        dataPath,
        JSON.stringify({
          categories: ["Jan", "Feb", "Mar"],
          series: [
            { name: "Revenue", values: [100, 150, 175] },
            { name: "Cost", values: [40, 60, 55] },
          ],
        })
      );
      const { io } = makeIO();
      const code = await runCli(
        [
          "pptx",
          "chart-set-data",
          "--file",
          CHART,
          "--out",
          out,
          "--slide",
          String(slide),
          "--shape",
          shapeId,
          "--data",
          dataPath,
        ],
        io
      );
      expect(code).toBe(0);
      const after = await loadDeterministic(out);
      const ch = after.getSnapshot().root.slides[slide].shapes.find((s) => s.kind === "chart");
      if (!ch || ch.kind !== "chart") throw new Error("chart missing");
      const part = after.getSnapshot().root.charts.get(ch.chartPartPath);
      expect(part?.categories).toEqual(["Jan", "Feb", "Mar"]);
      expect(part?.series.map((s) => s.name)).toEqual(["Revenue", "Cost"]);
      expect(part?.series.map((s) => Array.from(s.values))).toEqual([
        [100, 150, 175],
        [40, 60, 55],
      ]);
    });

    it("chart-set-data validates mismatched categories/series lengths", async () => {
      const { slide, shapeId } = await findChart(CHART);
      const dir = mkdtempSync(join(tmpdir(), "pptx-cli-chart-bad-"));
      const dataPath = join(dir, "bad.json");
      const out = join(dir, "out.pptx");
      writeFileSync(
        dataPath,
        JSON.stringify({
          categories: ["A", "B"],
          series: [{ name: "X", values: [1, 2, 3] }],
        })
      );
      const { io, stderr } = makeIO();
      const code = await runCli(
        [
          "pptx",
          "chart-set-data",
          "--file",
          CHART,
          "--out",
          out,
          "--slide",
          String(slide),
          "--shape",
          shapeId,
          "--data",
          dataPath,
        ],
        io
      );
      expect(code).not.toBe(0);
      expect(stderr.text()).toMatch(/length/);
    });

    it("insert-chart authors a brand-new chart shape backed by an embedded xlsx", async () => {
      const dir = mkdtempSync(join(tmpdir(), "pptx-cli-insert-chart-"));
      const dataPath = join(dir, "data.json");
      const out = join(dir, "out.pptx");
      writeFileSync(
        dataPath,
        JSON.stringify({
          categories: ["Mon", "Tue", "Wed"],
          series: [{ name: "Visits", values: [3, 7, 11] }],
        })
      );
      const { io } = makeIO();
      const code = await runCli(
        [
          "pptx",
          "insert-chart",
          "--file",
          MULTI,
          "--out",
          out,
          "--slide",
          "0",
          "--x",
          "500000",
          "--y",
          "500000",
          "--chart-type",
          "line",
          "--title",
          "Daily traffic",
          "--data",
          dataPath,
        ],
        io
      );
      expect(code).toBe(0);
      const after = await loadDeterministic(out);
      const slide = after.getSnapshot().root.slides[0];
      const ch = slide.shapes.find((s) => s.kind === "chart");
      expect(ch).toBeDefined();
      if (!ch || ch.kind !== "chart") throw new Error("chart missing");
      const part = after.getSnapshot().root.charts.get(ch.chartPartPath);
      expect(part?.chartType).toBe("line");
      expect(part?.title).toBe("Daily traffic");
      expect(part?.categories).toEqual(["Mon", "Tue", "Wed"]);
      expect(part?.series[0]?.values).toEqual([3, 7, 11]);
      // Embedded workbook must exist so PowerPoint's "Edit Data" works.
      expect(part?.embeddingPartPath).toMatch(/Microsoft_Excel_Worksheet\d+\.xlsx$/);
    });

    it("insert-spreadsheet authors a typed OLE-embedded Excel shape", async () => {
      const dir = mkdtempSync(join(tmpdir(), "pptx-cli-insert-ss-"));
      const dataPath = join(dir, "data.json");
      const out = join(dir, "out.pptx");
      writeFileSync(
        dataPath,
        JSON.stringify([
          ["Name", "Score"],
          ["Ada", 91],
          ["Linus", 88],
        ])
      );
      const { io } = makeIO();
      const code = await runCli(
        [
          "pptx",
          "insert-spreadsheet",
          "--file",
          MULTI,
          "--out",
          out,
          "--slide",
          "0",
          "--x",
          "500000",
          "--y",
          "500000",
          "--data",
          dataPath,
        ],
        io
      );
      expect(code).toBe(0);
      const after = await loadDeterministic(out);
      const slide = after.getSnapshot().root.slides[0];
      const ole = slide.shapes.find((s) => s.kind === "ole-spreadsheet");
      expect(ole).toBeDefined();
      if (!ole || ole.kind !== "ole-spreadsheet") throw new Error("OLE shape missing");
      expect(ole.embeddingKind).toBe("xlsx");
      expect(ole.embeddingPartPath).toMatch(/^ppt\/embeddings\//);
      // The workbook bytes are materialised by the serializer, so the
      // round-tripped agent's embeddings map carries real bytes.
      const part = after.getSnapshot().root.embeddings.get(ole.embeddingPartPath);
      expect(part).toBeDefined();
      expect(part?.bytes && part.bytes.byteLength).toBeGreaterThan(0);
    });
  });

  describe("animation CLI subcommands", () => {
    it("inspect reports animation + transition counts and read --format json projects them", async () => {
      const { io, stdout } = makeIO();
      let code = await runCli(["pptx", "inspect", "--file", ANIM], io);
      expect(code).toBe(0);
      const summary = JSON.parse(stdout.text());
      expect(summary.animations).toBeGreaterThan(0);
      expect(summary.transitions).toBeGreaterThan(0);

      const { io: io2, stdout: stdout2 } = makeIO();
      code = await runCli(["pptx", "read", "--file", ANIM, "--format", "json"], io2);
      expect(code).toBe(0);
      const projection = JSON.parse(stdout2.text());
      const slide0 = projection.slides[0];
      expect(slide0.transition?.kind).toBe("fade");
      expect(slide0.animations.length).toBe(2);
      expect(slide0.animations[0]).toMatchObject({
        preset: "appear",
        targetCNvPrId: 2,
        order: 0,
      });
    });

    it("set-slide-transition writes a different transition that survives reload", async () => {
      const dir = mkdtempSync(join(tmpdir(), "pptx-cli-trans-"));
      const out = join(dir, "out.pptx");
      const { io } = makeIO();
      const code = await runCli(
        [
          "pptx",
          "set-slide-transition",
          "--file",
          ANIM,
          "--out",
          out,
          "--slide",
          "0",
          "--kind",
          "wipe",
          "--speed",
          "fast",
        ],
        io
      );
      expect(code).toBe(0);
      const after = await loadDeterministic(out);
      const t = after.getSnapshot().root.slides[0]!.transition!;
      expect(t.kind).toBe("wipe");
      expect(t.speed).toBe("fast");
    });

    it("set-slide-transition --kind none removes an existing transition", async () => {
      const dir = mkdtempSync(join(tmpdir(), "pptx-cli-trans-none-"));
      const out = join(dir, "out.pptx");
      const { io } = makeIO();
      const code = await runCli(
        ["pptx", "set-slide-transition", "--file", ANIM, "--out", out, "--slide", "0", "--kind", "none"],
        io
      );
      expect(code).toBe(0);
      const after = await loadDeterministic(out);
      expect(after.getSnapshot().root.slides[0]!.transition).toBeUndefined();
    });

    it("add/remove/reorder animations roundtrip through CLI invocations", async () => {
      const dir = mkdtempSync(join(tmpdir(), "pptx-cli-anim-"));
      const a = join(dir, "a.pptx");
      const b = join(dir, "b.pptx");
      const c = join(dir, "c.pptx");

      const baseAgent = await loadDeterministic(SINGLE);
      const target = baseAgent.getSnapshot().root.slides[0]!.shapes.find((s) => s.cNvPrId > 0)!;

      const { io } = makeIO();
      let code = await runCli(
        [
          "pptx",
          "add-shape-animation",
          "--file",
          SINGLE,
          "--out",
          a,
          "--slide",
          "0",
          "--shape",
          target.id,
          "--effect",
          "appear",
        ],
        io
      );
      expect(code).toBe(0);
      code = await runCli(
        [
          "pptx",
          "add-shape-animation",
          "--file",
          a,
          "--out",
          b,
          "--slide",
          "0",
          "--shape",
          target.id,
          "--effect",
          "fade",
          "--duration-ms",
          "400",
        ],
        io
      );
      expect(code).toBe(0);
      const afterAdd = await loadDeterministic(b);
      const anims = afterAdd.getSnapshot().root.slides[0]!.animations;
      expect(anims.length).toBe(2);
      expect(anims[0]!.preset).toBe("appear");
      expect(anims[1]!.preset).toBe("fade");
      expect(anims[1]!.durationMs).toBe(400);

      const reverseOrder = [anims[1]!.id, anims[0]!.id].join(",");
      code = await runCli(
        [
          "pptx",
          "reorder-shape-animations",
          "--file",
          b,
          "--out",
          c,
          "--slide",
          "0",
          "--order",
          reverseOrder,
        ],
        io
      );
      expect(code).toBe(0);
      const afterReorder = await loadDeterministic(c);
      const reordered = afterReorder.getSnapshot().root.slides[0]!.animations;
      expect(reordered[0]!.preset).toBe("fade");
      expect(reordered[1]!.preset).toBe("appear");

      const dropId = reordered[0]!.id;
      const d = join(dir, "d.pptx");
      code = await runCli(
        ["pptx", "remove-shape-animation", "--file", c, "--out", d, "--slide", "0", "--animation", dropId],
        io
      );
      expect(code).toBe(0);
      const afterDel = await loadDeterministic(d);
      const left = afterDel.getSnapshot().root.slides[0]!.animations;
      expect(left.length).toBe(1);
      expect(left[0]!.preset).toBe("appear");
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
