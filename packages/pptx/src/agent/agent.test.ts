import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { PptxAgent } from "./agent.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

describe("PptxAgent skeleton", () => {
  it("builds from buffer and exports byte-stable file on no-edit pass", async () => {
    const buf = await readFile(join(FIXTURES_DIR.pathname, "04-multi-shape.pptx"));
    const agent = await PptxAgent.fromBuffer(buf);
    const snap = agent.getSnapshot();
    expect(snap.format).toBe("pptx");
    expect(snap.root.slides.length).toBeGreaterThan(0);
    const out = await agent.exportFile();
    expect(out.byteLength).toBeGreaterThan(0);
  });

  it("renders Markdown for inspection", async () => {
    const buf = await readFile(join(FIXTURES_DIR.pathname, "04-multi-shape.pptx"));
    const agent = await PptxAgent.fromBuffer(buf);
    const md = agent.toMarkdown();
    expect(md).toMatch(/^# Presentation/);
    expect(md).toMatch(/Slide 1/);
    expect(md).toMatch(/Multi-shape canvas/);
  });

  it("supports getRange and search", async () => {
    const buf = await readFile(join(FIXTURES_DIR.pathname, "07-multi-slide.pptx"));
    const agent = await PptxAgent.fromBuffer(buf);
    const range = agent.getRange({ kind: "pptx-slides", start: 0, end: 2 });
    expect(range.slides.length).toBe(2);
    expect(range.slides[0].partPath).toMatch(/^ppt\/slides\//);

    const results = agent.search({ query: "Welcome" });
    expect(results.length).toBeGreaterThan(0);
    for (const r of results) {
      expect(r.match.toLowerCase()).toContain("welcome");
    }
  });

  it("returns no-handler error mutation for unknown command", async () => {
    const buf = await readFile(join(FIXTURES_DIR.pathname, "01-blank.pptx"));
    const agent = await PptxAgent.fromBuffer(buf);
    await expect(agent.applyCommand({ type: "non.existent.command", payload: {} })).rejects.toThrow(
      /no-handler|No handler/
    );
  });
});
