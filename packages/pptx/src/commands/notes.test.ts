import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { ooxml } from "@officeai/core";
import { PptxAgent } from "../agent/agent.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

async function loadAgent(name: string): Promise<PptxAgent> {
  const buf = await readFile(join(FIXTURES_DIR.pathname, name));
  return PptxAgent.fromBuffer(buf);
}

function bodyAsText(body: { paragraphs: ReadonlyArray<{ runs: ReadonlyArray<{ text: string }> }> }): string {
  return body.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join("\n");
}

describe("pptx:set-slide-notes (creates notes part on demand)", () => {
  it("writes the body text into the notes part (creating one if needed)", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const m = await agent.applyCommand({
      type: "pptx:set-slide-notes",
      payload: { slideIndex: 0, text: "Talk track:\nBe concise." },
      source: "human",
    });
    expect(m.status).toBe("approved");

    const after = agent.getSnapshot();
    const newSlide = after.root.slides[0];
    expect(newSlide.notesSlidePartPath).toBeTruthy();
    const notes = after.root.notesSlides.get(newSlide.notesSlidePartPath!);
    expect(notes).toBeTruthy();
    expect(bodyAsText(notes!.body)).toBe("Talk track:\nBe concise.");
  });

  it("rewrites the body in-place on subsequent calls without minting a new part", async () => {
    const agent = await loadAgent("01-blank.pptx");
    await agent.applyCommand({
      type: "pptx:set-slide-notes",
      payload: { slideIndex: 0, text: "first" },
      source: "human",
    });
    const path1 = agent.getSnapshot().root.slides[0].notesSlidePartPath;
    expect(path1).toBeTruthy();
    await agent.applyCommand({
      type: "pptx:set-slide-notes",
      payload: { slideIndex: 0, text: "second" },
      source: "human",
    });
    const after = agent.getSnapshot();
    const path2 = after.root.slides[0].notesSlidePartPath;
    expect(path2).toBe(path1);
    const notes = after.root.notesSlides.get(path2!)!;
    expect(bodyAsText(notes.body)).toBe("second");
  });

  it("survives serialize → re-parse with the notes body intact", async () => {
    const agent = await loadAgent("01-blank.pptx");
    await agent.applyCommand({
      type: "pptx:set-slide-notes",
      payload: { slideIndex: 0, text: "Reminder:\nbreathe." },
      source: "human",
    });
    const out = await agent.exportFile();
    // Sanity: container has the notes part registered.
    const c = await ooxml.OoxmlContainer.load(out);
    const path = agent.getSnapshot().root.slides[0].notesSlidePartPath!;
    expect(c.has(path)).toBe(true);

    const reparsed = await PptxAgent.fromBuffer(out);
    const slide = reparsed.getSnapshot().root.slides[0];
    expect(slide.notesSlidePartPath).toBe(path);
    const notes = reparsed.getSnapshot().root.notesSlides.get(slide.notesSlidePartPath!)!;
    expect(bodyAsText(notes.body)).toBe("Reminder:\nbreathe.");
  });

  it("rejects non-string text payloads", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const memo = await agent.applyCommand({
      type: "pptx:set-slide-notes",
      payload: { slideIndex: 0, text: 42 as unknown as string },
      source: "human",
    });
    expect(memo.status).toBe("rejected");
    if (memo.status === "rejected") {
      expect(memo.rejection.code).toBe("invalid-payload");
    }
  });
});
