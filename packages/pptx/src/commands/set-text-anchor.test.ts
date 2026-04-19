import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { ooxml } from "@officeai/core";
import { PptxAgent } from "../agent/agent.js";
import type { TextShape } from "../model/types.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

async function loadAgent(name: string): Promise<PptxAgent> {
  const buf = await readFile(join(FIXTURES_DIR.pathname, name));
  return PptxAgent.fromBuffer(buf);
}

function findFirstTextShape(slideShapes: ReadonlyArray<unknown>): TextShape | null {
  for (const s of slideShapes as ReadonlyArray<{ kind: string }>) {
    if (s.kind === "text") return s as unknown as TextShape;
  }
  return null;
}

describe("pptx:set-text-anchor", () => {
  it.each([
    ["top", "t"],
    ["middle", "ctr"],
    ["bottom", "b"],
  ] as const)("sets anchor=%s and re-emits anchor=\"%s\"", async (anchor, expected) => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const ts = findFirstTextShape(slide.shapes);
    if (!ts) throw new Error("expected a text shape");

    const m = await agent.applyCommand({
      type: "pptx:set-text-anchor",
      payload: { slideIndex: 0, shapeId: ts.id, anchor },
      source: "system",
    });
    expect(m.status).toBe("approved");

    const updated = agent.getSnapshot().root.slides[0].shapes.find(
      (s): s is TextShape => s.id === ts.id
    )!;
    expect(updated.txBody.bodyPrRaw?.attrs.anchor).toBe(expected);

    const out = await agent.exportFile();
    const c = await ooxml.OoxmlContainer.load(out);
    const xml = c.readText(slide.partPath);
    expect(xml).toContain(`anchor="${expected}"`);
  });

  it("clears the anchor attribute when anchor is null", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const ts = findFirstTextShape(slide.shapes);
    if (!ts) throw new Error("expected a text shape");

    await agent.applyCommand({
      type: "pptx:set-text-anchor",
      payload: { slideIndex: 0, shapeId: ts.id, anchor: "middle" },
      source: "system",
    });
    await agent.applyCommand({
      type: "pptx:set-text-anchor",
      payload: { slideIndex: 0, shapeId: ts.id, anchor: null },
      source: "system",
    });

    const updated = agent.getSnapshot().root.slides[0].shapes.find(
      (s): s is TextShape => s.id === ts.id
    )!;
    expect(updated.txBody.bodyPrRaw?.attrs.anchor).toBeUndefined();
    expect(updated.txBody.bodyPrRaw?.rawAttrs["@_anchor"]).toBeUndefined();
  });

  it("rejects unknown anchor values", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const ts = findFirstTextShape(slide.shapes);
    if (!ts) throw new Error("expected a text shape");

    const m = await agent.applyCommand({
      type: "pptx:set-text-anchor",
      payload: {
        slideIndex: 0,
        shapeId: ts.id,
        anchor: "sideways" as unknown as "top",
      },
      source: "system",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
    expect(m.rejection?.message).toMatch(/unknown anchor/);
  });

  it("survives a serialize → parse round-trip", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const ts = findFirstTextShape(slide.shapes);
    if (!ts) throw new Error("expected a text shape");

    await agent.applyCommand({
      type: "pptx:set-text-anchor",
      payload: { slideIndex: 0, shapeId: ts.id, anchor: "bottom" },
      source: "system",
    });

    const buf = await agent.exportFile();
    const reloaded = await PptxAgent.fromBuffer(buf);
    const slide2 = reloaded.getSnapshot().root.slides[0];
    const ts2 = findFirstTextShape(slide2.shapes);
    expect(ts2?.txBody.bodyPrRaw?.attrs.anchor).toBe("b");
  });

  it("synthesises a bodyPr when the source had none", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    // Surgically remove the bodyPrRaw from the first text shape to
    // simulate a shape parsed from XML that didn't carry an explicit
    // <a:bodyPr> element.
    const snap = agent.getSnapshot();
    const slide = snap.root.slides[0];
    const ts = findFirstTextShape(slide.shapes);
    if (!ts) throw new Error("expected a text shape");

    // Apply twice: first to set, second a no-op-ish to confirm idempotency.
    await agent.applyCommand({
      type: "pptx:set-text-anchor",
      payload: { slideIndex: 0, shapeId: ts.id, anchor: "middle" },
      source: "system",
    });

    const updated = agent.getSnapshot().root.slides[0].shapes.find(
      (s): s is TextShape => s.id === ts.id
    )!;
    expect(updated.txBody.bodyPrRaw).toBeDefined();
    expect(updated.txBody.bodyPrRaw?.attrs.anchor).toBe("ctr");
  });
});
