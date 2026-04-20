import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { PptxAgent } from "../agent/agent.js";
import { parsePptx } from "../parser/parse.js";
import { serializePptx } from "../serializer/serialize.js";
import type { Shape, Slide } from "../model/types.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);
const ANIM_FIXTURE = join(FIXTURES_DIR.pathname, "10-with-anim.pptx");
const SIMPLE_FIXTURE = join(FIXTURES_DIR.pathname, "04-multi-shape.pptx");

async function loadAgent(path: string): Promise<PptxAgent> {
  const buf = await readFile(path);
  return PptxAgent.fromBuffer(buf);
}

function firstSlide(agent: PptxAgent): Slide {
  return agent.getSnapshot().root.slides[0]!;
}

describe("F4: pptx:set-slide-transition", () => {
  it("replaces an existing transition kind+speed (delta-merging the raw blob)", async () => {
    const agent = await loadAgent(ANIM_FIXTURE);
    const before = firstSlide(agent).transition;
    expect(before?.kind).toBe("fade");
    await agent.applyCommand({
      type: "pptx:set-slide-transition",
      payload: { slideIndex: 0, kind: "push", speed: "fast" },
    });
    const after = firstSlide(agent).transition!;
    expect(after.kind).toBe("push");
    expect(after.speed).toBe("fast");
    // Delta-merge: raw is patched in-place, not dropped, so unmodeled
    // attrs/children survive. The patched raw must reflect the new
    // kind tag and the new speed attribute.
    expect(after.raw).toBeDefined();
    expect(after.raw!.attrs.spd).toBe("fast");
    const childTags = (after.raw!.subtree as ReadonlyArray<Record<string, unknown>>).map(
      (n) => Object.keys(n).find((k) => !k.startsWith(":") && !k.startsWith("#"))
    );
    expect(childTags).toContain("p:push");
    expect(childTags).not.toContain("p:fade");
  });

  it("removes the transition when kind=none", async () => {
    const agent = await loadAgent(ANIM_FIXTURE);
    await agent.applyCommand({
      type: "pptx:set-slide-transition",
      payload: { slideIndex: 0, kind: "none" },
    });
    expect(firstSlide(agent).transition).toBeUndefined();
  });

  it("rejects no-op when kind+speed already match", async () => {
    const agent = await loadAgent(ANIM_FIXTURE);
    const m = await agent.applyCommand({
      type: "pptx:set-slide-transition",
      payload: { slideIndex: 0, kind: "fade", speed: "med" },
    });
    expect(m.rejection?.code).toBe("no-op");
  });

  it("rejects no-op when removing a transition that does not exist", async () => {
    const agent = await loadAgent(SIMPLE_FIXTURE);
    const m = await agent.applyCommand({
      type: "pptx:set-slide-transition",
      payload: { slideIndex: 0, kind: "none" },
    });
    expect(m.rejection?.code).toBe("no-op");
  });

  it("survives roundtrip when added to a slide that had no transition", async () => {
    const agent = await loadAgent(SIMPLE_FIXTURE);
    await agent.applyCommand({
      type: "pptx:set-slide-transition",
      payload: { slideIndex: 0, kind: "wipe", speed: "slow" },
    });
    const out = await serializePptx(agent.getSnapshot());
    const reparsed = await parsePptx(out);
    expect(reparsed.root.slides[0]!.transition?.kind).toBe("wipe");
    expect(reparsed.root.slides[0]!.transition?.speed).toBe("slow");
  });
});

describe("F4: pptx:add-shape-animation / remove / reorder", () => {
  it("adds an entrance animation targeting the shape's cNvPrId", async () => {
    const agent = await loadAgent(SIMPLE_FIXTURE);
    const slide = firstSlide(agent);
    const target = slide.shapes.find((s: Shape) => s.cNvPrId > 0)!;
    expect(slide.animations.length).toBe(0);
    await agent.applyCommand({
      type: "pptx:add-shape-animation",
      payload: { slideIndex: 0, shapeId: target.id, effect: "fly-in", durationMs: 750 },
    });
    const after = firstSlide(agent);
    expect(after.animations.length).toBe(1);
    expect(after.animations[0]).toMatchObject({
      category: "entrance",
      preset: "flyIn",
      targetCNvPrId: target.cNvPrId,
      durationMs: 750,
      order: 0,
    });
  });

  it("removes an animation by id and renumbers remaining order", async () => {
    const agent = await loadAgent(ANIM_FIXTURE);
    const slide = firstSlide(agent);
    expect(slide.animations.length).toBe(2);
    const drop = slide.animations[0]!;
    await agent.applyCommand({
      type: "pptx:remove-shape-animation",
      payload: { slideIndex: 0, animationId: drop.id },
    });
    const after = firstSlide(agent);
    expect(after.animations.length).toBe(1);
    expect(after.animations[0]!.order).toBe(0);
    expect(after.animations[0]!.targetCNvPrId).toBe(slide.animations[1]!.targetCNvPrId);
  });

  it("reorders animations atomically as a permutation", async () => {
    const agent = await loadAgent(ANIM_FIXTURE);
    const slide = firstSlide(agent);
    const ids = slide.animations.map((a) => a.id);
    const reversed = [...ids].reverse();
    await agent.applyCommand({
      type: "pptx:reorder-shape-animations",
      payload: { slideIndex: 0, order: reversed },
    });
    const after = firstSlide(agent);
    expect(after.animations.map((a) => a.id)).toEqual(reversed);
    expect(after.animations.map((a) => a.order)).toEqual([0, 1]);
  });

  it("rejects reorder when ids don't match the current set", async () => {
    const agent = await loadAgent(ANIM_FIXTURE);
    const m = await agent.applyCommand({
      type: "pptx:reorder-shape-animations",
      payload: { slideIndex: 0, order: ["bogus-1", "bogus-2"] },
    });
    expect(m.rejection?.code).toBe("invalid-payload");
  });

  it("rejects remove with unknown-target when id is not present", async () => {
    const agent = await loadAgent(ANIM_FIXTURE);
    const m = await agent.applyCommand({
      type: "pptx:remove-shape-animation",
      payload: { slideIndex: 0, animationId: "nope" },
    });
    expect(m.rejection?.code).toBe("unknown-target");
  });

  it("rebuilds <p:timing> from typed animations after edits (roundtrip)", async () => {
    const agent = await loadAgent(SIMPLE_FIXTURE);
    const target = firstSlide(agent).shapes.find((s: Shape) => s.cNvPrId > 0)!;
    await agent.applyCommand({
      type: "pptx:add-shape-animation",
      payload: { slideIndex: 0, shapeId: target.id, effect: "appear" },
    });
    await agent.applyCommand({
      type: "pptx:add-shape-animation",
      payload: { slideIndex: 0, shapeId: target.id, effect: "fade", durationMs: 400 },
    });
    const out = await serializePptx(agent.getSnapshot());
    const reparsed = await parsePptx(out);
    const slide2 = reparsed.root.slides[0]!;
    expect(slide2.animations.length).toBe(2);
    expect(slide2.animations[0]!.preset).toBe("appear");
    expect(slide2.animations[0]!.category).toBe("entrance");
    expect(slide2.animations[1]!).toMatchObject({
      category: "entrance",
      preset: "fade",
      durationMs: 400,
      targetCNvPrId: target.cNvPrId,
    });
  });

  it("removing the last animation drops <p:timing> entirely on roundtrip", async () => {
    const agent = await loadAgent(ANIM_FIXTURE);
    for (const a of [...firstSlide(agent).animations]) {
      await agent.applyCommand({
        type: "pptx:remove-shape-animation",
        payload: { slideIndex: 0, animationId: a.id },
      });
    }
    const out = await serializePptx(agent.getSnapshot());
    const reparsed = await parsePptx(out);
    expect(reparsed.root.slides[0]!.animations.length).toBe(0);
    expect(reparsed.root.slides[0]!.timingTailRaw).toBeUndefined();
  });
});
