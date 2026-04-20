import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { PptxAgent } from "../agent/agent.js";
import { parsePptx } from "../parser/parse.js";
import { serializePptx } from "../serializer/serialize.js";
import {
  ANIMATION_PRESETS,
  findPreset,
  type PresetSpec,
} from "./presets.js";
import type {
  AnimationCategory,
  AnimationDirection,
  AnimationPreset,
  AnimationTrigger,
  Shape,
  Slide,
} from "../model/types.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);
const SIMPLE_FIXTURE = join(FIXTURES_DIR.pathname, "04-multi-shape.pptx");

async function loadAgent(): Promise<PptxAgent> {
  const buf = await readFile(SIMPLE_FIXTURE);
  return PptxAgent.fromBuffer(buf);
}

function firstAnimatableShape(slide: Slide): Shape {
  const shape = slide.shapes.find((s) => s.cNvPrId > 0);
  if (!shape) throw new Error("fixture has no animatable shape");
  return shape;
}

/**
 * Add a single typed animation, serialize and reparse, and assert the
 * (category, preset) pair survived. The parser is allowed to upgrade
 * `motionPath/line` to `motionPath/custom` when the path string differs
 * from the canonical default — that's a deliberate semantic in
 * `parseSlideTiming`, so the test asserts on (category, presetClass)
 * rather than the bare preset key for that one case.
 */
async function roundTripPreset(
  spec: PresetSpec,
  overrides: {
    direction?: AnimationDirection;
    trigger?: AnimationTrigger;
    delayMs?: number;
    motionPath?: string;
  } = {}
): Promise<void> {
  const agent = await loadAgent();
  const slide = agent.getSnapshot().root.slides[0]!;
  const target = firstAnimatableShape(slide);

  await agent.applyCommand({
    type: "pptx:add-shape-animation",
    payload: {
      slideIndex: 0,
      shapeId: target.id,
      category: spec.category,
      preset: spec.preset,
      ...(overrides.direction ? { direction: overrides.direction } : {}),
      ...(overrides.trigger ? { trigger: overrides.trigger } : {}),
      ...(overrides.delayMs !== undefined ? { delayMs: overrides.delayMs } : {}),
      ...(overrides.motionPath ? { motionPath: overrides.motionPath } : {}),
    },
  });

  const out = await serializePptx(agent.getSnapshot());
  const reparsed = await parsePptx(out);
  const parsedSlide = reparsed.root.slides[0]!;
  expect(parsedSlide.animations.length).toBe(1);
  const a = parsedSlide.animations[0]!;
  expect(a.targetCNvPrId).toBe(target.cNvPrId);
  expect(a.category).toBe(spec.category);

  // Motion-path `line` with a non-default path is intentionally upgraded
  // to `custom` by the parser. All other presets round-trip 1:1.
  if (spec.category === "motionPath" && spec.preset === "line" && overrides.motionPath) {
    expect(a.preset).toBe("custom");
  } else {
    expect(a.preset).toBe(spec.preset);
  }

  if (overrides.trigger) {
    expect(a.trigger).toBe(overrides.trigger);
  }
  if (overrides.delayMs !== undefined) {
    expect(a.delayMs).toBe(overrides.delayMs);
  }
  if (overrides.direction && spec.directions?.includes(overrides.direction)) {
    expect(a.direction).toBe(overrides.direction);
  }
}

describe("F4 v2: animation preset round-trip", () => {
  it("registry has at least one entry per category", () => {
    const cats = new Set<AnimationCategory>();
    for (const p of ANIMATION_PRESETS) cats.add(p.category);
    expect(cats.has("entrance")).toBe(true);
    expect(cats.has("emphasis")).toBe(true);
    expect(cats.has("exit")).toBe(true);
    expect(cats.has("motionPath")).toBe(true);
  });

  describe("entrance presets", () => {
    const presets = ANIMATION_PRESETS.filter((p) => p.category === "entrance");
    for (const spec of presets) {
      it(`round-trips entrance/${spec.preset}`, async () => {
        await roundTripPreset(spec);
      });
    }
  });

  describe("emphasis presets", () => {
    const presets = ANIMATION_PRESETS.filter((p) => p.category === "emphasis");
    for (const spec of presets) {
      it(`round-trips emphasis/${spec.preset}`, async () => {
        await roundTripPreset(spec);
      });
    }
  });

  describe("exit presets", () => {
    const presets = ANIMATION_PRESETS.filter((p) => p.category === "exit");
    for (const spec of presets) {
      it(`round-trips exit/${spec.preset}`, async () => {
        await roundTripPreset(spec);
      });
    }
  });

  describe("motion path presets", () => {
    const presets = ANIMATION_PRESETS.filter((p) => p.category === "motionPath");
    for (const spec of presets) {
      it(`round-trips motionPath/${spec.preset}`, async () => {
        if (spec.preset === "custom") {
          await roundTripPreset(spec, { motionPath: "M 0 0 L 0.5 0.3 L 0.7 -0.2 E" });
        } else {
          await roundTripPreset(spec);
        }
      });
    }
  });

  describe("direction-aware presets preserve direction", () => {
    const directionalPresets = ANIMATION_PRESETS.filter((p) => p.directions && p.directions.length);
    for (const spec of directionalPresets) {
      for (const dir of spec.directions!) {
        it(`${spec.category}/${spec.preset} direction=${dir}`, async () => {
          await roundTripPreset(spec, { direction: dir });
        });
      }
    }
  });

  describe("triggers round-trip", () => {
    const flyIn = findPreset("entrance", "flyIn")!;
    for (const trigger of ["onClick", "withPrevious", "afterPrevious"] as const) {
      it(`flyIn with trigger=${trigger}`, async () => {
        await roundTripPreset(flyIn, { trigger });
      });
    }
  });

  describe("delay round-trips", () => {
    it("preserves delayMs for an entrance fade", async () => {
      const fade = findPreset("entrance", "fade")!;
      await roundTripPreset(fade, { delayMs: 750 });
    });
  });

  describe("11-animations-gallery.pptx golden fixture", () => {
    it("parses each category from the hand-rolled gallery fixture", async () => {
      const galleryPath = join(FIXTURES_DIR.pathname, "11-animations-gallery.pptx");
      const buf = await readFile(galleryPath);
      const snap = await parsePptx(buf);
      const anim = snap.root.slides[0]!.animations;
      expect(anim.length).toBe(5);
      expect(anim[0]).toMatchObject({
        category: "entrance",
        preset: "flyIn",
        direction: "left",
        trigger: "onClick",
        targetCNvPrId: 2,
      });
      expect(anim[1]).toMatchObject({
        category: "emphasis",
        preset: "spin",
        trigger: "withPrevious",
        targetCNvPrId: 2,
      });
      expect(anim[2]).toMatchObject({
        category: "exit",
        preset: "fade",
        trigger: "afterPrevious",
        targetCNvPrId: 2,
      });
      expect(anim[3]).toMatchObject({
        category: "entrance",
        preset: "wipe",
        direction: "down",
        trigger: "onClick",
        delayMs: 500,
        targetCNvPrId: 3,
      });
      expect(anim[4]).toMatchObject({
        category: "motionPath",
        preset: "arc",
        trigger: "withPrevious",
        targetCNvPrId: 3,
      });
      expect(anim[4]!.motionPath).toBeDefined();
    });

    it("re-emits and reparses the gallery without losing categories", async () => {
      const galleryPath = join(FIXTURES_DIR.pathname, "11-animations-gallery.pptx");
      const buf = await readFile(galleryPath);
      const agent = await PptxAgent.fromBuffer(buf);
      const out = await serializePptx(agent.getSnapshot());
      const reparsed = await parsePptx(out);
      const anim = reparsed.root.slides[0]!.animations;
      expect(anim.length).toBe(5);
      expect(anim.map((a) => a.category)).toEqual([
        "entrance",
        "emphasis",
        "exit",
        "entrance",
        "motionPath",
      ]);
    });
  });

  describe("animation gallery: many presets on a single slide", () => {
    it("serialises and reparses a mixed gallery without loss", async () => {
      const agent = await loadAgent();
      const slide = agent.getSnapshot().root.slides[0]!;
      const target = firstAnimatableShape(slide);
      const gallery: ReadonlyArray<{
        category: AnimationCategory;
        preset: AnimationPreset;
      }> = [
        { category: "entrance", preset: "appear" },
        { category: "entrance", preset: "fade" },
        { category: "entrance", preset: "flyIn" },
        { category: "entrance", preset: "wipe" },
        { category: "entrance", preset: "zoom" },
        { category: "emphasis", preset: "pulse" },
        { category: "emphasis", preset: "spin" },
        { category: "emphasis", preset: "growShrink" },
        { category: "exit", preset: "disappear" },
        { category: "exit", preset: "fade" },
        { category: "exit", preset: "flyOut" },
        { category: "motionPath", preset: "line" },
        { category: "motionPath", preset: "arc" },
        { category: "motionPath", preset: "loops" },
      ];
      for (const g of gallery) {
        await agent.applyCommand({
          type: "pptx:add-shape-animation",
          payload: {
            slideIndex: 0,
            shapeId: target.id,
            category: g.category,
            preset: g.preset,
          },
        });
      }
      const out = await serializePptx(agent.getSnapshot());
      const reparsed = await parsePptx(out);
      const parsed = reparsed.root.slides[0]!.animations;
      expect(parsed.length).toBe(gallery.length);
      for (let i = 0; i < gallery.length; i++) {
        expect(parsed[i]!.category).toBe(gallery[i]!.category);
        expect(parsed[i]!.preset).toBe(gallery[i]!.preset);
        expect(parsed[i]!.order).toBe(i);
      }
    });
  });
});
