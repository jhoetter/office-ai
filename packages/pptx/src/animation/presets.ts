/**
 * F4 v2 — Centralised PowerPoint animation preset registry.
 *
 * Single source of truth for the four animation categories (Entrance,
 * Emphasis, Exit and Motion Paths). The parser, serializer, picker UI
 * and (Phase 4) the playback engine all consume this table so adding
 * a new preset is a one-place change.
 *
 * Preset IDs and classes follow the OOXML spec (ECMA-376 Part 1
 * §19.5 / [MS-OI29500]). The picked subset matches PowerPoint's default
 * gallery for each category — anything outside this set still parses
 * (recognised by `presetClass` alone, dropped into `timingTailRaw`)
 * but does not round-trip through the typed model.
 *
 * Each spec carries:
 *   - `category`        – which gallery section it belongs to
 *   - `preset`          – typed key
 *   - `presetClass`     – OOXML `<p:cTn @presetClass>` (entr|emph|exit|path)
 *   - `presetId`        – OOXML `<p:cTn @presetID>`
 *   - `defaultSubtype`  – OOXML `<p:cTn @presetSubtype>` baseline; mutated
 *                          per direction by `subtypeFor()` for direction-aware
 *                          presets so the playback target survives roundtrip
 *   - `defaultDurationMs` / `directions` – picker defaults + Effect-options
 *   - `emitBody`        – builds the per-preset child of `<p:par>`'s `<p:cTn>`
 *
 * The `<p:par><p:cTn>` envelope (id, presetID, presetClass, nodeType,
 * delay, dur, fill) is built by the serializer; emitters only synthesise
 * the *body* (typically a `<p:childTnLst>` containing one or more of
 * `<p:set>`, `<p:anim>`, `<p:animEffect>`, `<p:animRot>`, `<p:animScale>`,
 * `<p:animMotion>`).
 */

import type {
  AnimationCategory,
  AnimationDirection,
  AnimationPreset,
  EmphasisPreset,
  EntrancePreset,
  ExitPreset,
  MotionPathPreset,
  ShapeAnimation,
} from "../model/types.js";

export interface PresetSpec {
  readonly category: AnimationCategory;
  readonly preset: AnimationPreset;
  readonly presetClass: "entr" | "emph" | "exit" | "path";
  readonly presetId: number;
  readonly defaultSubtype: number;
  readonly defaultDurationMs: number;
  readonly directions?: ReadonlyArray<AnimationDirection>;
  /** Defaults injected when the user picks the preset from the gallery. */
  readonly defaultDirection?: AnimationDirection;
  /** Builds the OOXML body emitted under `<p:par><p:cTn>...<p:childTnLst>`. */
  readonly emitBody: (a: ShapeAnimation, helpers: EmitHelpers) => unknown;
  /** Translation hint for the picker and badges. */
  readonly i18nKey: string;
}

export interface EmitHelpers {
  /** Build a generic `<p:set>` setting an attribute to a value. */
  readonly setAttr: (
    spid: number,
    attrName: string,
    value: string,
    durMs: number
  ) => Record<string, unknown>;
  /** Build a `<p:anim>` interpolating one attribute. */
  readonly anim: (
    spid: number,
    attrName: string,
    durMs: number,
    fromTo: { from?: string; to: string },
    extra?: Record<string, string>
  ) => Record<string, unknown>;
  /** Build a `<p:animEffect>` filter (used for `wipe`, `circle`, `blinds`, …). */
  readonly animEffect: (
    spid: number,
    durMs: number,
    transition: "in" | "out",
    filter: string,
    extra?: Record<string, string>
  ) => Record<string, unknown>;
  /** Build a `<p:animRot>` for spin/swivel. */
  readonly animRot: (
    spid: number,
    durMs: number,
    fromDeg: number,
    toDeg: number,
    extra?: Record<string, string>
  ) => Record<string, unknown>;
  /** Build a `<p:animScale>` for grow/shrink/zoom. */
  readonly animScale: (
    spid: number,
    durMs: number,
    from: { x: number; y: number },
    to: { x: number; y: number },
    extra?: Record<string, string>
  ) => Record<string, unknown>;
  /** Build a `<p:animMotion>` for motion paths. */
  readonly animMotion: (
    spid: number,
    durMs: number,
    path: string,
    extra?: Record<string, string>
  ) => Record<string, unknown>;
  /** Wrap several behavior nodes in a single `<p:childTnLst>`. */
  readonly childTnLst: (children: ReadonlyArray<unknown>) => Record<string, unknown>;
}

// ─── Entrance presets ────────────────────────────────────────────────────

const ENTRANCE_PRESETS: ReadonlyArray<PresetSpec> = [
  {
    category: "entrance",
    preset: "appear",
    presetClass: "entr",
    presetId: 1,
    defaultSubtype: 0,
    defaultDurationMs: 1,
    i18nKey: "entranceAppear",
    emitBody: (a, h) =>
      h.childTnLst([
        h.setAttr(a.targetCNvPrId, "style.visibility", "visible", a.durationMs ?? 1),
      ]),
  },
  {
    category: "entrance",
    preset: "flyIn",
    presetClass: "entr",
    presetId: 2,
    defaultSubtype: 4,
    defaultDurationMs: 500,
    defaultDirection: "left",
    directions: ["left", "right", "up", "down"],
    i18nKey: "entranceFlyIn",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 500;
      const dir = a.direction ?? "left";
      const { axisAttr, from, to } = flyInVector(dir);
      return h.childTnLst([
        h.setAttr(a.targetCNvPrId, "style.visibility", "visible", 1),
        h.anim(a.targetCNvPrId, axisAttr, dur, { from, to }),
      ]);
    },
  },
  {
    category: "entrance",
    preset: "fade",
    presetClass: "entr",
    presetId: 3,
    defaultSubtype: 0,
    defaultDurationMs: 500,
    i18nKey: "entranceFade",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 500;
      return h.childTnLst([
        h.setAttr(a.targetCNvPrId, "style.visibility", "visible", 1),
        h.anim(a.targetCNvPrId, "style.opacity", dur, { from: "0", to: "1" }, { calcmode: "lin" }),
      ]);
    },
  },
  {
    category: "entrance",
    preset: "split",
    presetClass: "entr",
    presetId: 4,
    defaultSubtype: 26, // horizontalIn
    defaultDurationMs: 500,
    defaultDirection: "horizontal",
    directions: ["horizontal", "vertical"],
    i18nKey: "entranceSplit",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 500;
      const dir = a.direction ?? "horizontal";
      const filter = dir === "vertical" ? "barn(inVertical)" : "barn(inHorizontal)";
      return h.childTnLst([
        h.setAttr(a.targetCNvPrId, "style.visibility", "visible", 1),
        h.animEffect(a.targetCNvPrId, dur, "in", filter),
      ]);
    },
  },
  {
    category: "entrance",
    preset: "wipe",
    presetClass: "entr",
    presetId: 10,
    defaultSubtype: 8, // fromBottom (matches PowerPoint default)
    defaultDurationMs: 500,
    defaultDirection: "up",
    directions: ["up", "down", "left", "right"],
    i18nKey: "entranceWipe",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 500;
      const dir = a.direction ?? "up";
      const filter = `wipe(${wipeDirection(dir)})`;
      return h.childTnLst([
        h.setAttr(a.targetCNvPrId, "style.visibility", "visible", 1),
        h.animEffect(a.targetCNvPrId, dur, "in", filter),
      ]);
    },
  },
  {
    category: "entrance",
    preset: "shape",
    presetClass: "entr",
    presetId: 5,
    defaultSubtype: 16, // in (circle)
    defaultDurationMs: 500,
    defaultDirection: "in",
    directions: ["in", "out"],
    i18nKey: "entranceShape",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 500;
      const dir = a.direction === "out" ? "out" : "in";
      return h.childTnLst([
        h.setAttr(a.targetCNvPrId, "style.visibility", "visible", 1),
        h.animEffect(a.targetCNvPrId, dur, "in", `circle(${dir})`),
      ]);
    },
  },
  {
    category: "entrance",
    preset: "wheel",
    presetClass: "entr",
    presetId: 21,
    defaultSubtype: 1,
    defaultDurationMs: 1000,
    i18nKey: "entranceWheel",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 1000;
      return h.childTnLst([
        h.setAttr(a.targetCNvPrId, "style.visibility", "visible", 1),
        h.animEffect(a.targetCNvPrId, dur, "in", "wheel(1)"),
      ]);
    },
  },
  {
    category: "entrance",
    preset: "randomBars",
    presetClass: "entr",
    presetId: 11,
    defaultSubtype: 10, // horizontal
    defaultDurationMs: 500,
    defaultDirection: "horizontal",
    directions: ["horizontal", "vertical"],
    i18nKey: "entranceRandomBars",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 500;
      const dir = a.direction ?? "horizontal";
      return h.childTnLst([
        h.setAttr(a.targetCNvPrId, "style.visibility", "visible", 1),
        h.animEffect(
          a.targetCNvPrId,
          dur,
          "in",
          `randombar(${dir === "vertical" ? "vertical" : "horizontal"})`
        ),
      ]);
    },
  },
  {
    category: "entrance",
    preset: "floatIn",
    presetClass: "entr",
    presetId: 42,
    defaultSubtype: 8, // up
    defaultDurationMs: 500,
    defaultDirection: "up",
    directions: ["up", "down"],
    i18nKey: "entranceFloatIn",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 500;
      const dir = a.direction ?? "up";
      const offset = dir === "down" ? "-#ppt_h/8" : "#ppt_h/8";
      return h.childTnLst([
        h.setAttr(a.targetCNvPrId, "style.visibility", "visible", 1),
        h.anim(a.targetCNvPrId, "style.opacity", dur, { from: "0", to: "1" }, { calcmode: "lin" }),
        h.anim(a.targetCNvPrId, "ppt_y", dur, { from: `#ppt_y+${offset}`, to: "#ppt_y" }),
      ]);
    },
  },
  {
    category: "entrance",
    preset: "growAndTurn",
    presetClass: "entr",
    presetId: 6,
    defaultSubtype: 0,
    defaultDurationMs: 1000,
    i18nKey: "entranceGrowAndTurn",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 1000;
      return h.childTnLst([
        h.setAttr(a.targetCNvPrId, "style.visibility", "visible", 1),
        h.animScale(a.targetCNvPrId, dur, { x: 0, y: 0 }, { x: 100, y: 100 }),
        h.animRot(a.targetCNvPrId, dur, -180, 0),
      ]);
    },
  },
  {
    category: "entrance",
    preset: "zoom",
    presetClass: "entr",
    presetId: 23,
    defaultSubtype: 0,
    defaultDurationMs: 500,
    i18nKey: "entranceZoom",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 500;
      return h.childTnLst([
        h.setAttr(a.targetCNvPrId, "style.visibility", "visible", 1),
        h.animScale(a.targetCNvPrId, dur, { x: 0, y: 0 }, { x: 100, y: 100 }),
        h.anim(a.targetCNvPrId, "style.opacity", dur, { from: "0", to: "1" }, { calcmode: "lin" }),
      ]);
    },
  },
  {
    category: "entrance",
    preset: "swivel",
    presetClass: "entr",
    presetId: 18,
    defaultSubtype: 0,
    defaultDurationMs: 1000,
    i18nKey: "entranceSwivel",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 1000;
      return h.childTnLst([
        h.setAttr(a.targetCNvPrId, "style.visibility", "visible", 1),
        h.animRot(a.targetCNvPrId, dur, -180, 0),
      ]);
    },
  },
  {
    category: "entrance",
    preset: "bounce",
    presetClass: "entr",
    presetId: 26,
    defaultSubtype: 0,
    defaultDurationMs: 1500,
    i18nKey: "entranceBounce",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 1500;
      return h.childTnLst([
        h.setAttr(a.targetCNvPrId, "style.visibility", "visible", 1),
        h.anim(
          a.targetCNvPrId,
          "ppt_y",
          dur,
          { from: "#ppt_y-#ppt_h", to: "#ppt_y" },
          { calcmode: "lin" }
        ),
      ]);
    },
  },
];

// ─── Emphasis presets ────────────────────────────────────────────────────

const EMPHASIS_PRESETS: ReadonlyArray<PresetSpec> = [
  {
    category: "emphasis",
    preset: "pulse",
    presetClass: "emph",
    presetId: 1,
    defaultSubtype: 0,
    defaultDurationMs: 700,
    i18nKey: "emphasisPulse",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 700;
      const half = Math.max(1, Math.round(dur / 2));
      return h.childTnLst([
        h.animScale(a.targetCNvPrId, half, { x: 100, y: 100 }, { x: 110, y: 110 }),
        h.animScale(a.targetCNvPrId, half, { x: 110, y: 110 }, { x: 100, y: 100 }),
      ]);
    },
  },
  {
    category: "emphasis",
    preset: "colorPulse",
    presetClass: "emph",
    presetId: 2,
    defaultSubtype: 0,
    defaultDurationMs: 700,
    i18nKey: "emphasisColorPulse",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 700;
      return h.childTnLst([
        h.anim(
          a.targetCNvPrId,
          "fillcolor",
          dur,
          { from: "#fillColor", to: "#fillColor" },
          { calcmode: "lin" }
        ),
      ]);
    },
  },
  {
    category: "emphasis",
    preset: "teeter",
    presetClass: "emph",
    presetId: 14,
    defaultSubtype: 0,
    defaultDurationMs: 1000,
    i18nKey: "emphasisTeeter",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 1000;
      const seg = Math.max(1, Math.round(dur / 4));
      return h.childTnLst([
        h.animRot(a.targetCNvPrId, seg, 0, 6),
        h.animRot(a.targetCNvPrId, seg, 6, -6),
        h.animRot(a.targetCNvPrId, seg, -6, 6),
        h.animRot(a.targetCNvPrId, seg, 6, 0),
      ]);
    },
  },
  {
    category: "emphasis",
    preset: "spin",
    presetClass: "emph",
    presetId: 8,
    defaultSubtype: 0,
    defaultDurationMs: 1500,
    defaultDirection: "clockwise",
    directions: ["clockwise", "counterclockwise"],
    i18nKey: "emphasisSpin",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 1500;
      const sign = a.direction === "counterclockwise" ? -1 : 1;
      return h.childTnLst([h.animRot(a.targetCNvPrId, dur, 0, 360 * sign)]);
    },
  },
  {
    category: "emphasis",
    preset: "growShrink",
    presetClass: "emph",
    presetId: 6,
    defaultSubtype: 0,
    defaultDurationMs: 1000,
    i18nKey: "emphasisGrowShrink",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 1000;
      const half = Math.max(1, Math.round(dur / 2));
      return h.childTnLst([
        h.animScale(a.targetCNvPrId, half, { x: 100, y: 100 }, { x: 150, y: 150 }),
        h.animScale(a.targetCNvPrId, half, { x: 150, y: 150 }, { x: 100, y: 100 }),
      ]);
    },
  },
  {
    category: "emphasis",
    preset: "desaturate",
    presetClass: "emph",
    presetId: 13,
    defaultSubtype: 0,
    defaultDurationMs: 700,
    i18nKey: "emphasisDesaturate",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 700;
      return h.childTnLst([
        h.anim(
          a.targetCNvPrId,
          "style.color",
          dur,
          { from: "#fillColor", to: "#888888" },
          { calcmode: "lin" }
        ),
      ]);
    },
  },
  {
    category: "emphasis",
    preset: "fontColor",
    presetClass: "emph",
    presetId: 3,
    defaultSubtype: 0,
    defaultDurationMs: 700,
    i18nKey: "emphasisFontColor",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 700;
      return h.childTnLst([
        h.anim(
          a.targetCNvPrId,
          "style.color",
          dur,
          { from: "#000000", to: "#FF0000" },
          { calcmode: "lin" }
        ),
      ]);
    },
  },
  {
    category: "emphasis",
    preset: "lineColor",
    presetClass: "emph",
    presetId: 9,
    defaultSubtype: 0,
    defaultDurationMs: 700,
    i18nKey: "emphasisLineColor",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 700;
      return h.childTnLst([
        h.anim(
          a.targetCNvPrId,
          "stroke.color",
          dur,
          { from: "#strokeColor", to: "#FF0000" },
          { calcmode: "lin" }
        ),
      ]);
    },
  },
];

// ─── Exit presets ────────────────────────────────────────────────────────
//
// Mirror entrance presets where possible: PowerPoint's preset IDs for
// exit effects parallel the entrance ones (`Disappear=1`, `Fly Out=2`,
// `Fade=10`, `Wipe=12`, `Zoom=23`, `Shape=5`, `Wheel=21`, `Split=4`).

const EXIT_PRESETS: ReadonlyArray<PresetSpec> = [
  {
    category: "exit",
    preset: "disappear",
    presetClass: "exit",
    presetId: 1,
    defaultSubtype: 0,
    defaultDurationMs: 1,
    i18nKey: "exitDisappear",
    emitBody: (a, h) =>
      h.childTnLst([h.setAttr(a.targetCNvPrId, "style.visibility", "hidden", a.durationMs ?? 1)]),
  },
  {
    category: "exit",
    preset: "flyOut",
    presetClass: "exit",
    presetId: 2,
    defaultSubtype: 4,
    defaultDurationMs: 500,
    defaultDirection: "right",
    directions: ["left", "right", "up", "down"],
    i18nKey: "exitFlyOut",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 500;
      const dir = a.direction ?? "right";
      const { axisAttr, from, to } = flyOutVector(dir);
      return h.childTnLst([
        h.anim(a.targetCNvPrId, axisAttr, dur, { from, to }),
        h.setAttr(a.targetCNvPrId, "style.visibility", "hidden", 1),
      ]);
    },
  },
  {
    category: "exit",
    preset: "fade",
    presetClass: "exit",
    presetId: 10,
    defaultSubtype: 0,
    defaultDurationMs: 500,
    i18nKey: "exitFade",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 500;
      return h.childTnLst([
        h.anim(a.targetCNvPrId, "style.opacity", dur, { from: "1", to: "0" }, { calcmode: "lin" }),
        h.setAttr(a.targetCNvPrId, "style.visibility", "hidden", 1),
      ]);
    },
  },
  {
    category: "exit",
    preset: "wipe",
    presetClass: "exit",
    presetId: 12,
    defaultSubtype: 8,
    defaultDurationMs: 500,
    defaultDirection: "down",
    directions: ["up", "down", "left", "right"],
    i18nKey: "exitWipe",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 500;
      const dir = a.direction ?? "down";
      return h.childTnLst([
        h.animEffect(a.targetCNvPrId, dur, "out", `wipe(${wipeDirection(dir)})`),
        h.setAttr(a.targetCNvPrId, "style.visibility", "hidden", 1),
      ]);
    },
  },
  {
    category: "exit",
    preset: "split",
    presetClass: "exit",
    presetId: 4,
    defaultSubtype: 26,
    defaultDurationMs: 500,
    defaultDirection: "horizontal",
    directions: ["horizontal", "vertical"],
    i18nKey: "exitSplit",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 500;
      const dir = a.direction ?? "horizontal";
      const filter = dir === "vertical" ? "barn(outVertical)" : "barn(outHorizontal)";
      return h.childTnLst([
        h.animEffect(a.targetCNvPrId, dur, "out", filter),
        h.setAttr(a.targetCNvPrId, "style.visibility", "hidden", 1),
      ]);
    },
  },
  {
    category: "exit",
    preset: "shape",
    presetClass: "exit",
    presetId: 5,
    defaultSubtype: 32,
    defaultDurationMs: 500,
    defaultDirection: "out",
    directions: ["in", "out"],
    i18nKey: "exitShape",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 500;
      const dir = a.direction === "in" ? "in" : "out";
      return h.childTnLst([
        h.animEffect(a.targetCNvPrId, dur, "out", `circle(${dir})`),
        h.setAttr(a.targetCNvPrId, "style.visibility", "hidden", 1),
      ]);
    },
  },
  {
    category: "exit",
    preset: "wheel",
    presetClass: "exit",
    presetId: 21,
    defaultSubtype: 1,
    defaultDurationMs: 1000,
    i18nKey: "exitWheel",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 1000;
      return h.childTnLst([
        h.animEffect(a.targetCNvPrId, dur, "out", "wheel(1)"),
        h.setAttr(a.targetCNvPrId, "style.visibility", "hidden", 1),
      ]);
    },
  },
  {
    category: "exit",
    preset: "zoom",
    presetClass: "exit",
    presetId: 23,
    defaultSubtype: 0,
    defaultDurationMs: 500,
    i18nKey: "exitZoom",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 500;
      return h.childTnLst([
        h.animScale(a.targetCNvPrId, dur, { x: 100, y: 100 }, { x: 0, y: 0 }),
        h.anim(a.targetCNvPrId, "style.opacity", dur, { from: "1", to: "0" }, { calcmode: "lin" }),
        h.setAttr(a.targetCNvPrId, "style.visibility", "hidden", 1),
      ]);
    },
  },
  {
    category: "exit",
    preset: "floatOut",
    presetClass: "exit",
    presetId: 42,
    defaultSubtype: 8,
    defaultDurationMs: 500,
    defaultDirection: "down",
    directions: ["up", "down"],
    i18nKey: "exitFloatOut",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 500;
      const dir = a.direction ?? "down";
      const delta = dir === "up" ? "-#ppt_h/8" : "#ppt_h/8";
      return h.childTnLst([
        h.anim(a.targetCNvPrId, "style.opacity", dur, { from: "1", to: "0" }, { calcmode: "lin" }),
        h.anim(a.targetCNvPrId, "ppt_y", dur, { from: "#ppt_y", to: `#ppt_y+${delta}` }),
        h.setAttr(a.targetCNvPrId, "style.visibility", "hidden", 1),
      ]);
    },
  },
];

// ─── Motion path presets ────────────────────────────────────────────────
//
// Every motion-path preset uses `<p:animMotion>` with a `path`
// attribute. The path string follows OOXML's compact syntax:
//   M x y    move
//   L x y    line
//   C x1 y1 x2 y2 x y   cubic bezier
//   E         end
// Coordinates are slide-relative (1.0 = full slide width / height).
// Built-in variants synthesise canonical paths; the `custom` variant
// reads `ShapeAnimation.motionPath` so authors can carry a shape's
// trajectory verbatim.

const MOTION_PATH_PRESETS: ReadonlyArray<PresetSpec> = [
  {
    category: "motionPath",
    preset: "line",
    presetClass: "path",
    presetId: 1,
    defaultSubtype: 0,
    defaultDurationMs: 2000,
    i18nKey: "motionPathLine",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 2000;
      const path = a.motionPath ?? "M 0 0 L 0.25 0 E";
      return h.childTnLst([h.animMotion(a.targetCNvPrId, dur, path)]);
    },
  },
  {
    category: "motionPath",
    preset: "arc",
    presetClass: "path",
    presetId: 2,
    defaultSubtype: 0,
    defaultDurationMs: 2000,
    i18nKey: "motionPathArc",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 2000;
      const path = a.motionPath ?? "M 0 0 C 0 -0.15 0.15 -0.25 0.3 -0.25 E";
      return h.childTnLst([h.animMotion(a.targetCNvPrId, dur, path)]);
    },
  },
  {
    category: "motionPath",
    preset: "turn",
    presetClass: "path",
    presetId: 3,
    defaultSubtype: 0,
    defaultDurationMs: 2000,
    i18nKey: "motionPathTurn",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 2000;
      const path = a.motionPath ?? "M 0 0 L 0.2 0 L 0.2 0.15 E";
      return h.childTnLst([h.animMotion(a.targetCNvPrId, dur, path)]);
    },
  },
  {
    category: "motionPath",
    preset: "loops",
    presetClass: "path",
    presetId: 4,
    defaultSubtype: 0,
    defaultDurationMs: 3000,
    i18nKey: "motionPathLoops",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 3000;
      const path =
        a.motionPath ??
        "M 0 0 C 0 -0.1 0.1 -0.15 0.15 -0.05 C 0.2 0.05 0.05 0.1 0.05 0 C 0.05 -0.1 0.2 -0.05 0.2 0.05 E";
      return h.childTnLst([h.animMotion(a.targetCNvPrId, dur, path)]);
    },
  },
  {
    category: "motionPath",
    preset: "custom",
    presetClass: "path",
    presetId: 1, // PowerPoint marks user paths with presetId=1, presetSubtype=0 too
    defaultSubtype: 0,
    defaultDurationMs: 2000,
    i18nKey: "motionPathCustom",
    emitBody: (a, h) => {
      const dur = a.durationMs ?? 2000;
      const path = a.motionPath ?? "M 0 0 L 0.2 0 E";
      return h.childTnLst([h.animMotion(a.targetCNvPrId, dur, path)]);
    },
  },
];

export const ANIMATION_PRESETS: ReadonlyArray<PresetSpec> = [
  ...ENTRANCE_PRESETS,
  ...EMPHASIS_PRESETS,
  ...EXIT_PRESETS,
  ...MOTION_PATH_PRESETS,
];

// First-wins: when two specs collide on `(presetClass, presetId)`
// (notably `motionPath/line` vs `motionPath/custom`, both `path:1`),
// the canonical built-in is registered first so the parser starts
// from `line` and only upgrades to `custom` when the path string
// differs from the canonical one (see `parseSlideTiming`).
const PRESETS_BY_KEY = (() => {
  const m = new Map<string, PresetSpec>();
  for (const p of ANIMATION_PRESETS) {
    const k = `${p.presetClass}:${p.presetId}`;
    if (!m.has(k)) m.set(k, p);
  }
  return m;
})();

const PRESETS_BY_CATEGORY_PRESET = new Map<string, PresetSpec>(
  ANIMATION_PRESETS.map((p) => [`${p.category}:${p.preset}`, p])
);

/** Look up a spec by `(presetClass, presetId)`. Used by the parser. */
export function findPresetByOoxmlIds(
  presetClass: string,
  presetId: number
): PresetSpec | undefined {
  // Custom motion-path collides with `path/1`; the parser distinguishes
  // the `custom` variant only when the path is user-specified, which we
  // can't tell from the tuple alone. Default to the typed `line` preset
  // and rely on the parser to upgrade to `custom` when the OOXML carries
  // a path that doesn't match a built-in.
  return PRESETS_BY_KEY.get(`${presetClass}:${presetId}`);
}

/** Look up a spec by `(category, preset)`. Used by the serializer / UI. */
export function findPreset(
  category: AnimationCategory,
  preset: AnimationPreset
): PresetSpec | undefined {
  return PRESETS_BY_CATEGORY_PRESET.get(`${category}:${preset}`);
}

/** Group the registry by category for the picker UI. */
export function presetsByCategory(): Readonly<Record<AnimationCategory, ReadonlyArray<PresetSpec>>> {
  return {
    entrance: ENTRANCE_PRESETS,
    emphasis: EMPHASIS_PRESETS,
    exit: EXIT_PRESETS,
    motionPath: MOTION_PATH_PRESETS,
  };
}

/**
 * OOXML emits direction by mutating `presetSubtype`. Round-trip works
 * fine when we re-emit because the parser also reads the direction-aware
 * helpers below; for brand-new authoring we set a stable subtype so
 * PowerPoint plays the right variant.
 */
export function subtypeFor(spec: PresetSpec, direction: AnimationDirection | undefined): number {
  if (!direction) return spec.defaultSubtype;
  const map: Partial<Record<AnimationDirection, number>> = {
    // Standard PowerPoint directional subtypes (ECMA-376 Part 1 §19.5.20).
    left: 4,
    right: 2,
    up: 1,
    down: 8,
    in: 16,
    out: 32,
    horizontal: 26,
    vertical: 10,
    clockwise: 0,
    counterclockwise: 1,
  };
  return map[direction] ?? spec.defaultSubtype;
}

/** Reverse of `subtypeFor` — best-effort lookup used by the parser. */
export function directionForSubtype(
  spec: PresetSpec,
  subtype: number | undefined
): AnimationDirection | undefined {
  if (!spec.directions || subtype === undefined) return undefined;
  // Subtype codes are not globally unique: `1` means `up` for
  // direction-aware presets like FlyIn / Wipe but `counterclockwise`
  // for Spin, and `0` is both the no-direction default and Spin's
  // `clockwise`. We disambiguate by only returning a direction the
  // spec actually accepts (the `spec.directions.includes(dir)` guard
  // below). Listing rotational directions alongside cardinal ones is
  // safe because no preset advertises both axes.
  const candidates: ReadonlyArray<[AnimationDirection, number]> = [
    ["left", 4],
    ["right", 2],
    ["up", 1],
    ["down", 8],
    ["in", 16],
    ["out", 32],
    ["horizontal", 26],
    ["vertical", 10],
    ["counterclockwise", 1],
    ["clockwise", 0],
  ];
  for (const [dir, code] of candidates) {
    if (code === subtype && spec.directions.includes(dir)) return dir;
  }
  return undefined;
}

// ─── Internal helpers used by emitBody implementations ──────────────────

interface FlyVector {
  readonly axisAttr: "ppt_x" | "ppt_y";
  readonly from: string;
  readonly to: string;
}

function flyInVector(direction: AnimationDirection): FlyVector {
  switch (direction) {
    case "left":
      return { axisAttr: "ppt_x", from: "#ppt_x-#ppt_w", to: "#ppt_x" };
    case "right":
      return { axisAttr: "ppt_x", from: "#ppt_x+#ppt_w", to: "#ppt_x" };
    case "up":
      return { axisAttr: "ppt_y", from: "#ppt_y-#ppt_h", to: "#ppt_y" };
    case "down":
      return { axisAttr: "ppt_y", from: "#ppt_y+#ppt_h", to: "#ppt_y" };
    default:
      return { axisAttr: "ppt_x", from: "#ppt_x-#ppt_w", to: "#ppt_x" };
  }
}

function flyOutVector(direction: AnimationDirection): FlyVector {
  switch (direction) {
    case "left":
      return { axisAttr: "ppt_x", from: "#ppt_x", to: "#ppt_x-#ppt_w" };
    case "right":
      return { axisAttr: "ppt_x", from: "#ppt_x", to: "#ppt_x+#ppt_w" };
    case "up":
      return { axisAttr: "ppt_y", from: "#ppt_y", to: "#ppt_y-#ppt_h" };
    case "down":
      return { axisAttr: "ppt_y", from: "#ppt_y", to: "#ppt_y+#ppt_h" };
    default:
      return { axisAttr: "ppt_x", from: "#ppt_x", to: "#ppt_x+#ppt_w" };
  }
}

function wipeDirection(direction: AnimationDirection): string {
  switch (direction) {
    case "up":
      return "fromBottom";
    case "down":
      return "fromTop";
    case "left":
      return "fromRight";
    case "right":
      return "fromLeft";
    default:
      return "fromBottom";
  }
}
