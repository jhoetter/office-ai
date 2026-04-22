"use client";

import * as React from "react";
import type { TextShape } from "@officeai/pptx";
import { useTranslator } from "@/lib/i18n";

/**
 * Floating "geometry" mini-bar shown when a single text shape with a
 * parameterised preset geometry (e.g. `roundRect`, `rightArrow`,
 * `parallelogram`, …) is selected. Mirrors the on-canvas yellow
 * adjustment handles PowerPoint shows on the same shapes.
 *
 * Each slider edits one `<a:gd name="adjN" fmla="val …"/>` entry
 * inside the shape's `<a:prstGeom>/<a:avLst>` via the
 * `pptx:set-shape-geometry` command. Number of sliders is driven by
 * the per-preset `adjCount` table below; presets we haven't catalogued
 * fall back to a single `adj1` slider, which covers the common case
 * (corner radius / slant / arrow body).
 */
export interface ShapeGeometryContextBarProps {
  readonly shape: TextShape;
  /**
   * Apply a single adjustment update. The implementation should
   * dispatch `pptx:set-shape-geometry` with `{ adjName, value }`.
   * Returning a promise is fine — the bar treats the call as fire-
   * and-forget; errors propagate via the editor's existing toast
   * channel.
   */
  readonly onChange: (adjName: string, value: number) => void;
  /** Optional positioning style applied to the wrapper. */
  readonly style?: React.CSSProperties;
}

const ADJ_MIN = 0;
const ADJ_MAX = 100_000;
const ADJ_STEP = 1_000;

/**
 * Number of `<a:avLst>` adjustments each preset declares. Anything
 * not in this table renders a single `adj1` slider — that's the
 * shape of every common parameterised preset we don't special-case
 * (parallelogram, trapezoid, bevel, frame, chevron, pentagon,
 * homePlate, pie, donut, noSmoking, …).
 */
const PRESET_ADJ_COUNT: Readonly<Record<string, 1 | 2 | 3>> = {
  roundRect: 1,
  round1Rect: 1,
  round2SameRect: 2,
  round2DiagRect: 2,
  snip1Rect: 1,
  snip2SameRect: 2,
  snip2DiagRect: 2,
  parallelogram: 1,
  trapezoid: 1,
  bevel: 1,
  frame: 1,
  chevron: 1,
  pentagon: 1,
  homePlate: 1,
  arc: 2,
  pie: 2,
  donut: 1,
  noSmoking: 1,
  arrow: 2,
  leftArrow: 2,
  rightArrow: 2,
  upArrow: 2,
  downArrow: 2,
  leftRightArrow: 2,
  upDownArrow: 2,
};

/**
 * Per-preset default for each `adjN`. PowerPoint's reference defaults
 * vary wildly between presets; we only special-case `roundRect` (the
 * one preset whose default isn't 50%) — every other entry resolves to
 * `50000` (50%), which lines up with the canonical OOXML defaults for
 * the arrow / snip / round-corner / pie family the way users perceive
 * them on a freshly-inserted shape.
 */
function presetDefault(prst: string, adjName: string): number {
  if (prst === "roundRect" && (adjName === "adj" || adjName === "adj1")) return 16667;
  return 50_000;
}

/** Cardinality (1, 2 or 3) of sliders we should expose for a preset. */
function adjCountFor(prst: string | null): 1 | 2 | 3 {
  if (!prst) return 1;
  return PRESET_ADJ_COUNT[prst] ?? 1;
}

/**
 * Pull the `prst` attribute off a shape's `<a:prstGeom>` (mirrors the
 * private `readPrstGeom` helper in `renderer/svg/shapes.ts` — kept
 * inline so the editor doesn't reach into the renderer's deep
 * imports).
 */
function readPrst(shape: TextShape): string | null {
  for (const c of shape.spPrTail) {
    if (c.tag !== "a:prstGeom") continue;
    const prst = c.attrs?.prst ?? c.rawAttrs?.["@_prst"];
    if (typeof prst === "string") return prst;
  }
  return null;
}

/**
 * Pull the `<a:gd name=adjN fmla="val 12345"/>` map out of a shape's
 * `<a:avLst>`. Mirrors `readPrstAdjustments` in
 * `renderer/svg/shapes.ts` — duplicated here so we don't need to
 * promote that helper into the package's public surface for one UI
 * call site.
 */
function readAdjustments(shape: TextShape): ReadonlyMap<string, number> {
  const out = new Map<string, number>();
  for (const c of shape.spPrTail) {
    if (c.tag !== "a:prstGeom") continue;
    for (const inner of c.subtree) {
      if (!inner || typeof inner !== "object" || Array.isArray(inner)) continue;
      const obj = inner as Record<string, unknown>;
      const keys = Object.keys(obj).filter((k) => k !== ":@");
      if (keys.length !== 1 || keys[0] !== "a:avLst") continue;
      const subs = (obj["a:avLst"] as unknown[] | undefined) ?? [];
      for (const gd of subs) {
        if (!gd || typeof gd !== "object" || Array.isArray(gd)) continue;
        const g = gd as Record<string, unknown>;
        if (Object.keys(g).filter((k) => k !== ":@")[0] !== "a:gd") continue;
        const a = (g[":@"] as Record<string, unknown> | undefined) ?? {};
        const name = typeof a["@_name"] === "string" ? a["@_name"] : undefined;
        const fmla = typeof a["@_fmla"] === "string" ? a["@_fmla"] : undefined;
        if (!name || !fmla) continue;
        const m = /^val\s+(-?\d+)/.exec(fmla);
        if (!m) continue;
        out.set(name, Number(m[1]));
      }
      return out;
    }
  }
  return out;
}

/**
 * Decide whether a given selected text shape exposes adjustable
 * preset geometry. Returns `false` for plain rectangles / ellipses /
 * lines — none of those have `<a:avLst>` knobs — and for any
 * `custGeom`-based shape (no `prst` at all). Returns `true` for
 * presets we know how to drive (the table above) and for presets we
 * don't know but where the file already declares one or more
 * `<a:gd>` entries (so the user can keep tuning whatever's there).
 */
export function shapeHasAdjustableGeometry(shape: TextShape): boolean {
  const prst = readPrst(shape);
  if (!prst) return false;
  if (PRESET_ADJ_COUNT[prst] !== undefined) return true;
  return readAdjustments(shape).size > 0;
}

export function ShapeGeometryContextBar({
  shape,
  onChange,
  style,
}: ShapeGeometryContextBarProps): React.ReactElement | null {
  const { t } = useTranslator();
  const prst = readPrst(shape);
  if (!prst) return null;
  const adjustments = readAdjustments(shape);
  const count = adjCountFor(prst);
  // For roundRect we display a friendlier label than the generic
  // "Adjustment 1" — corner radius is what 99% of users came here to
  // tweak.
  const isRoundRectFamily =
    prst === "roundRect" || prst === "round1Rect" || prst === "round2SameRect" || prst === "round2DiagRect";

  const sliders: React.ReactElement[] = [];
  for (let i = 1; i <= count; i++) {
    const adjName = `adj${i}`;
    // OOXML legacy: `roundRect` writes `adj` (no number) for its
    // single knob. Read either form so an existing file's value
    // shows up in the slider.
    const current =
      adjustments.get(adjName) ??
      (i === 1 ? adjustments.get("adj") : undefined) ??
      presetDefault(prst, adjName);
    const label =
      isRoundRectFamily && i === 1
        ? t("pptx.geometry.cornerRadius")
        : t("pptx.geometry.adjustment", { n: i });
    sliders.push(
      // Keying on `value` remounts the slider when the underlying
      // snapshot changes externally (undo / redo / collaborator
      // edit), so we don't need a setState-in-effect dance to keep
      // the local drag state in sync. During an in-flight drag the
      // value prop is stable (we only commit on mouseup) so the
      // remount never disrupts the user's interaction.
      <GeometrySlider
        key={`${adjName}-${current}`}
        adjName={adjName}
        label={label}
        initialValue={current}
        onChange={(v) => onChange(adjName, v)}
      />
    );
  }

  return (
    <div
      data-testid="pptx-shape-geometry-bar"
      className="flex items-center gap-3 rounded-md border border-divider bg-surface px-3 py-1.5 shadow-md"
      style={style}
      onMouseDown={(e) => {
        // Prevent the canvas from clearing selection / starting a
        // marquee when the user clicks one of our sliders.
        e.stopPropagation();
      }}
    >
      <span className="select-none text-[11px] font-medium uppercase tracking-wide text-secondary">
        {t("pptx.geometry.title")}
      </span>
      {sliders}
    </div>
  );
}

interface GeometrySliderProps {
  readonly adjName: string;
  readonly label: string;
  readonly initialValue: number;
  readonly onChange: (next: number) => void;
}

/**
 * Single `<a:gd>` slider. We update local state on every `input`
 * event so the value label tracks the drag, and only dispatch a
 * command on `change` (mouseup / focus-out / keyboard commit) so the
 * undo stack records one entry per drag rather than dozens of
 * intermediate steps.
 *
 * The component is mounted with `key={value}` from the parent, so
 * external value changes (undo / redo / collaborator edit) replace
 * the slider rather than fighting the local drag state — keeps the
 * component's render path free of setState-in-effect / setState-in-
 * render gymnastics.
 */
function GeometrySlider({ adjName, label, initialValue, onChange }: GeometrySliderProps): React.ReactElement {
  const [draft, setDraft] = React.useState<number>(initialValue);
  const clamped = Math.max(ADJ_MIN, Math.min(ADJ_MAX, draft));
  const percent = Math.round((clamped / 100_000) * 100);
  const id = `pptx-shape-geometry-${adjName}`;
  return (
    <label htmlFor={id} className="flex items-center gap-2 text-xs text-foreground">
      <span className="select-none text-secondary">{label}</span>
      <input
        id={id}
        type="range"
        min={ADJ_MIN}
        max={ADJ_MAX}
        step={ADJ_STEP}
        value={clamped}
        data-testid={`pptx-shape-geometry-${adjName}-slider`}
        onInput={(e) => setDraft(Number((e.target as HTMLInputElement).value))}
        onChange={(e) => {
          const next = Number(e.target.value);
          setDraft(next);
          onChange(Math.max(ADJ_MIN, Math.min(ADJ_MAX, next)));
        }}
        className="h-1 w-32 cursor-pointer accent-[var(--accent)]"
        aria-label={label}
      />
      <span
        data-testid={`pptx-shape-geometry-${adjName}-value`}
        className="w-9 select-none text-right text-[10px] tabular-nums text-secondary"
      >
        {percent}%
      </span>
    </label>
  );
}
