import type {
  PptxAgent,
  Shape,
  Slide,
  TextRun,
  TextRunProperties,
  TextShape,
  TextFormatPayload,
} from "@officeai/pptx";
import type { PptxTextSelection } from "@officeai/pptx/renderer/react";
import {
  collapse,
  hundredthsOfPtToPt,
  MIXED,
  normalizeColor,
  ptToHundredthsOfPt,
  type ActiveTextFormat,
  type MaybeMixed,
  type TextFormat,
  type TextFormatProvider,
} from "@officeai/text-formatting";

/**
 * PPTX adapter for the shared TextFormatBar.
 *
 * Reads run-level formatting for the currently-edited paragraph from
 * the agent snapshot and translates canonical patches into
 * `pptx:format-text` commands using PPTX's native units (font size in
 * hundredths of a point) and RRGGBB colours. Highlight is reported
 * as `"native"` since we plumb a real `<a:highlight>` element through
 * the model + parser + serializer (Phase 6a).
 *
 * Selection state is sourced from a ref written by `SlideCanvas` on
 * every `selectionchange` event. When no shape is being edited the
 * provider reports an empty `ActiveTextFormat` and `apply` no-ops so
 * toolbar clicks while the canvas is in shape-selection mode don't
 * accidentally hammer the agent.
 */
export interface PptxProviderDeps {
  readonly agentRef: { readonly current: PptxAgent | null };
  readonly slideIndexRef: { readonly current: number };
  readonly selectionRef: { readonly current: PptxTextSelection | null };
  /**
   * Currently-selected shape id (set by the canvas's shape-selection
   * handler). When the user has selected a text shape but isn't in
   * text-edit mode, the toolbar treats the entire shape as the
   * formatting target — matching PowerPoint, where you can click a
   * shape and immediately change the font without first entering edit
   * mode and selecting text.
   */
  readonly selectedShapeIdRef: { readonly current: string | null };
  readonly pushToast: (kind: "info" | "warn" | "error", text: string) => void;
}

const CAPABILITIES = {
  highlight: "native",
  underlineVariants: false,
  fontFamily: true,
  fontSize: true,
  strike: true,
} as const;

export function createPptxFormatProvider(deps: PptxProviderDeps): TextFormatProvider {
  return {
    capabilities: CAPABILITIES,
    hasSelection: () => {
      // Toolbar is enabled in three situations:
      //   1. A real text range is selected → format that range.
      //   2. The text-edit caret is parked inside a shape → format the
      //      whole shape (PowerPoint treats this as "no text selected
      //      but a shape is being edited", and font changes apply to
      //      the entire text body).
      //   3. A text shape is selected from the canvas (no edit mode
      //      open) → format the whole shape.
      // Cases 2 and 3 funnel through the same shape-wide code path in
      // `dispatchPatch`/`readActive` below.
      const sel = deps.selectionRef.current;
      if (sel != null) return true;
      return resolveSelectedTextShape(deps) != null;
    },
    getActive: () =>
      readActive(
        deps.agentRef.current,
        deps.slideIndexRef.current,
        deps.selectionRef.current,
        deps.selectedShapeIdRef.current
      ),
    apply: (patch) => {
      void dispatchPatch(deps, patch);
    },
  };
}

function resolveSelectedTextShape(deps: PptxProviderDeps): TextShape | null {
  const agent = deps.agentRef.current;
  const id = deps.selectedShapeIdRef.current;
  if (!agent || !id) return null;
  const slide = agent.getSnapshot().root.slides[deps.slideIndexRef.current];
  if (!slide) return null;
  const shape = findShape(slide.shapes, id);
  if (!shape || shape.kind !== "text") return null;
  return shape;
}

/**
 * Pure render-path helper that mirrors `provider.getActive()`.
 * Exported so the editor can compute `ActiveTextFormat` from values
 * already in scope without bouncing through the provider's refs.
 */
export function computePptxActive(
  agent: PptxAgent | null,
  slideIndex: number,
  selection: PptxTextSelection | null,
  selectedShapeId: string | null = null
): ActiveTextFormat {
  return readActive(agent, slideIndex, selection, selectedShapeId);
}

const EMPTY: ActiveTextFormat = {
  bold: undefined,
  italic: undefined,
  underline: undefined,
  strike: undefined,
  fontFamily: undefined,
  fontSizePt: undefined,
  color: undefined,
  highlight: undefined,
};

function readActive(
  agent: PptxAgent | null,
  slideIndex: number,
  selection: PptxTextSelection | null,
  selectedShapeId: string | null
): ActiveTextFormat {
  if (!agent) return EMPTY;
  const slide = agent.getSnapshot().root.slides[slideIndex];
  if (!slide) return EMPTY;

  // Prefer a live text-edit selection; fall back to the
  // canvas's shape selection so the toolbar reflects (and applies to)
  // the entire shape's runs even before the user enters edit mode.
  let shape: Shape | null = null;
  let runs: TextRun[] = [];
  if (selection) {
    shape = findShape(slide.shapes, selection.shapeId);
    if (!shape || shape.kind !== "text") return EMPTY;
    const para = shape.txBody.paragraphs[selection.paragraph];
    if (!para) return EMPTY;
    runs = runsInRange(para.runs, selection.start, selection.end);
  } else if (selectedShapeId) {
    shape = findShape(slide.shapes, selectedShapeId);
    if (!shape || shape.kind !== "text") return EMPTY;
    runs = collectAllRuns(shape);
  } else {
    return EMPTY;
  }
  if (runs.length === 0) return EMPTY;

  const props = runs.map((r) => r.properties);
  const bold = collapse(props.map((p) => p.bold === true));
  const italic = collapse(props.map((p) => p.italic === true));
  const underline = collapse(props.map((p) => p.underline === true || typeof p.underline === "string"));
  const strike = collapse(props.map((p) => p.strike === true));
  // PowerPoint defaults: when the runs don't carry an explicit
  // `latin@typeface` / `sz` / `solidFill` we fall back to the
  // theme-equivalent defaults the SVG renderer effectively shows
  // (Calibri 18pt black). Showing concrete values in the dropdowns —
  // rather than the placeholder "Font" / "Size" — matches PowerPoint
  // and makes the toolbar feel "live" the moment a shape is selected,
  // even before the user enters edit mode.
  const fontFamily = collapseWithDefault(
    props.map((p) => p.fontFamily),
    DEFAULT_FONT_FAMILY
  );
  const fontSizePt = collapseWithDefault(
    props.map((p) =>
      p.fontSizeHundredths !== undefined ? hundredthsOfPtToPt(p.fontSizeHundredths) : undefined
    ),
    DEFAULT_FONT_SIZE_PT
  );
  const color = collapseWithDefault(
    props.map((p) => normalizeColor(p.color)),
    DEFAULT_COLOR
  );
  const highlight = collapse(props.map((p) => normalizeColor(p.highlight)));

  return { bold, italic, underline, strike, fontFamily, fontSizePt, color, highlight };
}

const DEFAULT_FONT_FAMILY = "Calibri";
const DEFAULT_FONT_SIZE_PT = 18;
const DEFAULT_COLOR = "000000";

/**
 * Like `collapse` but substitutes `defaultValue` for `undefined` entries
 * before collapsing. Lets us treat "no explicit run-level value" as
 * the PowerPoint default so the toolbar always reflects something —
 * while still surfacing MIXED when a real conflict exists.
 */
function collapseWithDefault<T>(values: ReadonlyArray<T | undefined>, defaultValue: T): MaybeMixed<T> {
  return collapse(values.map((v) => (v === undefined ? defaultValue : v)));
}

function collectAllRuns(shape: TextShape): TextRun[] {
  const out: TextRun[] = [];
  for (const para of shape.txBody.paragraphs) {
    for (const r of para.runs) {
      if (!r.isLineBreak) out.push(r);
    }
  }
  return out;
}

function runsInRange(runs: ReadonlyArray<TextRun>, start: number, end: number): TextRun[] {
  if (start === end) {
    // Caret-only: report formatting at the caret position so the
    // toolbar shows the about-to-be-typed style. We pick the run
    // immediately before the caret, falling back to the first run.
    let pos = 0;
    let last: TextRun | null = null;
    for (const r of runs) {
      if (r.isLineBreak) continue;
      const len = r.text.length;
      if (start <= pos + len) return [r];
      pos += len;
      last = r;
    }
    return last ? [last] : [];
  }
  const out: TextRun[] = [];
  let pos = 0;
  for (const r of runs) {
    if (r.isLineBreak) continue;
    const len = r.text.length;
    const runStart = pos;
    const runEnd = pos + len;
    pos = runEnd;
    if (runEnd <= start || runStart >= end) continue;
    out.push(r);
  }
  return out;
}

function findShape(shapes: ReadonlyArray<Shape>, id: string): Shape | null {
  for (const s of shapes) {
    if (s.id === id) return s;
    if (s.kind === "group") {
      const inner = findShape(s.children, id);
      if (inner) return inner;
    }
  }
  return null;
}

async function dispatchPatch(deps: PptxProviderDeps, patch: TextFormat): Promise<void> {
  const agent = deps.agentRef.current;
  if (!agent) return;

  const sel = deps.selectionRef.current;
  const slide: Slide | undefined = agent.getSnapshot().root.slides[deps.slideIndexRef.current];
  if (!slide) return;

  // Resolve the target shape: prefer the text-edit selection's shape,
  // fall back to the canvas's selected shape.
  let targetShape: TextShape | null = null;
  if (sel) {
    const s = findShape(slide.shapes, sel.shapeId);
    if (s && s.kind === "text") targetShape = s;
  }
  if (!targetShape) {
    targetShape = resolveSelectedTextShape(deps);
  }
  if (!targetShape) {
    deps.pushToast("info", "Select a text shape first.");
    return;
  }

  const active = readActive(agent, deps.slideIndexRef.current, sel, deps.selectedShapeIdRef.current);
  const fmt = canonicalToPptx(patch, active);
  if (Object.keys(fmt).length === 0) return;

  // Build the list of (paragraph, start, end) ranges to format. With a
  // real text-edit range we hit just that paragraph; otherwise we fan
  // out across every paragraph in the shape so the change behaves as a
  // shape-wide format (matches PowerPoint's "select shape, change
  // font" interaction).
  const ranges: { paragraph: number; start: number; end: number }[] = [];
  if (sel && sel.start !== sel.end) {
    ranges.push({ paragraph: sel.paragraph, start: sel.start, end: sel.end });
  } else {
    targetShape.txBody.paragraphs.forEach((p, i) => {
      const len = p.runs.reduce((acc, r) => acc + (r.isLineBreak ? 0 : r.text.length), 0);
      if (len > 0) ranges.push({ paragraph: i, start: 0, end: len });
    });
  }
  if (ranges.length === 0) {
    deps.pushToast("info", "No text to format.");
    return;
  }

  try {
    for (const range of ranges) {
      await agent.applyCommand({
        type: "pptx:format-text",
        payload: {
          slideIndex: deps.slideIndexRef.current,
          shapeId: targetShape.id,
          range,
          format: fmt,
        },
        source: "human",
      });
    }
  } catch (err) {
    deps.pushToast("error", err instanceof Error ? err.message : String(err));
  }
}

function canonicalToPptx(patch: TextFormat, _active: ActiveTextFormat): TextFormatPayload {
  const out: { -readonly [K in keyof TextFormatPayload]: TextFormatPayload[K] } = {};
  if (patch.bold !== undefined) out.bold = patch.bold;
  if (patch.italic !== undefined) out.italic = patch.italic;
  if (patch.underline !== undefined) {
    if (patch.underline === false) out.underline = false;
    else if (patch.underline === true) out.underline = true;
    else out.underline = patch.underline;
  }
  if (patch.strike !== undefined) out.strike = patch.strike;
  if (patch.fontFamily !== undefined) out.fontFamily = patch.fontFamily;
  if (patch.fontSizePt !== undefined) out.fontSizeHundredths = ptToHundredthsOfPt(patch.fontSizePt);
  if (patch.color !== undefined) {
    const norm = normalizeColor(patch.color);
    out.color = norm ?? "";
  }
  if (patch.highlight !== undefined) {
    if (patch.highlight === "") out.highlight = "";
    else {
      const norm = normalizeColor(patch.highlight);
      if (norm) out.highlight = norm;
    }
  }
  return out;
}

export { MIXED };
export type { MaybeMixed, TextRunProperties };
