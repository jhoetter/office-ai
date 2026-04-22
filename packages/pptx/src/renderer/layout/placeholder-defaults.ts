import { wrapFontFamily } from "@officeai/text-formatting";
import type { PlaceholderSpec, Shape, SlideLayout, TextShape } from "../../model/types.js";
import type { BoundingBox } from "./shape.js";

/**
 * Resolved typography for an empty (or partially empty) placeholder.
 *
 * PowerPoint placeholders inherit their default font / size / colour
 * / alignment from a chain that walks: shape's own run props →
 * layout placeholder's `lstStyle` → master's `titleStyle` /
 * `bodyStyle` / `otherStyle` → theme's `majorFont` / `minorFont`.
 *
 * This resolver currently terminates at the layout placeholder (the
 * master / theme tail is preserved as opaque XML in our model and
 * would require a real OOXML walk to introspect). The type-based
 * fallback table below mirrors PowerPoint's built-in placeholder
 * defaults — large left-anchored title, smaller top-anchored body,
 * etc — so an empty placeholder still renders with the expected
 * size and weight even when the layout doesn't override anything.
 *
 * The returned shape is the same one used by the SVG hint renderer
 * AND the HTML edit overlay, which is the whole point — both paths
 * must agree or "click to edit" looks like the font visibly changes.
 */
export interface PlaceholderTextDefaults {
  readonly fontFamily: string;
  readonly fontSizePt: number;
  readonly fontWeight: number;
  readonly align: "left" | "center" | "right";
  readonly anchor: "top" | "middle" | "bottom";
  readonly fillHex: string | null;
}

/**
 * Default font family for every placeholder type. Matches the SVG
 * placeholder hint renderer (`packages/pptx/src/renderer/svg/shapes.ts`)
 * so the SVG hint and the edit overlay use the same family.
 *
 * `wrapFontFamily` appends a `system-ui, sans-serif` tail; the
 * `@font-face` aliases in `apps/web/app/globals.css` redefine
 * `Calibri` itself to resolve to a bundled metric-equivalent
 * open-source twin (Carlito) on systems without Office, so the
 * placeholder hint visually matches what PowerPoint would render.
 * Non-null assertion is safe — `"Calibri"` is a non-empty literal.
 */
const DEFAULT_PLACEHOLDER_FAMILY: string = wrapFontFamily("Calibri")!;

/**
 * Type-based defaults that approximate PowerPoint's built-in
 * placeholder styles. Chosen to match what a fresh PowerPoint slide
 * shows when the user types into an empty placeholder (40pt
 * centered Title-Slide title, 36pt left-anchored content title,
 * 24pt centered subtitle, 18pt left-top body, …). Lets the editor
 * render visually-correct prompts without parsing the master /
 * theme XML chain.
 */
function typeDefaults(type: string | undefined): PlaceholderTextDefaults {
  switch (type) {
    case "ctrTitle":
      return {
        fontFamily: DEFAULT_PLACEHOLDER_FAMILY,
        fontSizePt: 40,
        fontWeight: 400,
        align: "center",
        anchor: "middle",
        fillHex: null,
      };
    case "title":
      return {
        fontFamily: DEFAULT_PLACEHOLDER_FAMILY,
        fontSizePt: 36,
        fontWeight: 400,
        align: "left",
        anchor: "middle",
        fillHex: null,
      };
    case "subTitle":
      return {
        fontFamily: DEFAULT_PLACEHOLDER_FAMILY,
        fontSizePt: 24,
        fontWeight: 400,
        align: "center",
        anchor: "middle",
        fillHex: null,
      };
    case "body":
      return {
        fontFamily: DEFAULT_PLACEHOLDER_FAMILY,
        fontSizePt: 18,
        fontWeight: 400,
        align: "left",
        anchor: "top",
        fillHex: null,
      };
    case "ftr":
    case "hdr":
      return {
        fontFamily: DEFAULT_PLACEHOLDER_FAMILY,
        fontSizePt: 12,
        fontWeight: 400,
        align: "center",
        anchor: "middle",
        fillHex: null,
      };
    case "dt":
      return {
        fontFamily: DEFAULT_PLACEHOLDER_FAMILY,
        fontSizePt: 12,
        fontWeight: 400,
        align: "left",
        anchor: "middle",
        fillHex: null,
      };
    case "sldNum":
      return {
        fontFamily: DEFAULT_PLACEHOLDER_FAMILY,
        fontSizePt: 12,
        fontWeight: 400,
        align: "right",
        anchor: "middle",
        fillHex: null,
      };
    default:
      return {
        fontFamily: DEFAULT_PLACEHOLDER_FAMILY,
        fontSizePt: 18,
        fontWeight: 400,
        align: "left",
        anchor: "top",
        fillHex: null,
      };
  }
}

/**
 * Walk the inheritance chain to compute the effective text defaults
 * for a shape. The shape's own run / paragraph properties always
 * win when present; otherwise we fall back to type-based defaults
 * keyed on `shape.placeholder.type`.
 *
 * `_layout` is currently unused but accepted for API stability — it
 * will become live once we walk the layout's `<a:lstStyle>` blob.
 */
export function resolvePlaceholderTextDefaults(
  shape: TextShape,
  _layout: SlideLayout | undefined
): PlaceholderTextDefaults {
  const fallback = typeDefaults(shape.placeholder?.type);
  const firstRun = shape.txBody.paragraphs[0]?.runs.find((r) => !r.isLineBreak);
  const firstPara = shape.txBody.paragraphs[0];
  // `wrapFontFamily` adds the trailing `system-ui, sans-serif` tail.
  // When the run has no font of its own we use the placeholder
  // fallback (already wrapped) verbatim.
  const fontFamily = wrapFontFamily(firstRun?.properties.fontFamily) ?? fallback.fontFamily;
  const fontSizePt =
    firstRun?.properties.fontSizeHundredths !== undefined
      ? firstRun.properties.fontSizeHundredths / 100
      : fallback.fontSizePt;
  const fontWeight = firstRun?.properties.bold ? 700 : fallback.fontWeight;
  const align: PlaceholderTextDefaults["align"] = (() => {
    const a = firstPara?.properties.alignment;
    if (a === "center") return "center";
    if (a === "right") return "right";
    if (a === "left") return "left";
    return fallback.align;
  })();
  const anchor: PlaceholderTextDefaults["anchor"] = (() => {
    const v = shape.txBody.bodyPrRaw?.attrs?.anchor ?? shape.txBody.bodyPrRaw?.rawAttrs?.["@_anchor"];
    if (v === "ctr") return "middle";
    if (v === "b") return "bottom";
    if (v === "t") return "top";
    return fallback.anchor;
  })();
  const fillHex = firstRun?.properties.color ?? fallback.fillHex;
  return { fontFamily, fontSizePt, fontWeight, align, anchor, fillHex };
}

/**
 * Look up the matching `<p:ph>` spec on the slide layout for a
 * placeholder shape. PowerPoint matches first by `(type, idx)`,
 * then by `idx` alone, then by `type` alone. Returns `undefined`
 * for non-placeholder shapes or when the layout doesn't declare a
 * matching slot.
 */
export function findLayoutPlaceholder(
  shape: Shape,
  layout: SlideLayout | undefined
): PlaceholderSpec | undefined {
  if (!layout) return undefined;
  const ph = (shape as TextShape).placeholder;
  if (!ph) return undefined;
  const type = ph.type;
  const idx = ph.idx;
  const placeholders = layout.placeholders;
  if (type !== undefined && idx !== undefined) {
    const exact = placeholders.find((p) => p.type === type && p.idx === idx);
    if (exact) return exact;
  }
  if (idx !== undefined) {
    const byIdx = placeholders.find((p) => p.idx === idx);
    if (byIdx) return byIdx;
  }
  if (type !== undefined) {
    const byType = placeholders.find((p) => p.type === type);
    if (byType) return byType;
  }
  return undefined;
}

/**
 * Like `shapeBoundingBox` but falls back to the matching layout
 * placeholder's geometry when the shape itself has no explicit
 * `position` / `size`. PowerPoint placeholders frequently inherit
 * their geometry from the layout; without this fallback the
 * editor's overlay code returns null and silently refuses to open
 * (the user perceives it as "the placeholder has no clickable
 * area").
 */
export function resolvedShapeBoundingBox(shape: Shape, layout: SlideLayout | undefined): BoundingBox | null {
  if (shape.position && shape.size) {
    return {
      x: shape.position.xEmu,
      y: shape.position.yEmu,
      cx: shape.size.cxEmu,
      cy: shape.size.cyEmu,
    };
  }
  const layoutPh = findLayoutPlaceholder(shape, layout);
  if (layoutPh?.position && layoutPh.size) {
    return {
      x: layoutPh.position.xEmu,
      y: layoutPh.position.yEmu,
      cx: layoutPh.size.cxEmu,
      cy: layoutPh.size.cyEmu,
    };
  }
  return null;
}
