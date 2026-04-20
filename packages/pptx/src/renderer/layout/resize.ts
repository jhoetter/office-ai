import type { BoundingBox } from "./shape.js";

/**
 * The 8 standard resize handles a shape exposes. Cardinal handles
 * (`n`/`s`/`e`/`w`) resize along one axis; corner handles
 * (`ne`/`se`/`sw`/`nw`) resize along both. Names mirror compass
 * directions on the *unrotated* shape, even when the shape is
 * displayed at a non-zero rotation.
 */
export type ResizeHandle = "n" | "s" | "e" | "w" | "nw" | "ne" | "sw" | "se";

export interface ResolveRotatedResizeOpts {
  /** Pre-resize bounding box in slide-space EMU (unrotated). */
  readonly o: BoundingBox;
  /** Shape rotation in degrees, clockwise. `0` is the unrotated path. */
  readonly rotDeg: number;
  /** Which handle the user grabbed (in the shape's local frame). */
  readonly h: ResizeHandle;
  /** Cursor delta in screen-space EMU since pointerdown. */
  readonly dxEmu: number;
  readonly dyEmu: number;
  /** Floor on the new width / height in EMU. Use `0` for line shapes. */
  readonly minSize: number;
}

/**
 * Rotation-aware resize for a single shape. Mirrors PowerPoint /
 * Figma: the corner OPPOSITE the dragged handle stays anchored in
 * SCREEN space, the dragged handle tracks the cursor, and the new
 * width/height grow along the shape's LOCAL axes.
 *
 * Returned box is still in unrotated slide-space EMU — the renderer
 * applies the rotation around the new centre, matching the OOXML
 * convention where `<a:off>`/`<a:ext>` describe the unrotated rect
 * and `<a:xfrm rot=…>` rotates around its centre.
 *
 * Math:
 *   1. Project screen-space cursor delta into the shape's local axes
 *      by rotating by `-rotDeg`. This gives the local-frame growth
 *      along each side.
 *   2. Apply the same width/height formulas the unrotated path uses,
 *      yielding the new local size.
 *   3. Compute the screen-space anchor from the original box
 *      (centre + R(rot) · localAnchorOffset). This point is fixed
 *      for the duration of the gesture.
 *   4. Re-derive the new screen-space centre so the anchor at the
 *      new local offset (relative to the new size) lands at the
 *      same screen-space anchor. From the centre we derive
 *      `(nx, ny)` for the unrotated bounding box.
 *
 * The `rotDeg === 0` path is intentionally bit-identical to the
 * legacy unrotated formulas so unrotated shapes round-trip without
 * subpixel drift; callers should branch on `rotDeg !== 0` and only
 * route rotated shapes through this helper.
 */
export function resolveRotatedResize(opts: ResolveRotatedResizeOpts): BoundingBox {
  const { o, rotDeg, h, dxEmu, dyEmu, minSize } = opts;
  const r = (rotDeg * Math.PI) / 180;
  const cos = Math.cos(r);
  const sin = Math.sin(r);

  // Project screen-space delta into local axes (rotate by -rotDeg).
  // R(-r) = [[ cos,  sin], [-sin,  cos]]
  const dxL = dxEmu * cos + dyEmu * sin;
  const dyL = -dxEmu * sin + dyEmu * cos;

  // Apply size formulas in the local frame. These match the legacy
  // axis-aligned path one-for-one — only the inputs are rotated.
  let nw = o.cx;
  let nh = o.cy;
  if (h.includes("e")) nw = Math.max(minSize, o.cx + dxL);
  if (h.includes("w")) nw = Math.max(minSize, o.cx - dxL);
  if (h.includes("s")) nh = Math.max(minSize, o.cy + dyL);
  if (h.includes("n")) nh = Math.max(minSize, o.cy - dyL);

  // Screen-space anchor (opposite-corner of the rotated body, fixed
  // for the gesture). For edge handles we anchor the opposite edge's
  // midpoint and let the perpendicular axis pivot around the centre.
  const aOld = anchorOffsetForHandle(h, o.cx, o.cy);
  const cx0 = o.x + o.cx / 2;
  const cy0 = o.y + o.cy / 2;
  // R(r) = [[ cos, -sin], [ sin,  cos]]
  const ax = cx0 + (aOld.x * cos - aOld.y * sin);
  const ay = cy0 + (aOld.x * sin + aOld.y * cos);

  // After resize, the same logical anchor sits at a new local offset
  // (corner of the new box, centred on the new origin).
  const aNew = anchorOffsetForHandle(h, nw, nh);
  const cxNew = ax - (aNew.x * cos - aNew.y * sin);
  const cyNew = ay - (aNew.x * sin + aNew.y * cos);

  return {
    x: cxNew - nw / 2,
    y: cyNew - nh / 2,
    cx: nw,
    cy: nh,
  };
}

/**
 * Local-frame offset (relative to the box centre) of the corner /
 * edge midpoint OPPOSITE the dragged handle. This is the point
 * pinned in screen space during a rotation-aware resize.
 *
 * `+x` = local right, `+y` = local bottom (matching the unrotated
 * EMU convention used everywhere else).
 */
function anchorOffsetForHandle(
  h: ResizeHandle,
  cx: number,
  cy: number
): { readonly x: number; readonly y: number } {
  // For corners both axes pin; for edge handles the perpendicular
  // axis stays at the centre line so the box pivots around its own
  // midline (matches PowerPoint's edge-handle behaviour).
  let ax = 0;
  let ay = 0;
  if (h.includes("e")) ax = -cx / 2;
  if (h.includes("w")) ax = cx / 2;
  if (h.includes("s")) ay = -cy / 2;
  if (h.includes("n")) ay = cy / 2;
  return { x: ax, y: ay };
}
