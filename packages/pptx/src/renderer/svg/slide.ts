import type { Slide } from "../../model/types.js";
import { slideViewBox } from "../layout/slide.js";
import { shapeToSvg, type SvgRenderCtx } from "./shapes.js";

export function slideToSvgString(slide: Slide, ctx: SvgRenderCtx): string {
  const parts: string[] = [
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${slideViewBox(ctx.slideSize)}" preserveAspectRatio="xMidYMid meet">`,
    `<rect width="100%" height="100%" fill="white"/>`,
  ];
  for (const s of slide.shapes) parts.push(shapeToSvg(s, ctx));
  parts.push(`</svg>`);
  return parts.join("");
}
