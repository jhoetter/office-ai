import type { SlideSize } from "../../model/types.js";

export function slideViewBox(size: SlideSize): string {
  return `0 0 ${size.cxEmu} ${size.cyEmu}`;
}

export function slideAspectRatio(size: SlideSize): number {
  return size.cxEmu / size.cyEmu;
}
