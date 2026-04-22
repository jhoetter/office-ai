import { ooxml, type IdMinter } from "@officeai/core";
import type { MediaShape, OpaqueXml } from "../model/types.js";
import { attrOf, captureOpaque, elementEntries, findElementEntry } from "./xml-helpers.js";

/**
 * If a `<p:pic>` entry hosts an `<a:videoFile>` or `<a:audioFile>` inside
 * its `<p:nvPicPr><p:nvPr>`, lift it into a typed {@link MediaShape}.
 * Returns `null` for ordinary picture shapes so the caller falls back
 * to the existing `parsePic` path.
 *
 * The Phase-1 implementation is intentionally a "stub": every
 * unmodelled child of `<p:pic>` is captured verbatim through `raw`
 * so byte-faithful round-trip works without us having to introspect
 * the entire media OOXML schema (poster `<p:blipFill>`, the optional
 * `<p14:media>` extLst that newer PowerPoint adds, audio waveform
 * `<p:cNvPr><a:hlinkClick>` etc.). Only `cNvPrId`, `name`,
 * position/size, the resolved media + poster paths, and the
 * media/poster rel ids are surfaced as typed fields.
 */
export function tryParseMediaShape(
  picEntry: Record<string, unknown>,
  mintNodeId: IdMinter,
  partPath: string,
  slideRelTargets: ReadonlyMap<string, string>
): MediaShape | null {
  const children = (picEntry["p:pic"] as unknown[] | undefined) ?? [];
  const nvPicPr = findElementEntry(children, "p:nvPicPr");
  if (!nvPicPr) return null;
  const nvPr = findElementEntry((nvPicPr["p:nvPicPr"] as unknown[] | undefined) ?? [], "p:nvPr");
  if (!nvPr) return null;

  let mediaType: "video" | "audio" | null = null;
  let mediaRelId = "";
  for (const c of elementEntries((nvPr["p:nvPr"] as unknown[] | undefined) ?? [])) {
    const tag = ooxml.getTag(c);
    if (tag === "a:videoFile") {
      mediaType = "video";
      mediaRelId = attrOf(c, "r:link") ?? attrOf(c, "r:embed") ?? "";
    } else if (tag === "a:audioFile") {
      mediaType = "audio";
      mediaRelId = attrOf(c, "r:link") ?? attrOf(c, "r:embed") ?? "";
    }
  }
  if (!mediaType) return null;

  let cNvPrId = 0;
  let name = "";
  const cNvPr = findElementEntry((nvPicPr["p:nvPicPr"] as unknown[] | undefined) ?? [], "p:cNvPr");
  if (cNvPr) {
    cNvPrId = Number(attrOf(cNvPr, "id") ?? "0");
    name = attrOf(cNvPr, "name") ?? "";
  }

  // Poster image rides on the standard `<p:blipFill>/<a:blip r:embed>`.
  let posterRelId: string | undefined;
  const blipFill = findElementEntry(children, "p:blipFill");
  if (blipFill) {
    const blip = findElementEntry((blipFill["p:blipFill"] as unknown[] | undefined) ?? [], "a:blip");
    if (blip) {
      const embed = attrOf(blip, "r:embed");
      if (embed) posterRelId = embed;
    }
  }

  let position: { xEmu: number; yEmu: number } | undefined;
  let size: { cxEmu: number; cyEmu: number } | undefined;
  let rotation: number | undefined;
  const spPr = findElementEntry(children, "p:spPr");
  if (spPr) {
    const xfrm = findElementEntry((spPr["p:spPr"] as unknown[] | undefined) ?? [], "a:xfrm");
    if (xfrm) {
      const rotAttr = attrOf(xfrm, "rot");
      if (rotAttr !== undefined) {
        const n = Number(rotAttr);
        if (Number.isFinite(n) && n !== 0) rotation = n / 60000;
      }
      const xfrmChildren = (xfrm["a:xfrm"] as unknown[] | undefined) ?? [];
      const off = findElementEntry(xfrmChildren, "a:off");
      const ext = findElementEntry(xfrmChildren, "a:ext");
      if (off) {
        position = {
          xEmu: Number(attrOf(off, "x") ?? "0"),
          yEmu: Number(attrOf(off, "y") ?? "0"),
        };
      }
      if (ext) {
        size = {
          cxEmu: Number(attrOf(ext, "cx") ?? "0"),
          cyEmu: Number(attrOf(ext, "cy") ?? "0"),
        };
      }
    }
  }

  const mediaPath = mediaRelId ? resolveSlideRel(slideRelTargets, partPath, mediaRelId) : "";
  const posterPath = posterRelId ? resolveSlideRel(slideRelTargets, partPath, posterRelId) : undefined;

  const raw: OpaqueXml = captureOpaque(picEntry);

  return {
    kind: "media",
    id: mintNodeId(),
    cNvPrId,
    name,
    ...(position ? { position } : {}),
    ...(size ? { size } : {}),
    ...(rotation !== undefined ? { rotation } : {}),
    mediaType,
    mediaRelId,
    mediaPath,
    ...(posterRelId ? { posterRelId } : {}),
    ...(posterPath ? { posterPath } : {}),
    raw,
  };
}

function resolveSlideRel(
  slideRelTargets: ReadonlyMap<string, string>,
  partPath: string,
  relId: string
): string {
  const target = slideRelTargets.get(relId);
  if (!target) return "";
  return resolveTarget(partPath, target);
}

function resolveTarget(ownerPart: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const ownerDir = ownerPart.split("/").slice(0, -1);
  const segs = (ownerDir.length ? `${ownerDir.join("/")}/${target}` : target).split("/");
  const out: string[] = [];
  for (const s of segs) {
    if (!s || s === ".") continue;
    if (s === "..") {
      out.pop();
      continue;
    }
    out.push(s);
  }
  return out.join("/");
}
