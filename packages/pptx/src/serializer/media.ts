import { ooxml } from "@officeai/core";
import type { MediaShape, OpaqueXml } from "../model/types.js";

const ATTR_KEY = ":@";

/**
 * Serialize a {@link MediaShape} back to its `<p:pic>` entry. The
 * Phase-1 pipeline keeps the entire `<p:pic>` subtree opaquely on
 * `MediaShape.raw`; here we walk that subtree and surgically patch:
 *
 *   - `<p:nvPicPr>/<p:cNvPr>` — id + name from the typed model.
 *   - `<p:spPr>/<a:xfrm>`     — position/size/rotation from the
 *                               typed model (clears `rot` when the
 *                               typed value is `undefined`/`0`).
 *
 * Every other captured child — the `<a:videoFile>` / `<a:audioFile>`
 * reference, the `<p14:media>` extLst, the poster `<p:blipFill>`,
 * any picture style — passes through verbatim. Untouched media
 * shapes therefore round-trip byte-identical; commands that mutate
 * the typed fields take effect on save.
 */
export function mediaShapeToEntry(shape: MediaShape): Record<string, unknown> {
  const subtree = patchPicSubtree(shape);
  const entry: Record<string, unknown> = { "p:pic": subtree };
  if (Object.keys(shape.raw.rawAttrs).length > 0) {
    entry[ATTR_KEY] = { ...shape.raw.rawAttrs };
  }
  return entry;
}

function patchPicSubtree(shape: MediaShape): unknown[] {
  const out: unknown[] = [];
  for (const child of shape.raw.subtree) {
    const patched = patchChild(child, shape);
    out.push(patched);
  }
  return out;
}

function patchChild(child: unknown, shape: MediaShape): unknown {
  if (!child || typeof child !== "object" || Array.isArray(child)) return child;
  const obj = child as Record<string, unknown>;
  const tag = ooxml.getTag(obj);
  if (tag === "p:nvPicPr") return patchNvPicPr(obj, shape);
  if (tag === "p:spPr") return patchSpPr(obj, shape);
  return child;
}

function patchNvPicPr(entry: Record<string, unknown>, shape: MediaShape): Record<string, unknown> {
  const children = (entry["p:nvPicPr"] as unknown[] | undefined) ?? [];
  const out: unknown[] = [];
  for (const c of children) {
    if (c && typeof c === "object" && !Array.isArray(c)) {
      const obj = c as Record<string, unknown>;
      if (ooxml.getTag(obj) === "p:cNvPr") {
        out.push(rebuildCNvPr(obj, shape.cNvPrId, shape.name));
        continue;
      }
    }
    out.push(c);
  }
  const next: Record<string, unknown> = { "p:nvPicPr": out };
  const attrs = entry[ATTR_KEY];
  if (attrs && typeof attrs === "object") next[ATTR_KEY] = { ...(attrs as Record<string, unknown>) };
  return next;
}

function rebuildCNvPr(entry: Record<string, unknown>, id: number, name: string): Record<string, unknown> {
  const next: Record<string, unknown> = { "p:cNvPr": entry["p:cNvPr"] ?? [] };
  const prevAttrs = entry[ATTR_KEY];
  const attrs: Record<string, string> = {};
  if (prevAttrs && typeof prevAttrs === "object") {
    for (const [k, v] of Object.entries(prevAttrs as Record<string, unknown>)) {
      attrs[k] = String(v);
    }
  }
  attrs["@_id"] = String(id);
  attrs["@_name"] = name;
  next[ATTR_KEY] = attrs;
  return next;
}

function patchSpPr(entry: Record<string, unknown>, shape: MediaShape): Record<string, unknown> {
  const children = (entry["p:spPr"] as unknown[] | undefined) ?? [];
  const out: unknown[] = [];
  let foundXfrm = false;
  for (const c of children) {
    if (c && typeof c === "object" && !Array.isArray(c)) {
      const obj = c as Record<string, unknown>;
      if (ooxml.getTag(obj) === "a:xfrm") {
        out.push(rebuildXfrm(obj, shape));
        foundXfrm = true;
        continue;
      }
    }
    out.push(c);
  }
  if (!foundXfrm && (shape.position || shape.size || shape.rotation !== undefined)) {
    out.unshift(rebuildXfrm(undefined, shape));
  }
  const next: Record<string, unknown> = { "p:spPr": out };
  const attrs = entry[ATTR_KEY];
  if (attrs && typeof attrs === "object") next[ATTR_KEY] = { ...(attrs as Record<string, unknown>) };
  return next;
}

function rebuildXfrm(
  captured: Record<string, unknown> | undefined,
  shape: MediaShape
): Record<string, unknown> {
  const attrs: Record<string, string> = {};
  if (captured) {
    const a = captured[ATTR_KEY];
    if (a && typeof a === "object") {
      for (const [k, v] of Object.entries(a as Record<string, unknown>)) attrs[k] = String(v);
    }
  }
  if (shape.rotation !== undefined && Number.isFinite(shape.rotation)) {
    const normalised = ((shape.rotation % 360) + 360) % 360;
    if (normalised === 0) {
      delete attrs["@_rot"];
    } else {
      attrs["@_rot"] = String(Math.round(normalised * 60000));
    }
  } else {
    delete attrs["@_rot"];
  }
  const sub: unknown[] = [];
  if (shape.position) {
    sub.push({
      "a:off": [],
      [ATTR_KEY]: {
        "@_x": String(shape.position.xEmu),
        "@_y": String(shape.position.yEmu),
      },
    });
  }
  if (shape.size) {
    sub.push({
      "a:ext": [],
      [ATTR_KEY]: {
        "@_cx": String(shape.size.cxEmu),
        "@_cy": String(shape.size.cyEmu),
      },
    });
  }
  // If neither position nor size were carried by the typed model and
  // we had a captured xfrm, fall through to the captured subtree so
  // we don't accidentally drop e.g. flipH/flipV-only xfrms.
  if (sub.length === 0 && captured) {
    const inner = (captured["a:xfrm"] as unknown[] | undefined) ?? [];
    for (const c of inner) sub.push(c);
  }
  const next: Record<string, unknown> = { "a:xfrm": sub };
  if (Object.keys(attrs).length > 0) next[ATTR_KEY] = attrs;
  return next;
}

/**
 * Build the `<p:pic>` raw blob for a freshly-inserted media shape.
 * Centralised here so both `commands/insert-media.ts` and any future
 * `pptx:set-media-*` commands stay structurally consistent.
 */
export interface BuildMediaPicOptions {
  readonly cNvPrId: number;
  readonly name: string;
  readonly mediaType: "video" | "audio";
  readonly mediaContentType: string;
  readonly mediaRelId: string;
  readonly posterRelId: string;
  readonly position: { xEmu: number; yEmu: number };
  readonly size: { cxEmu: number; cyEmu: number };
}

export function buildMediaPicRaw(opts: BuildMediaPicOptions): OpaqueXml {
  const mediaTag = opts.mediaType === "video" ? "a:videoFile" : "a:audioFile";
  const subtree: unknown[] = [
    {
      "p:nvPicPr": [
        {
          "p:cNvPr": [],
          [ATTR_KEY]: { "@_id": String(opts.cNvPrId), "@_name": opts.name },
        },
        { "p:cNvPicPr": [{ "a:picLocks": [], [ATTR_KEY]: { "@_noChangeAspect": "1" } }] },
        {
          "p:nvPr": [
            {
              [mediaTag]: [],
              [ATTR_KEY]: {
                "@_r:link": opts.mediaRelId,
                "@_contentType": opts.mediaContentType,
              },
            },
          ],
        },
      ],
    },
    {
      "p:blipFill": [
        { "a:blip": [], [ATTR_KEY]: { "@_r:embed": opts.posterRelId } },
        { "a:stretch": [{ "a:fillRect": [] }] },
      ],
    },
    {
      "p:spPr": [
        {
          "a:xfrm": [
            {
              "a:off": [],
              [ATTR_KEY]: { "@_x": String(opts.position.xEmu), "@_y": String(opts.position.yEmu) },
            },
            {
              "a:ext": [],
              [ATTR_KEY]: { "@_cx": String(opts.size.cxEmu), "@_cy": String(opts.size.cyEmu) },
            },
          ],
        },
        { "a:prstGeom": [{ "a:avLst": [] }], [ATTR_KEY]: { "@_prst": "rect" } },
      ],
    },
  ];
  return {
    tag: "p:pic",
    attrs: {},
    rawAttrs: {},
    subtree,
  };
}
