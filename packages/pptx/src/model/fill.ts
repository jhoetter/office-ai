import type { OpaqueXml } from "./types.js";

/**
 * Discriminated union describing every fill the editor can author.
 *
 * Mirrors the OOXML choice group `EG_FillProperties`:
 *   - `none`     → `<a:noFill/>`
 *   - `solid`    → `<a:solidFill><a:srgbClr val="…"/>(<a:alpha …/>)?</a:solidFill>`
 *   - `gradient` → `<a:gradFill rotWithShape="1"><a:gsLst>…</a:gsLst><a:lin/|<a:path/></a:gradFill>`
 *   - `pattern`  → `<a:pattFill prst="…"><a:fgClr/><a:bgClr/></a:pattFill>`
 *   - `picture`  → `<a:blipFill><a:blip r:embed="…"/><a:srcRect/><a:stretch|tile/></a:blipFill>`
 *
 * Used by both the slide-background command and the shape-fill command
 * so the same `FillPicker` UI can drive either one.
 */
export type FillSpec =
  | { readonly type: "none" }
  | SolidFillSpec
  | GradientFillSpec
  | PatternFillSpec
  | PictureFillSpec;

export interface SolidFillSpec {
  readonly type: "solid";
  /** Lowercase or uppercase RRGGBB (no `#`). The serializer normalises. */
  readonly color: string;
  /** Opacity in `[0, 1]`. `undefined` means fully opaque. */
  readonly alpha?: number;
}

export type GradientKind = "linear" | "radial";

export interface GradientFillSpec {
  readonly type: "gradient";
  readonly kind: GradientKind;
  /**
   * Direction in degrees clockwise from horizontal-right (12-o-clock = 270).
   * Ignored when `kind === "radial"`. OOXML expresses this as
   * 60_000-units of an angle on `<a:lin ang="…"/>`.
   */
  readonly angleDeg: number;
  /**
   * 2..10 stops, sorted by `pos` ascending. Each stop's `pos` lives in
   * `[0, 1]`; the OOXML form multiplies by 100_000.
   */
  readonly stops: ReadonlyArray<GradientStop>;
}

export interface GradientStop {
  readonly pos: number;
  readonly color: string;
  readonly alpha?: number;
}

/**
 * `<a:pattFill prst="…">` — see ECMA-376 §20.1.8.20 for the full enum.
 * The picker exposes a curated subset; serialization accepts any string.
 */
export type PatternPreset =
  | "pct5"
  | "pct10"
  | "pct20"
  | "pct25"
  | "pct30"
  | "pct40"
  | "pct50"
  | "pct60"
  | "pct70"
  | "pct75"
  | "pct80"
  | "pct90"
  | "horz"
  | "vert"
  | "ltHorz"
  | "ltVert"
  | "dkHorz"
  | "dkVert"
  | "narHorz"
  | "narVert"
  | "dashHorz"
  | "dashVert"
  | "cross"
  | "dnDiag"
  | "upDiag"
  | "ltDnDiag"
  | "ltUpDiag"
  | "dkDnDiag"
  | "dkUpDiag"
  | "wdDnDiag"
  | "wdUpDiag"
  | "dashDnDiag"
  | "dashUpDiag"
  | "diagCross"
  | "smCheck"
  | "lgCheck"
  | "smGrid"
  | "lgGrid"
  | "dotGrid"
  | "smConfetti"
  | "lgConfetti"
  | "horzBrick"
  | "diagBrick"
  | "solidDmnd"
  | "openDmnd"
  | "dotDmnd"
  | "plaid"
  | "sphere"
  | "weave"
  | "divot"
  | "shingle"
  | "wave"
  | "trellis"
  | "zigZag";

export interface PatternFillSpec {
  readonly type: "pattern";
  readonly preset: PatternPreset;
  readonly fgColor: string;
  readonly bgColor: string;
}

export interface PictureFillSpec {
  readonly type: "picture";
  /** Slide-rels relationship id pointing at a media part already in the deck. */
  readonly embedRelId: string;
  /** When true, the bitmap tiles inside the shape; otherwise it stretches. */
  readonly tile?: boolean;
}

// ─── Normalisation ────────────────────────────────────────────────────────

const HEX_RE = /^[0-9a-fA-F]{6}$/;

export function normaliseHex(input: string): string {
  const v = input.trim().replace(/^#/, "").toUpperCase();
  if (!HEX_RE.test(v)) throw new Error(`invalid hex color: ${input}`);
  return v;
}

function clamp01(n: number): number {
  if (!Number.isFinite(n)) return 0;
  if (n < 0) return 0;
  if (n > 1) return 1;
  return n;
}

/** Round-trip-safe normalisation; throws on malformed input. */
export function normaliseFillSpec(spec: FillSpec): FillSpec {
  switch (spec.type) {
    case "none":
      return spec;
    case "solid":
      return { type: "solid", color: normaliseHex(spec.color), alpha: normaliseAlpha(spec.alpha) };
    case "gradient": {
      if (spec.stops.length < 2) throw new Error("gradient requires at least 2 stops");
      if (spec.stops.length > 10) throw new Error("gradient supports at most 10 stops");
      const stops = spec.stops
        .map((s) => ({ pos: clamp01(s.pos), color: normaliseHex(s.color), alpha: normaliseAlpha(s.alpha) }))
        .sort((a, b) => a.pos - b.pos);
      const angle = ((Number(spec.angleDeg) % 360) + 360) % 360;
      return { type: "gradient", kind: spec.kind, angleDeg: angle, stops };
    }
    case "pattern":
      return {
        type: "pattern",
        preset: spec.preset,
        fgColor: normaliseHex(spec.fgColor),
        bgColor: normaliseHex(spec.bgColor),
      };
    case "picture":
      return { type: "picture", embedRelId: spec.embedRelId, tile: spec.tile === true ? true : undefined };
    default: {
      const _exhaustive: never = spec;
      void _exhaustive;
      throw new Error(`unknown fill type`);
    }
  }
}

function normaliseAlpha(a: number | undefined): number | undefined {
  if (a === undefined) return undefined;
  const v = clamp01(a);
  return v >= 1 ? undefined : v;
}

// ─── Serialisation: FillSpec → OpaqueXml ──────────────────────────────────

/**
 * Build the single `<a:solidFill>` / `<a:gradFill>` / … node that should
 * splice into a shape's `spPrTail` (or a slide's `<p:bgPr>`). Returns
 * `null` for `{type:"none"}` callers that want to emit `<a:noFill>`
 * inline; use `fillSpecToOpaqueOrNoFill` for that case.
 */
export function fillSpecToOpaque(spec: FillSpec): OpaqueXml {
  const normalised = normaliseFillSpec(spec);
  switch (normalised.type) {
    case "none":
      return { tag: "a:noFill", attrs: {}, rawAttrs: {}, subtree: [] };
    case "solid":
      return {
        tag: "a:solidFill",
        attrs: {},
        rawAttrs: {},
        subtree: [solidColorChild(normalised.color, normalised.alpha)],
      };
    case "gradient":
      return {
        tag: "a:gradFill",
        attrs: { rotWithShape: "1", flip: "none" },
        rawAttrs: { "@_rotWithShape": "1", "@_flip": "none" },
        subtree: gradFillChildren(normalised),
      };
    case "pattern":
      return {
        tag: "a:pattFill",
        attrs: { prst: normalised.preset },
        rawAttrs: { "@_prst": normalised.preset },
        subtree: [
          {
            "a:fgClr": [solidColorChild(normalised.fgColor)],
          },
          {
            "a:bgClr": [solidColorChild(normalised.bgColor)],
          },
        ],
      };
    case "picture": {
      const blipChild = {
        "a:blip": [],
        ":@": { "@_r:embed": normalised.embedRelId },
      };
      const stretchOrTile = normalised.tile
        ? {
            "a:tile": [],
            ":@": {
              "@_tx": "0",
              "@_ty": "0",
              "@_sx": "100000",
              "@_sy": "100000",
              "@_flip": "none",
              "@_algn": "tl",
            },
          }
        : { "a:stretch": [{ "a:fillRect": [] }] };
      return {
        tag: "a:blipFill",
        attrs: {},
        rawAttrs: {},
        subtree: [blipChild, stretchOrTile],
      };
    }
    default: {
      const _exhaustive: never = normalised;
      void _exhaustive;
      throw new Error("unreachable");
    }
  }
}

function solidColorChild(hex: string, alpha?: number): unknown {
  const inner: unknown[] = [];
  if (alpha !== undefined) {
    inner.push({ "a:alpha": [], ":@": { "@_val": String(Math.round(alpha * 100_000)) } });
  }
  return {
    "a:srgbClr": inner,
    ":@": { "@_val": hex },
  };
}

function gradFillChildren(spec: GradientFillSpec): unknown[] {
  const gsLst = spec.stops.map((s) => ({
    "a:gs": [solidColorChild(s.color, s.alpha)],
    ":@": { "@_pos": String(Math.round(s.pos * 100_000)) },
  }));
  if (spec.kind === "linear") {
    // OOXML angles are 60_000ths of a degree, measured clockwise from
    // 3-o-clock. Our model already uses degrees clockwise from
    // horizontal-right so the conversion is a straight scale.
    const angleOoxml = String(Math.round(spec.angleDeg * 60_000));
    return [{ "a:gsLst": gsLst }, { "a:lin": [], ":@": { "@_ang": angleOoxml, "@_scaled": "0" } }];
  }
  // Radial. PowerPoint uses `<a:path path="circle"><a:fillToRect/></a:path>`.
  return [
    { "a:gsLst": gsLst },
    {
      "a:path": [
        { "a:fillToRect": [], ":@": { "@_l": "50000", "@_t": "50000", "@_r": "50000", "@_b": "50000" } },
      ],
      ":@": { "@_path": "circle" },
    },
  ];
}

// ─── Parsing: OpaqueXml → FillSpec ────────────────────────────────────────

/**
 * Walk a list of `spPrTail` (or `<p:bgPr>`) children and return the
 * first recognised fill. Returns `null` when no fill node is present
 * (the caller should treat this as "inherits from layout / theme").
 *
 * Theme / scheme colours collapse to their literal `lastClr` if
 * available, otherwise the `val` is returned verbatim — the renderer
 * still needs the theme to actually paint them.
 */
export function readFillSpec(children: ReadonlyArray<OpaqueXml>): FillSpec | null {
  for (const node of children) {
    switch (node.tag) {
      case "a:noFill":
        return { type: "none" };
      case "a:solidFill": {
        const c = readSolidColor(node.subtree);
        if (c) return { type: "solid", color: c.color, alpha: c.alpha };
        break;
      }
      case "a:gradFill": {
        const grad = readGradFill(node.subtree);
        if (grad) return grad;
        break;
      }
      case "a:pattFill": {
        const patt = readPattFill(node);
        if (patt) return patt;
        break;
      }
      case "a:blipFill": {
        const pic = readBlipFill(node.subtree);
        if (pic) return pic;
        break;
      }
      default:
        break;
    }
  }
  return null;
}

interface ReadColor {
  color: string;
  alpha: number | undefined;
}

function readSolidColor(subtree: ReadonlyArray<unknown>): ReadColor | null {
  for (const inner of subtree) {
    const node = inner as Record<string, unknown> | null;
    if (!node || typeof node !== "object" || Array.isArray(node)) continue;
    const keys = Object.keys(node).filter((k) => k !== ":@");
    if (keys.length !== 1) continue;
    const tag = keys[0];
    const attrs = node[":@"] as Record<string, unknown> | undefined;
    const val = attrs?.["@_val"];
    if (typeof val !== "string") continue;
    let color: string | null = null;
    if (tag === "a:srgbClr") color = val;
    else if (tag === "a:sysClr") {
      const last = attrs?.["@_lastClr"];
      color = typeof last === "string" ? last : val;
    } else if (tag === "a:schemeClr") {
      // Best-effort: keep the scheme name's lastClr if encoded, otherwise
      // bail. The caller can still see the *kind* of fill is solid.
      color = "808080";
    }
    if (!color) continue;
    return { color: normaliseHex(color), alpha: readAlphaFromChildren(node[tag]) };
  }
  return null;
}

function readAlphaFromChildren(subtree: unknown): number | undefined {
  if (!Array.isArray(subtree)) return undefined;
  for (const inner of subtree) {
    const node = inner as Record<string, unknown> | null;
    if (!node || typeof node !== "object" || Array.isArray(node)) continue;
    if ("a:alpha" in node) {
      const attrs = node[":@"] as Record<string, unknown> | undefined;
      const v = attrs?.["@_val"];
      if (typeof v === "string") {
        const n = Number(v);
        if (Number.isFinite(n)) return Math.max(0, Math.min(1, n / 100_000));
      }
    }
  }
  return undefined;
}

function readGradFill(subtree: ReadonlyArray<unknown>): GradientFillSpec | null {
  const stops: GradientStop[] = [];
  let kind: GradientKind = "linear";
  let angleDeg = 0;
  for (const inner of subtree) {
    const node = inner as Record<string, unknown> | null;
    if (!node || typeof node !== "object" || Array.isArray(node)) continue;
    if ("a:gsLst" in node) {
      const list = node["a:gsLst"];
      if (!Array.isArray(list)) continue;
      for (const gs of list) {
        const obj = gs as Record<string, unknown> | null;
        if (!obj || typeof obj !== "object") continue;
        const attrs = obj[":@"] as Record<string, unknown> | undefined;
        const posRaw = attrs?.["@_pos"];
        const pos = typeof posRaw === "string" ? Number(posRaw) / 100_000 : 0;
        const color = readSolidColor(Array.isArray(obj["a:gs"]) ? (obj["a:gs"] as unknown[]) : []);
        if (color) stops.push({ pos, color: color.color, alpha: color.alpha });
      }
    } else if ("a:lin" in node) {
      kind = "linear";
      const attrs = node[":@"] as Record<string, unknown> | undefined;
      const ang = attrs?.["@_ang"];
      if (typeof ang === "string") {
        const n = Number(ang);
        if (Number.isFinite(n)) angleDeg = (((n / 60_000) % 360) + 360) % 360;
      }
    } else if ("a:path" in node) {
      const attrs = node[":@"] as Record<string, unknown> | undefined;
      if (attrs?.["@_path"] === "circle") kind = "radial";
    }
  }
  if (stops.length < 2) return null;
  stops.sort((a, b) => a.pos - b.pos);
  return { type: "gradient", kind, angleDeg, stops };
}

function readPattFill(node: OpaqueXml): PatternFillSpec | null {
  const preset = (node.attrs.prst ?? node.rawAttrs["@_prst"]) as PatternPreset | undefined;
  if (!preset) return null;
  let fg: string | null = null;
  let bg: string | null = null;
  for (const inner of node.subtree) {
    const obj = inner as Record<string, unknown> | null;
    if (!obj || typeof obj !== "object" || Array.isArray(obj)) continue;
    if ("a:fgClr" in obj) {
      const c = readSolidColor(Array.isArray(obj["a:fgClr"]) ? (obj["a:fgClr"] as unknown[]) : []);
      if (c) fg = c.color;
    } else if ("a:bgClr" in obj) {
      const c = readSolidColor(Array.isArray(obj["a:bgClr"]) ? (obj["a:bgClr"] as unknown[]) : []);
      if (c) bg = c.color;
    }
  }
  if (!fg || !bg) return null;
  return { type: "pattern", preset, fgColor: fg, bgColor: bg };
}

function readBlipFill(subtree: ReadonlyArray<unknown>): PictureFillSpec | null {
  let embed: string | null = null;
  let tile = false;
  for (const inner of subtree) {
    const obj = inner as Record<string, unknown> | null;
    if (!obj || typeof obj !== "object" || Array.isArray(obj)) continue;
    if ("a:blip" in obj) {
      const attrs = obj[":@"] as Record<string, unknown> | undefined;
      const v = attrs?.["@_r:embed"];
      if (typeof v === "string") embed = v;
    } else if ("a:tile" in obj) {
      tile = true;
    }
  }
  if (!embed) return null;
  return { type: "picture", embedRelId: embed, tile: tile || undefined };
}

// ─── Splice helper ────────────────────────────────────────────────────────

/**
 * Replace any existing fill node in `spPrTail` with the one derived from
 * `spec`. The result preserves PowerPoint's invariant that at most one
 * of `solidFill / noFill / gradFill / pattFill / blipFill` lives at the
 * `<a:spPr>` level. The replacement is inserted immediately after
 * `<a:prstGeom>` when present (PowerPoint refuses to render the fill
 * otherwise), otherwise prepended.
 *
 * Pass `spec === null` to clear the fill entirely (the shape inherits
 * from its style — equivalent to PowerPoint's "Fill: Automatic").
 */
export function spliceFillIntoSpPr(
  tail: ReadonlyArray<OpaqueXml>,
  spec: FillSpec | null
): ReadonlyArray<OpaqueXml> {
  const filtered = tail.filter(
    (c) =>
      c.tag !== "a:solidFill" &&
      c.tag !== "a:noFill" &&
      c.tag !== "a:gradFill" &&
      c.tag !== "a:pattFill" &&
      c.tag !== "a:blipFill"
  );
  if (spec === null) return filtered;
  const replacement = fillSpecToOpaque(spec);
  const idx = filtered.findIndex((c) => c.tag === "a:prstGeom");
  return idx >= 0
    ? [...filtered.slice(0, idx + 1), replacement, ...filtered.slice(idx + 1)]
    : [replacement, ...filtered];
}

/**
 * Same as `spliceFillIntoSpPr` but for the `<p:cSld>` head opaque tree:
 * walks for an existing `<p:bg>` and replaces (or removes) it. Returns
 * a new array; the input is not mutated.
 *
 * Pass `spec === null` to drop `<p:bg>` entirely (slide inherits from
 * the layout/master, matching PowerPoint's "Reset Background").
 */
export function spliceSlideBackground(
  cSldHead: ReadonlyArray<OpaqueXml>,
  spec: FillSpec | null
): ReadonlyArray<OpaqueXml> {
  const filtered = cSldHead.filter((c) => c.tag !== "p:bg");
  if (spec === null) return filtered;
  const fillNode = fillSpecToOpaque(spec);
  // <p:bg><p:bgPr><…fill…/><a:effectLst/></p:bgPr></p:bg>
  const bg: OpaqueXml = {
    tag: "p:bg",
    attrs: {},
    rawAttrs: {},
    subtree: [
      {
        "p:bgPr": [serializeOpaqueAsObj(fillNode), { "a:effectLst": [] }],
      },
    ],
  };
  // <p:bg> must come BEFORE <p:spTree>; cSldHead conceptually ends just
  // before spTree, so prepending is always safe.
  return [bg, ...filtered];
}

/**
 * Convert an `OpaqueXml` node back to fast-xml-parser's `{tag: [...]}`
 * shape so it can live inside another OpaqueXml `subtree`. Mirrors the
 * inverse of how the parser captures opaque trees.
 */
function serializeOpaqueAsObj(node: OpaqueXml): Record<string, unknown> {
  const result: Record<string, unknown> = { [node.tag]: node.subtree as unknown[] };
  const rawAttrs = Object.keys(node.rawAttrs).length > 0 ? node.rawAttrs : null;
  if (rawAttrs) result[":@"] = { ...rawAttrs };
  return result;
}
