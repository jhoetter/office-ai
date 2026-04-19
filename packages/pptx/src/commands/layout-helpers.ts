/**
 * Helpers for resolving and cloning slide layouts.
 *
 * A "layout" in PPTX is a `<p:sldLayout>` part referenced from a slide's
 * `_rels` file. When a user picks a layout we need three pieces of glue:
 *
 *   1) Resolve a `LayoutKind` (or partPath) to an actual `SlideLayout`,
 *      synthesising a built-in layout part on demand if the deck doesn't
 *      have one of that kind yet.
 *   2) Stamp the layout's placeholders as concrete `TextShape`s on the
 *      target slide so the user has prompts to type into.
 *   3) Wire up the slide → layout relationship + content-types override
 *      when a synthesised layout is added.
 */

import type { HandlerContext } from "@officeai/core";
import type {
  ContentTypesSnap,
  LayoutKind,
  OpaqueXml,
  PlaceholderSpec,
  PptxSnapshot,
  RelationshipsSnap,
  Slide,
  SlideLayout,
  TextShape,
  TextParagraph,
} from "../model/types.js";
import { BUILTIN_LAYOUTS } from "../layouts/builtin.js";
import { parseSlideLayoutFromXml } from "../parser/parse.js";
import type { LayoutKindPayload } from "./payloads.js";

const LAYOUT_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";

const REL_TYPE_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const REL_TYPE_SLIDE_LAYOUT_FROM_MASTER =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const REL_TYPE_SLIDE_MASTER =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";

const SLIDE_RELS_PREFIX = "ppt/slides/_rels/";
const LAYOUT_DIR = "ppt/slideLayouts/";

export interface LayoutResolution {
  /** The (possibly newly synthesised) layout to apply. */
  readonly layout: SlideLayout;
  /** When set, the snapshot must be augmented with this new part. */
  readonly added?: AddedLayout;
}

export interface AddedLayout {
  readonly partPath: string;
  readonly xml: string;
  readonly relsPath: string;
  readonly entries: ReadonlyArray<{ id: string; type: string; target: string }>;
  readonly contentTypeOverride: { partName: string; contentType: string };
  readonly masterRelTarget?: { masterPartPath: string; relId: string; target: string };
}

/**
 * Find a layout in the deck matching `kind`, or produce a synthesised
 * built-in layout part the caller can attach. Returns the resolved layout
 * + (when synthesised) all the side-effects needed to install it.
 *
 * When the deck already carries a layout of the requested kind but it has
 * NO placeholder slots (e.g. minimal sample decks ship a `type="title"`
 * layout with an empty `<p:spTree>`), we still want the user-visible
 * placeholder hints ("Click to add title", …) to appear when the user
 * picks that layout from the menu. We layer the built-in template's
 * placeholder specs onto the existing layout in-memory so cloning stamps
 * them onto the slide; the on-disk layout part is left untouched, which
 * keeps the deck's other layout-specific data (theme, colour map,
 * relationships) intact.
 */
export function resolveLayoutForKind(snapshot: PptxSnapshot, kind: LayoutKindPayload): LayoutResolution {
  for (const l of snapshot.root.layouts.values()) {
    if (l.kind !== (kind as LayoutKind)) continue;
    if (l.placeholders.length > 0) return { layout: l };
    const enriched = builtinPlaceholdersFor(kind, snapshot.root.slideSize);
    if (enriched.length === 0) return { layout: l };
    return { layout: { ...l, placeholders: enriched } };
  }
  return synthesiseBuiltinLayout(snapshot, kind);
}

/**
 * Parse the built-in template for `kind` purely to recover its
 * placeholder specs. The synthesised layout is discarded — only the
 * placeholders are returned, so callers can stamp them onto a slide
 * without installing a duplicate layout part.
 *
 * Built-in templates are authored against the standard 16:9 footprint
 * (12_192_000 × 6_858_000 EMU). When the deck is a different size —
 * 4:3 (9_144_000 × 6_858_000) is the common case — we proportionally
 * rescale every placeholder rectangle so the cloned shapes stay
 * inside the slide bounds. Without this, a deck-with-empty-title-
 * layout would produce placeholders whose right edge sat ~3 inches
 * past the slide.
 */
function builtinPlaceholdersFor(
  kind: LayoutKindPayload,
  slideSize: { readonly cxEmu: number; readonly cyEmu: number }
): ReadonlyArray<PlaceholderSpec> {
  const tmpl = BUILTIN_LAYOUTS[kind];
  if (!tmpl) return [];
  const layout = parseSlideLayoutFromXml(`${LAYOUT_DIR}__builtin_${kind}.xml`, tmpl.xml);
  return rescalePlaceholders(layout.placeholders, slideSize);
}

const BUILTIN_TEMPLATE_W = 12_192_000;
const BUILTIN_TEMPLATE_H = 6_858_000;

function rescalePlaceholders(
  phs: ReadonlyArray<PlaceholderSpec>,
  slideSize: { readonly cxEmu: number; readonly cyEmu: number }
): ReadonlyArray<PlaceholderSpec> {
  const sx = slideSize.cxEmu / BUILTIN_TEMPLATE_W;
  const sy = slideSize.cyEmu / BUILTIN_TEMPLATE_H;
  if (sx === 1 && sy === 1) return phs;
  return phs.map((ph) => ({
    ...ph,
    ...(ph.position
      ? {
          position: {
            xEmu: Math.round(ph.position.xEmu * sx),
            yEmu: Math.round(ph.position.yEmu * sy),
          },
        }
      : {}),
    ...(ph.size
      ? {
          size: {
            cxEmu: Math.round(ph.size.cxEmu * sx),
            cyEmu: Math.round(ph.size.cyEmu * sy),
          },
        }
      : {}),
  }));
}

function synthesiseBuiltinLayout(snapshot: PptxSnapshot, kind: LayoutKindPayload): LayoutResolution {
  const tmpl = BUILTIN_LAYOUTS[kind];
  // Mint a fresh layout part path that doesn't collide with existing ones.
  let n = 1;
  const taken = new Set<string>();
  for (const p of snapshot.root.layouts.keys()) taken.add(p);
  let partPath = `${LAYOUT_DIR}slideLayout${n}.xml`;
  while (taken.has(partPath)) {
    n++;
    partPath = `${LAYOUT_DIR}slideLayout${n}.xml`;
  }
  const parsed = parseSlideLayoutFromXml(partPath, tmpl.xml);
  // Built-in templates target a 12.192M × 6.858M canvas; rescale to
  // match the deck so placeholders land inside the slide bounds (see
  // `builtinPlaceholdersFor` for the rationale). The on-disk XML is
  // left at the template's coordinates — PowerPoint reflows layout
  // placeholders against the deck's slide size on open, and our own
  // renderer pulls coordinates from the in-memory `placeholders`
  // array (which we rescale here), so this stays consistent without
  // having to rewrite the XML.
  const layout: SlideLayout = {
    ...parsed,
    placeholders: rescalePlaceholders(parsed.placeholders, snapshot.root.slideSize),
  };
  const relsPath = relsPathFor(partPath);

  // Pick a master to point the new layout at — every layout part needs
  // exactly one slideMaster relationship so PowerPoint can resolve its
  // theme + colour map. We use whatever master comes first (decks
  // virtually always have just one).
  const masterPaths = [...snapshot.root.masters.keys()];
  const masterPartPath = masterPaths[0];
  const entries = masterPartPath
    ? [
        {
          id: "rId1",
          type: REL_TYPE_SLIDE_MASTER,
          target: relativeFrom(relsPath, masterPartPath),
        },
      ]
    : [];

  // The master also needs a back-reference rel pointing at the new
  // layout — without it PowerPoint refuses to surface the layout in the
  // picker. We compute the rel id later when attaching.
  const masterRelTarget = masterPartPath
    ? {
        masterPartPath,
        relId: "",
        target: relativeFrom(relsPathFor(masterPartPath), partPath),
      }
    : undefined;

  return {
    layout,
    added: {
      partPath,
      xml: tmpl.xml,
      relsPath,
      entries,
      contentTypeOverride: { partName: `/${partPath}`, contentType: LAYOUT_CONTENT_TYPE },
      ...(masterRelTarget ? { masterRelTarget } : {}),
    },
  };
}

/**
 * Stamp a layout's placeholders into a slide as concrete `TextShape`s.
 * Existing shapes on the slide are preserved when their placeholder idx
 * matches one in the new layout (so swapping layouts doesn't wipe user
 * content); placeholders without a matching existing shape get an empty
 * prompt.
 */
export function clonePlaceholdersIntoSlide(
  slide: Slide,
  layout: SlideLayout,
  ctx: HandlerContext,
  baseCNvPrId: number
): Slide {
  // Index existing placeholder shapes by their `idx` so we keep user
  // content where the placeholder slot still exists in the new layout.
  const existingByIdx = new Map<number, TextShape>();
  for (const s of slide.shapes) {
    if (s.kind !== "text" || !s.placeholder) continue;
    const idx = s.placeholder.idx ?? 0;
    existingByIdx.set(idx, s);
  }

  const cloned: TextShape[] = [];
  let nextId = baseCNvPrId;
  for (const ph of layout.placeholders) {
    const existing = existingByIdx.get(ph.idx);
    if (existing) {
      // Re-stamp the position/size from the new layout — but keep the
      // existing text body so users don't lose what they typed.
      cloned.push({
        ...existing,
        ...(ph.position ? { position: ph.position } : {}),
        ...(ph.size ? { size: ph.size } : {}),
        placeholder: {
          ...(ph.type ? { type: ph.type } : {}),
          idx: ph.idx,
        },
      });
      continue;
    }
    nextId++;
    cloned.push(buildPlaceholderShape(ph, ctx, nextId));
  }

  // Anything else on the slide (non-placeholder shapes — pictures users
  // dropped in, connectors, etc.) is preserved verbatim.
  const nonPlaceholders = slide.shapes.filter((s) => s.kind !== "text" || !s.placeholder);
  return { ...slide, shapes: [...cloned, ...nonPlaceholders] };
}

function buildPlaceholderShape(ph: PlaceholderSpec, ctx: HandlerContext, cNvPrId: number): TextShape {
  const para: TextParagraph = {
    id: ctx.mintNodeId(),
    properties: {},
    runs: [],
  };
  return {
    kind: "text",
    id: ctx.mintNodeId(),
    cNvPrId,
    name: placeholderName(ph),
    placeholder: { type: ph.type, idx: ph.idx },
    ...(ph.position ? { position: ph.position } : {}),
    ...(ph.size ? { size: ph.size } : {}),
    nvSpPrTail: defaultPlaceholderNvSpPrTail(ph, cNvPrId),
    spPrTail: [],
    txBody: { paragraphs: [para] },
  };
}

function placeholderName(ph: PlaceholderSpec): string {
  switch (ph.type) {
    case "title":
    case "ctrTitle":
      return "Title";
    case "subTitle":
      return "Subtitle";
    case "body":
      return "Content placeholder";
    case "pic":
      return "Picture placeholder";
    default:
      return ph.type ? `${ph.type} placeholder` : "Placeholder";
  }
}

function defaultPlaceholderNvSpPrTail(ph: PlaceholderSpec, cNvPrId: number): OpaqueXml[] {
  const phAttrs: Record<string, string> = { type: ph.type };
  const phRawAttrs: Record<string, string> = { "@_type": ph.type };
  if (ph.idx !== undefined) {
    phAttrs.idx = String(ph.idx);
    phRawAttrs["@_idx"] = String(ph.idx);
  }
  if (ph.sz) {
    phAttrs.sz = ph.sz;
    phRawAttrs["@_sz"] = ph.sz;
  }
  return [
    {
      tag: "p:cNvPr",
      attrs: { id: String(cNvPrId), name: placeholderName(ph) },
      rawAttrs: { "@_id": String(cNvPrId), "@_name": placeholderName(ph) },
      subtree: [],
    },
    {
      tag: "p:cNvSpPr",
      attrs: {},
      rawAttrs: {},
      subtree: [{ "a:spLocks": [], ":@": { "@_noGrp": "1" } }],
    },
    {
      tag: "p:nvPr",
      attrs: {},
      rawAttrs: {},
      subtree: [{ "p:ph": [], ":@": phRawAttrs }],
    },
  ];
}

/**
 * Apply an `AddedLayout` to the snapshot's auxiliary maps + dirty
 * buckets. Mutates copies, returns updated structures the caller can
 * pass back into `evolveSnapshot`.
 */
export interface ApplyAddedLayoutResult {
  readonly layouts: ReadonlyMap<string, SlideLayout>;
  readonly relationships: ReadonlyMap<string, RelationshipsSnap>;
  readonly contentTypes: ContentTypesSnap;
  readonly dirtyLayouts: ReadonlyArray<string>;
  readonly dirtyRels: ReadonlyArray<string>;
  readonly dirtyContentTypes: boolean;
}

export function applyAddedLayout(
  snapshot: PptxSnapshot,
  added: AddedLayout,
  layout: SlideLayout
): ApplyAddedLayoutResult {
  // Layouts map.
  const layouts = new Map(snapshot.root.layouts);
  layouts.set(added.partPath, layout);

  // Relationships: layout → master.
  const relationships = new Map(snapshot.relationships);
  relationships.set(added.relsPath, {
    relsPath: added.relsPath,
    entries: added.entries.map((e) => ({ id: e.id, type: e.type, target: e.target })),
  });
  const dirtyRels: string[] = [added.relsPath];

  // Master → layout back-reference (so the layout is discoverable).
  if (added.masterRelTarget) {
    const masterRelsPath = relsPathFor(added.masterRelTarget.masterPartPath);
    const existing = relationships.get(masterRelsPath);
    const existingEntries = existing?.entries ?? [];
    const relId = nextRelId(existingEntries.map((e) => e.id));
    const newEntries = [
      ...existingEntries,
      {
        id: relId,
        type: REL_TYPE_SLIDE_LAYOUT_FROM_MASTER,
        target: added.masterRelTarget.target,
      },
    ];
    relationships.set(masterRelsPath, { relsPath: masterRelsPath, entries: newEntries });
    dirtyRels.push(masterRelsPath);
  }

  // Content types: add the layout override unless it's already present.
  const overrideExists = snapshot.contentTypes.overrides.some(
    (o) => o.partName === added.contentTypeOverride.partName
  );
  const contentTypes: ContentTypesSnap = overrideExists
    ? snapshot.contentTypes
    : {
        ...snapshot.contentTypes,
        overrides: [...snapshot.contentTypes.overrides, added.contentTypeOverride],
      };

  return {
    layouts,
    relationships,
    contentTypes,
    dirtyLayouts: [added.partPath],
    dirtyRels,
    dirtyContentTypes: !overrideExists,
  };
}

/**
 * Update or create the slide → layout relationship so the slide points
 * at `layoutPartPath`. Returns the new relationships map + the slide-rels
 * path for the dirty bucket.
 */
export function setSlideLayoutRel(
  snapshot: PptxSnapshot,
  slide: Slide,
  layoutPartPath: string
): { relationships: ReadonlyMap<string, RelationshipsSnap>; relsPath: string } {
  const relsPath = `${SLIDE_RELS_PREFIX}${slide.partPath.split("/").pop()}.rels`;
  const relationships = new Map(snapshot.relationships);
  const existing = relationships.get(relsPath);
  const target = relativeFrom(relsPath, layoutPartPath);
  if (!existing) {
    relationships.set(relsPath, {
      relsPath,
      entries: [{ id: "rId1", type: REL_TYPE_LAYOUT, target }],
    });
    return { relationships, relsPath };
  }
  const layoutEntry = existing.entries.find((e) => e.type === REL_TYPE_LAYOUT);
  if (layoutEntry) {
    relationships.set(relsPath, {
      relsPath,
      entries: existing.entries.map((e) => (e === layoutEntry ? { ...e, target } : e)),
    });
  } else {
    const relId = nextRelId(existing.entries.map((e) => e.id));
    relationships.set(relsPath, {
      relsPath,
      entries: [...existing.entries, { id: relId, type: REL_TYPE_LAYOUT, target }],
    });
  }
  return { relationships, relsPath };
}

/**
 * Mirror of `add-slide.ts`'s helper — kept private to this module to
 * avoid pulling that file's surface into the layout-helpers boundary.
 */
function nextRelId(existing: ReadonlyArray<string>): string {
  let max = 0;
  for (const id of existing) {
    const m = /^rId(\d+)$/.exec(id);
    if (m) {
      const n = Number(m[1]);
      if (n > max) max = n;
    }
  }
  return `rId${max + 1}`;
}

function relsPathFor(partPath: string): string {
  const lastSlash = partPath.lastIndexOf("/");
  const dir = partPath.slice(0, lastSlash);
  const file = partPath.slice(lastSlash + 1);
  return `${dir}/_rels/${file}.rels`;
}

function relativeFrom(relsPath: string, targetPath: string): string {
  const ownerDir = relsPathOwnerDir(relsPath);
  const targetSegments = targetPath.split("/");
  const ownerSegments = ownerDir.split("/").filter((s) => s.length > 0);
  let i = 0;
  while (
    i < ownerSegments.length &&
    i < targetSegments.length - 1 &&
    ownerSegments[i] === targetSegments[i]
  ) {
    i++;
  }
  const ups = ownerSegments.length - i;
  const downs = targetSegments.slice(i);
  return [...Array(ups).fill(".."), ...downs].join("/");
}

function relsPathOwnerDir(relsPath: string): string {
  const m = /^(.*?)_rels\/[^/]+\.rels$/.exec(relsPath);
  if (!m) return "";
  return (m[1] ?? "").replace(/\/$/, "");
}
