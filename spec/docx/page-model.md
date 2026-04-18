# DOCX page model

> Status: P3.2 spec. Drives W5 (typed `SectionProperties`), W6 (typed
> `PageBreak`), W7 (round-trip), W8 (header/footer reference graph).

## Why

Today `<w:sectPr>` parses to `SectionBreak { kind, id, raw }` — opaque
bytes only. The renderer can't show page boundaries because it has no
typed access to page size, margins, or which header/footer applies.
[`packages/docx/src/parser/parse.ts`](../../packages/docx/src/parser/parse.ts)
line 217.

Every text-flow page break (`<w:br w:type="page"/>`) is also opaque inside
a run today (`OpaqueRunChild`). The renderer needs typed access to split
the body into pages.

## Typed model

Add to `packages/docx/src/model/types.ts`:

```ts
/** OOXML twips (1/20 pt) for distances; positive integers. */
export type Twips = number;

export interface PageSize {
  readonly w: Twips;
  readonly h: Twips;
  readonly orient?: "portrait" | "landscape";
  /** Word also writes `w:code` (size enum). Carried through opaqueAttrs. */
}

export interface PageMargins {
  readonly top: Twips;
  readonly right: Twips;
  readonly bottom: Twips;
  readonly left: Twips;
  readonly header: Twips;
  readonly footer: Twips;
  readonly gutter?: Twips;
}

export interface PageColumns {
  readonly num: number;
  readonly sep?: boolean;
  readonly equalWidth?: boolean;
  readonly space?: Twips;
}

export interface HeaderFooterRef {
  readonly type: "default" | "first" | "even";
  /** `r:id` resolving via `word/_rels/document.xml.rels` to a header/footer part. */
  readonly relationshipId: string;
}

export interface SectionProperties {
  readonly pgSz?: PageSize;
  readonly pgMar?: PageMargins;
  readonly cols?: PageColumns;
  readonly headerRefs: ReadonlyArray<HeaderFooterRef>;
  readonly footerRefs: ReadonlyArray<HeaderFooterRef>;
  readonly titlePg?: boolean;
  /** `<w:type w:val>` — continuous, nextPage, oddPage, evenPage, nextColumn. */
  readonly sectionType?: "continuous" | "nextPage" | "oddPage" | "evenPage" | "nextColumn";
  /** Catch-all for sectPr children we don't model (lineNumbers, pgNumType, etc.). */
  readonly opaqueProps?: ReadonlyArray<OpaqueXml>;
}
```

Update `SectionBreak`:

```ts
export interface SectionBreak {
  readonly kind: "section-break";
  readonly id: NodeId;
  readonly properties: SectionProperties;
  /**
   * Original `<w:sectPr>` subtree. Present on freshly-parsed nodes and
   * on nodes that have not been touched by a mutating command. The
   * serializer emits `raw` verbatim when present so untouched
   * sections round-trip byte-identical. Mutating commands MUST drop
   * `raw` on the `SectionBreak` they produce. Same contract as
   * `Table.raw` / `InlineImageDrawing.raw`.
   */
  readonly raw?: OpaqueXml;
}
```

### Inline page break

New `RunChild` variant:

```ts
export type RunChild =
  | TextLeaf
  | BreakLeaf
  | TabLeaf
  | DrawingLeaf
  | PageBreakLeaf
  | LastRenderedPageBreakLeaf
  | OpaqueRunChild;

export interface PageBreakLeaf {
  readonly kind: "page-break";
  readonly id: NodeId;
}

/**
 * `<w:lastRenderedPageBreak/>` — Word writes this hint at the position
 * where pagination broke during the last save. Layout-only; carries no
 * formatting. We use it as a cheap hint for the page chunker but never
 * treat it as authoritative (Word may not have written it; the geometry
 * may have changed since).
 */
export interface LastRenderedPageBreakLeaf {
  readonly kind: "last-rendered-page-break";
  readonly id: NodeId;
}
```

Today these collapse into `BreakLeaf { breakType: "page" }` and
`OpaqueRunChild` respectively. Promoting to discriminated variants keeps
the renderer pure and lets the chunker switch on `kind` directly.

Backwards compatibility: `BreakLeaf.breakType === "page"` is migrated to
the new typed `PageBreakLeaf` at parse time. The legacy `BreakLeaf` shape
stays for `column` and `textWrapping` (rare, layout-only).

## Parser

`packages/docx/src/parser/sections.ts` (new):

- `parseSectionProperties(entry: Record<string, unknown>): SectionProperties`
  walking `<w:pgSz>`, `<w:pgMar>`, `<w:cols>`, `<w:titlePg>`, `<w:type>`,
  `<w:headerReference>`, `<w:footerReference>`. Anything else captured to
  `opaqueProps`.
- Wire into `parseSectionBreak` (existing function) to attach
  `properties` alongside the existing `raw`.

`packages/docx/src/parser/parse.ts` `pushRunChild`: detect
`<w:br w:type="page"/>` → `PageBreakLeaf`. Detect
`<w:lastRenderedPageBreak/>` → `LastRenderedPageBreakLeaf`.

## Serializer

`packages/docx/src/serializer/sections.ts` (new):

- When `SectionBreak.raw` is present and `properties` has not been
  mutated (new helper `sectionPropertiesAreDerivedFromRaw(props, raw)`),
  emit `raw` verbatim.
- Otherwise rebuild `<w:sectPr>` from typed fields, appending
  `opaqueProps` in original order.

For the run leaves: emit `<w:br w:type="page"/>` and
`<w:lastRenderedPageBreak/>` from typed leaves; fall back to legacy
`BreakLeaf`/`OpaqueRunChild` paths unchanged.

## Header/footer reference graph (W8)

`packages/docx/src/agent/header-footer-graph.ts`:

```ts
export interface ResolvedHeaderFooter {
  readonly default?: HeaderFooterPart;
  readonly first?: HeaderFooterPart;
  readonly even?: HeaderFooterPart;
}

export function resolveHeaderFooterParts(
  snapshot: DocxSnapshot,
  sectionIndex: number
): { headers: ResolvedHeaderFooter; footers: ResolvedHeaderFooter };
```

Walks `body[sectionIndex]` (a `SectionBreak`), follows
`headerRefs[].relationshipId` through
`relationships.get("word/document.xml")` to a target like
`"header1.xml"`, then looks up the corresponding `HeaderFooterPart` in
`document.headersAndFooters`. Drives the renderer's choice of which
header/footer body to display in each page slot.

When `titlePg === true`, the first page of the section uses the `first`
header/footer (falling back to `default` when `first` isn't defined).
When `<w:settings><w:evenAndOddHeaders/>` is set (model addition deferred
to P4), even-numbered pages use the `even` variant.

## Round-trip tests (W7)

In `tests/roundtrip/docx/`:

- `section-properties.test.ts` — for each real-world fixture that
  contains a `<w:sectPr>`, parse → serialize → assert SHA-256
  byte-equality on `word/document.xml`.
- `page-break.test.ts` — synthetic fixture with two paragraphs split by
  `<w:br w:type="page"/>`. Parse → serialize → byte-identical.
- `last-rendered-page-break.test.ts` — synthetic fixture with the hint
  embedded mid-paragraph; parse → serialize → byte-identical.

The existing `real-world-roundtrip.test.ts` SHA-256 sweep across all
fixtures already gates this; the dedicated tests pin the specific
mechanism so a regression has a one-test failure pointing at the cause.

## Acceptance

- All 7 real-world fixtures continue to pass `make verify`'s round-trip
  byte-equality.
- The masterthesis fixture (which has multiple section breaks, a title
  page, and per-section page numbering) parses with non-empty
  `properties` on every section break.
- `resolveHeaderFooterParts` returns the correct part for every
  section in every fixture (smoke test).

## Out of scope (P3.2)

- Mutation commands for sections (`docx:set-page-margins`, etc.) — land
  in P3.4 / P3.6.
- `<w:settings>` parsing (evenAndOddHeaders, etc.) — needed for full
  even-page header support in P4.
- Footnote references — orthogonal, P4 (R8).
- Continuous section break flow vs nextPage — the chunker treats both
  identically for page splitting in P3.3, then P3.4 adds the distinction.
