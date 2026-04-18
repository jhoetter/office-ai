# DOCX style cascade

> Status: P3.1 spec. Drives W1 (parser), W2 (resolver), W3 (toolbar binding).
> Source-of-truth for how a run's effective formatting is computed in the
> editor and exposed to the agent.

## Why

Today [`activeMarkAttr`](../../apps/web/app/lib/format-helpers.ts) only inspects
ProseMirror marks. When a paragraph carries `styleId=Heading1` and the run
inside it has empty `<w:rPr>`, the toolbar has no signal that the effective
font is Calibri 16pt — so the dropdowns show "Size" / "Font" placeholders
instead of `16` / `Calibri`. Word resolves through a four-level cascade
defined in `word/styles.xml`; we don't parse that file at all.

## Cascade order (low → high precedence)

```
docDefaults.rPrDefault           ← word/styles.xml/<w:docDefaults>
  └─ paragraph style chain        ← styleId → basedOn → ... (root)
       └─ paragraph's own pPr.rPr  ← <w:pPr><w:rPr>...</w:rPr></w:pPr>
            └─ run's rPr            ← <w:r><w:rPr>...</w:rPr></w:r>
```

Every later layer overrides individual fields of the earlier layers. A run
with `<w:rPr><w:b/></w:rPr>` makes the run bold but inherits its font and
size from above. A run with empty `<w:rPr/>` inherits everything.

For paragraph properties (`pPr`) the equivalent chain is:

```
docDefaults.pPrDefault
  └─ paragraph style chain
       └─ paragraph's own pPr
```

(Numbering / list paragraphs add a fifth layer — `pPr` inherited from the
list level — which we'll handle but flag as out-of-scope for the toolbar
display in P3.1; see "Out of scope" below.)

## Typed model

New file: `packages/docx/src/model/styles.ts`. Imported and re-exported
through `model/types.ts`.

```ts
export interface StylesPart {
  readonly docDefaults: {
    readonly rPrDefault?: RunProperties;
    readonly pPrDefault?: ParagraphProperties;
  };
  /** Keyed by styleId. Insertion order matches load order. */
  readonly styles: ReadonlyMap<string, StyleDefinition>;
  /**
   * Captured but not modeled: <w:latentStyles>, <w:lsdException>, etc.
   * Re-emitted verbatim. Present on freshly-parsed StylesPart, dropped if
   * any field of the typed cascade is mutated.
   */
  readonly raw?: OpaqueXml;
}

export interface StyleDefinition {
  readonly id: string;
  readonly type: "paragraph" | "character" | "table" | "numbering";
  readonly name?: string;
  readonly basedOn?: string;
  readonly next?: string;
  readonly link?: string;
  readonly hidden?: boolean;
  readonly default?: boolean;
  readonly rPr?: RunProperties;
  readonly pPr?: ParagraphProperties;
  /** Anything we don't model on this <w:style>. */
  readonly opaqueProps?: ReadonlyArray<OpaqueXml>;
}
```

Threaded through `DocxDocument` as a new optional field:

```ts
export interface DocxDocument {
  // ... existing fields
  readonly styles?: StylesPart;
}
```

`undefined` when `word/styles.xml` is absent — synthetic test fixtures
without a styles part stay valid.

## Parser

`packages/docx/src/parser/styles.ts`

- Entry point: `parseStylesPart(bytes: Uint8Array): StylesPart`.
- Walks `<w:styles>` children:
  - `<w:docDefaults>` → `docDefaults.{rPrDefault, pPrDefault}` reusing the
    existing `parseRunProperties` / `parseParagraphProperties` helpers.
  - `<w:style w:type w:styleId>` → `StyleDefinition`. `<w:basedOn w:val>`
    populates `basedOn`.
- Anything else (latent styles, doc parts, etc.) is captured into
  `StylesPart.raw`. The parser does NOT walk into those subtrees.
- Wired into `parseDocxBuffer` after `headersAndFooters`.

`docDefaults.rPrDefault` ends up holding things like `Calibri 11pt` for a
modern Word doc. `Heading1` typically has `basedOn="Normal"` and overrides
`fontSize=32` (`<w:sz w:val="32"/>` = 16pt because OOXML stores half-points).

## Resolver

`packages/docx/src/agent/style-resolver.ts`

```ts
export interface ResolvedRunProperties {
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: boolean | string;
  readonly strike: boolean;
  readonly fontFamily: string | undefined;
  /** Half-points (OOXML w:sz units). 22 = 11pt, 32 = 16pt. */
  readonly fontSize: number | undefined;
  readonly color: string | undefined;
  readonly highlight: string | undefined;
}

export interface ResolvedParagraphProperties {
  readonly styleId: string | undefined;
  readonly alignment: "left" | "center" | "right" | "justify" | undefined;
  readonly indentation: NormalizedIndentation;
  readonly spacing: NormalizedSpacing;
  readonly numbering: { numId: number; ilvl: number } | undefined;
}

export function resolveEffectiveRpr(
  snapshot: DocxSnapshot,
  paragraphIndex: number,
  runIndex?: number
): ResolvedRunProperties;

export function resolveEffectivePpr(
  snapshot: DocxSnapshot,
  paragraphIndex: number
): ResolvedParagraphProperties;
```

### Algorithm (rPr)

1. Start with `docDefaults.rPrDefault ?? {}`.
2. Walk the paragraph's `styleId` chain. For each style from root → leaf
   (i.e. resolve `basedOn` recursively first, then layer leaves on top),
   merge `style.rPr` into the accumulator (each defined field overrides).
3. Merge the paragraph's `properties.opaqueProps`-derived `pPr.rPr` if
   present (Word stores per-paragraph default run formatting here).
4. If `runIndex` is provided, merge the run's own `RunProperties`.
5. Return the accumulator with `undefined` collapsed for fields that were
   never set.

### basedOn cycle handling

Real-world docs have been seen with cycles (rare but legal-ish).
The resolver hard-caps walks at depth 16 and breaks on already-visited
ids. Logged via the existing parser warning sink, not thrown.

### Caching

`resolveEffectiveRpr` is hot — called per selection-change in the
toolbar. Resolver memoizes per (snapshotId, styleId) inside the
StylesPart resolution; the outer call is then O(1) for the style chain
plus O(1) for the leaf merges. The cache is cleared whenever
`StylesPart` itself is mutated (rare; not in P3 scope).

## Toolbar binding (W3)

`apps/web/app/lib/format-helpers.ts`:

- New `activeRunAttr<T>(state, snapshot, attrName: keyof ResolvedRunProperties): MaybeMixed<T>`.
- For each text node in selection (or stored marks at caret), compute the
  effective value via `resolveEffectiveRpr` and merge with the run's PM
  marks (PM marks always win — they represent the typed `<w:rPr>` on the
  run). Mixed when values differ, undefined when no path resolves.
- `FontSizePicker` and `FontFamilyPicker` consume the new helper.
- The placeholder `"Size"` / `"Font"` is reserved for the explicit "no
  style cascade resolved a value" case (which should be impossible after
  W1 lands for any well-formed Word doc).

## Round-trip

- Parser keeps `StylesPart.raw` of the original bytes.
- Serializer emits original bytes when `dirty.styles === false` (new
  dirty flag). Mutating commands (none in P3) flip the flag.
- For the no-style fixture set (synthetic tests without
  `word/styles.xml`), `StylesPart` stays `undefined` and serialization
  is a no-op for that part.
- Acceptance: every real-world fixture in `fixtures/docx/real-world/`
  re-serializes byte-identical with the new typed parser in place.

## Out of scope (P3.1)

- Numbering-level inheritance (the fifth cascade layer). Documented in
  the resolver as a known gap; the toolbar reads `pPr.numbering` directly
  from the paragraph for now.
- Theme color resolution (`<w:themeColor>`). Treated as opaque; the
  toolbar still resolves explicit `<w:color w:val>` values.
- Conditional table styles (`<w:tblStylePr w:type="firstRow">`).
- Style mutation commands (`docx:set-style-rpr`, etc.). Adding to the
  P4 backlog under R10.
