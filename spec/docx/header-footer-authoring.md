# DOCX header/footer authoring (P3.4)

> Status: P3.4 spec. Drives W14 (focus model), W15 (header/footer
> schema integration), W16 (insert-page-number / set-different-first
> commands), W17 (insert-section-break command).
>
> Builds on P3.2 (typed `SectionProperties`, `HeaderFooterRef`,
> `resolveHeaderFooterParts`) and P3.3 (page chunker, page divider
> decorations). Followed by P3.5 (page-aware editing UX).

## Why

Today header / footer parts parse into typed `HeaderFooterPart`
objects (W2/W3 of P1) and their text can be rewritten via
`docx:set-header-text` / `docx:set-footer-text`. But the editor
surface treats them as invisible — the user can neither see them
on the page, focus into them, nor format them with the toolbar.
Word users expect to:

1. **See** the header/footer at the top/bottom of every visible page,
   styled like the body but greyed when the body has focus.
2. **Click in** the header/footer (or use Insert → Header) to
   activate it. Body content greys out, the toolbar retargets the
   active part, the active part shows a label like
   `Header — First Page`.
3. **Insert a page number field** (`<w:fldSimple w:instr=" PAGE "/>`)
   that updates per page, styled by the surrounding run.
4. **Toggle "Different first page"** — Word's most common section
   knob. Backed by `<w:titlePg/>`.
5. **Insert a section break** (next page, continuous, even, odd) at
   the caret, splitting the current section's geometry / header
   refs across the new boundary.

These are the four building blocks the rest of P3 / P4 layers more
advanced authoring on (different-odd-even, per-section margins,
restart numbering, etc.).

## Typed model additions

### `PageNumberField` run child

`<w:fldSimple w:instr=" PAGE "/>` and the equivalent
`<w:fldChar>`-bracketed sequence both render as "current page
number". Today they parse as `OpaqueRunChild`, which means the
serializer can preserve them but no command can produce them.
P3.4 promotes the simple-field form to a typed leaf:

```ts
export interface PageNumberFieldLeaf {
  readonly kind: "page-number-field";
  readonly id: NodeId;
  /**
   * Either "PAGE" (current page) or "NUMPAGES" (total pages).
   * The full Word field grammar is far larger; we type only the
   * two variants the toolbar can produce, leaving the rest to
   * `OpaqueRunChild` for byte-identical round-trip.
   */
  readonly field: "PAGE" | "NUMPAGES";
  /**
   * The literal field instruction string, captured verbatim so
   * the serializer can re-emit the exact `w:instr` (including
   * `\* MERGEFORMAT` switches Word's UI typically appends).
   * Round-trip invariant: `parse → serialize → parse` produces
   * the same `instr`.
   */
  readonly instr: string;
  /**
   * Optional cached display value (e.g. "3"). Word writes this as
   * a `<w:t>` child of the field for offline rendering. Recomputed
   * by the renderer; preserved as a hint for byte round-trip when
   * the field hasn't been touched.
   */
  readonly cachedText?: string;
}
```

Added to `RunChild`. The complex `<w:fldChar>`-bracketed form
stays as `OpaqueRunChild` in P3.4 — it requires multi-run
re-assembly that lands in P4.

### Section break payload

W17 needs a typed payload separate from the existing
`SectionProperties` model:

```ts
export interface InsertSectionBreakPayload {
  /** Insert before this paragraph index (0 = top of doc). */
  paragraphIndex: number;
  /**
   * Section type for the new boundary. Maps to `<w:type>` inside
   * the inserted `<w:sectPr>`.
   *
   * Default: "nextPage" (Word's "Next Page" section break).
   */
  type?: "nextPage" | "continuous" | "evenPage" | "oddPage";
}
```

## Commands (W16, W17)

### `docx:insert-page-number`

Inserts a `PageNumberFieldLeaf` at the caret inside the active
header / footer paragraph. Shape:

```ts
export interface InsertPageNumberPayload {
  /** Stable id of the target paragraph (must live inside a header or footer part). */
  paragraphId: NodeId;
  /** Byte offset inside the paragraph's flat-text. */
  offset: number;
  /** Defaults to "PAGE". */
  field?: "PAGE" | "NUMPAGES";
}
```

Behavior:
- Errors with `unknown-target` if the paragraph isn't inside a
  header/footer part.
- Errors with `invalid-payload` if `offset` is out of range.
- Splits the targeted run at `offset`, splices a fresh run
  containing only the `PageNumberFieldLeaf`, and merges back the
  trailing half. Run properties are inherited from the split run.
- Dirties the owning header/footer part path.

### `docx:set-section-different-first`

Toggles `titlePg` on the section that contains a given paragraph
(or the trailing implicit section if `paragraphIndex >=
body.length`).

```ts
export interface SetSectionDifferentFirstPayload {
  /**
   * Any paragraph index *inside* the target section. The handler
   * walks forward from this index to find the next `SectionBreak`
   * (or the trailing implicit section).
   */
  paragraphIndex: number;
  /** New value. Pass `false` to disable. */
  enabled: boolean;
}
```

Behavior:
- Locates the section by walking forward from `paragraphIndex` to
  the next `SectionBreak` block or the document's trailing
  implicit `sectPr`.
- Mutates the section's `properties.titlePg` to the new value.
- When enabling, **does not** auto-create a `first` header part
  — that's P4. The toolbar surfaces a follow-up affordance for
  the user to run `docx:set-header-text` against the (yet-to-be
  created) first-page part. P3.4 ships only the typed flag flip
  + serializer support for `<w:titlePg/>`.
- Dirties `body` (since `<w:sectPr>` lives in `word/document.xml`).

### `docx:insert-section-break`

Splits a section at `paragraphIndex` by inserting a new
`SectionBreak` block immediately before that paragraph. The new
block inherits the geometry of the parent section (so margins,
size, header/footer refs are preserved); the caller can subsequently
mutate page geometry via P4 commands.

```ts
export interface InsertSectionBreakPayload {
  paragraphIndex: number;
  type?: "nextPage" | "continuous" | "evenPage" | "oddPage";
}
```

Behavior:
- Errors `invalid-payload` if `paragraphIndex` is out of range
  (inserting at `body.length` is permitted — appends a section
  break before the trailing sectPr).
- Mints a fresh `SectionBreak` whose `properties` is a structural
  clone of the section's existing properties, with `type` set to
  the requested value (defaults to `"nextPage"`).
- Dirties `body`.

## Editor focus model (W14, W15)

The body PM instance stays the source of truth for the document
body. Headers / footers each get their own lightweight PM
instance, mounted lazily when the user activates one. State:

```ts
type ActivePart =
  | { kind: "body" }
  | { kind: "header-footer"; partId: string };
```

Stored in `DocxEditor` React state. Transitions:

| Trigger                                   | New `ActivePart`        |
| ----------------------------------------- | ----------------------- |
| Click inside body                         | `{ kind: "body" }`      |
| Click inside a rendered header / footer   | `{ kind: "header-footer", partId }` |
| Press `Esc` while in header/footer        | `{ kind: "body" }`      |
| Toolbar "Close Header/Footer" button      | `{ kind: "body" }`      |

While `kind === "body"`:

- Header / footer PM instances render with `editable: false` and a
  `data-active-part="false"` attribute that lowers their opacity
  (CSS-only; doesn't change layout).
- Toolbar acts on the body PM (current behavior, unchanged).

While `kind === "header-footer"`:

- Body PM gets `editable: false` + greyed-out attribute.
- Active part PM gets `editable: true` and shows a label
  `Header — Default` / `Footer — First Page` etc., derived from
  the part's relationship type.
- Toolbar dispatches commands against the active part. A small
  set of toolbar buttons becomes header-only (Insert Page Number,
  Different First Page).

W15 reuses the existing docx PM schema for header/footer
instances — they're paragraphs of runs, no schema differences.
The mount helper lives in
`packages/docx/src/renderer/mount-header-footer.ts` (new file)
and accepts:

```ts
interface MountHeaderFooterOptions {
  agent: DocxAgent;
  partId: string;
  source: "human" | "agent";
  onUnsupported?: (events: ReadonlyArray<UnsupportedTx>) => void;
  onError?: (err: unknown) => void;
}
```

It mirrors `mountDocxEditor` but scopes the schema to the
header/footer part: the PM doc is built from
`HeaderFooterPart.body` instead of `DocxDocument.body`, and the
transaction-to-commands translator emits
`docx:set-header-text` / `docx:set-footer-text` /
`docx:insert-page-number` instead of body commands.

For P3.4 the simplest viable mount routes only **typing** through
the existing `set-header-text` (whole-paragraph rewrite) command;
inline formatting from the toolbar (bold, page-number insert) is
the only other channel. Inline formatting via run mutation
inside a header is a P4 concern (it requires the same
`format-range` machinery the body has, scoped to header parts).

## Renderer integration (W11 follow-up)

`pageDecorationsPlugin` already inserts a "Page N of M" widget
between page chunks. P3.4 extends the widget to render the
**section's resolved header / footer** above / below the divider,
using `resolveHeaderFooterParts` (W8).

The header/footer block above the divider is a
`Decoration.widget(side: -1)` whose DOM is a non-editable preview
of the part's body, optionally overlaid by an active PM instance
when the user has focused that part. The active PM is mounted at
the top of the page geometry (one PM per visible page is
expensive — instead we mount **one** PM per part type and reposition
it via CSS when the user switches pages).

Header/footer DOM hierarchy:

```
<div class="pm-page-divider" data-page-number="2">
  <div class="pm-page-header" data-part-id="word/header1.xml">
    [non-editable preview OR active PM]
  </div>
  <span class="pm-page-divider__label">Page 2 of N</span>
</div>
```

The footer mirrors the header, attached to the *previous* page
chunk's bottom rather than the next chunk's top. P3.4 ships the
header preview only; footer rendering and active-part PM mounting
are wired up but the integration test cases focus on the
typed-model + commands. The full visual end-state (active PM
overlay, focus routing) lands as part of W14's UI work and is
verified manually against the masterthesis fixture in P3.7 / W25.

## Round-trip invariants

- Snapshots without P3.4 mutations round-trip byte-identically
  (existing `dirty.body === false`, `dirty.headersAndFooters` empty,
  no new `<w:sectPr>` produced).
- `PageNumberFieldLeaf` re-serializes to the *exact* `w:instr`
  string captured at parse time, including switches.
- A fresh-from-command page-number leaf serializes to
  `<w:fldSimple w:instr=" PAGE \* MERGEFORMAT "/>` (matching
  Word's own emit).
- `set-section-different-first` produces a `<w:titlePg/>`
  child of `<w:sectPr>` only when `enabled === true`. Disabling
  removes the element.
- `insert-section-break` writes the new `<w:sectPr>` inside a
  paragraph's `<w:pPr>` (the OOXML idiom for embedded section
  breaks), with `<w:type>` matching the payload `type`.

## Acceptance criteria

A1. **Page-number field round-trip.**
A header that already contains `<w:fldSimple w:instr=" PAGE \* MERGEFORMAT "/>` parses to a `PageNumberFieldLeaf`, re-serializes byte-identical, and re-parses to an equal leaf.

A2. **Insert page number.**
`docx:insert-page-number` inside an empty header paragraph produces a single run containing a `PageNumberFieldLeaf` with `field: "PAGE"`.

A3. **Different-first toggle.**
`docx:set-section-different-first({ paragraphIndex: 0, enabled: true })` flips `properties.titlePg` to `true` on the trailing implicit section. Disabling reverses it.

A4. **Insert section break.**
`docx:insert-section-break({ paragraphIndex: 1, type: "continuous" })` produces a new `SectionBreak` with `properties.type === "continuous"` and the geometry of the inheriting section.

A5. **Focus model UI.**
Clicking a header preview swaps `activePart` to that part, greys the body PM, and shows the part label. Pressing `Esc` returns focus to the body.

A6. **No regressions.**
All existing tests stay green (208+ docx, 51 integration). New tests cover the four commands and the page-number leaf round-trip.

## Out of scope (P3.4)

- Auto-creating header / footer parts when toggling
  `differentFirst` on a section that lacks them (P4).
- `<w:fldChar>`-bracketed multi-run field syntax (P4).
- Different odd / even header support (`<w:settings>`
  `evenAndOddHeaders`, P4).
- Per-section restart of page numbering (`<w:pgNumType>`).
- Inline run formatting inside header / footer parts via the
  toolbar — typing works, but bold / italic / font / size etc.
  inside a header are deferred to P4.
- Mounting a separate PM per visible page (we mount one per part
  per active state).
