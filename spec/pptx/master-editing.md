# PPTX — Master / Layout / Theme editing

> Status: F1. Promotes `ppt/slideMasters/*`, `ppt/slideLayouts/*`,
> and `ppt/theme/themeN.xml` from opaque preservation to typed,
> editable parts. Adds an in-app "Master view" with theme color &
> font editing. Supersedes the "Master / layout editing — preserve
> only" line in [`feature-scope.md`](feature-scope.md).

## Why

The user wants a "master design" they can define once so every
new slide automatically inherits it (PowerPoint's Slide Master
view). Today our slide masters and layouts are byte-clean opaque
blobs — readable for placeholder cloning during `add-slide`, but
not editable through any command, and definitely not editable
visually.

## OOXML mapping

| Part                                           | Today        | After F1                                                                        |
| ---------------------------------------------- | ------------ | ------------------------------------------------------------------------------- |
| `ppt/slideMasters/slideMasterN.xml`            | `OpaquePart` | typed `SlideMaster { background, placeholders, shapes, colorMapOverride, raw }` |
| `ppt/slideLayouts/slideLayoutN.xml`            | `OpaquePart` | typed `SlideLayout { master, type, name, placeholders, shapes, raw }`           |
| `ppt/theme/themeN.xml`                         | `OpaquePart` | typed `Theme { colorScheme, fontScheme, formatScheme, raw }`                    |
| `ppt/slideMasters/_rels/slideMasterN.xml.rels` | preserved    | preserved (links to layouts + theme + media)                                    |
| `ppt/slideLayouts/_rels/slideLayoutN.xml.rels` | preserved    | preserved (links to its master + media)                                         |

The relationship graph (master → layouts, master → theme, slide
→ layout) stays the source of truth for navigation; we don't
flatten it.

## Typed model

```ts
// packages/pptx/src/model/master.ts

export interface ThemeColor {
  /** Either an explicit hex (`#RRGGBB`) or a theme-relative ref. */
  readonly hex: string;
}

export interface ThemeColorScheme {
  readonly name: string;
  /** 12 named slots: dk1, lt1, dk2, lt2, accent1..accent6, hyperlink, followedHyperlink. */
  readonly dk1: ThemeColor;
  readonly lt1: ThemeColor;
  readonly dk2: ThemeColor;
  readonly lt2: ThemeColor;
  readonly accent1: ThemeColor;
  readonly accent2: ThemeColor;
  readonly accent3: ThemeColor;
  readonly accent4: ThemeColor;
  readonly accent5: ThemeColor;
  readonly accent6: ThemeColor;
  readonly hyperlink: ThemeColor;
  readonly followedHyperlink: ThemeColor;
}

export interface ThemeFontScheme {
  readonly name: string;
  readonly major: { readonly latin: string; readonly ea?: string; readonly cs?: string };
  readonly minor: { readonly latin: string; readonly ea?: string; readonly cs?: string };
}

export interface Theme {
  readonly partPath: string;
  readonly name: string;
  readonly colorScheme: ThemeColorScheme;
  readonly fontScheme: ThemeFontScheme;
  /** `formatScheme` (fillStyleLst, lnStyleLst, effectStyleLst, bgFillStyleLst) — preserved verbatim for F1. */
  readonly formatSchemeRaw: string;
  readonly raw?: Readonly<Record<string, unknown>>;
}

export interface SlideMaster {
  readonly partPath: string;
  /** Background fill (typed for solid/gradient/picture; raw for the rest). */
  readonly background: SlideBackground;
  readonly placeholders: ReadonlyArray<TextShape>;
  readonly shapes: ReadonlyArray<Shape>;
  readonly colorMapOverride?: Readonly<Record<string, string>>;
  /** Linked theme part path. */
  readonly themePartPath: string;
  /** Linked layout part paths in order. */
  readonly layoutPartPaths: ReadonlyArray<string>;
  readonly raw?: Readonly<Record<string, unknown>>;
}

export interface SlideLayout {
  readonly partPath: string;
  readonly masterPartPath: string;
  /** OOXML layout type attribute (`title`, `obj`, `twoObj`, `tx`, `blank`, …). */
  readonly type: string;
  readonly name: string;
  readonly background?: SlideBackground;
  readonly placeholders: ReadonlyArray<TextShape>;
  readonly shapes: ReadonlyArray<Shape>;
  readonly raw?: Readonly<Record<string, unknown>>;
}

export type SlideBackground =
  | { kind: "solid"; color: string /* hex */ }
  | { kind: "gradient"; raw: string }
  | { kind: "picture"; mediaPath: string }
  | { kind: "themed"; idx: number /* index into theme.bgFillStyleLst */ }
  | { kind: "raw"; raw: string };
```

`PptxPresentation` gains:

```ts
readonly masters: ReadonlyArray<SlideMaster>;
readonly layouts: ReadonlyArray<SlideLayout>;
readonly themes: ReadonlyArray<Theme>;
/** Optional default layout for new slides, by part path. */
readonly defaultLayoutPartPath?: string;
```

The existing `OpaquePart` arrays are retained for any part we
fail to type (defensive — if parsing throws on an exotic master,
fall back to opaque).

## Parser

`packages/pptx/src/parser/master.ts` reads each master:
background, placeholder tree, plain shapes (reuses the slide
shape parser — same OOXML grammar). Walks `colorMapOverride` and
`hf` children. Stores rest as `raw`.

`packages/pptx/src/parser/layout.ts` similarly; reads `type`
from the root element, `name` from `cSld/@name`.

`packages/pptx/src/parser/theme.ts` already reads color scheme
for rendering; extend to also extract font scheme + persist as
typed `Theme`.

Promotion: remove from `OpaquePart` paths in `parse.ts`; add
typed loaders in the read sequence after slides.

## Serializer

Mirror writers under `packages/pptx/src/serializer/master.ts`,
`layout.ts`. The same dirty-tracking pattern as slides: only
re-emit parts whose typed model differs from parse.

## Master view UI

Toggled from the View tab in the new ribbon (or a temporary
"Edit Master" button until the ribbon ships — see
[`spec/shared/ribbon-design.md`](../shared/ribbon-design.md)).

```
┌─────────┬──────────────────────────────────────────┐
│ MASTER  │                                          │
│  ┌────┐ │     [active master/layout canvas]        │
│  │ M1 │ │                                          │
│  └────┘ │     uses existing SlideCanvas            │
│ LAYOUTS │     in master/layout edit mode           │
│  ┌────┐ │                                          │
│  │ L1 │ │                                          │
│  └────┘ │                                          │
│  ┌────┐ │                                          │
│  │ L2 │ │                                          │
│  └────┘ │                                          │
└─────────┴──────────────────────────────────────────┘
```

`apps/web/app/pptx-editor/MasterView.tsx`:

- Left rail: master at top, then numbered layouts beneath. Click
  to switch the canvas target.
- Canvas: reuses `SlideCanvas` with a `masterMode: { kind:
"master" | "layout"; partPath }` prop. Selection / shape moves
  / text edits dispatch the same shape commands but addressed
  against the master/layout part instead of a slide.
- Top of the canvas: a **Theme** sub-toolbar with two pickers:
  _Colors_ and _Fonts_, opening the editors below.

### Theme editor

`apps/web/app/pptx-editor/ThemeEditor.tsx`:

- 12 color swatches (the slots in `ThemeColorScheme`) with a
  color picker each.
- Two font pickers (Major / Minor) with Latin font selection.
- "Apply" dispatches `pptx:set-theme-colors` / `pptx:set-theme-fonts`.
- Live preview: every open slide re-renders because the existing
  `themeDefault` resolution path already consults the active
  theme.

## Commands

| Command                          | Payload                                                                                          | Effect                                                                    |
| -------------------------------- | ------------------------------------------------------------------------------------------------ | ------------------------------------------------------------------------- |
| `pptx:set-theme-colors`          | `{ themePartPath; colors: Partial<ThemeColorScheme> }`                                           | Mutates theme color scheme (single or multiple slots)                     |
| `pptx:set-theme-fonts`           | `{ themePartPath; major?; minor? }`                                                              | Mutates font scheme                                                       |
| `pptx:set-master-background`     | `{ masterPartPath; background: SlideBackground }`                                                | Replaces master background                                                |
| `pptx:set-layout-name`           | `{ layoutPartPath; name }`                                                                       | Renames a layout                                                          |
| `pptx:set-default-layout`        | `{ layoutPartPath \| null }`                                                                     | Records `defaultLayoutPartPath` on presentation; new slides clone from it |
| `pptx:add-shape-to-master`       | `{ masterPartPath; shape }`                                                                      | Inserts a shape onto master (logos, footer text, etc.)                    |
| `pptx:add-shape-to-layout`       | `{ layoutPartPath; shape }`                                                                      | Same for a layout                                                         |
| (Existing shape / text commands) | Augmented to accept `{ target: { kind: "master" \| "layout"; partPath } }` instead of a slide id | Master/layout placeholders edit through the same handlers                 |

The "augment shape commands" pattern is critical: every shape
command (`set-shape-text`, `set-shape-position`, `set-shape-fill`,
`delete-shape`, …) gets a `target` field so the master view can
reuse them without duplication.

## New-slide inheritance

`packages/pptx/src/commands/add-slide.ts` already supports
`layoutPartPath`. We add:

- If `layoutPartPath` is omitted **and**
  `presentation.defaultLayoutPartPath` is set, use that.
- Otherwise: existing behaviour (resolve by `layoutKind` or
  pick first usable layout).

This makes "Set as default layout" a one-toggle setting from the
master view.

## Round-trip invariants

1. **No-edit.** Load + save without master/layout/theme edits
   leaves all three parts byte-identical.
2. **Single-theme edit.** Change one color → only the affected
   `themeN.xml` is re-emitted; layouts/masters byte-identical.
3. **PowerPoint compatibility.** A theme color edit saved via
   our serializer reopens in PowerPoint with the same value
   (validated via a LibreOffice headless roundtrip + a fixture
   diff in `tests/roundtrip/pptx/`).
4. **Layout add-slide hint.** `defaultLayoutPartPath` survives
   round-trip via a `<p:extLst>` extension on the
   `presentation.xml` part (custom URI
   `urn:officeai:default-layout`); when not set, no extension is
   emitted.

## Acceptance criteria

A1. **Load.** Real-world fixtures parse into non-empty
`masters`, `layouts`, `themes` collections.

A2. **Theme color edit.** `pptx:set-theme-colors` flips
`accent1` from blue to red; the editor canvas re-renders all
slides with the new accent; saved file opens with the new color
in PowerPoint.

A3. **Master shape add.** `pptx:add-shape-to-master` inserting a
text "Confidential" appears on every slide that uses a layout
derived from that master.

A4. **Default layout.** `pptx:set-default-layout` followed by
`pptx:add-slide` (no `layoutPartPath`) produces a slide cloned
from the chosen layout.

A5. **Round-trip discipline.** `make audit-roundtrip` stays at
100% across fixtures; new fixture
`fixtures/pptx/09-themed-master.pptx` exercises a non-default
master.

## Out of scope (F1)

- Master variant generation (Office's built-in palette of
  starter themes).
- Color theme palettes (saved gallery of color sets).
- Importing themes from other files.
- `formatScheme` typed model (fill/line/effect style lists) —
  preserved as `formatSchemeRaw` for now; F2 may type if needed.
- Per-layout backgrounds with rich gradients beyond `SlideBackground.kind = "gradient"` (raw preserve).
- Master-of-master inheritance (only one master level today —
  PowerPoint's "Slide Master" + its child layouts).
