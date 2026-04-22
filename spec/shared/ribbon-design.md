# Shared — Office-style ribbon design

> Status: F1. Replaces the single overloaded toolbar rows in
> DOCX, XLSX, PPTX with a tabbed ribbon shell + contextual tabs.
> Adopts PowerPoint / Word / Excel UX conventions: persistent
> tabs (Home, Insert, …), contextual tabs (Picture Format, Shape
> Format, Pivot Analyze, …) that appear and recede with
> selection.

## Why

The PPTX toolbar today is one ~40px row that scrolls
horizontally; on narrow viewports the user thrashes between
slide-level operations, character formatting, paragraph
formatting, and shape-fill controls — all in the same strip.
DOCX and XLSX are calmer but still dense.

Office's solution — and the user's explicit reference — is a
**tabbed ribbon**: only one task domain is primary at a time,
contextual tabs appear when their selection holds.

## Tab inventory

### Persistent tabs

| Editor | Tabs                                                                               |
| ------ | ---------------------------------------------------------------------------------- |
| PPTX   | **Home** · Insert · Design · Transitions · Animations · Slide Show · Review · View |
| DOCX   | **Home** · Insert · Layout · References · Review · View                            |
| XLSX   | **Home** · Insert · Layout · Formulas · Data · Review · View                       |

`**Home**` is the default active tab on first open.

### Contextual tabs

Visible only while the matching condition holds:

| Editor | Tab                          | Condition                                                              |
| ------ | ---------------------------- | ---------------------------------------------------------------------- |
| PPTX   | Picture Format               | A picture (or media shape) is selected                                 |
| PPTX   | Shape Format                 | Any non-picture shape (rect, ellipse, connector, group, …) is selected |
| PPTX   | Table Design                 | A table shape is selected (any cell selection within)                  |
| PPTX   | Chart Design                 | A chart shape is selected                                              |
| PPTX   | Media Playback               | A `MediaShape` is selected                                             |
| PPTX   | Master                       | The view is in master/layout edit mode                                 |
| DOCX   | Picture Format               | An image is selected                                                   |
| DOCX   | Table Design                 | The caret is inside a table                                            |
| DOCX   | Header & Footer              | The active part is a header/footer (Word-mode)                         |
| XLSX   | Pivot Analyze + Pivot Design | A cell inside a pivot table is selected                                |
| XLSX   | Picture Format               | An image is selected                                                   |
| XLSX   | Table Design                 | A structured table is selected                                         |
| XLSX   | Chart Design                 | A chart is selected                                                    |

Contextual tabs render with a coloured accent strip above the
tab label (PowerPoint convention: purple for Picture, orange for
Shape, green for Pivot/Table). Multiple contextual tabs may be
visible simultaneously (e.g. Picture + Header & Footer).

## Group convention

Every ribbon tab is composed of `RibbonGroup`s. A group is the
existing
[`packages/ui/src/primitives/toolbar-group.tsx`](../../packages/ui/src/primitives/toolbar-group.tsx)
primitive: a row of controls with an optional uppercase 9px
caption underneath. The caption is the user-readable group name
("Font", "Paragraph", "Slides", …).

```
┌──────────────────────────────┬─────────────────────────────┐
│ [B] [I] [U] [S] [Aa] ...     │ [≡] [≡] [≡] [≡]             │
│ Font                         │ Paragraph                   │
└──────────────────────────────┴─────────────────────────────┘
```

A group MUST have a caption unless its contents are
self-describing (e.g. a sole `Format Painter` button).

## Shell

```ts
// packages/ui/src/primitives/ribbon-shell.tsx

export interface RibbonShellProps {
  readonly tabs: ReadonlyArray<RibbonTabSpec>;
  readonly activeTabId: string;
  readonly onActiveTabChange: (id: string) => void;
  readonly trailing?: ReactNode; // Notes, Present, Save, etc.
  readonly height?: number; // default 96px
  readonly onCloseContextual?: (id: string) => void;
}

export interface RibbonTabSpec {
  readonly id: string;
  readonly label: string; // i18n'd by caller
  readonly accentColor?: string; // hex; if present the tab is contextual
  readonly groupId?: string; // group label above contextual tabs (e.g. "Picture Tools")
  readonly content: ReactNode; // the active-tab body
  readonly visible?: boolean; // contextual tabs default true; we hide via removal
}
```

Layout:

```
┌─────────────────────────────────────────────────────────────────┐
│ Tab strip:  Home  Insert  ...  | (Picture Tools) Picture Format │  ← 28px
├─────────────────────────────────────────────────────────────────┤
│ Active tab body (groups + captions)                             │  ← ~68px
└─────────────────────────────────────────────────────────────────┘
```

Total ~96px (vs today's 40px). On narrow viewports (<1100px
wide), groups collapse to icon-only with the caption replaced by
a small chevron — clicking the chevron pops a hover overlay
that shows the expanded group.

## Hooks

```ts
export function useContextualTab(id: string, visible: boolean, spec: Omit<RibbonTabSpec, "id">): void;
```

Editors register their contextual tabs once with the shell;
`visible` is the selection-driven condition. The shell
auto-shows / auto-hides them and switches the active tab to the
first newly-visible contextual tab if the user wasn't already on
a contextual.

## Migration path

The migration is **behaviour-preserving**: every control on
today's toolbars stays addressable by the same accessible name
in the new ribbon, so existing Playwright e2e selectors keep
working. Only the host UI moves.

### PPTX

Map of today's `PptxToolbar.tsx` controls → new tabs:

| Today                                             | Tab                           | Group                                   |
| ------------------------------------------------- | ----------------------------- | --------------------------------------- |
| Add slide / Duplicate / Delete slide              | Home                          | Slides                                  |
| Layout picker                                     | Home                          | Slides (split button)                   |
| Text box / Shape / Connector / Image / From XLSX  | Insert                        | Insert                                  |
| Replace (picture context)                         | Picture Format                | Adjust                                  |
| Delete (selection)                                | Home                          | Editing                                 |
| Six align icons + target toggle                   | Shape Format / Picture Format | Arrange                                 |
| Distribute H / V                                  | Shape Format / Picture Format | Arrange                                 |
| Arrange menu (z-order, group, ungroup)            | Shape Format                  | Arrange                                 |
| TextFormatBar (font/size/B/I/U/S/color/highlight) | Home                          | Font                                    |
| Align text L/C/R/J                                | Home                          | Paragraph                               |
| Vertical anchor T/M/B                             | Home                          | Paragraph                               |
| Fill picker                                       | Shape Format                  | Shape Styles                            |
| Comment                                           | Review                        | Comments                                |
| Connector controls (today floating)               | Shape Format                  | Lines (visible when connector selected) |
| Animations panel rail                             | Animations                    | (entire tab; panel becomes optional)    |
| Transitions controls (today in panel)             | Transitions                   | Transitions                             |

Notes & Present remain pinned in the trailing slot (top-right of
the shell, outside the tab strip).

### DOCX

Map of today's `Toolbar.tsx`:

| Today                                       | Tab                          | Group               |
| ------------------------------------------- | ---------------------------- | ------------------- |
| Style picker                                | Home                         | Styles              |
| TextFormatBar                               | Home                         | Font                |
| Align / indent / spacing / lists            | Home                         | Paragraph           |
| Formatting marks toggle                     | View                         | Show                |
| Insert image / table / xlsx / section break | Insert                       | Insert              |
| Comment / Review menu                       | Review                       | Comments / Tracking |
| Edit-mode picker                            | Review                       | Mode                |
| Page-number / Header & Footer commands      | Header & Footer (contextual) | Insert              |

`TableContextToolbar` and `ImageContextToolbar` collapse into
Table Design / Picture Format contextual tabs but **also** stay
as a near-cursor mini-toolbar for the most-used 4 controls
(Word's selection mini-toolbar pattern).

### XLSX

| Today                       | Tab                         | Group               |
| --------------------------- | --------------------------- | ------------------- |
| Undo / Redo                 | Home                        | Clipboard / History |
| Format painter              | Home                        | Clipboard           |
| TextFormatBar               | Home                        | Font                |
| Cell align L/C/R            | Home                        | Alignment           |
| Borders                     | Home                        | Font                |
| Merge / Unmerge             | Home                        | Alignment           |
| Text-to-columns             | Data                        | Data Tools          |
| Filter                      | Data                        | Sort & Filter       |
| Comment                     | Review                      | Comments            |
| Image                       | Insert                      | Illustrations       |
| Freeze menu                 | View                        | Window              |
| Number format `<select>`    | Home                        | Number              |
| Pivot creation entry point  | Insert                      | Tables              |
| Pivot Analyze + Design tabs | Pivot Analyze, Pivot Design | (contextual)        |

## i18n

All ribbon strings flow through `useTranslator()`. New i18n
namespace `ribbon.*`:

```
ribbon.tabs.home
ribbon.tabs.insert
ribbon.tabs.design
ribbon.groups.font
ribbon.groups.paragraph
ribbon.groups.slides
ribbon.actions.bold
ribbon.actions.bold.title       (tooltip / aria-label)
ribbon.actions.bold.shortcut    ("Cmd+B")
...
```

Catalogues: `apps/web/app/lib/i18n/en.json` + `de.json`. ~300
strings; we ship en + de in this session, other locales follow.

## Active-context indicator

When a contextual tab group is visible, the shell renders an
accent strip across the top of the tabs:

```
            ┌──────── Picture Tools (purple) ────────┐
Home Insert Design ... [Picture Format]  Animations
```

The strip uses the `accentColor` from the contextual tab spec.
PowerPoint's standard map:

| Tab group             | Color                   |
| --------------------- | ----------------------- |
| Picture Tools         | `#a020f0` (purple)      |
| Drawing Tools         | `#ff8c00` (orange)      |
| Table Tools           | `#36b37e` (green)       |
| Chart Tools           | `#1e88e5` (blue)        |
| PivotTable Tools      | `#1e88e5` (blue)        |
| Header & Footer Tools | `#5c2d91` (deep purple) |
| Master View           | `#d32f2f` (red)         |

Strip + label use the accent color; tab body styling is unchanged.

## Selection mini-toolbar

Word's "selection mini-toolbar" (the tiny floating bar that
appears next to a selection) is a separate primitive:

```ts
// packages/ui/src/primitives/selection-mini-toolbar.tsx
```

DOCX uses it for text selections (B/I/U/font/size/color/highlight).
PPTX uses it for shape selections in connector mode (the four
most-used connector controls). XLSX uses it for cell-range
operations (cut/copy/paste/delete contents).

## Action plumbing

Every ribbon button dispatches via the **same handlers** as
today's toolbar buttons. There is **no** new command surface
introduced by the ribbon migration — only the host UI changes.
The contract guarantees:

- `data-testid` and `aria-label` strings are preserved (with
  i18n updates only).
- Keyboard shortcuts stay registered globally (independent of
  ribbon visibility).
- Right-click context menus continue to work.

## Acceptance criteria

A1. **PPTX ribbon renders.** Open the PPTX editor; the tabbed
ribbon shows Home/Insert/Design/Transitions/Animations/Slide
Show/Review/View. Active tab defaults to Home.

A2. **Picture Format appears.** Insert / select an image; the
Picture Format contextual tab appears with a purple accent
strip; switching to Insert and back to Home doesn't make the
contextual disappear (it persists while selection holds).

A3. **DE locale.** Switch locale to DE; tab labels and group
captions render in German.

A4. **Behaviour preservation.** Every existing Playwright
toolbar e2e passes after the migration with at most label
updates.

A5. **Compact mode.** Resize the viewport to 900px wide; groups
collapse to icon-only; clicking a chevron expands a group
overlay.

## Out of scope (F1)

- Drag-to-rearrange tabs.
- User-defined custom tabs (Office's "Customize Ribbon"
  dialog).
- Quick Access Toolbar above the tabs (different convention).
- Touch-optimised "Touch / Mouse" mode toggle.
- Backstage view (File menu).
