# PPTX — Feature Scope

> Per-feature confidence ratings. Read this first.

| P-code  | Meaning                                                                                     |
| ------- | ------------------------------------------------------------------------------------------- |
| **P0**  | Lands this session. Tested. Documented.                                                     |
| **P1**  | Lands as **opaque-blob preservation only** this session. Roundtrip-clean. Editing deferred. |
| **P2**  | Out of scope this session. Not parsed/serialized beyond opaque blob.                        |
| **OUT** | Explicitly excluded by [`prompt.md`](../../prompt.md).                                      |

## In scope (per [`prompt.md`](../../prompt.md) lines 277–289)

| Feature                                                                                  | Open | Edit | Roundtrip |          This session           |
| ---------------------------------------------------------------------------------------- | :--: | :--: | :-------: | :-----------------------------: |
| Open any real-world `.pptx`                                                              |  ✓   |  —   |     ✓     |             **P0**              |
| Render slides faithfully (typed shapes + opaque fallback)                                |  ✓   |  —   |     ✓     |             **P0**              |
| Edit text in text boxes / shapes (content)                                               |  ✓   |  ✓   |     ✓     |             **P0**              |
| Edit run-level formatting in text frames (font / size / color / bold / italic / under-)  |  ✓   |  ✓   |     ✓     |             **P0**              |
| Reposition shapes / images / text boxes (drag + property-set)                            |  ✓   |  ✓   |     ✓     |             **P0**              |
| Resize shapes / images                                                                   |  ✓   |  ✓   |     ✓     |             **P0**              |
| Add new slides (blank + clone-from-layout)                                               |  ✓   |  ✓   |     ✓     |             **P0**              |
| Duplicate existing slides                                                                |  ✓   |  ✓   |     ✓     |             **P0**              |
| Delete slides (and any attached notes slide)                                             |  ✓   |  ✓   |     ✓     |             **P0**              |
| Reorder slides                                                                           |  ✓   |  ✓   |     ✓     |             **P0**              |
| Insert images (PNG / JPEG / GIF, with SHA-256 media dedup)                               |  ✓   |  ✓   |     ✓     |             **P0**              |
| Add new text boxes                                                                       |  ✓   |  ✓   |     ✓     |             **P0**              |
| Theme (color scheme / font scheme) — **read for resolution**, preserve verbatim          |  ✓   |  —   |     ✓     |             **P0**              |
| Slide masters & layouts — read for placeholder resolution; preserved verbatim            |  ✓   |  —   |     ✓     |             **P0**              |
| Notes slides — preserved verbatim per slide; cleaned up on `delete-slide`                |  ✓   |  —   |     ✓     |             **P0**              |
| SmartArt graphics                                                                        |  ✓   |  —   |     ✓     | **P0** (preserve as `Opaque*`)  |
| Animations & transitions                                                                 |  ✓   |  —   |     ✓     | **P0** (preserve as raw blocks) |
| Embedded charts (`graphicFrame` referencing `ppt/charts/`)                               |  ✓   |  —   |     ✓     |     **P0** (preserve only)      |
| In-slide tables (`graphicFrame` referencing `a:tbl`)                                     |  ✓   |  ✓   |     ✓     | **F2** (typed model + 5 commands) |
| Embedded videos / audio                                                                  |  ✓   |  —   |     ✓     |     **P0** (preserve only)      |
| Connectors (`p:cxnSp`) — render bounding box only; preserve geometry verbatim            |  ✓   |  —   |     ✓     |     **P0** (preserve only)      |
| Group shapes (`p:grpSp`) — group bounding box typed; children preserved verbatim         |  ✓   |  —   |     ✓     |          **P0** (move)          |
| `[Content_Types].xml`, `_rels/.rels`, `docProps/*`, `viewProps`, `presProps`, etc.       |  ✓   |  —   |     ✓     |    **P0** (preserve verbatim)   |

## Explicitly out

Per [`prompt.md`](../../prompt.md) lines 291–298:

| Feature                                                | Status                                  |
| ------------------------------------------------------ | --------------------------------------- |
| SmartArt **editing** (render + preserve only)          | **OUT** (preserved as `OpaqueShape`)    |
| Animation / transition **editing** (preserve only)     | **OUT** (preserved verbatim per slide)  |
| Chart editing within slides (preserve only)            | **OUT** (chart parts preserved opaque)  |
| Master / layout **editing** (preserve only)            | **OUT** (master & layout parts opaque)  |
| Notes-pages **editing** (preserve only)                | **OUT** (notes parts preserved opaque)  |
| ~~Table **editing** within slides (render + preserve)~~ | **F2** — typed `TableShape`; cell-text + add/delete row/column edits. Visual styling stays opaque (re-emit `<a:tblPr>` verbatim). |

## Roundtrip-integrity bar

All features above (P0, P1, OUT) must satisfy:

- Open a real-world `.pptx`, save without edit, file is byte-identical to input.
- Open, perform a P0 edit, save, reopen — only the edited region is changed; every other part is byte-identical to the input.

This is the **only** non-negotiable acceptance criterion.

## Tracking

`docs/build-log/pptx.md` keeps a row per feature with status, test reference, and any deviation from this spec.
