# DOCX — Feature Scope

> Per-feature confidence ratings. Read this first.

| P-code | Meaning                                                              |
| ------ | -------------------------------------------------------------------- |
| **P0** | Lands this session. Tested. Documented.                              |
| **P1** | Lands as **opaque-blob preservation only** this session. Roundtrip-clean. Editing deferred. |
| **P2** | Out of scope this session. Not parsed/serialized beyond opaque blob. |
| **OUT**| Explicitly excluded by [`prompt.md`](../../prompt.md).               |

## In scope (per [`prompt.md`](../../prompt.md) lines 205–217)

| Feature                                  | Open | Edit | Roundtrip | This session |
| ---------------------------------------- | :--: | :--: | :-------: | :----------: |
| Plain text content                       | ✓    | ✓    | ✓         | **P0**       |
| Paragraph: alignment / indentation / spacing | ✓ | ✓ | ✓ | **P0** |
| Paragraph: heading styles (H1–H6) / style by name | ✓ | ✓ | ✓ | **P0** |
| Run formatting: bold, italic, underline, strikethrough | ✓ | ✓ | ✓ | **P0** |
| Run formatting: font family / size / color / highlight | ✓ | ✓ | ✓ | **P0** |
| Comments: add new                        | ✓    | ✓    | ✓         | **P0**       |
| Comments: read existing & preserve threads | ✓ | (read) | ✓ | **P0** |
| Comments: resolve / reply / delete       | ✓    | —    | ✓         | **P1**       |
| Tracked changes: parse + preserve revision markup | ✓ | — | ✓ | **P0** |
| Tracked changes: accept / reject         | ✓    | —    | ✓         | **P1**       |
| Hyperlinks                               | ✓    | (preserve) | ✓ | **P0** |
| Lists: bullet / numbered / nested (preserve) | ✓ | (style only) | ✓ | **P0** (preserve) / **P1** (mutate) |
| Tables: open + render                    | ✓    | —    | ✓         | **P0** (preserve) / **P1** (mutate) |
| Tables: insert / merge-split / borders / column widths | ✓ | — | ✓ | **P1** |
| Images: inline + floating                | ✓    | (preserve) | ✓ | **P0** (preserve) / **P1** (insert) |
| Images: insert / resize / reposition     | ✓    | —    | ✓         | **P1**       |
| Headers and footers                      | ✓    | (preserve) | ✓ | **P0** (preserve) / **P1** (edit) |
| Page breaks (preserve)                   | ✓    | (preserve) | ✓ | **P0**       |
| Section breaks (preserve, do not edit)   | ✓    | —    | ✓         | **P0**       |
| Numbering definitions (preserve)         | ✓    | —    | ✓         | **P0**       |
| Styles part (preserve)                   | ✓    | —    | ✓         | **P0**       |
| Custom XML, font tables, web settings, theme, etc. (preserve as opaque parts) | ✓ | — | ✓ | **P0** |

## Explicitly out

Per [`prompt.md`](../../prompt.md) lines 219–225:

| Feature                                            | Status |
| -------------------------------------------------- | ------ |
| Mail merge / field codes (preserve only)           | **OUT** (preserved as opaque blocks) |
| Complex cross-references and footnotes (preserve only) | **OUT** (preserved as opaque blocks) |
| Embedded OLE / linked Excel charts (preserve as blobs) | **OUT** (preserved as opaque parts) |
| VBA macros (preserve, do not execute)              | **OUT** (preserved as opaque parts) |
| Custom XML data binding                            | **OUT** (preserved as opaque blocks) |

## Roundtrip-integrity bar

All features above (P0, P1, OUT) must satisfy:

- Open a real-world `.docx`, save without edit, file is byte-identical to input.
- Open, perform a P0 edit, save, reopen — only the edited region is changed; every other part is byte-identical to the input.

This is the **only** non-negotiable acceptance criterion.

## Tracking

`docs/build-log/docx.md` keeps a row per feature with status, test reference, and any deviation from this spec.
