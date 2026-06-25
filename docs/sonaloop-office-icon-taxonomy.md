# Sonaloop Office icon taxonomy

office-ai needs a consistent Office-domain icon family in
`sonaloop-design`, not one-off Lucide replacements. This taxonomy is the
input for adding missing glyphs and for the Office-AI icon adapter.

## Categories

| Category           | Scope                                                                           | Naming pattern                                                                                       |
| ------------------ | ------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------- |
| File and format    | DOCX, XLSX, PPTX, PDF, generic document, archive, image, code                   | `documentFile`, `spreadsheetFile`, `presentationFile`, `pdfFile`, `archive`, `imageFile`             |
| Text formatting    | Bold, italic, underline, strikethrough, font size, color, highlight, alignment  | `textBold`, `textItalic`, `textUnderline`, `textStrike`, `textColor`, `textHighlight`, `alignCenter` |
| Tables and grid    | Spreadsheet cells, table insert/delete, merge, filters, sort, functions, charts | `table`, `grid`, `mergeCells`, `filter`, `sortAsc`, `functionFx`, `chart`                            |
| Slides and shapes  | Slide layout, shape insert/edit, connector, group, arrange, present mode        | `slide`, `layoutTemplate`, `shape`, `connector`, `group`, `arrange`, `present`                       |
| PDF and annotation | PDF page, bookmark, link, highlight, annotation, outline                        | `pdfFile`, `bookmark`, `link`, `annotation`, `highlight`, `outline`                                  |
| Review and command | Pending change, approve, reject, undo/redo, history, diagnostics                | `check`, `close`, `undo`, `redo`, `history`, `alert`, `info`                                         |
| Export and assets  | Upload, download, export, asset handoff, package                                | `upload`, `download`, `export`, `asset`, `archive`                                                   |

## Regular vs hifi

- **Regular 24x24**: editor chrome, ribbons, toolbars, command palette,
  inspector rows, status badges.
- **Hifi 48x48**: empty states, product cards, onboarding and visual
  parity references.

Office-AI editor chrome should prefer regular icons. Hifi icons are for
larger Sonaloop app moments and should not replace toolbar glyphs.

## Mapping rule

1. If a generic Sonaloop icon already exists, reuse it (`check`,
   `close`, `search`, `folderOpen`, `link`, `alert`, `monitor`,
   `expand`, `collapse`).
2. If the concept is Office-domain-specific, add it under the category
   above instead of importing a local SVG.
3. If a Lucide name is implementation-specific (`FileType2`,
   `Loader2`, `Clock3`), map it to a semantic Sonaloop key (`pdfFile`,
   `loader`, `clock`).
4. New icon geometry lives only in `sonaloop-design/icons.data.mjs` and
   is generated to React/Python consumers.

## Verification

Run:

```bash
pnpm design:audit
```

The generated `docs/sonaloop-design-adoption.md` table must assign
every Office-AI icon need to one taxonomy category and one proposed
`IconKey`.
