# UX Audit — Night of 2026-04-20

Auditing the toolbars, status bars, and canvas interactions of all three editors against the canonical Word/Excel/PowerPoint mental model. Concrete fixes shipped in this pass live under "Fixed". Findings that didn't make tonight's cut are tracked under "Backlog" with a one-liner so the next pass has a starting point.

---

## DOCX (Word)

### Current grouping (left → right)

`paragraph-style ▸ text-format-bar (B/I/U/S, font, size, color) ▸ alignment ▸ indent ▸ spacing ▸ lists ▸ formatting-marks ▸ insert (image, table, section break) ▸ comment` and trailing `doc-info ▸ review menu ▸ edit-mode pill`.

This is already very close to the Word "Home" tab. The only material gripe is that there are **no visible group captions** — eight clusters are separated by 1 px dividers but there's nothing telling the user "this is alignment, this is paragraph". Word's toolbar has tiny labels under each cluster.

- **Fixed**: shipped a `ToolbarGroup` primitive in `@officeai/ui` (`primitives/toolbar-group.tsx`) that adds an optional 9 px tertiary caption beneath a cluster of buttons. Toolbars adopt it incrementally; nothing in the existing rows changes shape.
- **Backlog**: actually wrap the existing DOCX toolbar groups in `ToolbarGroup` and add captions ("Style", "Font", "Paragraph", "Lists", "Insert", "Review"). The wrap is mechanical; deferred to keep this pass scoped.

### Other observations

- "Show formatting marks" sits inside the lists cluster but is conceptually a _view_ toggle, not an _edit_. Belongs next to the (future) ruler / page-grid toggles in a "View" group.
- The trailing `doc-info` strip ("N paragraphs · rev N · N comments") repeats information the status bar already shows. Candidate for removal once the status bar carries the equivalent.

---

## XLSX (Excel)

### Current grouping

Toolbar header → font/format → number formats → alignment → borders/fill → insert (image/comment) → freeze/filter → format-painter, with the formula bar living below.

### Fixed

- **Formula-reference picking hint** in the status bar. When the formula bar is open (`= …` draft), the status bar now surfaces a high-contrast pill: `Click cells to insert reference · Esc to cancel`. Previously, users routinely thought clicking a cell would _deselect_ the formula draft when in fact it inserts a reference — a [classic Excel onboarding stumble][1]. Translatable via `xlsx.selection.formulaPickHint`.

### Backlog

- The selection summary (Sum / Avg / Count / Min / Max) lives only on multi-cell ranges. Add a quieter single-cell readout (cell address + type) so the status bar is never blank.
- Number-format chips (`123 ▾`) repeat the dropdown's contents — collapse into a single split button.

---

## PPTX (PowerPoint)

### Current grouping

`slide ops ▸ insert (text, shape, image, connector) ▸ arrange (forward/backward) ▸ align ▸ distribute ▸ group ▸ text alignment / anchor ▸ comment ▸ present`. Trailing: zoom control.

### Fixed

- **Selection-status hint in the status bar**: now shows `Slide N of M` plus either `N shapes selected` (accent pill) or `No selection` (dimmed). Previously the status bar was empty on the slide editor, so users had no canonical "what's selected" affordance — the only hint was the visual selection chrome on the canvas itself, which can be off-screen on small viewports. Translatable via `status.slideOf` / `status.shapesSelected` / `status.selectionEmpty`.

### Backlog

- **Crop image mode**: opening a picture should reveal a "Crop" button that swaps the resize chrome for crop handles. The plumbing exists at the OOXML level (`p:pic/p:blipFill/a:srcRect`) but the canvas has no crop UI yet. Material work; deferred.
- **Resize handles disambiguation**: corner vs edge handles use the same chrome at small sizes. PowerPoint differentiates with a tiny inset square inside the corner handle — a low-cost fidelity improvement.
- **Empty-canvas click cursor**: the slide background should show a `default` cursor (not `text`) so users know clicking it deselects.

---

## Cross-product

- **Theme/locale toggles**: the editor top bar now hosts both `LocaleToggle` and `ThemeToggle` next to each other, which read as a "preferences" pair — good. The home page header was updated to match.
- **Save-state pill**: clear and consistent across products. Translatable.
- **Find/replace panel**: title + tooltips now go through `t()`. The right-aligned "Replace / All" buttons are localised; the option toggles (`Aa`, `ab`, `.*`) keep their universal glyph but the tooltip is translated.
- **Empty state**: open-or-drop wording now adapts to the active product (`Word document` / `Excel workbook` / `PowerPoint presentation`) and locale. Drop overlay says `Drop .docx to open` etc. with a real translation.

[1]: https://support.microsoft.com/excel — Excel "click to add a cell reference while editing a formula"
