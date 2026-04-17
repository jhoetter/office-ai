# DOCX — Reference Analysis (Step A)

> Notes derived from reading the public READMEs, architecture docs, and OOXML
> specification. **No code from the references is copied.** This file informs
> our spec; it is not the spec.

References surveyed (per [`prompt.md`](../../prompt.md) §Reference Repositories):

- `eigenpal/docx-js-editor` (MIT) — ProseMirror substrate, direct OOXML manipulation, tracked changes, comments, Yjs collaboration
- `eigenpal/docx-editor` (MIT) — sibling project, slightly different cut
- `dolanmiu/docx` (MIT) — mature DOCX **builder** (write only); informative for serialization edge cases
- ECMA-376 / OfficeOpenXML.com — canonical truth, prefer over any implementation

---

## 1. In-memory document model

The references converge on a **paragraph-of-runs** model that mirrors WordprocessingML:

- A `Document` is an ordered sequence of `Paragraph`s and `Table`s (block-level nodes).
- A `Paragraph` carries a `ParagraphProperties` bag (`pPr`: alignment, indentation, style, numbering reference) and an ordered list of inline children.
- An inline `Run` carries a `RunProperties` bag (`rPr`: bold, italic, font, size, color, highlight, etc.) and a list of text/break/tab/drawing leaves.
- A `Table` is rows of cells; each `Cell` is itself a sequence of paragraphs (and may contain nested tables). Cell merging is encoded in `gridSpan` (horizontal) and `vMerge` (vertical, with a `restart` value on the top cell).
- `Hyperlink` wraps runs (a single hyperlink may contain multiple runs with different formatting).
- Comments use **range markers** (`commentRangeStart`/`commentRangeEnd`) plus a separate `comments.xml` part for the actual comment text and threading metadata.
- Tracked changes wrap content in `w:ins` / `w:del` elements that carry author + date.

`dolanmiu/docx` is a **builder** API (you construct `Paragraph`, `TextRun`, `Table` objects programmatically) and serializes to OOXML. It does not have a parser at all — i.e. it cannot roundtrip an existing document, which is exactly the gap we close. We use it as a **fixture-generation helper**, not as a runtime dependency for the editor.

`eigenpal/docx-js-editor` projects the OOXML model into a ProseMirror schema for editing. The ProseMirror state is the working surface; serialization back to OOXML happens on save.

### Decisions for our model

- Mirror this paragraph-of-runs structure but add an explicit `OpaqueBlock` node so anything we don't fully understand survives roundtrip byte-for-byte.
- The model is **immutable** at the shape we expose — mutations produce new snapshots so the command bus can compute diffs.
- The model is **format-aware but renderer-agnostic** — ProseMirror is a projection, not the truth.

---

## 2. Parsing strategy

OOXML files are zip archives. Inside, `word/document.xml` is the main story; styles, numbering, comments, headers, footers, images, and relationships live in sibling parts. The `_rels/` directory describes how parts reference each other (`rId1`, `rId2`, …).

The references use a **DOM walk** (parse the XML to a tree, then walk it) rather than streaming SAX. This is fine for typical document sizes (tens of MB at most) and it makes opaque-blob preservation trivial: if we encounter an element type we don't recognize, we keep the raw subtree.

Namespace handling matters: `w:`, `r:`, `wp:`, `pic:`, `a:`, `xml:` and friends. We must keep the original namespace declarations on each part, otherwise Word will reject the file.

### Decisions for our parser

- Use `fast-xml-parser` (MIT) with `preserveOrder: true` and `attributeNamePrefix: '@_'`. Order matters in OOXML (run order, paragraph order). Attribute order does **not** matter to Word but we preserve it to maximize byte stability.
- Walk the tree element-by-element. Recognized elements become typed model nodes. Unrecognized elements become `OpaqueBlock { tag, attrs, rawXml }` nodes attached to the nearest recognized ancestor.
- Parse the `_rels/document.xml.rels` graph eagerly so paragraph-level hyperlink/comment/image references can be resolved during model construction.
- Record on each model node which OOXML part it came from (for byte-preservation accounting in the serializer).

---

## 3. Serialization & untouched-parts preservation

The hard problem is **roundtrip integrity**: a file produced by Word that we open, edit one word in, and save must be byte-clean on every part the editor did not touch. The references handle this with varying degrees of rigor; `dolanmiu/docx` does not address it at all (it builds from scratch).

Approach we will take, derived from the OOXML spec:

- The `OoxmlContainer` keeps a **byte-level cache** of every part loaded from the zip. On export, any part not explicitly marked dirty by a command is re-emitted from the cache — byte-identical.
- Parts we change (typically just `word/document.xml`, sometimes `word/comments.xml` and `word/_rels/document.xml.rels`) are re-serialized from the model.
- `[Content_Types].xml` is mostly preserved verbatim; we only touch it when adding a new content type (e.g. inserting an image of a type not previously registered).
- Relationship IDs (`rId*`) are preserved when re-serializing modified parts. New relationships get fresh IDs that don't collide.

### Decisions for our serializer

- Serializer operates on a **diff between the current model and the original parsed snapshot**. If a part's contributing model nodes are unchanged, we don't re-emit — we use the cached bytes.
- Opaque blocks emit their `rawXml` verbatim, with namespace declarations preserved.
- We never reformat XML (no pretty-printing, no whitespace normalization). We respect the original part's whitespace mode.

---

## 4. Mutation / command pattern

`eigenpal/docx-js-editor` uses ProseMirror transactions internally. Tracked changes and comments are layered on top via marks/decorations.

`dolanmiu/docx` has **no mutation pattern** — it's a one-shot builder.

### Decisions for our command bus

- The pattern from [`prompt.md`](../../prompt.md) lines 312–332 is non-negotiable. **All** edits — ProseMirror keystrokes, agent calls, CLI invocations — produce a `Command<T,P>`, are dispatched through the bus, and yield a `Mutation` with a structured `Diff`.
- The ProseMirror plugin **intercepts every transaction** at `dispatchTransaction`, translates it into one or more commands, dispatches them, then applies the resulting model state back into the editor view. Direct ProseMirror mutation outside this funnel is a bug.
- Pending agent mutations live in a separate `MutationStore` slice (`approved + pending = working`). The view renders the working state with visual marks distinguishing pending from approved (Notion-style, using the `aiViolet` design token).

---

## 5. The hard parts

### Tables with merged cells

Horizontal merging uses `w:gridSpan`. Vertical merging uses `w:vMerge` (with `val="restart"` on the top cell of a vertical merge group, omitted on continuation cells). The grid is described separately by `w:tblGrid` (column widths). Editing here is genuinely tricky: inserting a column requires updating `tblGrid` and adjusting `gridSpan` on every spanning cell.

**Plan:** parse + preserve in this session. `docx:insert-table` and `docx:set-cell-content` are stubs that throw `NotImplementedError`. Build-log entry tracks the deferral.

### Tracked changes (`w:ins` / `w:del`)

These wrap runs (or ranges of runs) and carry `w:author`, `w:date`, `w:id`. Accept means: unwrap the `w:ins` (keep content) or remove the `w:del` (drop content). Reject is the inverse.

**Plan:** parse + preserve in this session. `docx:accept-change` and `docx:reject-change` are stubs.

### Comments

`word/comments.xml` holds the comment bodies. The story uses `w:commentRangeStart`/`w:commentRangeEnd` (range markers) and `w:commentReference` (the inline marker that triggers the UI). Modern Word also writes `word/commentsExtended.xml` (resolved state, threading) and `word/commentsIds.xml` — we preserve these verbatim if we don't fully understand them.

**Plan:** add new comments (P0). Resolve / reply / delete are deferred and stubbed.

### Hyperlinks

Hyperlinks have two forms: external (resolved via a relationship — `r:id="rId7"` pointing into `_rels`) and internal (anchor to a bookmark via `w:anchor`). When inserting a hyperlink into modified content, we may need to mint a new relationship.

**Plan:** preserve in roundtrip (P0). Programmatic insertion is not in the six initial commands.

### Numbering / lists

Numbering definitions live in `word/numbering.xml` and are referenced from `pPr` via `numId` + `ilvl`. The numbering format is non-trivial (abstract numbering definitions, num overrides, level definitions).

**Plan:** preserve verbatim (P0). Setting a paragraph to a list style via `set-paragraph-style` works if the style already references a numbering definition.

---

## 6. What references get wrong / sacrifice that we improve

- **`dolanmiu/docx` cannot read existing files.** This is the central limitation. We support open → edit → save.
- **`eigenpal/docx-js-editor`** appears to focus on the visual editing experience. Strict byte-preservation of untouched parts is not its priority. We make it the central invariant.
- **None of the references treat agents as first-class users.** All have a UI-first design where edits originate from the view layer. We invert this: the agent API is primary, the UI is a skin.
- **None expose a structured diff per mutation** that an agent can introspect, approve, or reject. We make this part of the core type system.

---

## 7. What's missing from the 80% scope we need

Beyond what the references cover, our 80% requires:

- **Headless-first I/O** — the editor must work in Node with zero DOM (the references assume a browser).
- **CLI** — none ship one. We do.
- **Pending-mutation staging** — the approved/pending/working tri-state from [`prompt.md`](../../prompt.md) lines 437–451 is not modeled in the references. We add it to the core.
- **Opaque-blob preservation as an invariant**, not as a best-effort behavior. Verified by SHA-256 on every untouched part.
- **Per-feature confidence ratings** in our spec (P0/P1/P2) — keeps deferrals explicit, not implicit.

---

## Summary

Our DOCX implementation borrows the **paragraph-of-runs / typed-properties** structural intuition from the public OOXML ecosystem and the ECMA-376 spec. Everything else — the command bus, the headless agent, the staging tri-state, the strict byte-preservation, the CLI, the agent-first ergonomics — is original to this project.

Next: produce `spec/shared/*` and `spec/docx/*`.
