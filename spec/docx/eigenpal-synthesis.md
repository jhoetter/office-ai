# Eigenpal deep-dive — synthesis & roadmap input (P2.6 / W28+W29)

Status: research note, source for follow-up roadmap items. Not a spec.

This document captures everything we learned from a code-level read of
the two reference implementations the project tracks against:

- [`eigenpal/docx-js-editor`](https://github.com/eigenpal/docx-js-editor) — full WYSIWYG, ProseMirror-based, packaged as a React component, with a paged renderer
- [`eigenpal/docx-editor`](https://github.com/eigenpal/docx-editor) — sibling project

We cloned both at HEAD on 2026-04-17. Both repos are byte-identical
on disk excluding `.git`:

```bash
$ diff -rq docx-js-editor docx-editor --exclude=.git
# (empty output)
```

So `docx-editor` is a mirror — there is exactly one upstream
implementation to study, with two npm publish targets. Everything
below refers to the active codebase.

---

## 1. Repo shape (what they ship)

Bun + Vite + Playwright monorepo with four workspace packages:

| Package                                               | Role                                                         |
| ----------------------------------------------------- | ------------------------------------------------------------ |
| `@eigenpal/docx-core`                                 | Headless OOXML parse / serialize / mutate. No React, no DOM. |
| `@eigenpal/docx-js-editor`                            | React UI: paged editor, toolbar, plugins, sidebars.          |
| `@eigenpal/docx-editor-vue`                           | Empty Vue scaffold ("contributions welcome").                |
| `@eigenpal/docx-editor-agents` (`packages/agent-use`) | Word-like high-level wrapper aimed specifically at LLM use.  |

Reference materials they ship in-tree (we don't):

- `reference/ecma-376/part1/schemas/wml.xsd` and friends — the actual
  ECMA-376 XSDs in the repo, used as the "source of truth" they
  compare parser/serializer behaviour against.
- `reference/quick-ref/wordprocessingml.md` — distilled cheat-sheet.
- `.openspec/` — proposal/spec workflow they apparently use to gate
  changes, complete with `tasks.md` / `spec.md` / `design.md` / `testing.md`.

Quality-gate philosophy is similar to ours but heavier on Playwright:
500+ specs, narrowly scoped quick-verify patterns, explicit
`--workers=4 --timeout=30s`.

## 2. Core architecture differences vs. ours

### 2.1 Dual-rendering ProseMirror

Their headline architectural call is the **dual rendering system**:

```
HIDDEN ProseMirror (off-screen, css `left:-9999px`)
  ↳ owns selection / undo / keyboard input
        │
        ▼ state changes
VISIBLE Pages (layout-painter)
  ↳ static DOM rebuilt from PM state on every change
  ↳ owns paged display, page-break math, fidelity
```

Click → `getPositionFromMouse(x, y)` → PM `setSelection(pos)` → state
update → `layout-painter` re-renders pages.

This buys them paged WYSIWYG fidelity (page numbers, margins, headers/
footers in their actual page slots). It costs them a permanent
"two truths" maintenance burden — `CLAUDE.md` warns repeatedly that
"if you fix `toDOM` for a visual bug, the user won't see the change".

**We do not need this for the 80% scope.** Our editor is a single
ProseMirror surface inside a `max-w-[720px]` column, no paged
rendering yet. The layout-painter pattern is a P3/P4 candidate when
we want true page geometry, but introducing it before then trades a
clear architecture for two coupled rendering paths.

### 2.2 Headless API — three layers

```
DocumentAgent  (fluent: agent.insertText(...).applyStyle(...))
    ▲
    │
executeCommand(doc, command) — pure function, immutable update
    ▲
    │
AgentCommand union: insertText | replaceText | deleteText | formatText |
                    formatParagraph | applyStyle | insertTable | insertImage |
                    insertHyperlink | removeHyperlink | insertParagraphBreak |
                    mergeParagraphs | splitParagraph | setVariable | applyVariables
```

- **Position model**: `{ paragraphIndex, offset, contentIndex?, sectionIndex? }`.
  Notably no `run` index in the public API — they collapse to a
  paragraph-wide character offset, the same pragmatic call we made
  for `format-range` / `delete-range`.
- **Range model**: `{ start, end, collapsed? }` — mirrors ours.
- **Mutations are pure functions returning new `Document`** — no event
  bus, no `Mutation` envelope, no `Diff`. The agent layer owns history
  by holding a stack of `Document`s.

This is more conservative than our `Command<T,P> → Mutation → Diff`
pipeline. The trade-offs we observed:

| Their model                                      | Ours                                               |
| ------------------------------------------------ | -------------------------------------------------- |
| Pure functions, easy to reason about per-command | Bus + middleware, easier to layer policy / staging |
| No structured diff per mutation                  | `DocumentDiff` per mutation, agents can introspect |
| LLM dispatcher must replay commands serially     | LLM dispatcher can stage a batch in `pending`      |
| Built-in "applyReview" batch helper hides errors | Per-command `Mutation.status` (approved/rejected)  |

Net: our model is right for "AI proposes, human approves" because the
diff and the staging area are first-class. Theirs is right for "LLM
script executes top-to-bottom" pipelines.

### 2.3 LLM-facing surface — `agent-use` (DocxReviewer)

The most directly relevant package. ~1.5k lines in `agent-use/src/`,
deliberately scoped to the **document-review use case**:

```ts
const reviewer = await DocxReviewer.fromBuffer(buf, "AI Reviewer");

reviewer.getContentAsText(); // [0] (h1) Title\n[1] paragraph...\n[2] (table, row 1, col 1) ...
reviewer.addComment(5, "Liability cap seems too low.");
reviewer.replace(5, "$50k", "$500k");
reviewer.applyReview({
  comments: [{ paragraphIndex: 5, text: "Too low." }],
  proposals: [{ paragraphIndex: 5, search: "$50k", replaceWith: "$500k" }],
});
const out = await reviewer.toBuffer();
```

Worth stealing in spirit:

- **`getContentAsText()` projection** with paragraph-indexed `[N]`
  prefixes (their "every paragraph including table cells gets its
  own `[N]`" rule). LLMs cite text by `[index]` instead of regex,
  which sidesteps escaping bugs.
  - Our `agent.toMarkdown()` is the rough equivalent today. We do
    NOT prefix paragraphs with stable indices, and table cells get
    folded into Markdown tables. **Action:** add a parallel
    `agent.toLlmText()` that emits `[N]` per paragraph (incl. table
    cells), preserves comment anchor markers `[comment:K]`/`[/comment]`
    and tracked-change markers `[+...+]{by:...}` / `[-...-]{by:...}`.
- **Comment anchor model**: stores `commentRangeStart`/`commentRangeEnd`
  inline in paragraph content, with optional `search` to anchor to
  a substring instead of the whole paragraph.
  - We already do this in `add-comment` but always anchor to the
    selection range — we don't have the "find this substring and
    anchor to it" affordance, which is the LLM-friendly mode.
    **Action:** add an optional `search?: string` to `add-comment`
    so the agent can target text without computing offsets.
- **Tracked-change-as-proposal API** (`replace`, `proposeInsertion`,
  `proposeDeletion`). This packages "find-and-replace as a tracked
  change" instead of asking the agent to compose `delete-range +
insert-text`. Less surface, harder to drift.
  - We have `accept-change` / `reject-change` on existing tracked
    revisions but no commands that _create_ a tracked change. The
    agent currently has to insert text plain and then mark it
    manually if it wants the change to be tracked.
    **Action:** add `docx:propose-replacement`, `docx:propose-insertion`,
    `docx:propose-deletion` commands that emit revisions instead of
    plain mutations.
- **Honest batch API**: `applyReview({ accept, reject, comments,
replies, proposals })` collects per-operation errors instead of
  throwing. The return shape is `{ accepted, rejected, commentsAdded,
..., errors: [{ operation, id, error }] }`.
  - Our `applyCommands()` already returns one `Mutation` per command
    with `status: "approved" | "rejected" | "pending"` and a
    `rejection` field — equivalent semantics, slightly more verbose.
    No action needed beyond documenting that pattern in the agent docs.
- **Default-author convention**: `new DocxReviewer(doc, "AI", buf)`
  sets a default author once; every per-call `author?` overrides it.
  - Our commands take `author` per call. **Action:** add an
    `agent.defaultCommentAuthor` / `defaultRevisionAuthor` setter
    that command handlers fall back to when the payload omits author.

### 2.4 MCP server — `core/src/mcp/`

They ship an MCP server (`startStdioServer`) and a fixed set of
"core tools" for document loading + command execution:

- `docx_load { content: base64 }` → `documentId`
- `docx_save { documentId }` → `base64`
- `docx_get_content { documentId, options? }` → blocks
- `docx_get_content_as_text` → `[N]`-prefixed text
- `docx_get_changes` / `docx_get_comments`
- `docx_add_comment` / `docx_reply_to_comment` / `docx_remove_comment`
- `docx_propose_replacement` / `..._insertion` / `..._deletion`
- `docx_accept_change` / `docx_reject_change` / `..._all`
- `docx_apply_review` (batch)

Tools come from a plugin registry — `pluginRegistry.getCommandHandler(type)`
is consulted before built-in handlers, so external plugins can
override or extend the surface without modifying core.

Our `office-agent mcp` server already exposes the DOCX command
surface. The differences worth stealing:

1. **Plugin-registered tools.** Our MCP tools are hard-coded; theirs
   are discovered. We don't need plugin support for P2 but the
   `getCommandHandler(type)` indirection would let third parties
   ship `docx:*` extensions without forking the agent.
2. **Document-id session model.** Their MCP keeps a `Map<docId, LoadedDocument>`
   so multiple LLM tool calls can reference the same doc. Our MCP
   currently treats each call as a one-shot file path. **Action:**
   spike a `docx:open` MCP tool that returns a doc handle and a
   matching `docx:close`, so multi-step LLM flows don't re-parse
   the file every call.

### 2.5 Plugin architecture

Their plugins register:

- ProseMirror NodeExtension / MarkExtension / Extension (Tiptap-style
  two-phase init: `buildSchema()` then `initializeRuntime()`).
- Command handlers (`type → fn(doc, command) → doc`).
- MCP tool definitions.
- React components (toolbar buttons, dialogs).

`docxtemplater` is the canonical example — adds `{variable}` syntax,
`setVariable` / `applyVariables` commands, MCP tools, and a sidebar.

We don't have a plugin surface yet. The dual-registration pattern
(command handler + MCP tool from one definition) is the only piece
that feels P2-applicable; everything else is P3+.

### 2.6 i18n from day one

They have a typed `useTranslation()` hook, `en.json` as the canonical
map, automatic union-type derivation of locale keys, CI gate that
fails when locale files drift, and a small `validate-i18n.mjs` CLI.
Currently our toolbar strings are inline English; this is a P3
candidate.

---

## 3. Concrete fidelity gaps they cover that we don't

Reading `core/src/docx/` revealed parsers / serializers we don't yet
have. Ordered by perceived value for our 80% scope:

| Feature                               | Their file                           | Our gap                                                          |
| ------------------------------------- | ------------------------------------ | ---------------------------------------------------------------- |
| Footnotes / endnotes                  | `footnoteParser.ts`                  | No model, no parser. Currently re-emitted opaque-only.           |
| Bookmarks                             | `bookmarkParser.ts`                  | No model. Re-emitted opaque.                                     |
| Field codes (`fldChar` / `instrText`) | `fieldParser.ts`                     | We unwrap `w:fldSimple` for SDT (P2.3) but not field-code pairs. |
| Text boxes / shapes                   | `textBoxParser.ts`, `shapeParser.ts` | Treated as opaque. Mostly fine; not editable in 80%.             |
| Theme parser                          | `themeParser.ts`                     | We have nothing. Theme colour resolution is missing.             |
| Style parser                          | `styleParser.ts`                     | We do not parse `word/styles.xml`; toolbar dropdown              |
|                                       |                                      | derives style ids dynamically (good enough for now).             |
| Selective XML patch                   | `selectiveXmlPatch.ts`               | They diff only mutated paragraphs back into the original         |
|                                       |                                      | `document.xml` byte stream as a fast-path save. We always        |
|                                       |                                      | re-serialize the full body unless `dirty.body === false`.        |
| Run consolidation                     | `runConsolidator.ts`                 | We do not collapse adjacent runs with identical rPr.             |
|                                       |                                      | Drift over many edits inflates the doc.                          |

Of these, **selective XML patch** is the highest-leverage idea. Their
`attemptSelectiveSave()` rewrites only the paragraphs that actually
changed and falls back to full serialization on conflict. For long
documents like the user's masterthesis fixture this is the difference
between a 3 ms save and a 300 ms save.

---

## 4. Things they do worse than we do (sanity check)

- **Byte-preservation as a second-class citizen.** Their parser+serializer
  pair is round-trip-tested against generated XML, but they don't have
  the SHA256-equality assertion across an untouched real-world fixture
  corpus the way our `tests/roundtrip/docx/real-world-roundtrip.test.ts`
  does. Their selective-save path is the byte-preservation story; the
  full path can drift on whitespace and attribute order.
- **No diff per mutation.** Discussed above. Without it the
  "human approves" surface either has to re-render the full doc and
  visually diff, or trust the LLM's text description.
- **Dual rendering tax.** The `CLAUDE.md` warning ("if you fix `toDOM`
  the user won't see the change") is exactly the maintenance burden
  we avoided by single-pathing through ProseMirror. Worth keeping.
- **Author conventions are loose.** Author defaults to literal "AI"
  string; no agentId concept, no provenance trail across mutations.
  Our `Command.source: "human" | "agent"` + `agentId` covers this.

---

## 5. Recommended next-roadmap items

Sorted by ROI (impact × proximity to the 80% goal):

| ID      | Item                                                                                                                             | Phase |
| ------- | -------------------------------------------------------------------------------------------------------------------------------- | ----- |
| **R1**  | `agent.toLlmText()` projection: `[N]`-prefixed paragraphs incl. table cells, with inline tracked-change + comment anchor markers | P3    |
| **R2**  | Optional `search?: string` on `docx:add-comment` for substring-anchored comments                                                 | P3    |
| **R3**  | `docx:propose-replacement` / `docx:propose-insertion` / `docx:propose-deletion` commands that emit tracked revisions             | P3    |
| **R4**  | `agent.defaultAuthor` setter so commands can omit author per-call                                                                | P3    |
| **R5**  | MCP doc-handle session model (`docx:open` → `docId`, `docx:close`) so multi-step LLM flows reuse one parse                       | P3    |
| **R6**  | Selective save: serialise only dirty paragraphs back into the original `document.xml` bytes                                      | P3    |
| **R7**  | Run consolidator: collapse adjacent runs with identical `rPr` after every commit to avoid drift                                  | P3    |
| **R8**  | Footnotes / endnotes / bookmarks / field codes — typed parsers + serializers                                                     | P4    |
| **R9**  | Plugin surface: registry that owns command handlers + MCP tools + ProseMirror extensions from one definition                     | P4    |
| **R10** | Theme + style parsers (`themeParser.ts` / `styleParser.ts`) so paragraph-style dropdown stops being heuristic                    | P4    |
| **R11** | i18n hooks across toolbar and sidebars                                                                                           | P4    |
| **R12** | Paged WYSIWYG renderer (layout-painter equivalent)                                                                               | P5    |

R1–R7 are the right shape for a P3 roadmap. R8 onwards is
post-80% polish.

## 6. What we deliberately keep different

- Single ProseMirror render path. No paged dual-renderer until P5.
- `Command<T,P> → Mutation → Diff` instead of pure `executeCommand`.
- SHA256-equality byte-preservation as the central invariant, not a
  fast-save optimization.
- Headless agent + bus + UI as three layers; theirs collapses
  agent+bus into the `DocumentAgent`.

These are the architectural calls in our `prompt.md` and
`spec/docx/analysis.md` that the eigenpal review reaffirmed —
removing this section would lose the "why" if the file is read in
isolation later.

---

## 7. Sources

- `eigenpal/docx-js-editor` @ `main` (cloned 2026-04-17, depth=1)
- `eigenpal/docx-editor` @ `main` (byte-identical mirror; not separately analyzed)
- Files referenced inline:
  - `packages/core/src/headless.ts`
  - `packages/core/src/agent/{DocumentAgent,executor,context}.ts`
  - `packages/core/src/types/agentApi.ts`
  - `packages/core/src/mcp/{server,core-tools}.ts`
  - `packages/core/src/docx/*` (parsers + serializer)
  - `packages/agent-use/src/{DocxReviewer,types,content,comments,changes,batch}.ts`
  - `docs/ARCHITECTURE.md`, `CLAUDE.md`
