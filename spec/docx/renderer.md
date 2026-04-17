# DOCX — Renderer

> Browser-side projection of `DocxDocument` into a ProseMirror EditorView.
> The renderer is **a skin** over the agent — never a parallel mutation path.

## Schema

ProseMirror nodes (subset that matches our model 1:1):

```
doc          → root: block+
paragraph    → group: block; content: inline*; attrs: { styleId?, alignment?, indentationJson?, spacingJson?, paragraphId, opaquePropsJson? }
table        → group: block; atom: true; attrs: { tableId, rawJson }
opaque_block → group: block; atom: true; attrs: { blockId, rawJson }
section_break→ group: block; atom: true; attrs: { blockId, rawJson }
text         → group: inline; marks: bold italic underline strikethrough font_family font_size color highlight comment_mark revision_mark hyperlink
hard_break   → group: inline; atom: true; attrs: { breakType? }
tab          → group: inline; atom: true
image        → group: inline; atom: true; attrs: { runId, drawingJson }
opaque_inline→ group: inline; atom: true; attrs: { inlineId, rawJson }
```

Marks:

```
bold, italic, underline, strikethrough            → boolean toggles
font_family   → attrs: { family: string }
font_size     → attrs: { halfPoints: number }
color         → attrs: { rgb: string }
highlight     → attrs: { name: string }
hyperlink     → attrs: { relationshipId?, anchor?, hyperlinkId }
comment_mark  → attrs: { commentId } (renders a colored span; resolved comment thread shown in side panel)
revision_mark → attrs: { revisionType: "ins"|"del", author, date, revisionId }
```

`opaqueProps`, `attrs.rawJson`, etc. carry the parser's preserveOrder
subtree as JSON so the round-trip can re-emit on save.

## model → PM doc

`docToPM(snapshot)`:

```
return PMDoc.create(schema, [], snapshot.root.body.map(blockToPM))
```

Each block becomes a corresponding PM node; each inline becomes a PM
text node with the right marks. Hyperlinks are flattened: a hyperlink
containing two runs becomes two text fragments both carrying the
`hyperlink` mark with the same `hyperlinkId`. (Adjacent text fragments
with identical marks are not coalesced — we keep run boundaries, which
makes serializing back deterministic.)

## PM doc → model (after every transaction)

`pmToDoc(pmDoc, prevSnapshot)`:

```
walk pmDoc nodes in order:
  paragraph → Paragraph { id: attrs.paragraphId ?? mintNodeId(), properties, children: marks → runs }
    - group adjacent text+marks runs into a single Run when marks are identical
    - hyperlink mark spans become Hyperlink wrappers
    - comment_mark spans become a CommentRangeStart..CommentRangeEnd wrapping the inline range
    - revision_mark spans become RevisionWrapper
  table → preserved Table from prevSnapshot.body (matched by tableId)
  opaque_block / section_break → preserved from prevSnapshot
```

The function reuses node ids from `prevSnapshot` whenever the PM node
carries a stable `*Id` attribute. Newly created PM nodes (from inserts)
get fresh ids minted here.

## The single-funnel plugin

The whole point of the renderer is that **no PM transaction reaches the
EditorView except via the command bus**. We achieve this with an EditorView
whose `dispatchTransaction` is overridden:

```typescript
const view = new EditorView(target, {
  state,
  dispatchTransaction(tx) {
    if (tx.getMeta("from-bus") === true) {
      // The bus is replaying a state into the view; let it through.
      this.updateState(this.state.apply(tx));
      return;
    }
    // Otherwise: convert to commands, dispatch through bus, and let the bus
    // call back to update the view.
    const commands = transactionToCommands(tx, this.state, agent);
    if (commands.length === 0) return;
    void agent.applyCommands(commands).catch(reportError);
  },
});

agent.subscribe((snapshot) => {
  const pmDoc = docToPM(snapshot);
  const tx = view.state.tr.replaceWith(0, view.state.doc.content.size, pmDoc.content);
  tx.setMeta("from-bus", true);
  view.dispatch(tx);
});
```

`transactionToCommands` translates standard editing transactions:

| PM transaction               | Command                              |
| ---------------------------- | ------------------------------------ |
| insertText at pos            | `docx:insert-text { at, text }`      |
| deleteRange (selection)      | `docx:delete-range { range }`        |
| addMark / removeMark over a range | `docx:format-range { range, format }` |
| Enter at end of paragraph    | `docx:insert-paragraph { at }`       |
| set styleId via menu / API   | `docx:set-paragraph-style { at, style }` |
| add comment via UI           | `docx:add-comment { range, text, author }` |

A keystroke that produces an unsupported transaction (e.g. table
insertion via the toolbar before P1 lands) is **rejected**, with a
toast: "This action is deferred — see build log."

## Visual styling for pending agent mutations

A separate ProseMirror Decoration plugin reads
`agent.getPendingMutations()` and adds inline decorations carrying the
class `pm-pending-agent` over each pending range. CSS in
`apps/web` (using the `--ai-violet` design token) renders these in the
brand "AI" purple.

## What the renderer does NOT do

- Toolbar UI lives in `apps/web`, not the renderer.
- Persisting / undo: undo is the bus's responsibility (rollback to a
  previous revision). The renderer does not register a PM history plugin.
- Any direct mutation of the model. The model is owned by the agent.

## Tests

The renderer's tests use `prosemirror-test-builder` (MIT) to construct
PM docs, run a transaction through the funnel, and assert the resulting
command sequence. No browser is required — only `jsdom`.
