# DOCX — Acceptance Criteria

> Done when all rows pass.

## Roundtrip integrity (the only non-negotiable)

| Fixture                                      | Open | Save unchanged byte-identical | Trivial-edit roundtrip valid |
| -------------------------------------------- | :--: | :---------------------------: | :--------------------------: |
| `fixtures/docx/synthetic-01-plain.docx`      |  ✓   |               ✓               |              ✓               |
| `fixtures/docx/synthetic-02-formatted.docx`  |  ✓   |               ✓               |              ✓               |
| `fixtures/docx/synthetic-03-with-table.docx` |  ✓   |               ✓               |              ✓               |
| (real-world fixtures listed in MANIFEST.md)  | TBD  |              TBD              |             TBD              |

A roundtrip is **valid** when:

- The output zip parses back without error.
- Every part not touched by the edit is **byte-identical** (SHA-256 match) to the loaded part.
- The touched part is **structurally equivalent** to the expected post-edit
  model when re-parsed.
- Word/LibreOffice opens the file without a repair dialog (manual check).

## Agent API

| Operation                                             | Expected                                                                 |
| ----------------------------------------------------- | ------------------------------------------------------------------------ |
| `DocxAgent.fromBuffer(buffer)`                        | returns an agent without DOM access                                      |
| `agent.getSnapshot()`                                 | returns the parsed `DocxSnapshot`                                        |
| `agent.toMarkdown()`                                  | renders body to markdown headings/paragraphs                             |
| `agent.applyCommand("docx:insert-text", ...)`         | returns Mutation; snapshot updated                                       |
| `agent.applyCommand("docx:delete-range", ...)`        | returns Mutation; snapshot updated                                       |
| `agent.applyCommand("docx:format-range", ...)`        | returns Mutation; runs split as needed                                   |
| `agent.applyCommand("docx:insert-paragraph", ...)`    | returns Mutation; new paragraph node                                     |
| `agent.applyCommand("docx:set-paragraph-style", ...)` | returns Mutation; styleId updated                                        |
| `agent.applyCommand("docx:add-comment", ...)`         | returns Mutation; comment in `comments` array                            |
| `agent.applyCommand("docx:insert-table", ...)`        | returns Mutation with `status: "rejected"`, error code `not-implemented` |
| `agent.exportFile()`                                  | returns a Buffer that round-trips                                        |
| Pending mutation flow (agent → approve)               | `status: "pending"` until `approveMutation(id)`                          |

## CLI

| Command                                                                                    | Expected                                       |
| ------------------------------------------------------------------------------------------ | ---------------------------------------------- |
| `office-agent docx inspect --file X.docx`                                                  | stdout: paragraph/table/comment counts as JSON |
| `office-agent docx read --file X.docx --format markdown`                                   | stdout: markdown                               |
| `office-agent docx write --file X.docx --at "paragraph:0" --text "X"`                      | rewrites file; stdout JSON Mutation            |
| `office-agent docx comment --file X.docx --range "paragraph:1" --text "..." --author "AI"` | rewrites file; comment added                   |
| `office-agent docx --help`                                                                 | exits 0; lists subcommands                     |

## Performance

Aspirational this session; measured on synthetic fixtures only:

- Open + parse 10-page docx: < 200 ms in Node.
- Apply 100 sequential `insert-text` commands: < 500 ms.
- Export modified docx: < 200 ms.

## Quality gates

| Gate                               | Pass criterion                                       |
| ---------------------------------- | ---------------------------------------------------- |
| `pnpm typecheck`                   | exit 0 across all packages                           |
| `pnpm test`                        | all roundtrip + agent tests green                    |
| License audit (`pnpm licenses ls`) | no AGPL anywhere                                     |
| Spec ↔ build log reconciliation    | every P0 row has a passing test or a build-log entry |
