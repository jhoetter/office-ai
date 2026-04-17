# office-agent CLI

The CLI is a thin shell over the `DocumentAgent` interface. Designed to be
**pipeable, scriptable, and composable** with standard UNIX tools per
[`prompt.md`](../../prompt.md) lines 453–493.

## Install / run

```bash
pnpm --filter @officeai/agent build
pnpm --filter @officeai/agent exec office-agent --help
```

Or, when published, `npm i -g @officeai/agent` and call `office-agent`.

## DOCX commands (in scope this session)

```text
office-agent docx inspect --file <path>
  → {"paragraphs": N, "tables": N, "comments": N, "tracked_changes": N, "parts": [...]}

office-agent docx read --file <path> --format markdown|json|text
  → markdown / structured / plain text projection of the document

office-agent docx write --file <path> --at <selector> --text <string>
       [--out <path>] [--source agent|human] [--agent-id <id>]
  → applies docx:insert-text; prints the Mutation as JSON to stdout

office-agent docx style --file <path> --at <selector> --style <styleId>
       [--out <path>]
  → applies docx:set-paragraph-style

office-agent docx comment --file <path> --range <selector>
       --text <string> --author <name> [--out <path>]
  → applies docx:add-comment

office-agent docx diff --before <path> --after <path>
  → structured DocumentDiff to stdout
```

### Selector grammar

`<selector>` is a slash-separated address:

```
paragraph:N                  → paragraph index N (0-based)
paragraph:N/run:M            → run M inside paragraph N
paragraph:N/text:OFFSET      → text offset OFFSET within paragraph N
section:S/paragraph:N        → paragraph N inside section S (sections deferred)
```

Range form: `paragraph:N..paragraph:M` (inclusive paragraphs), or
`paragraph:N/text:A..B` (text range within one paragraph).

### Output conventions

- All structured output is **JSON Lines** by default (one JSON object per
  line). Add `--pretty` for indented JSON.
- Errors go to stderr with `exit code 1`.
- File-not-found, parse failure, command-handler failure, and serializer
  failure each have a distinct exit code (`2`, `3`, `4`, `5`) and a
  structured error payload on stderr.

### `--out` semantics

If omitted, write back in place. If provided, write to the given path
(useful for safe round-trips in tests).

## XLSX / PPTX (deferred)

Stubbed in the CLI as `office-agent xlsx --help` and
`office-agent pptx --help` printing "deferred to a future session" and
exit code 0. This keeps the surface stable for follow-ups.

## Exit codes

| Code | Meaning                         |
| ---- | ------------------------------- |
| 0    | Success                         |
| 1    | Generic error                   |
| 2    | File not found / unreadable     |
| 3    | OOXML parse error               |
| 4    | Command handler error           |
| 5    | Serializer error                |
| 64   | Usage error (commander default) |
