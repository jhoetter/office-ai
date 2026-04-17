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

## DOCX commands

### Read-only

```text
office-agent docx inspect --file <path>
  → {"paragraphs": N, "tables": N, "comments": N, "tracked_changes": N, "parts": [...]}

office-agent docx read --file <path> --format markdown|json|text
  → markdown / structured / plain text projection of the document

office-agent docx search --file <path> --query <string>
       [--case-sensitive] [--regex]
  → JSON list of matches with paragraph indices and offsets

office-agent docx diff --before <path> --after <path>
  → structured DocumentDiff to stdout
```

### Text mutations

```text
office-agent docx write --file <path> --at <selector> --text <string>
       [--out <path>] [--source agent|human] [--agent-id <id>] [--no-approve]
  → applies docx:insert-text

office-agent docx format-range --file <path> --range <selector>
       [--bold true|false] [--italic true|false] [--underline true|false]
       [--strike true|false] [--font <name>] [--size <half-points>]
       [--color <hex>] [--highlight <name>]
       [--out <path>] [--no-approve]
  → applies docx:format-range; at least one formatting flag is required

office-agent docx style --file <path> --at <selector> --style <styleId>
       [--out <path>] [--no-approve]
  → applies docx:set-paragraph-style

office-agent docx set-list --file <path> --at <selector>
       --num-id <id> [--ilvl <level>] [--out <path>] [--no-approve]
office-agent docx remove-list --file <path> --at <selector>
       [--out <path>] [--no-approve]
  → applies docx:set-paragraph-list / docx:remove-paragraph-list
```

### Structural mutations

```text
office-agent docx insert-table --file <path> --at <selector>
       --rows <N> --cols <N> [--widths <csv-of-pixels>] [--out <path>] [--no-approve]
office-agent docx insert-row --file <path> --table-id <id> --at <row-index>
       [--out <path>] [--no-approve]
office-agent docx insert-column --file <path> --table-id <id> --at <col-index>
       [--out <path>] [--no-approve]
office-agent docx set-cell-text --file <path> --table-id <id>
       --row <N> --col <N> --text <string> [--out <path>] [--no-approve]

office-agent docx insert-image --file <path> --at <selector> --image <path>
       [--cx <emu>] [--cy <emu>] [--alt <text>] [--out <path>] [--no-approve]

office-agent docx insert-hyperlink --file <path> --range <selector>
       --href <url> [--tooltip <string>] [--out <path>] [--no-approve]
office-agent docx remove-hyperlink --file <path> --link-id <id>
       [--out <path>] [--no-approve]

office-agent docx header --file <path> --text <string> [--out <path>] [--no-approve]
office-agent docx footer --file <path> --text <string> [--out <path>] [--no-approve]
```

### Review surface (tracked changes & comments)

```text
office-agent docx accept-change --file <path> --change-id <id> [--out <path>]
office-agent docx reject-change --file <path> --change-id <id> [--out <path>]

office-agent docx comment --file <path> --range <selector>
       --text <string> --author <name> [--initials <str>] [--out <path>] [--no-approve]
office-agent docx resolve-comment --file <path> --comment-id <id>
       [--out <path>] [--no-approve]
office-agent docx reply-comment --file <path> --parent-id <id>
       --text <string> --author <name> [--initials <str>] [--out <path>] [--no-approve]
office-agent docx delete-comment --file <path> --comment-id <id>
       [--out <path>] [--no-approve]
```

### Pending mutation review

When any write command is invoked with `--no-approve`, the resulting
mutation is left in `pending` state on the bus instead of being committed
immediately. The pending queue is per-process — the CLI snapshots it into
the output file so that a subsequent invocation can inspect or clear it.

```text
office-agent docx pending list --file <path>
  → JSON list of pending mutation summaries (id, command type, source, agentId)
```

> **Note**: dedicated `docx pending approve <id>` and
> `docx pending reject <id>` shells are deferred until pending state is
> persisted across CLI invocations (today the bus state lives in-process,
> so an approve flow only makes sense via the MCP server).

### Generic escape hatch

```text
office-agent docx apply --file <path> --command-file <path-to-json>
       [--source agent|human|system] [--agent-id <id>] [--out <path>] [--no-approve]
  → dispatches an arbitrary Command/CommandLite parsed from JSON
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
