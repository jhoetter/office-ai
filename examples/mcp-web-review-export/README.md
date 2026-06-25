# MCP + web review + export

## What this proves

- A real DOCX fixture can enter the canonical MCP session flow.
- An MCP command can be previewed and queued as a pending review change.
- The same document can be opened in the web editor for human review.
- Export produces a real DOCX plus diagnostics asset metadata.

## Run

```bash
pnpm --filter @officeai/agent build
export OFFICEAI_DATA_DIR="$(pwd)/examples/_generated/mcp-web-review-export/data"
pnpm --filter @officeai/agent exec office-agent mcp
```

Replay the calls in [`mcp-transcript.json`](mcp-transcript.json) from
your MCP host. After `apply_command` returns a pending mutation, open:

```text
http://localhost:3100/sessions/<documentId>
```

Approve or reject the change in the web review panel, then run the
`export_document` step. The expected outputs are listed in
[`manifest.json`](manifest.json).

## Screenshot regeneration

After the transcript has produced `sessionId` and `documentId`, run:

```bash
make dev
node scripts/capture-example-screenshots.mjs mcp-web-review-export \
  --base-url http://localhost:3100 \
  --session-id session_... \
  --document-id doc_...
```
