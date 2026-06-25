# Sonaloop deliverable demo

## What this proves

- A neutral synthesis payload can be converted to DOCX, PPTX and XLSX.
- Provenance survives as MCP response data and visible/reviewable
  document content.
- Generated deliverables are exportable and reimportable through the
  same canonical session API.

## Run

```bash
pnpm --filter @officeai/agent build
export OFFICEAI_DATA_DIR="$(pwd)/examples/_generated/sonaloop-deliverable-demo/data"
pnpm --filter @officeai/agent exec office-agent mcp
```

Replay [`mcp-transcript.json`](mcp-transcript.json). The transcript uses
[`input/synthesis.json`](input/synthesis.json) as the payload for
`create_deliverable_from_synthesis`.
