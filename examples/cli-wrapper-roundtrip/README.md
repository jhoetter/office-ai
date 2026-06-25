# CLI wrapper roundtrip

## What this proves

- CLI wrappers are a convenience layer over the same document engine.
- Small DOCX/XLSX/PPTX/PDF fixtures can be read through one local script.
- Session import/projection/export can be demonstrated without an MCP
  host UI.

## Run

```bash
pnpm --filter @officeai/agent build
bash examples/cli-wrapper-roundtrip/cli-transcript.sh
```

The script writes projections and exported files under
`examples/_generated/cli-wrapper-roundtrip/`.
