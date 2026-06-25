# office-ai examples

These examples are product demos, not test fixtures. They show how a new
user can move from real DOCX/XLSX/PPTX/PDF files to MCP-first mutation,
web review and export without relying on a private host integration.

Run the structural smoke first:

```bash
pnpm examples:check
```

The smoke checks manifests, fixture paths, MCP transcripts, CLI scripts,
web screenshot specs and common secret/path leaks. Generated artifacts go
under `examples/_generated/` and are intentionally not committed.

## Examples

| Example                                                    | Proves                                                                                   |
| ---------------------------------------------------------- | ---------------------------------------------------------------------------------------- |
| [`mcp-web-review-export`](mcp-web-review-export)           | MCP imports a DOCX, queues a pending mutation, the web editor reviews it, export works.  |
| [`sonaloop-deliverable-demo`](sonaloop-deliverable-demo)   | A Sonaloop-style synthesis becomes DOCX/PPTX/XLSX deliverables with provenance.          |
| [`sonaloop-cloud-mount-spike`](sonaloop-cloud-mount-spike) | A native host React island opens XLSX bytes through presigned GET and saves through PUT. |
| [`pdf-diagnostics`](pdf-diagnostics)                       | Unsupported/high-risk PDF features produce diagnostics instead of silent success.        |
| [`cli-wrapper-roundtrip`](cli-wrapper-roundtrip)           | CLI wrappers read and export small DOCX/XLSX/PPTX/PDF fixtures through local sessions.   |

## Web screenshots

Screenshots are regenerated, not checked in. Start the local editor and
run the capture helper after the MCP transcript has produced a document
id:

```bash
make dev
node scripts/capture-example-screenshots.mjs mcp-web-review-export \
  --base-url http://localhost:3100 \
  --document-id doc_... \
  --session-id session_...
```

The helper writes PNG files to the example's `_generated` screenshot
folder. It uses Playwright when available and fails with a clear install
hint otherwise.
