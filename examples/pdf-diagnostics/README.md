# PDF diagnostics

## What this proves

- PDF is part of the same canonical MCP session flow.
- High-risk PDF inputs surface explicit diagnostics.
- A demo can show limitations without pretending the happy path worked.

## Run

Use [`mcp-transcript.json`](mcp-transcript.json) against
`office-agent mcp`. The input is the signed fixture
`fixtures/pdf/signed-then-modified.pdf`, which should report signature
diagnostics and export policy caveats.
