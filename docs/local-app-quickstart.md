# Local app quickstart

Status: 2026-06-25.

This is the supported local installation path for a cloned `office-ai`
checkout. It covers the MCP server, the web editor and the CLI wrapper
without any external product environment.

## Fresh checkout

```bash
git clone git@github.com:jhoetter/office-ai.git
cd office-ai
make install
make doctor
make smoke-local-install
```

`make smoke-local-install` is the end-to-end install smoke. It runs the
doctor, builds the web app, imports a real DOCX fixture into a temporary
session store, reads a projection, exports it, initializes the MCP stdio
server, calls `tools/list`, starts the built web editor, checks HTTP 200
and writes a screenshot to:

```text
apps/web/test-results/local-install-home.png
```

The smoke uses a temporary `OFFICEAI_DATA_DIR` and deletes it afterwards.
If Chromium is not installed for Playwright, run:

```bash
pnpm --filter @officeai/web e2e:install
```

## Start the product locally

```bash
make dev
```

Default URLs:

| Surface  | URL                            |
| -------- | ------------------------------ |
| Web      | `http://localhost:3100`        |
| Realtime | `ws://localhost:1234`          |
| Health   | `http://127.0.0.1:1234/health` |

Forwarded profiles use stable tunnel-friendly ports:

```bash
make dev-forwarded
make dev-forwarded-fugu
```

| Profile   | Web                      | Realtime               |
| --------- | ------------------------ | ---------------------- |
| Local     | `http://localhost:3100`  | `ws://localhost:1234`  |
| Forwarded | `http://localhost:23003` | `ws://localhost:21234` |
| Fugu      | `http://localhost:63003` | `ws://localhost:61234` |

Tunnel both web and realtime ports when using a remote host:

```bash
ssh -L 23003:127.0.0.1:23003 -L 21234:127.0.0.1:21234 <host>
ssh -L 63003:127.0.0.1:63003 -L 61234:127.0.0.1:61234 <fugu-host>
```

## MCP server

Build the CLI first:

```bash
make cli
```

The MCP stdio command is:

```bash
pnpm --dir /path/to/office-ai --filter @officeai/agent exec office-agent mcp
```

MCP host configs that accept `command` and `args` can use this shape:

```json
{
  "mcpServers": {
    "office-ai": {
      "command": "pnpm",
      "args": ["--dir", "/path/to/office-ai", "--filter", "@officeai/agent", "exec", "office-agent", "mcp"],
      "env": {
        "OFFICEAI_DATA_DIR": "/path/to/office-ai/.office-ai-data"
      }
    }
  }
}
```

Hosts that prefer a direct Node entrypoint can use:

```json
{
  "mcpServers": {
    "office-ai": {
      "command": "node",
      "args": ["/path/to/office-ai/packages/agent/dist/cli.js", "mcp"],
      "env": {
        "OFFICEAI_DATA_DIR": "/path/to/office-ai/.office-ai-data"
      }
    }
  }
}
```

Canonical MCP tools to use first:

- `create_session`
- `import_document`
- `create_document`
- `list_documents`
- `get_document_projection`
- `plan_command`
- `preview_command`
- `apply_command`
- `export_document`

## CLI session flow

The CLI is a wrapper around the same local session store. Paths only
enter at import/export boundaries.

```bash
export OFFICEAI_DATA_DIR="$PWD/.office-ai-data"

pnpm --filter @officeai/agent exec office-agent sessions import \
  --file fixtures/docx/real-world/01-styled-letter.docx \
  --json --pretty

pnpm --filter @officeai/agent exec office-agent sessions projection \
  --document-id doc_... \
  --projection markdown

pnpm --filter @officeai/agent exec office-agent sessions export \
  --document-id doc_... \
  --out reviewed.docx \
  --json --pretty
```

Open `http://localhost:3100` while using the same `OFFICEAI_DATA_DIR` to
review the imported document, diagnostics, pending changes, command log
and export history in the web editor.

## Fresh-clone verification

Before declaring a local install path healthy, run:

```bash
make install
make smoke-local-install
make help
```

The smoke proves:

- workspace dependencies are present;
- `doctor` passes for the built CLI/MCP runtime;
- a real Office file imports, projects and exports through session IDs;
- the MCP stdio entrypoint handles `initialize` and `tools/list`;
- the built web editor returns HTTP 200;
- a screenshot exists for visual inspection.
