# Integration adapter boundary

office-ai is an MCP-first document engine with an optional web editor. It
must stay embeddable in other products without assuming a concrete host,
storage backend, identity system, asset pipeline or MCP runtime. The
host boundary is therefore expressed as small ports in `@officeai/core`;
concrete products attach adapters at the edge.

## Core ports

`@officeai/core/integration-adapters` exports the generic extension
surface:

| Port     | Purpose                                                                                  |
| -------- | ---------------------------------------------------------------------------------------- |
| Storage  | Read/write/list/remove byte artifacts and metadata behind a logical key/path contract.   |
| Identity | Resolve the current actor when the host has one; anonymous/local use returns `null`.     |
| Assets   | Import/export DOCX, XLSX, PPTX, PDF and derived JSON/Markdown/HTML/binary artifacts.     |
| Events   | Emit product-neutral lifecycle events such as import/export/review actions.              |
| UI       | Describe how a host can open office-ai: web URL or `@officeai/react-editors` entrypoint. |
| MCP Host | Describe the MCP endpoint a host should launch or connect to.                            |

The ports are TypeScript contracts only. They do not import React, Next,
the web app, the agent package or a concrete MCP SDK. This keeps the
format packages usable standalone and lets consumers choose whether
office-ai runs as a local CLI/MCP server, embedded React editor,
custom-hosted MCP service or web editor.

## Reference local adapters

`createLocalIntegrationAdapters()` returns a complete local reference
set:

- `MemoryOfficeAiStorageAdapter` for process-local storage tests and
  examples;
- `StaticOfficeAiIdentityAdapter` for anonymous or fixed-actor runs;
- `MemoryOfficeAiAssetAdapter` for byte-backed generated artifacts;
- `MemoryOfficeAiEventHookAdapter` for inspectable event capture;
- `LocalOfficeAiUiEmbeddingAdapter` for local web URLs or
  `@officeai/react-editors` entrypoints;
- `LocalOfficeAiMcpHostAdapter` for a stdio MCP command descriptor.

The reference storage adapter is intentionally in-memory and advertises
`localPaths: false`. Production local persistence remains in
`@officeai/agent` as `LocalFilesystemSessionStorageAdapter`, but that
session adapter now implements the core `OfficeAiStorageAdapter` port.

## What belongs where

Core and format packages:

- define document models, commands, OOXML/PDF parsing and serialization;
- define adapter contracts;
- may accept adapter instances as options;
- must not import `apps/web`, `@officeai/react-editors`,
  `@modelcontextprotocol/sdk`, Next or React host code.

Agent package:

- owns CLI and MCP orchestration;
- may use the core ports for storage/asset/event integration;
- must not import the web app or React editor package.

Web app and embedding packages:

- compose concrete adapters;
- may call the agent/MCP surface or consume `@officeai/react-editors`;
- remain optional consumers, not dependencies of the core engine.

## Architecture gate

`pnpm architecture` checks two layers:

1. package-manifest dependencies via `ALLOWED_INTERNAL_DEPS` and
   `FORBIDDEN_EXTERNAL_DEPS`;
2. source imports via `FORBIDDEN_SOURCE_IMPORTS`, so undeclared hard
   imports from Core/Agent/format packages into host packages are caught
   even before they appear in `package.json`.

The gate intentionally keeps the optional PPTX React renderer exception:
`@officeai/pptx` may expose React renderer entrypoints, but it still must
not depend on the web app, agent package, MCP SDK or Next.js host code.

## Integration rule

New integrations should be represented as adapter implementations outside
the core packages. A cloud storage provider, enterprise identity system,
downstream asset registry, telemetry hook or hosted MCP service should
plug into the ports above. If a feature needs a new host capability, add
a small core port or extend an existing one before wiring a concrete
product.
