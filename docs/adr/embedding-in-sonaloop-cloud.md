# ADR: native office-ai embedding in sonaloop-cloud

Status: accepted

Date: 2026-06-25

## Context

office-ai exposes `@officeai/react-editors` for host apps and
`@officeai/core/integration-adapters` for storage, identity, assets,
events, UI embedding and MCP hosting. `docs/embedding.md` already
rejects iframe embedding because the editor should share host theme,
input handling and event loop.

`sonaloop-cloud` is Python-first. It can provide object storage,
identity, presigned asset access and its own app routing, but it should
not become a dependency of the office-ai core packages.

## Options

| Option                                           | Assessment                                                                                          |
| ------------------------------------------------ | --------------------------------------------------------------------------------------------------- |
| React island only                                | Best human UX and shared theme, but does not cover agent automation by itself.                      |
| office-ai web/MCP service only                   | Simple deployment boundary, but feels remote and duplicates app chrome.                             |
| Hybrid: MCP for agents + React island for people | Best fit: native UX for humans, stable MCP/session API for agents, and explicit adapter boundaries. |

## Decision

Use the **hybrid path**:

- people open embedded React editor islands from
  `@officeai/react-editors/components/<format>`;
- agents use office-ai MCP/session tools;
- `sonaloop-cloud` supplies storage/identity through adapter
  implementations and presigned object-byte GET/PUT;
- no iframe is used.

## Adapter plan

| Port     | sonaloop-cloud mapping                                                                                           |
| -------- | ---------------------------------------------------------------------------------------------------------------- |
| Storage  | Object storage asset key plus revision metadata. Use presigned GET for editor load and presigned PUT for save.   |
| Identity | Current authenticated workspace actor, passed as `OfficeAiActor` and command provenance.                         |
| Assets   | Existing cloud asset records map to `OfficeAiAssetRef`; exports return hash, format, size and source refs.       |
| Events   | Cloud audit/event stream receives import, edit, save, export and diagnostic events.                              |
| UI       | Cloud route mounts a format-specific React island with Sonaloop app shell/tokens.                                |
| MCP      | Hosted MCP profile points to office-ai agent runtime or a sidecar service using the same session IDs/asset refs. |

## Consequences

- `@officeai/react-editors` must keep per-format entrypoints and bundle
  budgets.
- Heavy format dependencies must lazy-load by format so a cloud page
  opening XLSX does not pull PDF/DOCX stacks.
- Save-back must be explicit: editor bytes go to a host `onSave`
  callback in cloud mode or a session save endpoint in local web mode.
- The local `apps/web` shell remains the reference implementation and
  smoke surface, not the only deployment shape.

## Follow-up tickets

- `sonaloop-cloud-mount-spike`
- `dependency-budget-gate`
- `per-format-lazy-loading`
- `session-bytes-load-endpoint`
- `editor-save-back-to-session`
