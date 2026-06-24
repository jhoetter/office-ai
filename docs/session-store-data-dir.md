# Local session store and data-dir

office-ai keeps canonical MCP/web document sessions in a local data-dir.
This is the durable backing store for imported documents, created blank
documents, projections/cache metadata, command diagnostics, pending
review metadata and export history.

## Location

Resolution order:

1. Explicit `LocalSessionStore({ dataDir })` option.
2. `OFFICEAI_DATA_DIR`.
3. `$XDG_DATA_HOME/office-ai`.
4. `~/.local/share/office-ai`.

Tests set `OFFICEAI_DATA_DIR` to a temporary directory so runs are
isolated and never write into a developer's real store.

## Layout

```text
<data-dir>/
  VERSION.json
  sessions/
    session_<id>/
      session.json
      documents/
        doc_<id>/
          document.json
          artifacts/
            original.docx
            working.docx
```

The artifact extension matches the document format (`docx`, `xlsx`,
`pptx` or `pdf`). Imported files store both `original.*` and
`working.*`; created blank documents store `working.*`.

## Durability contract

- `create_session`, `import_document` and `create_document` persist the
  session/document metadata and available artifacts immediately.
- `plan_command` and `preview_command` append command-log metadata and
  diagnostics without mutating bytes.
- `apply_command`, `approve_change`, `reject_change`, `undo_command` and
  `export_document` persist the current working artifact plus command-log
  metadata.
- `list_sessions`, `list_documents`, `get_document`,
  `get_document_projection` and command lifecycle tools can hydrate
  documents from the data-dir after an MCP server restart.
- The web editor reads the same store through `GET /api/sessions` for
  path-free inspection metadata.

Writes use temp-file plus atomic rename for JSON metadata and artifacts.
Session metadata records a lightweight local lease (`pid`, host marker,
timestamp) so future cleanup/inspect tooling can reason about stale
writers without changing the file contract.

## Corruption and privacy

Metadata is versioned and schema-tagged. Invalid JSON or schema mismatch
raises `SessionStoreCorruptError`; API callers surface a clear
`corrupt-session-store` error instead of silently dropping records.

The web route deliberately omits `sourcePath`, exported paths,
`dataDir`, and artifact paths. Local paths remain available to the local
MCP/CLI process where they are needed for import/export, but browser
payloads should not expose them.

## Current boundary

Pending change metadata is durable and visible after restart. Replaying
the exact live pending stack for `approve_change` / `reject_change`
after a full server restart is not implemented yet; review actions fail
explicitly in that case instead of pretending to mutate hidden state.
