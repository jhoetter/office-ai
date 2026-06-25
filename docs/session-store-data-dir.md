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

## Storage adapter boundary

`LocalSessionStore` is backed by a `SessionStorageAdapter` port. The
default implementation is `LocalFilesystemSessionStorageAdapter`, which
maps the store to the data-dir layout above. Tests and future
integrations can pass their own adapter through
`new LocalSessionStore({ storage })`.

The adapter contract covers the storage primitives the session layer
needs:

- resolve logical child paths or keys;
- create directories/containers;
- list artifact and metadata entries;
- check existence;
- read and atomically write bytes;
- copy a local source file into the store;
- remove a subtree during cleanup.

Adapter capabilities also describe whether writes are atomic, whether
paths are local filesystem paths, and whether locks or watch/refresh are
available. The current local adapter advertises atomic writes, advisory
leases and no watcher.

Storage failures surface as `SessionStoreStorageError` with a structured
diagnostic (`level`, `code`, `message`) so MCP/web callers can report
environment or adapter failures without treating them as corrupt
document metadata.

## Corruption and privacy

Metadata is versioned and schema-tagged. Invalid JSON or schema mismatch
raises `SessionStoreCorruptError`; API callers surface a clear
`corrupt-session-store` error instead of silently dropping records.

Session metadata, document metadata and the data-dir manifest carry a
`schemaVersion`. `LocalSessionStore.inspectDataDir()` reports old or
corrupt metadata without mutating files. `migrateDataDir()` backs up
each changed metadata file under `backups/migration-*/...` before
rewriting it to the current schema. `cleanupTemporaryArtifacts()`
removes only atomic-write temp files (`.*.tmp`) from known store
directories and preserves `original.*` / `working.*` artifacts.

CLI maintenance commands:

```bash
office-agent sessions inspect --json --pretty
office-agent sessions migrate --json --pretty
office-agent sessions cleanup --json --pretty
```

`office-agent doctor` also inspects the store. Old metadata is reported
as a warning with a migration hint; corrupt metadata is reported as an
error.

CLI document commands use the same store:

```bash
office-agent sessions create --title "Q3 report"
office-agent sessions import --file report.docx --json --pretty
office-agent sessions projection --document-id doc_... --projection markdown
office-agent sessions export --document-id doc_... --out reviewed.docx
```

`sessions import` and `sessions export` are explicit edge operations:
paths enter or leave the system there, while the persisted core state is
addressed by `sessionId` and `documentId`. Imports and exports append
`office-ai/audit-log-entry@1` metadata so CLI-created documents are
visible to the MCP tools and to the web session browser.

The web route deliberately omits `sourcePath`, exported paths,
`dataDir`, and artifact paths. Local paths remain available to the local
MCP/CLI process where they are needed for import/export, but browser
payloads should not expose them.

## Current boundary

Pending change metadata is durable and visible after restart. Replaying
the exact live pending stack for `approve_change` / `reject_change`
after a full server restart is not implemented yet; review actions fail
explicitly in that case instead of pretending to mutate hidden state.
