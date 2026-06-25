# Sonaloop Cloud mount spike

The `sonaloop-cloud-mount-spike` ticket is implemented as an Office-AI
example plus smoke gate because the current `sonaloop-cloud` repository
is still Python/FastAPI SSR and has no React/Vite asset pipeline to mount
an editor island directly.

## What is proven

- A Sonaloop-hosted React route can import
  `@officeai/react-editors/components/xlsx` directly.
- The host loads OOXML bytes before mount and passes them through
  `initialBytes`.
- Save goes through the editor's `onSave(bytes, mime, filename)` callback
  and back into host-owned object storage.
- Cloud storage is modeled as presigned GET for open and presigned PUT
  for save.
- The mount uses host semantic classes (`sl-app-main`, `sl-app-topbar`,
  `sl-card`) and the editor package CSS, with no iframe.

## Files

- [`../examples/sonaloop-cloud-mount-spike/README.md`](../examples/sonaloop-cloud-mount-spike/README.md)
  documents the host contract and mount shape.
- [`../examples/sonaloop-cloud-mount-spike/host/src/App.tsx`](../examples/sonaloop-cloud-mount-spike/host/src/App.tsx)
  is the native React island.
- [`../examples/sonaloop-cloud-mount-spike/host/src/cloud-object-store.mjs`](../examples/sonaloop-cloud-mount-spike/host/src/cloud-object-store.mjs)
  is the storage adapter used by both the host and the smoke test.
- [`../scripts/check-cloud-mount-spike.mjs`](../scripts/check-cloud-mount-spike.mjs)
  starts local HTTP presign endpoints and verifies the byte roundtrip.

Run:

```bash
pnpm spike:cloud-mount
```

## Contract for `sonaloop-cloud`

The production cloud route should provide:

- current actor/workspace identity for `presenceUser` and provenance;
- a stable object key for the office document;
- `GET /api/office-ai/presign-get?key=<key>` returning a read URL,
  required headers, `etag` and display filename;
- `POST /api/office-ai/presign-put` accepting key, content type,
  filename and current `etag`, then returning a write URL and required
  headers;
- a React asset pipeline capable of importing
  `@officeai/react-editors/components/<format>` and
  `@officeai/react-editors/styles.css`.

The spike intentionally passes `room={null}`. Production can pass a
stable room id once cloud collaboration policy exists; until then,
object open/save should stay deterministic and revision-checked.
