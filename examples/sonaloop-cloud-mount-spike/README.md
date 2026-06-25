# Sonaloop Cloud mount spike

This spike shows the intended `sonaloop-cloud` integration shape for
office-ai:

- the host owns object storage, identity and revisions;
- the host loads the document with a presigned GET;
- the host mounts `@officeai/react-editors/components/xlsx` with
  `initialBytes`;
- the editor's Save button calls host `onSave`;
- the host writes the returned OOXML bytes with a presigned PUT;
- no iframe is involved.

The current `sonaloop-cloud` repo is Python/FastAPI SSR and does not yet
ship a React asset pipeline. For that reason this example is a portable
React island source tree plus a Node smoke that validates the storage
contract against real local HTTP presign endpoints.

Run the smoke from the repository root:

```bash
pnpm spike:cloud-mount
```

## Host contract

The cloud host exposes two small endpoints to the browser route that
mounts the editor.

`GET /api/office-ai/presign-get?key=<cloud-object-key>` returns:

```json
{
  "url": "https://object-store.example/read-url",
  "headers": { "x-required-header": "..." },
  "etag": "\"rev-42\"",
  "filename": "revenue-model.xlsx"
}
```

`POST /api/office-ai/presign-put` receives:

```json
{
  "key": "workspaces/demo/assets/revenue-model.xlsx",
  "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "filename": "revenue-model.xlsx",
  "ifMatch": "\"rev-42\""
}
```

and returns:

```json
{
  "url": "https://object-store.example/write-url",
  "headers": { "x-required-header": "..." }
}
```

The host then sends the editor-exported bytes to that URL with
`method: "PUT"`. The object response's `ETag` becomes the next revision
shown in the host chrome.

## React mount

The example host lives in [`host/src/App.tsx`](host/src/App.tsx). The
critical shape is:

```tsx
<XlsxEditor
  initialBytes={cloudDocument.bytes}
  initialFilename={cloudDocument.filename}
  onSave={handleSave}
  room={null}
  hideLocalFileOpen
/>
```

`room={null}` keeps the spike deterministic and avoids implicit realtime
rooms. A production cloud mount can pass a stable room id derived from
the cloud object key once collaboration policy is decided.
