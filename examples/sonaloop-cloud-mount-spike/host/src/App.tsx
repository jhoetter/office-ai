import { useCallback, useEffect, useMemo, useState } from "react";
import { XlsxEditor } from "@officeai/react-editors/components/xlsx";
import "@officeai/react-editors/styles.css";
import { loadCloudObjectBytes, saveCloudObjectBytes } from "./cloud-object-store.mjs";

const OBJECT_KEY = "workspaces/demo/assets/revenue-model.xlsx";
const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

interface CloudDocument {
  readonly bytes: Uint8Array;
  readonly etag: string | null;
  readonly filename: string;
}

export function App() {
  const [cloudDocument, setCloudDocument] = useState<CloudDocument | null>(null);
  const [status, setStatus] = useState("Loading workbook from cloud storage...");
  const [error, setError] = useState<string | null>(null);

  const endpoints = useMemo(
    () => ({
      presignGetUrl: "/api/office-ai/presign-get",
      presignPutUrl: "/api/office-ai/presign-put",
      objectKey: OBJECT_KEY,
    }),
    []
  );

  useEffect(() => {
    let cancelled = false;
    void loadCloudObjectBytes(endpoints)
      .then((document) => {
        if (cancelled) return;
        setCloudDocument(document);
        setStatus(`Loaded ${document.filename}${document.etag ? ` at ${document.etag}` : ""}`);
      })
      .catch((err: unknown) => {
        if (cancelled) return;
        setError(err instanceof Error ? err.message : String(err));
      });
    return () => {
      cancelled = true;
    };
  }, [endpoints]);

  const handleSave = useCallback(
    async (bytes: Uint8Array, mime: string, filename: string) => {
      if (!cloudDocument) return;
      const saved = await saveCloudObjectBytes({
        ...endpoints,
        bytes,
        mime: mime || XLSX_MIME,
        filename,
        etag: cloudDocument.etag,
      });
      setCloudDocument({
        bytes,
        etag: saved.etag,
        filename,
      });
      setStatus(`Saved ${filename}${saved.etag ? ` as ${saved.etag}` : ""}`);
    },
    [cloudDocument, endpoints]
  );

  if (error) {
    return (
      <main className="sl-app-main sl-app-main--error">
        <section className="sl-card">
          <h1>Could not open cloud workbook</h1>
          <p>{error}</p>
        </section>
      </main>
    );
  }

  if (!cloudDocument) {
    return (
      <main className="sl-app-main sl-app-main--loading">
        <section className="sl-card">
          <h1>Opening cloud workbook</h1>
          <p>{status}</p>
        </section>
      </main>
    );
  }

  return (
    <main className="sl-app-main sl-app-main--editor">
      <header className="sl-app-topbar">
        <div>
          <h1>Office AI cloud workbook</h1>
          <p>{status}</p>
        </div>
      </header>
      <XlsxEditor
        initialBytes={cloudDocument.bytes}
        initialFilename={cloudDocument.filename}
        onSave={handleSave}
        locale="en"
        theme="light"
        room={null}
        hideLocalFileOpen
      />
    </main>
  );
}
