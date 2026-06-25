"use client";

import dynamic from "next/dynamic";
import { useSearchParams } from "next/navigation";
import { Suspense, useMemo, useState } from "react";
import { LoadingScreen } from "@/lib/shell";
import { SessionEditorLoadState, useSessionEditorDocument } from "@/lib/sessions/session-editor-client";

// See `apps/web/app/editor/page.tsx` for the rationale behind owning
// the splash here instead of inside the editor: a single page-level
// mount of the spinner badge survives the dynamic-import handoff so
// the user never sees the loader "reload".
const XlsxEditor = dynamic(() => import("./XlsxEditor").then((m) => m.XlsxEditor), {
  ssr: false,
  loading: () => null,
});

function XlsxEditorPageInner(): React.ReactNode {
  const params = useSearchParams();
  const sessionDocumentId = params.get("session");
  const sessionState = useSessionEditorDocument({
    documentId: sessionDocumentId,
    expectedFormat: "xlsx",
  });
  const initialSource = useMemo(() => {
    if (sessionDocumentId) return undefined;
    const url = params.get("src");
    if (!url) return undefined;
    const name = params.get("name") ?? url.split("/").pop() ?? "workbook.xlsx";
    return { url, name };
  }, [params, sessionDocumentId]);
  const initialBlank = params.get("new") === "1";
  const [ready, setReady] = useState(false);
  if (sessionDocumentId && (sessionState.loading || sessionState.error || !sessionState.document)) {
    return <SessionEditorLoadState format="xlsx" state={sessionState} />;
  }
  const sessionDocument = sessionState.document;
  return (
    <main className="flex h-screen w-full flex-col overflow-hidden">
      <XlsxEditor
        onBootstrapReady={setReady}
        initialSource={initialSource}
        initialBlank={initialBlank}
        {...(sessionDocument
          ? {
              initialBytes: sessionDocument.bytes,
              initialFilename: sessionDocument.filename,
              onSave: sessionState.save,
              hideLocalFileOpen: true,
              room: `session:${sessionDocument.documentId}`,
            }
          : {})}
      />
      <LoadingScreen variant="splash" product="xlsx" show={!ready} />
    </main>
  );
}

export default function XlsxEditorPage(): React.ReactNode {
  return (
    <Suspense fallback={<LoadingScreen variant="splash" product="xlsx" show />}>
      <XlsxEditorPageInner />
    </Suspense>
  );
}
