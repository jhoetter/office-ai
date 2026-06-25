"use client";

import dynamic from "next/dynamic";
import { useSearchParams } from "next/navigation";
import { Suspense, useMemo, useState } from "react";
import { LoadingScreen } from "@/lib/shell";
import { SessionEditorLoadState, useSessionEditorDocument } from "@/lib/sessions/session-editor-client";

// The page owns the bootstrap splash so the badge `<span>` is mounted
// exactly once, from page hydration through agent-ready. Returning
// `null` from the dynamic loading callback keeps the editor subtree
// transparent during the chunk fetch — the page-level splash is the
// only loader the user sees, and it never remounts across the
// `dynamic()` resolution → editor mount → agent-ready handoff. See
// `LoadingScreen` for the splash variant's JSDoc.
const DocxEditor = dynamic(() => import("./DocxEditor").then((m) => m.DocxEditor), {
  ssr: false,
  loading: () => null,
});

function EditorPageInner() {
  const params = useSearchParams();
  const sessionDocumentId = params.get("session");
  const sessionState = useSessionEditorDocument({
    documentId: sessionDocumentId,
    expectedFormat: "docx",
  });
  const initialSource = useMemo(() => {
    if (sessionDocumentId) return undefined;
    const url = params.get("src");
    if (!url) return undefined;
    const name = params.get("name") ?? url.split("/").pop() ?? "document.docx";
    return { url, name };
  }, [params, sessionDocumentId]);
  const initialBlank = params.get("new") === "1";
  const [ready, setReady] = useState(false);
  if (sessionDocumentId && (sessionState.loading || sessionState.error || !sessionState.document)) {
    return <SessionEditorLoadState format="docx" state={sessionState} />;
  }
  const sessionDocument = sessionState.document;
  return (
    <main className="flex h-screen w-full flex-col overflow-hidden">
      <DocxEditor
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
      <LoadingScreen variant="splash" product="docx" show={!ready} />
    </main>
  );
}

export default function EditorPage() {
  return (
    <Suspense fallback={<LoadingScreen variant="splash" product="docx" show />}>
      <EditorPageInner />
    </Suspense>
  );
}
