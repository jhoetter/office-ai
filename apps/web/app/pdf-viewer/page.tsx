"use client";

import dynamic from "next/dynamic";
import { useSearchParams } from "next/navigation";
import { Suspense, useMemo, useState } from "react";
import { LoadingScreen } from "@/lib/shell";
import { SessionEditorLoadState, useSessionEditorDocument } from "@/lib/sessions/session-editor-client";

// Mirrors `pptx-editor/page.tsx`: the splash is owned by the page
// (not the editor) so the dynamic-import handoff doesn't visually
// "reload" the loader. The `<LoadingScreen>` mounts before the
// dynamic chunk arrives and stays mounted across the boundary,
// fading out only once `PdfEditor` reports it has finished
// bootstrapping the agent + engine.
const PdfEditor = dynamic(() => import("./PdfEditor").then((m) => m.PdfEditor), {
  ssr: false,
  loading: () => null,
});

function PdfEditorPageInner(): React.ReactNode {
  const params = useSearchParams();
  const sessionDocumentId = params.get("session");
  const sessionState = useSessionEditorDocument({
    documentId: sessionDocumentId,
    expectedFormat: "pdf",
  });
  const initialSource = useMemo(() => {
    if (sessionDocumentId) return undefined;
    const url = params.get("src");
    if (!url) return undefined;
    const name = params.get("name") ?? url.split("/").pop() ?? "document.pdf";
    return { url, name };
  }, [params, sessionDocumentId]);
  const initialBlank = params.get("new") === "1";
  const [ready, setReady] = useState(false);
  if (sessionDocumentId && (sessionState.loading || sessionState.error || !sessionState.document)) {
    return <SessionEditorLoadState format="pdf" state={sessionState} />;
  }
  const sessionDocument = sessionState.document;
  return (
    <main className="flex h-screen w-full flex-col overflow-hidden">
      <PdfEditor
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
      <LoadingScreen variant="splash" product="pdf" show={!ready} />
    </main>
  );
}

export default function PdfEditorPage(): React.ReactNode {
  return (
    <Suspense fallback={<LoadingScreen variant="splash" product="pdf" show />}>
      <PdfEditorPageInner />
    </Suspense>
  );
}
