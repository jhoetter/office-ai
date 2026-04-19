"use client";

import dynamic from "next/dynamic";
import { useSearchParams } from "next/navigation";
import { Suspense, useMemo, useState } from "react";
import { LoadingScreen } from "@/lib/shell";

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
  const initialSource = useMemo(() => {
    const url = params.get("src");
    if (!url) return undefined;
    const name = params.get("name") ?? url.split("/").pop() ?? "document.docx";
    return { url, name };
  }, [params]);
  const initialBlank = params.get("new") === "1";
  const [ready, setReady] = useState(false);
  return (
    <main className="flex h-screen w-full flex-col overflow-hidden">
      <DocxEditor
        onBootstrapReady={setReady}
        initialSource={initialSource}
        initialBlank={initialBlank}
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
