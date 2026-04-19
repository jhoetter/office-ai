"use client";

import dynamic from "next/dynamic";
import { useSearchParams } from "next/navigation";
import { Suspense, useMemo, useState } from "react";
import { LoadingScreen } from "@/lib/shell";

// See `apps/web/app/editor/page.tsx` for the rationale behind owning
// the splash here instead of inside the editor: a single page-level
// mount of the spinner badge survives the dynamic-import handoff so
// the user never sees the loader "reload".
const PptxEditor = dynamic(() => import("./PptxEditor").then((m) => m.PptxEditor), {
  ssr: false,
  loading: () => null,
});

function PptxEditorPageInner(): React.ReactNode {
  const params = useSearchParams();
  const initialSource = useMemo(() => {
    const url = params.get("src");
    if (!url) return undefined;
    const name = params.get("name") ?? url.split("/").pop() ?? "presentation.pptx";
    return { url, name };
  }, [params]);
  const initialBlank = params.get("new") === "1";
  const [ready, setReady] = useState(false);
  return (
    <main className="flex h-screen w-full flex-col overflow-hidden">
      <PptxEditor
        onBootstrapReady={setReady}
        initialSource={initialSource}
        initialBlank={initialBlank}
      />
      <LoadingScreen variant="splash" product="pptx" show={!ready} />
    </main>
  );
}

export default function PptxEditorPage(): React.ReactNode {
  return (
    <Suspense fallback={<LoadingScreen variant="splash" product="pptx" show />}>
      <PptxEditorPageInner />
    </Suspense>
  );
}
