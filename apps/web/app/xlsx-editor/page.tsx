"use client";

import dynamic from "next/dynamic";
import { useSearchParams } from "next/navigation";
import { Suspense, useMemo, useState } from "react";
import { LoadingScreen } from "@/lib/shell";

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
  const initialSource = useMemo(() => {
    const url = params.get("src");
    if (!url) return undefined;
    const name = params.get("name") ?? url.split("/").pop() ?? "workbook.xlsx";
    return { url, name };
  }, [params]);
  const initialBlank = params.get("new") === "1";
  const [ready, setReady] = useState(false);
  return (
    <main className="flex h-screen w-full flex-col overflow-hidden">
      <XlsxEditor
        onBootstrapReady={setReady}
        initialSource={initialSource}
        initialBlank={initialBlank}
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
