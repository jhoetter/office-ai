"use client";

import dynamic from "next/dynamic";
import { useState } from "react";
import { LoadingScreen } from "@/lib/shell";

// See `apps/web/app/editor/page.tsx` for the rationale behind owning
// the splash here instead of inside the editor: a single page-level
// mount of the spinner badge survives the dynamic-import handoff so
// the user never sees the loader "reload".
const XlsxEditor = dynamic(() => import("./XlsxEditor").then((m) => m.XlsxEditor), {
  ssr: false,
  loading: () => null,
});

export default function XlsxEditorPage(): React.ReactNode {
  const [ready, setReady] = useState(false);
  return (
    <main className="flex h-screen w-full flex-col overflow-hidden">
      <XlsxEditor onBootstrapReady={setReady} />
      <LoadingScreen variant="splash" product="xlsx" show={!ready} />
    </main>
  );
}
