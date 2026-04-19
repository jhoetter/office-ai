"use client";

import dynamic from "next/dynamic";
import { useState } from "react";
import { LoadingScreen } from "@/lib/shell";

// See `apps/web/app/editor/page.tsx` for the rationale behind owning
// the splash here instead of inside the editor: a single page-level
// mount of the spinner badge survives the dynamic-import handoff so
// the user never sees the loader "reload".
const PptxEditor = dynamic(() => import("./PptxEditor").then((m) => m.PptxEditor), {
  ssr: false,
  loading: () => null,
});

export default function PptxEditorPage(): React.ReactNode {
  const [ready, setReady] = useState(false);
  return (
    <main className="flex h-screen w-full flex-col">
      <PptxEditor onBootstrapReady={setReady} />
      <LoadingScreen variant="splash" product="pptx" show={!ready} />
    </main>
  );
}
