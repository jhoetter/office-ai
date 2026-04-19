"use client";

import dynamic from "next/dynamic";
import { useState } from "react";
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

export default function EditorPage() {
  const [ready, setReady] = useState(false);
  return (
    <main className="flex h-screen w-full flex-col">
      <DocxEditor onBootstrapReady={setReady} />
      <LoadingScreen variant="splash" product="docx" show={!ready} />
    </main>
  );
}
