"use client";

import dynamic from "next/dynamic";

const PptxEditor = dynamic(() => import("./PptxEditor").then((m) => m.PptxEditor), {
  ssr: false,
  loading: () => (
    <div className="flex h-full items-center justify-center text-sm text-secondary">Loading editor…</div>
  ),
});

export default function PptxEditorPage(): React.ReactNode {
  return (
    <main className="flex h-screen w-full flex-col">
      <PptxEditor />
    </main>
  );
}
