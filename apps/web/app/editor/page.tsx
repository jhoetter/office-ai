"use client";

import dynamic from "next/dynamic";

const DocxEditor = dynamic(() => import("./DocxEditor").then((m) => m.DocxEditor), {
  ssr: false,
  loading: () => (
    <div className="flex h-full items-center justify-center text-sm text-secondary">Loading editor…</div>
  ),
});

export default function EditorPage() {
  return (
    <main className="flex h-screen w-full flex-col">
      <DocxEditor />
    </main>
  );
}
