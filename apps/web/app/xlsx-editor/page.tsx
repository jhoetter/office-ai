"use client";

import dynamic from "next/dynamic";

const XlsxEditor = dynamic(() => import("./XlsxEditor").then((m) => m.XlsxEditor), {
  ssr: false,
  loading: () => (
    <div className="flex h-full items-center justify-center text-sm text-secondary">Loading editor…</div>
  ),
});

export default function XlsxEditorPage(): React.ReactNode {
  return (
    <main className="flex h-screen w-full flex-col">
      <XlsxEditor />
    </main>
  );
}
