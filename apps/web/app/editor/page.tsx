"use client";

import dynamic from "next/dynamic";
import { LoadingScreen } from "@/lib/shell";

const DocxEditor = dynamic(() => import("./DocxEditor").then((m) => m.DocxEditor), {
  ssr: false,
  loading: () => <LoadingScreen product="docx" label="Loading editor…" />,
});

export default function EditorPage() {
  return (
    <main className="flex h-screen w-full flex-col">
      <DocxEditor />
    </main>
  );
}
