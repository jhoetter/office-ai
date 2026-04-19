"use client";

import dynamic from "next/dynamic";
import { LoadingScreen } from "@/lib/shell";

const PptxEditor = dynamic(() => import("./PptxEditor").then((m) => m.PptxEditor), {
  ssr: false,
  loading: () => <LoadingScreen product="pptx" label="Loading editor…" />,
});

export default function PptxEditorPage(): React.ReactNode {
  return (
    <main className="flex h-screen w-full flex-col">
      <PptxEditor />
    </main>
  );
}
