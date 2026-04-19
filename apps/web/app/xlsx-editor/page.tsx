"use client";

import dynamic from "next/dynamic";
import { LoadingScreen } from "@/lib/shell";

const XlsxEditor = dynamic(() => import("./XlsxEditor").then((m) => m.XlsxEditor), {
  ssr: false,
  loading: () => <LoadingScreen product="xlsx" label="Loading editor…" />,
});

export default function XlsxEditorPage(): React.ReactNode {
  return (
    <main className="flex h-screen w-full flex-col">
      <XlsxEditor />
    </main>
  );
}
