"use client";

import { useMemo } from "react";
import type { ExportFormat } from "@/lib/shell";

/**
 * Export catalogue for the PDF viewer.
 *
 * Three formats:
 *
 *   - **PDF** — round-trips the agent buffer via incremental save
 *     (`PdfAgent.exportFile()`). The native format and the most
 *     common ask, so it leads.
 *   - **Markdown** — `PdfAgent.toMarkdown()` produces a best-effort
 *     reading-order projection. Useful for piping into LLMs.
 *   - **Plain text** — concatenates per-page text. Lossy by design;
 *     no headings, no tables — for downstream search / grep.
 *
 * Returned as a `ReadonlyArray` so the editor's `useMemo` consumer
 * can pass it straight into the {@link ProductAdapter}.
 */
export function usePdfFormatProvider(): ReadonlyArray<ExportFormat> {
  return useMemo<ReadonlyArray<ExportFormat>>(
    () => [
      {
        id: "pdf",
        label: "PDF document (.pdf)",
        description: "Incremental save preserves untouched objects byte-for-byte.",
        extension: "pdf",
        mime: "application/pdf",
        kind: "instant",
        group: "deck",
        icon: "pdf",
      },
      {
        id: "markdown",
        label: "Markdown (.md)",
        description: "Reading-order projection. Best for handing to an LLM.",
        extension: "md",
        mime: "text/markdown",
        kind: "instant",
        group: "deck",
        icon: "text",
      },
      {
        id: "text",
        label: "Plain text (.txt)",
        description: "Concatenated per-page text. Lossy — no layout preserved.",
        extension: "txt",
        mime: "text/plain",
        kind: "instant",
        group: "deck",
        icon: "text",
      },
    ],
    []
  );
}
