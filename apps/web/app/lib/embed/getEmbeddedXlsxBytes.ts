/**
 * Resolve the bytes of the `.xlsx` workbook backing a given embedded
 * shape. Centralises the four very-similar "where does this part
 * actually live?" lookups so the editor double-click handlers stay
 * tiny and the rules for charts vs OLE objects are documented in
 * exactly one place.
 *
 * Three storage tiers, queried in order:
 *
 *   1. `snapshot.root.embeddings.get(path)?.bytes` — authoritative for
 *      everything loaded from a real package and for OLE inserts that
 *      have already been serialized once.
 *
 *   2. `pendingGrid` — fresh OLE / chart inserts whose workbook is
 *      built lazily by the serializer; we materialise it on demand
 *      via `buildEmbeddedXlsx` instead of waiting for serialize.
 *
 *   3. `snapshot.container.readBytes(path)` — chart workbooks are
 *      written into the OOXML container by the serializer but are
 *      not always mirrored back into `root.embeddings`. Fall back to
 *      the container so the user can still edit a chart that's been
 *      serialized once but never re-parsed.
 *
 * Returns `null` if no tier can produce bytes (typically a brand-new
 * chart that has never been serialized AND has no `pendingGrid`).
 */

import { buildEmbeddedXlsx } from "@officeai/xlsx";
import type { DocxAgent } from "@officeai/docx";
import type { PptxAgent } from "@officeai/pptx";

export interface EmbeddedXlsxRef {
  /** Package-absolute path of the embedded `.xlsx` part. */
  readonly embeddingPartPath: string;
  /**
   * For chart-backed embeds: the chart part path so the caller can
   * also dispatch `*:set-chart-data` once the user finishes editing.
   * `null` for OLE spreadsheet embeds.
   */
  readonly chartPartPath: string | null;
}

export type EmbeddingSource =
  | { readonly kind: "docx"; readonly agent: DocxAgent }
  | { readonly kind: "pptx"; readonly agent: PptxAgent };

/**
 * Locate the embedded-xlsx + chart paths for a given target. Pass
 * either `chartPartPath` (chart drawing / chart shape) OR
 * `embeddingPartPath` (OLE spreadsheet) — whichever the shape has on
 * its typed model.
 */
export function resolveEmbeddedXlsxRef(args: {
  readonly source: EmbeddingSource;
  readonly chartPartPath?: string;
  readonly embeddingPartPath?: string;
}): EmbeddedXlsxRef | null {
  const { source, chartPartPath, embeddingPartPath } = args;
  if (chartPartPath) {
    const charts = source.agent.getSnapshot().root.charts;
    const chart = charts.get(chartPartPath);
    if (!chart || !chart.embeddingPartPath) return null;
    return { embeddingPartPath: chart.embeddingPartPath, chartPartPath };
  }
  if (embeddingPartPath) {
    return { embeddingPartPath, chartPartPath: null };
  }
  return null;
}

/**
 * Produce the workbook bytes for {@link resolveEmbeddedXlsxRef}'s
 * resolved part path. Async because lazy `pendingGrid` inserts have
 * to round-trip through `buildEmbeddedXlsx`'s `JSZip.generateAsync`.
 */
export async function readEmbeddedXlsxBytes(args: {
  readonly source: EmbeddingSource;
  readonly embeddingPartPath: string;
}): Promise<Uint8Array | null> {
  const { source, embeddingPartPath } = args;
  const snapshot = source.agent.getSnapshot();
  const part = snapshot.root.embeddings.get(embeddingPartPath);

  if (part?.bytes) return part.bytes;

  if (part?.pendingGrid) {
    const built = await buildEmbeddedXlsx(part.pendingGrid, {
      sheetName: part.pendingSheetName ?? "Sheet1",
    });
    return built.bytes;
  }

  if (snapshot.container.has(embeddingPartPath)) {
    return snapshot.container.readBytes(embeddingPartPath);
  }

  return null;
}
