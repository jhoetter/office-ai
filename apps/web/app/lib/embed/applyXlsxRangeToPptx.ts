import type { PptxAgent } from "@officeai/pptx";
import type { XlsxClipboardSnapshot } from "@officeai/xlsx";
import { snapshotToGrid, snapshotToChartSpec, type XlsxEmbedMode } from "./xlsxEmbedShared";

/**
 * Drop an XLSX range onto a PPTX slide.
 *
 * Three modes (default `materialized`, mirrors PowerPoint's "Paste
 * Special" submenu):
 *   - `materialized`: a real `TableShape` (`<a:tbl>` graphic frame).
 *     Cells stay editable as cells, theme colors apply.
 *   - `live`: an OLE-embedded `.xlsx` package
 *     (`pptx:insert-spreadsheet`). PowerPoint activates Excel on
 *     double-click; the embedded bytes round-trip with the deck.
 *   - `chart`: a typed chart (`pptx:insert-chart`) projecting the
 *     snapshot's first row as series names and first column as
 *     categories. Embeds a backing `.xlsx` workbook so PowerPoint's
 *     "Edit Data" works.
 *
 * Sizing for materialized/live: 80% of slide width × 60% of slide
 * height, positioned at 10%, 20% — leaves room for an existing
 * title and gives the user something obviously "newly inserted" to
 * drag. Chart mode uses a smaller default (~5×3.75 in) since charts
 * usually want to sit alongside other content.
 */
export async function applyXlsxRangeToPptx(args: {
  readonly agent: PptxAgent;
  readonly snapshot: XlsxClipboardSnapshot;
  readonly slideIndex: number;
  /** Default `materialized`. */
  readonly mode?: XlsxEmbedMode;
}): Promise<void> {
  const { agent, snapshot, slideIndex } = args;
  const mode: XlsxEmbedMode = args.mode ?? "materialized";
  if (snapshot.width <= 0 || snapshot.height <= 0) return;

  // Slide dimensions live on the snapshot's `presentation.slideSize`.
  // Defaults match PowerPoint's 16:9 (9144000 × 6858000 EMU).
  const snap = agent.getSnapshot();
  const slideW = snap.root.slideSize?.cxEmu ?? 9_144_000;
  const slideH = snap.root.slideSize?.cyEmu ?? 6_858_000;
  const x = Math.round(slideW * 0.1);
  const y = Math.round(slideH * 0.2);

  if (mode === "live") {
    const cx = Math.round(slideW * 0.8);
    const cy = Math.round(slideH * 0.6);
    const result = await agent.applyCommand({
      type: "pptx:insert-spreadsheet",
      payload: {
        slideIndex,
        x,
        y,
        cx,
        cy,
        data: snapshotToGrid(snapshot),
        sheetName: snapshot.origin.sheet,
        name: `XLSX paste ${snapshot.origin.sheet}!${snapshot.origin.range}`,
      },
      source: "human",
    });
    if (result.rejection) {
      throw new Error(
        `pptx:insert-spreadsheet rejected: ${result.rejection.code} ${result.rejection.message ?? ""}`
      );
    }
    return;
  }

  if (mode === "chart") {
    const spec = snapshotToChartSpec(snapshot);
    if (!spec) {
      // Fall back to materialised when the grid can't be projected
      // into a chart (e.g. a single column / single row of values).
      return applyAsTable(agent, snapshot, slideIndex, { x, y, slideW, slideH });
    }
    const cx = 4_572_000; // ~5 in
    const cy = 3_429_000; // ~3.75 in
    const result = await agent.applyCommand({
      type: "pptx:insert-chart",
      payload: {
        slideIndex,
        x,
        y,
        cx,
        cy,
        chartType: "bar",
        categories: spec.categories,
        series: spec.series,
        name: `XLSX chart ${snapshot.origin.sheet}!${snapshot.origin.range}`,
      },
      source: "human",
    });
    if (result.rejection) {
      throw new Error(
        `pptx:insert-chart rejected: ${result.rejection.code} ${result.rejection.message ?? ""}`
      );
    }
    return;
  }

  return applyAsTable(agent, snapshot, slideIndex, { x, y, slideW, slideH });
}

async function applyAsTable(
  agent: PptxAgent,
  snapshot: XlsxClipboardSnapshot,
  slideIndex: number,
  geom: { x: number; y: number; slideW: number; slideH: number }
): Promise<void> {
  const cx = Math.round(geom.slideW * 0.8);
  const cy = Math.round(geom.slideH * 0.6);
  const result = await agent.applyCommand({
    type: "pptx:insert-table",
    payload: {
      slideIndex,
      x: geom.x,
      y: geom.y,
      cx,
      cy,
      data: snapshotToGrid(snapshot),
      name: `XLSX paste ${snapshot.origin.sheet}!${snapshot.origin.range}`,
    },
    source: "human",
  });
  if (result.rejection) {
    throw new Error(
      `pptx:insert-table rejected: ${result.rejection.code} ${result.rejection.message ?? ""}`
    );
  }
}
