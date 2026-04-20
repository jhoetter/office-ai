import type { PptxAgent } from "@officeai/pptx";
import { snapshotToTsv, type XlsxClipboardSnapshot } from "@officeai/xlsx";

/**
 * Drop an XLSX range into PPTX as a freshly-added text box on the
 * current slide.
 *
 * PPTX has no `insert-table` command yet (only edits to existing
 * `<p:graphicFrame>` tables), so we fall back to a text-box rendering
 * for now. This is the same compromise PowerPoint itself makes when
 * pasting an Excel range with "Keep Text Only" — and it round-trips
 * cleanly through OOXML because text boxes are first-class shapes.
 *
 * Sizing: the box is 80% of slide width × 60% of slide height,
 * positioned at 10%, 20% — leaves room for an existing title and
 * gives the user something obviously "newly inserted" to drag.
 *
 * Future work (see UX backlog): add a `pptx:insert-table` command
 * that materialises a real `<a:tbl>` so the cells survive a font /
 * fill change in PowerPoint without losing their row/column
 * structure.
 */
export async function applyXlsxRangeToPptx(args: {
  readonly agent: PptxAgent;
  readonly snapshot: XlsxClipboardSnapshot;
  readonly slideIndex: number;
}): Promise<void> {
  const { agent, snapshot, slideIndex } = args;
  if (snapshot.width <= 0 || snapshot.height <= 0) return;

  const tsv = snapshotToTsv(snapshot);
  const text = tsv.replace(/\t/g, "    ");

  // Slide dimensions live on the snapshot's `presentation.slideSize`.
  // We pull them from the agent so a 16:9 vs 4:3 deck both lay out
  // cleanly. Defaults match PowerPoint's 16:9 (9144000 × 6858000 EMU).
  const snap = agent.getSnapshot();
  const slideW = snap.root.slideSize?.cxEmu ?? 9_144_000;
  const slideH = snap.root.slideSize?.cyEmu ?? 6_858_000;
  const x = Math.round(slideW * 0.1);
  const y = Math.round(slideH * 0.2);
  const width = Math.round(slideW * 0.8);
  const height = Math.round(slideH * 0.6);

  const result = await agent.applyCommand({
    type: "pptx:add-text-box",
    payload: {
      slideIndex,
      text,
      x,
      y,
      width,
      height,
      name: `XLSX paste ${snapshot.origin.sheet}!${snapshot.origin.range}`,
    },
    source: "human",
  });
  if (result.rejection) {
    throw new Error(
      `pptx:add-text-box rejected: ${result.rejection.code} ${result.rejection.message ?? ""}`
    );
  }
}
