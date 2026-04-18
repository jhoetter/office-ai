import { readFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { CommandBus, defaultIdMinter } from "@officeai/core";
import type { XlsxSnapshot } from "../model/types.js";
import { parseXlsx } from "../parser/index.js";
import { serializeXlsx } from "../serializer/index.js";
import { allXlsxHandlers } from "./registry.js";

const here = dirname(fileURLToPath(import.meta.url));
const fixtures = resolve(here, "../../../../fixtures/xlsx/synthetic");

async function loadFixture(name: string): Promise<Uint8Array> {
  return new Uint8Array(await readFile(resolve(fixtures, name)));
}

async function makeBus(fixture: string): Promise<{
  bus: CommandBus<XlsxSnapshot>;
  initial: XlsxSnapshot;
}> {
  const buf = await loadFixture(fixture);
  const initial = await parseXlsx(buf, { idMinter: defaultIdMinter });
  const bus = new CommandBus<XlsxSnapshot>(initial);
  bus.registerAll(allXlsxHandlers);
  return { bus, initial };
}

/**
 * Smallest valid PNG payload: a 1×1 transparent pixel. Hand-rolled so
 * the suite doesn't depend on a Node `Buffer` PNG encoder. We emit
 * two distinct variants (different IDAT chunks via different
 * checksums) so the dedupe-by-hash test can verify two visually
 * different blobs both land on disk.
 */
function tinyPngA(): Uint8Array {
  // 1x1 transparent PNG (well-known canonical bytes).
  return new Uint8Array([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
    0x42, 0x60, 0x82,
  ]);
}

/**
 * A *different* tiny PNG (1x1 fully opaque red) so we can prove that
 * two distinct images coexist in `xl/media/` instead of being
 * collapsed by the hash dedupe.
 */
function tinyPngB(): Uint8Array {
  return new Uint8Array([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
    0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0x99, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
    0x00, 0x00, 0x03, 0x00, 0x01, 0x5b, 0x37, 0x88, 0x6c, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
    0x44, 0xae, 0x42, 0x60, 0x82,
  ]);
}

describe("xlsx:add-image — basic insertion", () => {
  it("appends a SheetImage and stages a fresh ImageBlob in xl/media/", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const sheetName = initial.root.sheets[0]!.name;
    const m = await bus.dispatch({
      type: "xlsx:add-image",
      payload: {
        sheet: sheetName,
        bytes: tinyPngA(),
        contentType: "image/png",
        fromRow: 2,
        fromCol: 3,
        widthPx: 96,
        heightPx: 96,
      },
    });
    expect(m.status).toBe("approved");
    const snap = bus.getWorking();
    const sheet = snap.root.sheets.find((s) => s.name === sheetName)!;
    expect(sheet.images).toHaveLength(1);
    expect(sheet.drawingPartPath).toMatch(/^xl\/drawings\/drawing\d+\.xml$/);
    expect(snap.root.images.size).toBe(1);
    const blob = [...snap.root.images.values()][0]!;
    expect(blob.contentType).toBe("image/png");
    expect(blob.partPath).toMatch(/^xl\/media\/image\d+\.png$/);
    expect(snap.dirty.drawings.has(sheet.partPath)).toBe(true);
    expect(snap.dirty.media.has(blob.partPath)).toBe(true);
    expect(snap.dirty.sheets.has(sheet.partPath)).toBe(true);
    expect(snap.dirty.contentTypes).toBe(true);
  });

  it("rejects empty bytes / non-positive dims / negative anchor", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const sheet = initial.root.sheets[0]!.name;
    const empty = await bus.dispatch({
      type: "xlsx:add-image",
      payload: {
        sheet,
        bytes: new Uint8Array(),
        contentType: "image/png",
        fromRow: 0,
        fromCol: 0,
        widthPx: 96,
        heightPx: 96,
      },
    });
    expect(empty.status).toBe("rejected");
    expect(empty.rejection?.code).toBe("empty-image");

    const bad = await bus.dispatch({
      type: "xlsx:add-image",
      payload: {
        sheet,
        bytes: tinyPngA(),
        contentType: "image/png",
        fromRow: 0,
        fromCol: 0,
        widthPx: 0,
        heightPx: 96,
      },
    });
    expect(bad.status).toBe("rejected");
    expect(bad.rejection?.code).toBe("invalid-image-size");

    const neg = await bus.dispatch({
      type: "xlsx:add-image",
      payload: {
        sheet,
        bytes: tinyPngA(),
        contentType: "image/png",
        fromRow: -1,
        fromCol: 0,
        widthPx: 96,
        heightPx: 96,
      },
    });
    expect(neg.status).toBe("rejected");
    expect(neg.rejection?.code).toBe("invalid-anchor");
  });
});

describe("xlsx:add-image — media dedupe", () => {
  it("collapses two identical-byte inserts onto one media blob", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const sheet = initial.root.sheets[0]!.name;
    const bytes = tinyPngA();
    await bus.dispatch({
      type: "xlsx:add-image",
      payload: { sheet, bytes, contentType: "image/png", fromRow: 0, fromCol: 0, widthPx: 96, heightPx: 96 },
    });
    await bus.dispatch({
      type: "xlsx:add-image",
      payload: { sheet, bytes, contentType: "image/png", fromRow: 5, fromCol: 5, widthPx: 96, heightPx: 96 },
    });
    const snap = bus.getWorking();
    expect(snap.root.images.size).toBe(1);
    expect(snap.root.sheets.find((s) => s.name === sheet)!.images).toHaveLength(2);
  });

  it("keeps two distinct media parts when bytes differ", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const sheet = initial.root.sheets[0]!.name;
    await bus.dispatch({
      type: "xlsx:add-image",
      payload: {
        sheet,
        bytes: tinyPngA(),
        contentType: "image/png",
        fromRow: 0,
        fromCol: 0,
        widthPx: 96,
        heightPx: 96,
      },
    });
    await bus.dispatch({
      type: "xlsx:add-image",
      payload: {
        sheet,
        bytes: tinyPngB(),
        contentType: "image/png",
        fromRow: 4,
        fromCol: 4,
        widthPx: 96,
        heightPx: 96,
      },
    });
    expect(bus.getWorking().root.images.size).toBe(2);
  });
});

describe("xlsx:move-image / xlsx:resize-image", () => {
  it("updates anchor + dimensions and dirties the drawing part", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const sheetName = initial.root.sheets[0]!.name;
    const add = await bus.dispatch({
      type: "xlsx:add-image",
      payload: { sheet: sheetName, bytes: tinyPngA(), contentType: "image/png", fromRow: 0, fromCol: 0, widthPx: 96, heightPx: 96 },
    });
    expect(add.status).toBe("approved");
    const sheet0 = bus.getWorking().root.sheets.find((s) => s.name === sheetName)!;
    const imageId = sheet0.images[0]!.id;

    const mv = await bus.dispatch({
      type: "xlsx:move-image",
      payload: { sheet: sheetName, imageId, fromRow: 7, fromCol: 4, fromOffsetXPx: 12, fromOffsetYPx: 3 },
    });
    expect(mv.status).toBe("approved");
    const moved = bus.getWorking().root.sheets.find((s) => s.name === sheetName)!.images[0]!;
    expect(moved.anchor.fromRow).toBe(7);
    expect(moved.anchor.fromCol).toBe(4);
    expect(moved.anchor.fromOffsetXPx).toBe(12);
    expect(moved.anchor.fromOffsetYPx).toBe(3);

    const rs = await bus.dispatch({
      type: "xlsx:resize-image",
      payload: { sheet: sheetName, imageId, widthPx: 200, heightPx: 150 },
    });
    expect(rs.status).toBe("approved");
    const resized = bus.getWorking().root.sheets.find((s) => s.name === sheetName)!.images[0]!;
    expect(resized.anchor.widthPx).toBe(200);
    expect(resized.anchor.heightPx).toBe(150);
  });

  it("rejects move/resize for unknown image ids", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const sheet = initial.root.sheets[0]!.name;
    const mv = await bus.dispatch({
      type: "xlsx:move-image",
      payload: { sheet, imageId: "does-not-exist", fromRow: 0, fromCol: 0, fromOffsetXPx: 0, fromOffsetYPx: 0 },
    });
    expect(mv.status).toBe("rejected");
    expect(mv.rejection?.code).toBe("image-not-found");
  });
});

describe("xlsx:remove-image — media garbage collection", () => {
  it("drops the orphaned blob when the last reference goes away", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const sheetName = initial.root.sheets[0]!.name;
    await bus.dispatch({
      type: "xlsx:add-image",
      payload: { sheet: sheetName, bytes: tinyPngA(), contentType: "image/png", fromRow: 0, fromCol: 0, widthPx: 96, heightPx: 96 },
    });
    const sheet0 = bus.getWorking().root.sheets.find((s) => s.name === sheetName)!;
    const imageId = sheet0.images[0]!.id;
    expect(bus.getWorking().root.images.size).toBe(1);

    const rm = await bus.dispatch({
      type: "xlsx:remove-image",
      payload: { sheet: sheetName, imageId },
    });
    expect(rm.status).toBe("approved");
    const snap = bus.getWorking();
    expect(snap.root.images.size).toBe(0);
    expect(snap.root.sheets.find((s) => s.name === sheetName)!.images).toHaveLength(0);
    expect(snap.root.sheets.find((s) => s.name === sheetName)!.drawingPartPath).toBeUndefined();
  });

  it("keeps the blob if other SheetImage entries still reference it", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const sheetName = initial.root.sheets[0]!.name;
    const bytes = tinyPngA();
    await bus.dispatch({
      type: "xlsx:add-image",
      payload: { sheet: sheetName, bytes, contentType: "image/png", fromRow: 0, fromCol: 0, widthPx: 96, heightPx: 96 },
    });
    await bus.dispatch({
      type: "xlsx:add-image",
      payload: { sheet: sheetName, bytes, contentType: "image/png", fromRow: 4, fromCol: 4, widthPx: 96, heightPx: 96 },
    });
    const sheet0 = bus.getWorking().root.sheets.find((s) => s.name === sheetName)!;
    const firstId = sheet0.images[0]!.id;
    await bus.dispatch({
      type: "xlsx:remove-image",
      payload: { sheet: sheetName, imageId: firstId },
    });
    const snap = bus.getWorking();
    expect(snap.root.images.size).toBe(1);
    expect(snap.root.sheets.find((s) => s.name === sheetName)!.images).toHaveLength(1);
  });
});

describe("xlsx:add-image — full round-trip through serialize / re-parse", () => {
  it("re-parses with sheet.images, drawing part, and identical media bytes", async () => {
    const { bus, initial } = await makeBus("02-multi-sheet.xlsx");
    const sheetName = initial.root.sheets[0]!.name;
    const bytes = tinyPngA();
    await bus.dispatch({
      type: "xlsx:add-image",
      payload: {
        sheet: sheetName,
        bytes,
        contentType: "image/png",
        fromRow: 1,
        fromCol: 2,
        fromOffsetXPx: 5,
        fromOffsetYPx: 7,
        widthPx: 128,
        heightPx: 64,
        name: "Logo",
        altText: "Company logo",
      },
    });
    const out = await serializeXlsx(bus.getWorking());
    const reparsed = await parseXlsx(new Uint8Array(out));
    const sheet = reparsed.root.sheets.find((s) => s.name === sheetName)!;
    expect(sheet.images).toHaveLength(1);
    const img = sheet.images[0]!;
    expect(img.anchor.fromRow).toBe(1);
    expect(img.anchor.fromCol).toBe(2);
    // EMU round-trip is lossy at the sub-pixel level (1 px = 9525 EMU,
    // we round to nearest int) so we tolerate ±1 px on offsets and dims.
    expect(Math.abs(img.anchor.fromOffsetXPx - 5)).toBeLessThanOrEqual(1);
    expect(Math.abs(img.anchor.fromOffsetYPx - 7)).toBeLessThanOrEqual(1);
    expect(Math.abs(img.anchor.widthPx - 128)).toBeLessThanOrEqual(1);
    expect(Math.abs(img.anchor.heightPx - 64)).toBeLessThanOrEqual(1);
    expect(img.name).toBe("Logo");
    expect(img.altText).toBe("Company logo");

    expect(reparsed.root.images.size).toBe(1);
    const blob = [...reparsed.root.images.values()][0]!;
    expect(blob.contentType).toBe("image/png");
    expect(blob.bytes.byteLength).toBe(bytes.byteLength);
    expect(Array.from(blob.bytes)).toEqual(Array.from(bytes));
  });

  it("removes the drawing part + media when the only image is deleted before save", async () => {
    const { bus, initial } = await makeBus("02-multi-sheet.xlsx");
    const sheetName = initial.root.sheets[0]!.name;
    await bus.dispatch({
      type: "xlsx:add-image",
      payload: {
        sheet: sheetName,
        bytes: tinyPngA(),
        contentType: "image/png",
        fromRow: 0,
        fromCol: 0,
        widthPx: 96,
        heightPx: 96,
      },
    });
    const sheet0 = bus.getWorking().root.sheets.find((s) => s.name === sheetName)!;
    const imageId = sheet0.images[0]!.id;
    await bus.dispatch({
      type: "xlsx:remove-image",
      payload: { sheet: sheetName, imageId },
    });
    const out = await serializeXlsx(bus.getWorking());
    const reparsed = await parseXlsx(new Uint8Array(out));
    const sheet = reparsed.root.sheets.find((s) => s.name === sheetName)!;
    expect(sheet.images).toHaveLength(0);
    expect(reparsed.root.images.size).toBe(0);
    // No xl/media/* parts should be present in the container at all.
    for (const path of reparsed.container.parts.keys()) {
      expect(path.startsWith("xl/media/")).toBe(false);
    }
  });
});
