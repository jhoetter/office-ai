import { readFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { CommandBus, defaultIdMinter, ooxml } from "@officeai/core";
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

describe("xlsx:add-comment — first comment on a sheet", () => {
  it("mints a comments part, dirties sheet rels + content types", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const sheetName = initial.root.sheets[0].name;
    expect(initial.root.sheets[0].commentsPartPath).toBeUndefined();
    expect(initial.root.sheets[0].comments).toEqual([]);

    const m = await bus.dispatch({
      type: "xlsx:add-comment",
      payload: { sheet: sheetName, ref: "B7", text: "Verify with finance", author: "OfficeAI" },
    });
    expect(m.status).toBe("approved");

    const snap = bus.getWorking();
    const sheet = snap.root.sheets[0];
    expect(sheet.commentsPartPath).toMatch(/^xl\/comments\d+\.xml$/);
    expect(sheet.comments).toHaveLength(1);
    expect(sheet.comments[0]).toMatchObject({
      ref: "B7",
      author: "OfficeAI",
      text: "Verify with finance",
      id: "comment-1",
    });
    expect(sheet.commentAuthors).toEqual(["OfficeAI"]);

    expect(snap.dirty.contentTypes).toBe(true);
    expect(snap.dirty.comments.has(sheet.commentsPartPath!)).toBe(true);
    const expectedSheetRels = ooxml.RelationshipGraph.relsPathFor(sheet.partPath);
    expect(snap.dirty.sheetRels.has(expectedSheetRels)).toBe(true);

    expect(m.diff.changes).toHaveLength(2);
    expect(m.diff.changes[0].kind).toBe("node-inserted");
    expect(m.diff.changes[0].summary).toBe(`Added comment by OfficeAI on ${sheetName}!B7`);
    expect(m.diff.changes[0].meta?.commentId).toBe("comment-1");
    expect(m.diff.changes[1].summary).toBe(`Created ${sheet.commentsPartPath}`);
  });

  it("emits exactly one diff change when no new part needs to be minted", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const sheetName = initial.root.sheets[0].name;

    await bus.dispatch({
      type: "xlsx:add-comment",
      payload: { sheet: sheetName, ref: "B7", text: "first", author: "OfficeAI" },
    });
    const snap1 = bus.getWorking();
    const partPath = snap1.root.sheets[0].commentsPartPath!;

    const m2 = await bus.dispatch({
      type: "xlsx:add-comment",
      payload: { sheet: sheetName, ref: "C8", text: "second", author: "OfficeAI" },
    });
    expect(m2.status).toBe("approved");

    const snap2 = bus.getWorking();
    expect(snap2.root.sheets[0].commentsPartPath).toBe(partPath);
    expect(snap2.root.sheets[0].comments).toHaveLength(2);
    expect(snap2.root.sheets[0].comments.map((c) => c.id)).toEqual(["comment-1", "comment-2"]);
    expect(m2.diff.changes).toHaveLength(1);
  });

  it("de-dupes the same author across two comments", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const sheetName = initial.root.sheets[0].name;

    await bus.dispatch({
      type: "xlsx:add-comment",
      payload: { sheet: sheetName, ref: "B7", text: "x", author: "OfficeAI" },
    });
    await bus.dispatch({
      type: "xlsx:add-comment",
      payload: { sheet: sheetName, ref: "C8", text: "y", author: "OfficeAI" },
    });
    const snap = bus.getWorking();
    expect(snap.root.sheets[0].commentAuthors).toEqual(["OfficeAI"]);
  });
});

describe("xlsx:add-comment — validation", () => {
  it("rejects comment-exists when the cell already has one", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const sheetName = initial.root.sheets[0].name;
    await bus.dispatch({
      type: "xlsx:add-comment",
      payload: { sheet: sheetName, ref: "B7", text: "x", author: "OfficeAI" },
    });
    const m = await bus.dispatch({
      type: "xlsx:add-comment",
      payload: { sheet: sheetName, ref: "B7", text: "y", author: "OfficeAI" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("comment-exists");
  });

  it("rejects empty-text", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:add-comment",
      payload: { sheet: initial.root.sheets[0].name, ref: "B7", text: "", author: "OfficeAI" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("empty-text");
  });

  it("rejects empty-author", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:add-comment",
      payload: { sheet: initial.root.sheets[0].name, ref: "B7", text: "hi", author: "" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("empty-author");
  });

  it("rejects unknown-sheet", async () => {
    const { bus } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:add-comment",
      payload: { sheet: "Nope", ref: "B7", text: "hi", author: "OfficeAI" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-sheet");
  });

  it("rejects invalid-ref for a range", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:add-comment",
      payload: { sheet: initial.root.sheets[0].name, ref: "B7:C9", text: "hi", author: "OfficeAI" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-ref");
  });

  it("rejects invalid-ref for a malformed cell", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:add-comment",
      payload: { sheet: initial.root.sheets[0].name, ref: "@@@", text: "hi", author: "OfficeAI" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-ref");
  });
});

describe("xlsx:add-comment — round-trip", () => {
  it("survives parse → add-comment → serialize → re-parse on a clean fixture", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const sheetName = initial.root.sheets[0].name;
    const originalCellValues = new Map<string, unknown>();
    for (const [k, v] of initial.root.sheets[0].cells) originalCellValues.set(k, v.value);

    await bus.dispatch({
      type: "xlsx:add-comment",
      payload: { sheet: sheetName, ref: "B7", text: "Verify with finance", author: "OfficeAI" },
    });

    const out = await serializeXlsx(bus.getWorking());
    const reparsed = await parseXlsx(new Uint8Array(out));

    const sheet = reparsed.root.sheets.find((s) => s.name === sheetName);
    expect(sheet).toBeDefined();
    expect(sheet!.comments).toHaveLength(1);
    expect(sheet!.comments[0]).toMatchObject({
      ref: "B7",
      author: "OfficeAI",
      text: "Verify with finance",
    });
    expect(sheet!.commentAuthors).toEqual(["OfficeAI"]);
    expect(sheet!.commentsPartPath).toMatch(/^xl\/comments\d+\.xml$/);

    for (const [k, v] of originalCellValues) {
      expect(sheet!.cells.get(k)?.value, `cell ${k}`).toEqual(v);
    }
  });

  it("preserves pre-existing comments on a fixture that already has them", async () => {
    const { bus, initial } = await makeBus("05-comments-hyperlinks.xlsx");
    const sheetName = initial.root.sheets[0].name;
    const before = initial.root.sheets[0];
    expect(before.comments.length).toBeGreaterThan(0);
    expect(before.commentsPartPath).toBeDefined();
    const beforeAuthors = before.commentAuthors;

    const m = await bus.dispatch({
      type: "xlsx:add-comment",
      payload: { sheet: sheetName, ref: "C1", text: "added by test", author: "OfficeAI" },
    });
    expect(m.status).toBe("approved");

    const snap = bus.getWorking();
    expect(snap.root.sheets[0].commentsPartPath).toBe(before.commentsPartPath);
    expect(snap.dirty.comments.has(before.commentsPartPath!)).toBe(true);
    expect(snap.dirty.sheetRels.size).toBe(0);

    const out = await serializeXlsx(snap);
    const reparsed = await parseXlsx(new Uint8Array(out));
    const sheet = reparsed.root.sheets[0];
    expect(sheet.comments.length).toBe(before.comments.length + 1);
    expect(sheet.commentAuthors).toEqual([...beforeAuthors, "OfficeAI"]);
    const added = sheet.comments.find((c) => c.ref === "C1");
    expect(added?.text).toBe("added by test");
    expect(added?.author).toBe("OfficeAI");
  });
});
