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

describe("xlsx:delete-sheet", () => {
  it("removes the named sheet from the workbook and reindexes neighbours", async () => {
    const { bus, initial } = await makeBus("02-multi-sheet.xlsx");
    const beforeNames = initial.root.sheets.map((s) => s.name);
    const target = beforeNames[1];

    const m = await bus.dispatch({
      type: "xlsx:delete-sheet",
      payload: { name: target },
    });

    expect(m.status).toBe("approved");
    const snap = bus.getWorking();
    const afterNames = snap.root.sheets.map((s) => s.name);
    expect(afterNames).not.toContain(target);
    expect(afterNames).toEqual(beforeNames.filter((n) => n !== target));
    snap.root.sheets.forEach((s, i) => expect(s.index).toBe(i));
  });

  it("emits a node-deleted diff with name/sheetId/partPath meta", async () => {
    const { bus, initial } = await makeBus("02-multi-sheet.xlsx");
    const target = initial.root.sheets[1];

    const m = await bus.dispatch({
      type: "xlsx:delete-sheet",
      payload: { name: target.name },
    });
    expect(m.status).toBe("approved");
    expect(m.diff.changes).toHaveLength(1);
    const change = m.diff.changes[0];
    expect(change.kind).toBe("node-deleted");
    expect(change.nodeId).toBe(target.id);
    expect(change.meta?.name).toBe(target.name);
    expect(change.meta?.sheetId).toBe(target.sheetId);
    expect(change.meta?.partPath).toBe(target.partPath);
  });

  it("rejects deletion of an unknown sheet", async () => {
    const { bus } = await makeBus("02-multi-sheet.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:delete-sheet",
      payload: { name: "Phantom" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-sheet");
  });

  it("rejects deletion of the only worksheet", async () => {
    const { bus, initial } = await makeBus("01-single-sheet-numbers.xlsx");
    const m = await bus.dispatch({
      type: "xlsx:delete-sheet",
      payload: { name: initial.root.sheets[0].name },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-position");
  });

  it("drops the sheet part, its rels sidecar, the workbook rel, and the content-types override on serialize", async () => {
    const { bus, initial } = await makeBus("02-multi-sheet.xlsx");
    const target = initial.root.sheets[1];
    const partPath = target.partPath;
    const relsPath = ooxml.RelationshipGraph.relsPathFor(partPath);

    const before = await ooxml.OoxmlContainer.load(await loadFixture("02-multi-sheet.xlsx"));
    expect(before.has(partPath)).toBe(true);

    await bus.dispatch({ type: "xlsx:delete-sheet", payload: { name: target.name } });
    const snap = bus.getWorking();
    expect(snap.dirty.removedSheetParts.has(partPath)).toBe(true);
    expect(snap.dirty.workbook).toBe(true);
    expect(snap.dirty.rels).toBe(true);
    expect(snap.dirty.contentTypes).toBe(true);

    const out = await serializeXlsx(snap);
    const after = await ooxml.OoxmlContainer.load(out);

    expect(after.has(partPath)).toBe(false);
    expect(after.has(relsPath)).toBe(false);

    const wbRels = ooxml.RelationshipGraph.loadFor(after, "xl/workbook.xml");
    const lingering = wbRels.relationships.filter((r) => {
      const t = r.target.startsWith("/") ? r.target.slice(1) : r.target;
      const normalized = t.startsWith("xl/") ? t : `xl/${t}`;
      return normalized === partPath;
    });
    expect(lingering).toHaveLength(0);

    const ct = ooxml.ContentTypes.load(after);
    expect(ct.hasOverride(`/${partPath}`)).toBe(false);
  });

  it("re-parses cleanly after a delete and round-trips the surviving sheets", async () => {
    const { bus, initial } = await makeBus("02-multi-sheet.xlsx");
    const target = initial.root.sheets[1];
    const expected = initial.root.sheets.filter((s) => s.name !== target.name).map((s) => s.name);

    await bus.dispatch({ type: "xlsx:delete-sheet", payload: { name: target.name } });
    const snap = bus.getWorking();
    const out = await serializeXlsx(snap);

    const reparsed = await parseXlsx(new Uint8Array(out), { idMinter: defaultIdMinter });
    expect(reparsed.root.sheets.map((s) => s.name)).toEqual(expected);
  });
});
