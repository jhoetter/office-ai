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

async function makeBus(): Promise<{
  bus: CommandBus<XlsxSnapshot>;
  initial: XlsxSnapshot;
}> {
  const buf = await loadFixture("01-single-sheet-numbers.xlsx");
  const initial = await parseXlsx(buf, { idMinter: defaultIdMinter });
  const bus = new CommandBus<XlsxSnapshot>(initial);
  bus.registerAll(allXlsxHandlers);
  return { bus, initial };
}

async function seed(bus: CommandBus<XlsxSnapshot>, sheet: string, ref = "B7") {
  const m = await bus.dispatch({
    type: "xlsx:add-comment",
    payload: { sheet, ref, text: "Verify with finance", author: "OfficeAI" },
  });
  if (m.status !== "approved") throw new Error("seed failed");
  return bus.getWorking().root.sheets[0].comments[0].id;
}

describe("xlsx threaded-comment CRUD", () => {
  it("xlsx:reply-comment threads a reply onto a top-level comment", async () => {
    const { bus, initial } = await makeBus();
    const sheetName = initial.root.sheets[0].name;
    const parentId = await seed(bus, sheetName);

    const m = await bus.dispatch({
      type: "xlsx:reply-comment",
      payload: { sheet: sheetName, parentId, author: "Alex", text: "Looks good" },
    });
    expect(m.status).toBe("approved");

    const snap = bus.getWorking();
    const comments = snap.root.sheets[0].comments;
    expect(comments).toHaveLength(2);
    expect(comments[1]).toMatchObject({
      author: "Alex",
      text: "Looks good",
      parentId,
      ref: "B7",
    });
    expect(comments[1].createdAt).toBeDefined();
    expect(snap.root.sheets[0].commentAuthors).toEqual(["OfficeAI", "Alex"]);
    expect(m.diff.changes[0].kind).toBe("node-inserted");
  });

  it("xlsx:reply-comment rejects replies-to-replies", async () => {
    const { bus, initial } = await makeBus();
    const sheetName = initial.root.sheets[0].name;
    const parentId = await seed(bus, sheetName);
    await bus.dispatch({
      type: "xlsx:reply-comment",
      payload: { sheet: sheetName, parentId, author: "Alex", text: "first reply" },
    });
    const replyId = bus.getWorking().root.sheets[0].comments[1].id;

    const m = await bus.dispatch({
      type: "xlsx:reply-comment",
      payload: { sheet: sheetName, parentId: replyId, author: "Alex", text: "nested" },
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
  });

  it("xlsx:resolve-comment toggles the resolved flag", async () => {
    const { bus, initial } = await makeBus();
    const sheetName = initial.root.sheets[0].name;
    const parentId = await seed(bus, sheetName);

    const m1 = await bus.dispatch({
      type: "xlsx:resolve-comment",
      payload: { sheet: sheetName, commentId: parentId, resolved: true },
    });
    expect(m1.status).toBe("approved");
    expect(bus.getWorking().root.sheets[0].comments[0].resolved).toBe(true);

    const m2 = await bus.dispatch({
      type: "xlsx:resolve-comment",
      payload: { sheet: sheetName, commentId: parentId, resolved: false },
    });
    expect(m2.status).toBe("approved");
    expect(bus.getWorking().root.sheets[0].comments[0].resolved).toBe(false);
  });

  it("xlsx:edit-comment rewrites the body in-place", async () => {
    const { bus, initial } = await makeBus();
    const sheetName = initial.root.sheets[0].name;
    const parentId = await seed(bus, sheetName);

    const m = await bus.dispatch({
      type: "xlsx:edit-comment",
      payload: { sheet: sheetName, commentId: parentId, text: "rewritten" },
    });
    expect(m.status).toBe("approved");
    expect(bus.getWorking().root.sheets[0].comments[0].text).toBe("rewritten");
  });

  it("xlsx:delete-comment cascades replies when the target is top-level", async () => {
    const { bus, initial } = await makeBus();
    const sheetName = initial.root.sheets[0].name;
    const parentId = await seed(bus, sheetName);
    await bus.dispatch({
      type: "xlsx:reply-comment",
      payload: { sheet: sheetName, parentId, author: "Alex", text: "reply A" },
    });
    await bus.dispatch({
      type: "xlsx:reply-comment",
      payload: { sheet: sheetName, parentId, author: "Alex", text: "reply B" },
    });
    expect(bus.getWorking().root.sheets[0].comments).toHaveLength(3);

    const m = await bus.dispatch({
      type: "xlsx:delete-comment",
      payload: { sheet: sheetName, commentId: parentId },
    });
    expect(m.status).toBe("approved");
    expect(bus.getWorking().root.sheets[0].comments).toHaveLength(0);
  });

  it("xlsx:delete-comment on a reply leaves the parent intact", async () => {
    const { bus, initial } = await makeBus();
    const sheetName = initial.root.sheets[0].name;
    const parentId = await seed(bus, sheetName);
    await bus.dispatch({
      type: "xlsx:reply-comment",
      payload: { sheet: sheetName, parentId, author: "Alex", text: "reply" },
    });
    const replyId = bus.getWorking().root.sheets[0].comments[1].id;

    await bus.dispatch({
      type: "xlsx:delete-comment",
      payload: { sheet: sheetName, commentId: replyId },
    });
    const remaining = bus.getWorking().root.sheets[0].comments;
    expect(remaining).toHaveLength(1);
    expect(remaining[0].id).toBe(parentId);
  });

  it("threaded metadata round-trips through serialize/re-parse", async () => {
    const { bus, initial } = await makeBus();
    const sheetName = initial.root.sheets[0].name;
    const parentId = await seed(bus, sheetName);
    await bus.dispatch({
      type: "xlsx:reply-comment",
      payload: { sheet: sheetName, parentId, author: "Alex", text: "looks good" },
    });
    await bus.dispatch({
      type: "xlsx:resolve-comment",
      payload: { sheet: sheetName, commentId: parentId, resolved: true },
    });

    const out = await serializeXlsx(bus.getWorking());
    const reparsed = await parseXlsx(new Uint8Array(out));
    const sheet = reparsed.root.sheets.find((s) => s.name === sheetName)!;
    expect(sheet.comments).toHaveLength(2);
    const parent = sheet.comments.find((c) => c.id === "comment-1");
    const reply = sheet.comments.find((c) => c.id === "comment-2");
    expect(parent?.resolved).toBe(true);
    expect(parent?.createdAt).toBeDefined();
    expect(reply?.parentId).toBe("comment-1");
    expect(reply?.author).toBe("Alex");
    expect(reply?.createdAt).toBeDefined();
  });
});
