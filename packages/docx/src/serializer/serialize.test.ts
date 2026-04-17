import { ooxml, sha256Hex, deterministicIdMinter } from "@officeai/core";
import { describe, expect, it } from "vitest";
import { parseDocx } from "../parser/parse.js";
import { makeSyntheticDocx, plainDocxXml } from "../test-utils/synthetic.js";
import { serializeDocx } from "./serialize.js";

describe("serializeDocx", () => {
  it("re-emits unchanged container byte-identically (untouched parts)", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: plainDocxXml([
        { text: "Hello, world." },
        { text: "Second paragraph." },
      ]),
    });
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter() });
    const out = await serializeDocx(snap);
    const reloaded = await ooxml.OoxmlContainer.load(out);
    for (const path of snap.container.parts.keys()) {
      const beforeHash = sha256Hex(snap.container.readBytes(path));
      const afterHash = sha256Hex(reloaded.readBytes(path));
      expect(afterHash, `part ${path} should be byte-identical`).toBe(beforeHash);
    }
  });

  it("re-parses to a structurally-equivalent body after a no-op pass", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: plainDocxXml([
        { text: "Title", styleId: "Title" },
        { text: "Body content here." },
      ]),
    });
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter("a") });
    const out = await serializeDocx(snap);
    const snap2 = await parseDocx(out, { idMinter: deterministicIdMinter("b") });
    expect(snap2.root.body.length).toBe(snap.root.body.length);
    const p0 = snap.root.body[0];
    const q0 = snap2.root.body[0];
    if (p0.kind !== "paragraph" || q0.kind !== "paragraph") throw new Error();
    expect(q0.properties.styleId).toBe(p0.properties.styleId);
  });

  it("forces serialization when body is dirty and remains valid", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: plainDocxXml([{ text: "before" }]),
    });
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter() });
    const dirtied = { ...snap, dirty: { ...snap.dirty, body: true } };
    const out = await serializeDocx(dirtied);
    const reparsed = await parseDocx(out, { idMinter: deterministicIdMinter("c") });
    const p0 = reparsed.root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    const r0 = p0.children[0];
    if (r0.kind !== "run") throw new Error();
    const t = r0.children[0];
    if (t.kind !== "text") throw new Error();
    expect(t.text).toBe("before");
  });
});
