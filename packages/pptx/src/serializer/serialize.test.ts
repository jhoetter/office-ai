import { readFile, readdir } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { ooxml, sha256Hex } from "@officeai/core";
import { parsePptx } from "../parser/parse.js";
import { serializePptx } from "./serialize.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

async function listFixtures(): Promise<string[]> {
  const entries = await readdir(FIXTURES_DIR.pathname);
  return entries.filter((e) => e.endsWith(".pptx")).sort();
}

describe("serializePptx — no-edit roundtrip", () => {
  it("produces byte-identical content for every part on a no-touch pass", async () => {
    const fixtures = await listFixtures();
    expect(fixtures.length).toBeGreaterThan(0);
    for (const name of fixtures) {
      const path = join(FIXTURES_DIR.pathname, name);
      const buf = await readFile(path);
      const snap = await parsePptx(buf);
      const out = await serializePptx(snap);
      const reload = await ooxml.OoxmlContainer.load(out);

      // Every original part must round-trip with the same content hash.
      for (const [partPath] of snap.container.parts) {
        const before = sha256Hex(snap.container.readBytes(partPath));
        expect(reload.has(partPath), `${name}: part missing after roundtrip: ${partPath}`).toBe(true);
        const after = sha256Hex(reload.readBytes(partPath));
        expect(after, `${name}: ${partPath} differs after roundtrip`).toBe(before);
      }
    }
  });
});
