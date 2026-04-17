import { readFile } from "node:fs/promises";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { sha256Hex } from "@officeai/core";
import { parseXlsx } from "../parser/parse.js";
import { serializeXlsx } from "./serialize.js";
import { XlsxSerializeError } from "./errors.js";

const here = dirname(fileURLToPath(import.meta.url));
const fixtures = resolve(here, "../../../../fixtures/xlsx/synthetic");

const FIXTURES = [
  "01-single-sheet-numbers.xlsx",
  "02-multi-sheet.xlsx",
  "03-formulas-basic.xlsx",
  "04-merged-and-formatted.xlsx",
  "05-comments-hyperlinks.xlsx",
  "06-large-grid.xlsx",
] as const;

async function loadFixture(name: string): Promise<Uint8Array> {
  return new Uint8Array(await readFile(resolve(fixtures, name)));
}

describe("serializeXlsx — round-trip byte-preservation", () => {
  for (const name of FIXTURES) {
    it(`${name}: every part-content survives parse → serialize → reparse byte-identical`, async () => {
      const buf = await loadFixture(name);
      const before = await parseXlsx(buf);
      const reEmitted = await serializeXlsx(before);
      const after = await parseXlsx(new Uint8Array(reEmitted));

      expect(Object.keys(after.partHashes).sort()).toEqual(Object.keys(before.partHashes).sort());
      for (const path of Object.keys(before.partHashes)) {
        expect(after.partHashes[path], `mismatch on ${path}`).toBe(before.partHashes[path]);
      }
    });

    it(`${name}: container.serialize itself preserves part-content hashes`, async () => {
      const buf = await loadFixture(name);
      const snap = await parseXlsx(buf);
      const out = await serializeXlsx(snap);
      const reparsed = await parseXlsx(new Uint8Array(out));
      for (const path of Object.keys(snap.partHashes)) {
        const original = snap.container.readBytes(path);
        const re = reparsed.container.readBytes(path);
        expect(sha256Hex(re), `byte-content drift on ${path}`).toBe(sha256Hex(original));
      }
    });
  }
});

describe("serializeXlsx — dirty-flag guard", () => {
  it("throws when caller hand-sets a dirty flag (Phase 5 will wire re-emission)", async () => {
    const buf = await loadFixture("01-single-sheet-numbers.xlsx");
    const snap = await parseXlsx(buf);
    const dirty = {
      ...snap.dirty,
      sheets: new Set<string>(["xl/worksheets/sheet1.xml"]),
    };
    const tampered = { ...snap, dirty };
    await expect(serializeXlsx(tampered)).rejects.toBeInstanceOf(XlsxSerializeError);
  });
});
