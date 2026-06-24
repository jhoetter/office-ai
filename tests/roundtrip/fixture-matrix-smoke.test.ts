import { readFile } from "node:fs/promises";
import { describe, expect, it } from "vitest";
import { ooxml } from "@officeai/core";
import { DocxAgent } from "@officeai/docx";
import { PdfAgent } from "@officeai/pdf";
import { PptxAgent } from "@officeai/pptx";
import { XlsxAgent } from "@officeai/xlsx";
import { fixturePath, requiredMatrixFixture, type FixtureFormat } from "../fixture-matrix.js";

const CORE_FIXTURES: FixtureFormat[] = ["docx", "xlsx", "pptx", "pdf"];

async function readFixture(format: FixtureFormat): Promise<Uint8Array> {
  const fixture = requiredMatrixFixture(format, {
    complexity: "simple",
    expectedBehavior: "noop-roundtrip",
  });
  const buf = await readFile(fixturePath(fixture));
  return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
}

describe("fixture matrix smoke", () => {
  for (const format of CORE_FIXTURES) {
    it(`${format}: imports and exports the matrix-selected core fixture`, async () => {
      const bytes = await readFixture(format);

      if (format === "docx") {
        const out = await (await DocxAgent.fromBuffer(bytes)).exportFile();
        await expect(ooxml.OoxmlContainer.load(out)).resolves.toBeTruthy();
        return;
      }

      if (format === "xlsx") {
        const out = await (await XlsxAgent.fromBuffer(bytes)).exportFile();
        await expect(ooxml.OoxmlContainer.load(out)).resolves.toBeTruthy();
        return;
      }

      if (format === "pptx") {
        const out = await (await PptxAgent.fromBuffer(bytes)).exportFile();
        await expect(ooxml.OoxmlContainer.load(out)).resolves.toBeTruthy();
        return;
      }

      const out = await (await PdfAgent.fromBuffer(bytes)).exportFile();
      expect(out.slice(0, 5)).toEqual(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]));
    });
  }
});
