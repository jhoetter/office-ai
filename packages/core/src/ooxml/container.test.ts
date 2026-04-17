import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { OoxmlContainer } from "./container.js";

async function makeZip(files: Record<string, string | Uint8Array>): Promise<ArrayBuffer> {
  const z = new JSZip();
  for (const [path, content] of Object.entries(files)) {
    z.file(path, content);
  }
  return z.generateAsync({ type: "arraybuffer" });
}

describe("OoxmlContainer", () => {
  it("loads parts and reads them back as text and bytes", async () => {
    const buf = await makeZip({
      "[Content_Types].xml": "<Types/>",
      "word/document.xml": "<doc/>",
    });
    const c = await OoxmlContainer.load(buf);
    expect(c.has("[Content_Types].xml")).toBe(true);
    expect(c.has("word/document.xml")).toBe(true);
    expect(c.readText("word/document.xml")).toBe("<doc/>");
    expect(Array.from(c.parts.keys()).sort()).toEqual(["[Content_Types].xml", "word/document.xml"]);
  });

  it("preserves part-content bytes byte-for-byte across an unchanged save", async () => {
    const buf = await makeZip({
      "[Content_Types].xml": "<Types/>",
      "word/document.xml": "<doc>hello</doc>",
      "word/styles.xml": "<styles/>",
    });
    const c = await OoxmlContainer.load(buf);
    const out = await c.serialize();
    const c2 = await OoxmlContainer.load(out);
    for (const path of c.parts.keys()) {
      expect(c2.hash(path)).toBe(c.hash(path));
    }
  });

  it("marks dirty on writeText", async () => {
    const buf = await makeZip({ "a.xml": "<a/>" });
    const c = await OoxmlContainer.load(buf);
    expect(c.isDirty("a.xml")).toBe(false);
    c.writeText("a.xml", "<a>changed</a>");
    expect(c.isDirty("a.xml")).toBe(true);
    expect(c.readText("a.xml")).toBe("<a>changed</a>");
  });

  it("strips a UTF-8 BOM on readText but preserves it in bytes", async () => {
    const bytes = new Uint8Array([0xef, 0xbb, 0xbf, 0x3c, 0x61, 0x2f, 0x3e]); // "<a/>" with BOM
    const buf = await makeZip({ "a.xml": bytes });
    const c = await OoxmlContainer.load(buf);
    expect(c.readText("a.xml")).toBe("<a/>");
    expect(c.readBytes("a.xml")[0]).toBe(0xef);
  });

  it("addPart and removePart work and update part list", async () => {
    const buf = await makeZip({ "a.xml": "<a/>" });
    const c = await OoxmlContainer.load(buf);
    c.addPart("b.xml", new TextEncoder().encode("<b/>"));
    expect(c.has("b.xml")).toBe(true);
    c.removePart("b.xml");
    expect(c.has("b.xml")).toBe(false);
  });
});
