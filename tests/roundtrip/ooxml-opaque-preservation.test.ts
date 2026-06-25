import { readFile } from "node:fs/promises";
import { describe, expect, it } from "vitest";
import { ooxml, sha256Hex } from "@officeai/core";
import { DocxAgent } from "@officeai/docx";
import { PptxAgent } from "@officeai/pptx";
import { XlsxAgent } from "@officeai/xlsx";
import { fixturePath, requiredMatrixFixture } from "../fixture-matrix.js";

type Format = "docx" | "xlsx" | "pptx";

interface PreservationProbe {
  readonly format: Format;
  readonly fixturePath: string;
  readonly relsPath: string;
  readonly xmlPart: string;
  readonly binPart: string;
  readonly touchedPart: string;
  readonly mutate: (input: Uint8Array) => Promise<ArrayBuffer>;
  readonly roundtrip: (input: Uint8Array) => Promise<ArrayBuffer>;
}

const TEXT_ENCODER = new TextEncoder();
const CONTENT_TYPES = "[Content_Types].xml";
const OPAQUE_XML_REL_ID = "rIdOfficeAiOpaqueXml";
const OPAQUE_BIN_REL_ID = "rIdOfficeAiOpaqueBin";
const CUSTOM_XML_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
const OLE_OBJECT_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";

const PROBES: ReadonlyArray<PreservationProbe> = [
  {
    format: "docx",
    fixturePath: fixturePath(requiredMatrixFixture("docx", { id: "docx.real-shape.styled-letter" })),
    relsPath: "word/_rels/document.xml.rels",
    xmlPart: "customXml/officeAiOpaque-docx.xml",
    binPart: "word/embeddings/officeAiOpaque-docx.bin",
    touchedPart: "word/document.xml",
    roundtrip: async (input) => (await DocxAgent.fromBuffer(input)).exportFile(),
    mutate: async (input) => {
      const agent = await DocxAgent.fromBuffer(input);
      const firstParagraphIndex = agent
        .getSnapshot()
        .root.body.findIndex((block) => block.kind === "paragraph");
      expect(firstParagraphIndex).toBeGreaterThanOrEqual(0);
      await agent.applyCommand({
        type: "docx:insert-text",
        payload: { at: { paragraph: firstParagraphIndex, offset: 0 }, text: "Opaque preservation " },
        source: "human",
      });
      return agent.exportFile();
    },
  },
  {
    format: "xlsx",
    fixturePath: fixturePath(requiredMatrixFixture("xlsx", { id: "xlsx.synthetic.single-sheet-numbers" })),
    relsPath: "xl/_rels/workbook.xml.rels",
    xmlPart: "customXml/officeAiOpaque-xlsx.xml",
    binPart: "xl/embeddings/officeAiOpaque-xlsx.bin",
    touchedPart: "xl/worksheets/sheet1.xml",
    roundtrip: async (input) => (await XlsxAgent.fromBuffer(input)).exportFile(),
    mutate: async (input) => {
      const agent = await XlsxAgent.fromBuffer(input);
      const sheet = agent.getSnapshot().root.sheets.find((candidate) => candidate.kind === "worksheet");
      if (!sheet) throw new Error("expected worksheet in xlsx probe fixture");
      await agent.applyCommand({
        type: "xlsx:set-cell-value",
        payload: { sheet: sheet.name, ref: "Z99", value: "Opaque preservation" },
        source: "human",
      });
      return agent.exportFile();
    },
  },
  {
    format: "pptx",
    fixturePath: fixturePath(requiredMatrixFixture("pptx", { id: "pptx.synthetic.multi-shape" })),
    relsPath: "ppt/slides/_rels/slide1.xml.rels",
    xmlPart: "customXml/officeAiOpaque-pptx.xml",
    binPart: "ppt/embeddings/officeAiOpaque-pptx.bin",
    touchedPart: "ppt/slides/slide1.xml",
    roundtrip: async (input) => (await PptxAgent.fromBuffer(input)).exportFile(),
    mutate: async (input) => {
      const agent = await PptxAgent.fromBuffer(input);
      const slide = agent.getSnapshot().root.slides[0];
      const textShape = slide.shapes.find((shape) => shape.kind === "text");
      if (!textShape) throw new Error("expected text shape in pptx probe fixture");
      await agent.applyCommand({
        type: "pptx:set-text",
        payload: { slideIndex: 0, shapeId: textShape.id, text: "Opaque preservation" },
        source: "human",
      });
      return agent.exportFile();
    },
  },
];

describe("OOXML opaque preservation contract", () => {
  for (const probe of PROBES) {
    it(`${probe.format}: no-op roundtrip preserves injected opaque parts and relationships`, async () => {
      const input = await augmentFixture(probe);
      const before = await ooxml.OoxmlContainer.load(input);
      const beforeHashes = partHashes(before);
      const beforeRels = relationshipSnapshot(before, probe.relsPath);

      const output = await probe.roundtrip(input);
      const after = await ooxml.OoxmlContainer.load(output);
      const afterHashes = partHashes(after);

      expect([...after.parts.keys()].sort()).toEqual([...before.parts.keys()].sort());
      assertOpaqueProbePreserved(probe, beforeHashes, afterHashes);
      expect(relationshipSnapshot(after, probe.relsPath)).toEqual(beforeRels);
    });

    it(`${probe.format}: known mutation does not drop injected opaque parts or relationships`, async () => {
      const input = await augmentFixture(probe);
      const before = await ooxml.OoxmlContainer.load(input);
      const beforeHashes = partHashes(before);
      const beforeRels = relationshipSnapshot(before, probe.relsPath);

      const output = await probe.mutate(input);
      const after = await ooxml.OoxmlContainer.load(output);
      const afterHashes = partHashes(after);

      expect([...after.parts.keys()].sort()).toEqual([...before.parts.keys()].sort());
      assertOpaqueProbePreserved(probe, beforeHashes, afterHashes);
      expect(relationshipSnapshot(after, probe.relsPath)).toEqual(beforeRels);
      expect(
        afterHashes[probe.touchedPart],
        `${probe.touchedPart} should be dirtied by the mutation`
      ).not.toBe(beforeHashes[probe.touchedPart]);
    });
  }

  it("defines snapshot-stable diagnostics for opaque preservation decisions", () => {
    expect([
      ooxml.ooxmlPreservationDiagnostic({
        format: "docx",
        code: "ooxml-opaque-part-preserved",
        partPath: "word/embeddings/legacy.bin",
      }),
      ooxml.ooxmlPreservationDiagnostic({
        format: "xlsx",
        code: "ooxml-opaque-preservation-risk",
        partPath: "xl/pivotCache/cacheDefinition1.xml",
        detail: "The owning command would rewrite a feature that is not in the typed model.",
      }),
      ooxml.ooxmlPreservationDiagnostic({
        format: "pptx",
        code: "ooxml-opaque-mutation-blocked",
        partPath: "ppt/embeddings/oleObject1.bin",
        detail: "Use a supported slide, media or chart command instead.",
      }),
    ]).toMatchInlineSnapshot(`
      [
        {
          "code": "ooxml-opaque-part-preserved",
          "format": "docx",
          "level": "info",
          "message": "DOCX opaque OOXML part word/embeddings/legacy.bin was preserved byte-for-byte.",
          "partPath": "word/embeddings/legacy.bin",
        },
        {
          "code": "ooxml-opaque-preservation-risk",
          "format": "xlsx",
          "level": "warning",
          "message": "XLSX opaque OOXML part xl/pivotCache/cacheDefinition1.xml may not be safe to rewrite. The owning command would rewrite a feature that is not in the typed model.",
          "partPath": "xl/pivotCache/cacheDefinition1.xml",
        },
        {
          "code": "ooxml-opaque-mutation-blocked",
          "format": "pptx",
          "level": "error",
          "message": "PPTX opaque OOXML part ppt/embeddings/oleObject1.bin cannot be safely mutated. Use a supported slide, media or chart command instead.",
          "partPath": "ppt/embeddings/oleObject1.bin",
        },
      ]
    `);
  });
});

async function augmentFixture(probe: PreservationProbe): Promise<Uint8Array> {
  const container = await ooxml.OoxmlContainer.load(await readFile(probe.fixturePath));
  container.writeText(
    probe.xmlPart,
    `<officeAiOpaque xmlns="urn:office-ai:test" format="${probe.format}">preserve me</officeAiOpaque>`
  );
  container.writeBytes(probe.binPart, TEXT_ENCODER.encode(`opaque-binary-${probe.format}`));
  appendRelationships(container, probe);
  appendContentTypeOverride(container, probe.xmlPart, "application/xml");
  appendContentTypeOverride(
    container,
    probe.binPart,
    "application/vnd.openxmlformats-officedocument.oleObject"
  );
  return new Uint8Array(await container.serialize());
}

function appendRelationships(container: ooxml.OoxmlContainer, probe: PreservationProbe): void {
  const relsXml = container.has(probe.relsPath)
    ? container.readText(probe.relsPath)
    : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  const xmlTarget = relativeTargetFor(probe, "xml");
  const binTarget = relativeTargetFor(probe, "bin");
  const insert =
    `<Relationship Id="${OPAQUE_XML_REL_ID}" Type="${CUSTOM_XML_REL_TYPE}" Target="${xmlTarget}"/>` +
    `<Relationship Id="${OPAQUE_BIN_REL_ID}" Type="${OLE_OBJECT_REL_TYPE}" Target="${binTarget}"/>`;
  container.writeText(probe.relsPath, relsXml.replace("</Relationships>", `${insert}</Relationships>`));
}

function appendContentTypeOverride(
  container: ooxml.OoxmlContainer,
  partPath: string,
  contentType: string
): void {
  const partName = `/${partPath}`;
  const contentTypes = container.readText(CONTENT_TYPES);
  if (contentTypes.includes(`PartName="${partName}"`)) return;
  const override = `<Override PartName="${partName}" ContentType="${contentType}"/>`;
  container.writeText(CONTENT_TYPES, contentTypes.replace("</Types>", `${override}</Types>`));
}

function relativeTargetFor(probe: PreservationProbe, kind: "xml" | "bin"): string {
  if (probe.format === "pptx" && kind === "xml") return "../../customXml/officeAiOpaque-pptx.xml";
  if (probe.format === "pptx" && kind === "bin") return "../embeddings/officeAiOpaque-pptx.bin";
  if (kind === "xml") return `../customXml/officeAiOpaque-${probe.format}.xml`;
  return `embeddings/officeAiOpaque-${probe.format}.bin`;
}

function partHashes(container: ooxml.OoxmlContainer): Record<string, string> {
  const out: Record<string, string> = {};
  for (const part of container.parts.keys()) {
    out[part] = sha256Hex(container.readBytes(part));
  }
  return out;
}

function relationshipSnapshot(container: ooxml.OoxmlContainer, relsPath: string): ReadonlyArray<string> {
  const rels = container.readText(relsPath);
  const matches = rels.match(/<Relationship [^>]+>/g) ?? [];
  return matches.filter((entry) => entry.includes(OPAQUE_XML_REL_ID) || entry.includes(OPAQUE_BIN_REL_ID));
}

function assertOpaqueProbePreserved(
  probe: PreservationProbe,
  beforeHashes: Record<string, string>,
  afterHashes: Record<string, string>
): void {
  for (const part of [probe.xmlPart, probe.binPart, probe.relsPath, CONTENT_TYPES]) {
    expect(afterHashes[part], `${probe.format} opaque preservation part ${part}`).toBe(beforeHashes[part]);
  }
}
