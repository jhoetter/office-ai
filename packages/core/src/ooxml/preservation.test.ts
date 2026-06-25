import { describe, expect, it } from "vitest";
import { OOXML_PRESERVATION_CONTRACT, ooxmlPreservationDiagnostic } from "./preservation.js";

describe("OOXML preservation contract", () => {
  it("covers every OOXML product format with parsed, opaque and diagnostic policy", () => {
    expect(Object.keys(OOXML_PRESERVATION_CONTRACT).sort()).toEqual(["docx", "pptx", "xlsx"]);
    for (const contract of Object.values(OOXML_PRESERVATION_CONTRACT)) {
      expect(contract.parsedPartPatterns.length).toBeGreaterThan(0);
      expect(contract.opaquePartPatterns.length).toBeGreaterThan(0);
      expect(contract.relationshipPolicy).toContain(".rels");
      expect(contract.mutationPolicy).toContain("diagnostic");
      expect(contract.diagnosticCodes).toEqual([
        "ooxml-opaque-part-preserved",
        "ooxml-opaque-preservation-risk",
        "ooxml-opaque-mutation-blocked",
      ]);
    }
  });

  it("maps preservation decisions to existing diagnostic levels", () => {
    expect([
      ooxmlPreservationDiagnostic({
        format: "docx",
        code: "ooxml-opaque-part-preserved",
        partPath: "word/embeddings/legacy.bin",
      }).level,
      ooxmlPreservationDiagnostic({
        format: "xlsx",
        code: "ooxml-opaque-preservation-risk",
        partPath: "xl/pivotCache/cacheDefinition1.xml",
      }).level,
      ooxmlPreservationDiagnostic({
        format: "pptx",
        code: "ooxml-opaque-mutation-blocked",
        partPath: "ppt/embeddings/oleObject1.bin",
      }).level,
    ]).toEqual(["info", "warning", "error"]);
  });
});
