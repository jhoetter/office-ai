import { describe, expect, it } from "vitest";
import { normalizeDocumentDiff } from "./semantic-diff.js";
import type { DocumentDiff } from "../types/document.js";

describe("semantic diff normalization", () => {
  it("maps raw document diffs into a reviewable schema", () => {
    const diff: DocumentDiff = {
      format: "docx",
      fromRevision: 1,
      toRevision: 2,
      changes: [
        {
          kind: "node-updated",
          nodeId: "p1",
          path: ["body", 0],
          field: "text",
          summary: "Updated paragraph text.",
          meta: { before: "old", after: "new" },
        },
      ],
    };

    const semantic = normalizeDocumentDiff(diff, { operation: "docx:set-text" });

    expect(semantic).toMatchObject({
      schema: "office-ai/semantic-diff@1",
      format: "docx",
      summary: {
        text: "docx:set-text: 1 DOCX change; low review risk.",
        changeCount: 1,
        risk: "low",
      },
      fallback: false,
      truncated: false,
    });
    expect(semantic.anchors).toEqual([{ id: "p1", label: "DOCX body > 0", path: ["body", 0] }]);
    expect(semantic.changes[0]).toMatchObject({
      kind: "node-updated",
      field: "text",
      before: "old",
      after: "new",
    });
  });

  it("reports opaque revision changes as fallback diffs", () => {
    const semantic = normalizeDocumentDiff(
      { format: "pdf", fromRevision: 3, toRevision: 4, changes: [] },
      { operation: "pdf:opaque-write" }
    );

    expect(semantic.fallback).toBe(true);
    expect(semantic.summary).toMatchObject({
      text: "pdf:opaque-write: opaque PDF change; review raw diagnostics.",
      risk: "unknown",
    });
    expect(semantic.diagnostics).toEqual([
      expect.objectContaining({ code: "semantic-diff-fallback", level: "warning" }),
    ]);
  });

  it("truncates large diffs for MCP payloads", () => {
    const changes = Array.from({ length: 3 }, (_, index) => ({
      kind: "node-deleted" as const,
      nodeId: `row-${index}`,
      path: ["sheets", "Sheet1", index],
      summary: `Deleted row ${index}.`,
    }));

    const semantic = normalizeDocumentDiff(
      { format: "xlsx", fromRevision: 1, toRevision: 2, changes },
      { maxChanges: 2 }
    );

    expect(semantic.truncated).toBe(true);
    expect(semantic.changes).toHaveLength(2);
    expect(semantic.summary).toMatchObject({ changeCount: 3, risk: "high" });
    expect(semantic.diagnostics.map((diagnostic) => diagnostic.code)).toContain("semantic-diff-truncated");
  });
});
