import type { XlsxClipboardSnapshot } from "@officeai/xlsx";
import { describe, expect, it } from "vitest";
import {
  EMBED_MIME,
  EMBED_VERSION,
  makeEnvelope,
  parseEnvelope,
  serializeEnvelope,
} from "./envelope";

const SAMPLE_SNAPSHOT: XlsxClipboardSnapshot = {
  origin: { sheet: "Sheet1", range: "A1:B2" },
  width: 2,
  height: 2,
  cells: [
    [
      { value: "h1" },
      { value: "h2" },
    ],
    [
      { value: 1 },
      { value: 2 },
    ],
  ],
  merges: [],
};

describe("OfficeAI embed envelope", () => {
  it("exposes a stable MIME type and version", () => {
    expect(EMBED_MIME).toBe("application/x-officeai-embed+json");
    expect(EMBED_VERSION).toBe(1);
  });

  it("round-trips an XLSX-range payload through serialize/parse", () => {
    const env = makeEnvelope("xlsx", {
      kind: "xlsx-range",
      snapshot: SAMPLE_SNAPSHOT,
      originLabel: "Sheet1!A1:B2",
    });
    const wire = serializeEnvelope(env);
    const back = parseEnvelope(wire);
    expect(back).not.toBeNull();
    expect(back?.source).toBe("xlsx");
    expect(back?.payload.kind).toBe("xlsx-range");
    if (back?.payload.kind === "xlsx-range") {
      expect(back.payload.snapshot.cells[0][0]?.value).toBe("h1");
      expect(back.payload.snapshot.merges).toEqual([]);
      expect(back.payload.originLabel).toBe("Sheet1!A1:B2");
    }
  });

  it("round-trips a chart-image payload", () => {
    const env = makeEnvelope("xlsx", {
      kind: "xlsx-chart-image",
      png: "iVBORw0KGgo=",
      width: 480,
      height: 280,
      chartKind: "column",
      title: "Revenue",
    });
    const wire = serializeEnvelope(env);
    const back = parseEnvelope(wire);
    expect(back?.payload.kind).toBe("xlsx-chart-image");
    if (back?.payload.kind === "xlsx-chart-image") {
      expect(back.payload.width).toBe(480);
      expect(back.payload.chartKind).toBe("column");
    }
  });

  it("round-trips a docx-table payload", () => {
    const env = makeEnvelope("docx", {
      kind: "docx-table",
      cells: [
        ["Name", "Score"],
        ["Ada", "99"],
        ["Linus", "42"],
      ],
      originLabel: "Document table",
    });
    const wire = serializeEnvelope(env);
    const back = parseEnvelope(wire);
    expect(back?.source).toBe("docx");
    expect(back?.payload.kind).toBe("docx-table");
    if (back?.payload.kind === "docx-table") {
      expect(back.payload.cells).toHaveLength(3);
      expect(back.payload.cells[1]?.[1]).toBe("99");
      expect(back.payload.originLabel).toBe("Document table");
    }
  });

  it("returns null for non-JSON, wrong-shape, and falsy inputs", () => {
    expect(parseEnvelope(null)).toBeNull();
    expect(parseEnvelope(undefined)).toBeNull();
    expect(parseEnvelope("")).toBeNull();
    expect(parseEnvelope("{not json")).toBeNull();
    expect(parseEnvelope(JSON.stringify({ type: "other" }))).toBeNull();
    expect(
      parseEnvelope(
        JSON.stringify({ type: "officeai/embed", version: 1, source: "xlsx", payload: { kind: "unknown" } })
      )
    ).toBeNull();
    expect(
      parseEnvelope(JSON.stringify({ type: "officeai/embed", version: 1, source: "elsewhere", payload: { kind: "xlsx-range" } }))
    ).toBeNull();
  });

  it("ignores unknown extra fields rather than rejecting them", () => {
    const wire = JSON.stringify({
      type: "officeai/embed",
      version: 1,
      source: "xlsx",
      createdAt: new Date().toISOString(),
      futureField: "ignored",
      payload: {
        kind: "xlsx-range",
        snapshot: SAMPLE_SNAPSHOT,
        originLabel: "x",
        anotherFutureField: 42,
      },
    });
    expect(parseEnvelope(wire)).not.toBeNull();
  });
});
