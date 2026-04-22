import { describe, expect, it } from "vitest";
import {
  fillSpecToOpaque,
  normaliseFillSpec,
  readFillSpec,
  spliceFillIntoSpPr,
  spliceSlideBackground,
  type FillSpec,
} from "./fill.js";
import type { OpaqueXml } from "./types.js";

function roundTrip(spec: FillSpec): FillSpec | null {
  const opaque = fillSpecToOpaque(spec);
  return readFillSpec([opaque]);
}

describe("FillSpec / opaque round-trip", () => {
  it("solid color round-trips with 6-char hex", () => {
    const out = roundTrip({ type: "solid", color: "ff0080" });
    expect(out).toEqual({ type: "solid", color: "FF0080", alpha: undefined });
  });

  it("solid color preserves alpha < 1", () => {
    const out = roundTrip({ type: "solid", color: "112233", alpha: 0.5 });
    expect(out).toMatchObject({ type: "solid", color: "112233".toUpperCase(), alpha: 0.5 });
  });

  it("solid color drops alpha === 1 (treats as opaque)", () => {
    const out = roundTrip({ type: "solid", color: "112233", alpha: 1 });
    expect(out).toMatchObject({ type: "solid", color: "112233", alpha: undefined });
  });

  it("none round-trips", () => {
    expect(roundTrip({ type: "none" })).toEqual({ type: "none" });
  });

  it("linear gradient round-trips with stops sorted and angle preserved", () => {
    const spec: FillSpec = {
      type: "gradient",
      kind: "linear",
      angleDeg: 90,
      stops: [
        { pos: 1, color: "ffffff" },
        { pos: 0, color: "000000" },
      ],
    };
    const out = roundTrip(spec);
    expect(out).toMatchObject({
      type: "gradient",
      kind: "linear",
      angleDeg: 90,
      stops: [
        { pos: 0, color: "000000" },
        { pos: 1, color: "FFFFFF" },
      ],
    });
  });

  it("radial gradient round-trips kind", () => {
    const spec: FillSpec = {
      type: "gradient",
      kind: "radial",
      angleDeg: 0,
      stops: [
        { pos: 0, color: "ff0000" },
        { pos: 1, color: "0000ff" },
      ],
    };
    const out = roundTrip(spec);
    expect(out?.type).toBe("gradient");
    if (out?.type === "gradient") expect(out.kind).toBe("radial");
  });

  it("pattern round-trips fg/bg/preset", () => {
    const out = roundTrip({ type: "pattern", preset: "diagCross", fgColor: "112233", bgColor: "445566" });
    expect(out).toMatchObject({
      type: "pattern",
      preset: "diagCross",
      fgColor: "112233",
      bgColor: "445566",
    });
  });

  it("picture round-trips embed rel id and tile flag", () => {
    const out = roundTrip({ type: "picture", embedRelId: "rId7", tile: true });
    expect(out).toMatchObject({ type: "picture", embedRelId: "rId7", tile: true });
  });
});

describe("normaliseFillSpec", () => {
  it("rejects malformed hex", () => {
    expect(() => normaliseFillSpec({ type: "solid", color: "xyz" })).toThrow();
  });
  it("rejects gradient with <2 stops", () => {
    expect(() =>
      normaliseFillSpec({
        type: "gradient",
        kind: "linear",
        angleDeg: 0,
        stops: [{ pos: 0, color: "ff0000" }],
      })
    ).toThrow();
  });
  it("normalises angle into [0,360)", () => {
    const out = normaliseFillSpec({
      type: "gradient",
      kind: "linear",
      angleDeg: -45,
      stops: [
        { pos: 0, color: "000000" },
        { pos: 1, color: "ffffff" },
      ],
    }) as Extract<FillSpec, { type: "gradient" }>;
    expect(out.angleDeg).toBe(315);
  });
});

describe("spliceFillIntoSpPr", () => {
  const prstGeom: OpaqueXml = { tag: "a:prstGeom", attrs: {}, rawAttrs: {}, subtree: [] };
  const oldFill: OpaqueXml = { tag: "a:solidFill", attrs: {}, rawAttrs: {}, subtree: [] };
  const oldNoFill: OpaqueXml = { tag: "a:noFill", attrs: {}, rawAttrs: {}, subtree: [] };
  const oldGrad: OpaqueXml = { tag: "a:gradFill", attrs: {}, rawAttrs: {}, subtree: [] };
  const oldPatt: OpaqueXml = { tag: "a:pattFill", attrs: {}, rawAttrs: {}, subtree: [] };
  const oldBlip: OpaqueXml = { tag: "a:blipFill", attrs: {}, rawAttrs: {}, subtree: [] };
  const ln: OpaqueXml = { tag: "a:ln", attrs: {}, rawAttrs: {}, subtree: [] };

  it("strips every existing fill kind before inserting the replacement", () => {
    const tail = [prstGeom, oldFill, oldNoFill, oldGrad, oldPatt, oldBlip, ln];
    const next = spliceFillIntoSpPr(tail, { type: "solid", color: "ff0000" });
    expect(next.filter((c) => c.tag === "a:solidFill").length).toBe(1);
    expect(next.some((c) => c.tag === "a:noFill")).toBe(false);
    expect(next.some((c) => c.tag === "a:gradFill")).toBe(false);
    expect(next.some((c) => c.tag === "a:pattFill")).toBe(false);
    expect(next.some((c) => c.tag === "a:blipFill")).toBe(false);
    expect(next.some((c) => c.tag === "a:ln")).toBe(true);
    expect(next.some((c) => c.tag === "a:prstGeom")).toBe(true);
  });

  it("inserts the new fill immediately after a:prstGeom", () => {
    const tail = [prstGeom, ln];
    const next = spliceFillIntoSpPr(tail, { type: "solid", color: "ff0000" });
    expect(next[0].tag).toBe("a:prstGeom");
    expect(next[1].tag).toBe("a:solidFill");
    expect(next[2].tag).toBe("a:ln");
  });

  it("clears the fill when spec is null", () => {
    const tail = [prstGeom, oldFill, ln];
    const next = spliceFillIntoSpPr(tail, null);
    expect(next.some((c) => c.tag === "a:solidFill")).toBe(false);
  });
});

describe("spliceSlideBackground", () => {
  it("inserts <p:bg> with the right fill node", () => {
    const next = spliceSlideBackground([], { type: "solid", color: "112233" });
    expect(next.length).toBe(1);
    expect(next[0].tag).toBe("p:bg");
    // Walk into <p:bgPr> > <a:solidFill> via the opaque subtree.
    const bgPr = (next[0].subtree[0] as Record<string, unknown>)["p:bgPr"] as unknown[];
    expect(Array.isArray(bgPr)).toBe(true);
    const fill = (bgPr[0] as Record<string, unknown>)["a:solidFill"];
    expect(Array.isArray(fill)).toBe(true);
  });

  it("removes <p:bg> when spec is null", () => {
    const existing: OpaqueXml = { tag: "p:bg", attrs: {}, rawAttrs: {}, subtree: [] };
    const other: OpaqueXml = { tag: "p:other", attrs: {}, rawAttrs: {}, subtree: [] };
    const next = spliceSlideBackground([existing, other], null);
    expect(next.some((c) => c.tag === "p:bg")).toBe(false);
    expect(next.some((c) => c.tag === "p:other")).toBe(true);
  });
});
