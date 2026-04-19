import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { ooxml } from "@officeai/core";
import { PptxAgent } from "../agent/agent.js";
import type { ConnectorShape, TextShape } from "../model/types.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

async function loadAgent(name: string): Promise<PptxAgent> {
  const buf = await readFile(join(FIXTURES_DIR.pathname, name));
  return PptxAgent.fromBuffer(buf);
}

/**
 * Add two rectangles + a connector wired between them so the rest of the
 * tests can drive lifecycle commands without re-doing the boilerplate.
 */
async function setupTwoShapesAndConnector(): Promise<{
  agent: PptxAgent;
  rect1: TextShape;
  rect2: TextShape;
  connector: ConnectorShape;
}> {
  const agent = await loadAgent("01-blank.pptx");
  await agent.applyCommand({
    type: "pptx:add-shape",
    payload: {
      slideIndex: 0,
      preset: "rect",
      x: 1_000_000,
      y: 1_000_000,
      width: 2_000_000,
      height: 1_000_000,
    },
    source: "human",
  });
  await agent.applyCommand({
    type: "pptx:add-shape",
    payload: {
      slideIndex: 0,
      preset: "rect",
      x: 6_000_000,
      y: 4_000_000,
      width: 2_000_000,
      height: 1_000_000,
    },
    source: "human",
  });
  const slide0 = agent.getSnapshot().root.slides[0];
  const rect1 = slide0.shapes[slide0.shapes.length - 2] as TextShape;
  const rect2 = slide0.shapes[slide0.shapes.length - 1] as TextShape;
  await agent.applyCommand({
    type: "pptx:add-connector",
    payload: {
      slideIndex: 0,
      connectorType: "straight",
      start: { kind: "anchored", targetCNvPrId: rect1.cNvPrId, side: "e" },
      end: { kind: "anchored", targetCNvPrId: rect2.cNvPrId, side: "w" },
    },
    source: "human",
  });
  const slide1 = agent.getSnapshot().root.slides[0];
  const connector = slide1.shapes[slide1.shapes.length - 1] as ConnectorShape;
  return { agent, rect1, rect2, connector };
}

describe("pptx:add-connector", () => {
  it("creates a typed ConnectorShape with anchored endpoints whose box matches the snapped target points", async () => {
    const { agent, rect1, rect2, connector } = await setupTwoShapesAndConnector();
    expect(connector.kind).toBe("connector");
    expect(connector.connectorType).toBe("straight");
    expect(connector.start).toMatchObject({ kind: "anchored", targetCNvPrId: rect1.cNvPrId, side: "e" });
    expect(connector.end).toMatchObject({ kind: "anchored", targetCNvPrId: rect2.cNvPrId, side: "w" });

    // rect1 east anchor → x=3_000_000, y=1_500_000
    // rect2 west anchor → x=6_000_000, y=4_500_000
    expect(connector.position?.xEmu).toBe(3_000_000);
    expect(connector.position?.yEmu).toBe(1_500_000);
    expect(connector.size?.cxEmu).toBe(3_000_000);
    expect(connector.size?.cyEmu).toBe(3_000_000);

    // Default head arrow + slate-700 stroke
    expect(connector.headEnd).toBe("arrow");
    expect(connector.stroke?.color).toBe("374151");

    // Round-trip: serialised slide should embed <p:cxnSp> with the resolved xfrm.
    const out = await agent.exportFile();
    const c = await ooxml.OoxmlContainer.load(out);
    const slide = agent.getSnapshot().root.slides[0];
    const xml = c.readText(slide.partPath);
    expect(xml).toContain("<p:cxnSp>");
    expect(xml).toContain('prst="line"');
    expect(xml).toContain('id="' + rect1.cNvPrId + '"');
  });

  it("rejects an anchored endpoint that targets an unknown cNvPrId", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const m = await agent.applyCommand({
      type: "pptx:add-connector",
      payload: {
        slideIndex: 0,
        connectorType: "straight",
        start: { kind: "anchored", targetCNvPrId: 99_999, side: "n" },
        end: { kind: "free", xEmu: 0, yEmu: 0 },
      },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-target");
  });
});

describe("pptx:set-connector-endpoint", () => {
  it("re-anchors one endpoint and rebuilds the bounding box from the new target", async () => {
    const { agent, rect1, connector } = await setupTwoShapesAndConnector();
    // Re-anchor end → north of rect1 (so both endpoints reference rect1).
    await agent.applyCommand({
      type: "pptx:set-connector-endpoint",
      payload: {
        slideIndex: 0,
        shapeId: connector.id,
        which: "end",
        endpoint: { kind: "anchored", targetCNvPrId: rect1.cNvPrId, side: "n" },
      },
      source: "human",
    });
    const slide = agent.getSnapshot().root.slides[0];
    const c = slide.shapes.find((s) => s.id === connector.id) as ConnectorShape;
    expect(c.end).toMatchObject({ kind: "anchored", side: "n" });
    // start = east anchor of rect1 (3_000_000, 1_500_000); end = north (2_000_000, 1_000_000).
    expect(c.position?.xEmu).toBe(2_000_000);
    expect(c.position?.yEmu).toBe(1_000_000);
    expect(c.size?.cxEmu).toBe(1_000_000);
    expect(c.size?.cyEmu).toBe(500_000);
  });
});

describe("pptx:set-connector-style", () => {
  it("updates type, stroke colour/width and head/tail without losing the endpoints", async () => {
    const { agent, connector } = await setupTwoShapesAndConnector();
    await agent.applyCommand({
      type: "pptx:set-connector-style",
      payload: {
        slideIndex: 0,
        shapeId: connector.id,
        connectorType: "elbow",
        strokeColor: "#3366ff",
        strokeWidthEmu: 19_050,
        headEnd: "triangle",
        tailEnd: "oval",
      },
      source: "human",
    });
    const slide = agent.getSnapshot().root.slides[0];
    const c = slide.shapes.find((s) => s.id === connector.id) as ConnectorShape;
    expect(c.connectorType).toBe("elbow");
    expect(c.stroke?.color).toBe("3366FF");
    expect(c.stroke?.widthEmu).toBe(19_050);
    expect(c.headEnd).toBe("triangle");
    expect(c.tailEnd).toBe("oval");
    // Endpoints survive untouched.
    expect(c.start).toEqual(connector.start);
    expect(c.end).toEqual(connector.end);
  });

  it("rejects when no fields are supplied", async () => {
    const { agent, connector } = await setupTwoShapesAndConnector();
    const m = await agent.applyCommand({
      type: "pptx:set-connector-style",
      payload: { slideIndex: 0, shapeId: connector.id },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
  });
});

describe("connector reflow on shape edits", () => {
  it("moves the connector's bounding box when an anchored target shape is moved", async () => {
    const { agent, rect2, connector } = await setupTwoShapesAndConnector();
    await agent.applyCommand({
      type: "pptx:set-position",
      payload: { slideIndex: 0, shapeId: rect2.id, x: 7_500_000, y: 5_000_000 },
      source: "human",
    });
    const slide = agent.getSnapshot().root.slides[0];
    const c = slide.shapes.find((s) => s.id === connector.id) as ConnectorShape;
    // rect2 west anchor moved to (7_500_000, 5_500_000).
    expect(c.position?.xEmu).toBe(3_000_000);
    expect(c.size?.cxEmu).toBe(4_500_000);
    expect(c.size?.cyEmu).toBe(4_000_000);
  });

  it("resizes the connector's bounding box when an anchored target is resized", async () => {
    const { agent, rect2, connector } = await setupTwoShapesAndConnector();
    await agent.applyCommand({
      type: "pptx:set-size",
      payload: { slideIndex: 0, shapeId: rect2.id, width: 4_000_000, height: 2_000_000 },
      source: "human",
    });
    const slide = agent.getSnapshot().root.slides[0];
    const c = slide.shapes.find((s) => s.id === connector.id) as ConnectorShape;
    // rect2 west anchor: x=6_000_000, y = 4_000_000 + 2_000_000/2 = 5_000_000.
    expect(c.position?.yEmu).toBe(1_500_000);
    expect(c.size?.cyEmu).toBe(3_500_000);
  });

  it("detaches anchored endpoints to their last resolved point when the target is deleted", async () => {
    const { agent, rect2, connector } = await setupTwoShapesAndConnector();
    await agent.applyCommand({
      type: "pptx:delete-shape",
      payload: { slideIndex: 0, shapeId: rect2.id },
      source: "human",
    });
    const slide = agent.getSnapshot().root.slides[0];
    const c = slide.shapes.find((s) => s.id === connector.id) as ConnectorShape;
    expect(c.end.kind).toBe("free");
    if (c.end.kind === "free") {
      expect(c.end.xEmu).toBe(6_000_000);
      expect(c.end.yEmu).toBe(4_500_000);
    }
    // start endpoint untouched.
    expect(c.start.kind).toBe("anchored");
  });
});

describe("pptx:set-connector-endpoint quarter-point t", () => {
  it("persists a clamped + 3-decimal-rounded t when supplied for an edge anchor", async () => {
    const { agent, rect2, connector } = await setupTwoShapesAndConnector();
    await agent.applyCommand({
      type: "pptx:set-connector-endpoint",
      payload: {
        slideIndex: 0,
        shapeId: connector.id,
        which: "end",
        endpoint: {
          kind: "anchored",
          targetCNvPrId: rect2.cNvPrId,
          side: "n",
          t: 0.7777,
        },
      },
      source: "human",
    });
    const slide = agent.getSnapshot().root.slides[0];
    const c = slide.shapes.find((s) => s.id === connector.id) as ConnectorShape;
    expect(c.end.kind).toBe("anchored");
    if (c.end.kind === "anchored") {
      expect(c.end.t).toBe(0.778);
    }
  });

  it("clamps an out-of-range t into [0, 1]", async () => {
    const { agent, rect2, connector } = await setupTwoShapesAndConnector();
    await agent.applyCommand({
      type: "pptx:set-connector-endpoint",
      payload: {
        slideIndex: 0,
        shapeId: connector.id,
        which: "end",
        endpoint: { kind: "anchored", targetCNvPrId: rect2.cNvPrId, side: "n", t: 1.5 },
      },
      source: "human",
    });
    const slide = agent.getSnapshot().root.slides[0];
    const c = slide.shapes.find((s) => s.id === connector.id) as ConnectorShape;
    if (c.end.kind === "anchored") expect(c.end.t).toBe(1);
  });

  it("drops t for the center anchor (it has no edge to interpolate along)", async () => {
    const { agent, rect2, connector } = await setupTwoShapesAndConnector();
    await agent.applyCommand({
      type: "pptx:set-connector-endpoint",
      payload: {
        slideIndex: 0,
        shapeId: connector.id,
        which: "end",
        endpoint: { kind: "anchored", targetCNvPrId: rect2.cNvPrId, side: "center", t: 0.4 },
      },
      source: "human",
    });
    const slide = agent.getSnapshot().root.slides[0];
    const c = slide.shapes.find((s) => s.id === connector.id) as ConnectorShape;
    if (c.end.kind === "anchored") expect(c.end.t).toBeUndefined();
  });
});

describe("pptx:set-connector-waypoint", () => {
  it("stores a finite waypoint at the requested segment index", async () => {
    const { agent, connector } = await setupTwoShapesAndConnector();
    await agent.applyCommand({
      type: "pptx:set-connector-waypoint",
      payload: { slideIndex: 0, shapeId: connector.id, segmentIndex: 0, valueEmu: 250_000 },
      source: "human",
    });
    const slide = agent.getSnapshot().root.slides[0];
    const c = slide.shapes.find((s) => s.id === connector.id) as ConnectorShape;
    expect(c.waypoints).toEqual([250_000]);
  });

  it("pads earlier slots with 0 when writing past the current end", async () => {
    const { agent, connector } = await setupTwoShapesAndConnector();
    await agent.applyCommand({
      type: "pptx:set-connector-waypoint",
      payload: { slideIndex: 0, shapeId: connector.id, segmentIndex: 2, valueEmu: 500_000 },
      source: "human",
    });
    const slide = agent.getSnapshot().root.slides[0];
    const c = slide.shapes.find((s) => s.id === connector.id) as ConnectorShape;
    expect(c.waypoints).toEqual([0, 0, 500_000]);
  });

  it("clears a waypoint by passing valueEmu: null", async () => {
    const { agent, connector } = await setupTwoShapesAndConnector();
    await agent.applyCommand({
      type: "pptx:set-connector-waypoint",
      payload: { slideIndex: 0, shapeId: connector.id, segmentIndex: 0, valueEmu: 250_000 },
      source: "human",
    });
    await agent.applyCommand({
      type: "pptx:set-connector-waypoint",
      payload: { slideIndex: 0, shapeId: connector.id, segmentIndex: 0, valueEmu: null },
      source: "human",
    });
    const slide = agent.getSnapshot().root.slides[0];
    const c = slide.shapes.find((s) => s.id === connector.id) as ConnectorShape;
    expect(c.waypoints).toBeUndefined();
  });

  it("rejects non-integer segmentIndex", async () => {
    const { agent, connector } = await setupTwoShapesAndConnector();
    const m = await agent.applyCommand({
      type: "pptx:set-connector-waypoint",
      payload: { slideIndex: 0, shapeId: connector.id, segmentIndex: 1.5, valueEmu: 100 },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
  });

  it("rejects when the target shape is not a connector", async () => {
    const { agent, rect1 } = await setupTwoShapesAndConnector();
    const m = await agent.applyCommand({
      type: "pptx:set-connector-waypoint",
      payload: { slideIndex: 0, shapeId: rect1.id, segmentIndex: 0, valueEmu: 100 },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-target");
  });
});

describe("pptx:set-connector-style strokeDash", () => {
  it("persists a dashed stroke style", async () => {
    const { agent, connector } = await setupTwoShapesAndConnector();
    await agent.applyCommand({
      type: "pptx:set-connector-style",
      payload: { slideIndex: 0, shapeId: connector.id, strokeDash: "dashed" },
      source: "human",
    });
    const slide = agent.getSnapshot().root.slides[0];
    const c = slide.shapes.find((s) => s.id === connector.id) as ConnectorShape;
    expect(c.stroke?.dash).toBe("dashed");
  });

  it("emits <a:prstDash> in OOXML and survives a round-trip", async () => {
    const { agent, connector } = await setupTwoShapesAndConnector();
    await agent.applyCommand({
      type: "pptx:set-connector-style",
      payload: { slideIndex: 0, shapeId: connector.id, strokeDash: "dotted" },
      source: "human",
    });
    const buf = await agent.exportFile();
    const c = await ooxml.OoxmlContainer.load(buf);
    const slide = agent.getSnapshot().root.slides[0];
    const xml = c.readText(slide.partPath);
    expect(xml).toContain("<a:prstDash");
    const reloaded = await PptxAgent.fromBuffer(buf);
    const cxn = reloaded
      .getSnapshot()
      .root.slides[0].shapes.find((s) => s.kind === "connector") as ConnectorShape | undefined;
    expect(cxn?.stroke?.dash).toBe("dotted");
  });

  it("round-trips new longDash / dashDot presets through OOXML", async () => {
    for (const dash of ["longDash", "dashDot"] as const) {
      const { agent, connector } = await setupTwoShapesAndConnector();
      await agent.applyCommand({
        type: "pptx:set-connector-style",
        payload: { slideIndex: 0, shapeId: connector.id, strokeDash: dash },
        source: "human",
      });
      const buf = await agent.exportFile();
      const reloaded = await PptxAgent.fromBuffer(buf);
      const cxn = reloaded
        .getSnapshot()
        .root.slides[0].shapes.find((s) => s.kind === "connector") as ConnectorShape | undefined;
      expect(cxn?.stroke?.dash).toBe(dash);
    }
  });

  it("round-trips headEnd / tailEnd shapes", async () => {
    const { agent, connector } = await setupTwoShapesAndConnector();
    await agent.applyCommand({
      type: "pptx:set-connector-style",
      payload: { slideIndex: 0, shapeId: connector.id, headEnd: "triangle", tailEnd: "oval" },
      source: "human",
    });
    const buf = await agent.exportFile();
    const reloaded = await PptxAgent.fromBuffer(buf);
    const cxn = reloaded
      .getSnapshot()
      .root.slides[0].shapes.find((s) => s.kind === "connector") as ConnectorShape | undefined;
    expect(cxn?.headEnd).toBe("triangle");
    expect(cxn?.tailEnd).toBe("oval");
  });

  it("rejects an unknown dash style", async () => {
    const { agent, connector } = await setupTwoShapesAndConnector();
    const m = await agent.applyCommand({
      type: "pptx:set-connector-style",
      payload: {
        slideIndex: 0,
        shapeId: connector.id,
        strokeDash: "wiggly" as unknown as "solid",
      },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
  });
});

describe("pptx:reroute-connector", () => {
  it("clears existing waypoints", async () => {
    const { agent, connector } = await setupTwoShapesAndConnector();
    await agent.applyCommand({
      type: "pptx:set-connector-waypoint",
      payload: { slideIndex: 0, shapeId: connector.id, segmentIndex: 0, valueEmu: 250_000 },
      source: "human",
    });
    const before = agent
      .getSnapshot()
      .root.slides[0].shapes.find((s) => s.id === connector.id) as ConnectorShape;
    expect(before.waypoints).toEqual([250_000]);

    await agent.applyCommand({
      type: "pptx:reroute-connector",
      payload: { slideIndex: 0, shapeId: connector.id },
      source: "human",
    });
    const after = agent
      .getSnapshot()
      .root.slides[0].shapes.find((s) => s.id === connector.id) as ConnectorShape;
    expect(after.waypoints).toBeUndefined();
  });

  it("is idempotent on a connector with no waypoints", async () => {
    const { agent, connector } = await setupTwoShapesAndConnector();
    const m = await agent.applyCommand({
      type: "pptx:reroute-connector",
      payload: { slideIndex: 0, shapeId: connector.id },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const c = agent
      .getSnapshot()
      .root.slides[0].shapes.find((s) => s.id === connector.id) as ConnectorShape;
    expect(c.waypoints).toBeUndefined();
  });

  it("rejects when the target shape is not a connector", async () => {
    const { agent, rect1 } = await setupTwoShapesAndConnector();
    const m = await agent.applyCommand({
      type: "pptx:reroute-connector",
      payload: { slideIndex: 0, shapeId: rect1.id },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-target");
  });
});

describe("pptx:swap-connector-direction", () => {
  it("swaps start and end endpoints", async () => {
    const { agent, rect1, rect2, connector } = await setupTwoShapesAndConnector();
    await agent.applyCommand({
      type: "pptx:swap-connector-direction",
      payload: { slideIndex: 0, shapeId: connector.id },
      source: "human",
    });
    const c = agent
      .getSnapshot()
      .root.slides[0].shapes.find((s) => s.id === connector.id) as ConnectorShape;
    expect(c.start).toMatchObject({ kind: "anchored", targetCNvPrId: rect2.cNvPrId, side: "w" });
    expect(c.end).toMatchObject({ kind: "anchored", targetCNvPrId: rect1.cNvPrId, side: "e" });
  });

  it("swaps headEnd and tailEnd shapes", async () => {
    const { agent, connector } = await setupTwoShapesAndConnector();
    await agent.applyCommand({
      type: "pptx:set-connector-style",
      payload: { slideIndex: 0, shapeId: connector.id, headEnd: "triangle", tailEnd: "oval" },
      source: "human",
    });
    await agent.applyCommand({
      type: "pptx:swap-connector-direction",
      payload: { slideIndex: 0, shapeId: connector.id },
      source: "human",
    });
    const c = agent
      .getSnapshot()
      .root.slides[0].shapes.find((s) => s.id === connector.id) as ConnectorShape;
    expect(c.headEnd).toBe("oval");
    expect(c.tailEnd).toBe("triangle");
  });

  it("twice-swapped yields the original ordering", async () => {
    const { agent, connector } = await setupTwoShapesAndConnector();
    const before = agent
      .getSnapshot()
      .root.slides[0].shapes.find((s) => s.id === connector.id) as ConnectorShape;
    await agent.applyCommand({
      type: "pptx:swap-connector-direction",
      payload: { slideIndex: 0, shapeId: connector.id },
      source: "human",
    });
    await agent.applyCommand({
      type: "pptx:swap-connector-direction",
      payload: { slideIndex: 0, shapeId: connector.id },
      source: "human",
    });
    const after = agent
      .getSnapshot()
      .root.slides[0].shapes.find((s) => s.id === connector.id) as ConnectorShape;
    expect(after.start).toEqual(before.start);
    expect(after.end).toEqual(before.end);
  });

  it("preserves the bounding box (order-independent)", async () => {
    const { agent, connector } = await setupTwoShapesAndConnector();
    const before = agent
      .getSnapshot()
      .root.slides[0].shapes.find((s) => s.id === connector.id) as ConnectorShape;
    await agent.applyCommand({
      type: "pptx:swap-connector-direction",
      payload: { slideIndex: 0, shapeId: connector.id },
      source: "human",
    });
    const after = agent
      .getSnapshot()
      .root.slides[0].shapes.find((s) => s.id === connector.id) as ConnectorShape;
    expect(after.position).toEqual(before.position);
    expect(after.size).toEqual(before.size);
  });

  it("rejects when the target shape is not a connector", async () => {
    const { agent, rect1 } = await setupTwoShapesAndConnector();
    const m = await agent.applyCommand({
      type: "pptx:swap-connector-direction",
      payload: { slideIndex: 0, shapeId: rect1.id },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-target");
  });
});

describe("connector edge cases", () => {
  it("anchors a connector to a shape inside a group", async () => {
    const agent = await loadAgent("01-blank.pptx");
    await agent.applyCommand({
      type: "pptx:add-shape",
      payload: { slideIndex: 0, preset: "rect", x: 1_000_000, y: 1_000_000, width: 2_000_000, height: 1_000_000 },
      source: "human",
    });
    await agent.applyCommand({
      type: "pptx:add-shape",
      payload: { slideIndex: 0, preset: "rect", x: 4_000_000, y: 1_000_000, width: 2_000_000, height: 1_000_000 },
      source: "human",
    });
    const before = agent.getSnapshot().root.slides[0];
    const a = before.shapes[before.shapes.length - 2] as TextShape;
    const b = before.shapes[before.shapes.length - 1] as TextShape;
    await agent.applyCommand({
      type: "pptx:group-shapes",
      payload: { slideIndex: 0, shapeIds: [a.id, b.id] },
      source: "human",
    });
    const grouped = agent.getSnapshot().root.slides[0];
    const groupShape = grouped.shapes[grouped.shapes.length - 1];
    expect(groupShape.kind).toBe("group");
    if (groupShape.kind !== "group") return;
    // Anchor the connector to a shape that's now nested inside the group.
    await agent.applyCommand({
      type: "pptx:add-connector",
      payload: {
        slideIndex: 0,
        connectorType: "straight",
        start: { kind: "anchored", targetCNvPrId: a.cNvPrId, side: "e" },
        end: { kind: "free", xEmu: 7_000_000, yEmu: 1_500_000 },
      },
      source: "human",
    });
    const slide = agent.getSnapshot().root.slides[0];
    const c = slide.shapes.find((s) => s.kind === "connector") as ConnectorShape | undefined;
    expect(c).toBeDefined();
    if (c?.start.kind === "anchored") expect(c.start.targetCNvPrId).toBe(a.cNvPrId);
  });

  it("supports off-slide free endpoints (negative + beyond-slide coords)", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const slideSize = agent.getSnapshot().root.slideSize;
    await agent.applyCommand({
      type: "pptx:add-connector",
      payload: {
        slideIndex: 0,
        connectorType: "straight",
        start: { kind: "free", xEmu: -500_000, yEmu: -200_000 },
        end: { kind: "free", xEmu: slideSize.cxEmu + 500_000, yEmu: slideSize.cyEmu + 200_000 },
      },
      source: "human",
    });
    const slide = agent.getSnapshot().root.slides[0];
    const c = slide.shapes.find((s) => s.kind === "connector") as ConnectorShape | undefined;
    expect(c).toBeDefined();
    // Endpoints survive verbatim — the renderer is responsible for clipping,
    // not the model. Position can be negative because the bbox spans the
    // off-slide rectangle.
    expect((c!.start as { xEmu: number }).xEmu).toBe(-500_000);
    expect(c!.position?.xEmu).toBe(-500_000);
  });

  it("self-loop (anchor-to-self) survives the round-trip without crashing", async () => {
    const agent = await loadAgent("01-blank.pptx");
    await agent.applyCommand({
      type: "pptx:add-shape",
      payload: { slideIndex: 0, preset: "rect", x: 2_000_000, y: 2_000_000, width: 2_000_000, height: 1_000_000 },
      source: "human",
    });
    const slide0 = agent.getSnapshot().root.slides[0];
    const rect = slide0.shapes[slide0.shapes.length - 1] as TextShape;
    const m = await agent.applyCommand({
      type: "pptx:add-connector",
      payload: {
        slideIndex: 0,
        connectorType: "elbow",
        start: { kind: "anchored", targetCNvPrId: rect.cNvPrId, side: "e" },
        end: { kind: "anchored", targetCNvPrId: rect.cNvPrId, side: "w" },
      },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const slide = agent.getSnapshot().root.slides[0];
    const c = slide.shapes.find((s) => s.kind === "connector") as ConnectorShape | undefined;
    expect(c).toBeDefined();
    if (c?.start.kind === "anchored" && c.end.kind === "anchored") {
      expect(c.start.targetCNvPrId).toBe(c.end.targetCNvPrId);
      expect(c.start.targetCNvPrId).toBe(rect.cNvPrId);
    }
  });

  it("converts anchored endpoints to free when their target shape is deleted", async () => {
    const { agent, rect1, connector } = await setupTwoShapesAndConnector();
    expect(connector.start.kind).toBe("anchored");
    await agent.applyCommand({
      type: "pptx:delete-shape",
      payload: { slideIndex: 0, shapeId: rect1.id },
      source: "human",
    });
    const slide = agent.getSnapshot().root.slides[0];
    const c = slide.shapes.find((s) => s.kind === "connector") as ConnectorShape | undefined;
    expect(c).toBeDefined();
    // The start endpoint that targeted rect1 should have been detached
    // to a free endpoint at rect1's east anchor (3_000_000, 1_500_000).
    expect(c!.start.kind).toBe("free");
    if (c!.start.kind === "free") {
      expect(c!.start.xEmu).toBe(3_000_000);
      expect(c!.start.yEmu).toBe(1_500_000);
    }
    // The end endpoint that targeted rect2 (still alive) is left intact.
    expect(c!.end.kind).toBe("anchored");
  });
});

describe("connector OOXML round-trip", () => {
  it("survives serialise → re-parse with anchors and style intact", async () => {
    const { agent, rect1, rect2 } = await setupTwoShapesAndConnector();
    const buf = await agent.exportFile();
    const reloaded = await PptxAgent.fromBuffer(buf);
    const slide = reloaded.getSnapshot().root.slides[0];
    const cxn = slide.shapes.find((s) => s.kind === "connector") as ConnectorShape | undefined;
    expect(cxn).toBeDefined();
    expect(cxn!.connectorType).toBe("straight");
    expect(cxn!.start.kind).toBe("anchored");
    expect(cxn!.end.kind).toBe("anchored");
    if (cxn!.start.kind === "anchored") expect(cxn!.start.targetCNvPrId).toBe(rect1.cNvPrId);
    if (cxn!.end.kind === "anchored") expect(cxn!.end.targetCNvPrId).toBe(rect2.cNvPrId);
  });
});
