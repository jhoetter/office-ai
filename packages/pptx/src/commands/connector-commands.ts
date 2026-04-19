import type { CommandHandler } from "@officeai/core";
import type {
  ConnectorDashStyle,
  ConnectorEndShape,
  ConnectorEndpoint,
  ConnectorShape,
  ConnectorSide,
  ConnectorStroke,
  ConnectorType,
  OpaqueXml,
  PptxSnapshot,
  Shape,
  Slide,
} from "../model/types.js";
import { bboxFromEndpoints, resolveEndpoint } from "../model/connector-geometry.js";
import { recomputeConnectorBox } from "./connector-helpers.js";
import {
  buildDiff,
  evolveSnapshot,
  findShapeInSlide,
  findSlide,
  isConnectorShape,
  makeError,
  maxCNvPrId,
  replaceShape,
} from "./helpers.js";
import type {
  AddConnectorPayload,
  ConnectorDashStylePayload,
  ConnectorEndpointPayload,
  RerouteConnectorPayload,
  SetConnectorEndpointPayload,
  SetConnectorStylePayload,
  SetConnectorWaypointPayload,
  SwapConnectorDirectionPayload,
} from "./payloads.js";

const DEFAULT_STROKE_COLOR = "374151"; // slate-700
const DEFAULT_STROKE_WIDTH_EMU = 9525; // ≈ 0.75pt

// ─── add-connector ────────────────────────────────────────────────────────

export const addConnectorHandler: CommandHandler<AddConnectorPayload, PptxSnapshot> = {
  type: "pptx:add-connector",
  apply(snapshot, payload, ctx) {
    if (!isKnownType(payload.connectorType)) {
      throw makeError("invalid-payload", `unknown connector type: ${payload.connectorType}`);
    }
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const start = validateEndpoint(payload.start, slide.shapes, "start");
    const end = validateEndpoint(payload.end, slide.shapes, "end");

    const stroke: ConnectorStroke = {
      color: normaliseHex(payload.strokeColor ?? DEFAULT_STROKE_COLOR),
      widthEmu: payload.strokeWidthEmu ?? DEFAULT_STROKE_WIDTH_EMU,
      ...(payload.strokeDash !== undefined ? { dash: validateDash(payload.strokeDash) } : {}),
    };
    if (stroke.widthEmu < 0) {
      throw makeError("invalid-payload", "strokeWidthEmu must be ≥ 0");
    }

    const cNvPrId = maxCNvPrId(slide.shapes) + 1;
    const name = payload.name ?? `Connector ${cNvPrId}`;

    // Compute initial bounding box from resolved endpoints. Anchored
    // endpoints walk through `shapesByCNvPrId` to read the target's
    // current geometry.
    const map = new Map<number, Shape>();
    for (const s of slide.shapes) collectShapesByCNvPrId(s, map);
    const startPt = resolveEndpoint(start, map) ?? { x: 0, y: 0 };
    const endPt = resolveEndpoint(end, map) ?? { x: 0, y: 0 };
    const box = bboxFromEndpoints(startPt, endPt);

    const connector: ConnectorShape = {
      kind: "connector",
      id: ctx.mintNodeId(),
      cNvPrId,
      name,
      position: { xEmu: box.x, yEmu: box.y },
      size: { cxEmu: box.cx, cyEmu: box.cy },
      connectorType: payload.connectorType,
      start,
      end,
      stroke,
      ...(payload.headEnd ? { headEnd: payload.headEnd } : { headEnd: "arrow" as ConnectorEndShape }),
      ...(payload.tailEnd ? { tailEnd: payload.tailEnd } : {}),
      nvCxnSpPrTail: defaultNvCxnSpPrTail(),
      spPrTail: [],
    };

    const newSlide: Slide = { ...slide, shapes: [...slide.shapes, connector] };
    const root = {
      ...snapshot.root,
      slides: snapshot.root.slides.map((s, i) => (i === sIdx ? newSlide : s)),
    };
    const next = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-inserted",
        nodeId: connector.id,
        path: ["slides", sIdx, "shapes", newSlide.shapes.length - 1],
        summary: `connector:${payload.connectorType}`,
      }),
    };
  },
};

// ─── set-connector-endpoint ───────────────────────────────────────────────

export const setConnectorEndpointHandler: CommandHandler<SetConnectorEndpointPayload, PptxSnapshot> = {
  type: "pptx:set-connector-endpoint",
  apply(snapshot, payload) {
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const found = findShapeInSlide(slide, payload.shapeId);
    const connector = found.shape;
    if (!isConnectorShape(connector)) {
      throw makeError("invalid-target", `shape ${payload.shapeId} is not a connector`);
    }
    if (payload.which !== "start" && payload.which !== "end") {
      throw makeError("invalid-payload", `which must be "start" or "end"`);
    }
    const newEndpoint = validateEndpoint(payload.endpoint, slide.shapes, payload.which);
    const updatedConnector: ConnectorShape = {
      ...connector,
      ...(payload.which === "start" ? { start: newEndpoint } : { end: newEndpoint }),
    };
    const map = new Map<number, Shape>();
    for (const s of slide.shapes) collectShapesByCNvPrId(s, map);
    const reflowed = recomputeConnectorBox(updatedConnector, map);
    const nextShapes = replaceShape(slide.shapes, found.path, reflowed);
    const newSlide: Slide = { ...slide, shapes: nextShapes };
    const root = {
      ...snapshot.root,
      slides: snapshot.root.slides.map((s, i) => (i === sIdx ? newSlide : s)),
    };
    const next = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: reflowed.id,
        path: ["slides", sIdx, "shapes", ...found.path],
        field: payload.which === "start" ? "start" : "end",
        summary: `connector-endpoint:${payload.which}`,
      }),
    };
  },
};

// ─── set-connector-style ──────────────────────────────────────────────────

export const setConnectorStyleHandler: CommandHandler<SetConnectorStylePayload, PptxSnapshot> = {
  type: "pptx:set-connector-style",
  apply(snapshot, payload) {
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const found = findShapeInSlide(slide, payload.shapeId);
    const connector = found.shape;
    if (!isConnectorShape(connector)) {
      throw makeError("invalid-target", `shape ${payload.shapeId} is not a connector`);
    }
    if (
      payload.connectorType === undefined &&
      payload.strokeColor === undefined &&
      payload.strokeWidthEmu === undefined &&
      payload.strokeDash === undefined &&
      payload.headEnd === undefined &&
      payload.tailEnd === undefined
    ) {
      throw makeError("invalid-payload", "set-connector-style requires at least one field");
    }
    if (payload.connectorType !== undefined && !isKnownType(payload.connectorType)) {
      throw makeError("invalid-payload", `unknown connector type: ${payload.connectorType}`);
    }
    const prevStroke: ConnectorStroke = connector.stroke ?? {
      color: DEFAULT_STROKE_COLOR,
      widthEmu: DEFAULT_STROKE_WIDTH_EMU,
    };
    const nextDash: ConnectorDashStyle | undefined =
      payload.strokeDash !== undefined ? validateDash(payload.strokeDash) : prevStroke.dash;
    const stroke: ConnectorStroke = {
      color: payload.strokeColor !== undefined ? normaliseHex(payload.strokeColor) : prevStroke.color,
      widthEmu: payload.strokeWidthEmu ?? prevStroke.widthEmu,
      ...(nextDash !== undefined ? { dash: nextDash } : {}),
    };
    if (stroke.widthEmu < 0) {
      throw makeError("invalid-payload", "strokeWidthEmu must be ≥ 0");
    }
    const updated: ConnectorShape = {
      ...connector,
      ...(payload.connectorType !== undefined ? { connectorType: payload.connectorType } : {}),
      stroke,
      ...(payload.headEnd !== undefined ? { headEnd: payload.headEnd } : {}),
      ...(payload.tailEnd !== undefined ? { tailEnd: payload.tailEnd } : {}),
    };
    const nextShapes = replaceShape(slide.shapes, found.path, updated);
    const newSlide: Slide = { ...slide, shapes: nextShapes };
    const root = {
      ...snapshot.root,
      slides: snapshot.root.slides.map((s, i) => (i === sIdx ? newSlide : s)),
    };
    const next = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: updated.id,
        path: ["slides", sIdx, "shapes", ...found.path],
        field: "style",
        summary: "connector-style",
      }),
    };
  },
};

// ─── set-connector-waypoint ───────────────────────────────────────────────

export const setConnectorWaypointHandler: CommandHandler<SetConnectorWaypointPayload, PptxSnapshot> = {
  type: "pptx:set-connector-waypoint",
  apply(snapshot, payload) {
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const found = findShapeInSlide(slide, payload.shapeId);
    const connector = found.shape;
    if (!isConnectorShape(connector)) {
      throw makeError("invalid-target", `shape ${payload.shapeId} is not a connector`);
    }
    if (!Number.isInteger(payload.segmentIndex) || payload.segmentIndex < 0) {
      throw makeError("invalid-payload", "segmentIndex must be a non-negative integer");
    }
    if (payload.valueEmu !== null && !Number.isFinite(payload.valueEmu)) {
      throw makeError("invalid-payload", "valueEmu must be a finite number or null");
    }
    const prev = connector.waypoints ? [...connector.waypoints] : [];
    if (payload.valueEmu === null) {
      // Clearing: drop the entry at this index, then trim trailing
      // undefined slots so we don't carry empty padding.
      if (payload.segmentIndex < prev.length) {
        prev.splice(payload.segmentIndex, 1);
      }
    } else {
      while (prev.length <= payload.segmentIndex) prev.push(0);
      prev[payload.segmentIndex] = Math.round(payload.valueEmu);
    }
    const updated: ConnectorShape = {
      ...connector,
      ...(prev.length > 0 ? { waypoints: prev } : { waypoints: undefined }),
    };
    const nextShapes = replaceShape(slide.shapes, found.path, updated);
    const newSlide: Slide = { ...slide, shapes: nextShapes };
    const root = {
      ...snapshot.root,
      slides: snapshot.root.slides.map((s, i) => (i === sIdx ? newSlide : s)),
    };
    const next = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: updated.id,
        path: ["slides", sIdx, "shapes", ...found.path],
        field: "waypoints",
        summary: `connector-waypoint:${payload.segmentIndex}`,
      }),
    };
  },
};

// ─── reroute-connector ────────────────────────────────────────────────────

/**
 * Drop every user-supplied waypoint on a connector and let the auto-
 * router pick the polyline from scratch. No-op (still emits a diff)
 * when the connector has no waypoints — keeps the agent's idempotent
 * "force a clean route" semantics.
 */
export const rerouteConnectorHandler: CommandHandler<RerouteConnectorPayload, PptxSnapshot> = {
  type: "pptx:reroute-connector",
  apply(snapshot, payload) {
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const found = findShapeInSlide(slide, payload.shapeId);
    const connector = found.shape;
    if (!isConnectorShape(connector)) {
      throw makeError("invalid-target", `shape ${payload.shapeId} is not a connector`);
    }
    const updated: ConnectorShape = { ...connector, waypoints: undefined };
    const nextShapes = replaceShape(slide.shapes, found.path, updated);
    const newSlide: Slide = { ...slide, shapes: nextShapes };
    const root = {
      ...snapshot.root,
      slides: snapshot.root.slides.map((s, i) => (i === sIdx ? newSlide : s)),
    };
    const next = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: updated.id,
        path: ["slides", sIdx, "shapes", ...found.path],
        field: "waypoints",
        summary: "connector-reroute",
      }),
    };
  },
};

// ─── swap-connector-direction ─────────────────────────────────────────────

/**
 * Reverse a connector's direction by swapping `start` ↔ `end` and
 * `headEnd` ↔ `tailEnd`. Waypoints are kept as-is — the routing
 * engine treats them as absolute coordinates so they apply equally
 * well in either direction. The bounding box doesn't change because
 * `bboxFromEndpoints` is order-independent, but we still recompute it
 * for symmetry with the endpoint command.
 */
export const swapConnectorDirectionHandler: CommandHandler<SwapConnectorDirectionPayload, PptxSnapshot> = {
  type: "pptx:swap-connector-direction",
  apply(snapshot, payload) {
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const found = findShapeInSlide(slide, payload.shapeId);
    const connector = found.shape;
    if (!isConnectorShape(connector)) {
      throw makeError("invalid-target", `shape ${payload.shapeId} is not a connector`);
    }
    const swapped: ConnectorShape = {
      ...connector,
      start: connector.end,
      end: connector.start,
      ...(connector.headEnd !== undefined || connector.tailEnd !== undefined
        ? { headEnd: connector.tailEnd, tailEnd: connector.headEnd }
        : {}),
    };
    const map = new Map<number, Shape>();
    for (const s of slide.shapes) collectShapesByCNvPrId(s, map);
    const reflowed = recomputeConnectorBox(swapped, map);
    const nextShapes = replaceShape(slide.shapes, found.path, reflowed);
    const newSlide: Slide = { ...slide, shapes: nextShapes };
    const root = {
      ...snapshot.root,
      slides: snapshot.root.slides.map((s, i) => (i === sIdx ? newSlide : s)),
    };
    const next = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: reflowed.id,
        path: ["slides", sIdx, "shapes", ...found.path],
        field: "direction",
        summary: "connector-swap",
      }),
    };
  },
};

// ─── helpers ──────────────────────────────────────────────────────────────

function isKnownType(t: string): t is ConnectorType {
  return t === "straight" || t === "elbow" || t === "curved";
}

function isKnownSide(s: string): s is ConnectorSide {
  return s === "n" || s === "s" || s === "e" || s === "w" || s === "center";
}

function validateDash(d: ConnectorDashStylePayload): ConnectorDashStyle {
  if (
    d === "solid" ||
    d === "dashed" ||
    d === "dotted" ||
    d === "longDash" ||
    d === "dashDot"
  ) {
    return d;
  }
  throw makeError("invalid-payload", `unknown dash style: ${d}`);
}

/**
 * Sanity-check an endpoint payload before persisting it. Anchored
 * endpoints need to point at a real shape on the slide so we don't end
 * up with a connector that's already broken; free endpoints just need
 * finite numbers. We deliberately allow self-anchoring (a connector
 * whose start and end target the same shape) — PowerPoint allows it
 * and it's useful for self-loops.
 */
function validateEndpoint(
  ep: ConnectorEndpointPayload,
  shapes: ReadonlyArray<Shape>,
  label: string
): ConnectorEndpoint {
  if (ep.kind === "free") {
    if (!Number.isFinite(ep.xEmu) || !Number.isFinite(ep.yEmu)) {
      throw makeError("invalid-payload", `${label} endpoint requires finite xEmu / yEmu`);
    }
    return { kind: "free", xEmu: Math.round(ep.xEmu), yEmu: Math.round(ep.yEmu) };
  }
  if (!Number.isFinite(ep.targetCNvPrId) || ep.targetCNvPrId <= 0) {
    throw makeError("invalid-payload", `${label} endpoint requires a positive targetCNvPrId`);
  }
  if (!isKnownSide(ep.side)) {
    throw makeError("invalid-payload", `${label} endpoint side must be n/s/e/w/center`);
  }
  // Walk groups too — PowerPoint allows anchoring to grouped shapes.
  const map = new Map<number, Shape>();
  for (const s of shapes) collectShapesByCNvPrId(s, map);
  if (!map.has(ep.targetCNvPrId)) {
    throw makeError(
      "unknown-target",
      `${label} endpoint targetCNvPrId ${ep.targetCNvPrId} not found on slide`
    );
  }
  // `t` is meaningful only for n/s/e/w (the four edges); `center`
  // ignores it. Clamp + round to 3 decimals so lossy round-trip stays
  // predictable.
  const t = normaliseT(ep.t, ep.side);
  return {
    kind: "anchored",
    targetCNvPrId: ep.targetCNvPrId,
    side: ep.side,
    ...(t !== undefined ? { t } : {}),
  };
}

function normaliseT(t: number | undefined, side: ConnectorSide): number | undefined {
  if (side === "center") return undefined;
  if (t === undefined) return undefined;
  if (!Number.isFinite(t)) {
    throw makeError("invalid-payload", "endpoint t must be a finite number in [0, 1]");
  }
  const clamped = t < 0 ? 0 : t > 1 ? 1 : t;
  return Math.round(clamped * 1000) / 1000;
}

function collectShapesByCNvPrId(shape: Shape, out: Map<number, Shape>): void {
  if (shape.cNvPrId > 0) out.set(shape.cNvPrId, shape);
  if (shape.kind === "group") {
    for (const c of shape.children) collectShapesByCNvPrId(c, out);
  }
}

function defaultNvCxnSpPrTail(): OpaqueXml[] {
  // Mirror the structure parsed from a freshly-authored connector — the
  // serializer rebuilds <p:cNvPr> and <p:cNvCxnSpPr> from the typed
  // model, so the only thing we need here is the empty <p:nvPr> sibling
  // PowerPoint expects. (Our serializer also synthesises one when
  // missing, but having it here keeps the parsed tail symmetrical.)
  return [
    { tag: "p:cNvPr", attrs: { id: "0", name: "" }, rawAttrs: { "@_id": "0", "@_name": "" }, subtree: [] },
    { tag: "p:cNvCxnSpPr", attrs: {}, rawAttrs: {}, subtree: [] },
    { tag: "p:nvPr", attrs: {}, rawAttrs: {}, subtree: [] },
  ];
}

function normaliseHex(input: string): string {
  const v = input.trim().replace(/^#/, "").toUpperCase();
  if (!/^[0-9A-F]{6}$/.test(v)) {
    throw makeError("invalid-payload", `invalid hex color: ${input}`);
  }
  return v;
}
