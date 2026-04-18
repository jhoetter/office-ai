import { CommandError, type DiffChange, type DocumentDiff, type NodeId } from "@officeai/core";
import type {
  ContentTypesSnap,
  PptxDirty,
  PptxPresentation,
  PptxSnapshot,
  RelationshipsSnap,
  Shape,
  Slide,
  TableShape,
  TextShape,
} from "../model/types.js";

// ─── Snapshot evolution ───────────────────────────────────────────────────

export interface DirtyMutator {
  presentation?: boolean;
  contentTypes?: boolean;
  slides?: ReadonlyArray<string>;
  removeSlides?: ReadonlyArray<string>;
  notesSlides?: ReadonlyArray<string>;
  media?: ReadonlyArray<string>;
  relationships?: ReadonlyArray<string>;
}

export function evolveSnapshot(
  snapshot: PptxSnapshot,
  root: PptxPresentation,
  mut: DirtyMutator,
  opts?: {
    relationships?: ReadonlyMap<string, RelationshipsSnap>;
    contentTypes?: ContentTypesSnap;
    removedParts?: ReadonlySet<string>;
  }
): PptxSnapshot {
  const dirty: PptxDirty = {
    presentation: snapshot.dirty.presentation || (mut.presentation ?? false),
    slides: extendSet(snapshot.dirty.slides, mut.slides ?? []),
    notesSlides: extendSet(snapshot.dirty.notesSlides, mut.notesSlides ?? []),
    masters: snapshot.dirty.masters,
    layouts: snapshot.dirty.layouts,
    theme: snapshot.dirty.theme,
    media: extendSet(snapshot.dirty.media, mut.media ?? []),
    relationships: extendSet(snapshot.dirty.relationships, mut.relationships ?? []),
    contentTypes: snapshot.dirty.contentTypes || (mut.contentTypes ?? false),
  };
  return {
    ...snapshot,
    revision: snapshot.revision + 1,
    root,
    dirty,
    relationships: opts?.relationships ?? snapshot.relationships,
    contentTypes: opts?.contentTypes ?? snapshot.contentTypes,
    removedParts: opts?.removedParts ?? snapshot.removedParts,
  };
}

function extendSet(s: ReadonlySet<string>, items: ReadonlyArray<string>): ReadonlySet<string> {
  if (items.length === 0) return s;
  const out = new Set(s);
  for (const i of items) out.add(i);
  return out;
}

// ─── Diff helpers ─────────────────────────────────────────────────────────

export function buildDiff(
  fromRevision: number,
  toRevision: number,
  ...changes: DiffChange[]
): DocumentDiff {
  return {
    format: "pptx",
    fromRevision,
    toRevision,
    changes,
  };
}

// ─── Slides + shapes ──────────────────────────────────────────────────────

export function findSlide(
  snapshot: PptxSnapshot,
  slideIndex: number
): { slide: Slide; index: number } {
  const slides = snapshot.root.slides;
  if (slideIndex < 0 || slideIndex >= slides.length) {
    throw makeError("unknown-target", `slideIndex ${slideIndex} out of range (0..${slides.length})`);
  }
  return { slide: slides[slideIndex], index: slideIndex };
}

export function findShapeInSlide(
  slide: Slide,
  shapeId: NodeId
): { shape: Shape; path: number[] } {
  for (let i = 0; i < slide.shapes.length; i++) {
    const s = slide.shapes[i];
    if (s.id === shapeId) return { shape: s, path: [i] };
    if (s.kind === "group") {
      const inner = findShapeInGroup(s.children, shapeId, [i]);
      if (inner) return inner;
    }
  }
  throw makeError("unknown-target", `shape ${shapeId} not found on slide`);
}

function findShapeInGroup(
  shapes: ReadonlyArray<Shape>,
  shapeId: NodeId,
  prefix: number[]
): { shape: Shape; path: number[] } | null {
  for (let i = 0; i < shapes.length; i++) {
    const s = shapes[i];
    const path = [...prefix, i];
    if (s.id === shapeId) return { shape: s, path };
    if (s.kind === "group") {
      const inner = findShapeInGroup(s.children, shapeId, path);
      if (inner) return inner;
    }
  }
  return null;
}

export function replaceShape(
  shapes: ReadonlyArray<Shape>,
  path: ReadonlyArray<number>,
  next: Shape
): Shape[] {
  if (path.length === 0) throw new Error("empty path");
  const [head, ...tail] = path;
  const out = [...shapes];
  if (tail.length === 0) {
    out[head] = next;
    return out;
  }
  const target = out[head];
  if (target.kind !== "group") throw new Error("path expects group");
  out[head] = {
    ...target,
    children: replaceShape(target.children, tail, next),
  };
  return out;
}

export function withSlide(
  root: PptxPresentation,
  slideIndex: number,
  fn: (slide: Slide) => Slide
): PptxPresentation {
  const slides = [...root.slides];
  slides[slideIndex] = fn(slides[slideIndex]);
  return { ...root, slides };
}

export function maxCNvPrId(shapes: ReadonlyArray<Shape>): number {
  let max = 0;
  walk(shapes, (s) => {
    if (s.cNvPrId > max) max = s.cNvPrId;
  });
  return max;
}

function walk(shapes: ReadonlyArray<Shape>, fn: (s: Shape) => void): void {
  for (const s of shapes) {
    fn(s);
    if (s.kind === "group") walk(s.children, fn);
  }
}

// ─── Text helpers ─────────────────────────────────────────────────────────

export function isTextShape(s: Shape): s is TextShape {
  return s.kind === "text";
}

export function isTableShape(s: Shape): s is TableShape {
  return s.kind === "table";
}

// ─── Errors ───────────────────────────────────────────────────────────────

export function makeError(code: string, message: string): CommandError {
  return new CommandError(code, message);
}
