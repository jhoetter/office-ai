import { ooxml } from "@officeai/core";
import type { AbstractNum, NumInstance, NumberingDefinitions, NumberingLevel } from "../model/types.js";
import { DocxParseError } from "./errors.js";
import { attrOf, captureOpaque, elementEntries, findElementEntry } from "./xml-helpers.js";

const NUMBERING_PART = "word/numbering.xml";

/**
 * Parse `word/numbering.xml` into a typed `NumberingDefinitions`. Returns
 * `undefined` when the part is absent (most documents have no list
 * paragraphs and therefore no numbering part at all).
 *
 * Byte-preservation contract: the typed projection is read-only metadata
 * for commands; the serializer never writes back from it unless
 * `dirty.numbering` is set, which only happens when a future workstream
 * mutates a numbering definition. Today's list commands (`set-paragraph
 * -list` / `remove-paragraph-list`) only flip `<w:numPr>` references in
 * `word/document.xml`, so the numbering part keeps round-tripping
 * verbatim through the container's part cache.
 */
export function parseNumberingPart(container: ooxml.OoxmlContainer): NumberingDefinitions | undefined {
  if (!container.has(NUMBERING_PART)) return undefined;
  let tree: unknown;
  try {
    tree = ooxml.parseXml(container.readText(NUMBERING_PART));
  } catch (err) {
    throw new DocxParseError("invalid-xml", "Failed to parse numbering.xml", {
      partPath: NUMBERING_PART,
      cause: err,
    });
  }
  if (!Array.isArray(tree)) {
    return { abstractNums: new Map(), nums: new Map() };
  }
  const root = findElementEntry(tree as unknown[], "w:numbering");
  if (!root) {
    return { abstractNums: new Map(), nums: new Map() };
  }

  const abstractNums = new Map<string, AbstractNum>();
  const nums = new Map<number, NumInstance>();
  for (const entry of elementEntries((root["w:numbering"] as unknown[] | undefined) ?? [])) {
    const tag = ooxml.getTag(entry);
    if (tag === "w:abstractNum") {
      const a = parseAbstractNum(entry);
      abstractNums.set(a.id, a);
    } else if (tag === "w:num") {
      const n = parseNum(entry);
      if (n) nums.set(n.numId, n);
    }
  }
  return { abstractNums, nums };
}

function parseAbstractNum(entry: Record<string, unknown>): AbstractNum {
  const id = attrOf(entry, "w:abstractNumId") ?? "";
  const children = (entry["w:abstractNum"] as unknown[] | undefined) ?? [];
  const levels: NumberingLevel[] = [];
  let multiLevelType: AbstractNum["multiLevelType"];
  for (const c of elementEntries(children)) {
    const tag = ooxml.getTag(c);
    if (tag === "w:multiLevelType") {
      const v = attrOf(c, "w:val");
      if (v === "singleLevel" || v === "multilevel" || v === "hybridMultilevel") {
        multiLevelType = v;
      }
    } else if (tag === "w:lvl") {
      levels.push(parseLevel(c));
    }
  }
  return {
    id,
    ...(multiLevelType ? { multiLevelType } : {}),
    levels,
    raw: captureOpaque(entry),
  };
}

function parseLevel(entry: Record<string, unknown>): NumberingLevel {
  const ilvlAttr = attrOf(entry, "w:ilvl");
  const ilvl = ilvlAttr !== undefined ? Number(ilvlAttr) : 0;
  const children = (entry["w:lvl"] as unknown[] | undefined) ?? [];
  let numFmt: string | undefined;
  let lvlText: string | undefined;
  let start: number | undefined;
  const opaqueProps: ReturnType<typeof captureOpaque>[] = [];
  for (const c of elementEntries(children)) {
    const tag = ooxml.getTag(c);
    if (tag === "w:numFmt") {
      const v = attrOf(c, "w:val");
      if (v) numFmt = v;
    } else if (tag === "w:lvlText") {
      const v = attrOf(c, "w:val");
      if (v !== undefined) lvlText = v;
    } else if (tag === "w:start") {
      const v = attrOf(c, "w:val");
      if (v !== undefined) start = Number(v);
    } else {
      opaqueProps.push(captureOpaque(c));
    }
  }
  return {
    ilvl,
    ...(numFmt !== undefined ? { numFmt } : {}),
    ...(lvlText !== undefined ? { lvlText } : {}),
    ...(start !== undefined ? { start } : {}),
    ...(opaqueProps.length > 0 ? { opaqueProps } : {}),
  };
}

function parseNum(entry: Record<string, unknown>): NumInstance | null {
  const numIdAttr = attrOf(entry, "w:numId");
  if (numIdAttr === undefined) return null;
  const numId = Number(numIdAttr);
  if (!Number.isFinite(numId)) return null;
  const children = (entry["w:num"] as unknown[] | undefined) ?? [];
  let abstractNumId = "";
  const lvlOverrides: Array<{ ilvl: number; startOverride?: number }> = [];
  for (const c of elementEntries(children)) {
    const tag = ooxml.getTag(c);
    if (tag === "w:abstractNumId") {
      const v = attrOf(c, "w:val");
      if (v !== undefined) abstractNumId = v;
    } else if (tag === "w:lvlOverride") {
      const ilvlAttr = attrOf(c, "w:ilvl");
      const ilvl = ilvlAttr !== undefined ? Number(ilvlAttr) : 0;
      let startOverride: number | undefined;
      const lvlOvChildren = (c["w:lvlOverride"] as unknown[] | undefined) ?? [];
      const startEl = findElementEntry(lvlOvChildren, "w:startOverride");
      if (startEl) {
        const sv = attrOf(startEl, "w:val");
        if (sv !== undefined) startOverride = Number(sv);
      }
      lvlOverrides.push(startOverride !== undefined ? { ilvl, startOverride } : { ilvl });
    }
  }
  return {
    numId,
    abstractNumId,
    ...(lvlOverrides.length > 0 ? { lvlOverrides } : {}),
    raw: captureOpaque(entry),
  };
}
