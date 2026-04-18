import { ooxml } from "@officeai/core";
import type {
  AbstractNum,
  DocxSnapshot,
  NumInstance,
  NumberingDefinitions,
  NumberingLevel,
} from "../model/types.js";
import { opaqueToEntry } from "../parser/xml-helpers.js";
import { DocxSerializeError } from "./errors.js";

const NUMBERING_PART = "word/numbering.xml";
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const NUMBERING_ROOT_ATTRS: Record<string, string> = {
  "@_xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
};

/**
 * Re-emit `word/numbering.xml` from the typed `NumberingDefinitions`
 * carrier. Only invoked when `dirty.numbering` is set; untouched
 * numbering parts ride the container's part cache so their bytes
 * survive round-trip exactly.
 *
 * Today's W10 commands (`set-paragraph-list` /
 * `remove-paragraph-list`) only flip pointers in `word/document.xml`
 * and never set `dirty.numbering`, so this code path is currently only
 * reached if a future workstream mutates the abstract / num
 * definitions themselves. Each `<w:abstractNum>` and `<w:num>`
 * carrying a `raw` opaque snapshot is re-emitted from that snapshot
 * verbatim — the typed projection only takes over when `raw` is absent
 * (i.e. the entry was synthesised in-memory and has nothing to
 * preserve).
 */
export function serializeNumberingPart(container: ooxml.OoxmlContainer, snapshot: DocxSnapshot): void {
  if (!snapshot.dirty.numbering) return;
  const defs = snapshot.root.numbering;
  if (!defs) {
    throw new DocxSerializeError(
      "numbering-missing",
      "dirty.numbering is set but snapshot.root.numbering is undefined"
    );
  }
  const xml = serializeNumberingXml(defs);
  if (!container.has(NUMBERING_PART)) {
    container.addPart(NUMBERING_PART, new TextEncoder().encode(xml));
  } else {
    container.writeText(NUMBERING_PART, xml);
  }
}

function serializeNumberingXml(defs: NumberingDefinitions): string {
  const children: unknown[] = [];
  for (const a of defs.abstractNums.values()) {
    children.push(serializeAbstractNum(a));
  }
  for (const n of defs.nums.values()) {
    children.push(serializeNumInstance(n));
  }
  const tree = [{ "w:numbering": children, ":@": { ...NUMBERING_ROOT_ATTRS } }];
  return ooxml.serializeXml(tree, { xmlDeclaration: XML_DECL });
}

function serializeAbstractNum(a: AbstractNum): unknown {
  if (a.raw) return opaqueToEntry(a.raw);
  const inner: unknown[] = [];
  if (a.multiLevelType) {
    inner.push({ "w:multiLevelType": [], ":@": { "@_w:val": a.multiLevelType } });
  }
  for (const lvl of a.levels) {
    inner.push(serializeLevel(lvl));
  }
  return { "w:abstractNum": inner, ":@": { "@_w:abstractNumId": a.id } };
}

function serializeLevel(lvl: NumberingLevel): unknown {
  const children: unknown[] = [];
  if (lvl.start !== undefined) {
    children.push({ "w:start": [], ":@": { "@_w:val": String(lvl.start) } });
  }
  if (lvl.numFmt !== undefined) {
    children.push({ "w:numFmt": [], ":@": { "@_w:val": lvl.numFmt } });
  }
  if (lvl.lvlText !== undefined) {
    children.push({ "w:lvlText": [], ":@": { "@_w:val": lvl.lvlText } });
  }
  if (lvl.opaqueProps) {
    for (const o of lvl.opaqueProps) children.push(opaqueToEntry(o));
  }
  return { "w:lvl": children, ":@": { "@_w:ilvl": String(lvl.ilvl) } };
}

function serializeNumInstance(n: NumInstance): unknown {
  if (n.raw) return opaqueToEntry(n.raw);
  const children: unknown[] = [{ "w:abstractNumId": [], ":@": { "@_w:val": n.abstractNumId } }];
  if (n.lvlOverrides) {
    for (const ov of n.lvlOverrides) {
      const ovChildren: unknown[] = [];
      if (ov.startOverride !== undefined) {
        ovChildren.push({ "w:startOverride": [], ":@": { "@_w:val": String(ov.startOverride) } });
      }
      children.push({ "w:lvlOverride": ovChildren, ":@": { "@_w:ilvl": String(ov.ilvl) } });
    }
  }
  return { "w:num": children, ":@": { "@_w:numId": String(n.numId) } };
}
