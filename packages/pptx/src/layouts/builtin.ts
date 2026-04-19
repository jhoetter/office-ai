/**
 * Built-in PowerPoint slide-layout templates.
 *
 * Authored as raw `<p:sldLayout>` XML strings keyed by `LayoutKind` so we
 * can synthesise a layout part on demand when the deck doesn't already
 * carry one of the requested kind. Placeholder rectangles use generous
 * percentages of the standard 16:9 slide (12_192_000 × 6_858_000 EMU)
 * so the cloned shapes look reasonable on any deck size — `pptx:add-
 * slide` rescales them when needed.
 *
 * Every template carries the `<p:sldLayout type="…">` attribute so the
 * classifier round-trips cleanly through `parseSlideLayout`.
 */

import type { LayoutKind } from "../model/types.js";

export interface BuiltinLayout {
  readonly name: string;
  readonly typeAttr: string;
  /** Display label for the toolbar picker. */
  readonly label: string;
  readonly xml: string;
}

const NS_ATTRS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

function layoutXml(typeAttr: string, name: string, placeholders: ReadonlyArray<PlaceholderEntry>): string {
  const phXml = placeholders.map((p, i) => placeholderShapeXml(p, i + 2)).join("");
  const preserveAttr = typeAttr === "title" ? ' preserve="1"' : "";
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:sldLayout ${NS_ATTRS} type="${typeAttr}"${preserveAttr}>` +
    `<p:cSld name="${escapeXml(name)}">` +
    `<p:spTree>` +
    `<p:nvGrpSpPr>` +
    `<p:cNvPr id="1" name=""/>` +
    `<p:cNvGrpSpPr/>` +
    `<p:nvPr/>` +
    `</p:nvGrpSpPr>` +
    `<p:grpSpPr>` +
    `<a:xfrm>` +
    `<a:off x="0" y="0"/>` +
    `<a:ext cx="0" cy="0"/>` +
    `<a:chOff x="0" y="0"/>` +
    `<a:chExt cx="0" cy="0"/>` +
    `</a:xfrm>` +
    `</p:grpSpPr>` +
    phXml +
    `</p:spTree>` +
    `</p:cSld>` +
    `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>` +
    `</p:sldLayout>`
  );
}

interface PlaceholderEntry {
  readonly type: string;
  readonly idx: number;
  readonly sz?: string;
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
  readonly label: string;
  readonly anchorCenter?: boolean;
}

function placeholderShapeXml(p: PlaceholderEntry, cNvPrId: number): string {
  const phAttrs = `type="${p.type}" idx="${p.idx}"` + (p.sz ? ` sz="${p.sz}"` : "");
  const anchorAttr = p.anchorCenter ? ` anchor="ctr"` : "";
  return (
    `<p:sp>` +
    `<p:nvSpPr>` +
    `<p:cNvPr id="${cNvPrId}" name="${escapeXml(p.label)}"/>` +
    `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
    `<p:nvPr><p:ph ${phAttrs}/></p:nvPr>` +
    `</p:nvSpPr>` +
    `<p:spPr>` +
    `<a:xfrm>` +
    `<a:off x="${p.x}" y="${p.y}"/>` +
    `<a:ext cx="${p.cx}" cy="${p.cy}"/>` +
    `</a:xfrm>` +
    `</p:spPr>` +
    `<p:txBody>` +
    `<a:bodyPr${anchorAttr}/>` +
    `<a:lstStyle/>` +
    `<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${escapeXml(p.label)}</a:t></a:r></a:p>` +
    `</p:txBody>` +
    `</p:sp>`
  );
}

function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// Standard 16:9 slide footprint in EMU.
const SLIDE_W = 12_192_000;
const SLIDE_H = 6_858_000;
const PAD_X = 838_200;
const PAD_TOP = 365_125;

const TITLE_BOX = {
  x: PAD_X,
  y: PAD_TOP,
  cx: SLIDE_W - PAD_X * 2,
  cy: 1_143_000,
};
const SUBTITLE_BOX = {
  x: PAD_X,
  y: TITLE_BOX.y + TITLE_BOX.cy + 200_000,
  cx: TITLE_BOX.cx,
  cy: 1_000_000,
};
const BODY_BOX = {
  x: PAD_X,
  y: TITLE_BOX.y + TITLE_BOX.cy + 365_125,
  cx: TITLE_BOX.cx,
  cy: SLIDE_H - (TITLE_BOX.y + TITLE_BOX.cy + 365_125) - 470_000,
};

export const BUILTIN_LAYOUTS: Readonly<Record<Exclude<LayoutKind, "unknown">, BuiltinLayout>> = {
  title: {
    name: "Title Slide",
    typeAttr: "title",
    label: "Title Slide",
    xml: layoutXml("title", "Title Slide", [
      {
        type: "ctrTitle",
        idx: 0,
        x: PAD_X,
        y: 2_133_600,
        cx: SLIDE_W - PAD_X * 2,
        cy: 1_500_000,
        label: "Click to add title",
        anchorCenter: true,
      },
      {
        type: "subTitle",
        idx: 1,
        x: PAD_X,
        y: 3_900_000,
        cx: SLIDE_W - PAD_X * 2,
        cy: 800_000,
        label: "Click to add subtitle",
        anchorCenter: true,
      },
    ]),
  },
  titleAndContent: {
    name: "Title and Content",
    typeAttr: "obj",
    label: "Title and Content",
    xml: layoutXml("obj", "Title and Content", [
      { type: "title", idx: 0, ...TITLE_BOX, label: "Click to add title" },
      { type: "body", idx: 1, ...BODY_BOX, label: "Click to add text" },
    ]),
  },
  sectionHeader: {
    name: "Section Header",
    typeAttr: "secHead",
    label: "Section Header",
    xml: layoutXml("secHead", "Section Header", [
      {
        type: "title",
        idx: 0,
        x: PAD_X,
        y: 2_500_000,
        cx: SLIDE_W - PAD_X * 2,
        cy: 1_500_000,
        label: "Section title",
      },
      {
        type: "body",
        idx: 1,
        x: PAD_X,
        y: 4_100_000,
        cx: SLIDE_W - PAD_X * 2,
        cy: 800_000,
        label: "Section description",
      },
    ]),
  },
  twoContent: {
    name: "Two Content",
    typeAttr: "twoObj",
    label: "Two Content",
    xml: layoutXml("twoObj", "Two Content", [
      { type: "title", idx: 0, ...TITLE_BOX, label: "Click to add title" },
      {
        type: "body",
        idx: 1,
        sz: "half",
        x: PAD_X,
        y: BODY_BOX.y,
        cx: (BODY_BOX.cx - 304_800) / 2,
        cy: BODY_BOX.cy,
        label: "Click to add text",
      },
      {
        type: "body",
        idx: 2,
        sz: "half",
        x: PAD_X + (BODY_BOX.cx - 304_800) / 2 + 304_800,
        y: BODY_BOX.y,
        cx: (BODY_BOX.cx - 304_800) / 2,
        cy: BODY_BOX.cy,
        label: "Click to add text",
      },
    ]),
  },
  comparison: {
    name: "Comparison",
    typeAttr: "twoTxTwoObj",
    label: "Comparison",
    xml: layoutXml("twoTxTwoObj", "Comparison", [
      { type: "title", idx: 0, ...TITLE_BOX, label: "Click to add title" },
      {
        type: "body",
        idx: 1,
        sz: "quarter",
        x: PAD_X,
        y: BODY_BOX.y,
        cx: (BODY_BOX.cx - 304_800) / 2,
        cy: 600_000,
        label: "Heading",
      },
      {
        type: "body",
        idx: 2,
        sz: "half",
        x: PAD_X,
        y: BODY_BOX.y + 700_000,
        cx: (BODY_BOX.cx - 304_800) / 2,
        cy: BODY_BOX.cy - 700_000,
        label: "Click to add text",
      },
      {
        type: "body",
        idx: 3,
        sz: "quarter",
        x: PAD_X + (BODY_BOX.cx - 304_800) / 2 + 304_800,
        y: BODY_BOX.y,
        cx: (BODY_BOX.cx - 304_800) / 2,
        cy: 600_000,
        label: "Heading",
      },
      {
        type: "body",
        idx: 4,
        sz: "half",
        x: PAD_X + (BODY_BOX.cx - 304_800) / 2 + 304_800,
        y: BODY_BOX.y + 700_000,
        cx: (BODY_BOX.cx - 304_800) / 2,
        cy: BODY_BOX.cy - 700_000,
        label: "Click to add text",
      },
    ]),
  },
  titleOnly: {
    name: "Title Only",
    typeAttr: "titleOnly",
    label: "Title Only",
    xml: layoutXml("titleOnly", "Title Only", [
      { type: "title", idx: 0, ...TITLE_BOX, label: "Click to add title" },
    ]),
  },
  blank: {
    name: "Blank",
    typeAttr: "blank",
    label: "Blank",
    xml: layoutXml("blank", "Blank", []),
  },
  contentWithCaption: {
    name: "Content with Caption",
    typeAttr: "objTx",
    label: "Content with Caption",
    xml: layoutXml("objTx", "Content with Caption", [
      {
        type: "title",
        idx: 0,
        x: PAD_X,
        y: PAD_TOP,
        cx: 4_000_000,
        cy: 1_200_000,
        label: "Click to add title",
      },
      {
        type: "body",
        idx: 1,
        x: PAD_X + 4_200_000,
        y: PAD_TOP,
        cx: SLIDE_W - PAD_X * 2 - 4_200_000,
        cy: SLIDE_H - PAD_TOP - 470_000,
        label: "Click to add content",
      },
      {
        type: "body",
        idx: 2,
        x: PAD_X,
        y: PAD_TOP + 1_400_000,
        cx: 4_000_000,
        cy: SLIDE_H - PAD_TOP - 1_870_000,
        label: "Click to add caption",
      },
    ]),
  },
  pictureWithCaption: {
    name: "Picture with Caption",
    typeAttr: "picTx",
    label: "Picture with Caption",
    xml: layoutXml("picTx", "Picture with Caption", [
      {
        type: "pic",
        idx: 0,
        x: PAD_X,
        y: PAD_TOP,
        cx: SLIDE_W - PAD_X * 2,
        cy: 4_500_000,
        label: "Click to add picture",
      },
      {
        type: "title",
        idx: 1,
        x: PAD_X,
        y: PAD_TOP + 4_700_000,
        cx: SLIDE_W - PAD_X * 2,
        cy: 700_000,
        label: "Click to add title",
      },
      {
        type: "body",
        idx: 2,
        x: PAD_X,
        y: PAD_TOP + 5_500_000,
        cx: SLIDE_W - PAD_X * 2,
        cy: 700_000,
        label: "Click to add caption",
      },
    ]),
  },
  titleSlide: {
    name: "Title and Caption",
    typeAttr: "tx",
    label: "Title with Body",
    xml: layoutXml("tx", "Title with Body", [
      { type: "title", idx: 0, ...TITLE_BOX, label: "Click to add title" },
      { type: "body", idx: 1, ...SUBTITLE_BOX, label: "Click to add subtitle" },
    ]),
  },
  bigNumber: {
    name: "Big Number",
    typeAttr: "obj",
    label: "Big Number",
    xml: layoutXml("obj", "Big Number", [
      {
        type: "title",
        idx: 0,
        x: PAD_X,
        y: 2_400_000,
        cx: SLIDE_W - PAD_X * 2,
        cy: 1_800_000,
        label: "100%",
        anchorCenter: true,
      },
      {
        type: "body",
        idx: 1,
        x: PAD_X,
        y: 4_300_000,
        cx: SLIDE_W - PAD_X * 2,
        cy: 800_000,
        label: "Caption",
        anchorCenter: true,
      },
    ]),
  },
};
