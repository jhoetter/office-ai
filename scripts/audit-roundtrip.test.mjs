/**
 * Tests for the round-trip audit walkers.
 *
 * Run with:
 *
 *   node --test scripts/audit-roundtrip.test.mjs
 *
 * The tests use hand-built snapshot mocks rather than going through
 * the real XLSX/DOCX/PPTX agents — that lets us cover individual
 * counters (font-color, hyperlink-rel, connector, animation-step)
 * with the smallest possible inputs and keeps the unit tests free
 * of an `pnpm --filter @officeai/{xlsx,docx,pptx} build` prerequisite.
 *
 * The tally-fn shapes mirror the real `XlsxSnapshot` / `DocxDocument`
 * / pptx slide shapes that the walker dereferences. Anything the
 * walker doesn't read can stay omitted.
 */

import { strict as assert } from "node:assert";
import { test } from "node:test";

import {
  fillIsNonDefault,
  fontHasNonDefaultColor,
  spotCheckHashDocx,
  spotCheckHashPptx,
  spotCheckHashXlsx,
  stableStringify,
  tallyDocx,
  tallyPptx,
  tallyXlsx,
} from "./audit-roundtrip.mjs";

/* ── stableStringify ──────────────────────────────────────────── */

test("stableStringify sorts keys deterministically", () => {
  const a = stableStringify({ b: 1, a: 2 });
  const b = stableStringify({ a: 2, b: 1 });
  assert.equal(a, b);
});

test("stableStringify ignores `id` fields (parse-time minted)", () => {
  const a = stableStringify({ id: "node-1", value: 42 });
  const b = stableStringify({ id: "node-99", value: 42 });
  assert.equal(a, b);
});

test("stableStringify expands Map and Set in stable order", () => {
  const m1 = new Map([
    ["b", 2],
    ["a", 1],
  ]);
  const m2 = new Map([
    ["a", 1],
    ["b", 2],
  ]);
  assert.equal(stableStringify(m1), stableStringify(m2));
  const s1 = new Set([3, 1, 2]);
  const s2 = new Set([2, 3, 1]);
  assert.equal(stableStringify(s1), stableStringify(s2));
});

/* ── XLSX font / fill helpers ────────────────────────────────── */

const defaultFont = { color: { theme: 1 } };

test("fontHasNonDefaultColor returns false for default font", () => {
  assert.equal(fontHasNonDefaultColor(defaultFont, defaultFont), false);
});

test("fontHasNonDefaultColor returns false when no color is set", () => {
  assert.equal(fontHasNonDefaultColor({ name: "Calibri" }, defaultFont), false);
});

test("fontHasNonDefaultColor returns true for an authored RGB", () => {
  const font = { color: { rgb: "FFFF0000" } };
  assert.equal(fontHasNonDefaultColor(font, defaultFont), true);
});

test("fillIsNonDefault returns false for the stock `none` fill", () => {
  assert.equal(fillIsNonDefault({ kind: "pattern", patternType: "none" }), false);
});

test("fillIsNonDefault returns false for the stock `gray125` fill", () => {
  assert.equal(fillIsNonDefault({ kind: "pattern", patternType: "gray125" }), false);
});

test("fillIsNonDefault returns true for a solid colored fill", () => {
  const fill = { kind: "pattern", patternType: "solid", fgColor: { rgb: "FF00FF00" } };
  assert.equal(fillIsNonDefault(fill), true);
});

test("fillIsNonDefault returns true for a gradient fill", () => {
  assert.equal(fillIsNonDefault({ kind: "gradient" }), true);
});

/* ── tallyXlsx ───────────────────────────────────────────────── */

function makeXlsxSnapshot({ cells, styles }) {
  return {
    root: {
      sheets: [
        {
          cells: new Map(cells.map((c) => [`${c.row}:${c.col}`, c])),
          merges: [],
          cols: [],
          charts: [],
          images: [],
        },
      ],
      styles,
    },
  };
}

test("tallyXlsx counts font-color and font-fill on authored cells", () => {
  const styles = {
    fonts: [
      { name: "Calibri", color: { theme: 1 } }, // 0 = default
      { name: "Calibri", color: { rgb: "FFFF0000" } }, // 1 = red
    ],
    fills: [
      { kind: "pattern", patternType: "none" }, // 0
      { kind: "pattern", patternType: "gray125" }, // 1
      { kind: "pattern", patternType: "solid", fgColor: { rgb: "FFFFFF00" } }, // 2
    ],
    cellXfs: [
      { fontId: 0, fillId: 0 }, // 0 = default
      { fontId: 1, fillId: 0 }, // 1 = colored font, default fill
      { fontId: 0, fillId: 2 }, // 2 = default font, yellow fill
      { fontId: 1, fillId: 2 }, // 3 = colored font + yellow fill
    ],
    numFmts: new Map(),
  };
  const snap = makeXlsxSnapshot({
    styles,
    cells: [
      { row: 0, col: 0, value: "plain" },
      { row: 0, col: 1, value: "red", styleId: 1 },
      { row: 1, col: 0, value: "yellow", styleId: 2 },
      { row: 1, col: 1, value: "both", styleId: 3 },
      { row: 2, col: 0, value: "default-styleid", styleId: 0 },
    ],
  });

  const { counts } = tallyXlsx(snap);
  assert.equal(counts.get("font-color"), 2, "font-color should fire on cells 1 and 3");
  assert.equal(counts.get("font-fill"), 2, "font-fill should fire on cells 2 and 3");
  assert.equal(counts.get("cells"), 5);
});

test("tallyXlsx font-color stays 0 when no cell carries an authored font", () => {
  const styles = {
    fonts: [{ name: "Calibri", color: { theme: 1 } }],
    fills: [
      { kind: "pattern", patternType: "none" },
      { kind: "pattern", patternType: "gray125" },
    ],
    cellXfs: [{ fontId: 0, fillId: 0 }],
    numFmts: new Map(),
  };
  const snap = makeXlsxSnapshot({
    styles,
    cells: [{ row: 0, col: 0, value: "x" }],
  });
  const { counts } = tallyXlsx(snap);
  assert.equal(counts.get("font-color") ?? 0, 0);
  assert.equal(counts.get("font-fill") ?? 0, 0);
});

test("tallyXlsx spot-check hash is stable across two identical builds", () => {
  const build = () =>
    makeXlsxSnapshot({
      styles: {
        fonts: [{ color: { theme: 1 } }],
        fills: [
          { kind: "pattern", patternType: "none" },
          { kind: "pattern", patternType: "gray125" },
        ],
        cellXfs: [{ fontId: 0, fillId: 0 }],
        numFmts: new Map(),
      },
      cells: [
        { row: 0, col: 0, value: "hello" },
        { row: 0, col: 1, value: 42 },
        { row: 1, col: 0, value: true },
      ],
    });
  const a = tallyXlsx(build()).spotCheckHash;
  const b = tallyXlsx(build()).spotCheckHash;
  assert.equal(a, b);
  assert.match(a, /^[a-f0-9]{40}$/);
});

test("tallyXlsx spot-check hash changes when a cell value changes", () => {
  const styles = {
    fonts: [{ color: { theme: 1 } }],
    fills: [
      { kind: "pattern", patternType: "none" },
      { kind: "pattern", patternType: "gray125" },
    ],
    cellXfs: [{ fontId: 0, fillId: 0 }],
    numFmts: new Map(),
  };
  const a = tallyXlsx(
    makeXlsxSnapshot({ styles, cells: [{ row: 0, col: 0, value: "a" }] })
  ).spotCheckHash;
  const b = tallyXlsx(
    makeXlsxSnapshot({ styles, cells: [{ row: 0, col: 0, value: "b" }] })
  ).spotCheckHash;
  assert.notEqual(a, b);
});

test("tallyXlsx spot-check uses row-major ordering, ignoring map insert order", () => {
  const styles = {
    fonts: [{ color: { theme: 1 } }],
    fills: [
      { kind: "pattern", patternType: "none" },
      { kind: "pattern", patternType: "gray125" },
    ],
    cellXfs: [{ fontId: 0, fillId: 0 }],
    numFmts: new Map(),
  };
  const ordered = tallyXlsx(
    makeXlsxSnapshot({
      styles,
      cells: [
        { row: 0, col: 0, value: "A" },
        { row: 0, col: 1, value: "B" },
        { row: 1, col: 0, value: "C" },
      ],
    })
  ).spotCheckHash;
  const shuffled = tallyXlsx(
    makeXlsxSnapshot({
      styles,
      cells: [
        { row: 1, col: 0, value: "C" },
        { row: 0, col: 1, value: "B" },
        { row: 0, col: 0, value: "A" },
      ],
    })
  ).spotCheckHash;
  assert.equal(ordered, shuffled);
});

/* ── tallyDocx ───────────────────────────────────────────────── */

function makeDocxSnapshot({ blocks, rels = [] }) {
  return {
    root: {
      body: blocks,
      sections: [],
      relationships: new Map([
        ["word/document.xml", rels.map((id) => ({ id, type: "hyperlink", target: "https://example.com" }))],
      ]),
    },
  };
}

test("tallyDocx counts hyperlink-rel only when relationshipId resolves", () => {
  const blocks = [
    {
      kind: "paragraph",
      properties: {},
      children: [
        { kind: "run", properties: {}, children: [{ kind: "text", text: "hi" }] },
        // External link with a resolvable rel — counted.
        {
          kind: "hyperlink",
          relationshipId: "rId7",
          children: [
            { kind: "run", properties: {}, children: [{ kind: "text", text: "click" }] },
          ],
        },
        // Internal anchor — must NOT count.
        {
          kind: "hyperlink",
          anchor: "Heading1",
          children: [
            { kind: "run", properties: {}, children: [{ kind: "text", text: "jump" }] },
          ],
        },
        // Unresolved external (rel id not in document.rels) — also skipped.
        {
          kind: "hyperlink",
          relationshipId: "rIdGhost",
          children: [
            { kind: "run", properties: {}, children: [{ kind: "text", text: "broken" }] },
          ],
        },
      ],
    },
  ];
  const snap = makeDocxSnapshot({ blocks, rels: ["rId7"] });
  const { counts } = tallyDocx(snap);
  assert.equal(counts.get("hyperlink-rel"), 1);
});

test("tallyDocx spot-check hash captures first 5 runs across paragraphs", () => {
  const blocks = [
    {
      kind: "paragraph",
      children: [
        { kind: "run", properties: { bold: true }, children: [{ kind: "text", text: "A" }] },
        { kind: "run", properties: {}, children: [{ kind: "text", text: "B" }] },
      ],
    },
    {
      kind: "paragraph",
      children: [
        { kind: "run", properties: {}, children: [{ kind: "text", text: "C" }] },
      ],
    },
  ];
  const a = tallyDocx(makeDocxSnapshot({ blocks })).spotCheckHash;
  const b = tallyDocx(makeDocxSnapshot({ blocks })).spotCheckHash;
  assert.equal(a, b);

  const tweaked = JSON.parse(JSON.stringify(blocks));
  tweaked[0].children[0].properties.bold = false;
  const c = tallyDocx(makeDocxSnapshot({ blocks: tweaked })).spotCheckHash;
  assert.notEqual(a, c, "spot-check should change when a run property toggles");
});

/* ── tallyPptx ───────────────────────────────────────────────── */

function makePptxSnapshot({ slides }) {
  return { root: { slides } };
}

test("tallyPptx counts connectors and animation steps", () => {
  const slides = [
    {
      shapes: [
        { kind: "text", name: "Title", text: { paragraphs: [] } },
        {
          kind: "connector",
          name: "Line 1",
          connectorType: "straight",
          start: { kind: "free", xEmu: 0, yEmu: 0 },
          end: { kind: "free", xEmu: 100, yEmu: 100 },
        },
        {
          kind: "connector",
          name: "Line 2",
          connectorType: "elbow",
          start: { kind: "free", xEmu: 0, yEmu: 0 },
          end: { kind: "free", xEmu: 200, yEmu: 200 },
        },
      ],
      animations: [
        { effect: "fade", targetCNvPrId: 1, order: 0 },
        { effect: "fly-in", targetCNvPrId: 2, order: 1 },
      ],
    },
    {
      shapes: [{ kind: "picture", name: "Pic" }],
      animations: [{ effect: "appear", targetCNvPrId: 1, order: 0 }],
    },
  ];
  const { counts } = tallyPptx(makePptxSnapshot({ slides }));
  assert.equal(counts.get("connector"), 2);
  assert.equal(counts.get("animation-step"), 3);
  assert.equal(counts.get("shapes"), 4);
});

test("tallyPptx spot-check hash is stable, and changes when a shape moves", () => {
  const slides = [
    {
      shapes: [
        {
          kind: "text",
          name: "Title",
          cNvPrId: 1,
          position: { xEmu: 0, yEmu: 0 },
          size: { cxEmu: 100, cyEmu: 100 },
        },
        { kind: "picture", name: "Pic", cNvPrId: 2 },
      ],
      animations: [],
    },
  ];
  const a = tallyPptx(makePptxSnapshot({ slides })).spotCheckHash;
  const b = tallyPptx(
    makePptxSnapshot({ slides: JSON.parse(JSON.stringify(slides)) })
  ).spotCheckHash;
  assert.equal(a, b);

  const moved = JSON.parse(JSON.stringify(slides));
  moved[0].shapes[0].position.xEmu = 5000;
  const c = tallyPptx(makePptxSnapshot({ slides: moved })).spotCheckHash;
  assert.notEqual(a, c);
});

/* ── direct spot-check helpers ───────────────────────────────── */

test("spot-check helpers are exported and produce SHA-1 hex", () => {
  const xlsxHash = spotCheckHashXlsx({ root: { sheets: [] } });
  const docxHash = spotCheckHashDocx({ root: { body: [] } });
  const pptxHash = spotCheckHashPptx({ root: { slides: [] } });
  for (const h of [xlsxHash, docxHash, pptxHash]) {
    assert.match(h, /^[a-f0-9]{40}$/, `expected SHA-1 hex, got ${h}`);
  }
});
