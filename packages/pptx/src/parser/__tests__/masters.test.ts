/**
 * F1 master-editing — typed model + parser/serializer round-trip tests.
 *
 * These are the contract tests for promoting `OpaquePart` → typed
 * `SlideMaster` / `Theme`. They guard four properties:
 *  1. The parser walks `ppt/slideMasters/`, `ppt/slideLayouts/`, and
 *     `ppt/theme/` strictly via the rels graph (no hardcoded paths).
 *  2. Each typed `SlideMaster` carries the master's numeric `masterId`
 *     from the presentation's `<p:sldMasterIdLst>`, the theme part it
 *     points at, and the layouts it owns.
 *  3. Each typed `SlideLayout` is enriched with `layoutId`,
 *     `masterPartPath`, and `type` derived from the parent master's
 *     rels + `<p:sldLayoutIdLst>`.
 *  4. Untouched master/layout/theme parts round-trip byte-identical
 *     through serialise → re-parse (no dirty flags set).
 */

import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { sha256Hex } from "@officeai/core";
import { parsePptx } from "../parse.js";
import { serializePptx } from "../../serializer/serialize.js";

const FIXTURES = resolve(
  fileURLToPath(import.meta.url),
  "..",
  "..",
  "..",
  "..",
  "..",
  "..",
  "fixtures",
  "pptx",
  "synthetic"
);

const FIXTURE = resolve(FIXTURES, "03-title-and-content.pptx");

describe("F1 master/layout/theme typed model", () => {
  it("promotes slide masters from OpaquePart to typed SlideMaster", async () => {
    const buf = await readFile(FIXTURE);
    const snap = await parsePptx(buf);

    expect(snap.root.masters.size).toBeGreaterThan(0);
    for (const [partPath, master] of snap.root.masters) {
      expect(master.partPath).toBe(partPath);
      expect(partPath).toMatch(/^ppt\/slideMasters\/slideMaster\d+\.xml$/);
      expect(master.raw.tag).toBe("p:sldMaster");
      expect(master.raw.subtree.length).toBeGreaterThan(0);
      // masterId comes from `<p:sldMasterIdLst>` and is always a real
      // (non-zero) number for masters listed by the presentation.
      expect(master.masterId).toBeGreaterThan(0);
      // Every layout the master owns must have a typed entry that
      // back-references this master.
      expect(master.layouts.length).toBeGreaterThan(0);
      for (const layout of master.layouts) {
        expect(layout.masterPartPath).toBe(partPath);
        expect(layout.partPath).toMatch(/^ppt\/slideLayouts\/slideLayout\d+\.xml$/);
      }
      // theme rel discovered via the master's rels.
      expect(master.themePartPath).toBeDefined();
      expect(master.themePartPath).toMatch(/^ppt\/theme\/theme\d+\.xml$/);
      // master rels XML is captured verbatim when present.
      expect(master.relsXml).toBeDefined();
      expect(master.relsXml).toContain("<Relationship");
    }
  });

  it("promotes themes from OpaquePart to typed Theme", async () => {
    const buf = await readFile(FIXTURE);
    const snap = await parsePptx(buf);

    expect(snap.root.theme.size).toBeGreaterThan(0);
    for (const [partPath, theme] of snap.root.theme) {
      expect(theme.partPath).toBe(partPath);
      expect(partPath).toMatch(/^ppt\/theme\/theme\d+\.xml$/);
      expect(theme.raw.tag).toBe("a:theme");
      expect(theme.raw.subtree.length).toBeGreaterThan(0);
      // Most fixtures carry a theme name; we don't require it but if
      // present it must round-trip.
      if (theme.name !== undefined) {
        expect(theme.name.length).toBeGreaterThan(0);
      }
    }
  });

  it("enriches layouts with masterPartPath / layoutId / type", async () => {
    const buf = await readFile(FIXTURE);
    const snap = await parsePptx(buf);

    expect(snap.root.layouts.size).toBeGreaterThan(0);
    const masterPaths = new Set(snap.root.masters.keys());
    for (const layout of snap.root.layouts.values()) {
      // Every layout should resolve to one of the parsed masters.
      expect(layout.masterPartPath).toBeDefined();
      expect(masterPaths.has(layout.masterPartPath!)).toBe(true);
      // `type` mirrors `<p:sldLayout type="…">`. Built-in PowerPoint
      // layouts always carry one; we don't fail when absent (defensive
      // for hand-rolled XML), only when present it must be non-empty.
      if (layout.type !== undefined) expect(layout.type.length).toBeGreaterThan(0);
      // `relsXml` is captured verbatim — every layout has at least
      // a master rel pointing at its parent master.
      expect(layout.relsXml).toBeDefined();
      expect(layout.relsXml).toContain("slideMaster");
    }
  });

  it("discovers masters/layouts/themes via rels — no hardcoded paths", async () => {
    // The fixture's presentation rels are the single source of truth;
    // sanity-check that the parser's collected counts match what's on
    // disk by scanning the OPC package directly.
    const buf = await readFile(FIXTURE);
    const snap = await parsePptx(buf);

    const onDiskMasters = [...snap.container.parts.keys()].filter((p) =>
      /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(p)
    );
    const onDiskLayouts = [...snap.container.parts.keys()].filter((p) =>
      /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(p)
    );
    const onDiskThemes = [...snap.container.parts.keys()].filter((p) =>
      /^ppt\/theme\/theme\d+\.xml$/.test(p)
    );

    expect(snap.root.masters.size).toBe(onDiskMasters.length);
    expect(snap.root.layouts.size).toBe(onDiskLayouts.length);
    expect(snap.root.theme.size).toBe(onDiskThemes.length);
  });

  it("round-trips master/layout/theme parts byte-identically when not dirtied", async () => {
    const buf = await readFile(FIXTURE);
    const snap = await parsePptx(buf);

    const out = await serializePptx(snap);
    const snap2 = await parsePptx(out);

    // 1) Container-level: every master/layout/theme part hash should
    //    match between the original and the re-emitted package. The
    //    typed promotion is supposed to be a pure metadata lift —
    //    nothing in the output bytes should change.
    const tracked = new Set<string>([
      ...snap.root.masters.keys(),
      ...snap.root.layouts.keys(),
      ...snap.root.theme.keys(),
    ]);
    for (const partPath of tracked) {
      const before = sha256Hex(snap.container.readBytes(partPath));
      const after = sha256Hex(snap2.container.readBytes(partPath));
      expect(after, `${partPath} drifted`).toBe(before);
    }

    // 2) Typed-model identity must also survive a round-trip — masterId,
    //    themePartPath, and layout list shouldn't shift.
    expect(snap2.root.masters.size).toBe(snap.root.masters.size);
    for (const [path, before] of snap.root.masters) {
      const after = snap2.root.masters.get(path);
      expect(after).toBeDefined();
      expect(after!.masterId).toBe(before.masterId);
      expect(after!.themePartPath).toBe(before.themePartPath);
      expect(after!.layouts.map((l) => l.partPath)).toEqual(before.layouts.map((l) => l.partPath));
    }
    for (const [path, before] of snap.root.theme) {
      const after = snap2.root.theme.get(path);
      expect(after).toBeDefined();
      expect(after!.name).toBe(before.name);
    }
  });
});
