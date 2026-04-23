import { CommandError, type CommandHandler } from "@officeai/core";
import type { Sheet, XlsxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { SetSheetViewPayload } from "./payloads.js";
import { resolveSheet } from "./validation.js";

/**
 * `xlsx:set-sheet-view` — patch the first `<sheetView>` element of a
 * worksheet. Mirrors Excel's View tab toggles (view mode, gridlines,
 * headings, zoom, ruler, RTL).
 *
 * Implementation: we surgically rewrite the open tag attributes inside
 * the verbatim `sheetViewsXml` blob (same pattern as `set-show-formulas`
 * and the freeze-panes handler). Other `<sheetView>` attributes —
 * `tabSelected`, `workbookViewId`, `showZeros`, etc. that we don't
 * touch — round-trip verbatim, as does any `<pane>` / `<selection>`
 * subtree.
 *
 * Spec: ECMA-376 Part 1, §18.3.1.83 (sheetView).
 */
export const setSheetViewHandler: CommandHandler<SetSheetViewPayload, XlsxSnapshot> = {
  type: "xlsx:set-sheet-view",
  apply(snapshot, payload) {
    if (!hasAnyKnob(payload)) {
      throw new CommandError(
        "invalid-payload",
        "set-sheet-view: pass at least one of view/showGridLines/showRowColHeaders/showZeros/showRuler/showOutlineSymbols/rightToLeft/zoomScale/zoomScaleNormal."
      );
    }
    if (payload.zoomScale !== undefined) validateZoom(payload.zoomScale, "zoomScale");
    if (payload.zoomScaleNormal !== undefined) validateZoom(payload.zoomScaleNormal, "zoomScaleNormal");

    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const next = mutateSheetView(sheet.sheetViewsXml, payload);
    if (next === sheet.sheetViewsXml) {
      return { next: snapshot, diff: buildDiff(snapshot.revision, snapshot.revision, []) };
    }

    const updated: Sheet = { ...sheet, sheetViewsXml: next };
    const root = replaceSheet(snapshot.root, updated);
    const evolved = evolveSnapshot(snapshot, root, { sheets: [sheet.partPath] });
    return {
      next: evolved,
      diff: buildDiff(snapshot.revision, evolved.revision, [
        {
          kind: "node-updated",
          nodeId: sheet.id,
          path: ["sheets", sheet.index, "sheetView"],
          field: "sheetView",
          summary: summarise(payload),
        },
      ]),
    };
  },
};

function hasAnyKnob(p: SetSheetViewPayload): boolean {
  return (
    p.view !== undefined ||
    p.showGridLines !== undefined ||
    p.showRowColHeaders !== undefined ||
    p.showZeros !== undefined ||
    p.showRuler !== undefined ||
    p.showOutlineSymbols !== undefined ||
    p.rightToLeft !== undefined ||
    p.zoomScale !== undefined ||
    p.zoomScaleNormal !== undefined
  );
}

function validateZoom(value: number, name: string): void {
  if (!Number.isFinite(value) || value < 10 || value > 400) {
    throw new CommandError("invalid-payload", `set-sheet-view: ${name} must be 10–400 (got ${value}).`);
  }
}

function mutateSheetView(current: string | undefined, p: SetSheetViewPayload): string | undefined {
  let xml = current;
  if (!xml) {
    xml = `<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>`;
  }

  const tagRe = /<sheetView\b([^>]*?)(\/?)>/;
  const m = tagRe.exec(xml);
  if (!m) return current;

  const head = xml.slice(0, m.index);
  const tail = xml.slice(m.index + m[0].length);
  let attrs = m[1];
  const selfClose = m[2] === "/";

  if (p.view !== undefined) attrs = setStrAttr(attrs, "view", p.view === "normal" ? null : p.view);
  if (p.showGridLines !== undefined) attrs = setBoolAttr(attrs, "showGridLines", p.showGridLines);
  if (p.showRowColHeaders !== undefined) attrs = setBoolAttr(attrs, "showRowColHeaders", p.showRowColHeaders);
  if (p.showZeros !== undefined) attrs = setBoolAttr(attrs, "showZeros", p.showZeros);
  if (p.showRuler !== undefined) attrs = setBoolAttr(attrs, "showRuler", p.showRuler);
  if (p.showOutlineSymbols !== undefined)
    attrs = setBoolAttr(attrs, "showOutlineSymbols", p.showOutlineSymbols);
  if (p.rightToLeft !== undefined) attrs = setBoolAttr(attrs, "rightToLeft", p.rightToLeft);
  if (p.zoomScale !== undefined) attrs = setNumAttr(attrs, "zoomScale", p.zoomScale);
  if (p.zoomScaleNormal !== undefined) attrs = setNumAttr(attrs, "zoomScaleNormal", p.zoomScaleNormal);

  return `${head}<sheetView${attrs}${selfClose ? "/" : ""}>${tail}`;
}

function setStrAttr(attrs: string, name: string, value: string | null): string {
  const re = new RegExp(`\\s${name}="[^"]*"`);
  if (value === null) return attrs.replace(re, "");
  if (re.test(attrs)) return attrs.replace(re, ` ${name}="${value}"`);
  return `${attrs} ${name}="${value}"`;
}

function setBoolAttr(attrs: string, name: string, value: boolean): string {
  // OOXML defaults for these toggles are mostly `true`; "true" can
  // round-trip as either omitting the attr or setting it to "1". We
  // standardise: setting `true` → drop the attr (Excel default), setting
  // `false` → emit `name="0"`. Exception: showZeros default is true and
  // setting it true should drop. Same for showGridLines/showRowColHeaders.
  // For `showRuler`, the default is also true — we follow the same rule.
  return setStrAttr(attrs, name, value ? null : "0");
}

function setNumAttr(attrs: string, name: string, value: number): string {
  return setStrAttr(attrs, name, String(Math.floor(value)));
}

function summarise(p: SetSheetViewPayload): string {
  const parts: string[] = [];
  if (p.view) parts.push(p.view);
  if (p.showGridLines !== undefined) parts.push(`grid=${p.showGridLines}`);
  if (p.showRowColHeaders !== undefined) parts.push(`hdrs=${p.showRowColHeaders}`);
  if (p.zoomScale !== undefined) parts.push(`zoom=${p.zoomScale}`);
  return parts.join(",") || "patched";
}
