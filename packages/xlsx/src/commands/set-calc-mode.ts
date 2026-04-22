import type { CommandHandler } from "@officeai/core";
import type { XlsxSnapshot, XlsxWorkbook } from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import type { SetCalcModePayload } from "./payloads.js";

/**
 * `xlsx:set-calc-mode` — toggle Excel's "Calculation Options" between
 * Automatic / Automatic-except-tables / Manual, plus the
 * "Recalculate workbook before saving" + "Enable iterative
 * calculation" knobs from the Formulas ribbon.
 *
 * The `<calcPr>` element on `xl/workbook.xml` is mutated by attribute
 * surgery so other attrs (`calcId`, `iterateCount`, `iterateDelta`)
 * survive untouched. When no `<calcPr>` element existed in the
 * source we synthesise a minimal one carrying just the requested
 * attributes; otherwise we patch in place.
 *
 * Spec: ECMA-376 Part 1, §18.2.2 (calcPr).
 */
export const setCalcModeHandler: CommandHandler<SetCalcModePayload, XlsxSnapshot> = {
  type: "xlsx:set-calc-mode",
  apply(snapshot, payload) {
    const current = snapshot.root.calcPrXml;
    const next = mutateCalcPrXml(current, payload);
    if (next === current) {
      return { next: snapshot, diff: buildDiff(snapshot.revision, snapshot.revision, []) };
    }

    const root: XlsxWorkbook = { ...snapshot.root, calcPrXml: next };
    const evolved = evolveSnapshot(snapshot, root, { workbook: true });
    return {
      next: evolved,
      diff: buildDiff(snapshot.revision, evolved.revision, [
        {
          kind: "node-updated",
          nodeId: snapshot.root.id,
          path: ["workbook", "calcPr"],
          field: "calcPr",
          summary: payload.calcMode ?? "patched",
        },
      ]),
    };
  },
};

function mutateCalcPrXml(current: string | undefined, payload: SetCalcModePayload): string {
  let xml = current ?? `<calcPr/>`;

  if (payload.calcMode !== undefined) {
    if (payload.calcMode === "auto") {
      xml = removeAttr(xml, "calcMode");
    } else {
      xml = setAttr(xml, "calcMode", payload.calcMode);
    }
  }
  if (payload.calcOnSave !== undefined) {
    xml = setBoolAttr(xml, "calcOnSave", payload.calcOnSave);
  }
  if (payload.iterate !== undefined) {
    xml = setBoolAttr(xml, "iterate", payload.iterate);
  }
  if (payload.iterateCount !== undefined) {
    xml = setAttr(xml, "iterateCount", String(Math.max(1, Math.floor(payload.iterateCount))));
  }
  if (payload.iterateDelta !== undefined) {
    xml = setAttr(xml, "iterateDelta", String(payload.iterateDelta));
  }
  return xml;
}

function setAttr(xml: string, name: string, value: string): string {
  const escaped = value.replace(/&/g, "&amp;").replace(/"/g, "&quot;");
  const attrRe = new RegExp(`(\\s${name}=")[^"]*(")`);
  if (attrRe.test(xml)) return xml.replace(attrRe, `$1${escaped}$2`);
  // No existing attr — splice into the open tag.
  return xml.replace(/<calcPr\b/, `<calcPr ${name}="${escaped}"`);
}

function setBoolAttr(xml: string, name: string, value: boolean): string {
  if (!value) return removeAttr(xml, name);
  return setAttr(xml, name, "1");
}

function removeAttr(xml: string, name: string): string {
  return xml.replace(new RegExp(`\\s${name}="[^"]*"`), "");
}
