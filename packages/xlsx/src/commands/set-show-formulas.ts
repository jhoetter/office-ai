import type { CommandHandler } from "@officeai/core";
import type { Sheet, XlsxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { SetShowFormulasPayload } from "./payloads.js";
import { resolveSheet } from "./validation.js";

/**
 * `xlsx:set-show-formulas` — toggle Excel's "Show Formulas" view
 * mode on a worksheet (Formulas tab → Formula Auditing). Surfaces as
 * a per-sheet `showFormulas="1"` attribute on `<sheetView>` inside
 * `xl/worksheets/sheetN.xml`.
 *
 * Rather than reparse the sheet view, we patch the verbatim
 * `sheetViewsXml` block — exactly the same pattern the freeze-panes
 * command uses for `<pane>`. Other sheetView attributes (zoomScale,
 * tabSelected, workbookViewId, …) are preserved verbatim.
 *
 * Spec: ECMA-376 Part 1, §18.3.1.83 (sheetView).
 */
export const setShowFormulasHandler: CommandHandler<SetShowFormulasPayload, XlsxSnapshot> = {
  type: "xlsx:set-show-formulas",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const next = mutateSheetView(sheet.sheetViewsXml, payload.show);
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
          path: ["sheets", sheet.index, "showFormulas"],
          field: "showFormulas",
          summary: payload.show ? "true" : "false",
        },
      ]),
    };
  },
};

function mutateSheetView(current: string | undefined, show: boolean): string | undefined {
  // No existing block — only synthesise when toggling ON, otherwise no-op.
  if (!current) {
    if (!show) return undefined;
    return `<sheetViews><sheetView tabSelected="1" workbookViewId="0" showFormulas="1"/></sheetViews>`;
  }
  if (show) return setShowFormulasAttr(current, true);
  return setShowFormulasAttr(current, false);
}

function setShowFormulasAttr(xml: string, show: boolean): string {
  // Patch the FIRST <sheetView> open tag (the one bound to the
  // primary view). Subsequent sheetViews — split-pane secondaries
  // when present — keep whatever attributes they already had.
  const tagRe = /<sheetView\b([^>]*)(\/?)>/;
  const m = tagRe.exec(xml);
  if (!m) return xml;
  const head = xml.slice(0, m.index);
  const tail = xml.slice(m.index + m[0].length);
  let attrs = m[1];
  const selfClose = m[2] === "/";

  const showRe = /\sshowFormulas="[^"]*"/;
  if (show) {
    if (showRe.test(attrs)) {
      attrs = attrs.replace(showRe, ` showFormulas="1"`);
    } else {
      attrs = `${attrs} showFormulas="1"`;
    }
  } else {
    attrs = attrs.replace(showRe, "");
  }
  return `${head}<sheetView${attrs}${selfClose ? "/" : ""}>${tail}`;
}
