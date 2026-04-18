import { CommandError, type CommandHandler } from "@officeai/core";
import type * as XLSX from "xlsx";
import type { Sheet, XlsxSnapshot, XlsxWorkbook } from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import type { AddSheetPayload } from "./payloads.js";
import { assertUniqueSheetName, validateSheetName } from "./validation.js";

/**
 * `xlsx:add-sheet` — append or insert a fresh, empty worksheet.
 *
 * Spec: `spec/xlsx/agent-commands.md` §11.
 *
 * Pipeline:
 *   1. Validate the proposed name (Excel rules + uniqueness, case-
 *      insensitive) and the optional `at` insert position.
 *   2. Mint a fresh `sheetId` (smallest unused positive integer) and a
 *      fresh part path (`xl/worksheets/sheetN.xml`) that does not
 *      collide with anything already in the container.
 *   3. Splice the typed `Sheet` into `workbook.sheets` and re-derive
 *      the `index` of every shifted neighbour.
 *   4. Mint a parallel SheetJS `WorkSheet` (dense, `!ref="A1"`,
 *      no cells) so the serializer's existing `rewriteDirtySheets`
 *      pipeline can emit the part bytes without a special case.
 *   5. Set `dirty.workbook | rels | contentTypes | sheets[partPath]`
 *      so the serializer rewrites the workbook `<sheets>` block, the
 *      workbook rels, the content-types overrides, and the new sheet
 *      part.
 *
 * The serializer is responsible for minting the new `r:id` (via the
 * `RelationshipGraph`) and for splicing the new `<sheet>` entry into
 * `xl/workbook.xml` at the correct tab index.
 */
export const addSheetHandler: CommandHandler<AddSheetPayload, XlsxSnapshot> = {
  type: "xlsx:add-sheet",
  apply(snapshot, payload, ctx) {
    validateSheetName(payload.name);
    assertUniqueSheetName(snapshot.root, payload.name);

    const total = snapshot.root.sheets.length;
    const at = payload.at ?? total;
    if (!Number.isInteger(at) || at < 0 || at > total) {
      throw new CommandError(
        "invalid-position",
        `Position ${payload.at} out of range; must be between 0 and ${total}`
      );
    }

    const newSheetId = mintSheetId(snapshot.root);
    const partPath = mintPartPath(snapshot);

    const newSheet: Sheet = {
      id: ctx.mintNodeId(),
      sheetId: String(newSheetId),
      name: payload.name,
      index: at,
      state: "visible",
      kind: "worksheet",
      partPath,
      cells: new Map(),
      merges: [],
      comments: [],
      commentAuthors: [],
      columnWidths: new Map(),
      rowHeights: new Map(),
    };

    const sheets = snapshot.root.sheets.slice();
    sheets.splice(at, 0, newSheet);
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].index !== i) {
        sheets[i] = { ...sheets[i], index: i };
      }
    }

    syncSheetjsAdd(snapshot.root, payload.name, at);

    const nextWorkbook: XlsxWorkbook = { ...snapshot.root, sheets };
    const next = evolveSnapshot(snapshot, nextWorkbook, {
      workbook: true,
      rels: true,
      contentTypes: true,
      sheets: [partPath],
    });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, [
        {
          kind: "node-inserted",
          nodeId: newSheet.id,
          path: ["sheets", at],
          summary: `Added sheet '${payload.name}' at index ${at}`,
          meta: {
            name: payload.name,
            at,
            sheetId: newSheet.sheetId,
            partPath,
          },
        },
      ]),
    };
  },
};

function mintSheetId(workbook: XlsxWorkbook): number {
  const taken = new Set<number>();
  for (const s of workbook.sheets) {
    const n = Number(s.sheetId);
    if (Number.isInteger(n) && n > 0) taken.add(n);
  }
  let i = 1;
  while (taken.has(i)) i++;
  return i;
}

function mintPartPath(snapshot: XlsxSnapshot): string {
  const taken = new Set<string>();
  for (const s of snapshot.root.sheets) taken.add(s.partPath);
  for (const path of snapshot.container.parts.keys()) taken.add(path);
  let i = 1;
  while (taken.has(`xl/worksheets/sheet${i}.xml`)) i++;
  return `xl/worksheets/sheet${i}.xml`;
}

function syncSheetjsAdd(workbook: XlsxWorkbook, name: string, at: number): void {
  const book = workbook.sheetjs;
  book.SheetNames.splice(at, 0, name);
  book.Sheets[name] = { "!ref": "A1", "!data": [] } as unknown as XLSX.WorkSheet;
}
